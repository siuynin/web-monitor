# -*- coding: utf-8 -*-
"""
Website Monitor - theo dõi các website liệt kê trong file Excel theo thời gian thực,
gửi email cảnh báo ngay khi có website lỗi và khi website hoạt động lại.

Cách dùng:
    python monitor.py               # chạy liên tục
    python monitor.py --once        # kiểm tra 1 lượt rồi thoát
    python monitor.py --test-email  # gửi email thử để kiểm tra cấu hình SMTP
"""
import argparse
import configparser
import csv
import html
import json
import logging
import logging.handlers
import os
import re
import smtplib
import ssl
import sys
import time
import warnings
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr

warnings.filterwarnings("ignore")  # bỏ cảnh báo phiên bản urllib3/chardet

import openpyxl  # noqa: E402
import requests  # noqa: E402
from openpyxl.styles import Font, PatternFill  # noqa: E402

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.ini")
STATE_FILE = os.path.join(BASE_DIR, "state.json")
HISTORY_FILE = os.path.join(BASE_DIR, "history.csv")
DASHBOARD_FILE = os.path.join(BASE_DIR, "dashboard.html")
REPORT_FILE = os.path.join(BASE_DIR, "status_report.xlsx")
LOG_FILE = os.path.join(BASE_DIR, "monitor.log")

DOMAIN_RE = re.compile(r"^(https?://)?([a-z0-9-]+\.)+[a-z]{2,}(:\d+)?(/\S*)?$", re.I)
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/126.0 Safari/537.36 WebsiteMonitor/1.0",
    "Accept": "text/html,application/xhtml+xml,*/*;q=0.8",
}

log = logging.getLogger("monitor")


# ----------------------------------------------------------------------------- config
ENV_OVERRIDES = {  # biến môi trường (GitHub Secrets) ghi đè config.ini
    "SMTP_HOST": "smtp_host", "SMTP_PORT": "smtp_port", "SMTP_USE_SSL": "use_ssl",
    "SMTP_USERNAME": "username", "SMTP_PASSWORD": "password", "MAIL_FROM": "from_addr",
    "MAIL_FROM_NAME": "from_name", "ALERT_TO": "to",
}


def load_config():
    path = CONFIG_FILE if os.path.exists(CONFIG_FILE) else os.path.join(BASE_DIR, "config.example.ini")
    if not os.path.exists(path):
        sys.exit(f"Không tìm thấy {CONFIG_FILE}. Hãy copy config.example.ini thành config.ini và điền thông tin.")
    cp = configparser.ConfigParser(interpolation=None)  # mật khẩu có ký tự % vẫn đọc đúng
    cp.read(path, encoding="utf-8")
    for env, key in ENV_OVERRIDES.items():
        if os.environ.get(env, "").strip():
            cp["email"][key] = os.environ[env].strip()
    m, e = cp["monitor"], cp["email"]
    excel = m.get("excel_file", "website serve.xlsx")
    return {
        "excel_file": excel if os.path.isabs(excel) else os.path.join(BASE_DIR, excel),
        "sheet": m.get("sheet", "").strip(),
        "interval": m.getint("interval_seconds", 60),
        "timeout": m.getint("timeout_seconds", 15),
        "retries": m.getint("retries", 2),
        "retry_delay": m.getint("retry_delay_seconds", 5),
        "slow_ms": m.getint("slow_threshold_ms", 0),
        "remind_minutes": m.getint("remind_minutes", 60),
        "max_alerts": m.getint("max_alerts_per_domain", 5),
        "workers": m.getint("workers", 20),
        "email": {
            "enabled": e.getboolean("enabled", True),
            "host": e.get("smtp_host", "smtp.gmail.com"),
            "port": e.getint("smtp_port", 587),
            # cổng 465 luôn là SSL trực tiếp; dùng STARTTLS ở cổng này sẽ bị server ngắt kết nối
            "use_ssl": e.getboolean("use_ssl", False) or e.getint("smtp_port", 587) == 465,
            "username": e.get("username", "").strip(),
            "password": _clean_password(e.get("password", "")),
            "from_addr": e.get("from_addr", "").strip() or e.get("username", "").strip(),
            "from_name": e.get("from_name", "Website Monitor").strip(),
            "to": [x.strip() for x in re.split(r"[,;]", e.get("to", "")) if x.strip()],
            "notify_recovery": e.getboolean("notify_recovery", True),
        },
    }


def _clean_password(pw):
    pw = pw.strip().strip("\"'").strip()
    # Mật khẩu ứng dụng Google hiển thị dạng "abcd efgh ijkl mnop" -> bỏ dấu cách
    if re.fullmatch(r"[a-zA-Z]{4}(\s?[a-zA-Z]{4}){3}", pw):
        pw = re.sub(r"\s", "", pw)
    return pw


def describe_credentials(ec):
    """Mô tả dạng thông tin đăng nhập mà KHÔNG lộ giá trị (để chẩn đoán lỗi 535)."""
    u, p = ec["username"], ec["password"]
    dom = u.rsplit("@", 1)[1].lower() if "@" in u else ""
    user = ("thiếu '@domain' (cần ghi đầy đủ địa chỉ email)" if not dom else
            "đuôi @gmail.com" if dom == "gmail.com" else "đuôi tên miền riêng (không phải gmail.com)")
    if any(c.isspace() for c in u):
        user += ", CÓ dấu cách"
    pw = f"dài {len(p)} ký tự"
    if re.fullmatch(r"[a-z]{16}", p):
        pw += ", đúng dạng mật khẩu ứng dụng Google (16 chữ thường)"
    else:
        pw += ", KHÔNG giống mật khẩu ứng dụng Google (phải đúng 16 chữ cái a-z)"
    if any(ord(c) > 127 for c in p):
        pw += ", có ký tự lạ/có dấu"
    return f"SMTP {ec['host']}:{ec['port']} - username: {user}; password: {pw}"


def setup_logging():
    log.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    fh = logging.handlers.RotatingFileHandler(LOG_FILE, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
    fh.setFormatter(fmt)
    sh = logging.StreamHandler(open(sys.stdout.fileno(), "w", encoding="utf-8", closefd=False))
    sh.setFormatter(fmt)
    log.addHandler(fh)
    log.addHandler(sh)


# ----------------------------------------------------------------------------- excel
_last_sites = []


def read_sites(cfg):
    """Đọc lại file Excel mỗi lượt để thêm/bớt website không cần khởi động lại.
    Mọi ô có dạng tên miền/URL đều được coi là 1 website."""
    global _last_sites
    try:
        wb = openpyxl.load_workbook(cfg["excel_file"], read_only=True, data_only=True)
        ws = wb[cfg["sheet"]] if cfg["sheet"] else wb.worksheets[0]
        sites, seen = [], set()
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if not isinstance(cell, str):
                    continue
                v = cell.strip().rstrip("/")
                if DOMAIN_RE.match(v) and v.lower() not in seen:
                    seen.add(v.lower())
                    sites.append(v)
        wb.close()
        if sites != _last_sites:
            log.info("Đọc được %d website từ %s", len(sites), os.path.basename(cfg["excel_file"]))
        _last_sites = sites
    except Exception as ex:  # file đang bị khoá/ghi dở -> dùng danh sách cũ
        log.warning("Không đọc được file Excel (%s), dùng danh sách cũ (%d site)", ex, len(_last_sites))
    return list(_last_sites)


# ----------------------------------------------------------------------------- check
def _request(url, timeout):
    t0 = time.perf_counter()
    r = requests.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True, stream=True)
    ms = int((time.perf_counter() - t0) * 1000)
    r.close()
    return r, ms


def check_once(site, timeout):
    """Trả về dict: ok, code, ms, error, final_url"""
    urls = [site] if site.lower().startswith("http") else ["https://" + site, "http://" + site]
    last_err = None
    for i, url in enumerate(urls):
        try:
            r, ms = _request(url, timeout)
            if r.status_code < 400:
                note = "Chỉ chạy HTTP (HTTPS không truy cập được)" if i == 1 else ""
                return {"ok": True, "code": r.status_code, "ms": ms, "error": note, "final_url": r.url}
            return {"ok": False, "code": r.status_code, "ms": ms,
                    "error": f"HTTP {r.status_code} {r.reason or ''}".strip(), "final_url": r.url}
        except requests.exceptions.SSLError as ex:
            last_err = f"Lỗi chứng chỉ SSL: {_short(ex)}"
            return {"ok": False, "code": None, "ms": None, "error": last_err, "final_url": url}
        except requests.exceptions.Timeout:
            last_err = f"Quá thời gian phản hồi ({timeout}s)"
            break  # timeout thì không thử http nữa
        except requests.exceptions.ConnectionError as ex:
            s = str(ex)
            if "getaddrinfo" in s or "NameResolution" in s or "Name or service" in s:
                last_err = "Không phân giải được tên miền (DNS lỗi / hết hạn domain)"
                break
            last_err = f"Không kết nối được: {_short(ex)}"
        except Exception as ex:
            last_err = f"Lỗi: {_short(ex)}"
    return {"ok": False, "code": None, "ms": None, "error": last_err, "final_url": urls[0]}


def _short(ex, n=160):
    s = re.sub(r"\s+", " ", str(ex))
    m = re.search(r"\(Caused by (.*)\)", s)
    s = m.group(1) if m else s
    return s[:n]


def check_site(site, cfg):
    """Kiểm tra lại vài lần trước khi kết luận lỗi để tránh báo nhầm."""
    res = check_once(site, cfg["timeout"])
    tries = 0
    while not res["ok"] and tries < cfg["retries"]:
        tries += 1
        time.sleep(cfg["retry_delay"])
        res = check_once(site, cfg["timeout"])
    if res["ok"] and cfg["slow_ms"] and res["ms"] and res["ms"] > cfg["slow_ms"]:
        res["error"] = (res["error"] + "; " if res["error"] else "") + f"Phản hồi chậm ({res['ms']} ms)"
    res["site"] = site
    res["checked_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return res


# ----------------------------------------------------------------------------- state
def load_state():
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_state(state):
    tmp = STATE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=1)
    os.replace(tmp, STATE_FILE)


def append_history(rows):
    new = not os.path.exists(HISTORY_FILE)
    with open(HISTORY_FILE, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["Thời gian", "Website", "Sự kiện", "Chi tiết"])
        w.writerows(rows)


def fmt_duration(sec):
    sec = int(sec)
    h, r = divmod(sec, 3600)
    m, s = divmod(r, 60)
    return (f"{h} giờ " if h else "") + (f"{m} phút " if m or h else "") + f"{s} giây"


# ----------------------------------------------------------------------------- email
def send_email(cfg, subject, body_html):
    ec = cfg["email"]
    if not ec["enabled"]:
        log.info("Email đang tắt, bỏ qua: %s", subject)
        return False
    if not (ec["username"] and ec["password"] and ec["to"]):
        log.error("Chưa cấu hình email (username/password/to) trong config.ini")
        return False
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = formataddr((ec["from_name"], ec["from_addr"]))
    msg["To"] = ", ".join(ec["to"])
    msg.attach(MIMEText(re.sub(r"<[^>]+>", "", body_html), "plain", "utf-8"))
    msg.attach(MIMEText(body_html, "html", "utf-8"))
    for attempt in range(3):
        try:
            ctx = ssl.create_default_context()
            if ec["use_ssl"]:
                s = smtplib.SMTP_SSL(ec["host"], ec["port"], context=ctx, timeout=30)
            else:
                s = smtplib.SMTP(ec["host"], ec["port"], timeout=30)
                s.starttls(context=ctx)
            _smtp_login(s, ec["username"], ec["password"])
            s.sendmail(ec["from_addr"], ec["to"], msg.as_string())
            s.quit()
            log.info("Đã gửi email: %s -> %s", subject, ", ".join(ec["to"]))
            return True
        except smtplib.SMTPAuthenticationError as ex:
            err = ex.smtp_error.decode(errors="replace") if isinstance(ex.smtp_error, bytes) else str(ex.smtp_error)
            _email_error(f"Sai tên đăng nhập hoặc mật khẩu SMTP ({ex.smtp_code} {' '.join(err.split())}). "
                         "Gmail cần 'Mật khẩu ứng dụng' 16 ký tự, không dùng mật khẩu đăng nhập thường.")
            _email_error("Chẩn đoán: " + describe_credentials(ec))
            return False  # sai mật khẩu thì thử lại cũng vô ích
        except Exception as ex:
            _email_error(f"Gửi email thất bại (lần {attempt + 1}): {ex}")
            time.sleep(3)
    return False


def _smtp_login(s, user, password):
    """Đăng nhập bằng đúng 1 cơ chế. smtplib.login() thử lần lượt nhiều cơ chế, nhưng Gmail
    ngắt kết nối sau lần sai đầu tiên -> lỗi thật (535 sai mật khẩu) bị che thành
    'Connection unexpectedly closed'."""
    s.ehlo_or_helo_if_needed()
    advertised = s.esmtp_features.get("auth", "").upper().split()
    s.user, s.password = user, password
    for mech in ("PLAIN", "LOGIN"):
        if mech in advertised:
            return s.auth(mech, getattr(s, "auth_" + mech.lower()))
    return s.login(user, password)


def _email_error(text):
    log.error(text)
    if os.environ.get("GITHUB_ACTIONS"):  # hiện lỗi ở trang tóm tắt lượt chạy (không cần đăng nhập)
        print(f"::warning::{text}", flush=True)


def _table(rows, header, color):
    th = "".join(f'<th style="padding:8px;text-align:left;background:{color};color:#fff">{h}</th>' for h in header)
    trs = "".join(
        "<tr>" + "".join(f'<td style="padding:8px;border-bottom:1px solid #ddd">{c}</td>' for c in r) + "</tr>"
        for r in rows)
    return f'<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px">' \
           f'<tr>{th}</tr>{trs}</table>'


def email_down(cfg, items, reminder=False):
    rows = [[f'<a href="{html.escape(r["final_url"])}">{html.escape(r["site"])}</a>',
             html.escape(r["error"] or ""), r["checked_at"], r.get("down_for", ""),
             f'{r.get("alert_no", 1)}/{cfg["max_alerts"]}' if cfg["max_alerts"] > 0 else r.get("alert_no", 1)]
            for r in items]
    names = ", ".join(r["site"] for r in items)
    tag = "[VẪN ĐANG LỖI]" if reminder else "[CẢNH BÁO]"
    subject = f"{tag} {len(items)} website không truy cập được: {names}"[:200]
    body = (f'<h2 style="color:#c62828;font-family:Arial">{"⏰ Website vẫn đang lỗi" if reminder else "🔴 Phát hiện website lỗi"}</h2>'
            + _table(rows, ["Website", "Lỗi", "Kiểm tra lúc", "Đã lỗi", "Lần báo"], "#c62828")
            + '<p style="font-family:Arial;color:#666;font-size:12px">Gửi tự động bởi Website Monitor.</p>')
    return send_email(cfg, subject, body)


def email_up(cfg, items):
    rows = [[html.escape(r["site"]), f'{r["code"]} - {r["ms"]} ms', r["checked_at"], r.get("down_for", "")]
            for r in items]
    names = ", ".join(r["site"] for r in items)
    subject = f"[ĐÃ KHÔI PHỤC] {len(items)} website hoạt động lại: {names}"[:200]
    body = ('<h2 style="color:#2e7d32;font-family:Arial">🟢 Website đã hoạt động lại</h2>'
            + _table(rows, ["Website", "Phản hồi", "Kiểm tra lúc", "Thời gian lỗi"], "#2e7d32")
            + '<p style="font-family:Arial;color:#666;font-size:12px">Gửi tự động bởi Website Monitor.</p>')
    return send_email(cfg, subject, body)


# ----------------------------------------------------------------------------- reports
def write_dashboard(results, cfg, cycle_time):
    down = [r for r in results if not r["ok"]]
    rows = []
    for r in sorted(results, key=lambda x: (x["ok"], x["site"])):
        st = '<span class="b ok">HOẠT ĐỘNG</span>' if r["ok"] else '<span class="b err">LỖI</span>'
        rows.append(
            f'<tr class="{"" if r["ok"] else "row-err"}"><td>{st}</td>'
            f'<td><a href="{html.escape(r["final_url"])}" target="_blank">{html.escape(r["site"])}</a></td>'
            f'<td>{r["code"] or "-"}</td><td>{(str(r["ms"]) + " ms") if r["ms"] else "-"}</td>'
            f'<td>{html.escape(r["error"] or "")}</td><td>{html.escape(r.get("since", ""))}</td>'
            f'<td>{r["checked_at"]}</td></tr>')
    page = f"""<!doctype html><html lang="vi"><head><meta charset="utf-8">
<meta http-equiv="refresh" content="{max(cfg['interval'], 10)}">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Website Monitor</title>
<style>
body{{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f4f6f8;color:#222}}
header{{background:#1e293b;color:#fff;padding:16px 24px}} h1{{margin:0;font-size:20px}}
.sum{{display:flex;gap:12px;padding:16px 24px;flex-wrap:wrap}}
.card{{background:#fff;border-radius:8px;padding:12px 20px;box-shadow:0 1px 3px #0002;min-width:120px}}
.card b{{display:block;font-size:28px}} .wrap{{padding:0 24px 24px;overflow-x:auto}}
table{{width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 3px #0002}}
th,td{{padding:9px 12px;text-align:left;border-bottom:1px solid #eee;font-size:14px}} th{{background:#f1f5f9}}
.b{{padding:3px 8px;border-radius:4px;font-size:12px;font-weight:600;color:#fff}} .ok{{background:#16a34a}} .err{{background:#dc2626}}
.row-err{{background:#fef2f2}} a{{color:#2563eb;text-decoration:none}}
</style></head><body>
<header><h1>Website Monitor</h1><div>Cập nhật: {datetime.now():%d/%m/%Y %H:%M:%S} · lượt kiểm tra {cycle_time:.1f}s · tự làm mới mỗi {cfg['interval']}s</div></header>
<div class="sum"><div class="card">Tổng<b>{len(results)}</b></div>
<div class="card" style="color:#16a34a">Hoạt động<b>{len(results) - len(down)}</b></div>
<div class="card" style="color:#dc2626">Lỗi<b>{len(down)}</b></div></div>
<div class="wrap"><table><tr><th>Trạng thái</th><th>Website</th><th>HTTP</th><th>Phản hồi</th><th>Ghi chú / Lỗi</th><th>Lỗi từ</th><th>Kiểm tra lúc</th></tr>
{''.join(rows)}</table></div></body></html>"""
    with open(DASHBOARD_FILE, "w", encoding="utf-8") as f:
        f.write(page)


def write_excel_report(results):
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Trạng thái"
        ws.append(["STT", "Website", "Trạng thái", "HTTP", "Phản hồi (ms)", "Ghi chú / Lỗi", "Lỗi từ", "Kiểm tra lúc"])
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1E293B")
        red, green = PatternFill("solid", fgColor="FDE2E2"), PatternFill("solid", fgColor="E3F6E8")
        for i, r in enumerate(results, 1):
            ws.append([i, r["site"], "HOẠT ĐỘNG" if r["ok"] else "LỖI", r["code"], r["ms"],
                       r["error"], r.get("since", ""), r["checked_at"]])
            for c in ws[ws.max_row]:
                c.fill = green if r["ok"] else red
        for col, w in zip("ABCDEFGH", [6, 32, 13, 8, 14, 60, 20, 20]):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A2"
        wb.save(REPORT_FILE)
    except PermissionError:
        pass  # file báo cáo đang mở trong Excel -> bỏ qua lượt này


def write_github_summary(results):
    """Khi chạy trên GitHub Actions: hiển thị bảng trạng thái ở trang kết quả của lượt chạy."""
    path = os.environ.get("GITHUB_STEP_SUMMARY")
    if not path:
        return
    down = sum(not r["ok"] for r in results)
    lines = [f"### Website Monitor - {datetime.now():%d/%m/%Y %H:%M}",
             f"Tổng **{len(results)}** · Hoạt động **{len(results) - down}** · Lỗi **{down}**", "",
             "| Trạng thái | Website | HTTP | Phản hồi | Ghi chú / Lỗi |", "|---|---|---|---|---|"]
    for r in sorted(results, key=lambda x: (x["ok"], x["site"])):
        lines.append(f'| {"🟢" if r["ok"] else "🔴"} | {r["site"]} | {r["code"] or "-"} | '
                     f'{(str(r["ms"]) + " ms") if r["ms"] else "-"} | {(r["error"] or "").replace("|", "/")} |')
    with open(path, "a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


# ----------------------------------------------------------------------------- main loop
def run_cycle(cfg, state):
    sites = read_sites(cfg)
    if not sites:
        log.warning("Không có website nào trong file Excel")
        return
    t0 = time.time()
    with ThreadPoolExecutor(max_workers=cfg["workers"]) as ex:
        results = list(ex.map(lambda s: check_site(s, cfg), sites))
    now = time.time()
    newly_down, still_down, recovered, history = [], [], [], []

    for r in results:
        key = r["site"].lower()
        st = state.get(key, {"down": False})
        if not r["ok"]:
            if not st["down"]:
                st = {"down": True, "since": now, "last_alert": 0, "alerts": 0}
                history.append([r["checked_at"], r["site"], "LỖI", r["error"]])
                log.warning("LỖI  %-30s %s", r["site"], r["error"])
            st["error"] = r["error"]
            sent = st.get("alerts", 1)
            if sent == 0:  # chưa báo được (lần đầu, hoặc lần trước gửi email thất bại)
                r["alert_no"] = 1
                newly_down.append(r)
            elif (cfg["max_alerts"] <= 0 or sent < cfg["max_alerts"]) and cfg["remind_minutes"] > 0                     and now - st.get("last_alert", 0) >= cfg["remind_minutes"] * 60:
                r["alert_no"] = sent + 1
                still_down.append(r)
            r["down_for"] = fmt_duration(now - st["since"])
            r["since"] = datetime.fromtimestamp(st["since"]).strftime("%d/%m %H:%M:%S")
        else:
            if st["down"]:
                r["down_for"] = fmt_duration(now - st["since"])
                recovered.append(r)
                history.append([r["checked_at"], r["site"], "KHÔI PHỤC", f"sau {r['down_for']}"])
                log.info("OK   %-30s đã hoạt động lại sau %s", r["site"], r["down_for"])
            st = {"down": False}
        state[key] = st

    # xoá site đã bị gỡ khỏi Excel
    for k in list(state):
        if k not in {s.lower() for s in sites}:
            del state[k]
    if history:
        append_history(history)

    # chỉ tính là "đã báo" khi gửi email thành công -> lỗi SMTP thì lượt sau gửi lại
    for items, reminder in ((newly_down, False), (still_down, True)):
        if items and email_down(cfg, items, reminder):
            for r in items:
                st = state[r["site"].lower()]
                st["alerts"], st["last_alert"] = r["alert_no"], now
                if 0 < cfg["max_alerts"] <= r["alert_no"]:
                    log.info("%s: đã gửi đủ %d email lỗi, ngừng nhắc đến khi hoạt động lại",
                             r["site"], cfg["max_alerts"])
    if recovered and cfg["email"]["notify_recovery"]:
        email_up(cfg, recovered)
    save_state(state)

    write_dashboard(results, cfg, time.time() - t0)
    write_excel_report(results)
    write_github_summary(results)
    n_down = sum(not r["ok"] for r in results)
    log.info("Đã kiểm tra %d site trong %.1fs: %d hoạt động, %d lỗi",
             len(results), time.time() - t0, len(results) - n_down, n_down)


def main():
    ap = argparse.ArgumentParser(description="Website Monitor")
    ap.add_argument("--once", action="store_true", help="kiểm tra 1 lượt rồi thoát")
    ap.add_argument("--test-email", action="store_true", help="gửi email thử")
    args = ap.parse_args()

    setup_logging()
    cfg = load_config()

    if args.test_email:
        ok = send_email(cfg, "[TEST] Website Monitor - email thử",
                        "<h3>Cấu hình email hoạt động tốt ✅</h3><p>Bạn sẽ nhận cảnh báo tại địa chỉ này.</p>")
        sys.exit(0 if ok else 1)

    state = load_state()
    log.info("Bắt đầu giám sát - chu kỳ %ds, timeout %ds, file: %s",
             cfg["interval"], cfg["timeout"], cfg["excel_file"])
    while True:
        start = time.time()
        try:
            cfg = load_config()  # cho phép sửa config.ini khi đang chạy
            run_cycle(cfg, state)
        except KeyboardInterrupt:
            raise
        except Exception:
            log.exception("Lỗi trong lượt kiểm tra")
        if args.once:
            break
        time.sleep(max(1, cfg["interval"] - (time.time() - start)))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        log.info("Đã dừng giám sát.")
