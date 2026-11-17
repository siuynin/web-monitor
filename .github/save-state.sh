#!/usr/bin/env bash
# Lưu state.json / history.csv về repo (gọi từ workflow và từ monitor.py khi có sự cố/khôi phục)
cd "$(git rev-parse --show-toplevel)" || exit 0
git config user.name "website-monitor[bot]"
git config user.email "41898282+github-actions[bot]@users.noreply.github.com"
for f in state.json history.csv; do [ -f "$f" ] && git add "$f"; done
if ! git diff --cached --quiet; then
  git commit -q -m "Cập nhật trạng thái giám sát"
elif [ $(( $(date +%s) - $(git log -1 --format=%ct) )) -gt $((50*86400)) ]; then
  # GitHub tắt lịch chạy nếu repo không có hoạt động 60 ngày -> commit giữ kết nối
  git commit -q --allow-empty -m "Keepalive: giữ lịch giám sát hoạt động"
else
  exit 0
fi
BR="${GITHUB_REF_NAME:-main}"
for i in 1 2 3; do
  git pull --rebase --autostash -q origin "$BR" && git push -q origin "HEAD:$BR" && exit 0
  sleep 5
done
echo "::warning::Không lưu được trạng thái lên repo"
exit 1
