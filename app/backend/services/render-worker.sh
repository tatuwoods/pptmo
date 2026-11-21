#!/bin/bash
# pptmo 渲染 worker：由 launchd 常驻（用户 Aqua 会话直接子进程），
# 轮询队列目录，逐个执行 LibreOffice 转换。
# 为什么存在：uvicorn(launchd) 直接 spawn 的 soffice 在 launchd 会话链里
# 会丢失全部 CJK 字形；launchd 直接派生的 bash 进程渲染正常。

QUEUE=/Users/tt06/.agents/skills/pptmo/app/backend/data/render-queue
export PATH=/Users/tt06/opt/node/bin:/usr/local/bin:/usr/bin:/bin
export HOME=/Users/tt06

mkdir -p "$QUEUE"

while true; do
  for task_file in "$QUEUE"/*.task; do
    [ -e "$task_file" ] || continue   # glob 未命中时跳过
    IFS=$'\t' read -r pptx outdir scale < "$task_file"
    [ -z "$pptx" ] && { rm -f "$task_file"; continue; }
    base=$(basename "$task_file" .task)
    prof="$outdir/.worker-profile"
    mkdir -p "$outdir" "$prof"
    /Applications/LibreOffice.app/Contents/MacOS/soffice \
      "-env:UserInstallation=file://$prof" \
      --headless --norestore --nologo --nodefault \
      --convert-to pdf --outdir "$outdir" "$pptx" > "$outdir/.worker-log" 2>&1
    rc=$?
    if [ $rc -eq 0 ] && ls "$outdir"/*.pdf >/dev/null 2>&1; then
      echo "ok" > "$task_file.done"
      rm -f "$task_file"
    else
      echo "fail rc=$rc" > "$task_file.done"
      mv "$task_file" "$task_file.failed"
    fi
  done
  sleep 1
done
