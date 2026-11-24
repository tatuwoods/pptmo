#!/usr/bin/env bash
# v0.4 live smoke: health=0.4.0 / share 链接 / outline 导出 / compare diff / undo。
# Usage: smoke_test_v04.sh [BASE_URL]   (default http://127.0.0.1:8090)
#
# 互斥与自愈（多会话/多进程共用同一服务时的冒烟竞态治理）：
#   - mkdir 原子锁串行化并发冒烟（此前多次会话各自 kickstart+冒烟互相打断，
#     表现为 000/空响应的"间歇性超时"，曾误判为渲染负载超时）；
#   - 整轮失败自动重试（默认 3 次，SMOKE_MAX_ATTEMPTS 可调），
#     轮间等待 /api/health 恢复，吸收外部 kickstart 重启窗口。
set -uo pipefail

BASE="${1:-http://127.0.0.1:8090}"
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DEMO="$HERE/app/samples/demo.pptx"
PY="$HERE/app/backend/.venv/bin/python"
[[ -x "$PY" ]] || PY=python3

MAX_ATTEMPTS="${SMOKE_MAX_ATTEMPTS:-3}"
LOCKDIR="/tmp/pptmo-v04-smoke.lock.$USER"
# 每轮独占临时目录：锁等待超时会降级为无锁继续，此路径下固定文件名仍会互删
SCRATCH=""
LOCK_STALE_S=900   # 锁持有超过 15 分钟视为残留（持有进程被杀），自动破锁
LOCK_WAIT_S=600    # 最多等锁 10 分钟，超时降级为警告后继续（不让兄弟会话永久卡死）

acquire_lock() {
  local waited=0 age
  while ! mkdir "$LOCKDIR" 2>/dev/null; do
    age=$(( $(date +%s) - $(stat -f %m "$LOCKDIR" 2>/dev/null || echo 0) ))
    if (( age > LOCK_STALE_S )); then
      echo "  ! breaking stale smoke lock (age ${age}s)" >&2
      rm -rf "$LOCKDIR"
      continue
    fi
    if (( waited >= LOCK_WAIT_S )); then
      echo "  ! smoke lock busy ${waited}s, proceeding without lock" >&2
      return 1
    fi
    sleep 5; waited=$((waited + 5))
  done
  date +%s > "$LOCKDIR/.acquired"
  return 0
}

release_lock() { rm -rf "$LOCKDIR" 2>/dev/null || true; }

wait_healthy() {  # 等服务就绪（吸收刚发生的 kickstart 重启窗口）
  local i
  for i in $(seq 1 30); do
    curl -sf --max-time 3 "$BASE/api/health" 2>/dev/null | grep -q '"ok":true' && return 0
    sleep 1
  done
  return 1
}

run_checks() {
  pass=0; fail=0
  ok()   { pass=$((pass+1)); echo "  ✓ $1"; }
  bad()  { fail=$((fail+1)); echo "  ✗ $1"; }
  check(){ if [[ "$2" == "$3" ]]; then ok "$1 ($2)"; else bad "$1 (got $2, want $3)"; fi }

  echo "== v0.4 smoke against $BASE =="

  # 1. health 版本号
  ver=$(curl -sf --max-time 10 "$BASE/api/health" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["version"])' 2>/dev/null)
  check "health version" "$ver" "0.4.0"

  # 2. 上传 + 改标题产生 v1（后续 compare/undo 的素材）
  doc=$(curl -sf --max-time 300 -X POST "$BASE/api/docs/upload" -F "file=@$DEMO")
  doc_id=$(echo "$doc" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["doc_id"])' 2>/dev/null)
  if [[ -z "${doc_id:-}" ]]; then bad "upload failed"; echo "== result: $pass passed, $fail failed =="; return 1; fi
  ok "upload $doc_id"
  rm -rf "$SCRATCH"; SCRATCH="$(mktemp -d "${TMPDIR:-/tmp}/pptmo-v04.XXXXXX")"
  cleanup(){ curl -s -o /dev/null -X DELETE "$BASE/api/docs/$doc_id"; rm -rf "$SCRATCH"; }
  trap 'cleanup; release_lock' EXIT

  target=$(echo "$doc" | "$PY" -c '
import json,sys
d=json.load(sys.stdin)
o=[x for x in d["objects"] if x["kind"]=="text" and "演示文稿标题" in x["text"]]
print(json.dumps(o[0]) if o else "")')
  if [[ -z "$target" ]]; then
    bad "title object not found"; echo "== result: $pass passed, $fail failed =="; return 1
  fi
  patch=$("$PY" -c "
import json
o=json.loads('''$target''')
print(json.dumps({'patches':[{'action':'edit','shape_id':o['shape_id'],'after':'冒烟测试标题','before':o['text']}],'label':'冒烟v04','base_version':0}))")
  vid=$(curl -sf --max-time 300 -X POST "$BASE/api/docs/$doc_id/apply" -H 'Content-Type: application/json' -d "$patch" \
    | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["version_id"])' 2>/dev/null)
  check "apply -> v1" "$vid" "v1"

  # 3. share：建链接 → 页面/图/下载可访问 → 撤销后失效
  share=$(curl -sf --max-time 60 -X POST "$BASE/api/docs/$doc_id/share")
  token=$(echo "$share" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["token"])' 2>/dev/null)
  url=$(echo "$share" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["share_url"])' 2>/dev/null)
  [[ "$url" == "/share/$token" ]] && ok "share_url=$url" || bad "share_url malformed ($url)"
  code=$(curl -s -o /dev/null -w '%{http_code}' --max-time 30 "$BASE/share/$token")
  check "share page 200" "$code" "200"
  code=$(curl -s -o /dev/null -w '%{http_code}' --max-time 60 "$BASE/share/$token/slides/0/image")
  check "share slide image 200" "$code" "200"
  code=$(curl -s -o /dev/null -w '%{http_code}' --max-time 60 "$BASE/share/$token/download")
  check "share download 200" "$code" "200"
  # 只读：share 路径下无 apply 编辑入口（前端页面不带编辑 API 调用，验证后端仅暴露读端点）
  code=$(curl -s -o /dev/null -w '%{http_code}' -X POST "$BASE/share/$token/apply" -H 'Content-Type: application/json' -d '{}')
  [[ "$code" != "200" ]] && ok "share has no apply endpoint ($code)" || bad "share unexpectedly editable"
  code=$(curl -s -o /dev/null -w '%{http_code}' -X DELETE "$BASE/api/docs/$doc_id/share/$token")
  check "revoke share 200" "$code" "200"
  code=$(curl -s -o /dev/null -w '%{http_code}' "$BASE/share/$token")
  [[ "$code" == "404" || "$code" == "410" ]] && ok "revoked share gone ($code)" || bad "revoked share still $code"

  # 4. outline 导出：text/plain、含改后标题、Content-Disposition
  hdr=$(curl -sD - -o "$SCRATCH/outline.txt" --max-time 30 "$BASE/api/docs/$doc_id/export/outline")
  grep -qi "content-type: text/plain" <<<"$hdr" && ok "outline content-type text/plain" || bad "outline content-type wrong"
  grep -q "outline.txt" <<<"$hdr" && ok "outline filename header" || bad "outline filename header missing"
  grep -q "冒烟测试标题" "$SCRATCH/outline.txt" && ok "outline contains edited title" || bad "outline missing edited title"
  head -1 "$SCRATCH/outline.txt" | grep -q "^# " && ok "outline starts with # title" || bad "outline heading malformed"

  # 5. compare：v0 vs v1 返回 diff JSON
  diff=$(curl -sf --max-time 30 "$BASE/api/docs/$doc_id/versions/v0/compare/v1")
  echo "$diff" | "$PY" -c '
import json,sys
d=json.load(sys.stdin)
assert len(d["changed"])>=1, "no changed pages"
p=d["changed"][0]
assert p["page"]==1, "first change on page %s" % p["page"]
flat=" ".join(p["additions"])
assert "冒烟测试标题" in flat, "edited title not in additions: %s" % flat
assert any("演示文稿标题" in x for x in p["deletions"]), "original title not in deletions"
print(d["summary"])' >"$SCRATCH/diff.out" 2>"$SCRATCH/diff.err" \
    && ok "compare diff: $(cat "$SCRATCH/diff.out")" || bad "compare diff wrong: $(cat "$SCRATCH/diff.err")"

  # 6. undo 仍工作：回到 v0 内容，版本前进到 v2
  undo=$(curl -sf --max-time 300 -X POST "$BASE/api/docs/$doc_id/undo" -H 'Content-Type: application/json' -d '{"base_version":1}')
  echo "$undo" | "$PY" -c '
import json,sys
d=json.load(sys.stdin)
assert d.get("restored_from")=="v0", d
print(d.get("version_id"))' >"$SCRATCH/undo.out" 2>"$SCRATCH/undo.err" \
    && ok "undo restored v0 -> $(cat "$SCRATCH/undo.out")" || bad "undo failed: $(cat "$SCRATCH/undo.err")"
  curl -sf --max-time 30 "$BASE/api/docs/$doc_id/export/outline" | grep -q "演示文稿标题" \
    && ok "outline after undo has original title" || bad "undo did not restore original title"

  # 7. audit 记录了 apply / undo
  curl -sf --max-time 10 "$BASE/api/audit" | "$PY" -c '
import json,sys
d=json.load(sys.stdin)
acts={e.get("action") for e in (d.get("entries") or d.get("items") or [])}
assert "apply" in acts and "undo" in acts, acts' 2>/dev/null \
    && ok "audit trail has apply+undo" || bad "audit missing apply/undo"

  echo
  echo "== result: $pass passed, $fail failed =="
  return $(( fail > 0 ? 1 : 0 ))
}

acquire_lock
attempt=1
while (( attempt <= MAX_ATTEMPTS )); do
  wait_healthy || echo "  ! service not healthy before attempt $attempt, trying anyway"
  if run_checks; then
    cleanup 2>/dev/null || true
    release_lock; trap - EXIT
    exit 0
  fi
  if (( attempt < MAX_ATTEMPTS )); then
    echo "  … attempt $attempt/$MAX_ATTEMPTS failed (并发 kickstart 竞态可致单轮失败)，8s 后重试"
    sleep 8
    attempt=$((attempt + 1))
  else
    break
  fi
done
cleanup 2>/dev/null || true
release_lock; trap - EXIT
exit 1
