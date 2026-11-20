#!/usr/bin/env bash
# Live end-to-end smoke test against a running ppt-light server (real LibreOffice render).
# Usage: smoke_test.sh [BASE_URL]   (default http://127.0.0.1:8090)
# Requires: curl, python3, and the repo's samples/demo.pptx.
set -uo pipefail

BASE="${1:-http://127.0.0.1:8090}"
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DEMO="$HERE/app/samples/demo.pptx"
PY="$(cd "$HERE/app/backend" 2>/dev/null && ls .venv/bin/python 2>/dev/null || command -v python3)"
[[ -x "$PY" ]] || PY=python3

pass=0; fail=0
ok()   { pass=$((pass+1)); echo "  ✓ $1"; }
bad()  { fail=$((fail+1)); echo "  ✗ $1"; }
check(){ if [[ "$2" == "$3" ]]; then ok "$1 ($2)"; else bad "$1 (got $2, want $3)"; fi; }

echo "== smoke against $BASE =="

# 1. health
health=$(curl -sf --max-time 10 "$BASE/api/health") && ok "health reachable" || bad "health unreachable"
echo "$health" | "$PY" -c 'import json,sys;d=json.load(sys.stdin);exit(0 if d.get("ok") else 1)' \
  && ok "health ok=true" || bad "health ok!=true"

# 2. upload
doc=$(curl -sf --max-time 120 -X POST "$BASE/api/docs/upload" -F "file=@$DEMO")
doc_id=$(echo "$doc" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["doc_id"])' 2>/dev/null)
[[ -n "${doc_id:-}" ]] && ok "upload created $doc_id" || { bad "upload failed"; exit 1; }

cleanup(){ curl -s -o /dev/null -X DELETE "$BASE/api/docs/$doc_id"; }
trap cleanup EXIT

slides=$(echo "$doc" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["slide_count"])')
check "slide_count=12" "$slides" "12"

# 3. real render pipeline (LibreOffice must have produced a PNG)
code=$(curl -s -o /dev/null -w '%{http_code}' --max-time 60 "$BASE/api/docs/$doc_id/slides/0/image")
check "rendered slide image" "$code" "200"
size=$(curl -s --max-time 60 "$BASE/api/docs/$doc_id/slides/0/image" | wc -c | tr -d ' ')
[[ "$size" -gt 5000 ]] && ok "image non-trivial (${size}B)" || bad "image suspiciously small (${size}B)"

# 4. manual apply: edit slide-1 title text
target=$(echo "$doc" | "$PY" -c '
import json,sys
d=json.load(sys.stdin)
o=[x for x in d["objects"] if x["kind"]=="text" and "演示文稿标题" in x["text"]]
print(json.dumps(o[0]) if o else "")')
if [[ -n "$target" ]]; then
  patch=$("$PY" -c "
import json,sys
o=json.loads('''$target''')
print(json.dumps({'patches':[{'action':'edit','shape_id':o['shape_id'],'after':'冒烟测试标题','before':o['text']}],'label':'冒烟测试','base_version':0}))")
  resp=$(curl -sf --max-time 120 -X POST "$BASE/api/docs/$doc_id/apply" -H 'Content-Type: application/json' -d "$patch")
  vid=$(echo "$resp" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["version_id"])' 2>/dev/null)
  check "apply produced v1" "$vid" "v1"
else
  bad "could not find title object for apply test"
fi

# 5. conflict guard
code=$(curl -s -o /dev/null -w '%{http_code}' -X POST "$BASE/api/docs/$doc_id/apply" \
  -H 'Content-Type: application/json' -d '{"patches":[{"action":"edit","shape_id":"s0_id4","after":"x"}],"base_version":0}')
check "stale base_version -> 409" "$code" "409"

# 6. versions + restore
count=$(curl -sf "$BASE/api/docs/$doc_id/versions" | "$PY" -c 'import json,sys;print(len(json.load(sys.stdin)["versions"]))')
check "2 versions recorded" "$count" "2"
newv=$(curl -sf -X POST "$BASE/api/docs/$doc_id/versions/v0/restore" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["version_id"])')
check "restore v0 -> v2" "$newv" "v2"

# 7. messages trail
kinds=$(curl -sf "$BASE/api/docs/$doc_id/messages" | "$PY" -c '
import json,sys
ms=json.load(sys.stdin)["messages"]
print(",".join(sorted({m["kind"] for m in ms})))')
echo "$kinds" | grep -q "apply" && ok "apply message in trail" || bad "no apply message"
echo "$kinds" | grep -q "restore" && ok "restore message in trail" || bad "no restore message"

# 8. download
hdr=$(curl -sf --max-time 60 -o /tmp/pptmo-smoke-dl.pptx -w '%{content_type}' "$BASE/api/docs/$doc_id/download")
head -c2 /tmp/pptmo-smoke-dl.pptx | grep -q "PK" && ok "download is a zip" || bad "download not zip"
rm -f /tmp/pptmo-smoke-dl.pptx

# 9. AI chat (only if key configured)
ai_on=$(echo "$health" | "$PY" -c 'import json,sys;print(json.load(sys.stdin).get("ai_configured",False))')
if [[ "$ai_on" == "True" ]]; then
  echo "-- AI chat (glm streaming, may take ~30s) --"
  stream=$(curl -sN --max-time 240 -X POST "$BASE/api/docs/$doc_id/chat" \
    -H 'Content-Type: application/json' \
    -d '{"message":"第一页的标题是什么？只回答不要修改。","scope":"current","slide_index":0,"base_version":2}')
  echo "$stream" | grep -q "^event: start" && ok "SSE start" || bad "no SSE start"
  echo "$stream" | grep -q "tool_call" && ok "agent called a tool" || bad "no tool_call seen"
  echo "$stream" | grep -q "^event: done" && ok "SSE done" || bad "no SSE done"
  echo "$stream" | grep -q "^event: applied" && bad "read-only chat unexpectedly applied" || ok "no write for read-only ask"
  versions=$(curl -sf "$BASE/api/docs/$doc_id/versions" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["current"])')
  check "no new version after read-only chat" "$versions" "v2"
else
  echo "-- AI not configured; skipping chat test --"
fi

# 10. delete cleanup path (fresh doc, then delete)
doc2=$(curl -sf -X POST "$BASE/api/docs/upload" -F "file=@$DEMO" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["doc_id"])')
code=$(curl -s -o /dev/null -w '%{http_code}' -X DELETE "$BASE/api/docs/$doc2")
check "delete second doc" "$code" "200"
code=$(curl -s -o /dev/null -w '%{http_code}' "$BASE/api/docs/$doc2")
check "deleted doc gone" "$code" "404"

echo
echo "== result: $pass passed, $fail failed =="
exit $(( fail > 0 ? 1 : 0 ))
