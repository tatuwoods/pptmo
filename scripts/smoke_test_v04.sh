#!/usr/bin/env bash
# v0.4 live smoke: health=0.4.0 / share 链接 / outline 导出 / compare diff / undo。
# Usage: smoke_test_v04.sh [BASE_URL]   (default http://127.0.0.1:8090)
set -uo pipefail

BASE="${1:-http://127.0.0.1:8090}"
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DEMO="$HERE/app/samples/demo.pptx"
PY="$HERE/app/backend/.venv/bin/python"
[[ -x "$PY" ]] || PY=python3

pass=0; fail=0
ok()   { pass=$((pass+1)); echo "  ✓ $1"; }
bad()  { fail=$((fail+1)); echo "  ✗ $1"; }
check(){ if [[ "$2" == "$3" ]]; then ok "$1 ($2)"; else bad "$1 (got $2, want $3)"; fi }

echo "== v0.4 smoke against $BASE =="

# 1. health 版本号
ver=$(curl -sf --max-time 10 "$BASE/api/health" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["version"])' 2>/dev/null)
check "health version" "$ver" "0.4.0"

# 2. 上传 + 改标题产生 v1（后续 compare/undo 的素材）
doc=$(curl -sf --max-time 120 -X POST "$BASE/api/docs/upload" -F "file=@$DEMO")
doc_id=$(echo "$doc" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["doc_id"])' 2>/dev/null)
[[ -n "${doc_id:-}" ]] && ok "upload $doc_id" || { bad "upload failed"; exit 1; }
cleanup(){ curl -s -o /dev/null -X DELETE "$BASE/api/docs/$doc_id"; }
trap cleanup EXIT

target=$(echo "$doc" | "$PY" -c '
import json,sys
d=json.load(sys.stdin)
o=[x for x in d["objects"] if x["kind"]=="text" and "演示文稿标题" in x["text"]]
print(json.dumps(o[0]) if o else "")')
if [[ -n "$target" ]]; then
  patch=$("$PY" -c "
import json
o=json.loads('''$target''')
print(json.dumps({'patches':[{'action':'edit','shape_id':o['shape_id'],'after':'冒烟测试标题','before':o['text']}],'label':'冒烟v04','base_version':0}))")
  vid=$(curl -sf --max-time 120 -X POST "$BASE/api/docs/$doc_id/apply" -H 'Content-Type: application/json' -d "$patch" \
    | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["version_id"])' 2>/dev/null)
  check "apply -> v1" "$vid" "v1"
else
  bad "title object not found"; exit 1
fi

# 3. share：建链接 → 页面/图/下载可访问 → 撤销后失效
share=$(curl -sf -X POST "$BASE/api/docs/$doc_id/share")
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
hdr=$(curl -sD - -o /tmp/pptmo-v04-outline.txt --max-time 30 "$BASE/api/docs/$doc_id/export/outline")
grep -qi "content-type: text/plain" <<<"$hdr" && ok "outline content-type text/plain" || bad "outline content-type wrong"
grep -q "outline.txt" <<<"$hdr" && ok "outline filename header" || bad "outline filename header missing"
grep -q "冒烟测试标题" /tmp/pptmo-v04-outline.txt && ok "outline contains edited title" || bad "outline missing edited title"
head -1 /tmp/pptmo-v04-outline.txt | grep -q "^# " && ok "outline starts with # title" || bad "outline heading malformed"
rm -f /tmp/pptmo-v04-outline.txt

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
print(d["summary"])' >/tmp/pptmo-v04-diff.out 2>/tmp/pptmo-v04-diff.err \
  && ok "compare diff: $(cat /tmp/pptmo-v04-diff.out)" || bad "compare diff wrong: $(cat /tmp/pptmo-v04-diff.err)"
rm -f /tmp/pptmo-v04-diff.out /tmp/pptmo-v04-diff.err

# 6. undo 仍工作：回到 v0 内容，版本前进到 v2
undo=$(curl -sf --max-time 120 -X POST "$BASE/api/docs/$doc_id/undo" -H 'Content-Type: application/json' -d '{"base_version":1}')
echo "$undo" | "$PY" -c '
import json,sys
d=json.load(sys.stdin)
assert d.get("restored_from")=="v0", d
print(d.get("version_id"))' >/tmp/pptmo-v04-undo.out 2>/tmp/pptmo-v04-undo.err \
  && ok "undo restored v0 -> $(cat /tmp/pptmo-v04-undo.out)" || bad "undo failed: $(cat /tmp/pptmo-v04-undo.err)"
curl -sf --max-time 30 "$BASE/api/docs/$doc_id/export/outline" | grep -q "演示文稿标题" \
  && ok "outline after undo has original title" || bad "undo did not restore original title"
rm -f /tmp/pptmo-v04-undo.out /tmp/pptmo-v04-undo.err

# 7. audit 记录了 apply / undo
curl -sf --max-time 10 "$BASE/api/audit" | "$PY" -c '
import json,sys
d=json.load(sys.stdin)
acts={e.get("action") for e in (d.get("entries") or d.get("items") or [])}
assert "apply" in acts and "undo" in acts, acts' 2>/dev/null \
  && ok "audit trail has apply+undo" || bad "audit missing apply/undo"

echo
echo "== result: $pass passed, $fail failed =="
exit $(( fail > 0 ? 1 : 0 ))
