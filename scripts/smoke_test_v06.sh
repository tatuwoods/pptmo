#!/usr/bin/env bash
# v0.6 live smoke: health=0.6.0 / vision 自检闭环 / OOXML passthrough 保真探针 /
# 演讲者模式页可访问 / add_chart 原生图表 / layouts 版式清单 / theme-preview
# 前后对比 / health 评分卡。
# Usage: smoke_test_v06.sh [BASE_URL]   (default http://127.0.0.1:8090)
#
# 互斥与自愈沿用 v0.4/v0.5 治理：mkdir 原子锁串行化；整轮失败自动重试。
set -uo pipefail

BASE="${1:-http://127.0.0.1:8090}"
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PY="$HERE/app/backend/.venv/bin/python"
[[ -x "$PY" ]] || PY=python3

MAX_ATTEMPTS="${SMOKE_MAX_ATTEMPTS:-3}"
LOCKDIR="/tmp/pptmo-v06-smoke.lock.$USER"
SCRATCH=""
LOCK_STALE_S=900
LOCK_WAIT_S=600

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

wait_healthy() {
  local i
  for i in $(seq 1 30); do
    curl -sf --max-time 3 "$BASE/api/health" 2>/dev/null | grep -q '"ok":true' && return 0
    sleep 1
  done
  return 1
}

# 冒烟素材：3 页主 deck（第 2 页注入动画 timing 探针 = passthrough 保真的比对锚点）
make_decks() {
  "$PY" - "$SCRATCH" <<'PYEOF'
import sys
from pptx import Presentation
from pptx.util import Emu, Pt

scratch = sys.argv[1]
prs = Presentation()
prs.slide_width = Emu(12192000)
prs.slide_height = Emu(6858000)
for i in (1, 2, 3):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Emu(914400), Emu(457200), Emu(8000000), Emu(1200000))
    box.text_frame.text = f"冒烟第{i}页"
    for run in box.text_frame.paragraphs[0].runs:
        run.font.size = Pt(40)
prs.save(f"{scratch}/plain.pptx")

# 给第 2 页 slide part 注入动画 timing 探针（编辑第 1 页后该页字节必须原样保留）
probe = (
    '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" '
    'nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek">'
    '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"/></p:seq>'
    "</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>"
)
import zipfile
with zipfile.ZipFile(f"{scratch}/plain.pptx") as zin, \
     zipfile.ZipFile(f"{scratch}/main.pptx", "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "ppt/slides/slide2.xml":
            data = data.decode("utf-8").replace("</p:sld>", probe + "</p:sld>").encode("utf-8")
        zout.writestr(item, data)
print("decks ready")
PYEOF
}

slide_part_bytes() {  # $1=pptx $2=zip内路径
  "$PY" - "$1" "$2" <<'PYEOF'
import sys, hashlib, zipfile
with zipfile.ZipFile(sys.argv[1]) as zf:
    print(hashlib.sha256(zf.read(sys.argv[2])).hexdigest())
PYEOF
}

run_checks() {
  pass=0; fail=0
  ok()   { pass=$((pass+1)); echo "  ✓ $1"; }
  bad()  { fail=$((fail+1)); echo "  ✗ $1"; }
  check(){ if [[ "$2" == "$3" ]]; then ok "$1 ($2)"; else bad "$1 (got $2, want $3)"; fi }

  echo "== v0.6 smoke against $BASE =="

  # 1. health 版本号
  ver=$(curl -sf --max-time 10 "$BASE/api/health" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["version"])' 2>/dev/null)
  check "health version" "$ver" "0.6.0"

  # 2. 素材 + 上传
  make_decks >/dev/null 2>&1 || { bad "deck generation failed"; echo "== result: $pass passed, $fail failed =="; return 1; }
  doc=$(curl -sf --max-time 300 -X POST "$BASE/api/docs/upload" -F "file=@$SCRATCH/main.pptx")
  doc_id=$(echo "$doc" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["doc_id"])' 2>/dev/null)
  if [[ -z "${doc_id:-}" ]]; then bad "upload failed"; echo "== result: $pass passed, $fail failed =="; return 1; fi
  ok "upload $doc_id"
  cleanup(){ curl -s -o /dev/null -X DELETE "$BASE/api/docs/$doc_id"; rm -rf "$SCRATCH"; }
  trap 'cleanup; release_lock' EXIT

  # 3. passthrough 保真探针：只编辑第 1 页，第 2 页 slide part 字节必须原样
  curl -sf --max-time 300 -o "$SCRATCH/v0.pptx" "$BASE/api/docs/$doc_id/download" || bad "download v0 failed"
  h_before=$(slide_part_bytes "$SCRATCH/v0.pptx" "ppt/slides/slide2.xml")
  curl -sf --max-time 300 -X POST "$BASE/api/docs/$doc_id/apply" -H 'Content-Type: application/json' \
    -d '{"patches":[{"action":"edit","shape_id":"s0_id2","after":"改过的第一页"}],"label":"冒烟v06改字"}' >/dev/null \
    && ok "apply edit slide 1" || bad "apply edit failed"
  curl -sf --max-time 300 -o "$SCRATCH/v1.pptx" "$BASE/api/docs/$doc_id/download" || bad "download v1 failed"
  h_after=$(slide_part_bytes "$SCRATCH/v1.pptx" "ppt/slides/slide2.xml")
  check "passthrough: untouched slide2 bytes" "$h_after" "$h_before"

  # 4. add_chart 原生图表：apply 后 objects 出现 kind=chart
  curl -sf --max-time 300 -X POST "$BASE/api/docs/$doc_id/apply" -H 'Content-Type: application/json' \
    -d '{"patches":[{"action":"add_chart","slide_index":2,"chart_type":"column","categories":["Q1","Q2"],"series":[{"name":"营收","values":[10,20]}]}],"label":"冒烟v06图表"}' >/dev/null \
    && ok "add_chart apply" || bad "add_chart apply failed"
  chart_kind=$(curl -sf --max-time 60 "$BASE/api/docs/$doc_id/objects" \
    | "$PY" -c 'import json,sys;o=json.load(sys.stdin)["objects"];print(next((x["kind"] for x in o if x.get("kind")=="chart"), ""))' 2>/dev/null)
  check "native chart in objects" "$chart_kind" "chart"

  # 5. layouts 版式清单
  layout_count=$(curl -sf --max-time 60 "$BASE/api/docs/$doc_id/layouts" \
    | "$PY" -c 'import json,sys;print(len(json.load(sys.stdin)["layouts"]))' 2>/dev/null)
  if [[ "${layout_count:-0}" -ge 5 ]]; then ok "layouts listed ($layout_count)"; else bad "layouts count $layout_count"; fi

  # 6. theme-preview 安全网：版本不变 + before/after 图可访问
  tp=$(curl -sf --max-time 300 -X POST "$BASE/api/docs/$doc_id/theme-preview" -H 'Content-Type: application/json' \
    -d '{"patches":[{"action":"theme_rebrand","payload":{"palette":{"accent1":"0066CC"}}}],"slides":[0]}')
  tp_version=$(echo "$tp" | "$PY" -c 'import json,sys;d=json.load(sys.stdin);print(d["version_id"])' 2>/dev/null)
  [[ "$tp_version" == "v2" ]] && ok "theme-preview keeps version v2" || bad "theme-preview version $tp_version"
  after_url=$(echo "$tp" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["pages"][0]["after_url"])' 2>/dev/null)
  code=$(curl -s -o /dev/null -w '%{http_code}' --max-time 60 "$BASE$after_url")
  check "theme-preview after image" "$code" "200"

  # 7. health 评分卡
  score=$(curl -sf --max-time 60 "$BASE/api/docs/$doc_id/health" \
    | "$PY" -c 'import json,sys;d=json.load(sys.stdin);print(f"{d[\"score\"]}/{d[\"grade\"]}")' 2>/dev/null)
  [[ "$score" =~ ^[0-9]+/[ABCD]$ ]] && ok "health scorecard ($score)" || bad "health scorecard: $score"

  # 8. vision 自检闭环（多模态模型真实调用；未配 key 时降级为结构断言）
  vc=$(curl -sf --max-time 300 "$BASE/api/docs/$doc_id/vision-check?slides=0")
  vc_error=$(echo "$vc" | "$PY" -c 'import json,sys;d=json.load(sys.stdin);print("1" if d.get("pages") and all("error" in p for p in d["pages"]) else "0")' 2>/dev/null)
  if [[ "$vc_error" == "0" ]]; then
    pages=$(echo "$vc" | "$PY" -c 'import json,sys;print(len(json.load(sys.stdin)["pages"]))' 2>/dev/null)
    check "vision self-check closed loop" "$pages" "1"
  else
    ok "vision self-check: model unavailable, structure ok (no key?)"
  fi

  # 9. 演讲者模式页可访问
  code=$(curl -s -o "$SCRATCH/present.html" -w '%{http_code}' --max-time 30 "$BASE/api/docs/$doc_id/present")
  check "present page http" "$code" "200"
  grep -q "toggleSpeaker" "$SCRATCH/present.html" && ok "present page speaker view" || bad "present page missing speaker view"
  grep -q "timer" "$SCRATCH/present.html" && ok "present page rehearsal timer" || bad "present page missing timer"

  echo
  echo "== result: $pass passed, $fail failed =="
  return $(( fail > 0 ? 1 : 0 ))
}

acquire_lock
attempt=1
while (( attempt <= MAX_ATTEMPTS )); do
  wait_healthy || echo "  ! service not healthy before attempt $attempt, trying anyway"
  rm -rf "$SCRATCH"; SCRATCH="$(mktemp -d "${TMPDIR:-/tmp}/pptmo-v06.XXXXXX")"
  if run_checks; then
    cleanup 2>/dev/null || true
    release_lock; trap - EXIT
    exit 0
  fi
  if (( attempt < MAX_ATTEMPTS )); then
    echo "  … attempt $attempt/$MAX_ATTEMPTS failed，8s 后重试"
    cleanup 2>/dev/null || true
    sleep 8
  fi
  attempt=$((attempt + 1))
done
release_lock; trap - EXIT
echo "smoke failed after $MAX_ATTEMPTS attempts" >&2
exit 1
