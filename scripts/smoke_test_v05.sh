#!/usr/bin/env bash
# v0.5 live smoke: theme_rebrand/recolor 换色后 PNG 色值验证 / merge 拉页页数 /
# 隐藏页导出 PDF 少一页 / watch SSE 推送 / health=0.5.0。
# Usage: smoke_test_v05.sh [BASE_URL]   (default http://127.0.0.1:8090)
#
# 互斥与自愈沿用 v0.4 治理（多会话共用服务时的冒烟竞态）：
#   - mkdir 原子锁串行化并发冒烟；整轮失败自动重试（SMOKE_MAX_ATTEMPTS 可调），
#     轮间等待 /api/health 恢复；每轮 mktemp 独占目录。
set -uo pipefail

BASE="${1:-http://127.0.0.1:8090}"
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PY="$HERE/app/backend/.venv/bin/python"
[[ -x "$PY" ]] || PY=python3

MAX_ATTEMPTS="${SMOKE_MAX_ATTEMPTS:-3}"
LOCKDIR="/tmp/pptmo-v05-smoke.lock.$USER"
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

# 生成冒烟素材：主 deck 3 页（红方块=硬编码色，accent1 方块=主题引用色），源 deck 1 页
make_decks() {
  "$PY" - "$SCRATCH" <<'PYEOF'
import sys
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

scratch = sys.argv[1]

def new_deck():
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    return prs

def add_text(slide, text, size=40):
    box = slide.shapes.add_textbox(Emu(914400), Emu(457200), Emu(8000000), Emu(1200000))
    run = box.text_frame.paragraphs[0].add_run()
    run.text = text
    run.font.size = Pt(size)
    return box

prs = new_deck()
s1 = prs.slides.add_slide(prs.slide_layouts[6])
add_text(s1, "冒烟第一页")
# 左：硬编码红方块（recolor 的目标）
red = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(914400), Emu(2286000), Emu(3657600), Emu(2286000))
red.fill.solid()
red.fill.fore_color.rgb = RGBColor.from_string("FF0000")
red.line.fill.background()
# 右：accent1 主题引用方块（theme_rebrand 的目标）
acc = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(6400800), Emu(2286000), Emu(3657600), Emu(2286000))
acc.line.fill.background()
spPr = acc._element.spPr
fill = etree.SubElement(spPr, qn("a:solidFill"))
scheme = etree.SubElement(fill, qn("a:schemeClr"))
scheme.set("val", "accent1")
prs.save(f"{scratch}/main.pptx")

src = new_deck()
s = src.slides.add_slide(src.slide_layouts[6])
add_text(s, "合并进来的页")
src.save(f"{scratch}/src.pptx")
print("decks ready")
PYEOF
}

# 像素校验：PNG 中相对坐标 (rx, ry) 处的颜色是否≈want（每通道容差 tol）
pixel_is() {  # $1=png $2=rx $3=ry $4=want_hex $5=tol
  "$PY" - "$1" "$2" "$3" "$4" "$5" <<'PYEOF'
import sys
from PIL import Image

png, rx, ry, want, tol = sys.argv[1], float(sys.argv[2]), float(sys.argv[3]), sys.argv[4], int(sys.argv[5])
img = Image.open(png).convert("RGB")
x, y = int(img.width * rx), int(img.height * ry)
r, g, b = img.getpixel((x, y))
wr, wg, wb = int(want[0:2], 16), int(want[2:4], 16), int(want[4:6], 16)
ok = abs(r - wr) <= tol and abs(g - wg) <= tol and abs(b - wb) <= tol
print("ok" if ok else f"bad: got #{r:02X}{g:02X}{b:02X} want #{want}")
sys.exit(0 if ok else 1)
PYEOF
}

pdf_pages() {  # $1=pdf path
  "$PY" - "$1" <<'PYEOF'
import sys, pymupdf
print(pymupdf.open(sys.argv[1]).page_count)
PYEOF
}

run_checks() {
  pass=0; fail=0
  ok()   { pass=$((pass+1)); echo "  ✓ $1"; }
  bad()  { fail=$((fail+1)); echo "  ✗ $1"; }
  check(){ if [[ "$2" == "$3" ]]; then ok "$1 ($2)"; else bad "$1 (got $2, want $3)"; fi }

  echo "== v0.5 smoke against $BASE =="

  # 1. health 版本号
  ver=$(curl -sf --max-time 10 "$BASE/api/health" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["version"])' 2>/dev/null)
  check "health version" "$ver" "0.5.0"

  # 2. 素材 + 上传
  make_decks >/dev/null 2>&1 || { bad "deck generation failed"; echo "== result: $pass passed, $fail failed =="; return 1; }
  doc=$(curl -sf --max-time 300 -X POST "$BASE/api/docs/upload" -F "file=@$SCRATCH/main.pptx")
  doc_id=$(echo "$doc" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["doc_id"])' 2>/dev/null)
  if [[ -z "${doc_id:-}" ]]; then bad "upload failed"; echo "== result: $pass passed, $fail failed =="; return 1; fi
  ok "upload $doc_id"
  cleanup(){ curl -s -o /dev/null -X DELETE "$BASE/api/docs/$doc_id"; rm -rf "$SCRATCH"; }
  trap 'cleanup; release_lock' EXIT

  # 3. recolor 硬编码色：红方块 → 0066CC，渲染 PNG 色值验证
  vid=$(curl -sf --max-time 300 -X POST "$BASE/api/docs/$doc_id/apply" -H 'Content-Type: application/json' \
    -d '{"patches":[{"action":"recolor","payload":{"mapping":{"FF0000":"0066CC"}}}],"label":"冒烟v05换色"}' \
    | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["version_id"])' 2>/dev/null)
  check "recolor apply -> v1" "$vid" "v1"
  curl -sf --max-time 300 -o "$SCRATCH/p1.png" "$BASE/api/docs/$doc_id/slides/0/image" || bad "slide image fetch failed"
  out=$(pixel_is "$SCRATCH/p1.png" 0.20 0.50 0066CC 12 2>&1)
  [[ "$out" == "ok" ]] && ok "recolor: red block -> #0066CC in PNG" || bad "recolor pixel check: $out"

  # 4. theme_rebrand 主题引用色：accent1 方块换色（XML 无硬编码，随主题变）
  vid=$(curl -sf --max-time 300 -X POST "$BASE/api/docs/$doc_id/apply" -H 'Content-Type: application/json' \
    -d '{"patches":[{"action":"theme_rebrand","payload":{"palette":{"accent1":"E94560"}}}],"label":"冒烟v05主题"}' \
    | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["version_id"])' 2>/dev/null)
  check "theme_rebrand apply -> v2" "$vid" "v2"
  curl -sf --max-time 300 -o "$SCRATCH/p2.png" "$BASE/api/docs/$doc_id/slides/0/image" || bad "slide image fetch failed"
  out=$(pixel_is "$SCRATCH/p2.png" 0.72 0.50 E94560 12 2>&1)
  [[ "$out" == "ok" ]] && ok "theme_rebrand: accent1 block -> #E94560 in PNG" || bad "theme_rebrand pixel check: $out"

  # 5. merge 拉页：1 页源 deck → 主文档 3→4 页
  merge=$(curl -sf --max-time 300 -X POST "$BASE/api/docs/$doc_id/merge" \
    -F "file=@$SCRATCH/src.pptx" -F "pages=1" -F "position=3")
  slide_count=$(echo "$merge" | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["slide_count"])' 2>/dev/null)
  check "merge page count 3+1" "$slide_count" "4"

  # 6. 隐藏页 + 导出联动：隐藏第 4 页 → PDF 3 页；include_hidden=1 → 4 页
  vid=$(curl -sf --max-time 300 -X POST "$BASE/api/docs/$doc_id/apply" -H 'Content-Type: application/json' \
    -d '{"patches":[{"action":"hide_slide","payload":{"pages":"4"}}],"label":"冒烟v05隐藏"}' \
    | "$PY" -c 'import json,sys;print(json.load(sys.stdin)["version_id"])' 2>/dev/null)
  check "hide_slide apply -> v4" "$vid" "v4"
  curl -sf --max-time 300 -o "$SCRATCH/out.pdf" "$BASE/api/docs/$doc_id/export/pdf" || bad "export pdf failed"
  check "hidden page skipped in PDF" "$(pdf_pages "$SCRATCH/out.pdf" 2>/dev/null)" "3"
  curl -sf --max-time 300 -o "$SCRATCH/out-all.pdf" "$BASE/api/docs/$doc_id/export/pdf?include_hidden=1" || bad "export pdf (include_hidden) failed"
  check "include_hidden keeps all pages" "$(pdf_pages "$SCRATCH/out-all.pdf" 2>/dev/null)" "4"

  # 7. watch SSE：连上即 hello；触发一次 apply 后收到 changed
  curl -sN --max-time 20 "$BASE/api/docs/$doc_id/events" >"$SCRATCH/sse.log" 2>/dev/null &
  sse_pid=$!
  sleep 2
  curl -sf --max-time 300 -X POST "$BASE/api/docs/$doc_id/apply" -H 'Content-Type: application/json' \
    -d '{"patches":[{"action":"show_slide","payload":{"pages":"4"}}],"label":"冒烟v05取消隐藏"}' >/dev/null
  sleep 3
  kill "$sse_pid" 2>/dev/null; wait "$sse_pid" 2>/dev/null
  grep -q "event: hello" "$SCRATCH/sse.log" && ok "watch SSE hello event" || bad "watch SSE missing hello"
  grep -q "event: changed" "$SCRATCH/sse.log" && ok "watch SSE changed event on apply" || bad "watch SSE missing changed"
  grep -q '"version": *5' "$SCRATCH/sse.log" && ok "watch SSE carries version 5" || bad "watch SSE version payload wrong"

  echo
  echo "== result: $pass passed, $fail failed =="
  return $(( fail > 0 ? 1 : 0 ))
}

acquire_lock
attempt=1
while (( attempt <= MAX_ATTEMPTS )); do
  wait_healthy || echo "  ! service not healthy before attempt $attempt, trying anyway"
  rm -rf "$SCRATCH"; SCRATCH="$(mktemp -d "${TMPDIR:-/tmp}/pptmo-v05.XXXXXX")"
  if run_checks; then
    cleanup 2>/dev/null || true
    release_lock; trap - EXIT
    exit 0
  fi
  if (( attempt < MAX_ATTEMPTS )); then
    echo "  … attempt $attempt/$MAX_ATTEMPTS failed，8s 后重试"
    sleep 8
    attempt=$((attempt + 1))
  else
    break
  fi
done
cleanup 2>/dev/null || true
release_lock; trap - EXIT
exit 1
