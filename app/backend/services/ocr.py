"""v0.6 #4：本机 OCR —— macOS Vision framework via pyobjc。

对标 ningzimu 的「图片→可编辑 PPT」：导入扫描件/截图时先用系统 Live Text
能力提取文字行（含行框位置与置信度），再由 import-images 在图片上层生成
原生文本框——用户可以直接改字，图片保留作背景层。全程本机推理，无网络依赖。

pyobjc 缺失（非 macOS / 未装依赖）时 ocr_available() 返回 False，导入行为
退回 v0.4 纯图片页，不阻塞主流程。
"""
from __future__ import annotations

from typing import Any, Dict, List

# 低于该置信度的文字行丢弃（Vision 在清晰扫描件上通常 ≥0.8）
MIN_CONFIDENCE = 0.3
EMU_PER_PT = 12700
FONT_SIZE_FIT = 0.82   # 行框高度 → 字号的折算系数（行框含上下留白）
MIN_FONT_PT = 9.0
MAX_FONT_PT = 96.0

# Vision 识别语言优先级：中文优先，英文兜底（两语混排扫描件常见）
_RECOGNITION_LANGUAGES = ["zh-Hans", "en-US"]


def _ns_value(x: Any) -> Any:
    """pyobjc 对 Swift 属性的暴露形态随版本不同：有的直接给值，有的给
    bound selector 需调用取值。统一在此取值，两种形态都兼容。"""
    return x() if callable(x) else x


def ocr_available() -> bool:
    try:
        import Foundation  # noqa: F401
        import Vision  # noqa: F401
        return True
    except Exception:
        return False


def ocr_image_bytes(data: bytes) -> List[Dict[str, Any]]:
    """识别图片中的文字行；返回按阅读顺序的
    [{text, confidence, x, y, w, h}]，坐标为相对图片的归一化框（原点左上）。

    Vision framework 不可用或推理失败抛异常——调用方决定降级行为。"""
    import Foundation
    import Vision

    ns_data = Foundation.NSData.dataWithBytes_length_(data, len(data))
    handler = Vision.VNImageRequestHandler.alloc().initWithData_options_(ns_data, None)
    request = Vision.VNRecognizeTextRequest.alloc().init()
    request.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)
    request.setRecognitionLanguages_(_RECOGNITION_LANGUAGES)
    request.setUsesLanguageCorrection_(True)
    ok, error = handler.performRequests_error_([request], None)
    if not ok:
        raise RuntimeError(f"Vision OCR failed: {error}")

    observations = request.results() or []
    lines: List[Dict[str, Any]] = []
    for obs in observations:
        candidates = obs.topCandidates_(1)
        if not candidates:
            continue
        candidate = candidates[0]
        confidence = float(_ns_value(candidate.confidence))
        if confidence < MIN_CONFIDENCE:
            continue
        text = str(_ns_value(candidate.string)).strip()
        if not text:
            continue
        box = _ns_value(obs.boundingBox)  # 归一化，原点在左下
        origin, size = box.origin, box.size
        lines.append({
            "text": text,
            "confidence": round(confidence, 3),
            # 转成原点左上，方便直接映射到画布坐标
            "x": float(origin.x),
            "y": float(1.0 - origin.y - size.height),
            "w": float(size.width),
            "h": float(size.height),
        })
    # 阅读顺序：先上后下、同行先左后右
    lines.sort(key=lambda item: (item["y"], item["x"]))
    return lines


def overlay_textboxes(
    slide, lines: List[Dict[str, Any]], img_left: int, img_top: int, img_w: int, img_h: int
) -> int:
    """把 OCR 文字行变成可编辑原生文本框，叠在图片上层（图片背景保留）。

    坐标映射：行框（归一化）× 图片放置区（EMU）→ 文本框 bbox；
    字号按行框高度折算。返回创建的文本框数量。"""
    from pptx.util import Emu, Pt

    added = 0
    for line in lines:
        text = str(line.get("text") or "").strip()
        if not text:
            continue
        x = int(img_left + float(line.get("x", 0)) * img_w)
        y = int(img_top + float(line.get("y", 0)) * img_h)
        w = max(1, int(float(line.get("w", 0)) * img_w))
        h = max(1, int(float(line.get("h", 0)) * img_h))
        size_pt = max(MIN_FONT_PT, min(MAX_FONT_PT, (h / EMU_PER_PT) * FONT_SIZE_FIT))
        textbox = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
        frame = textbox.text_frame
        frame.word_wrap = False
        frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
        frame.text = text
        for paragraph in frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(size_pt)
        try:
            textbox.name = f"OCR 文字层 · {text[:16]}"
        except Exception:
            pass
        added += 1
    return added
