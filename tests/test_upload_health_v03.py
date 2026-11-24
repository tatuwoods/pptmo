"""v0.3 E 组：上传校验增强（item 22）+ 健康检查版本（0.3.0）。"""
from __future__ import annotations


def _upload(client, name, payload, content_type="application/octet-stream"):
    return client.post(
        "/api/docs/upload",
        files={"file": (name, payload, content_type)},
    )


def test_upload_ppt_rejected_with_conversion_hint(client):
    resp = _upload(client, "legacy.ppt", b"x" * 5000, "application/vnd.ms-powerpoint")
    assert resp.status_code == 400
    detail = resp.json()["detail"]
    assert "转存" in detail and ".pptx" in detail


def test_upload_pptm_rejected_with_conversion_hint(client):
    resp = _upload(client, "macro.pptm", b"x" * 5000)
    assert resp.status_code == 400
    assert "转存" in resp.json()["detail"]


def test_upload_pps_rejected_with_conversion_hint(client):
    resp = _upload(client, "show.pps", b"x" * 5000)
    assert resp.status_code == 400
    assert "转存" in resp.json()["detail"]


def test_upload_empty_file_friendly(client):
    resp = _upload(client, "empty.pptx", b"")
    assert resp.status_code == 400
    assert "空" in resp.json()["detail"]


def test_upload_tiny_corrupt_friendly(client):
    resp = _upload(client, "tiny.pptx", b"PK")  # <100 字节且非有效 zip
    assert resp.status_code == 400
    detail = resp.json()["detail"]
    assert detail != "只支持 .pptx 文件"  # 不是扩展名问题，而是内容问题


def test_upload_oversize_friendly(client, monkeypatch):
    import main as main_module

    monkeypatch.setattr(main_module, "MAX_UPLOAD_BYTES", 1000)
    resp = _upload(client, "big.pptx", b"x" * 5000)
    assert resp.status_code == 413
    assert "MB" in resp.json()["detail"]


def test_health_reports_v040(client):
    resp = client.get("/api/health")
    assert resp.status_code == 200
    body = resp.json()
    assert body["version"] == "0.4.0"
    assert body["ok"] is True
