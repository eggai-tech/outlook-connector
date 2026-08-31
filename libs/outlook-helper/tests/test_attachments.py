import base64

from outlook_helper.attachments import (
    LARGE_ATTACHMENT_THRESHOLD,
    Attachment,
    is_large,
    iter_upload_chunks,
    load_attachment,
    to_inline_payload,
)


def test_load_attachment_from_path(tmp_path):
    p = tmp_path / "report.pdf"
    p.write_bytes(b"%PDF-1.4 data")
    att = load_attachment(str(p))
    assert att.name == "report.pdf"
    assert att.content == b"%PDF-1.4 data"
    assert att.content_type == "application/pdf"


def test_load_attachment_passthrough():
    att = Attachment(name="x.txt", content=b"hi", content_type="text/plain")
    assert load_attachment(att) is att


def test_load_attachment_unknown_type_defaults_octet_stream(tmp_path):
    p = tmp_path / "blob.weirdext"
    p.write_bytes(b"\x00\x01")
    att = load_attachment(str(p))
    assert att.content_type == "application/octet-stream"


def test_is_large_threshold():
    small = Attachment("s", b"x" * 10, "text/plain")
    big = Attachment("b", b"x" * (LARGE_ATTACHMENT_THRESHOLD + 1), "text/plain")
    assert is_large(small) is False
    assert is_large(big) is True


def test_to_inline_payload_base64_encodes():
    att = Attachment("note.txt", b"hello", "text/plain")
    payload = to_inline_payload(att)
    assert payload["@odata.type"] == "#microsoft.graph.fileAttachment"
    assert payload["name"] == "note.txt"
    assert payload["contentType"] == "text/plain"
    assert base64.b64decode(payload["contentBytes"]) == b"hello"


def test_iter_upload_chunks_splits_and_sets_content_range():
    data = b"0123456789"  # 10 bytes
    chunks = list(iter_upload_chunks(data, chunk_size=4))
    assert [c[0] for c in chunks] == [b"0123", b"4567", b"89"]
    # (data, content_range) with total size 10
    assert chunks[0][1] == "bytes 0-3/10"
    assert chunks[1][1] == "bytes 4-7/10"
    assert chunks[2][1] == "bytes 8-9/10"
