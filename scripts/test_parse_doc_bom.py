import io
import mimetypes
import os
import uuid
from urllib import error, request
import glob


def pick_doc() -> str | None:
    docs = [p for p in glob.glob("*.doc") if not os.path.basename(p).startswith("~$")]
    if not docs:
        return None
    preferred = [p for p in docs if "евги" in os.path.basename(p).lower()]
    if preferred:
        docs = preferred
    return max(docs, key=lambda p: os.path.getsize(p))


def upload(url: str, path: str):
    boundary = "----WebKitFormBoundary" + uuid.uuid4().hex
    fn = os.path.basename(path)
    ct = mimetypes.guess_type(fn)[0] or "application/octet-stream"
    raw = open(path, "rb").read()

    b = io.BytesIO()
    w = b.write
    w(("--%s\r\n" % boundary).encode("utf-8"))
    w(('Content-Disposition: form-data; name="file"; filename="%s"\r\n' % fn).encode("utf-8"))
    w(("Content-Type: %s\r\n\r\n" % ct).encode("utf-8"))
    w(raw)
    w(("\r\n--%s--\r\n" % boundary).encode("utf-8"))
    data = b.getvalue()

    req = request.Request(url, data=data, method="POST")
    req.add_header("Content-Type", "multipart/form-data; boundary=%s" % boundary)
    req.add_header("Content-Length", str(len(data)))
    with request.urlopen(req, timeout=120) as resp:
        body = resp.read().decode("utf-8", errors="replace")
        return resp.status, body


if __name__ == "__main__":
    url = "http://127.0.0.1:5000/api/parse_bom_docx"
    p = pick_doc()
    print("picked:", p)
    if not p:
        raise SystemExit(1)
    try:
        status, body = upload(url, p)
        print("http:", status)
        print(body[:600])
    except error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace") if hasattr(e, "read") else str(e)
        print("http:", e.code)
        print(body[:600])
