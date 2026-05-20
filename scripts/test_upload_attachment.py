import http.client
import mimetypes
import os
import uuid
from urllib.parse import urlparse


def main() -> None:
    url = os.environ.get("UPLOAD_URL", "http://127.0.0.1:5000/api/upload_attachment")
    file_path = os.environ.get("UPLOAD_FILE", "7d490b2fca4dac0075c9aeaf5cd90837.jpg")

    u = urlparse(url)
    boundary = uuid.uuid4().hex
    content_type = f"multipart/form-data; boundary={boundary}"

    filename = os.path.basename(file_path)
    with open(file_path, "rb") as f:
        content = f.read()

    mime = mimetypes.guess_type(filename)[0] or "application/octet-stream"

    body = b""
    body += f"--{boundary}\r\n".encode("utf-8")
    body += f"Content-Disposition: form-data; name=file; filename={filename}\r\n".encode("utf-8")
    body += f"Content-Type: {mime}\r\n\r\n".encode("utf-8")
    body += content + b"\r\n"
    body += f"--{boundary}--\r\n".encode("utf-8")

    conn = http.client.HTTPConnection(u.hostname, u.port, timeout=15)
    conn.request(
        "POST",
        u.path,
        body=body,
        headers={"Content-Type": content_type, "Content-Length": str(len(body))},
    )
    resp = conn.getresponse()
    raw = resp.read()
    print("status:", resp.status)
    print("content-type:", resp.getheader("Content-Type"))
    print(raw[:800].decode("utf-8", errors="replace"))


if __name__ == "__main__":
    main()
