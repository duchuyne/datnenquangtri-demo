from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse
import json

ROOT = Path(__file__).resolve().parent
SITE = ROOT / "metrocity-hoalac.com"
ZONES = ROOT / "images" / "images cac khu dat"
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".svg", ".avif"}

class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path.endswith("/api-list-images.php"):
            query = parse_qs(parsed.query)
            if query.get("source", [""])[0] == "zones":
                items = []
                for path in ZONES.rglob("*"):
                    if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS:
                        relative = path.relative_to(ZONES).as_posix()
                        items.append("../images/images cac khu dat/" + relative.replace(" ", "%20"))
                items.sort(key=str.casefold)
                payload = json.dumps({"ok": True, "count": len(items), "items": items}, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
        super().do_GET()

if __name__ == "__main__":
    server = ThreadingHTTPServer(("127.0.0.1", 8000), Handler)
    print("Serving http://127.0.0.1:8000/metrocity-hoalac.com/index.html")
    server.serve_forever()
