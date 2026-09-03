"""
Dev-only CORS proxy for testing the GitHub Pages build locally.

Forwards GET requests to the Nium Reference Data API and adds the
Access-Control-Allow-Origin header the real service does not send yet.
Never deploy this — it exists only so the in-browser build can be verified
end-to-end before the refdata team allows the Pages origin.

    python3 dev/cors_proxy.py            # listens on http://localhost:8787
"""
import gzip
import sys
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

UPSTREAM = "https://refdata.prod.nium.com/ref-data-service"
PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8787


class Handler(BaseHTTPRequestHandler):
    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "x-client-id, accept, content-type")

    def do_OPTIONS(self):
        self.send_response(204); self._cors(); self.end_headers()

    def do_GET(self):
        req = urllib.request.Request(UPSTREAM + self.path, headers={
            "X-CLIENT-ID": "cors-dev-proxy", "Accept-Encoding": "gzip"})
        try:
            with urllib.request.urlopen(req, timeout=300) as r:
                body = r.read()
                if (r.headers.get("Content-Encoding") or "").lower() == "gzip":
                    body = gzip.decompress(body)
                status, ctype = r.status, r.headers.get("Content-Type", "application/json")
        except urllib.error.HTTPError as e:
            body, status, ctype = e.read(), e.code, "application/json"
        self.send_response(status); self._cors()
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers(); self.wfile.write(body)

    def log_message(self, fmt, *args):
        sys.stderr.write("proxy %s %s\n" % (self.command, self.path[:80]))


if __name__ == "__main__":
    print("dev CORS proxy → %s on http://localhost:%d" % (UPSTREAM, PORT))
    ThreadingHTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
