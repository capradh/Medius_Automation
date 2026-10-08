import json
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import pytest

from automation.api_client import ApiClient, ApiError
from automation.config import ApiSettings, ConfigError

PAGES = {1: [{"id": 1}, {"id": 2}], 2: [{"id": 3}], 3: []}


class Handler(BaseHTTPRequestHandler):
    def do_GET(self):  # noqa: N802
        if self.path.startswith("/boom"):
            self.send_response(404); self.end_headers(); self.wfile.write(b"nope"); return
        page = int(self.path.split("page=")[1]) if "page=" in self.path else 0
        body = json.dumps({"data": PAGES.get(page, []), "auth": self.headers.get("Authorization")}).encode()
        self.send_response(200); self.send_header("Content-Type", "application/json"); self.end_headers()
        self.wfile.write(body)

    def log_message(self, *a):  # silence
        pass


@pytest.fixture
def client():
    srv = HTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    s = ApiSettings(base_url=f"http://127.0.0.1:{srv.server_port}", api_key="secret", max_retries=0, timeout=5)
    yield ApiClient(s)
    srv.shutdown()


def test_get_json_sends_auth(client):
    assert client.get_json("/x?page=1")["auth"] == "Bearer secret"


def test_pagination(client):
    assert [r["id"] for r in client.get_all_pages("/items", records_path="data")] == [1, 2, 3]


def test_http_error(client):
    with pytest.raises(ApiError, match="404"):
        client.get_json("/boom")


def test_needs_base_url():
    with pytest.raises(ConfigError):
        ApiClient(ApiSettings(base_url=None, api_key=None)).url("/x")
