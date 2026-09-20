# -*- coding: utf-8 -*-
"""
Pruebas para el servidor HTTP local y endpoints de la API.
"""

import threading
import time
import json
import urllib.request
import pytest

from gui_server import ThreadedHTTPServer, V2GUIHandler


@pytest.fixture(scope="module")
def local_server():
    server_address = ("127.0.0.1", 5099)
    server = ThreadedHTTPServer(server_address, V2GUIHandler)
    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()
    time.sleep(0.5)
    yield "http://127.0.0.1:5099"
    server.shutdown()
    server.server_close()


def test_server_static_files(local_server):
    # 1. Test index.html
    resp = urllib.request.urlopen(f"{local_server}/")
    assert resp.status == 200
    html = resp.read().decode("utf-8")
    assert "Auditor de Excel" in html
    assert "Modo B: Rúbrica Inteligente" in html

    # 2. Test style.css
    resp_css = urllib.request.urlopen(f"{local_server}/style.css")
    assert resp_css.status == 200
    css = resp_css.read().decode("utf-8")
    assert "mesh-background" in css

    # 3. Test app.js
    resp_js = urllib.request.urlopen(f"{local_server}/app.js")
    assert resp_js.status == 200
    js = resp_js.read().decode("utf-8")
    assert "DOMContentLoaded" in js


def test_server_api_defaults(local_server):
    resp = urllib.request.urlopen(f"{local_server}/api/defaults")
    assert resp.status == 200
    data = json.loads(resp.read().decode("utf-8"))
    assert "plantilla" in data
    assert "trabajos" in data


def test_server_api_inspect_rubric(local_server):
    from core.config import DEFAULT_PLANTILLA_PATH
    url = f"{local_server}/api/inspect-rubric?plantilla={urllib.parse.quote(DEFAULT_PLANTILLA_PATH)}"
    resp = urllib.request.urlopen(url)
    assert resp.status == 200
    data = json.loads(resp.read().decode("utf-8"))
    assert data.get("status") == "ok"
    assert data.get("puntos_totales") == 100.0
    assert len(data.get("criterios")) == 8


def test_server_api_generate_template(local_server):
    resp = urllib.request.urlopen(f"{local_server}/api/generate-rubric-template")
    assert resp.status == 200
    data = json.loads(resp.read().decode("utf-8"))
    assert data.get("status") == "ok"
    assert "path" in data
