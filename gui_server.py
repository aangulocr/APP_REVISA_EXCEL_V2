# -*- coding: utf-8 -*-
# ============================================================================
# gui_server.py — Servidor Web Local para la GUI de APP_REVISA_EXCEL_V2
# ============================================================================
"""
Servidor local en Python puro (sin dependencias externas) que:
  1. Sirve la interfaz web moderna (HTML/CSS/JS) en el puerto 5000 (100% offline).
  2. Expone APIs para diálogos nativos de archivo/carpeta (tkinter en subproceso).
  3. Pre-chequea e inspecciona rúbricas antes de auditar (/api/inspect-rubric).
  4. Genera plantillas modelo de rúbricas Excel (/api/generate-rubric-template).
  5. Ejecuta la auditoría en segundo plano transmitiendo la consola vía SSE.
"""

import os
import sys
import json
import urllib.parse
import subprocess
import threading
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from socketserver import ThreadingMixIn

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from core.config import (
    BASE_DIR,
    DEFAULT_PLANTILLA_PATH,
    TRABAJOS_DIR,
    PLANTILLAS_DIR,
)
from core.rubric_parser import cargar_rubrica
from core.rubric_template import generar_plantilla_rubrica

PORT = 5000
proceso_activo = None
server_instance = None


def abrir_dialogo_archivo() -> str:
    """Abre ventana nativa de selección de archivo .xlsx mediante subproceso aislado."""
    script = (
        "import tkinter as tk; from tkinter import filedialog; "
        "root=tk.Tk(); root.withdraw(); root.attributes('-topmost', True); "
        "print(filedialog.askopenfilename(filetypes=[('Archivos de Excel', '*.xlsx')]))"
    )
    try:
        res = subprocess.run([sys.executable, "-c", script], capture_output=True, text=True, timeout=120)
        return res.stdout.strip()
    except Exception as e:
        print(f"Error abriendo diálogo de archivo: {e}")
        return ""


def abrir_dialogo_carpeta() -> str:
    """Abre ventana nativa de selección de directorio mediante subproceso aislado."""
    script = (
        "import tkinter as tk; from tkinter import filedialog; "
        "root=tk.Tk(); root.withdraw(); root.attributes('-topmost', True); "
        "print(filedialog.askdirectory())"
    )
    try:
        res = subprocess.run([sys.executable, "-c", script], capture_output=True, text=True, timeout=120)
        return res.stdout.strip()
    except Exception as e:
        print(f"Error abriendo diálogo de carpeta: {e}")
        return ""


class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    """Servidor HTTP multihilo para atender streaming SSE y peticiones API en paralelo."""
    daemon_threads = True


class V2GUIHandler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        # Desactivar logs estándar en consola para mantener terminal limpia
        pass

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)

        # Rutas de Archivos Estáticos
        if path == "/":
            self.servir_archivo(os.path.join(BASE_DIR, "gui", "index.html"), "text/html; charset=utf-8")
        elif path == "/style.css":
            self.servir_archivo(os.path.join(BASE_DIR, "gui", "style.css"), "text/css; charset=utf-8")
        elif path == "/app.js":
            self.servir_archivo(os.path.join(BASE_DIR, "gui", "app.js"), "application/javascript; charset=utf-8")

        # Endpoints API
        elif path == "/api/defaults":
            # Si no existe la plantilla por defecto pero existe la de ejemplo, usarla
            plantilla_sugerida = DEFAULT_PLANTILLA_PATH
            if not os.path.exists(plantilla_sugerida):
                ejemplo = os.path.join(PLANTILLAS_DIR, "PLANTILLA_RUBRICA_EJEMPLO.xlsx")
                if os.path.exists(ejemplo):
                    plantilla_sugerida = ejemplo

            self.enviar_json({
                "plantilla": plantilla_sugerida if os.path.exists(plantilla_sugerida) else "",
                "trabajos": TRABAJOS_DIR if os.path.exists(TRABAJOS_DIR) else ""
            })

        elif path == "/api/select-file":
            ruta = abrir_dialogo_archivo()
            self.enviar_json({"path": ruta})

        elif path == "/api/select-folder":
            ruta = abrir_dialogo_carpeta()
            self.enviar_json({"path": ruta})

        elif path == "/api/generate-rubric-template":
            try:
                ruta_gen = generar_plantilla_rubrica()
                self.enviar_json({
                    "status": "ok",
                    "path": ruta_gen,
                    "message": "Plantilla con rúbrica creada exitosamente."
                })
            except Exception as e:
                self.enviar_json({"status": "error", "message": str(e)})

        elif path == "/api/inspect-rubric":
            plantilla = query.get("plantilla", [""])[0]
            if not plantilla or not os.path.isfile(plantilla):
                self.enviar_json({"status": "error", "message": "Archivo de plantilla no encontrado."})
                return

            try:
                criterios, puntos_totales, advertencias, origen = cargar_rubrica(plantilla)
                self.enviar_json({
                    "status": "ok",
                    "origen": origen,
                    "puntos_totales": puntos_totales,
                    "advertencias": advertencias,
                    "criterios": [c.to_dict() for c in criterios]
                })
            except Exception as e:
                self.enviar_json({"status": "error", "message": str(e)})

        elif path == "/api/run-audit":
            self.ejecutar_auditoria_sse(query)

        elif path == "/api/abort":
            self.abortar_auditoria()

        elif path == "/api/shutdown":
            self.apagar_servidor()

        else:
            self.send_error(404, "Recurso no encontrado")

    def servir_archivo(self, ruta_archivo: str, content_type: str):
        try:
            with open(ruta_archivo, "rb") as f:
                contenido = f.read()
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(contenido)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(contenido)
        except Exception as e:
            self.send_error(500, f"Error interno: {e}")

    def enviar_json(self, data: dict):
        try:
            contenido = json.dumps(data, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(contenido)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(contenido)
        except Exception as e:
            self.send_error(500, f"Error al enviar JSON: {e}")

    def ejecutar_auditoria_sse(self, query):
        """Ejecuta main.py en subproceso y transmite su salida línea por línea en SSE."""
        plantilla = query.get("plantilla", [""])[0]
        trabajos = query.get("trabajos", [""])[0]
        modo = query.get("modo", ["auto"])[0]
        seccion = query.get("seccion", ["SEC_DEFAULT"])[0]
        fecha = query.get("fecha", [""])[0]

        if not plantilla or not trabajos:
            self.send_response(400)
            self.end_headers()
            self.wfile.write(b"data: " + json.dumps({"status": "error", "message": "Faltan parametros"}).encode("utf-8") + b"\n\n")
            return

        # Localizar ejecutable python del venv
        python_bin = os.path.join(BASE_DIR, ".venv", "Scripts", "python.exe")
        if not os.path.exists(python_bin):
            python_bin = os.path.join(BASE_DIR, ".venv", "bin", "python")
        if not os.path.exists(python_bin):
            python_bin = sys.executable

        cmd = [python_bin, "-u", "main.py", "--plantilla", plantilla, "--trabajos", trabajos, "--modo", modo, "--seccion", seccion]
        if fecha:
            cmd.extend(["--fecha", fecha])

        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("Connection", "keep-alive")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()

        global proceso_activo
        try:
            proceso_activo = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                cwd=BASE_DIR,
                bufsize=1
            )

            for linea in proceso_activo.stdout:
                linea_limpia = linea.rstrip()
                data_msg = json.dumps({"text": linea_limpia}, ensure_ascii=False)
                self.wfile.write(f"data: {data_msg}\n\n".encode("utf-8"))
                self.wfile.flush()

            proceso_activo.wait()
            code = proceso_activo.returncode
            data_done = json.dumps({"status": "done", "code": code}, ensure_ascii=False)
            self.wfile.write(f"data: {data_done}\n\n".encode("utf-8"))
            self.wfile.flush()

        except (ConnectionError, BrokenPipeError):
            print("Cliente web desconectado.")
            if proceso_activo:
                try:
                    proceso_activo.terminate()
                except Exception:
                    pass
        except Exception as e:
            err_str = json.dumps({"status": "error", "message": str(e)}, ensure_ascii=False)
            try:
                self.wfile.write(f"data: {err_str}\n\n".encode("utf-8"))
                self.wfile.flush()
            except Exception:
                pass

    def abortar_auditoria(self):
        global proceso_activo
        if proceso_activo and proceso_activo.poll() is None:
            try:
                proceso_activo.terminate()
                print("Auditoría cancelada a petición del usuario.")
            except Exception:
                pass
        self.enviar_json({"status": "aborted"})

    def apagar_servidor(self):
        self.enviar_json({"status": "server_stopping"})
        global server_instance
        if server_instance:
            threading.Thread(target=server_instance.shutdown).start()


def iniciar_servidor(abrir_navegador=True):
    global server_instance
    puerto = PORT
    server_address = ("0.0.0.0", puerto)

    try:
        server_instance = ThreadedHTTPServer(server_address, V2GUIHandler)
    except OSError:
        puerto = PORT + 1
        server_address = ("0.0.0.0", puerto)
        server_instance = ThreadedHTTPServer(server_address, V2GUIHandler)

    url_local = f"http://localhost:{puerto}"
    print("=" * 65)
    print("   APP_REVISA_EXCEL_V2 — SERVIDOR LOCAL DE INTERFAZ GRÁFICA")
    print("=" * 65)
    print(f"Servidor escuchando en : {url_local}")
    print("Modo de ejecución      : 100% Offline (Sin dependencias externas)")
    print("Presiona Ctrl+C o usa el botón 'Apagar' para detener el servidor.")
    print("=" * 65)

    if abrir_navegador and "--no-browser" not in sys.argv:
        threading.Timer(1.2, lambda: webbrowser.open(url_local)).start()

    try:
        server_instance.serve_forever()
    except KeyboardInterrupt:
        print("\nCerrando servidor...")
    finally:
        server_instance.server_close()
        print("Servidor detenido correctamente.")


if __name__ == "__main__":
    iniciar_servidor()
