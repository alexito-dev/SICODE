"""
Boletín Oficial - SICODE (app independiente)
Levanta el servidor FastAPI en un hilo y lo muestra en una ventana nativa con pywebview.
"""
import os
import socket
import threading
import time

import uvicorn
import webview

from app.core_config import get_base_path

# Bypass compatibilidad rich/Python 3.14 (mismo fix que el launcher principal de SICODE)
try:
    import rich.cells
    rich.cells.get_character_cell_size = lambda *args, **kwargs: 1
except ImportError:
    pass

ICON_PATH = os.path.join(get_base_path(), "assets", "icon.ico")


def find_free_port() -> int:
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def run_server(port: int):
    from app.main import app
    config = uvicorn.Config(app, host="127.0.0.1", port=port, log_level="error")
    server = uvicorn.Server(config=config)
    server.install_signal_handlers = lambda: None
    server.run()


def wait_for_server(port: int, timeout: float = 20.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                return True
        except OSError:
            time.sleep(0.15)
    return False


def main():
    port = find_free_port()
    threading.Thread(target=run_server, args=(port,), daemon=True).start()

    icon_arg = ICON_PATH if os.path.exists(ICON_PATH) else None

    if not wait_for_server(port):
        webview.create_window(
            "Boletín Oficial - SICODE",
            html="<h2 style='font-family:sans-serif;padding:40px'>"
                 "No se pudo iniciar el servidor local. Cierra la aplicación e inténtalo de nuevo.</h2>",
        )
        webview.start(icon=icon_arg)
        return

    webview.create_window(
        "Boletín Oficial - SICODE",
        f"http://127.0.0.1:{port}/boletin/",
        width=1440,
        height=880,
        min_size=(1100, 650),
    )
    webview.start(icon=icon_arg)


if __name__ == "__main__":
    main()
