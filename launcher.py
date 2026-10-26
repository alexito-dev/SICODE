import os
import sys
import socket
import threading
import traceback
import webbrowser
import uvicorn
import tkinter as tk
from tkinter import messagebox
import contextlib


def _global_exception_handler(exc_type, exc_value, exc_traceback):
    err_text = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    print("\n" + "=" * 60)
    print("SE HA PRODUCIDO UN ERROR EN SICODE:")
    print(err_text)
    print("=" * 60)

    # Se guarda en D:\SICODE\LOGS (o C:\SICODE\LOGS); si falla, junto al .exe
    log_path = "ERROR_SICODE.txt"
    exe_dir = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
    for base in (r"D:\SICODE", r"C:\SICODE", exe_dir):
        try:
            if base.startswith("D:") and not os.path.exists("D:\\"):
                continue
            logs_dir = os.path.join(base, "LOGS") if base != exe_dir else base
            os.makedirs(logs_dir, exist_ok=True)
            log_path = os.path.join(logs_dir, "ERROR_SICODE.txt")
            with open(log_path, "w", encoding="utf-8") as lf:
                lf.write(err_text)
            break
        except Exception:
            continue

    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Error de Ejecución en SICODE", f"Se produjo un error:\n\n{err_text[:600]}\n\n(Revisa {log_path})"
        )
    except Exception:
        pass

    with contextlib.suppress(Exception):
        input("\nPresione ENTER para salir...")


sys.excepthook = _global_exception_handler

# --- CONFIGURACIÓN DE COLORES PORTAL ---
COLOR_PRIMARY = "#0163ac"  # Azul Principal
COLOR_PRIMARY_DARK = "#014f88"  # Azul Oscuro
COLOR_BG_LIGHT = "#f3f4f6"  # Fondo Gris Claro
COLOR_WHITE = "#ffffff"
COLOR_TEXT_DARK = "#1e2939"
COLOR_ACCENT_RED = "#e7000b"
COLOR_CELESTE = "#e4f3ff"  # Celeste Fondo

# Importar el Monitor

# Bypass para Python 3.14
import rich.cells

rich.cells.get_character_cell_size = lambda *args, **kwargs: 1


def is_port_in_use(port: int = 8000) -> bool:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.5)
            return s.connect_ex(("127.0.0.1", port)) == 0
    except Exception:
        return False


# --- CONFIGURACIÓN DE RUTAS ---
# Asegurar que el directorio raíz está en el path para importar 'app'
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# --- PRE-CHECK: BASE DE DATOS ---
# Asegurar que la carpeta de la BD existe antes de importar la app (que intenta conectar)
try:
    db_dir = r"D:\SICODE\DB"
    if not os.path.exists(db_dir):
        os.makedirs(db_dir, exist_ok=True)
except Exception as e:
    # Si falla (ej: permisos), lo mostramos pero intentamos seguir
    messagebox.showwarning("Advertencia", f"No se pudo crear el directorio de BD: {e}")

from app.main import app
import contextlib

# --- VARIABLE GLOBAL DEL SERVIDOR ---
server_instance = None
running_thread = None


class ServerThread(threading.Thread):
    def __init__(self, app, host="0.0.0.0", port=8000):
        threading.Thread.__init__(self)
        self.server = uvicorn.Server(
            config=uvicorn.Config(app, host=host, port=port, log_level="warning", loop="asyncio")
        )

    def run(self):
        self.server.run()

    def stop(self):
        self.server.should_exit = True


def buscar_icono():
    """Busca el icono .ico si existe, para la ventana"""
    # Si tienes un icono, pon su nombre aquí. Si no, usa el de defecto
    possible_paths = ["app.ico", "static/favicon.ico", "app/static/favicon.ico"]
    for p in possible_paths:
        if os.path.exists(p):
            return p
    return None


# --- LÓGICA DE LA INTERFAZ ---


def stop_server():
    global server_instance
    if server_instance:
        lbl_status.config(text="DETENIENDO...", fg="#FFA500")
        btn_stop.config(state="disabled")
        btn_open.config(state="disabled")
        root.update()

        server_instance.stop()

        # Esperar a que se detenga realmente (join bloquea, mejor after)
        root.after(1000, confirm_stop)


def confirm_stop():
    lbl_status.config(text="● DETENIDO", fg="#e74c3c")  # Rojo
    btn_start.config(state="normal")


def open_browser():
    webbrowser.open("http://localhost:8000")


def on_closing():
    if running_thread and running_thread.is_alive():
        if messagebox.askokcancel("Salir", "¿El servidor está corriendo. Quieres detenerlo y salir?"):
            stop_server()
            # Dar tiempo al shutdown
            root.after(1000, root.destroy)
    else:
        root.destroy()


# --- GUI ---
root = tk.Tk()
root.title("Launcher SICODE")
root.geometry("350x250")
root.resizable(False, False)

# Intentar poner icono
icon_path = buscar_icono()
if icon_path:
    with contextlib.suppress(Exception):
        root.iconbitmap(icon_path)

# Estilos simples
font_title = ("Segoe UI", 16, "bold")
font_status = ("Segoe UI", 12, "bold")
font_btn = ("Segoe UI", 10)

# Frame Principal
main_frame = tk.Frame(root, padx=20, pady=20)
main_frame.pack(expand=True, fill="both")

# Título
tk.Label(main_frame, text="SERVIDOR SICODE", font=font_title).pack(pady=(0, 20))

# Estado
lbl_status = tk.Label(main_frame, text="● DETENIDO", font=font_status, fg="#e74c3c")
lbl_status.pack(pady=(0, 10))

# Info de conexión
lbl_ip = tk.Label(main_frame, text="", font=("Segoe UI", 9), fg="#7f8c8d")
lbl_ip.pack(pady=(0, 20))


def get_local_ip():
    try:
        # Conectar a un DNS público para determinar la IP de salida (no envía datos reales)
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


# --- LÓGICA DE LA INTERFAZ ---
def start_server():
    global server_instance, running_thread

    if running_thread and running_thread.is_alive():
        return

    btn_start.config(state="disabled")
    lbl_status.config(text="INICIANDO...", fg="#FFA500")  # Naranja
    root.update()

    # Crear e iniciar el hilo del servidor en 0.0.0.0 para acceso LAN
    server_instance = ServerThread(app, host="0.0.0.0", port=8000)
    server_instance.start()
    running_thread = server_instance

    # Esperar un momento visualmente y confirmar
    root.after(1500, confirm_start)


def confirm_start():
    local_ip = get_local_ip()
    url = f"http://{local_ip}:8000"

    lbl_status.config(text="● EJECUTANDO", fg="#2ecc71")  # Verde
    lbl_ip.config(text=f"Disponible en:\n{url}")

    btn_stop.config(state="normal")
    btn_open.config(state="normal")


# Botones
btn_frame = tk.Frame(main_frame)
btn_frame.pack(pady=10)

btn_start = tk.Button(
    btn_frame,
    text="INICIAR SISTEMA",
    command=start_server,
    bg="#2ecc71",
    fg="white",
    font=font_btn,
    width=20,
    height=2,
    cursor="hand2",
)
btn_start.pack(pady=5)

btn_stop = tk.Button(
    btn_frame,
    text="DETENER",
    command=stop_server,
    bg="#e74c3c",
    fg="white",
    font=font_btn,
    width=20,
    state="disabled",
    cursor="hand2",
)
btn_stop.pack(pady=5)

btn_open = tk.Button(
    main_frame,
    text="Abrir en Navegador ↗",
    command=open_browser,
    font=("Segoe UI", 9, "underline"),
    fg="#3498db",
    relief="flat",
    cursor="hand2",
    state="disabled",
)
btn_open.pack(pady=(10, 0))

# Copyright
tk.Label(main_frame, text="v1.0.0 - CDAPP Corporate", font=("Arial", 8), fg="#7f8c8d").pack(side="bottom")

root.protocol("WM_DELETE_WINDOW", on_closing)

if __name__ == "__main__":
    root.mainloop()
