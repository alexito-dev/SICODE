from collections import deque
from datetime import datetime

# Buffer central para mensajes
log_buffer = deque(maxlen=15)

def add_log(msg):
    """Agrega un mensaje al buffer de logs"""
    now = datetime.now().strftime("%H:%M:%S")
    log_buffer.append(f"[{now}] {msg}")

def get_logs():
    """Retorna la lista de logs actuales"""
    return list(log_buffer)
