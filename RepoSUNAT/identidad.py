"""
Reconocimiento automatico del usuario en red institucional Windows / AD.

Orden de resolucion en get_current_user(request):
  1. Cookie de sesion (identidad ya resuelta antes)
  2. NTLM Type 3 (SSO), solo si REPOSUNAT_NTLM=1 y el navegador lo envio
  3. Peticion local: usuario de Windows + Display Name de Active Directory
  4. Peticion LAN: nombre del equipo por DNS inverso

Ninguna etapa lanza excepciones: si falla, se pasa a la siguiente.
"""
import base64
import ctypes
import os
import socket
import struct
import sys
from typing import Any, Dict, Optional
from urllib.parse import unquote

COOKIE = "reposunat_user"
NTLM_ENABLED = os.environ.get("REPOSUNAT_NTLM") == "1"

# Challenge NTLM Type 2 minimo (sin validar credenciales: solo identifica)
_NTLM_TYPE2 = base64.b64encode(
    b"NTLMSSP\x00" + struct.pack("<I", 2)
    + struct.pack("<HHI", 0, 0, 48)
    + struct.pack("<I", 0x00008201)
    + b"\x01\x23\x45\x67\x89\xab\xcd\xef"
    + b"\x00" * 8
    + struct.pack("<HHI", 0, 0, 48)
).decode()


def formatear_nombre(valor: str) -> str:
    """'APELLIDOS, NOMBRES' -> 'Nombres Apellidos'."""
    valor = (valor or "").strip()
    if "," in valor:
        apellidos, nombres = valor.split(",", 1)
        return f"{nombres.strip().title()} {apellidos.strip().title()}"
    return valor


def display_name_windows() -> Optional[str]:
    """Nombre completo oficial (NameDisplay = 3) via secur32.GetUserNameExW."""
    if sys.platform != "win32":
        return None
    try:
        fn = ctypes.windll.secur32.GetUserNameExW
        size = ctypes.c_ulong(0)
        fn(3, None, ctypes.byref(size))
        if size.value:
            buf = ctypes.create_unicode_buffer(size.value)
            if fn(3, buf, ctypes.byref(size)) and buf.value.strip():
                return formatear_nombre(buf.value)
    except Exception:
        pass
    return None


def ip_local() -> str:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.settimeout(0.5)
            s.connect(("10.255.255.255", 1))
            return s.getsockname()[0]
    except Exception:
        return "127.0.0.1"


def _ntlm_type3(header: str) -> Optional[Dict[str, Any]]:
    """Extrae usuario y dominio de un mensaje NTLM Type 3."""
    try:
        if not header.upper().startswith("NTLM "):
            return None
        msg = base64.b64decode(header[5:])
        if msg[:8] != b"NTLMSSP\x00" or struct.unpack("<I", msg[8:12])[0] != 3:
            return None
        unicode_ = bool(struct.unpack("<I", msg[60:64])[0] & 1) if len(msg) >= 64 else True
        enc = "utf-16-le" if unicode_ else "latin-1"

        def campo(off):
            ln, _, pos = struct.unpack("<HHI", msg[off:off + 8])
            return msg[pos:pos + ln].decode(enc, "ignore")

        dominio, usuario = campo(28), campo(36)
        return {"login": usuario, "dominio": dominio} if usuario else None
    except Exception:
        return None


def ntlm_challenge(request) -> Optional[str]:
    """Devuelve el header WWW-Authenticate a enviar si hay que continuar el handshake, o None."""
    auth = request.headers.get("authorization", "")
    if auth.upper().startswith("NTLM "):
        try:
            tipo = struct.unpack("<I", base64.b64decode(auth[5:])[8:12])[0]
        except Exception:
            return "NTLM"
        if tipo == 1:
            return f"NTLM {_NTLM_TYPE2}"
        return None
    return "NTLM"


def _cliente_ip(request) -> str:
    return request.client.host if request.client else "127.0.0.1"


def _ips_propias() -> set:
    ips = {"127.0.0.1", "::1", "localhost", ip_local()}
    try:
        ips.update(socket.gethostbyname_ex(socket.gethostname())[2])
    except Exception:
        pass
    return ips


def _es_local(ip: str) -> bool:
    return ip in _ips_propias()


def get_current_user(request) -> Dict[str, Any]:
    """Identifica al usuario de la peticion. Siempre devuelve un dict con login y nombre."""
    ip = _cliente_ip(request)
    try:
        # 1. Cookie
        cookie = request.cookies.get(COOKIE)
        if cookie:
            login, _, nombre = unquote(cookie).partition("|")
            if login and nombre:
                return {"login": login, "nombre": nombre, "origen": "cookie", "ip": ip}

        # 2. NTLM (SSO)
        if NTLM_ENABLED:
            ntlm = _ntlm_type3(request.headers.get("authorization", ""))
            if ntlm:
                return {"login": ntlm["login"], "nombre": ntlm["login"].title(),
                        "origen": "ntlm", "ip": ip}

        # 3. Equipo local
        if _es_local(ip):
            login = os.environ.get("USERNAME") or os.environ.get("USER") or "usuario"
            return {"login": login, "nombre": display_name_windows() or login.title(),
                    "origen": "windows_local", "ip": ip}

        # 4. LAN remota
        host = ip
        try:
            host = socket.gethostbyaddr(ip)[0].split(".")[0]
        except Exception:
            pass
        nombre = host if host != ip else f"Terminal {ip.split('.')[-1]}"
        return {"login": host, "nombre": nombre, "origen": "lan_remoto", "ip": ip}
    except Exception:
        return {"login": "invitado", "nombre": "Invitado", "origen": "desconocido", "ip": ip}
