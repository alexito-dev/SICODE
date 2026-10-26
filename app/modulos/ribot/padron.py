"""Cliente del web service interno de SUNAT (IvrConsultasFacade.consultaRegistro).

Devuelve los datos de registro de un RUC (domicilio fiscal, razón social, estado...).
Primero se invoca por SOAP; si falla, se usa la página de prueba de WebLogic
(?ruc=...&operation.invoke=consultaRegistro), que muestra la misma respuesta en HTML.
"""
import html
import logging
import os
import re
import socket
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import OrderedDict
from dataclasses import dataclass
from threading import BoundedSemaphore, Lock

logger = logging.getLogger("ribot.padron")

PADRON_URL = os.environ.get("RIBOT_PADRON_URL", "http://wssunat:8080/cl-at-iaivr/ws-ivrconsultas")
PADRON_TIMEOUT = float(os.environ.get("RIBOT_PADRON_TIMEOUT", "8"))
PADRON_ENABLED = os.environ.get("RIBOT_PADRON_ENABLED", "true").strip().lower() not in ("0", "false", "no")

RUC_RE = re.compile(r"^(10|15|16|17|20)\d{9}$")

SOAP_ENVELOPE = """<?xml version="1.0" encoding="utf-8"?>
<env:Envelope xmlns:soapenc="http://schemas.xmlsoap.org/soap/encoding/"
  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
  xmlns:env="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <env:Body env:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
    <m:consultaRegistro xmlns:m="http://tempuri.org/">
      <ruc xsi:type="xsd:string">{ruc}</ruc>
    </m:consultaRegistro>
  </env:Body>
</env:Envelope>"""

FIELDS = ("razonSocial", "domicilio", "estado", "condicionDomicilio", "dependencia", "tipo", "ruc")


class PadronError(Exception):
    """Error consultando el web service (red, HTTP o respuesta sin datos)."""


class PadronUnavailableError(PadronError):
    """El servicio no responde (no se debe seguir intentando en el mismo lote)."""


@dataclass(frozen=True)
class RegistroContribuyente:
    ruc: str
    razon_social: str | None
    domicilio: str | None
    estado: str | None
    condicion_domicilio: str | None
    dependencia: str | None
    tipo: str | None


def _clean(value: str | None) -> str | None:
    if value is None:
        return None
    value = re.sub(r"\s+", " ", html.unescape(value)).strip()
    return value if value and value not in ("-", "null") else None


def parse_registro(payload: str) -> dict[str, str | None]:
    """Extrae los campos del <result> tanto de la respuesta SOAP como de la página HTML de prueba."""
    text = html.unescape(payload)
    data: dict[str, str | None] = {}
    for field in FIELDS:
        m = re.search(rf"<(?:\w+:)?{field}\b[^>]*>(.*?)</(?:\w+:)?{field}\s*>", text, re.DOTALL | re.IGNORECASE)
        data[field] = _clean(m.group(1)) if m else None
    return data


CACHE_TTL_SECONDS = 24 * 3600
CACHE_MAX_ENTRIES = 5000
COOLDOWN_SECONDS = 60
MAX_CONCURRENT_CALLS = 4


class PadronClient:
    """Cliente con memoria (24 h), pausa tras una caída y tope de consultas simultáneas.

    - Memoria: un RUC ya consultado no vuelve a pedirse al servicio.
    - Pausa: si el servicio no responde, durante COOLDOWN_SECONDS no se intenta de nuevo, así decenas de
      usuarios no esperan el tiempo de espera completo uno por uno.
    - Tope: como máximo MAX_CONCURRENT_CALLS consultas a la vez, para no saturar el servicio de SUNAT.
    """

    def __init__(
        self,
        url: str,
        timeout: float = 8.0,
        enabled: bool = True,
        cooldown: float = COOLDOWN_SECONDS,
        max_concurrent: int = MAX_CONCURRENT_CALLS,
    ):
        self.url = url.rstrip("?")
        self.timeout = timeout
        self.enabled = enabled
        self._cooldown = cooldown
        self._cache: OrderedDict[str, tuple[float, RegistroContribuyente]] = OrderedDict()
        self._lock = Lock()
        self._slots = BoundedSemaphore(max_concurrent)
        self._down_until = 0.0

    def _cached(self, ruc: str) -> RegistroContribuyente | None:
        with self._lock:
            item = self._cache.get(ruc)
            if item is None:
                return None
            if time.monotonic() - item[0] > CACHE_TTL_SECONDS:
                del self._cache[ruc]
                return None
            self._cache.move_to_end(ruc)
            return item[1]

    def _remember(self, ruc: str, registro: RegistroContribuyente) -> None:
        with self._lock:
            self._cache[ruc] = (time.monotonic(), registro)
            self._cache.move_to_end(ruc)
            while len(self._cache) > CACHE_MAX_ENTRIES:
                self._cache.popitem(last=False)

    def _raise_if_paused(self) -> None:
        remaining = self._down_until - time.monotonic()
        if remaining > 0:
            raise PadronUnavailableError(f"servicio sin respuesta; se reintentará en {int(remaining) + 1} s")

    def _http(self, request: urllib.request.Request) -> str:
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as resp:  # noqa: S310 - URL interna configurada
                charset = resp.headers.get_content_charset() or "utf-8"
                return resp.read().decode(charset, errors="replace")
        except urllib.error.HTTPError as exc:
            # Un SOAP Fault llega como HTTP 500 con cuerpo: se devuelve para que el parser decida
            body = exc.read().decode("utf-8", errors="replace")
            if body:
                return body
            raise PadronError(f"HTTP {exc.code}") from exc
        except (urllib.error.URLError, TimeoutError, ConnectionError) as exc:
            reason = getattr(exc, "reason", exc)
            raise PadronUnavailableError(f"servicio no disponible ({reason})") from exc

    def _via_soap(self, ruc: str) -> str:
        req = urllib.request.Request(  # noqa: S310 - URL interna configurada
            self.url,
            data=SOAP_ENVELOPE.format(ruc=ruc).encode("utf-8"),
            headers={"Content-Type": "text/xml; charset=utf-8", "SOAPAction": '""'},
            method="POST",
        )
        return self._http(req)

    def _via_test_page(self, ruc: str) -> str:
        query = urllib.parse.urlencode({"ruc": ruc, "operation.invoke": "consultaRegistro"})
        return self._http(urllib.request.Request(f"{self.url}?{query}"))  # noqa: S310

    def consultar(self, ruc: str) -> RegistroContribuyente:
        ruc = (ruc or "").strip()
        if not RUC_RE.match(ruc):
            raise PadronError(f"RUC inválido: '{ruc}'")
        if not self.enabled:
            raise PadronUnavailableError("consulta al padrón desactivada (RIBOT_PADRON_ENABLED=false)")
        cached = self._cached(ruc)
        if cached:
            return cached
        self._raise_if_paused()

        with self._slots:
            self._raise_if_paused()  # otro hilo pudo detectar la caída mientras esperaba su turno
            try:
                data = parse_registro(self._via_soap(ruc))
                if not data["domicilio"] and not data["razonSocial"]:
                    logger.info("SOAP sin datos para %s; probando la página de prueba de WebLogic", ruc)
                    data = parse_registro(self._via_test_page(ruc))
            except PadronUnavailableError:
                self._down_until = time.monotonic() + self._cooldown
                logger.warning("Padrón SUNAT sin respuesta; pausa de %s s", self._cooldown)
                raise
        if not data["domicilio"] and not data["razonSocial"]:
            raise PadronError("el servicio no devolvió datos para este RUC")

        registro = RegistroContribuyente(
            ruc=ruc,
            razon_social=data["razonSocial"],
            domicilio=data["domicilio"],
            estado=data["estado"],
            condicion_domicilio=data["condicionDomicilio"],
            dependencia=data["dependencia"],
            tipo=data["tipo"],
        )
        self._remember(ruc, registro)
        return registro

    def is_reachable(self) -> bool:
        """Comprobación rápida por TCP (sin consumir el servicio)."""
        if not self.enabled:
            return False
        parsed = urllib.parse.urlparse(self.url)
        port = parsed.port or (443 if parsed.scheme == "https" else 80)
        try:
            with socket.create_connection((parsed.hostname or "", port), timeout=min(self.timeout, 3)):
                return True
        except OSError:
            return False


_CONECTORES = {"de", "del", "la", "las", "los", "el", "y"}  # en español van en minúscula (salvo al inicio)
_PALABRA = re.compile(r"[^\W\d_]+")


def formato_domicilio(texto: str) -> str:
    """'JR. MIGUEL GRAU - N 345' -> 'Jr. Miguel Grau - N 345': cada palabra inicia en mayúscula y sigue en minúscula."""

    def corregir(m: re.Match) -> str:
        palabra = m.group(0).lower()
        if m.start() > 0 and palabra in _CONECTORES:
            return palabra
        return palabra[0].upper() + palabra[1:]

    return _PALABRA.sub(corregir, texto)


_client: PadronClient | None = None


def get_padron() -> PadronClient:
    """Cliente único del módulo (comparte la memoria y la pausa entre todos los usuarios)."""
    global _client
    if _client is None:
        _client = PadronClient(PADRON_URL, timeout=PADRON_TIMEOUT, enabled=PADRON_ENABLED)
    return _client
