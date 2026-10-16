import os
import re
import logging
import asyncio
import io
import subprocess
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None
from calendar import monthrange
from datetime import datetime

logger = logging.getLogger("Boletin-Processor")

def _ps_quote(valor: str) -> str:
    """Escapa un valor para insertarlo en un literal de cadena simple de PowerShell."""
    return str(valor).replace("'", "''")

async def fetch_with_powershell_binary(url, output_path):
    """Descarga archivos binarios (PDF) saltando el Proxy de origen, usando modo fantasma asíncrono."""
    url = _ps_quote(url)
    output_path = _ps_quote(output_path)
    ps_command = (
        f"$url='{url}'; $out='{output_path}'; "
        "$wc=New-Object System.Net.WebClient; "
        "$wc.Proxy=[System.Net.WebRequest]::GetSystemWebProxy(); "
        "$wc.Proxy.Credentials=[System.Net.CredentialCache]::DefaultNetworkCredentials; "
        "$wc.DownloadFile($url, $out)"
    )
    system_root = os.environ.get('SystemRoot', 'C:\\Windows')
    ps_exe = os.path.join(system_root, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
    if not os.path.exists(ps_exe): ps_exe = "powershell"

    cmd = [ps_exe, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_command]
    
    def _run():
        try:
            subprocess.run(cmd, capture_output=True, text=True, timeout=120, creationflags=0x08000000)
            return os.path.exists(output_path)
        except Exception as e:
            logger.error(f"Error PS Binary: {e}")
            return False

    return await asyncio.to_thread(_run)

async def fetch_with_powershell_text(url, post_data=None):
    """Peticiones POST/GET de texto para buscar en el Boletín, usando modo fantasma asíncrono."""
    url = _ps_quote(url)
    if post_data:
        body = _ps_quote("&".join([f"{k}={v}" for k, v in post_data.items()]))
        ps_command = (
            f"$url='{url}'; $body='{body}'; "
            "$wc=New-Object System.Net.WebClient; "
            "$wc.Proxy=[System.Net.WebRequest]::GetSystemWebProxy(); "
            "$wc.Proxy.Credentials=[System.Net.CredentialCache]::DefaultNetworkCredentials; "
            "$wc.Headers.Add('Content-Type','application/x-www-form-urlencoded'); "
            "$wc.UploadString($url, 'POST', $body)"
        )
    else:
        ps_command = (
            f"$url='{url}'; "
            "$wc=New-Object System.Net.WebClient; "
            "$wc.Proxy=[System.Net.WebRequest]::GetSystemWebProxy(); "
            "$wc.Proxy.Credentials=[System.Net.CredentialCache]::DefaultNetworkCredentials; "
            "$wc.DownloadString($url)"
        )
    
    system_root = os.environ.get('SystemRoot', 'C:\\Windows')
    ps_exe = os.path.join(system_root, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
    if not os.path.exists(ps_exe): ps_exe = "powershell"

    cmd = [ps_exe, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_command]
    
    def _run():
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore', timeout=60, creationflags=0x08000000)
            return result.stdout
        except Exception as e:
            logger.error(f"Error PS Text: {e}")
            return None

    return await asyncio.to_thread(_run)


class BoletinProcessor:
    """Procesador para obtener y gestionar PDFs del Boletín Oficial usando PowerShell Ghost Mode.

    Desde 2026 el Boletín se publica en el portal de búsquedas de El Peruano
    (https://busquedas.elperuano.pe). El sitio anterior (diariooficial.elperuano.pe
    /BoletinOficial/Filtro) fue dado de baja y responde 404.

    El listado se obtiene del buscador por rango de fechas
    (`/?fechaIni=YYYY-MM-DD&fechaFin=YYYY-MM-DD`, paginado con `start`) y el PDF real
    se resuelve desde la ficha del cuadernillo (`/cuadernillo/BO/YYYYMMDD`).
    """

    BASE_URL = "https://busquedas.elperuano.pe"
    # Tamaño de página del buscador y tope de seguridad para no encadenar peticiones infinitas.
    PAGINA_TAM = 20
    MAX_PAGINAS = 20

    def __init__(self, cache_dir=r"D:\SICODE\APPS\BUSCADORBOLETIN\cache"):
        self.carpeta_cache = cache_dir
        self.meses_nombres = {
            "01": "01-Enero", "02": "02-Febrero", "03": "03-Marzo",
            "04": "04-Abril", "05": "05-Mayo", "06": "06-Junio",
            "07": "07-Julio", "08": "08-Agosto", "09": "09-Septiembre",
            "10": "10-Octubre", "11": "11-Noviembre", "12": "12-Diciembre"
        }
        if not os.path.exists(self.carpeta_cache):
            os.makedirs(self.carpeta_cache, exist_ok=True)
            
    def _obtener_ruta_pdf(self, fecha: str, titulo: str) -> str:
        """Calcula la ruta completa siguiendo la estructura Año/Mes/BOxxxx.pdf"""
        # fecha es YYYY-MM-DD
        anio = fecha.split('-')[0]
        mes_num = fecha.split('-')[1]
        mes_nombre = self.meses_nombres.get(mes_num, mes_num)
        
        # El título viene como "BO20250101 - 01/01/2025"
        codigo_boletin = titulo.split(' - ')[0]
        
        carpeta_anio = os.path.join(self.carpeta_cache, anio)
        carpeta_mes = os.path.join(carpeta_anio, mes_nombre)
        
        if not os.path.exists(carpeta_mes):
            os.makedirs(carpeta_mes, exist_ok=True)
            
        return os.path.join(carpeta_mes, f"{codigo_boletin}.pdf")

    def _limpiar_nombre(self, nombre: str) -> str:
        nombre = re.sub(r'[\\/*?:"<>|]', "_", nombre)
        return nombre[:200]

    def _normalizar_texto(self, texto: str) -> str:
        """Quita tildes y normaliza texto"""
        reemplazos = {
            'á': 'a', 'Á': 'a',
            'é': 'e', 'É': 'e', 
            'í': 'i', 'Í': 'i',
            'ó': 'o', 'Ó': 'o',
            'ú': 'u', 'Ú': 'u',
            'ñ': 'n', 'Ñ': 'n'
        }
        for orig, reemplazo in reemplazos.items():
            texto = texto.replace(orig, reemplazo)
        return texto.lower()

    def esta_descargado(self, fecha: str, titulo: str) -> bool:
        if not fecha or not titulo: return False
        ruta = self._obtener_ruta_pdf(fecha, titulo)
        return os.path.exists(ruta)

    def obtener_ruta_local(self, fecha: str, titulo: str) -> str | None:
        if not fecha or not titulo: return None
        ruta = self._obtener_ruta_pdf(fecha, titulo)
        return ruta if os.path.exists(ruta) else None

    async def descargar_pdf(self, url: str, fecha: str, titulo: str) -> str | None:
        if not url: return None
        ruta_completa = self._obtener_ruta_pdf(fecha, titulo)
        
        if os.path.exists(ruta_completa):
            return ruta_completa
            
        # Si la URL apunta a la ficha del cuadernillo en lugar de a un PDF directo,
        # resolvemos primero la URL real del PDF.
        if "/cuadernillo/" in url:
            logger.info(f"Resolviendo URL real del PDF desde cuadernillo: {url}")
            real_url = await self._resolver_url_pdf(url)
            if not real_url:
                logger.error(f"No se encontró el link de PDF en la ficha del cuadernillo: {url}")
                return None
            logger.info(f"URL real resuelta: {real_url}")
            url = real_url

        logger.info(f"Descargando PDF via PowerShell: {url} -> {os.path.basename(ruta_completa)}")
        success = await fetch_with_powershell_binary(url, ruta_completa)
        if success:
            return ruta_completa
        return None

    def get_file_bytes(self, path: str) -> bytes | None:
        if not path or not os.path.exists(path):
            return None
        with open(path, 'rb') as f:
            return f.read()

    def _url_cuadernillo(self, fecha_compacta: str) -> str:
        """Ficha del cuadernillo del Boletín para una fecha YYYYMMDD."""
        return f"{self.BASE_URL}/cuadernillo/BO/{fecha_compacta}"

    def _url_listado(self, fecha_ini: str, fecha_fin: str, start: int = 0) -> str:
        """Buscador por rango de fechas. `start` es el offset de paginación."""
        url = f"{self.BASE_URL}/?fechaIni={fecha_ini}&fechaFin={fecha_fin}"
        return url if start <= 0 else f"{url}&start={start}"

    @staticmethod
    def _extraer_fechas_bo(texto: str) -> set[str]:
        """Devuelve las fechas YYYYMMDD de los cuadernillos BO presentes en la respuesta."""
        return set(re.findall(r'/cuadernillo/BO/(\d{8})', texto or ""))

    @staticmethod
    def _leer_entero(texto: str, clave: str, defecto: int) -> int:
        """Lee un entero del stream SSR (`"clave",123`), tolerando las comillas escapadas."""
        match = re.search(rf'{clave}\\?"\s*,\s*(\d+)', texto or "")
        return int(match.group(1)) if match else defecto

    async def _resolver_url_pdf(self, url_cuadernillo: str) -> str | None:
        """Obtiene la URL absoluta del PDF a partir de la ficha del cuadernillo.

        Se consulta primero el endpoint de datos del router (`.data`), que devuelve el
        mismo payload sin el HTML; si falla, se cae a la página completa.
        """
        base = url_cuadernillo.split("?")[0].rstrip("/")
        for candidata in (f"{base}.data", base):
            contenido = await fetch_with_powershell_text(candidata)
            # En la página el payload viaja dentro de un string JS, con las comillas escapadas.
            match = re.search(r'urlPDF\\?"\s*,\s*\\?"([^"\\]+)', contenido or "")
            if match:
                ruta = match.group(1)
                return ruta if ruta.startswith("http") else f"{self.BASE_URL}{ruta}"
        return None

    async def obtener_pdfs_mes(self, anio: str, mes: str) -> list[dict]:
        """Lista los boletines publicados en un año/mes consultando el buscador de El Peruano."""
        try:
            anio_str = str(anio)
            mes_str = str(mes).zfill(2)
            ultimo_dia = monthrange(int(anio_str), int(mes_str))[1]
            fecha_ini = f"{anio_str}-{mes_str}-01"
            fecha_fin = f"{anio_str}-{mes_str}-{ultimo_dia:02d}"

            primera = await fetch_with_powershell_text(self._url_listado(fecha_ini, fecha_fin))
            if not primera:
                logger.error("No se pudo obtener el listado del buscador usando powershell")
                return []

            fechas = self._extraer_fechas_bo(primera)

            # El buscador devuelve todas las publicaciones del rango (NL, SE, EX, BO...),
            # así que hay que recorrer las páginas restantes para no perder boletines.
            total = self._leer_entero(primera, "totalHits", 0)
            por_pagina = self._leer_entero(primera, "paginatedBy", self.PAGINA_TAM) or self.PAGINA_TAM
            if total:
                paginas = min((total + por_pagina - 1) // por_pagina, self.MAX_PAGINAS)
            else:
                # Si el portal cambia el marcado y no podemos leer el total, barremos un
                # número fijo de páginas siempre que la primera haya traído boletines.
                paginas = self.MAX_PAGINAS if fechas else 1

            if paginas > 1:
                restantes = await asyncio.gather(*[
                    fetch_with_powershell_text(self._url_listado(fecha_ini, fecha_fin, start))
                    for start in range(por_pagina, paginas * por_pagina, por_pagina)
                ])
                for pagina in restantes:
                    fechas |= self._extraer_fechas_bo(pagina)

            prefijo = f"{anio_str}{mes_str}"
            pdf_list = []
            for fecha_compacta in sorted(f for f in fechas if f.startswith(prefijo)):
                dia_str = fecha_compacta[6:8]
                fecha_iso = f"{anio_str}-{mes_str}-{dia_str}"
                titulo = f"BO{fecha_compacta} - {dia_str}/{mes_str}/{anio_str}"
                pdf_list.append({
                    "titulo": titulo,
                    "url_online": self._url_cuadernillo(fecha_compacta),
                    "dia": int(dia_str),
                    "fecha": fecha_iso,
                    "local": self.esta_descargado(fecha_iso, titulo),
                })
            return pdf_list
        except Exception as e:
            logger.error(f"Error scraping El Peruano: {e}")
            return []

    async def buscar_palabra_en_pdfs(self, pdfs: list[dict], palabra: str):
        """Generador asíncrono para buscar en múltiples PDFs sin colgar el servidor."""
        palabra_norm = self._normalizar_texto(palabra)
        # Búsqueda de palabra exacta (con límites de palabra \b)
        patron = re.compile(
            rf"(?i)(\b{re.escape(palabra_norm)}\b|"
            rf"\b{re.escape(palabra_norm[:-1])}-\s*\n\s*{re.escape(palabra_norm[-1])}\b)"
        )
        total = len(pdfs)
        
        if not fitz:
            yield {"type": "complete", "error": "PyMuPDF no está instalado. No se puede analizar PDFs."}
            return

        for index, pdf in enumerate(pdfs):
            percent = int(((index + 1) / total) * 100)
            yield {
                "type": "progress", 
                "current": index + 1,
                "total": total,
                "percent": percent, 
                "archivo": pdf.get("titulo", "")
            }
            
            local_path = self.obtener_ruta_local(pdf["fecha"], pdf["titulo"])
            if not local_path:
                local_path = await self.descargar_pdf(pdf["url_online"], pdf["fecha"], pdf["titulo"])
                
            if local_path:
                def _busqueda():
                    found_pages = []
                    ocurrencias = 0
                    try:
                        doc = fitz.open(local_path)
                        for num_pagina, pagina in enumerate(doc):
                            texto = pagina.get_text("text") or ""
                            # Unir palabras cortadas por guion
                            texto_limpio = re.sub(r'(\w)-\s*\n\s*(\w)', r'\1\2', texto)
                            texto_norm = self._normalizar_texto(texto_limpio)
                            
                            busqueda = list(re.finditer(patron, texto_norm))
                            if busqueda:
                                occurrences = len(busqueda)
                                ocurrencias += occurrences
                                found_pages.append(num_pagina + 1)

                        doc.close()
                    except Exception as e:
                        logger.error(f"Error analizando {local_path}: {e}")
                    return found_pages, ocurrencias

                # Ejecutar el parseo PDF en un thread aparte 
                found_pages, ocurrencias = await asyncio.to_thread(_busqueda)
                
                if found_pages:
                    yield {
                        "type": "found",
                        "pdf": pdf,
                        "paginas": found_pages,
                        "veces": ocurrencias
                    }
                    
        yield {"type": "progress", "percent": 100, "archivo": "Búsqueda finalizada"}
        yield {"type": "complete", "message": "Búsqueda terminada", "total": total}
