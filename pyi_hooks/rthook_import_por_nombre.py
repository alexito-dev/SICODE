# Runtime hook de PyInstaller (se ejecuta ANTES que los hooks estándar).
#
# Problema: en algunas PCs el antivirus oculta los archivos recién extraídos al
# listar la carpeta _MEIxxxx mientras los escanea. El archivo existe y se puede
# abrir por su nombre, pero no aparece en os.listdir(). El importador normal de
# Python (FileFinder) se basa en ese listado, así que falla con errores como
# "ModuleNotFoundError: No module named '_socket'" o "... 'cv2'".
#
# Solución: un buscador de respaldo, al final de sys.meta_path, que busca el
# módulo directamente por su nombre de archivo (paquete, extensión .pyd o .py),
# en el mismo orden que el importador normal. Solo actúa cuando el importador
# normal no encontró el módulo. También resuelve la metadata de paquetes
# (importlib.metadata.version(...), carpetas *.dist-info) con una lista de esas
# carpetas que genera SICODE.spec (sicode_distinfo.txt).

import os
import sys
import importlib.machinery
import importlib.util

_M = importlib.machinery


class _BuscadorPorNombre:
    @classmethod
    def find_spec(cls, fullname, path=None, target=None):
        nombre = fullname.rpartition(".")[2]
        carpetas = list(path) if path is not None else list(sys.path)
        meipass = getattr(sys, "_MEIPASS", "")
        if path is None and meipass and meipass not in carpetas:
            carpetas.append(meipass)

        for carpeta in carpetas:
            if not isinstance(carpeta, str) or not carpeta or not os.path.isdir(carpeta):
                continue

            # 1) Paquete: carpeta/nombre/__init__.*
            paquete = os.path.join(carpeta, nombre)
            if os.path.isdir(paquete):
                for sufijos, loader_cls in ((_M.SOURCE_SUFFIXES, _M.SourceFileLoader),
                                            (_M.BYTECODE_SUFFIXES, _M.SourcelessFileLoader)):
                    for sufijo in sufijos:
                        init = os.path.join(paquete, "__init__" + sufijo)
                        if os.path.isfile(init):
                            return importlib.util.spec_from_file_location(
                                fullname, init, loader=loader_cls(fullname, init),
                                submodule_search_locations=[paquete])

            # 2) Módulo suelto: extensión (.pyd), fuente (.py) o bytecode (.pyc)
            for sufijos, loader_cls in ((_M.EXTENSION_SUFFIXES, _M.ExtensionFileLoader),
                                        (_M.SOURCE_SUFFIXES, _M.SourceFileLoader),
                                        (_M.BYTECODE_SUFFIXES, _M.SourcelessFileLoader)):
                for sufijo in sufijos:
                    archivo = os.path.join(carpeta, nombre + sufijo)
                    if os.path.isfile(archivo):
                        return importlib.util.spec_from_file_location(
                            fullname, archivo, loader=loader_cls(fullname, archivo))
        return None

    @classmethod
    def find_distributions(cls, context=None):
        """Metadata de paquetes (importlib.metadata) usando la lista generada en el .spec."""
        from importlib.metadata import PathDistribution
        import pathlib
        import re

        meipass = getattr(sys, "_MEIPASS", "")
        try:
            with open(os.path.join(meipass, "sicode_distinfo.txt"), encoding="utf-8") as f:
                carpetas = [linea.strip() for linea in f if linea.strip()]
        except OSError:
            return

        def normalizar(n):
            return re.sub(r"[-_.]+", "_", n).lower()

        buscado = getattr(context, "name", None)
        for carpeta in carpetas:
            nombre_dist = carpeta.rsplit(".", 1)[0].split("-", 1)[0]
            if buscado is None or normalizar(nombre_dist) == normalizar(buscado):
                ruta = os.path.join(meipass, carpeta)
                if os.path.isdir(ruta):
                    yield PathDistribution(pathlib.Path(ruta))

    @classmethod
    def invalidate_caches(cls):
        pass


if getattr(sys, "frozen", False):
    sys.meta_path.append(_BuscadorPorNombre)
