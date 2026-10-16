"""
Resolución de rutas para modo desarrollo y modo .exe (PyInstaller onefile).
Equivalente reducido al app/core/config.py de SICODE, solo para esta app aislada.
"""
import os
import sys


def get_base_path() -> str:
    """Carpeta desde la que se leen los datos empaquetados (templates, static)."""
    if getattr(sys, "frozen", False):
        return sys._MEIPASS  # carpeta temporal donde PyInstaller descomprime el onefile
    # En desarrollo: la raíz del proyecto BOLETIN (un nivel arriba de app/)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def get_template_path(relative_path: str) -> str:
    """relative_path ej: 'app/boletin/templates'"""
    return os.path.join(get_base_path(), relative_path)


STATIC_DIR = os.path.join(get_base_path(), "app", "static")
