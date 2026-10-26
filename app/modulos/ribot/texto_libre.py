"""Marcador {{TEXTO_LIBRE}}: redacción opcional que escribe el colaborador.

- Marcador solo en su línea + texto  → se reemplaza; cada salto de línea crea un párrafo nuevo
  con el mismo formato (estilo, sangría, fuente) que la línea del marcador.
- Marcador solo en su línea + vacío  → se elimina la línea completa.
- Marcador dentro de una frase       → se reemplaza en el sitio (saltos de línea como <w:br/>).
"""

import re
from copy import deepcopy

from docx.text.paragraph import Paragraph

MARKER = "{{TEXTO_LIBRE}}"
MAX_LENGTH = 5000

# Caracteres de control que no se pueden guardar en el XML de Word
_INVALID_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def normalize_lines(texto: str | None) -> list[str]:
    """Líneas del texto sin caracteres inválidos ni líneas vacías al inicio/final. [] si está vacío."""
    clean = _INVALID_XML.sub("", texto or "").replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.rstrip() for line in clean.split("\n")]
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    return lines


def _remove_paragraph(p: Paragraph) -> None:
    parent = p._p.getparent()
    # Una celda de tabla debe conservar al menos un párrafo: se vacía en lugar de borrarlo
    if parent.tag.endswith("}tc") and len(parent.findall(p._p.tag)) == 1:
        for r in list(p._p.r_lst):
            p._p.remove(r)
        return
    parent.remove(p._p)


def _marker_run_index(p: Paragraph) -> int | None:
    return next((i for i, r in enumerate(p.runs) if MARKER in (r.text or "")), None)


def _fill_standalone(p: Paragraph, lines: list[str]) -> None:
    idx = _marker_run_index(p)
    if idx is None:  # marcador repartido de forma irrecuperable: se escribe como texto plano
        for r in p.runs:
            r.text = ""
        p.runs[0].text = lines[0] if p.runs else p.add_run(lines[0]).text
        idx = 0
    original = deepcopy(p._p)
    prev = p._p
    for i, line in enumerate(lines):
        element = p._p if i == 0 else deepcopy(original)
        if i > 0:
            prev.addnext(element)
            prev = element
        para = Paragraph(element, p._parent)
        run = para.runs[idx]
        run.text = (run.text or "").replace(MARKER, line)


def apply_texto_libre(paragraphs: list[Paragraph], texto: str | None) -> bool:
    """Aplica el texto en todos los {{TEXTO_LIBRE}}. Devuelve True si la plantilla tenía el marcador."""
    lines = normalize_lines(texto)
    found = False
    for p in paragraphs:
        if MARKER not in p.text:
            continue
        found = True
        if p.text.strip() == MARKER:
            if lines:
                _fill_standalone(p, lines)
            else:
                _remove_paragraph(p)
        else:
            inline = "\n".join(lines)
            for r in p.runs:
                if r.text and MARKER in r.text:
                    r.text = r.text.replace(MARKER, inline)
            # Si se vació, evita dobles espacios en la frase
            if not lines:
                for r in p.runs:
                    if r.text and "  " in r.text:
                        r.text = re.sub(r" {2,}", " ", r.text)
    return found
