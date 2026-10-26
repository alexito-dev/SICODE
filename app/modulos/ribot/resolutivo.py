"""Ajusta, sin etiquetas, la frase "(corresponde|procede) modificar|revocar la orden de pago|resolución de multa".

- Verbo: FUNDADO -> "revocar"; FUNDADO EN PARTE -> "modificar".
- Tipo: el que dice la línea "Valor" de la Ficha (Orden de Pago, Resolución de Multa...).

Se trabaja sobre el texto completo del párrafo y se conserva el formato de cada run (negrita, etc.).
"""

import re

from docx.text.paragraph import Paragraph

_FRASE = re.compile(
    r"\b(?P<verbo>modificar|revocar)(?P<puente>\s+la\s+)(?P<tipo>orden\s+de\s+pago|resoluci[oó]n\s+de\s+(?:multa|determinaci[oó]n))",
    re.IGNORECASE,
)

VERBO_FUNDADO = "revocar"
VERBO_EN_PARTE = "modificar"


def verbo_para(fundado: str | None) -> str | None:
    """'revocar' si es FUNDADO, 'modificar' si es FUNDADO EN PARTE; None si no se sabe."""
    if fundado == "FUNDADO":
        return VERBO_FUNDADO
    if fundado == "FUNDADO EN PARTE":
        return VERBO_EN_PARTE
    return None


def _same_case(original: str, new: str) -> str:
    return new[:1].upper() + new[1:] if original[:1].isupper() else new


def _replace_range(p: Paragraph, start: int, end: int, new: str) -> None:
    """Reemplaza los caracteres [start, end) del texto del párrafo sin perder el formato de los runs."""
    runs = p.runs
    pos, placed = 0, False
    for run in runs:
        text = run.text or ""
        lo, hi = max(pos, start), min(pos + len(text), end)
        if lo < hi:
            insert = "" if placed else new
            run.text = text[: lo - pos] + insert + text[hi - pos :]
            placed = True
        pos += len(text)


def adapt_paragraph(p: Paragraph, verbo: str | None, tipo: str | None) -> bool:
    """Aplica verbo y tipo en la frase del párrafo. Devuelve True si la frase estaba."""
    match = _FRASE.search(p.text)
    if not match:
        return False
    # De atrás hacia adelante, para que las posiciones anteriores no se muevan
    if tipo:
        _replace_range(p, match.start("tipo"), match.end("tipo"), tipo.lower())
    if verbo:
        _replace_range(p, match.start("verbo"), match.end("verbo"), _same_case(match.group("verbo"), verbo))
    return True


def adapt_resolutivo(paragraphs: list[Paragraph], fundado: str | None, tipo: str | None) -> tuple[bool, bool]:
    """Devuelve (había_frase, se_aplicó_verbo)."""
    verbo = verbo_para(fundado)
    found = False
    for p in paragraphs:
        found = adapt_paragraph(p, verbo, tipo) or found
    return found, verbo is not None
