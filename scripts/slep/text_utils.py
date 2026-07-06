"""Utilidades de normalizacion de texto y RUT, sin dependencias del resto del paquete."""

import re
import unicodedata

def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def norm(s) -> str:
    if s is None:
        return ""
    txt = strip_accents(str(s)).lower().strip()
    txt = re.sub(r"\s+", " ", txt)
    return txt.replace("n°", "n ").replace("n °", "n ").strip()

def norm_rut(rut):
    if rut is None:
        return None
    s = re.sub(r"[^0-9kK]", "", str(rut)).upper()
    return f"{s[:-1]}-{s[-1]}" if len(s) >= 2 else None

def get_by_any(headers, row, *names):
    """Devuelve el primer valor de `row` cuyo header coincide con alguno de `names`."""
    for n in names:
        if n in headers:
            return row[headers.index(n)]
    return None