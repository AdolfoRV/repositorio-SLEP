"""Utilidades: manejo de errores, normalizacion de texto y RUT"""

import re
import unicodedata
from .constants import EXCEL_ERRORS

# Text utils
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


# Excel utils
def _is_excel_error(v):
    """Detecta si un valor es un error cacheado de Excel."""
    if v is None:
        return False
    return str(v).strip().lower() in EXCEL_ERRORS


def _clean_excel_error(v, default=None):
    """Reemplaza errores de Excel por un valor por defecto."""
    if _is_excel_error(v):
        return default
    return v