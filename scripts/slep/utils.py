"""Utilidades del Migrador SLEP: normalización de texto, RUT y errores de Excel.

Contiene las funciones de más bajo nivel del paquete. Son las que garantizan
que las reglas de negocio de ``constants.py`` operen sobre texto comparable
(sin tildes, sin mayúsculas, sin espaciado irregular) y que los errores
cacheados de Excel no contaminen la clasificación.

Reglas de negocio implementadas aquí:
    * RB-01: normalización canónica de texto (:func:`norm`).
    * RB-02: normalización de RUT chileno (:func:`norm_rut`).
    * RB-03: errores de Excel tratados como vacío (:func:`_is_excel_error`).
"""

import re
import unicodedata
from .constants import EXCEL_ERRORS


# TEXT UTILS

def strip_accents(s: str) -> str:
    """Elimina los signos diacríticos (tildes, diéresis, etc.) de un texto.

    Descompone cada carácter en su forma NFD (carácter base + marcas
    combinantes) y descarta las marcas (categoría Unicode ``Mn``).

    Args:
        s: Texto de entrada.

    Returns:
        Texto equivalente sin acentos. ``"El Mañío"`` -> ``"El Manio"``.
    """
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def norm(s) -> str:
    """Normaliza un valor cualquiera a su forma canónica comparable (RB-01).

    Es la puerta de entrada de todo texto libre del pipeline: las regex de
    ``constants.py`` asumen este formato.

    Transformaciones aplicadas, en orden:
        1. ``None`` -> cadena vacía.
        2. Conversión a ``str``, eliminación de tildes (:func:`strip_accents`),
           minúsculas y recorte de extremos.
        3. Colapso de espacios múltiples/saltos de línea a un solo espacio.
        4. ``"n°"`` / ``"n °"`` -> ``"n "`` (uniforma la numeración de
           establecimientos: "Escuela N°334" -> "escuela n 334").

    Args:
        s: Valor crudo proveniente de una celda de Excel (puede ser ``None``,
           número, fecha, etc.; se convierte a ``str``).

    Returns:
        Texto normalizado en minúsculas, sin tildes y con espaciado simple.
        Cadena vacía si la entrada era ``None``.
    """
    if s is None:
        return ""
    txt = strip_accents(str(s)).lower().strip()
    txt = re.sub(r"\s+", " ", txt)
    return txt.replace("n°", "n ").replace("n °", "n ").strip()


def norm_rut(rut):
    """Normaliza un RUT chileno al formato canónico ``NNNNNNNN-DV`` (RB-02).

    Elimina todo lo que no sea dígito o la letra K (puntos, guiones,
    espacios) y reconstruye el RUT como ``cuerpo-dígito_verificador`` en
    mayúsculas. No valida el dígito verificador (módulo 11): la planilla
    madre contiene RUT históricos con DV incorrecto y la regla de negocio
    es conservarlos tal cual para no perder trazabilidad.

    Args:
        rut: Valor crudo del RUT (acepta formatos como ``"12.345.678-k"``,
            ``12345678k`` o un número de celda).

    Returns:
        RUT normalizado (ej. ``"12345678-K"``), o ``None`` si la entrada es
        ``None`` o queda con menos de 2 caracteres útiles.
    """
    if rut is None:
        return None
    s = re.sub(r"[^0-9kK]", "", str(rut)).upper()
    return f"{s[:-1]}-{s[-1]}" if len(s) >= 2 else None


# EXCEL UTILS

def _is_excel_error(v) -> bool:
    """Indica si un valor es un error cacheado de Excel (RB-03).

    Al leer las planillas con ``data_only=True`` se obtiene el último valor
    calculado de cada fórmula; las fórmulas rotas aparecen como cadenas
    ``"#N/A"``, ``"#REF!"``, etc. (ver ``constants.EXCEL_ERRORS``).

    Args:
        v: Valor de la celda.

    Returns:
        ``True`` si el valor (insensible a mayúsculas y espacios) está en el
        conjunto de errores conocidos; ``False`` en caso contrario, incluido
        ``None``.
    """
    if v is None:
        return False
    return str(v).strip().lower() in EXCEL_ERRORS


def _clean_excel_error(v, default=None):
    """Reemplaza un error de Excel por un valor por defecto (RB-03).

    Args:
        v: Valor de la celda.
        default: Valor a devolver cuando ``v`` es un error de Excel
            (``None`` por defecto).

    Returns:
        ``default`` si ``v`` es un error cacheado de Excel; el propio ``v``
        en caso contrario.
    """
    if _is_excel_error(v):
        return default
    return v
