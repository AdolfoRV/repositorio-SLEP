"""Migrador SLEP.

Uso:
    from slep import procesar
    resultados = procesar(licencias_bytes, establecimientos_bytes)
"""

from .core import procesar

__all__ = ["procesar"]
