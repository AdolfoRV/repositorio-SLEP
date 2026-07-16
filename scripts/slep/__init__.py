"""Migrador SLEP Los Libertadores.

Paquete de procesamiento que transforma la planilla madre histórica de
licencias médicas (Excel, desde 2024) en un modelo estrella normalizado de
cinco archivos, listo para imputación y para Power BI.

La API pública es una única función:

Uso:
    from slep import procesar
    resultados = procesar(licencias_bytes, establecimientos_bytes)

``resultados`` es un ``dict`` ``{nombre_archivo: bytes}`` con los cinco
``.xlsx`` generados y un ``SLEP_files.zip`` que los agrupa.

Estructura del paquete:
    * :mod:`slep.constants`: reglas de negocio (regex y mapas canónicos).
    * :mod:`slep.utils`: normalización de texto, RUT y errores de Excel.
    * :mod:`slep.core`: pipeline completo (lectura -> clasificación ->
      migración -> escritura de salidas).
"""

from .core import procesar

__all__ = ["procesar"]
