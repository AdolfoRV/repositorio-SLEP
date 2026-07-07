"""Migrador SLEP - lógica consolidada en un solo archivo.

Importa únicamente constants y text_utils (que se mantienen sin cambios).

Uso:
    from slep.core import procesar
    resultados = procesar(licencias_bytes, establecimientos_bytes)
"""

import io
import re
import zipfile
from datetime import datetime, date
from difflib import get_close_matches

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table

from .constants import RE_TIPO, RE_INST, RE_RESOL, RE_ESTABLECIMIENTO, AFP_MAP, MESES_MAP
from .utils import norm, norm_rut, _is_excel_error, get_by_any

# ═══════════════════════════════════════════════════════════════════════════════
# CLASIFICADORES
# ═══════════════════════════════════════════════════════════════════════════════

def _match_canon(n, patterns):
    for canon, pats in patterns.items():
        if any(re.search(p, n) for p in pats):
            return canon
    return None


def extraer_valores_unicos(hojas_hechos, *alt_names):
    vals = set()
    for _, ws, hrow, headers in hojas_hechos:
        for n in alt_names:
            if n not in headers:
                continue
            idx = headers.index(n)
            for r in range(hrow + 1, ws.max_row + 1):
                v = ws.cell(r, idx + 1).value
                if v is not None and str(v).strip():
                    vals.add(str(v).strip())
            break
    return sorted(vals)


def clasificar_generico(raw, canon_list, patterns, cutoff=0.6):
    if _is_excel_error(raw):
        return None, "Vacio (error Excel)"
    n = norm(raw)
    if not n:
        return None, "Vacio"
    canon = _match_canon(n, patterns)
    if canon:
        return canon, "OK"
    match = get_close_matches(n, [norm(x) for x in canon_list], n=1, cutoff=cutoff)
    if match:
        idx = [norm(x) for x in canon_list].index(match[0])
        return canon_list[idx], "Corregido (revisar)"
    return raw, "REVISAR: no reconocido"


def clasificar_resolucion(raw, canon_list):
    if _is_excel_error(raw):
        return None, "Vacio (error Excel)"
    n = norm(raw)
    if not n:
        return None, "Vacio"
    if re.fullmatch(r"[\d\.\-]+", n.replace(" ", "")):
        return None, "LEGACY: es un N de resolucion, no un estado"
    canon = _match_canon(n, RE_RESOL)
    if canon:
        return canon, "OK"
    match = get_close_matches(n, [norm(x) for x in canon_list], n=1, cutoff=0.6)
    if match:
        idx = [norm(x) for x in canon_list].index(match[0])
        return canon_list[idx], "Corregido (revisar)"
    return raw, "REVISAR: no reconocido"


def clasificar_afp(raw):
    if not raw or not str(raw).strip():
        return None, None, "Vacio"
    if _is_excel_error(raw):
        return None, None, "Vacio (error Excel)"
    m = re.match(r"^([^\(]+?)\s*(?:\(([\d.,]+)\))?\s*$", str(raw).strip())
    if not m:
        return raw, None, "REVISAR: no reconocido"
    nombre_n = norm(m.group(1))
    canon = AFP_MAP.get(nombre_n)
    estado = "OK"
    if canon is None:
        match = get_close_matches(nombre_n, list(AFP_MAP.keys()), n=1, cutoff=0.6)
        if not match:
            return raw, None, "REVISAR: AFP no reconocida"
        canon, estado = AFP_MAP[match[0]], "Corregido (revisar)"
    tasa = None
    if m.group(2):
        try:
            tasa = float(m.group(2).replace(",", "."))
        except ValueError:
            pass
    return canon, tasa, estado


def clasificar_establecimiento(raw, catalogo_norm, catalogo_patterns):
    if _is_excel_error(raw):
        return None, "Vacio (error Excel)"
    n = norm(raw)
    if not n:
        return None, "Vacio"
    if n in catalogo_norm:
        return catalogo_norm[n], "OK"
    canon = _match_canon(n, catalogo_patterns)
    if canon:
        return canon, "OK (regex)"
    match = get_close_matches(n, list(catalogo_norm.keys()), n=1, cutoff=0.82)
    if match:
        return catalogo_norm[match[0]], "Corregido (revisar)"
    return raw, "NUEVO: no esta en Dim_Establecimiento, se agrega y debe revisarse"


def generar_listas_canonicas(hojas_hechos, re_tipo, re_inst):
    """Genera las listas canónicas de Tipo Licencia, Institución Salud y Resolución Médica
    observadas realmente en los datos de origen (para las listas desplegables)."""

    def _gen(alt_names, patterns):
        raw = extraer_valores_unicos(hojas_hechos, *alt_names)
        clasif = {k: [] for k in patterns}
        clasif["_sin_clasificar"] = []
        for r in raw:
            n = norm(r)
            if re.fullmatch(r"[\d\.\-]+", n.replace(" ", "")):
                clasif["_sin_clasificar"].append(r)
                continue
            for canon, pats in patterns.items():
                if any(re.search(p, n) for p in pats):
                    clasif[canon].append(r)
                    break
            else:
                clasif["_sin_clasificar"].append(r)
        return sorted([k for k in clasif if k != "_sin_clasificar"])

    return (
        _gen(("Tipo Licencia",), re_tipo),
        _gen(("Institucion Salud", "Institucion Salud"), re_inst),
        _gen(("Resolucion Medica", "Resolucion Medica", "Estado"), RE_RESOL),
    )


# ═══════════════════════════════════════════════════════════════════════════════
# EXTRACTOR DE MONTOS DOBLES (Sistema / Pagado)
# ═══════════════════════════════════════════════════════════════════════════════

def extraer_montos_dobles(headers, row):
    """
    Extrae los dos bloques de montos (Sistema Dep/Netcore  vs  Pagado FONASA/ISAPRE)
    adaptándose a la estructura de cada hoja (2024, 2025 o 2026).

    Devuelve dict con:
      monto_subsidio_sistema, monto_cotizacion_previsional_sistema,
      monto_previsional_salud_sistema, total_sistema,
      monto_subsidio_pagado, monto_cotizacion_previsional_pagado,
      monto_previsional_salud_pagado, total_pagado
    """
    h_norm = [str(h).strip().lower() if h else "" for h in headers]

    def find_all(substr):
        return [i for i, h in enumerate(h_norm) if substr in h]

    def find_first(substr):
        idxs = find_all(substr)
        return idxs[0] if idxs else None

    def find_second(substr):
        idxs = find_all(substr)
        return idxs[1] if len(idxs) > 1 else None

    def val(idx):
        if idx is None or idx >= len(row):
            return None
        v = row[idx]
        if v is None:
            return None
        try:
            return float(v)
        except (ValueError, TypeError):
            return v

    # LM-2025/2026 usan 'AFP' y 'SALUD' como columnas de montos en el primer bloque,
    # mientras que LM-2024 usa 'cotizacion previsional' directamente.
    # Buscamos una columna exactamente llamada 'afp' (sin puntos) y 'salud'.
    tiene_afp_monto = any(h == "afp" for h in h_norm)
    tiene_salud_monto = any(h == "salud" for h in h_norm)

    if tiene_afp_monto and tiene_salud_monto:
        # ── Patrón 2025 / 2026 ──
        # Primer bloque (Sistema): Monto de Subsidio | AFP | AFC | SALUD | Total
        # Segundo bloque (Pagado): Monto de Subsidio | Monto cotizacion previsional | Monto Previsional Salud | Total / Total Recuperado
        idx_sub_sist = find_first("monto de subsidio")
        idx_afp      = next((i for i, h in enumerate(h_norm) if h == "afp"), None)
        idx_salud    = next((i for i, h in enumerate(h_norm) if h == "salud"), None)
        idx_total_sist = find_first("total")

        idx_sub_pag  = find_second("monto de subsidio")
        idx_cot_pag  = find_first("cotizacion previsional")
        idx_prev_pag = find_first("monto previsional salud")
        # El total del segundo bloque puede llamarse 'Total' o 'Total Recuperado'
        idx_total_pag = find_first("total recuperado") or find_second("total")

        return {
            "monto_subsidio_sistema": val(idx_sub_sist),
            "monto_cotizacion_previsional_sistema": val(idx_afp),
            "monto_previsional_salud_sistema": val(idx_salud),
            "total_sistema": val(idx_total_sist),
            "monto_subsidio_pagado": val(idx_sub_pag),
            "monto_cotizacion_previsional_pagado": val(idx_cot_pag),
            "monto_previsional_salud_pagado": val(idx_prev_pag),
            "total_pagado": val(idx_total_pag),
        }
    else:
        # ── Patrón 2024 ──
        # Primer bloque (Sistema): Monto de Subsidio | Monto cotizacion previsional | Monto Previsional Salud | Total
        # Segundo bloque (Pagado): idem, segunda ocurrencia
        idx_sub_sist = find_first("monto de subsidio")
        idx_cot_sist = find_first("cotizacion previsional")
        idx_prev_sist = find_first("monto previsional salud")
        idx_total_sist = find_first("total")

        idx_sub_pag = find_second("monto de subsidio")
        idx_cot_pag = find_second("cotizacion previsional")
        idx_prev_pag = find_second("monto previsional salud")
        idx_total_pag = find_second("total")

        return {
            "monto_subsidio_sistema": val(idx_sub_sist),
            "monto_cotizacion_previsional_sistema": val(idx_cot_sist),
            "monto_previsional_salud_sistema": val(idx_prev_sist),
            "total_sistema": val(idx_total_sist),
            "monto_subsidio_pagado": val(idx_sub_pag),
            "monto_cotizacion_previsional_pagado": val(idx_cot_pag),
            "monto_previsional_salud_pagado": val(idx_prev_pag),
            "total_pagado": val(idx_total_pag),
        }


# ═══════════════════════════════════════════════════════════════════════════════
# LECTORES
# ═══════════════════════════════════════════════════════════════════════════════

def leer_fuente(data_bytes):
    """Lee la planilla de licencias: hoja DATOS (funcionarios), LM01-2024 (unidades)
    y todas las hojas LM* (hechos de licencias médicas)."""
    wb = openpyxl.load_workbook(io.BytesIO(data_bytes), data_only=True)
    datos = wb["DATOS"]
    h = [c.value for c in datos[1]]

    def c(name):
        return h.index(name)

    funcionarios = {}
    establecimientos_raw = set()
    for r in range(2, datos.max_row + 1):
        row = [datos.cell(r, i + 1).value for i in range(len(h))]
        rut = norm_rut(row[c("RUN")])
        if not rut:
            continue
        cc = (row[c("Centro de Costo")] or "").strip() or None
        if cc:
            establecimientos_raw.add(cc)
        funcionarios[rut] = {
            "rut": rut,
            "nombre": (row[c("Nombre")] or "").strip(),
            "fecha_nacimiento": row[c("Fecha Nacimiento")],
            "sexo": row[c("Sexo")],
            "estado_civil": row[c("Estado Civil")],
            "direccion": row[c("Dirección")],
            "comuna": row[c("Comuna")],
            "telefono": row[c("Teléfono")],
            "telefono_emergencia": row[c("Teléfono Emergencia")],
            "nacionalidad": row[c("Nacionalidad")],
            "formacion_profesional": row[c("Formación Profesional")],
            "cargo": row[c("Cargo")],
            "establecimiento": cc,
        }

    lm1 = wb["LM01-2024"]
    h1 = [c.value for c in lm1[2]]
    if "Unidad" in h1:
        idxu = h1.index("Unidad")
        for r in range(3, lm1.max_row + 1):
            v = lm1.cell(r, idxu + 1).value
            if v:
                establecimientos_raw.add(str(v).strip())

    hojas_hechos = []
    for name in wb.sheetnames:
        if not name.startswith("LM") or name == "DATOS":
            continue
        ws = wb[name]
        for hr in (1, 2):
            headers = [c.value for c in ws[hr]]
            if any(v and "rut" in str(v).lower() for v in headers):
                hojas_hechos.append((name, ws, hr, headers))
                break

    return funcionarios, sorted(establecimientos_raw), hojas_hechos


def leer_dim_establecimiento(data_bytes):
    """Lee el archivo maestro Establecimientos.xlsx."""
    wb = openpyxl.load_workbook(io.BytesIO(data_bytes), data_only=True)
    ws = wb.active
    header_row = 1
    headers = []
    for r in range(1, 5):
        row_vals = [c.value for c in ws[r]]
        if any(v and "tipo" in str(v).lower() for v in row_vals):
            headers = row_vals
            header_row = r
            break

    idx_tipo = headers.index("Tipo") if "Tipo" in headers else None
    idx_nombre = headers.index("Nombre establecimiento") if "Nombre establecimiento" in headers else None
    idx_comuna = headers.index("Comuna") if "Comuna" in headers else None
    idx_direccion = headers.index("Dirección") if "Dirección" in headers else None
    idx_telefono = headers.index("Telefono") if "Telefono" in headers else None
    idx_web = headers.index("Sitio web") if "Sitio web" in headers else None

    establecimientos = []
    vistos = set()
    for r in range(header_row + 1, ws.max_row + 1):
        tipo = ws.cell(r, idx_tipo + 1).value if idx_tipo is not None else None
        nombre = ws.cell(r, idx_nombre + 1).value if idx_nombre is not None else None
        if not nombre or not str(nombre).strip():
            continue
        nombre_limpio = str(nombre).strip()
        if nombre_limpio in vistos:
            continue
        vistos.add(nombre_limpio)
        establecimientos.append({
            "tipo": tipo,
            "establecimiento": nombre_limpio,
            "comuna": ws.cell(r, idx_comuna + 1).value if idx_comuna is not None else None,
            "direccion": ws.cell(r, idx_direccion + 1).value if idx_direccion is not None else None,
            "telefono": ws.cell(r, idx_telefono + 1).value if idx_telefono is not None else None,
            "sitio_web": ws.cell(r, idx_web + 1).value if idx_web is not None else None,
        })
    return establecimientos


# ═══════════════════════════════════════════════════════════════════════════════
# TRANSFORMACIÓN
# ═══════════════════════════════════════════════════════════════════════════════

def construir_dim_afp(hojas_hechos):
    # AFPs canónicas válidas: las del mapa + No Aplica
    afps_validas = set(AFP_MAP.values()) | {"No Aplica"}

    # Paso 1: extraer tasas conocidas por AFP (de cualquier hoja)
    tasas_conocidas = {}
    for _, ws, hrow, headers in hojas_hechos:
        if "A.F.P." not in headers:
            continue
        idx = headers.index("A.F.P.")
        for r in range(hrow + 1, ws.max_row + 1):
            canon, tasa, _ = clasificar_afp(ws.cell(r, idx + 1).value)
            if canon and canon in afps_validas and "no aplica" not in canon.lower() and tasa is not None:
                tasas_conocidas[canon] = tasa

    # Paso 2: indexar solo AFPs válidas observadas
    vistos = {}
    for _, ws, hrow, headers in hojas_hechos:
        if "A.F.P." not in headers:
            continue
        idx = headers.index("A.F.P.")
        for r in range(hrow + 1, ws.max_row + 1):
            canon, tasa, _ = clasificar_afp(ws.cell(r, idx + 1).value)
            if canon and canon in afps_validas and "no aplica" not in canon.lower():
                if tasa is None:
                    tasa = tasas_conocidas.get(canon, -1)
                vistos[(canon, tasa)] = vistos.get((canon, tasa), 0) + 1

    # Asegurar registros fijos en Dim_AFP
    vistos[("No Aplica", 0)] = vistos.get(("No Aplica", 0), 0)
    vistos[("Pensionado (no aplica AFP)", 0)] = vistos.get(("Pensionado (no aplica AFP)", 0), 0)

    return [{"afp": k[0], "tasa": k[1]} for k, _ in sorted(vistos.items(), key=lambda x: (x[0][0], -x[0][1]))]

def migrar_hechos(hojas_hechos, funcionarios, catalogo_norm, catalogo_patterns,
                   TIPO_LICENCIA_CANON, INSTITUCION_SALUD_CANON, RESOLUCION_MEDICA_CANON):
    nuevos_establecimientos = {}
    salida = []
    vistos = set()  # para deduplicación
    tasas_conocidas = {} # Precalcular tasas conocidas por AFP para filas sin tasa explícita
    for _, ws2, hrow2, headers2 in hojas_hechos:
        if "A.F.P." not in headers2:
            continue
        idx2 = headers2.index("A.F.P.")
        for r2 in range(hrow2 + 1, ws2.max_row + 1):
            canon2, tasa2, _ = clasificar_afp(ws2.cell(r2, idx2 + 1).value)
            if canon2 and "no aplica" not in canon2.lower() and tasa2 is not None:
                tasas_conocidas[canon2] = tasa2

    for name, ws, hrow, headers in hojas_hechos:
        for r in range(hrow + 1, ws.max_row + 1):
            row = [ws.cell(r, i + 1).value for i in range(len(headers))]
            rut_raw = get_by_any(headers, row, "Rut")
            if rut_raw is None and get_by_any(headers, row, "Folio licencia") is None:
                continue

            rut = norm_rut(rut_raw)
            nombre_raw = get_by_any(headers, row, "Nombre")
            estado_migracion = []
            func = funcionarios.get(rut) if rut else None

            if func is None and nombre_raw:
                nombres_idx = {norm(f["nombre"]): k for k, f in funcionarios.items()}
                cand = get_close_matches(norm(nombre_raw), nombres_idx.keys(), n=1, cutoff=0.85)
                if cand:
                    rut = nombres_idx[cand[0]]
                    func = funcionarios.get(rut)

            # FIX: crear placeholder en Dim_Funcionario para evitar RUTs huérfanos
            if func is None:
                estado_migracion.append("RUT/Nombre no encontrado en Dim_Funcionario")
                if rut and rut not in funcionarios:
                    funcionarios[rut] = {
                        "rut": rut,
                        "nombre": (nombre_raw or "").strip(),
                        "fecha_nacimiento": None,
                        "sexo": None,
                        "estado_civil": None,
                        "direccion": None,
                        "comuna": None,
                        "telefono": None,
                        "telefono_emergencia": None,
                        "nacionalidad": None,
                        "formacion_profesional": None,
                        "cargo": None,
                        "establecimiento": None,
                    }
                func = funcionarios.get(rut)

            # ── Establecimiento: buscar en múltiples columnas posibles ──
            estab_raw = get_by_any(headers, row,
                "Estableciemiento", "Establecimiento", "Unidad",
                "Centro de Costo", "Lugar", "Sede", "Ubicacion")
            estab_canon = None
            if estab_raw:
                estab_canon, est_estado = clasificar_establecimiento(estab_raw, catalogo_norm, catalogo_patterns)
                if est_estado.startswith("NUEVO"):
                    nuevos_establecimientos[estab_canon] = True
                    estado_migracion.append(est_estado)
                elif est_estado != "OK" and not est_estado.startswith("OK"):
                    estado_migracion.append("Establecimiento: " + est_estado)
            else:
                # Fallback: usar el del funcionario; si es None, dejarlo como None
                estab_canon = func["establecimiento"] if func else None
                if estab_canon:
                    estab_canon2, est_estado2 = clasificar_establecimiento(estab_canon, catalogo_norm, catalogo_patterns)
                    if est_estado2.startswith("NUEVO"):
                        nuevos_establecimientos[estab_canon2] = True
                    estab_canon = estab_canon2

            tipo_canon, tipo_estado = clasificar_generico(
                get_by_any(headers, row, "Tipo Licencia"), TIPO_LICENCIA_CANON, RE_TIPO)
            if tipo_estado not in ("OK", "Vacio"):
                estado_migracion.append("Tipo Licencia: " + tipo_estado)

            inst_canon, inst_estado = clasificar_generico(
                get_by_any(headers, row, "Institución Salud", "Institucion Salud"), INSTITUCION_SALUD_CANON, RE_INST)
            if inst_estado not in ("OK", "Vacio"):
                estado_migracion.append("Institucion Salud: " + inst_estado)

            afp_canon, tasa, afp_estado = clasificar_afp(get_by_any(headers, row, "A.F.P."))
            if afp_estado not in ("OK", "Vacio"):
                estado_migracion.append("AFP: " + afp_estado)
                afp_canon = "No Aplica"
                tasa = 0
            if tasa is None and afp_canon:
                if "no aplica" in afp_canon.lower():
                    tasa = 0
                else:
                    tasa = tasas_conocidas.get(afp_canon, -1)

            resol_raw = get_by_any(headers, row, "Resolución Médica", "Resolucion Medica")
            resol_canon, resol_estado = clasificar_resolucion(resol_raw, RESOLUCION_MEDICA_CANON)
            if resol_canon is None:
                resol_canon, resol_estado = clasificar_resolucion(
                    get_by_any(headers, row, "Estado"), RESOLUCION_MEDICA_CANON)
            if resol_estado not in ("OK", "Vacio") and not str(resol_estado).startswith("LEGACY"):
                estado_migracion.append("Resolucion Medica: " + resol_estado)

            # ── Extraer ambos bloques de montos ──
            montos = extraer_montos_dobles(headers, row)

            # ── Deduplicación por llave compuesta ──
            folio = get_by_any(headers, row, "Folio licencia", "Folio Minsal")
            fecha_ini = get_by_any(headers, row, "Fecha Inicio", "Fech. Inicio")
            fecha_ter = get_by_any(headers, row, "Fecha Termino", "Fech. Termino")
            dedup_key = (str(rut or ""), str(folio or ""), str(fecha_ini or ""), str(fecha_ter or ""))
            if dedup_key in vistos:
                continue
            vistos.add(dedup_key)

            # FIX: validar coherencia de fechas
            if (fecha_ini and fecha_ter and isinstance(fecha_ini, datetime)
                    and isinstance(fecha_ter, datetime) and fecha_ter < fecha_ini):
                estado_migracion.append("Fecha Termino anterior a Fecha Inicio")

            salida.append({
                "rut": rut or (norm_rut(rut_raw) or ""),
                "nombre": func["nombre"] if func else (nombre_raw or ""),
                "fecha_nacimiento": func["fecha_nacimiento"] if func else None,
                "sexo": func["sexo"] if func else get_by_any(headers, row, "Sexo"),
                "establecimiento": estab_canon,
                "folio_licencia": folio,
                "fecha_inicio": fecha_ini,
                "fecha_termino": fecha_ter,
                "dias_lm": get_by_any(headers, row, "Días LM", "Días Lic"),
                "tipo_licencia": tipo_canon,
                "institucion_salud": inst_canon,
                "afp": afp_canon or "No Aplica",
                "tasa_afp": tasa if tasa is not None else 0,
                "resolucion_medica": resol_canon,
                **montos,
                "observaciones": " | ".join(
                    str(x) for x in [
                        get_by_any(headers, row, "Observaciones"),
                        get_by_any(headers, row, "Observaciones 2"),
                    ] if x
                ),
                "origen": name,
                "estado_migracion": "; ".join(estado_migracion) if estado_migracion else "OK",
            })

    return salida, nuevos_establecimientos


def migrar_descuentos(hojas_hechos):
    """Extrae las columnas 'MONTO DESCONTADO MES AÑO' de todas las hojas LM
    y devuelve una lista de diccionarios para Hechos_Descuentos."""
    descuentos = []
    re_desc = re.compile(r"MONTO\s+DESCONTADO\s+(\w+)\s+(\d{4})", re.IGNORECASE)

    for name, ws, hrow, headers in hojas_hechos:
        # 1) Detectar qué columnas son de descuento y a qué período corresponden
        desc_cols = []   # [(índice_en_row, periodo YYYY-MM), ...]
        for idx, h in enumerate(headers):
            if h and isinstance(h, str):
                m = re_desc.match(h.strip())
                if m:
                    mes_nombre = m.group(1).upper()
                    anio = m.group(2)
                    mes_num = MESES_MAP.get(mes_nombre)
                    if mes_num:
                        desc_cols.append((idx, f"{anio}-{mes_num}"))

        if not desc_cols:
            continue

        # 2) Recorrer filas y extraer valores no nulos / != 0
        for r in range(hrow + 1, ws.max_row + 1):
            row = [ws.cell(r, i + 1).value for i in range(len(headers))]
            rut_raw = get_by_any(headers, row, "Rut")
            folio = get_by_any(headers, row, "Folio licencia", "Folio Minsal")
            rut = norm_rut(rut_raw) if rut_raw else ""

            # Si no hay ni RUT ni Folio, saltamos la fila
            if not rut and not folio:
                continue

            for idx, periodo in desc_cols:
                valor = row[idx]
                if valor is None:
                    continue
                try:
                    monto = float(valor)
                    if monto != 0:
                        descuentos.append({
                            "folio_licencia": folio,
                            "rut": rut,
                            "periodo": periodo,
                            "monto_descuento": monto,
                            "origen": name,
                        })
                except (ValueError, TypeError):
                    pass

    return descuentos


# ═══════════════════════════════════════════════════════════════════════════════
# WRITERS
# ═══════════════════════════════════════════════════════════════════════════════

FONT = Font(name="Calibri", size=9)
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
HEADER_FILL = PatternFill("solid", start_color="1F4E78")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _estilizar_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(1, c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.freeze_panes = ws.cell(2, 1).coordinate
    ws.row_dimensions[1].height = 30


def _autoancho(ws, max_w=30):
    def _clean(s):
        if not s:
            return ""
        s = re.sub(r"[\r\n\t]+", " ", str(s))
        s = re.sub(r"\s+", " ", s)
        return s.strip()

    def _vlen(cell):
        v = cell.value
        if v is None:
            return 0
        if isinstance(v, str) and v.startswith("="):
            return 0
        f = cell.number_format or "General"
        if isinstance(v, (datetime, date)):
            if "YYYY-MM-DD" in f:
                return 10
            if "DD-MM-YYYY" in f:
                return 10
            if "DD/MM/YYYY" in f:
                return 10
            return len(v.strftime("%d-%m-%Y"))
        if isinstance(v, (int, float)):
            m = re.search(r"\.(0+)", f)
            d = len(m.group(1)) if m else 0
            p, t = "%" in f, "#,##0" in f or "#,###" in f
            n = v * 100 if p else v
            s = (f"{n:,.{d}f}" if t else f"{n:.{d}f}") + ("%" if p else "")
            if v < 0 and ";" in f:
                s = f"({s.lstrip('-')})"
            return len(s) + (1 if any(c in f for c in "$€£¥") else 0)
        return len(_clean(v))

    real_max_row = 1
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        if any(cell.value is not None for cell in row):
            real_max_row = row[0].row
    ml = {}
    for col in ws.iter_cols(min_row=1, max_row=real_max_row):
        for cell in col:
            c = cell.column
            l = _vlen(cell)
            if l > ml.get(c, 0):
                ml[c] = l
    for c, l in ml.items():
        ws.column_dimensions[get_column_letter(c)].width = min(l + 2, max_w)


def _escribir_dim(title, headers, rows, num_fmt=None):
    wb = Workbook()
    ws = wb.active
    ws.title = title
    ws.append(headers)
    for r in rows:
        ws.append(r)
    for r in range(2, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(r, c)
            cell.font = FONT
            cell.border = BORDER
            if num_fmt and c in num_fmt:
                cell.number_format = num_fmt[c]
    _estilizar_header(ws, len(headers))
    ws.row_dimensions[1].height = None
    _autoancho(ws)
    ws.add_table(Table(displayName=title.replace(" ", ""), ref=f"A1:{get_column_letter(len(headers))}{ws.max_row}"))
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def _add_ref(wb, name, headers, rows):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for r in rows:
        ws.append(r)
    ws.sheet_state = "hidden"
    return ws.max_row


def _add_dv(ws, col, formula, n_rows, title="Valor no valido", msg="Debe seleccionar un elemento de la lista."):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True, showErrorMessage=True, showInputMessage=True)
    dv.error, dv.errorTitle = msg, title
    dv.prompt, dv.promptTitle = "Seleccione un valor de la lista desplegable", "Lista de valores"
    rng = f"{get_column_letter(col)}2:{get_column_letter(col)}{1 + n_rows}"
    dv.add(rng)
    ws.add_data_validation(dv)


def escribir_dim_funcionario(funcionarios):
    rows = [
        [f["rut"], f["nombre"], f["fecha_nacimiento"], f["sexo"], f["estado_civil"],
         f["direccion"], f["comuna"], f["telefono"], f["telefono_emergencia"],
         f["nacionalidad"], f["formacion_profesional"], f["cargo"], f["establecimiento"]]
        for f in sorted(funcionarios.values(), key=lambda x: x["nombre"])
    ]
    return _escribir_dim(
        "Funcionario",
        ["RUT", "Nombre", "Fecha Nacimiento", "Sexo", "Estado Civil", "Direccion", "Comuna",
         "Telefono", "Telefono Emergencia", "Nacionalidad", "Formacion Profesional", "Cargo", "Establecimiento"],
        rows, num_fmt={3: "DD-MM-YYYY"},
    )


def escribir_dim_establecimiento(establecimientos):
    # Deduplicar por nombre
    vistos = set()
    unicos = []
    for e in sorted(establecimientos, key=lambda x: x["establecimiento"]):
        if e["establecimiento"] not in vistos:
            vistos.add(e["establecimiento"])
            unicos.append(e)
    rows = [
        [e["establecimiento"], e["tipo"], e.get("comuna"), e.get("direccion"), e.get("telefono"), e.get("sitio_web")]
        for e in unicos
    ]
    return _escribir_dim("Establecimiento", ["Establecimiento", "Tipo", "Comuna", "Direccion", "Telefono", "Sitio Web"], rows)


def escribir_dim_afp(afp_filas):
    rows = [[f["afp"], f["tasa"]] for f in afp_filas]
    return _escribir_dim("AFP", ["AFP", "Tasa"], rows, num_fmt={2: "0.00"})


def escribir_hechos_descuentos(descuentos):
    wb = Workbook()
    if descuentos:
        wb.remove(wb.active)
        ws = wb.create_sheet("Hechos_Descuentos")
        headers = ["Folio Licencia", "RUT", "Periodo", "Monto Descuento", "Fuente"]
        ws.append(headers)
        ncols = len(headers)
        C = {h: i + 1 for i, h in enumerate(headers)}

        for d in descuentos:
            ws.append([
                d["folio_licencia"],
                d["rut"],
                d["periodo"],
                d["monto_descuento"],
                d["origen"],
            ])

        _estilizar_header(ws, ncols)
        for r in range(2, ws.max_row + 1):
            for c in range(1, ncols + 1):
                cell = ws.cell(r, c)
                cell.font = FONT
                cell.border = BORDER
            ws.cell(r, C["Monto Descuento"]).number_format = "#,##0"

        _autoancho(ws)
        ws.add_table(Table(
            displayName="HechosDescuentos",
            ref=f"A1:{get_column_letter(ncols)}{ws.max_row}"
        ))
    else:
        wb.active.title = "Hechos_Descuentos"
        wb.active.append(["Folio Licencia", "RUT", "Periodo", "Monto Descuento", "Fuente"])
        _estilizar_header(wb.active, 5)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def escribir_hechos(funcionarios, afp_filas, hechos, TIPO_LICENCIA_CANON,
                    INSTITUCION_SALUD_CANON, RESOLUCION_MEDICA_CANON, descuentos=None):
    wb = Workbook()
    hoja = wb.active
    hoja.title = "Hechos_Licencias"

    # ── refs ocultas (igual que antes) ──
    func_rows = [
        [f["rut"], f["nombre"], f["fecha_nacimiento"], f["sexo"], f["establecimiento"]]
        for f in sorted(funcionarios.values(), key=lambda x: x["nombre"])
    ]
    n_func = _add_ref(wb, "ref_Funcionario", ["RUT", "Nombre", "Fecha Nacimiento", "Sexo", "Establecimiento"], func_rows)

    afp_rows = [[f["afp"], f["tasa"]] for f in afp_filas]
    n_afp = _add_ref(wb, "ref_AFP", ["AFP", "Tasa"], afp_rows)

    ref_listas = wb.create_sheet("ref_Listas")
    ref_listas.append(["Tipo Licencia", "Institucion Salud", "Resolucion Medica", "AFP"])
    for i, v in enumerate(TIPO_LICENCIA_CANON, 2):
        ref_listas.cell(i, 1, v)
    for i, v in enumerate(INSTITUCION_SALUD_CANON, 2):
        ref_listas.cell(i, 2, v)
    for i, v in enumerate(RESOLUCION_MEDICA_CANON, 2):
        ref_listas.cell(i, 3, v)
    for i, v in enumerate(["No Aplica"] + [f["afp"] for f in afp_filas], 2):
        ref_listas.cell(i, 4, v)
    ref_listas.sheet_state = "hidden"

    n_tipos = 1 + len(TIPO_LICENCIA_CANON)
    n_inst = 1 + len(INSTITUCION_SALUD_CANON)
    n_resol = 1 + len(RESOLUCION_MEDICA_CANON)
    n_afp_names = 1 + len(afp_filas) + 1

    # ── NUEVO: ref_Descuentos (hoja oculta, igual patrón que ref_Funcionario) ──
    n_desc = 0
    if descuentos:
        desc_rows = [
            [d["folio_licencia"], d["rut"], d["periodo"], d["monto_descuento"], d["origen"]]
            for d in descuentos
        ]
        n_desc = _add_ref(wb, "ref_Descuentos", ["Folio Licencia", "RUT", "Periodo", "Monto Descuento", "Fuente"], desc_rows)

    # ── Headers de Hechos_Licencias (con 8 columnas de montos: 4 Sistema + 4 Pagado) ──
    headers = [
        "Folio Licencia", "RUT", "Nombre", "Fecha Nacimiento", "Sexo", "Establecimiento",
        "Fecha Inicio", "Fecha Termino", "Dias LM", "Tipo Licencia", "Institucion Salud",
        "A.F.P.", "Tasa AFP", "Resolucion Medica",
        "Monto Subsidio Sistema", "Monto Cotizacion Previsional Sistema", "Monto Previsional Salud Sistema", "Total Sistema",
        "Monto Subsidio Pagado", "Monto Cotizacion Previsional Pagado", "Monto Previsional Salud Pagado", "Total Pagado",
        "Observaciones", "Aplica Descuentos",
        "Fuente", "Detalle inconsistencia"
    ]
    hoja.append(headers)
    ncols = len(headers)
    C = {h: i + 1 for i, h in enumerate(headers)}
    n_hist, n_rows = len(hechos), len(hechos) + 40
    FECHA_MAP = {"fecha_inicio": "Fecha Inicio", "fecha_termino": "Fecha Termino"}

    # Columnas de montos para aplicar formato numérico
    COLS_MONTO_SISTEMA = [
        C["Monto Subsidio Sistema"], C["Monto Cotizacion Previsional Sistema"],
        C["Monto Previsional Salud Sistema"], C["Total Sistema"]
    ]
    COLS_MONTO_PAGADO = [
        C["Monto Subsidio Pagado"], C["Monto Cotizacion Previsional Pagado"],
        C["Monto Previsional Salud Pagado"], C["Total Pagado"]
    ]

    for i in range(n_rows):
        r = i + 2
        h = hechos[i] if i < n_hist else None
        rut_val = h["rut"] if h else None
        hoja.cell(r, C["RUT"], rut_val)
        f_rut = f"{get_column_letter(C['RUT'])}{r}"

        # FIX: Establecimiento se escribe DIRECTAMENTE desde el hecho, no como fórmula
        # Solo Nombre, Fecha Nacimiento y Sexo usan fórmula INDEX/MATCH
        for col_name, ref_col in [("Nombre", "B"), ("Fecha Nacimiento", "C"), ("Sexo", "D")]:
            formula = (
                '=IFERROR(INDEX(ref_Funcionario!$' + ref_col + '$2:$' + ref_col + '$' + str(n_func)
                + ',MATCH(' + f_rut + ',ref_Funcionario!$A$2:$A$' + str(n_func) + ',0)),"")'
            )
            hoja.cell(r, C[col_name], formula)
        hoja.cell(r, C["Fecha Nacimiento"]).number_format = "DD-MM-YYYY"

        # Establecimiento: valor directo del hecho (puede ser None)
        if h:
            hoja.cell(r, C["Establecimiento"], h["establecimiento"])

        if h:
            for campo, header in [
                ("folio_licencia", "Folio Licencia"), ("fecha_inicio", "Fecha Inicio"),
                ("fecha_termino", "Fecha Termino"), ("dias_lm", "Dias LM"),
                ("tipo_licencia", "Tipo Licencia"), ("institucion_salud", "Institucion Salud"),
                ("afp", "A.F.P."), ("resolucion_medica", "Resolucion Medica"),
                ("monto_subsidio_sistema", "Monto Subsidio Sistema"),
                ("monto_cotizacion_previsional_sistema", "Monto Cotizacion Previsional Sistema"),
                ("monto_previsional_salud_sistema", "Monto Previsional Salud Sistema"),
                ("total_sistema", "Total Sistema"),
                ("monto_subsidio_pagado", "Monto Subsidio Pagado"),
                ("monto_cotizacion_previsional_pagado", "Monto Cotizacion Previsional Pagado"),
                ("monto_previsional_salud_pagado", "Monto Previsional Salud Pagado"),
                ("total_pagado", "Total Pagado"),
                ("observaciones", "Observaciones"), ("origen", "Fuente"),
            ]:
                hoja.cell(r, C[header], h[campo])
            hoja.cell(r, C["Detalle inconsistencia"], h["estado_migracion"] if h["estado_migracion"] != "OK" else "")
            for fkey, hkey in FECHA_MAP.items():
                if isinstance(h[fkey], datetime):
                    hoja.cell(r, C[hkey]).number_format = "DD-MM-YYYY"

        # Formato numérico para montos
        for col_monto in COLS_MONTO_SISTEMA + COLS_MONTO_PAGADO:
            hoja.cell(r, col_monto).number_format = "#,##0"

        # ── NUEVO: fórmula Aplica Descuentos con HYPERLINK externo ──
        if descuentos and n_desc > 0:
            folio_col = get_column_letter(C["Folio Licencia"])
            folio_cell = f"{folio_col}{r}"
            formula = (
                f'=IF(AND({folio_cell}<>"",COUNTIF(ref_Descuentos!$A:$A,{folio_cell})>0),'
                f'HYPERLINK("[05_Hechos_Descuentos.xlsx]Hechos_Descuentos!A"&MATCH({folio_cell},ref_Descuentos!$A:$A,0),'
                f'"Sí ("&COUNTIF(ref_Descuentos!$A:$A,{folio_cell})&" desc.)"),"No")'
            )
            hoja.cell(r, C["Aplica Descuentos"], formula)
        else:
            hoja.cell(r, C["Aplica Descuentos"], "No")

        # FIX: Tasa AFP automática. Para filas históricas escribir el valor directo
        # (dato histórico); para filas nuevas (blancas) dejar la fórmula INDEX/MATCH.
        if i < n_hist and h and h.get("tasa_afp") is not None:
            hoja.cell(r, C["Tasa AFP"], h["tasa_afp"])
        else:
            f_afp = f"{get_column_letter(C['A.F.P.'])}{r}"
            formula_afp = "=IFERROR(INDEX(ref_AFP!$B$2:$B$" + str(n_afp) + ",MATCH(" + f_afp + ",ref_AFP!$A$2:$A$" + str(n_afp) + ",0)),0)"
            hoja.cell(r, C["Tasa AFP"], formula_afp)
        hoja.cell(r, C["Tasa AFP"]).number_format = "0.00"

        for c in range(1, ncols + 1):
            cell = hoja.cell(r, c)
            cell.border = BORDER
            cell.font = FONT

    _estilizar_header(hoja, ncols)
    hoja.row_dimensions[1].height = None
    _autoancho(hoja)
    hoja.add_table(Table(displayName="HechosLicencias", ref=f"A1:{get_column_letter(ncols)}{1 + n_rows}"))

    fills = {
        "2E7D32": ["RUT", "Nombre", "Fecha Nacimiento", "Sexo", "Establecimiento"],
        "1565C0": ["Folio Licencia", "Fecha Inicio", "Fecha Termino", "Dias LM", "Tipo Licencia",
                    "Institucion Salud", "A.F.P.", "Tasa AFP", "Resolucion Medica"],
        "C62828": ["Monto Subsidio Sistema", "Monto Cotizacion Previsional Sistema",
                   "Monto Previsional Salud Sistema", "Total Sistema"],
        "AD1457": ["Monto Subsidio Pagado", "Monto Cotizacion Previsional Pagado",
                   "Monto Previsional Salud Pagado", "Total Pagado"],
        "F57C00": ["Observaciones", "Aplica Descuentos", "Fuente", "Detalle inconsistencia"],
    }
    for color, campos in fills.items():
        fill = PatternFill("solid", start_color=color, end_color=color)
        for campo in campos:
            hoja.cell(1, C[campo]).fill = fill

    # Validaciones de listas (igual que antes)
    _add_dv(hoja, C["Tipo Licencia"], "=ref_Listas!$A$2:$A$" + str(n_tipos), n_rows)
    _add_dv(hoja, C["Institucion Salud"], "=ref_Listas!$B$2:$B$" + str(n_inst), n_rows)
    _add_dv(hoja, C["Resolucion Medica"], "=ref_Listas!$C$2:$C$" + str(n_resol), n_rows)
    _add_dv(hoja, C["A.F.P."], "=ref_Listas!$D$2:$D$" + str(n_afp_names), n_rows)

    for col_fecha in [C["Fecha Inicio"], C["Fecha Termino"]]:
        dv = DataValidation(type="date", operator="between", formula1="43831", formula2="73415",
                             allow_blank=True, showErrorMessage=True, showInputMessage=True)
        dv.error, dv.errorTitle = "Debe ingresar una fecha valida.", "Fecha invalida"
        dv.prompt, dv.promptTitle = "Haga doble clic o presione Ctrl+; para insertar fecha.", "Seleccion de fecha"
        dv.add(f"{get_column_letter(col_fecha)}2:{get_column_letter(col_fecha)}{1 + n_rows}")
        hoja.add_data_validation(dv)

    dv_rut = DataValidation(type="list", formula1="=ref_Funcionario!$A$2:$A$" + str(n_func),
                             allow_blank=True, showErrorMessage=True, showInputMessage=True)
    dv_rut.error = "El RUT no existe en Dim_Funcionario. Agregue el funcionario en el archivo correspondiente."
    dv_rut.errorTitle = "Funcionario no encontrado"
    dv_rut.prompt = "Seleccione o escriba un RUT registrado en Dim_Funcionario"
    dv_rut.promptTitle = "Validacion de RUT"
    dv_rut.add(f"{get_column_letter(C['RUT'])}2:{get_column_letter(C['RUT'])}{1 + n_rows}")
    hoja.add_data_validation(dv_rut)


    # Instrucciones actualizadas
    inst = wb.create_sheet("Instrucciones", 0)
    inst.sheet_view.showGridLines = False
    inst.column_dimensions["A"].width = 110
    textos = [
        "INSTRUCCIONES - Hechos_Licencias (planilla madre nueva, esquema estrella)",
        "",
        "1) Escriba el RUT del funcionario. Nombre, Fecha de Nacimiento, Sexo se completan automaticamente desde Dim_Funcionario. Establecimiento se escribe directamente desde el hecho migrado.",
        "",
        "2) 'Tipo Licencia', 'Institucion Salud', 'A.F.P.' y 'Resolucion Medica' tienen lista desplegable. Para agregar valores nuevos, actualice la dimension correspondiente y regenere.",
        "",
        "3) Al escribir la A.F.P., la 'Tasa AFP' se completa automaticamente desde Dim_AFP. Use 'No Aplica' para pensionados.",
        "",
        "4) Las filas migradas indican su origen en 'Fuente' y posibles inconsistencias en 'Detalle inconsistencia'.",
        "",
        "5) MONTOS DOBLES: las columnas con sufijo 'Sistema' (fondo rojo) corresponden al calculo Dep/Netcore. Las columnas con sufijo 'Pagado' (fondo fucsia) corresponden a lo efectivamente pagado por FONASA/ISAPRE. En hojas 2025/2026 el primer bloque usa columnas AFP/SALUD; el mapeo se realiza automaticamente.",
        "",
        "6) La columna 'Aplica Descuentos' muestra un enlace si el folio tiene descuentos en 05_Hechos_Descuentos.xlsx. Al hacer clic se abre ese archivo y salta a la primera fila del folio. Desde alli puede filtrar la tabla por Folio para ver todos los descuentos.",
        "",
        "7) Las hojas ref_* son copias internas para las formulas. NO se editan aqui; se actualizan desde los archivos 01/02/03/05 y regenerando.",
        "",
        "8) Para Power BI: importe los 5 archivos y arme relaciones: Hechos.RUT -> Funcionario.RUT, Hechos.Establecimiento -> Establecimiento.Establecimiento, Hechos.A.F.P. -> AFP.AFP.",
    ]
    for i, t in enumerate(textos, 1):
        c = inst.cell(i, 1, t)
        c.font = Font(name="Calibri", bold=(i == 1), size=14 if i == 1 else 11)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        inst.row_dimensions[i].height = 34 if i > 1 and t else 14

    wb.active = 1
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
# PIPELINE (punto de entrada)
# ═══════════════════════════════════════════════════════════════════════════════

def procesar(src_bytes: bytes, dim_est_bytes: bytes) -> dict:
    """Ejecuta el pipeline completo y devuelve un dict {nombre_archivo: bytes},
    incluyendo un 'SLEP_files.zip' con todo junto."""

    funcionarios, _establecimientos_raw, hojas_hechos = leer_fuente(src_bytes)
    dim_establecimientos = leer_dim_establecimiento(dim_est_bytes)

    catalogo_norm = {norm(e["establecimiento"]): e["establecimiento"] for e in dim_establecimientos}
    catalogo_patterns = RE_ESTABLECIMIENTO

    tipo_canon, institucion_canon, resolucion_canon = generar_listas_canonicas(hojas_hechos, RE_TIPO, RE_INST)

    hechos, nuevos_estab = migrar_hechos(
        hojas_hechos, funcionarios, catalogo_norm, catalogo_patterns,
        tipo_canon, institucion_canon, resolucion_canon,
    )

    # Extraer descuentos de todas las hojas LM
    descuentos = migrar_descuentos(hojas_hechos)

    for ne in nuevos_estab:
        dim_establecimientos.append({
            "tipo": "Otro",
            "establecimiento": ne,
            "comuna": None,
            "direccion": None,
            "telefono": None,
            "sitio_web": None,
        })

    dim_afp = construir_dim_afp(hojas_hechos)

    out = {
        "01_Dim_Funcionario.xlsx": escribir_dim_funcionario(funcionarios),
        "02_Dim_Establecimiento.xlsx": escribir_dim_establecimiento(dim_establecimientos),
        "03_Dim_AFP.xlsx": escribir_dim_afp(dim_afp),
        "04_Hechos_Licencias.xlsx": escribir_hechos(
            funcionarios, dim_afp, hechos, tipo_canon, institucion_canon, resolucion_canon,
            descuentos=descuentos,
        ),
        "05_Hechos_Descuentos.xlsx": escribir_hechos_descuentos(descuentos),
    }

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in out.items():
            zf.writestr(name, data)
    zip_buf.seek(0)
    out["SLEP_files.zip"] = zip_buf.read()
    return out