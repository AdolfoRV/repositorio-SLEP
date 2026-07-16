"""Constantes de clasificación y canonización del Migrador SLEP.

Este módulo es la **fuente única de verdad** de las reglas de negocio
basadas en texto libre. Todas las expresiones regulares y mapas de
canonización del pipeline viven aquí, separados de la lógica de
procesamiento (``core.py``), para poder mantenerlos sin tocar código.

Convenciones de escritura de las expresiones regulares:

* Se escriben **sin tildes y en minúsculas**, porque todo texto crudo pasa
  antes por :func:`slep.utils.norm`, que elimina acentos, deja en
  minúsculas y colapsa espacios (regla de negocio RB-01).
* Se usa ``\\s*`` / ``\\s+`` entre palabras para tolerar espaciado irregular
  y ``.*`` cuando pueden existir palabras intermedias.
* **El orden de las claves importa**: el clasificador
  (:func:`slep.core._match_canon`) recorre el diccionario en orden de
  inserción y gana la primera coincidencia (RB-04).

Los identificadores RB-* referencian el catálogo de reglas de negocio del
documento técnico (``docs/Documento_Tecnico_Migrador_SLEP.md``).
"""

# TIPO DE LICENCIA MÉDICA (RB-04)
# Mapea el valor canónico (categoría oficial según normativa de licencias
# médicas en Chile) -> lista de patrones regex que lo reconocen.
#
# Notas de negocio:
# * "Enfermedad o Accidente Comun" va primero para capturar las variantes
#   con las palabras en cualquier orden ("accidente común", "común o
#   accidente", etc.).
# * "Ley SANNA" incluye variantes de "acompañamiento niño" y "condición
#   grave" (patrón duplicado a propósito en el histórico; se conserva).
RE_TIPO = {
    "Enfermedad o Accidente Comun": [r"enfermedad.*accidente|accidente.*comun|comun.*accidente|enfermedad.*comun"],
    "Medicina Preventiva": [r"medicina.*preventiva|preventiva|prorroga.*preventiva"],
    "Licencia maternal": [r"maternal|maternidad|pre.*natal|post.*natal"],
    "Licencia parental": [r"parental|paternidad"],
    "Enfermedad grave del hijo menor de un ano": [r"grave.*hijo|hijo.*grave|menor.*un.*ano|menor.*de.*1"],
    "Accidente del trabajo o del trayecto": [r"accidente.*trabajo|trabajo.*accidente|trayecto"],
    "Enfermedad profesional": [r"enfermedad.*profesional|profesional.*enfermedad"],
    "Patologia del embarazo": [r"patologia.*embarazo|embarazo.*patologia|sintoma.*aborto"],
    "Ley SANNA": [r"sanna|acompanamiento.*nino|acompanamiento.*nino|condicion.*grave"],
}

# INSTITUCIÓN DE SALUD (RB-04)
# FONASA, las Isapres vigentes, la Mutual de Seguridad y el ISP
# (Instituto de Salud Pública, para licencias de medicina preventiva).
#
# Notas de negocio:
# * "Banm[eé]dica" admite la tilde por si algún registro no pasó por norm().
# * "Nueva Masvida" incluye el patrón genérico "masvida": como se evalúa
#   en orden de inserción, cualquier "masvida" cae en esta categoría.
# * "Isapre Banco Estado" también reconoce "fundación" (Fundación Banestado).
RE_INST = {
    "Fonasa": [r"fonasa"],
    "Banmedica": [r"banm[eé]dica"],
    "Consalud": [r"consalud"],
    "Cruz Blanca": [r"cruz\s*blanca"],
    "Colmena": [r"colmena"],
    "Vida Tres": [r"vida\s*tres"],
    "Nueva Masvida": [r"nueva\s*masvida|masvida"],
    "Esencial": [r"esencial"],
    "Isapre Banco Estado": [r"banco\s*estado|fundaci[oó]n"],
    "Mutual": [r"mutual"],
    "ISP": [r"isp"],
}

# RESOLUCIÓN / ESTADO DE LA LICENCIA (RB-04 y RB-05)
# Estado de tramitación de la licencia ante COMPIN/Isapre.
# Se usan raíces de palabra ("autoriz", "rechaz", "anul", ...) para cubrir
# todas las flexiones: autorizada/autorizado/autoriza, etc.
# "Pendiente" incluye "tramit" (en tramitación) y la marca legacy "n/c".
RE_RESOL = {
    "Autorizada": [r"autoriz"],
    "Rechazada": [r"rechaz", r"anul"],
    "Ampliada": [r"ampli"],
    "Reducida": [r"reduc"],
    "Pendiente": [r"pendient", r"tramit", r"n/c"],
}

# ESTABLECIMIENTOS DEL SLEP (RB-08)
# Catálogo regex de los establecimientos del Servicio Local. La clave es el
# nombre canónico oficial (coincide con la tabla maestra
# ``Establecimientos.xlsx``); la lista contiene las variantes ortográficas
# observadas en la planilla madre histórica.
#
# Patrones de escritura frecuentes:
# * ``n[°o]?`` : tolera "n°334", "no 334", "n 334" (norm() ya convirtió
#   "n°" -> "n ").
# * ``d-?120`` : tolera "D-120" y "D120".
# * Alternativas con/sin apodo: "poeta", "dr.", "polivalente", etc.
#
# El último bloque ("Servicio Local Los Libertadores") es especial: no es un
# establecimiento educacional, sino el comodín para las unidades
# administrativas de la Dirección Ejecutiva (cualquier "unidad de ..." del
# SLEP se canoniza al Servicio Local).
RE_ESTABLECIMIENTO = {
    "Escuela #334 Luis Cruz Martínez": [
        r"esc\.?\s*n[°o]?\s*334\s*luis\s*cruz\s*martinez",
        r"escuela\s*n[°o]?\s*334\s*luis\s*cruz\s*martinez",
    ],
    "Escuela N°1414 Mercedes Fontecilla": [
        r"esc\.?\s*n[°o]?\s*1414\s*mercedes\s*fontecilla",
    ],
    "Escuela N°336 Estado de Michigan": [
        r"esc\.?\s*n[°o]?\s*336\s*estado\s*de\s*michigan",
    ],
    "Escuela N°337 El Mañío": [
        r"esc\.?\s*n[°o]?\s*337\s*el\s*manio",
        r"esc\.?\s*n[°o]?\s*337\s*el\s*manio",
    ],
    "Escuela D-120 Horacio Johnson": [
        r"escuela\s*d-?120\s*horacio\s*johnson",
        r"escuela\s*d-?120\s*horacio\s*johnson\s*gana",
    ],
    "Escuela D-124 Eloísa Díaz Insunza": [
        r"escuela\s*e-?124\s*dra?\.?\s*eloisa\s*diaz\s*insunza",
        r"escuela\s*d-?124\s*eloisa\s*diaz\s*insunza",
    ],
    "Escuela Profesor Humberto Aranda Iribarren": [
        r"escuela\s*especial\s*profesor\s*humberto\s*aranda\s*iribarren",
        r"escuela\s*profesor\s*humberto\s*aranda\s*iribarren",
        r"escuela\s*especial\s*profesor\s*humberto\s*aranda\s*iribarren\s*e-?153",
    ],
    "Escuela N°1584 María Luisa Sepúlveda": [
        r"escuela\s*n[°o]?\s*1584\s*maria\s*luisa\s*sepulveda",
    ],
    "Escuela N° 1963 Ana Frank": [
        r"escuela\s*n[°o]?\s*1963\s*ana\s*frank",
    ],
    "Jardín Infantil Rigoberto Puebla": [
        r"jardin\s*infantil\s*rigoberto\s*puebla\s*pizarro",
        r"jardin\s*infantil\s*rigoberto\s*puebla",
    ],
    "Liceo A-33 Federico García Lorca": [
        r"liceo\s*a-?33\s*poeta\s*federico\s*garcia\s*lorca",
        r"liceo\s*a-?33\s*federico\s*garcia\s*lorca",
    ],
    "Liceo A-41 Abdón Cifuentes": [
        r"liceo\s*a-?41\s*polivalente\s*abdon\s*cifuentes",
        r"liceo\s*a-?41\s*abdon\s*cifuentes",
    ],
    "Liceo D-135 Almirante Riveros": [
        r"liceo\s*d-?135\s*almirante\s*(galvarino\s*)?riveros",
    ],
    "Escuela D-339 Poeta Eusebio Lillo": [
        r"escuela\s*d-?339\s*poeta\s*eusebio\s*lillo",
    ],
    "Escuela D-338 Allipén": [
        r"escuela\s*d-?338\s*allipen",
    ],
    "Escuela D-151 Pedro Aguirre Cerda": [
        r"escuela\s*d-?151\s*pedro\s*aguirre\s*cerda",
    ],
    "Escuela D-144 Valle del Inca": [
        r"escuela\s*d-?144\s*valle\s*del\s*inca",
    ],
    "Escuela E-140 Likan Antai": [
        r"escuela\s*e-?140\slikan\s*antai",
    ],
    "Escuela D-139 Araucarias de Chile": [
        r"escuela\s*d-?139\s*araucarias\s*de\s*chile",
    ],
    "Escuela F-127 Camilo Henríquez": [
        r"escuela\s*f-?127\s*camilo\s*henriquez",
    ],
    "Escuela E-125 Aviador Dagoberto Godoy": [
        r"escuela\s*e-?125\s*aviador\s*dagoberto\s*godoy",
    ],
    "Escuela E-118 Atenea": [
        r"escuela\s*e-?118\s*atenea",
    ],
    "Escuela D-116 Sol Naciente": [
        r"escuela\s*d-?116\s*sol\s*naciente",
    ],
    "Escuela D-114 José Alejandro Bernales": [
        r"escuela\s*d-?114\s*jose\s*alejandro\s*bernales",
    ],
    "Escuela D-110 Unesco": [
        r"escuela\s*d-?110\s*unesco",
    ],
    "Escuela N°1668 Valle de la Luna": [
        r"escuela\s*n[°o]?\s*1668\s*valle\s*de\s*la\s*luna",
    ],
    "Escuela N°1968 Pucará Lasana": [
        r"escuela\s*n[°o]?\s*1968\s*pucara\s*lasana",
    ],
    "Jardín Infantil y sala cuna Doña Letizia": [
        r"jardin\s*infantil\s*y\s*sala\s*cuna\s*dona\s*letizia",
    ],
    "Jardín Infantil y sala cuna Allipén": [
        r"jardin\s*infantil\s*y\s*sala\s*cuna\s*allipen",
    ],
    "Jardín Infantil y sala cuna Ayenhué": [
        r"jardin\s*infantil\s*y\s*sala\s*cuna\s*ayenhue",
    ],
    "Jardín Infantil Ayin Antu": [
        r"jardin\s*infantil\s*ayin\s*antu",
    ],
    "Jardín Infantil Juan XXIII": [
        r"jardin\s*infantil\s*juan\s*xxiii",
    ],
    "Jardín Infantil Peumayén": [
        r"jardin\s*infantil\s*peumayen",
    ],
    "Jardín Infantil Elena Caffarena": [
        r"jardin\s*infantil\s*elena\s*caffarena",
    ],
    "Jardín Infantil Manuel Guerrero": [
        r"jardin\s*infantil\s*manuel\s*guerrero",
    ],
    "Jardín Infantil José Manuel Parada": [
        r"jardin\s*infantil\s*jose\s*manuel\s*parada",
    ],
    "Jardín Infantil Colmenita": [
        r"jardin\s*infantil\s*colmenita",
    ],
    "Jardín Infantil Los Molinos": [
        r"jardin\s*infantil\s*los\s*molinos",
    ],
    "Jardín Infantil Andrés Aylwin Azócar": [
        r"jardin\s*infantil\s*andres\s*aylwin\s*azocar",
    ],
    "Jardín Infantil Beato Padre Hurtado": [
        r"jardin\s*infantil\s*beato\s*padre\s*hurtado",
    ],
    "Jardín Infantil Hugo Marchant": [
        r"jardin\s*infantil\s*hugo\s*marchant",
    ],
    "Jardín Infantil Gabriela Mistral": [
        r"jardin\s*infantil\s*gabriela\s*mistral",
    ],
    "Jardín Infantil Parinacota": [
        r"jardin\s*infantil\s*parinacota",
    ],
    "Sala Cuna Pascual Gambino": [
        r"sala\s*cuna\s*pascual\s*gambino",
    ],
    "Liceo Técnico Profesional José Miguel Carrera": [
        r"liceo\s*tecnico\s*profesional\s*jose\s*miguel\s*carrera",
    ],
    "Liceo Municipal Alcalde Jorge Indo": [
        r"liceo\s*municipal\s*alcalde\s*jorge\s*indo",
    ],
    "Liceo Bicentenario Francisco Bilbao Barquín": [
        r"liceo\s*bicentenario\s*francisco\s*bilbao\s*barquin",
    ],
    "Liceo Municipal de Adultos Poeta Vicente Huidobro": [
        r"liceo\s*municipal\s*de\s*adultos\s*poeta\s*vicente\s*huidobro",
    ],
    "Servicio Local Los Libertadores": [
        r"direccion\s*ejecutiva",
        r"unidad\s*de\s*administracion\s*y\s*finanzas",
        r"unidad\s*de\s*comunicaciones",
        r"unidad\s*de\s*desarrollo\s*estrategico\s*y\s*gestion\s*de\s*la\s*informacion",
        r"unidad\s*de\s*gestion\s*juridica\s*y\s*transparencia",
        r"unidad\s*de\s*gestion\s*y\s*desarrollo\s*de\s*las\s*personas",
        r"unidad\s*de\s*planificacion\s*y\s*control\s*de\s*gestion",
        r"uatp\s*slep",
    ],
}

# AFP (RB-06 y RB-07)
# Mapea el nombre de AFP **ya normalizado** (minúsculas, sin tildes) hacia
# su nombre canónico. A diferencia de los mapas anteriores, la clave es el
# texto normalizado exacto, no una regex.
#
# Valores especiales de negocio:
# * "pensionado" -> el funcionario no cotiza en AFP; se registra como
#   "Pensionado (no aplica AFP)" con tasa 0.
# * "empart", "ee municipales de la republica", "sss": regímenes
#   previsionales antiguos/especiales presentes en el histórico.
AFP_MAP = {
    "modelo": "Modelo",
    "capital": "Capital",
    "provida": "Provida",
    "habitat": "Habitat",
    "plan vital": "Plan Vital",
    "uno": "Uno",
    "cuprum": "Cuprum",
    "pensionado": "Pensionado (no aplica AFP)",
    "empart": "Empart",
    "ee municipales de la republica": "EE Municipales de la República",
    "sss": "S.S.S",
}

# MESES (RB-12)
# Convierte el nombre del mes (tal como aparece en las columnas
# "MONTO DESCONTADO <MES> <AÑO>" de las hojas LM) a su número MM,
# para construir el período estándar ``YYYY-MM`` de Hechos_Descuentos.
MESES_MAP = {
    "ENERO": "01", "FEBRERO": "02", "MARZO": "03", "ABRIL": "04",
    "MAYO": "05", "JUNIO": "06", "JULIO": "07", "AGOSTO": "08",
    "SEPTIEMBRE": "09", "OCTUBRE": "10", "NOVIEMBRE": "11", "DICIEMBRE": "12",
}

# ERRORES DE EXCEL (RB-03)
# Valores de error cacheados por Excel (se leen con ``data_only=True``).
# Regla de negocio: cualquier celda con uno de estos valores se trata como
# **vacía**, porque corresponde a una fórmula rota en la planilla madre,
# no a un dato real.
EXCEL_ERRORS = {"#error!", "#n/a", "#ref!", "#value!", "#num!", "#name?", "#null!", "#div/0!", "#n/d"}
