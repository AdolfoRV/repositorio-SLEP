"""
Constantes de clasificacion para el Migrador SLEP.

Todas las expresiones regulares y mapas de canonizacion viven aqui
"""

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

RE_RESOL = {
    "Autorizada": [r"autoriz"],
    "Rechazada": [r"rechaz", r"anul"],
    "Ampliada": [r"ampli"],
    "Reducida": [r"reduc"],
    "Pendiente": [r"pendient", r"tramit", r"n/c"],
}

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

MESES_MAP = {
    "ENERO": "01", "FEBRERO": "02", "MARZO": "03", "ABRIL": "04",
    "MAYO": "05", "JUNIO": "06", "JULIO": "07", "AGOSTO": "08",
    "SEPTIEMBRE": "09", "OCTUBRE": "10", "NOVIEMBRE": "11", "DICIEMBRE": "12",
}

EXCEL_ERRORS = {"#error!", "#n/a", "#ref!", "#value!", "#num!", "#name?", "#null!", "#div/0!", "#n/d"}