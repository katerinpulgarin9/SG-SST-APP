# -*- coding: utf-8 -*-
"""Familias de documento: estructuras base cuando no hay plantilla (o queda corta)."""
from __future__ import annotations

from generador.normativa import texto_base_normativa


def contexto_actividad(empresa: dict) -> str:
    codigo = (empresa.get("ciiu_codigo") or "").strip()
    desc = (empresa.get("ciiu_descripcion") or "").strip()
    if codigo and desc:
        return f"actividad econ\u00f3mica CIIU {codigo} \u2014 {desc}"
    if desc:
        return desc
    if codigo:
        return f"actividad econ\u00f3mica CIIU {codigo}"
    return "la actividad econ\u00f3mica registrada por la empresa"


def peligros_sugeridos(empresa: dict) -> list[str]:
    texto = f"{empresa.get('ciiu_codigo','')} {empresa.get('ciiu_descripcion','')}".lower()
    base = [
        "Condiciones locativas en instalaciones de trabajo",
        "Riesgos biomec\u00e1nicos por posturas y manipulaci\u00f3n de cargas",
        "Riesgo el\u00e9ctrico en equipos e instalaciones",
        "Emergencias (incendio, evacuaci\u00f3n, primeros auxilios)",
        "Riesgos psicosociales asociados a la organizaci\u00f3n del trabajo",
    ]
    if any(k in texto for k in ("ganad", "bovin", "subasta", "pecuaria", "animal", "4610")):
        base.extend(
            [
                "Atropello o aplastamiento por semovientes",
                "Riesgo biol\u00f3gico (zoonosis, fluidos, excrementos, mordeduras/coces)",
                "Cargue y descargue de animales en embarcaderos",
                "Ruido en remates o \u00e1reas de operaci\u00f3n",
                "Estampidas y emergencias asociadas a manejo de ganado",
            ]
        )
    elif any(k in texto for k in ("construcc", "obra", "edific")):
        base.extend(
            [
                "Trabajo en alturas",
                "Ca\u00eddas a distinto nivel y mismo nivel",
                "Manejo de herramientas y maquinaria de obra",
                "Exposici\u00f3n a material particulado",
            ]
        )
    elif any(k in texto for k in ("oficina", "admin", "servicios")):
        base.extend(
            [
                "Fatiga visual por pantallas",
                "Trastornos musculoesquel\u00e9ticos por trabajo de escritorio",
            ]
        )
    else:
        base.append(
            f"Peligros propios de la {contexto_actividad(empresa)}, a detallar con la "
            "participaci\u00f3n de los trabajadores"
        )
    return base


def _roles(empresa: dict) -> list[str]:
    return [
        f"Alta direcci\u00f3n / {empresa.get('rep_legal_cargo') or 'Representante legal'}: "
        f"{empresa.get('rep_legal_nombre') or 'Por definir'} \u2014 aprueba el documento, "
        "asigna recursos y hace seguimiento al SG-SST.",
        f"Responsable SST ({empresa.get('resp_sst_cargo') or 'Responsable SST'}): "
        f"{empresa.get('resp_sst_nombre') or 'Por definir'} \u2014 elabora, actualiza, "
        "socializa y verifica la aplicaci\u00f3n de este documento.",
        f"Vig\u00eda / COPASST: {(empresa.get('vigia_nombre') or '').strip() or 'Por definir'} "
        "\u2014 participa en la identificaci\u00f3n de peligros, promueve la cultura de "
        "prevenci\u00f3n y revisa avances.",
        "Jefes de \u00e1rea y trabajadores: cumplen controles, reportan condiciones "
        "inseguras e incidentes, y participan en capacitaciones e inspecciones.",
    ]


def _marco_normativo(norma: str) -> list[str]:
    return [
        norma,
        "Este documento es una base operativa del SG-SST. Debe complementarse con la "
        "realidad de cada centro de trabajo, la matriz de peligros vigente y los "
        "requisitos espec\u00edficos de la ARL y de la autoridad competente.",
        "La empresa conserva la versi\u00f3n controlada y registra cambios en el control "
        "maestro documental.",
    ]


def _seguimiento_mejora(empresa: dict) -> list[str]:
    return [
        "Indicadores sugeridos: cumplimiento del cronograma, hallazgos cerrados, "
        "cobertura de capacitaci\u00f3n, tasa de accidentalidad e investigaciones a tiempo.",
        "Frecuencia de revisi\u00f3n: al menos una vez al a\u00f1o, o cuando cambien procesos, "
        "normativa, infraestructura o resultados de auditor\u00eda / inspecci\u00f3n.",
        f"La alta direcci\u00f3n de {empresa.get('razon_social') or 'la empresa'} revisa "
        "resultados y define acciones de mejora continua.",
        "Las no conformidades generan plan de acci\u00f3n con responsable, fecha y evidencia.",
    ]


def _registros_evidencias(titulo: str) -> list[str]:
    return [
        f"Registro de elaboraci\u00f3n, revisi\u00f3n y aprobaci\u00f3n de \u00ab{titulo}\u00bb.",
        "Listas de asistencia a socializaci\u00f3n y capacitaci\u00f3n relacionadas.",
        "Actas de COPASST / Vig\u00eda, inspecciones, reportes de actos/condiciones inseguras.",
        "Investigaciones de incidentes/accidentes y planes de acci\u00f3n derivados.",
        "Evidencias fotogr\u00e1ficas, checklists y soportes de entrega de EPP cuando aplique.",
    ]


def _anexos_sugeridos() -> list[str]:
    return [
        "Matriz de identificaci\u00f3n de peligros, evaluaci\u00f3n y valoraci\u00f3n de riesgos (IPVR).",
        "Plan de trabajo anual del SG-SST y cronograma de capacitaciones.",
        "Plan de prevenci\u00f3n, preparaci\u00f3n y respuesta ante emergencias.",
        "Formatos de inspecci\u00f3n, entrega de EPP e investigaci\u00f3n de incidentes.",
        "Pol\u00edtica de SST y organigrama / designaciones vigentes.",
    ]


def secciones_familia(familia: str, empresa: dict, titulo: str) -> list[tuple[str, list[str]]]:
    act = contexto_actividad(empresa)
    norma = texto_base_normativa()
    peligros = peligros_sugeridos(empresa)
    clases = ", ".join(empresa.get("clases_riesgo_list") or []) or (
        empresa.get("clases_riesgo") or "seg\u00fan afiliaci\u00f3n ARL"
    )
    razon = empresa.get("razon_social") or "La empresa"
    arl = empresa.get("arl") or "la ARL afiliada"

    common_intro = [
        f"Documento: \u00ab{titulo}\u00bb.",
        f"Empresa: {razon} \u2014 NIT {empresa.get('nit', '')}.",
        f"Aplicable a la {act}.",
        f"ARL: {arl}. Clase(s) de riesgo: {clases}.",
    ]

    fam = (familia or "").strip().lower()

    if fam == "politica":
        return [
            ("1. Objetivo", common_intro + [
                f"Establecer la pol\u00edtica de Seguridad y Salud en el Trabajo de {razon}, "
                f"como compromiso de la alta direcci\u00f3n frente a la {act}.",
                "Definir lineamientos de prevenci\u00f3n, cumplimiento legal, participaci\u00f3n "
                "de trabajadores y mejora continua del SG-SST.",
            ]),
            ("2. Alcance", [
                "Aplica a todos los trabajadores directos, en misi\u00f3n, contratistas, "
                "proveedores y visitantes que ingresen a instalaciones o participen en "
                "actividades bajo control de la empresa.",
                "Cubre sedes, frentes de trabajo, \u00e1reas administrativas y operativas "
                "relacionadas con la actividad econ\u00f3mica.",
            ]),
            ("3. Marco normativo", _marco_normativo(norma)),
            ("4. Compromisos de la pol\u00edtica", [
                f"{razon} se compromete a proteger la vida, la salud y la integridad "
                "f\u00edsica y mental de las personas en el trabajo.",
                "Identificar peligros, evaluar y valorar riesgos, e implementar controles "
                "en la fuente, el medio y el individuo, priorizando la jerarqu\u00eda de controles.",
                "Cumplir la normatividad colombiana en riesgos laborales y los est\u00e1ndares "
                "m\u00ednimos aplicables seg\u00fan tama\u00f1o y clase de riesgo.",
                "Asignar recursos humanos, t\u00e9cnicos y financieros para el SG-SST.",
                "Promover la consulta, participaci\u00f3n y formaci\u00f3n de los trabajadores.",
                "Investigar incidentes y accidentes, adoptar acciones correctivas y "
                "preventivas, y difundir lecciones aprendidas.",
                "Revisar esta pol\u00edtica al menos una vez al a\u00f1o y actualizarla cuando "
                "cambien procesos, normativa o resultados del sistema.",
            ]),
            ("5. Objetivos del SG-SST asociados", [
                "Mantener actualizada la matriz IPVR con participaci\u00f3n de los trabajadores.",
                "Ejecutar el plan de trabajo anual y el plan de emergencias.",
                "Capacitar al personal en peligros prioritarios de la actividad.",
                "Hacer seguimiento a indicadores de accidentalidad, ausentismo y cumplimiento.",
                "Garantizar la disponibilidad de EPP, se\u00f1alizaci\u00f3n y condiciones locativas seguras.",
            ]),
            ("6. Roles y responsabilidades", _roles(empresa)),
            ("7. Difusi\u00f3n y comunicaci\u00f3n", [
                "La pol\u00edtica se socializa en inducci\u00f3n, reinducci\u00f3n, carteleras, "
                "reuniones de seguridad y medios digitales internos.",
                "Debe estar disponible y visible para trabajadores y partes interesadas "
                "que la requieran.",
            ]),
            ("8. Registros y evidencias", _registros_evidencias(titulo)),
            ("9. Seguimiento y mejora", _seguimiento_mejora(empresa)),
            ("10. Anexos sugeridos", _anexos_sugeridos()),
        ]

    if fam == "acta":
        return [
            ("1. Objetivo", common_intro + [
                f"Dejar constancia formal del acto administrativo o decisi\u00f3n "
                f"relacionada con \u00ab{titulo}\u00bb en el marco del SG-SST de {razon}.",
            ]),
            ("2. Alcance", [
                "Aplica a la designaci\u00f3n, constancia o acuerdo objeto de esta acta "
                "y a las personas e instancias mencionadas.",
            ]),
            ("3. Marco normativo", _marco_normativo(norma)),
            ("4. Hechos / designaci\u00f3n", [
                f"En desarrollo de la {act}, {razon} deja constancia de lo relacionado "
                f"con \u00ab{titulo}\u00bb.",
                f"Se reconoce como Responsable SST a "
                f"{empresa.get('resp_sst_nombre') or '[NOMBRE RESPONSABLE SST]'} "
                f"({empresa.get('resp_sst_cargo') or 'Responsable SST'}), quien coordinar\u00e1 "
                "la implementaci\u00f3n, mantenimiento y mejora del SG-SST.",
                f"Representante legal / alta direcci\u00f3n: "
                f"{empresa.get('rep_legal_nombre') or 'Por definir'} "
                f"({empresa.get('rep_legal_cargo') or 'Representante legal'}).",
                f"Vig\u00eda / representante de los trabajadores: "
                f"{(empresa.get('vigia_nombre') or '').strip() or 'Por definir'}.",
            ]),
            ("5. Funciones y compromisos", [
                "Coordinar la documentaci\u00f3n del SG-SST y el control maestro.",
                "Identificar peligros, valorar y controlar riesgos propios de la actividad.",
                "Promover capacitaci\u00f3n, inspecciones, investigaci\u00f3n de incidentes "
                "y respuesta ante emergencias.",
                "Reportar avances e indicadores a la alta direcci\u00f3n y coordinar con la ARL.",
                "Garantizar la participaci\u00f3n de los trabajadores en el sistema.",
            ]),
            ("6. Roles y responsabilidades", _roles(empresa)),
            ("7. Vigencia", [
                "Esta acta rige a partir de la fecha de firma y permanece vigente hasta "
                "nueva designaci\u00f3n, renuncia o actualizaci\u00f3n formal.",
            ]),
            ("8. Registros y evidencias", _registros_evidencias(titulo)),
            ("9. Seguimiento y mejora", _seguimiento_mejora(empresa)),
            ("10. Anexos sugeridos", _anexos_sugeridos()),
        ]

    if fam == "procedimiento":
        return [
            ("1. Objetivo", common_intro + [
                f"Establecer el procedimiento \u00ab{titulo}\u00bb para estandarizar la forma "
                f"segura y controlada de ejecutar actividades asociadas a la {act}.",
            ]),
            ("2. Alcance", [
                "Aplica a trabajadores, contratistas y dem\u00e1s personas que intervengan "
                "en las operaciones descritas en este procedimiento.",
                "Incluye actividades rutinarias y no rutinarias relacionadas con el objeto "
                "del documento, en las instalaciones y frentes de trabajo de la empresa.",
            ]),
            ("3. Marco normativo", _marco_normativo(norma)),
            ("4. Definiciones", [
                "Peligro: fuente, situaci\u00f3n o acto con potencial de daño.",
                "Riesgo: combinaci\u00f3n de la probabilidad de que ocurra un evento "
                "peligroso y la severidad de la lesi\u00f3n o enfermedad.",
                "Control: medida para eliminar el peligro o reducir el riesgo "
                "(fuente, medio, individuo).",
                "EPP: elemento de protecci\u00f3n personal requerido seg\u00fan la tarea.",
                "Incidente / accidente de trabajo: evento que pudo causar o caus\u00f3 "
                "lesi\u00f3n, conforme a la normativa de riesgos laborales.",
            ]),
            ("5. Roles y responsabilidades", _roles(empresa)),
            ("6. Condiciones generales / requisitos previos", [
                "Verificar inducci\u00f3n / reinducci\u00f3n y competencia para la tarea.",
                "Disponer de EPP, herramientas e instalaciones en condiciones seguras.",
                "Conocer rutas de evacuaci\u00f3n, puntos de encuentro y n\u00fameros de emergencia.",
                "Reportar de inmediato condiciones inseguras antes de iniciar la labor.",
            ]),
            ("7. Desarrollo del procedimiento", [
                "Paso 1. Identificar la necesidad u operaci\u00f3n cr\u00edtica asociada a "
                f"\u00ab{titulo}\u00bb y a la {act}.",
                "Paso 2. Revisar peligros y controles definidos en la matriz IPVR "
                "para el proceso/actividad.",
                "Paso 3. Preparar el \u00e1rea, equipos y EPP; se\u00f1alizar y delimitar "
                "si aplica.",
                "Paso 4. Ejecutar la actividad aplicando controles en fuente, medio "
                "e individuo; no improvisar atajos que aumenten el riesgo.",
                "Paso 5. Supervisar el cumplimiento de instrucciones y detener el "
                "trabajo ante peligro inminente (derecho a rehusarse).",
                "Paso 6. Registrar hallazgos, desviaciones, casi-accidentes o "
                "incidentes y reportarlos al Responsable SST.",
                "Paso 7. Al finalizar, ordenar el \u00e1rea, devolver equipos y "
                "dejar evidencias del control realizado.",
                "Paso 8. Revisar peri\u00f3dicamente la efectividad del procedimiento "
                "y actualizarlo cuando cambien procesos o normativa.",
            ]),
            ("8. Peligros asociados (referencia)", [
                "Los peligros prioritarios a considerar en este procedimiento incluyen:",
                *[f"- {p}" for p in peligros],
            ]),
            ("9. Controles y recursos", [
                "Controles ingenieriles y administrativos definidos en la matriz IPVR.",
                "EPP espec\u00edfico de la tarea (casco, guantes, calzado, protecci\u00f3n "
                "visual/auditiva u otros, seg\u00fan evaluaci\u00f3n).",
                f"Apoyo de {arl} en promoci\u00f3n y prevenci\u00f3n cuando se requiera.",
                "Se\u00f1alizaci\u00f3n, orden y aseo, y mantenimiento preventivo de equipos.",
            ]),
            ("10. Registros y evidencias", _registros_evidencias(titulo)),
            ("11. Seguimiento y mejora", _seguimiento_mejora(empresa)),
            ("12. Documentos relacionados / anexos", _anexos_sugeridos()),
        ]

    if fam == "plan":
        return [
            ("1. Objetivo", common_intro + [
                f"Definir el plan \u00ab{titulo}\u00bb para prevenir, controlar y dar "
                f"respuesta a los riesgos propios de la {act} en {razon}.",
            ]),
            ("2. Alcance", [
                "Cubre instalaciones, procesos, personal administrativo y operativo, "
                "contratistas y visitantes bajo control de la empresa.",
                "Incluye actividades de planificaci\u00f3n, ejecuci\u00f3n, verificaci\u00f3n "
                "y mejora asociadas al objeto del plan.",
            ]),
            ("3. Marco normativo", _marco_normativo(norma)),
            ("4. Diagn\u00f3stico y peligros / escenarios prioritarios", [
                "Con base en la actividad econ\u00f3mica y el contexto de la empresa, "
                "se priorizan los siguientes peligros / escenarios:",
                *[f"- {p}" for p in peligros],
                "Este listado debe validarse con la matriz IPVR vigente y la "
                "participaci\u00f3n de los trabajadores.",
            ]),
            ("5. Roles y responsabilidades", _roles(empresa)),
            ("6. Estrategias y acciones del plan", [
                "Actualizar la identificaci\u00f3n de peligros y la valoraci\u00f3n de riesgos.",
                "Programar intervenciones (controles, mantenimientos, se\u00f1alizaci\u00f3n, EPP).",
                "Ejecutar capacitaciones, simulacros e inducciones seg\u00fan cronograma.",
                "Realizar inspecciones peri\u00f3dicas y cerrar hallazgos con acciones correctivas.",
                "Coordinar con la ARL recursos de promoci\u00f3n y prevenci\u00f3n.",
                "Mantener listos los elementos y protocolos de emergencia aplicables.",
                f"Socializar \u00ab{titulo}\u00bb a todo el personal involucrado.",
            ]),
            ("7. Cronograma y recursos (orientaci\u00f3n)", [
                "Corto plazo (0\u20133 meses): socializaci\u00f3n, ajustes de matriz, "
                "inspecciones iniciales y priorizaci\u00f3n de controles cr\u00edticos.",
                "Mediano plazo (3\u20136 meses): capacitaciones, simulacros, cierre de "
                "hallazgos y mejora de condiciones locativas.",
                "Anual: revisi\u00f3n por la alta direcci\u00f3n, actualizaci\u00f3n del plan "
                "y alineaci\u00f3n con el plan de trabajo del SG-SST.",
                "Recursos: tiempo de personal, presupuesto de SST, apoyo ARL, "
                "proveedores de EPP y mantenimiento.",
            ]),
            ("8. Indicadores de seguimiento", [
                "% de actividades del plan ejecutadas vs programadas.",
                "N\u00famero de hallazgos abiertos / cerrados a tiempo.",
                "Cobertura de personal capacitado / simulado.",
                "Incidentes y accidentes relacionados con los escenarios del plan.",
            ]),
            ("9. Registros y evidencias", _registros_evidencias(titulo)),
            ("10. Seguimiento y mejora", _seguimiento_mejora(empresa)),
            ("11. Anexos sugeridos", _anexos_sugeridos()),
        ]

    if fam == "programa":
        return [
            ("1. Objetivo", common_intro + [
                f"Orientar el programa \u00ab{titulo}\u00bb para desarrollar actividades "
                f"peri\u00f3dicas de prevenci\u00f3n y promoci\u00f3n alineadas a la {act}.",
            ]),
            ("2. Alcance", [
                "Aplica al personal objetivo del programa (operativo, administrativo "
                "y/o contratistas, seg\u00fan se defina en la socializaci\u00f3n).",
                f"Cubre las sedes y frentes donde {razon} ejerce control de las actividades.",
            ]),
            ("3. Marco normativo", _marco_normativo(norma)),
            ("4. Roles y responsabilidades", _roles(empresa)),
            ("5. Actividades del programa", [
                "Diagnosticar necesidades (matriz IPVR, accidentalidad, inspecciones, "
                "requerimientos legales).",
                "Elaborar cronograma de actividades (diarias, semanales, mensuales o anuales).",
                "Asignar responsables, poblaci\u00f3n objetivo, recursos y evidencias esperadas.",
                "Ejecutar capacitaciones, pausas activas, inspecciones, campañas u otras "
                f"acciones propias de \u00ab{titulo}\u00bb.",
                "Registrar asistencia, hallazgos y resultados; archivar evidencias.",
                "Evaluar cumplimiento y efectividad; ajustar el programa.",
            ]),
            ("6. Peligros / temas prioritarios", [
                "Temas a abordar de forma prioritaria:",
                *[f"- {p}" for p in peligros],
            ]),
            ("7. Recursos", [
                "Tiempo de trabajadores y responsables SST.",
                f"Apoyo t\u00e9cnico de {arl} cuando corresponda.",
                "Material did\u00e1ctico, EPP de demostraci\u00f3n, se\u00f1alizaci\u00f3n y espacios "
                "adecuados para formaci\u00f3n.",
            ]),
            ("8. Registros y evidencias", _registros_evidencias(titulo)),
            ("9. Seguimiento y mejora", _seguimiento_mejora(empresa)),
            ("10. Anexos sugeridos", _anexos_sugeridos()),
        ]

    if fam in {"matriz", "formato", "evaluacion"}:
        return [
            ("1. Objetivo", common_intro + [
                f"Disponer el documento \u00ab{titulo}\u00bb como instrumento de registro, "
                f"control o evaluaci\u00f3n del SG-SST para la {act}.",
            ]),
            ("2. Alcance", [
                "Aplica a los procesos, \u00e1reas y personas que deban diligenciar, "
                "revisar o consolidar la informaci\u00f3n de este formato / matriz / evaluaci\u00f3n.",
            ]),
            ("3. Marco normativo", _marco_normativo(norma)),
            ("4. Roles y responsabilidades", _roles(empresa)),
            ("5. Instrucciones de diligenciamiento", [
                "Completar todos los campos obligatorios con letra legible o de forma digital.",
                "Usar la informaci\u00f3n vigente de la empresa (raz\u00f3n social, NIT, ARL, CIIU).",
                "Registrar fechas, responsables y evidencias asociadas a cada \u00edtem.",
                "No dejar filas cr\u00edticas en blanco; si no aplica, indicar N/A con justificaci\u00f3n.",
                "Conservar el registro firmado o con trazabilidad electr\u00f3nica en el "
                "archivo del SG-SST.",
            ]),
            ("6. Contenido / criterios de referencia", [
                "Peligros o \u00edtems de referencia para la actividad:",
                *[f"- {p}" for p in peligros],
                "Ajustar filas y columnas seg\u00fan el proceso real de la empresa.",
            ]),
            ("7. Frecuencia de uso", [
                "Seg\u00fan lo definido en el plan de trabajo anual, inspecciones programadas "
                "o cuando ocurra un cambio significativo en procesos o instalaciones.",
            ]),
            ("8. Registros y evidencias", _registros_evidencias(titulo)),
            ("9. Seguimiento y mejora", _seguimiento_mejora(empresa)),
            ("10. Anexos sugeridos", _anexos_sugeridos()),
        ]

    # Gen\u00e9rico (manual, instructivo, otro)
    return [
        ("1. Objetivo", common_intro + [
            f"Establecer el contenido base del documento \u00ab{titulo}\u00bb para "
            f"soportar la gesti\u00f3n del SG-SST en {razon}, en el marco de la {act}.",
        ]),
        ("2. Alcance", [
            "Aplica a las \u00e1reas, procesos y personas vinculadas al objeto del documento.",
            "Debe adaptarse a la estructura organizacional y a los peligros reales "
            "identificados en la matriz IPVR.",
        ]),
        ("3. Marco normativo", _marco_normativo(norma)),
        ("4. Roles y responsabilidades", _roles(empresa)),
        ("5. Desarrollo / contenido del documento", [
            f"Descripci\u00f3n del prop\u00f3sito de \u00ab{titulo}\u00bb dentro del ciclo "
            "Planear-Hacer-Verificar-Actuar del SG-SST.",
            "Condiciones de aplicaci\u00f3n, exclusiones y relacionamiento con otros "
            "documentos del sistema.",
            "Pasos, criterios o lineamientos que el personal debe seguir.",
            "Controles preventivos asociados (fuente, medio, individuo) y EPP cuando aplique.",
            "Mecanismos de reporte de desviaciones, incidentes y oportunidades de mejora.",
        ]),
        ("6. Peligros de referencia para la actividad", [
            *[f"- {p}" for p in peligros],
        ]),
        ("7. Registros y evidencias", _registros_evidencias(titulo)),
        ("8. Seguimiento y mejora", _seguimiento_mejora(empresa)),
        ("9. Anexos sugeridos", _anexos_sugeridos()),
    ]


COLUMNAS_EXCEL = {
    "evaluacion": [
        "Ciclo",
        "Est\u00e1ndar",
        "\u00cdtem",
        "Descripci\u00f3n",
        "Peso %",
        "Cumple",
        "No cumple",
        "N/A",
        "Observaciones",
    ],
    "plan": ["#", "Actividad", "Responsable", "Mes", "Recursos", "Indicador", "Estado"],
    "programa": [
        "#",
        "Tema / Actividad",
        "Poblaci\u00f3n",
        "Fecha prog.",
        "Duraci\u00f3n",
        "Responsable",
        "Estado",
    ],
    "formato": ["#", "\u00cdtem / Campo", "Registro", "Fecha", "Responsable", "Observaciones"],
    "matriz": ["#", "Proceso", "Actividad", "Peligro", "Riesgo", "Controles", "Responsable", "Fecha"],
}
