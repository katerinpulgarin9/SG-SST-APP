# -*- coding: utf-8 -*-
"""Familias de documento: estructuras base cuando no hay plantilla."""
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
            f"Peligros propios de la {contexto_actividad(empresa)}, a detallar con la participaci\u00f3n de los trabajadores"
        )
    return base


def secciones_familia(familia: str, empresa: dict, titulo: str) -> list[tuple[str, list[str]]]:
    act = contexto_actividad(empresa)
    norma = texto_base_normativa()
    peligros = peligros_sugeridos(empresa)
    clases = ", ".join(empresa.get("clases_riesgo_list") or []) or (
        empresa.get("clases_riesgo") or "seg\u00fan afiliaci\u00f3n ARL"
    )

    common_intro = [
        norma,
        f"Empresa: {empresa.get('razon_social','')} \u2014 NIT {empresa.get('nit','')}.",
        f"Documento aplicable a la {act}.",
        f"ARL: {empresa.get('arl','')}. Clase(s) de riesgo: {clases}.",
    ]

    if familia == "politica":
        return [
            (
                "1. Pol\u00edtica de Seguridad y Salud en el Trabajo",
                common_intro
                + [
                    f"{empresa.get('razon_social','')} se compromete a proteger la vida, salud e integridad de trabajadores, contratistas y visitantes, en el marco de la {act}.",
                    "Se cumple la normatividad vigente, se identifican peligros, se valoran riesgos y se implementan controles, promoviendo una cultura de prevenci\u00f3n y autocuidado.",
                    "La pol\u00edtica es de cumplimiento obligatorio y se revisa al menos anualmente.",
                ],
            ),
            (
                "2. Objetivos del SG-SST",
                [
                    "Cumplir los est\u00e1ndares m\u00ednimos aplicables seg\u00fan el tama\u00f1o de la empresa (Resoluci\u00f3n 0312 de 2019).",
                    "Capacitar al personal en peligros prioritarios de la actividad econ\u00f3mica.",
                    "Mantener actualizada la identificaci\u00f3n de peligros y el plan de emergencias.",
                    "Hacer seguimiento a indicadores de accidentalidad y al plan de trabajo anual.",
                ],
            ),
        ]

    if familia == "acta":
        return [
            (
                "1. Objeto",
                common_intro
                + [
                    "Designar formalmente al responsable del SG-SST conforme a la Resoluci\u00f3n 0312 de 2019.",
                ],
            ),
            (
                "2. Designaci\u00f3n",
                [
                    f"Se designa a {empresa.get('resp_sst_nombre') or '[NOMBRE RESPONSABLE SST]'} "
                    f"como {empresa.get('resp_sst_cargo') or 'Responsable SST'}, "
                    f"quien coordinar\u00e1 la implementaci\u00f3n y mejora del SG-SST de {empresa.get('razon_social','')}.",
                ],
            ),
            (
                "3. Funciones principales",
                [
                    "Coordinar el SG-SST y la documentaci\u00f3n del sistema.",
                    "Identificar peligros, valorar y controlar riesgos propios de la actividad.",
                    "Promover capacitaci\u00f3n, inspecciones, investigaci\u00f3n de incidentes y respuesta ante emergencias.",
                    "Reportar avances e indicadores a la alta direcci\u00f3n y a la ARL.",
                ],
            ),
        ]

    if familia == "procedimiento":
        return [
            ("1. Objetivo", common_intro + [f"Establecer el procedimiento \u00ab{titulo}\u00bb aplicable a la {act}."]),
            ("2. Alcance", ["Aplica a trabajadores, contratistas y dem\u00e1s personas que intervengan en las actividades descritas."]),
            (
                "3. Responsables",
                [
                    f"Responsable SST: {empresa.get('resp_sst_nombre') or 'Por definir'}.",
                    f"Alta direcci\u00f3n / {empresa.get('rep_legal_cargo') or 'Representante legal'}: {empresa.get('rep_legal_nombre') or 'Por definir'}.",
                ],
            ),
            (
                "4. Desarrollo",
                [
                    "Identificar la necesidad u operaci\u00f3n cr\u00edtica asociada a la actividad econ\u00f3mica.",
                    "Aplicar controles preventivos (fuente, medio e individuo) y EPP cuando corresponda.",
                    "Registrar hallazgos, incidentes o desviaciones y reportarlos al Responsable SST.",
                    "Revisar peri\u00f3dicamente la efectividad del procedimiento.",
                ],
            ),
            (
                "5. Documentos relacionados",
                ["Matriz IPVR, plan de emergencias, formatos de inspecci\u00f3n y plan de trabajo anual del SG-SST."],
            ),
        ]

    if familia == "plan":
        return [
            ("1. Objetivo", common_intro + [f"Definir el plan \u00ab{titulo}\u00bb para prevenir y controlar riesgos de la {act}."]),
            ("2. Alcance", ["Cubre instalaciones, procesos y personal de la empresa, seg\u00fan roles administrativos y operativos."]),
            ("3. Peligros / escenarios prioritarios", peligros),
            (
                "4. Acciones e intervenciones",
                [
                    "Actualizar la matriz de peligros con participaci\u00f3n de los trabajadores.",
                    "Ejecutar capacitaciones y simulacros seg\u00fan el plan anual.",
                    "Realizar inspecciones y cerrar hallazgos con acciones correctivas.",
                    "Coordinar con la ARL recursos de promoci\u00f3n y prevenci\u00f3n.",
                ],
            ),
            ("5. Seguimiento", ["Indicadores, cronograma y revisi\u00f3n por la alta direcci\u00f3n al menos una vez al a\u00f1o."]),
        ]

    if familia == "programa":
        return [
            ("1. Objetivo", common_intro + [f"Orientar el programa \u00ab{titulo}\u00bb alineado a la {act}."]),
            (
                "2. Actividades",
                [
                    "Programar actividades peri\u00f3dicas (diarias, semanales, mensuales o anuales seg\u00fan el caso).",
                    "Asignar responsables y recursos.",
                    "Registrar evidencias (listas de asistencia, inspecciones, mantenimientos).",
                ],
            ),
            ("3. Evaluaci\u00f3n", ["Verificar cumplimiento del cronograma y efectividad de las medidas."]),
        ]

    return [
        (
            "Contenido base",
            common_intro
            + [
                f"Documento tipo \u00ab{titulo}\u00bb para registro y control del SG-SST.",
                "Los campos espec\u00edficos se diligencian en la tabla del archivo generado.",
                "Peligros de referencia para la actividad:",
                *peligros,
            ],
        ),
    ]


COLUMNAS_EXCEL = {
    "evaluacion": ["Ciclo", "Est\u00e1ndar", "\u00cdtem", "Descripci\u00f3n", "Peso %", "Cumple", "No cumple", "N/A", "Observaciones"],
    "plan": ["#", "Actividad", "Responsable", "Mes", "Recursos", "Indicador", "Estado"],
    "programa": ["#", "Tema / Actividad", "Poblaci\u00f3n", "Fecha prog.", "Duraci\u00f3n", "Responsable", "Estado"],
    "formato": ["#", "\u00cdtem / Campo", "Registro", "Fecha", "Responsable", "Observaciones"],
    "matriz": ["#", "Proceso", "Actividad", "Peligro", "Riesgo", "Controles", "Responsable", "Fecha"],
}
