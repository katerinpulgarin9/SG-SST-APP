# -*- coding: utf-8 -*-
"""Referencias normativas del SG-SST (fácil de actualizar)."""

NORMATIVA = [
    {
        "nombre": "Resolución 0312 de 2019",
        "descripcion": "Estándares Mínimos del Sistema de Gestión de Seguridad y Salud en el Trabajo SG-SST.",
        "cita_corta": "Resolución 0312 de 2019",
    },
    {
        "nombre": "Decreto 1072 de 2015",
        "descripcion": "Decreto Único Reglamentario del Sector Trabajo. Libro 2, Parte 2, Título 4, Capítulo 6 — Sistema de Gestión de la Seguridad y Salud en el Trabajo.",
        "cita_corta": "Decreto 1072 de 2015",
    },
    {
        "nombre": "Ley 1562 de 2012",
        "descripcion": "Sistema General de Riesgos Laborales y disposiciones en materia de salud ocupacional.",
        "cita_corta": "Ley 1562 de 2012",
    },
]


def citas_normativas():
    return "; ".join(n["cita_corta"] for n in NORMATIVA)


def texto_base_normativa():
    return (
        "Este documento se elabora en cumplimiento de la "
        + citas_normativas()
        + ", aplicables al Sistema de Gestión de Seguridad y Salud en el Trabajo."
    )
