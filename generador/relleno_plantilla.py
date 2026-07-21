# -*- coding: utf-8 -*-
"""Rellena plantillas Word/Excel con datos de empresa sin perder su estructura.

Las plantillas del pack suelen traer datos de ejemplo (FOREST GREEN, NIT, etc.).
Este modulo los reemplaza por la empresa activa configurada en la app.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

from generador.familias_documento import contexto_actividad

# Datos de ejemplo frecuentes en el pack GESTION DOCUMENTAL SST
_EMPRESA_EJEMPLO_REGEX = [
    re.compile(r"FOREST\s+GREEN\s+SERVICIOS\s+AMBIENTALES\s+S\.?\s*A\.?\s*S\.?", re.I),
    re.compile(r"FOREST\s+GREEN\s+SERVICIOS\s+AMBIENTALES", re.I),
    re.compile(r"\bFOREST\s+GREEN\b", re.I),
    re.compile(r"Laboratorios\s+LT\s+S\.?\s*A\.?\s*S\.?", re.I),
    re.compile(r"\bXXXXXX(?:\s+XXXXXX)+\b", re.I),
    re.compile(r"\[?\s*NOMBRE\s+DE\s+LA\s+EMPRESA\s*\]?", re.I),
    re.compile(r"\[?\s*RAZ[O\u00d3]N\s+SOCIAL\s*\]?", re.I),
]

_NIT_EJEMPLO_REGEX = [
    re.compile(r"\b901[\.\s]?989[\.\s]?693\s*-?\s*7\b"),
    re.compile(r"\b901989693-?7\b"),
]


def _ctx_val(ctx: dict[str, str], key: str, fallback: str = "________________") -> str:
    v = (ctx.get(key) or "").strip()
    return v if v else fallback


def _contexto(empresa: dict, doc_meta: dict, version: int, fecha: str) -> dict[str, str]:
    clases = ", ".join(empresa.get("clases_riesgo_list") or []) or (
        empresa.get("clases_riesgo") or ""
    )
    return {
        "razon_social": (empresa.get("razon_social") or "").strip(),
        "empresa": (empresa.get("razon_social") or "").strip(),
        "nit": (empresa.get("nit") or "").strip(),
        "arl": (empresa.get("arl") or "").strip(),
        "ciiu": f"{empresa.get('ciiu_codigo', '')} {empresa.get('ciiu_descripcion', '')}".strip(),
        "ciiu_codigo": (empresa.get("ciiu_codigo") or "").strip(),
        "ciiu_descripcion": (empresa.get("ciiu_descripcion") or "").strip(),
        "actividad": contexto_actividad(empresa),
        "clases_riesgo": clases,
        "codigo": doc_meta.get("codigo", "") or "",
        "version": str(version),
        "fecha": fecha,
        "titulo": doc_meta.get("nombre", "") or "",
        "nombre_documento": doc_meta.get("nombre", "") or "",
        "rep_legal_nombre": (empresa.get("rep_legal_nombre") or "").strip(),
        "rep_legal_cargo": (empresa.get("rep_legal_cargo") or "").strip(),
        "resp_sst_nombre": (empresa.get("resp_sst_nombre") or "").strip(),
        "resp_sst_cargo": (empresa.get("resp_sst_cargo") or "").strip() or "Responsable SST",
        "vigia_nombre": (empresa.get("vigia_nombre") or "").strip() or "________________",
        "elaboro": (empresa.get("resp_sst_nombre") or "").strip(),
        "aprobo": (empresa.get("rep_legal_nombre") or "").strip(),
        "reviso": (empresa.get("vigia_nombre") or "").strip() or "________________",
    }


def _reemplazar_texto(texto: str, ctx: dict[str, str]) -> str:
    if not texto or not isinstance(texto, str):
        return texto
    out = texto
    razon = _ctx_val(ctx, "razon_social")
    nit = _ctx_val(ctx, "nit")
    arl = _ctx_val(ctx, "arl", fallback="")
    resp = _ctx_val(ctx, "resp_sst_nombre", fallback="")
    rep = _ctx_val(ctx, "rep_legal_nombre", fallback="")
    vigia = _ctx_val(ctx, "vigia_nombre")
    codigo = _ctx_val(ctx, "codigo", fallback="")
    version = _ctx_val(ctx, "version", fallback="")
    fecha = _ctx_val(ctx, "fecha", fallback="")
    ciiu = _ctx_val(ctx, "ciiu", fallback="")

    # 1) Placeholders tipo Jinja / corchetes
    for k, v in ctx.items():
        val = (v or "").strip() or "________________"
        out = out.replace("{{" + k + "}}", val)
        out = out.replace("{{ " + k + " }}", val)
        out = out.replace("[" + k.upper() + "]", val)

    # 2) Datos de ejemplo del pack (FOREST GREEN, NIT demo, etc.)
    for rx in _EMPRESA_EJEMPLO_REGEX:
        out = rx.sub(razon, out)
    for rx in _NIT_EJEMPLO_REGEX:
        out = rx.sub(nit, out)

    # CIIU de ejemplo frecuente en el pack
    out = re.sub(
        r"\(CIIU\s*0230\s*y\s*0163\)",
        f"(CIIU {ciiu})" if ciiu else "(CIIU segun actividad de la empresa)",
        out,
        flags=re.I,
    )
    out = re.sub(r"\bCIIU\s*0230\b", f"CIIU {ctx.get('ciiu_codigo') or '____'}", out, flags=re.I)
    out = re.sub(r"\bCIIU\s*0163\b", "", out, flags=re.I)

    # 3) Etiquetas con espacios en blanco
    def _pref(pat: str, valor: str) -> None:
        nonlocal out
        out = re.sub(pat, lambda m: m.group(1) + valor, out)

    _pref(r"(?i)(raz[o\u00f3]n\s+social\s*:?\s*)_{3,}", razon)
    _pref(r"(?i)(nit\s*:?\s*)_{3,}", nit)
    _pref(r"(?i)(empresa\s*:?\s*)_{3,}", razon)
    _pref(r"(?i)(c[o\u00f3]digo\s*:?\s*)_{3,}", codigo)
    _pref(r"(?i)(versi[o\u00f3]n\s*:?\s*)_{3,}", version)
    _pref(r"(?i)(fecha\s*:?\s*)_{3,}", fecha)

    simples = [
        (r"(?i)\[NOMBRE DE LA EMPRESA\]", razon),
        (r"(?i)\[RAZON SOCIAL\]", razon),
        (r"(?i)\[RAZ\u00d3N SOCIAL\]", razon),
        (r"(?i)\[NIT\]", nit),
        (r"(?i)\[CODIGO\]", codigo),
        (r"(?i)\[C\u00d3DIGO\]", codigo),
        (r"(?i)\[VERSION\]", version),
        (r"(?i)\[VERSI\u00d3N\]", version),
        (r"(?i)\[FECHA\]", fecha),
        (r"(?i)\[ARL\]", arl or "________________"),
        (r"(?i)\[RESPONSABLE SST\]", resp or "________________"),
        (r"(?i)\[REPRESENTANTE LEGAL\]", rep or "________________"),
        (r"(?i)\[VIGIA SST\]", vigia),
        (r"(?i)\[NOMBRE DEL VIG[I\u00cd]A SST\]", vigia),
        (r"(?i)NOMBRE DE LA EMPRESA", razon),
        (r"(?i)Nombre de la empresa", razon),
    ]
    for pat, rep_txt in simples:
        out = re.sub(pat, lambda m, r=rep_txt: r, out)

    # 4) NIT demo con etiqueta
    out = re.sub(
        r"(?i)(NIT\s*:?\s*)901[\.\s]?989[\.\s]?693\s*-?\s*7",
        lambda m: m.group(1) + nit,
        out,
    )

    return out


def _logo_path(empresa: dict) -> Path | None:
    """Solo logo adjunto por empresa; sin fallback al logo por defecto."""
    p = (empresa.get("logo_path") or "").strip()
    if not p:
        return None
    path = Path(p)
    if not path.exists():
        return None
    if path.name == "logo.png":
        return None
    return path


def _set_paragraph_text(p, nuevo: str) -> None:
    """Reemplaza el texto del parrafo conservando el estilo del primer run."""
    if p.runs:
        p.runs[0].text = nuevo
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(nuevo)


def _iter_all_paragraphs(doc) -> Iterable:
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        yield from _iter_table_paragraphs(table)
    for section in doc.sections:
        for part in (section.header, section.footer):
            for p in part.paragraphs:
                yield p
            for table in part.tables:
                yield from _iter_table_paragraphs(table)


def _iter_table_paragraphs(table) -> Iterable:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                yield p
            for nested in cell.tables:
                yield from _iter_table_paragraphs(nested)


def _aplicar_documento_word(doc, ctx: dict[str, str]) -> None:
    for p in _iter_all_paragraphs(doc):
        if not p.text:
            continue
        nuevo = _reemplazar_texto(p.text, ctx)
        if nuevo != p.text:
            _set_paragraph_text(p, nuevo)


def rellenar_word(
    plantilla: Path,
    empresa: dict,
    doc_meta: dict,
    version: int,
    fecha: str,
    out_path: Path,
) -> Path:
    ctx = _contexto(empresa, doc_meta, version, fecha)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Intentar docxtpl (si hay {{placeholders}}); si falla, copiar y reemplazar.
    rendered = False
    try:
        from docxtpl import DocxTemplate

        tpl = DocxTemplate(str(plantilla))
        context: dict[str, Any] = {k: _ctx_val(ctx, k) for k in ctx}
        logo = _logo_path(empresa)
        if logo:
            try:
                from docxtpl import InlineImage
                from docx.shared import Mm

                context["logo"] = InlineImage(tpl, str(logo), width=Mm(25))
            except Exception:
                pass
        tpl.render(context)
        tpl.save(str(out_path))
        rendered = True
    except Exception:
        rendered = False

    from docx import Document

    if not rendered:
        doc = Document(str(plantilla))
        doc.save(str(out_path))

    doc = Document(str(out_path))
    _aplicar_documento_word(doc, ctx)
    doc.save(str(out_path))
    return out_path


def rellenar_excel(
    plantilla: Path,
    empresa: dict,
    doc_meta: dict,
    version: int,
    fecha: str,
    out_path: Path,
) -> Path:
    ctx = _contexto(empresa, doc_meta, version, fecha)
    wb = load_workbook(plantilla)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                if isinstance(cell.value, str):
                    nuevo = _reemplazar_texto(cell.value, ctx)
                    if nuevo != cell.value:
                        cell.value = nuevo
        logo = _logo_path(empresa)
        if logo and not list(getattr(ws, "_images", []) or []):
            try:
                img = XLImage(str(logo))
                img.width = 90
                img.height = 55
                ws.add_image(img, "A1")
            except Exception:
                pass
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))
    return out_path
