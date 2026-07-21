# -*- coding: utf-8 -*-
"""Rellena plantillas Word/Excel con datos de empresa sin perder su estructura."""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

from generador.familias_documento import contexto_actividad


def _contexto(empresa: dict, doc_meta: dict, version: int, fecha: str) -> dict[str, str]:
    clases = ", ".join(empresa.get("clases_riesgo_list") or []) or (
        empresa.get("clases_riesgo") or ""
    )
    return {
        "razon_social": empresa.get("razon_social", ""),
        "empresa": empresa.get("razon_social", ""),
        "nit": empresa.get("nit", ""),
        "arl": empresa.get("arl", ""),
        "ciiu": f"{empresa.get('ciiu_codigo', '')} {empresa.get('ciiu_descripcion', '')}".strip(),
        "ciiu_codigo": empresa.get("ciiu_codigo", ""),
        "ciiu_descripcion": empresa.get("ciiu_descripcion", ""),
        "actividad": contexto_actividad(empresa),
        "clases_riesgo": clases,
        "codigo": doc_meta.get("codigo", ""),
        "version": str(version),
        "fecha": fecha,
        "titulo": doc_meta.get("nombre", ""),
        "nombre_documento": doc_meta.get("nombre", ""),
        "rep_legal_nombre": empresa.get("rep_legal_nombre", ""),
        "rep_legal_cargo": empresa.get("rep_legal_cargo", ""),
        "resp_sst_nombre": empresa.get("resp_sst_nombre", ""),
        "resp_sst_cargo": empresa.get("resp_sst_cargo", "") or "Responsable SST",
        "vigia_nombre": (empresa.get("vigia_nombre") or "").strip() or "________________",
        "elaboro": empresa.get("resp_sst_nombre", ""),
        "aprobo": empresa.get("rep_legal_nombre", ""),
        "reviso": (empresa.get("vigia_nombre") or "").strip() or "________________",
    }


def _reemplazar_texto(texto: str, ctx: dict[str, str]) -> str:
    if not texto or not isinstance(texto, str):
        return texto
    out = texto
    for k, v in ctx.items():
        out = out.replace("{{" + k + "}}", v)
        out = out.replace("{{ " + k + " }}", v)
        out = out.replace("[" + k.upper() + "]", v)
    pares = [
        ("(?i)raz[o\u00f3]n social\\s*:?\\s*_{3,}", f"Razon social: {ctx['razon_social']}"),
        ("(?i)nit\\s*:?\\s*_{3,}", f"NIT: {ctx['nit']}"),
        ("(?i)empresa\\s*:?\\s*_{3,}", f"Empresa: {ctx['razon_social']}"),
        ("(?i)c[o\u00f3]digo\\s*:?\\s*_{3,}", f"Codigo: {ctx['codigo']}"),
        ("(?i)versi[o\u00f3]n\\s*:?\\s*_{3,}", f"Version: {ctx['version']}"),
        ("(?i)fecha\\s*:?\\s*_{3,}", f"Fecha: {ctx['fecha']}"),
        ("(?i)\\[NOMBRE DE LA EMPRESA\\]", ctx["razon_social"]),
        ("(?i)\\[RAZON SOCIAL\\]", ctx["razon_social"]),
        ("(?i)\\[NIT\\]", ctx["nit"]),
        ("(?i)\\[CODIGO\\]", ctx["codigo"]),
        ("(?i)\\[VERSION\\]", ctx["version"]),
        ("(?i)\\[FECHA\\]", ctx["fecha"]),
        ("(?i)\\[ARL\\]", ctx["arl"]),
        ("(?i)\\[RESPONSABLE SST\\]", ctx["resp_sst_nombre"]),
        ("(?i)\\[REPRESENTANTE LEGAL\\]", ctx["rep_legal_nombre"]),
        ("(?i)\\[VIGIA SST\\]", ctx["vigia_nombre"]),
        ("(?i)\\[NOMBRE DEL VIG[I\u00cd]A SST\\]", ctx["vigia_nombre"]),
    ]
    for pat, rep in pares:
        out = re.sub(pat, rep, out)
    if "SUGANUS" not in out.upper() and ctx["razon_social"]:
        out = out.replace("NOMBRE DE LA EMPRESA", ctx["razon_social"])
        out = out.replace("Nombre de la empresa", ctx["razon_social"])
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


def _aplicar_parrafos(doc, ctx: dict[str, str]) -> None:
    for p in doc.paragraphs:
        if p.text:
            nuevo = _reemplazar_texto(p.text, ctx)
            if nuevo != p.text:
                if p.runs:
                    p.runs[0].text = nuevo
                    for r in p.runs[1:]:
                        r.text = ""
                else:
                    p.add_run(nuevo)


def rellenar_word(
    plantilla: Path,
    empresa: dict,
    doc_meta: dict,
    version: int,
    fecha: str,
    out_path: Path,
) -> Path:
    ctx = _contexto(empresa, doc_meta, version, fecha)
    try:
        from docxtpl import DocxTemplate

        tpl = DocxTemplate(str(plantilla))
        context: dict[str, Any] = dict(ctx)
        logo = _logo_path(empresa)
        if logo:
            try:
                from docxtpl import InlineImage
                from docx.shared import Mm

                context["logo"] = InlineImage(tpl, str(logo), width=Mm(25))
            except Exception:
                pass
        tpl.render(context)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        tpl.save(str(out_path))
    except Exception:
        from docx import Document

        doc = Document(str(plantilla))
        out_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out_path))

    from docx import Document

    doc = Document(str(out_path))
    _aplicar_parrafos(doc, ctx)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text:
                        nuevo = _reemplazar_texto(p.text, ctx)
                        if nuevo != p.text:
                            if p.runs:
                                p.runs[0].text = nuevo
                                for r in p.runs[1:]:
                                    r.text = ""
                            else:
                                p.add_run(nuevo)
    for section in doc.sections:
        for part in (section.header, section.footer):
            for p in part.paragraphs:
                if p.text:
                    nuevo = _reemplazar_texto(p.text, ctx)
                    if nuevo != p.text:
                        if p.runs:
                            p.runs[0].text = nuevo
                            for r in p.runs[1:]:
                                r.text = ""
                        else:
                            p.add_run(nuevo)
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
                if cell.value is None or not isinstance(cell.value, str):
                    continue
                nuevo = _reemplazar_texto(cell.value, ctx)
                val = cell.value.strip()
                mapa = {
                    "{{razon_social}}": ctx["razon_social"],
                    "{{nit}}": ctx["nit"],
                    "{{codigo}}": ctx["codigo"],
                    "{{version}}": ctx["version"],
                    "{{fecha}}": ctx["fecha"],
                    "{{titulo}}": ctx["titulo"],
                    "{{arl}}": ctx["arl"],
                }
                if val in mapa:
                    nuevo = mapa[val]
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
