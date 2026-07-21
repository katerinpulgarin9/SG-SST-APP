# -*- coding: utf-8 -*-
"""Rellena plantillas Word/Excel con datos de la empresa activa.

Las plantillas del pack traen datos de ejemplo (FOREST GREEN, NIT demo, etc.).
Se reemplazan a nivel de parrafo y tambien en el XML interno del .docx/.xlsx
(para atrapar texto partido en varios runs / cuadros de texto).
"""
from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path
from typing import Any, Iterable
from xml.sax.saxutils import escape as xml_escape

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

from generador.familias_documento import contexto_actividad

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

    for k, v in ctx.items():
        val = (v or "").strip() or "________________"
        out = out.replace("{{" + k + "}}", val)
        out = out.replace("{{ " + k + " }}", val)
        out = out.replace("[" + k.upper() + "]", val)

    for rx in _EMPRESA_EJEMPLO_REGEX:
        out = rx.sub(razon, out)
    for rx in _NIT_EJEMPLO_REGEX:
        out = rx.sub(nit, out)

    out = re.sub(
        r"\(CIIU\s*0230\s*y\s*0163\)",
        f"(CIIU {ciiu})" if ciiu else "(CIIU segun actividad de la empresa)",
        out,
        flags=re.I,
    )
    out = re.sub(r"\bCIIU\s*0230\b", f"CIIU {ctx.get('ciiu_codigo') or '____'}", out, flags=re.I)
    out = re.sub(r"\bCIIU\s*0163\b", "", out, flags=re.I)

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
        (r"(?i)\[NIT\]", nit),
        (r"(?i)\[CODIGO\]", codigo),
        (r"(?i)\[VERSION\]", version),
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

    out = re.sub(
        r"(?i)(NIT\s*:?\s*)901[\.\s]?989[\.\s]?693\s*-?\s*7",
        lambda m: m.group(1) + nit,
        out,
    )
    return out


def _logo_path(empresa: dict) -> Path | None:
    p = (empresa.get("logo_path") or "").strip()
    if not p:
        return None
    path = Path(p)
    if not path.exists() or path.name == "logo.png":
        return None
    return path


def _set_paragraph_text(p, nuevo: str) -> None:
    if p.runs:
        p.runs[0].text = nuevo
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(nuevo)


def _iter_table_paragraphs(table) -> Iterable:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                yield p
            for nested in cell.tables:
                yield from _iter_table_paragraphs(nested)


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


def _aplicar_documento_word(doc, ctx: dict[str, str]) -> None:
    for p in _iter_all_paragraphs(doc):
        if not p.text:
            continue
        nuevo = _reemplazar_texto(p.text, ctx)
        if nuevo != p.text:
            _set_paragraph_text(p, nuevo)


def _flex_phrase_pattern(phrase: str) -> str:
    """Permite etiquetas XML entre palabras (texto partido en runs de Word)."""
    gap = r"(?:\s|&nbsp;|&#160;|<[^>]+>)*"
    parts = [re.escape(p) for p in phrase.split() if p]
    return gap.join(parts)


def _ooxml_scrub_patterns(ctx: dict[str, str]) -> list[tuple[re.Pattern[str], str]]:
    razon = xml_escape(_ctx_val(ctx, "razon_social"))
    nit = xml_escape(_ctx_val(ctx, "nit"))
    phrases = [
        "FOREST GREEN SERVICIOS AMBIENTALES S.A.S.",
        "FOREST GREEN SERVICIOS AMBIENTALES S.A.S",
        "FOREST GREEN SERVICIOS AMBIENTALES",
        "FOREST GREEN",
        "Laboratorios LT S.A.S.",
        "Laboratorios LT S.A.S",
        "XXXXXX XXXXXX",
        "NOMBRE DE LA EMPRESA",
    ]
    pats: list[tuple[re.Pattern[str], str]] = []
    for ph in phrases:
        pats.append((re.compile(_flex_phrase_pattern(ph), re.I), razon))
    # NIT demo con posibles tags entre digitos
    nit_flex = (
        r"901"
        + r"(?:\s|<[^>]+>)*"
        + r"[\.]?"
        + r"(?:\s|<[^>]+>)*"
        + r"989"
        + r"(?:\s|<[^>]+>)*"
        + r"[\.]?"
        + r"(?:\s|<[^>]+>)*"
        + r"693"
        + r"(?:\s|<[^>]+>)*"
        + r"-?"
        + r"(?:\s|<[^>]+>)*"
        + r"7"
    )
    pats.append((re.compile(nit_flex, re.I), nit))
    return pats


def _logo_bytes_for_ext(logo_path: Path, ext: str) -> bytes | None:
    """Convierte el logo de la empresa al formato del media embebido (jpg/png)."""
    try:
        from PIL import Image
    except Exception:
        return logo_path.read_bytes() if logo_path.suffix.lower() == f".{ext}" else None
    try:
        img = Image.open(logo_path)
        if img.mode in ("RGBA", "LA", "P"):
            background = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == "P":
                img = img.convert("RGBA")
            background.paste(img, mask=img.split()[-1] if img.mode == "RGBA" else None)
            img = background
        elif img.mode != "RGB":
            img = img.convert("RGB")
        # Limitar tamano tipico de encabezado
        img.thumbnail((220, 120))
        bio = io.BytesIO()
        ext_l = ext.lower().lstrip(".")
        if ext_l in {"jpg", "jpeg"}:
            img.save(bio, format="JPEG", quality=90)
        elif ext_l == "png":
            img.save(bio, format="PNG")
        else:
            img.save(bio, format="PNG")
        return bio.getvalue()
    except Exception:
        try:
            return logo_path.read_bytes()
        except Exception:
            return None


def _blank_logo_bytes(ext: str) -> bytes:
    """Imagen blanca pequena para quitar logos de ejemplo si no hay logo propio."""
    from PIL import Image

    img = Image.new("RGB", (160, 80), (255, 255, 255))
    bio = io.BytesIO()
    ext_l = ext.lower().lstrip(".")
    if ext_l in {"jpg", "jpeg"}:
        img.save(bio, format="JPEG", quality=85)
    else:
        img.save(bio, format="PNG")
    return bio.getvalue()


def _ooxml_header_meta_patterns(ctx: dict[str, str]) -> list[tuple[re.Pattern[str], str]]:
    """Fecha / Version / Codigo FG-FOR-SST del encabezado de plantilla."""
    fecha = xml_escape(_ctx_val(ctx, "fecha"))
    version = xml_escape(_ctx_val(ctx, "version"))
    codigo = xml_escape(_ctx_val(ctx, "codigo"))
    gap = r"(?:\s|<[^>]+>)*"
    pats: list[tuple[re.Pattern[str], str]] = []
    # Fecha: 01/02/2026  (con posibles tags XML entre tokens)
    pats.append(
        (
            re.compile(
                rf"(Fecha{gap}:{gap})\d{{1,2}}{gap}/{gap}\d{{1,2}}{gap}/{gap}\d{{2,4}}",
                re.I,
            ),
            rf"\g<1>{fecha}",
        )
    )
    pats.append(
        (
            re.compile(
                rf"(Versi(?:\u00f3|o)n{gap}:{gap})\d{{1,4}}",
                re.I,
            ),
            rf"\g<1>{version}",
        )
    )
    pats.append(
        (
            re.compile(
                rf"(C\u00f3digo{gap}:{gap}|Codigo{gap}:{gap})FG{gap}-{gap}FOR{gap}-{gap}SST{gap}-{gap}\d+",
                re.I,
            ),
            rf"\g<1>{codigo}",
        )
    )
    # Codigo suelto FG-FOR-SST-39
    pats.append(
        (
            re.compile(
                rf"FG{gap}-{gap}FOR{gap}-{gap}SST{gap}-{gap}\d+",
                re.I,
            ),
            codigo,
        )
    )
    return pats


def _scrub_ooxml_file(path: Path, ctx: dict[str, str], logo_path: Path | None = None) -> None:
    """Reemplazo en XML + sustitucion de logos embebidos en encabezado/pie."""
    if not path.exists():
        return
    pats = _ooxml_scrub_patterns(ctx) + _ooxml_header_meta_patterns(ctx)

    # Detectar que archivos media se usan en headers/footers (logos de plantilla)
    logo_media_targets: set[str] = set()
    with zipfile.ZipFile(path, "r") as zin:
        names = set(zin.namelist())
        for rel_name in list(names):
            low = rel_name.lower()
            if not (
                low.startswith("word/_rels/header")
                or low.startswith("word/_rels/footer")
                or low == "word/_rels/document.xml.rels"
            ):
                continue
            if not low.endswith(".rels"):
                continue
            try:
                rel_xml = zin.read(rel_name).decode("utf-8")
            except Exception:
                continue
            # Solo headers/footers: sus .rels apuntan a media/imageN
            if "header" in low or "footer" in low:
                for m in re.finditer(r'Target="([^"]*media/[^"]+)"', rel_xml, re.I):
                    target = m.group(1).replace("\\", "/")
                    # Relativo a word/
                    if target.startswith("../"):
                        media_path = "word/" + target[3:]
                    elif target.startswith("media/"):
                        media_path = "word/" + target
                    else:
                        media_path = target
                    # Normalizar
                    for cand in (media_path, media_path.lstrip("/")):
                        if cand in names:
                            logo_media_targets.add(cand)
        # Si no se detecto nada pero hay una sola imagen, es casi seguro el logo
        medias = [n for n in names if "/media/" in n.lower() and n.lower().endswith((".jpg", ".jpeg", ".png"))]
        if not logo_media_targets and len(medias) == 1:
            logo_media_targets.add(medias[0])

    buf = io.BytesIO()
    changed = False
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(
        buf, "w", compression=zipfile.ZIP_DEFLATED
    ) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            name = info.filename
            name_l = name.lower()

            if name in logo_media_targets or name.replace("\\", "/") in logo_media_targets:
                ext = Path(name).suffix.lstrip(".").lower()
                nuevo_bytes = None
                if logo_path and logo_path.exists():
                    nuevo_bytes = _logo_bytes_for_ext(logo_path, ext)
                if nuevo_bytes is None:
                    try:
                        nuevo_bytes = _blank_logo_bytes("jpg" if ext == "jpeg" else ext)
                    except Exception:
                        nuevo_bytes = None
                if nuevo_bytes is not None:
                    data = nuevo_bytes
                    changed = True

            elif name_l.endswith((".xml", ".rels")) and not name_l.endswith(".bin"):
                try:
                    text = data.decode("utf-8")
                except UnicodeDecodeError:
                    zout.writestr(info, data)
                    continue
                nuevo = text
                for rx, rep in pats:
                    try:
                        nuevo2 = rx.sub(rep, nuevo)
                    except re.error:
                        continue
                    if nuevo2 != nuevo:
                        changed = True
                        nuevo = nuevo2
                data = nuevo.encode("utf-8")
            zout.writestr(info, data)
    if changed:
        path.write_bytes(buf.getvalue())


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
    logo = _logo_path(empresa)
    _scrub_ooxml_file(out_path, ctx, logo_path=logo)
    # Reparar layouts rotos de plantillas (encuestas con columnas fingidas)
    try:
        from generador.layout_encuestas import reparar_layouts_word

        reparar_layouts_word(out_path, fecha=ctx.get("fecha"))
        # Reaplicar scrub de logo/meta por si el repair reescribio el docx
        _scrub_ooxml_file(out_path, ctx, logo_path=logo)
    except Exception:
        pass
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
    logo = _logo_path(empresa)
    _scrub_ooxml_file(out_path, ctx, logo_path=logo)
    return out_path
