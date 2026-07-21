# -*- coding: utf-8 -*-
"""Indice y emparejamiento de plantillas Word/Excel."""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Optional

from generador import db

ROOT = Path(__file__).resolve().parent.parent
WORKSPACE = ROOT.parent
USUARIO_DIR = db.PLANTILLAS_DIR / "usuario"
MAPEO_PATH = db.PLANTILLAS_DIR / "mapeo_usuario.json"

EXT_OK = {".docx", ".xlsx"}

# Cache de listados (el rglob del pack es lento; no reescanear por cada documento)
_PL_CACHE: dict[str, list[Path]] = {}


def invalidar_cache_plantillas() -> None:
    _PL_CACHE.clear()


def _cache_get(key: str):
    return _PL_CACHE.get(key)


def _cache_set(key: str, value: list[Path]) -> list[Path]:
    _PL_CACHE[key] = value
    return value



def _norm(texto: str) -> str:
    t = (texto or "").upper()
    for a, b in (
        ("\u00c1", "A"),
        ("\u00c9", "E"),
        ("\u00cd", "I"),
        ("\u00d3", "O"),
        ("\u00da", "U"),
        ("\u00d1", "N"),
    ):
        t = t.replace(a, b)
    for ch in ("\u0301", "\u0300", "\u0303", "\u0308"):
        t = t.replace(ch, "")
    return t


def carpeta_base_gestion_documental() -> Optional[Path]:
    """Localiza el pack GESTION DOCUMENTAL SST (local Windows/Linux o nube)."""
    candidatos_root = [
        WORKSPACE / "PLANTILLAS",
        ROOT / "PLANTILLAS",
        ROOT / "plantillas",
        db.PLANTILLAS_DIR,
        ROOT / "data" / "PLANTILLAS",
    ]
    vistos: set[Path] = set()
    for plantillas_root in candidatos_root:
        try:
            resolved = plantillas_root.resolve()
        except OSError:
            continue
        if resolved in vistos or not plantillas_root.is_dir():
            continue
        vistos.add(resolved)
        for p in plantillas_root.iterdir():
            if not p.is_dir():
                continue
            if p.name.lower() == "usuario":
                continue
            n = _norm(p.name)
            if "165501" in p.name or ("DOCUMENTAL SST" in n and "GEST" in n):
                return p
    return None


def dirs_plantillas() -> list[Path]:
    """Orden: usuario (prioridad), luego pack base. Evita listar dos veces el mismo pack."""
    dirs: list[Path] = []
    USUARIO_DIR.mkdir(parents=True, exist_ok=True)
    db.PLANTILLAS_DIR.mkdir(parents=True, exist_ok=True)
    dirs.append(USUARIO_DIR)

    if db.get_modo_plantillas() == "solo_usuario":
        return dirs

    base = carpeta_base_gestion_documental()
    if base:
        dirs.append(base)
    return dirs


def listar_plantillas_usuario() -> list[Path]:
    """Solo archivos en plantillas/usuario."""
    cached = _cache_get("usuario")
    if cached is not None:
        return list(cached)
    USUARIO_DIR.mkdir(parents=True, exist_ok=True)
    files: list[Path] = []
    for p in USUARIO_DIR.rglob("*"):
        if not p.is_file():
            continue
        if p.suffix.lower() not in EXT_OK:
            continue
        if p.name.startswith("~$"):
            continue
        files.append(p)
    files = sorted(files, key=lambda x: x.name.lower())
    return _cache_set("usuario", files)


def listar_plantillas() -> list[Path]:
    modo = db.get_modo_plantillas()
    cache_key = f"todas:{modo}"
    cached = _cache_get(cache_key)
    if cached is not None:
        return list(cached)
    vistos: set[str] = set()
    files: list[Path] = []
    for d in dirs_plantillas():
        if not d.is_dir():
            continue
        for p in d.rglob("*"):
            if not p.is_file():
                continue
            if p.suffix.lower() not in EXT_OK:
                continue
            if p.name.startswith("~$"):
                continue
            # Evitar resolve() (lento / falla con rutas largas en Windows)
            try:
                key = str(p).lower()
            except OSError:
                continue
            if key in vistos:
                continue
            vistos.add(key)
            files.append(p)
    files = sorted(files, key=lambda x: x.name.lower())
    return _cache_set(cache_key, files)


def es_plantilla_usuario(path: Path) -> bool:
    try:
        return USUARIO_DIR.resolve() in path.resolve().parents or path.parent.resolve() == USUARIO_DIR.resolve()
    except Exception:
        return False


def numero_fg(texto: str) -> Optional[int]:
    m = re.search(r"FG-FOR-SST-0*(\d+)", texto or "", flags=re.I)
    return int(m.group(1)) if m else None


def codigo_origen_doc(doc: dict[str, Any]) -> Optional[str]:
    for campo in ("referencia", "ruta", "codigo", "nombre"):
        n = numero_fg(str(doc.get(campo) or ""))
        if n is not None:
            return f"FG-FOR-SST-{n}"
    return None


def cargar_mapeo_usuario() -> dict[str, str]:
    if not MAPEO_PATH.exists():
        return {}
    try:
        return json.loads(MAPEO_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def guardar_mapeo_usuario(mapeo: dict[str, str]) -> None:
    MAPEO_PATH.parent.mkdir(parents=True, exist_ok=True)
    MAPEO_PATH.write_text(json.dumps(mapeo, ensure_ascii=False, indent=2), encoding="utf-8")


def _convertir_xls_a_xlsx(data: bytes) -> bytes:
    import io

    import pandas as pd

    buf_in = io.BytesIO(data)
    sheets = pd.read_excel(buf_in, sheet_name=None, engine="xlrd")
    buf_out = io.BytesIO()
    with pd.ExcelWriter(buf_out, engine="openpyxl") as writer:
        for nombre, df in sheets.items():
            sheet = str(nombre)[:31] or "Hoja1"
            df.to_excel(writer, sheet_name=sheet, index=False)
    return buf_out.getvalue()



def _convertir_doc_a_docx(data: bytes, nombre_hint: str = "plantilla.doc") -> bytes:
    """Convierte .doc a .docx con Microsoft Word (Windows) o LibreOffice."""
    import shutil
    import subprocess
    import tempfile

    errores: list[str] = []

    with tempfile.TemporaryDirectory(prefix="sst_doc_") as tmp:
        tmp_path = Path(tmp)
        stem = Path(nombre_hint).stem or "plantilla"
        stem = re.sub(r"[^\w\-]+", "_", stem)[:40] or "plantilla"
        src = tmp_path / f"{stem}.doc"
        dst = tmp_path / f"{stem}.docx"
        src.write_bytes(data)
        src_abs = str(src.resolve())
        dst_abs = str(dst.resolve())

        # 0) Word COM en proceso hijo (evita fallos de hilos de Streamlit)
        try:
            import sys
            import subprocess

            worker = Path(__file__).resolve().parent / "_word_convert_worker.py"
            proc = subprocess.run(
                [sys.executable, str(worker), src_abs, dst_abs],
                capture_output=True,
                text=True,
                timeout=180,
            )
            if proc.returncode == 0 and dst.exists() and dst.stat().st_size > 0:
                return dst.read_bytes()
            err = (proc.stderr or proc.stdout or "").strip()
            errores.append(f"Word subprocess: {err or f'code {proc.returncode}'}")
        except Exception as e:
            errores.append(f"Word subprocess: {type(e).__name__}: {e}")

        # 1) Microsoft Word via COM (mismo proceso)
        try:
            import pythoncom
            import win32com.client  # type: ignore

            pythoncom.CoInitialize()
            word = None
            try:
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                doc = word.Documents.Open(
                    src_abs,
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                )
                try:
                    doc.SaveAs2(dst_abs, FileFormat=16)
                except Exception:
                    doc.SaveAs(dst_abs, FileFormat=16)
                doc.Close(False)
            finally:
                if word is not None:
                    try:
                        word.Quit()
                    except Exception:
                        pass
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            if dst.exists() and dst.stat().st_size > 0:
                return dst.read_bytes()
            errores.append("Word abrió el archivo pero no generó el .docx")
        except Exception as e:
            errores.append(f"Word COM: {type(e).__name__}: {e}")

        # 2) LibreOffice
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if soffice:
            try:
                subprocess.run(
                    [
                        soffice,
                        "--headless",
                        "--convert-to",
                        "docx",
                        "--outdir",
                        str(tmp_path),
                        src_abs,
                    ],
                    check=True,
                    capture_output=True,
                    timeout=120,
                )
                if dst.exists() and dst.stat().st_size > 0:
                    return dst.read_bytes()
                cand = list(tmp_path.glob("*.docx"))
                if cand:
                    return cand[0].read_bytes()
                errores.append("LibreOffice no generó el .docx")
            except Exception as e:
                errores.append(f"LibreOffice: {type(e).__name__}: {e}")
        else:
            errores.append("LibreOffice no está instalado")

    detalle = " | ".join(errores) if errores else "sin detalle"
    raise ValueError(
        "No se pudo convertir el .doc a .docx. "
        "Ábrelo en Word y guárdalo como .docx. "
        f"Detalle: {detalle}"
    )


def registrar_plantilla_usuario(
    uploaded_name: str,
    data: bytes,
    codigo_catalogo: str = "",
) -> Path:
    """Guarda plantilla en plantillas/usuario y opcionalmente la asocia a un codigo.

    - .xls  -> se convierte a .xlsx
    - .doc  -> se convierte a .docx (Word o LibreOffice)
    """
    USUARIO_DIR.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^\w.\- ]+", "_", uploaded_name).strip() or "plantilla.docx"
    # Evitar nombres cortos tipo FG-FOR~1.XLS / FG-FOR~1.DOC
    if "~" in safe:
        stem = Path(safe).stem.replace("~", "_")
        suf = Path(safe).suffix.lower() or ".docx"
        safe = f"{stem}{suf}"

    ext = Path(safe).suffix.lower()
    if ext == ".xls":
        try:
            data = _convertir_xls_a_xlsx(data)
            safe = str(Path(safe).with_suffix(".xlsx"))
        except Exception as e:
            raise ValueError(
                f"No se pudo convertir el .xls a .xlsx. "
                f"Ábrelo en Excel y guárdalo como .xlsx. Detalle: {e}"
            ) from e
    elif ext == ".doc":
        try:
            data = _convertir_doc_a_docx(data, safe)
            safe = str(Path(safe).with_suffix(".docx"))
        except ValueError:
            raise
        except Exception as e:
            raise ValueError(
                f"No se pudo convertir el .doc a .docx. "
                f"Ábrelo en Word y guárdalo como .docx. Detalle: {e}"
            ) from e
    elif ext not in {".docx", ".xlsx"}:
        raise ValueError(
            "Formato no soportado. Usa .docx, .doc, .xlsx o .xls."
        )

    if codigo_catalogo:
        dest = USUARIO_DIR / f"{codigo_catalogo}__{safe}"
    else:
        dest = USUARIO_DIR / safe
    dest.write_bytes(data)
    if codigo_catalogo:
        mapeo = cargar_mapeo_usuario()
        mapeo[codigo_catalogo] = str(dest)
        guardar_mapeo_usuario(mapeo)
    return dest


def _score_nombre(doc: dict[str, Any], path: Path) -> int:
    nombre = _norm(doc.get("nombre") or "")
    fname = _norm(path.stem)
    score = 0
    tokens = [t for t in re.split(r"\W+", nombre) if len(t) > 3]
    for t in tokens:
        if t in fname:
            score += 2
    familia = (doc.get("familia") or "").lower()
    fam_keys = {
        "politica": ["POLITICA"],
        "acta": ["ACTA"],
        "procedimiento": ["PROCEDIMIENTO"],
        "plan": ["PLAN"],
        "programa": ["PROGRAMA"],
        "matriz": ["MATRIZ"],
        "formato": ["FORMATO", "INSPECCION", "ENCUESTA"],
        "evaluacion": ["EVALUACION", "AUTOEVALU"],
    }
    for k in fam_keys.get(familia, []):
        if k in fname:
            score += 3
    return score


def encontrar_plantilla(
    doc: dict[str, Any],
    pool: Optional[list[Path]] = None,
) -> Optional[Path]:
    """Prioridad: mapeo usuario > FG-FOR-SST > palabras clave.

    Si modo_plantillas = solo_usuario, ignora el pack base de la app.
    Pasa `pool` para no reescanear el disco en bucles (catalogo).
    """
    solo_usuario = db.get_modo_plantillas() == "solo_usuario"
    codigo_cat = doc.get("codigo") or ""
    mapeo = cargar_mapeo_usuario()
    if codigo_cat in mapeo:
        p = Path(mapeo[codigo_cat])
        if p.exists() and (not solo_usuario or es_plantilla_usuario(p)):
            return p

    pool = list(pool) if pool is not None else listar_plantillas()
    propias = [p for p in pool if es_plantilla_usuario(p)] if pool else listar_plantillas_usuario()

    # Prefijo en nombre de archivo usuario: SST-SG-XXX-001__archivo.docx
    for p in propias:
        if p.name.startswith(f"{codigo_cat}__"):
            return p

    n_doc = numero_fg(codigo_origen_doc(doc) or "")
    fmt = doc.get("formato") or ""
    candidatos_fg: list[Path] = []
    if n_doc is not None:
        for p in pool:
            if solo_usuario and not es_plantilla_usuario(p):
                continue
            if fmt == "Word" and p.suffix.lower() != ".docx":
                continue
            if fmt == "Excel" and p.suffix.lower() != ".xlsx":
                continue
            if numero_fg(p.name) == n_doc:
                candidatos_fg.append(p)
        if candidatos_fg:
            candidatos_fg.sort(
                key=lambda x: (0 if es_plantilla_usuario(x) else 1, x.name)
            )
            return candidatos_fg[0]

    mejores: list[tuple[int, Path]] = []
    for p in pool:
        if solo_usuario and not es_plantilla_usuario(p):
            continue
        if fmt == "Word" and p.suffix.lower() != ".docx":
            continue
        if fmt == "Excel" and p.suffix.lower() != ".xlsx":
            continue
        score = _score_nombre(doc, p)
        if score > 0:
            mejores.append((score, p))
    if not mejores:
        return None
    mejores.sort(
        key=lambda x: (
            -x[0],
            0 if es_plantilla_usuario(x[1]) else 1,
            x[1].name,
        )
    )
    return mejores[0][1]


def estado_plantillas_catalogo() -> list[dict[str, Any]]:
    solo = db.get_modo_plantillas() == "solo_usuario"
    pool = listar_plantillas()
    rows = []
    for doc in db.list_catalogo():
        plantilla = encontrar_plantilla(doc, pool=pool)
        origen = codigo_origen_doc(doc) or ""
        if plantilla:
            fuente = "adjuntada" if es_plantilla_usuario(plantilla) else "pack base"
            plantilla_txt = f"{plantilla.name} [{fuente}]"
            tiene = "SI"
        else:
            tiene = "NO"
            if solo:
                plantilla_txt = "(sin adjunto; se generara desde cero)"
            else:
                plantilla_txt = "(sin match; se generara desde cero)"
        rows.append(
            {
                "codigo": doc["codigo"],
                "nombre": doc["nombre"][:70],
                "formato": doc["formato"],
                "origen_fg": origen,
                "tiene_plantilla": tiene,
                "plantilla": plantilla_txt,
            }
        )
    return rows
