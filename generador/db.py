# -*- coding: utf-8 -*-
"""Persistencia SQLite: empresas, catalogo, consecutivos e historial por empresa."""
from __future__ import annotations

import json
import re
import sqlite3
from pathlib import Path
from typing import Any, Optional

from generador.catalogo_maestro import CODIGO_SISTEMA, cargar_control_maestro, codigo_sst

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
DB_PATH = DATA_DIR / "empresa.db"
LOGO_PATH = DATA_DIR / "logo.png"
PLANTILLAS_DIR = ROOT / "plantillas"
SALIDA_WORD = ROOT / "salida" / "01_Word"
SALIDA_EXCEL = ROOT / "salida" / "02_Excel"


def _conn() -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _numero_codigo(codigo: str) -> int:
    m = re.search(r"(\d+)$", codigo or "")
    return int(m.group(1)) if m else 0


def _ensure_columns(conn: sqlite3.Connection) -> None:
    cols = {r["name"] for r in conn.execute("PRAGMA table_info(catalogo)").fetchall()}
    if "estado" not in cols:
        conn.execute("ALTER TABLE catalogo ADD COLUMN estado TEXT DEFAULT 'CONFORME'")
    if "ruta" not in cols:
        conn.execute("ALTER TABLE catalogo ADD COLUMN ruta TEXT DEFAULT ''")

    hcols = {r["name"] for r in conn.execute("PRAGMA table_info(historial)").fetchall()}
    if "empresa_id" not in hcols:
        conn.execute("ALTER TABLE historial ADD COLUMN empresa_id INTEGER")


def _migrate_multi_empresa(conn: sqlite3.Connection) -> None:
    """Migra esquema antiguo (1 empresa) a multi-empresa."""
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS empresas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            razon_social TEXT,
            nit TEXT,
            ciiu_codigo TEXT,
            ciiu_descripcion TEXT,
            arl TEXT,
            clases_riesgo TEXT,
            rep_legal_nombre TEXT,
            rep_legal_cargo TEXT,
            resp_sst_nombre TEXT,
            resp_sst_cargo TEXT,
            resp_sst_licencia TEXT,
            vigia_nombre TEXT,
            logo_path TEXT
        );

        CREATE TABLE IF NOT EXISTS app_meta (
            clave TEXT PRIMARY KEY,
            valor TEXT
        );

        CREATE TABLE IF NOT EXISTS consecutivo_empresa (
            empresa_id INTEGER PRIMARY KEY,
            ultimo INTEGER NOT NULL DEFAULT 0
        );
        """
    )

    n_emp = conn.execute("SELECT COUNT(*) AS n FROM empresas").fetchone()["n"]
    if n_emp == 0:
        # Copiar de tabla legacy `empresa` si existe
        tablas = {
            r["name"]
            for r in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        }
        if "empresa" in tablas:
            old = conn.execute("SELECT * FROM empresa WHERE id = 1").fetchone()
            if old:
                d = dict(old)
                cur = conn.execute(
                    """
                    INSERT INTO empresas (
                        razon_social, nit, ciiu_codigo, ciiu_descripcion, arl, clases_riesgo,
                        rep_legal_nombre, rep_legal_cargo, resp_sst_nombre, resp_sst_cargo,
                        resp_sst_licencia, vigia_nombre, logo_path
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        d.get("razon_social", ""),
                        d.get("nit", ""),
                        d.get("ciiu_codigo", ""),
                        d.get("ciiu_descripcion", ""),
                        d.get("arl", ""),
                        d.get("clases_riesgo", "[]"),
                        d.get("rep_legal_nombre", ""),
                        d.get("rep_legal_cargo", ""),
                        d.get("resp_sst_nombre", ""),
                        d.get("resp_sst_cargo", ""),
                        d.get("resp_sst_licencia", ""),
                        d.get("vigia_nombre", ""),
                        d.get("logo_path", ""),
                    ),
                )
                eid = cur.lastrowid
                ultimo = 0
                if "consecutivo" in tablas:
                    row = conn.execute(
                        "SELECT ultimo FROM consecutivo WHERE id = 1"
                    ).fetchone()
                    if row:
                        ultimo = int(row["ultimo"] or 0)
                conn.execute(
                    "INSERT OR REPLACE INTO consecutivo_empresa (empresa_id, ultimo) VALUES (?, ?)",
                    (eid, ultimo),
                )
                conn.execute(
                    "INSERT OR REPLACE INTO app_meta (clave, valor) VALUES ('empresa_activa_id', ?)",
                    (str(eid),),
                )
                conn.execute(
                    "UPDATE historial SET empresa_id = ? WHERE empresa_id IS NULL",
                    (eid,),
                )

    # Asegurar empresa activa
    act = conn.execute(
        "SELECT valor FROM app_meta WHERE clave = 'empresa_activa_id'"
    ).fetchone()
    if act is None:
        first = conn.execute("SELECT id FROM empresas ORDER BY id LIMIT 1").fetchone()
        if first:
            conn.execute(
                "INSERT INTO app_meta (clave, valor) VALUES ('empresa_activa_id', ?)",
                (str(first["id"]),),
            )


def init_db() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    PLANTILLAS_DIR.mkdir(parents=True, exist_ok=True)

    with _conn() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS empresa (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                razon_social TEXT,
                nit TEXT,
                ciiu_codigo TEXT,
                ciiu_descripcion TEXT,
                arl TEXT,
                clases_riesgo TEXT,
                rep_legal_nombre TEXT,
                rep_legal_cargo TEXT,
                resp_sst_nombre TEXT,
                resp_sst_cargo TEXT,
                resp_sst_licencia TEXT,
                vigia_nombre TEXT,
                logo_path TEXT
            );

            CREATE TABLE IF NOT EXISTS catalogo (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT UNIQUE,
                abreviatura TEXT NOT NULL,
                nombre TEXT NOT NULL,
                formato TEXT NOT NULL,
                familia TEXT NOT NULL,
                referencia TEXT,
                version INTEGER DEFAULT 0,
                activo INTEGER DEFAULT 1,
                estado TEXT DEFAULT 'CONFORME',
                ruta TEXT DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS consecutivo (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                ultimo INTEGER NOT NULL DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS historial (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT NOT NULL,
                nombre TEXT NOT NULL,
                version INTEGER NOT NULL,
                fecha TEXT NOT NULL,
                archivo TEXT,
                generado_en TEXT DEFAULT (datetime('now','localtime')),
                empresa_id INTEGER
            );
            """
        )
        _ensure_columns(conn)
        _migrate_multi_empresa(conn)

        row = conn.execute("SELECT ultimo FROM consecutivo WHERE id = 1").fetchone()
        if row is None:
            conn.execute("INSERT INTO consecutivo (id, ultimo) VALUES (1, 0)")

        n_emp = conn.execute("SELECT COUNT(*) AS n FROM empresas").fetchone()["n"]
        if n_emp == 0:
            seed_empresa_suganus(conn)

        cat = conn.execute("SELECT COUNT(*) AS n FROM catalogo").fetchone()["n"]
        sample = conn.execute(
            "SELECT codigo FROM catalogo ORDER BY id LIMIT 1"
        ).fetchone()
        needs_seed = cat == 0
        if sample and not str(sample["codigo"]).startswith(f"{CODIGO_SISTEMA}-"):
            needs_seed = True
        if needs_seed:
            seed_catalogo(conn, reemplazar=True)


def get_empresa_activa_id() -> int:
    init_db()
    with _conn() as conn:
        row = conn.execute(
            "SELECT valor FROM app_meta WHERE clave = 'empresa_activa_id'"
        ).fetchone()
        if not row:
            first = conn.execute("SELECT id FROM empresas ORDER BY id LIMIT 1").fetchone()
            if not first:
                return 0
            return int(first["id"])
        return int(row["valor"])


def get_meta(clave: str, default: str = "") -> str:
    init_db()
    with _conn() as conn:
        row = conn.execute(
            "SELECT valor FROM app_meta WHERE clave = ?", (clave,)
        ).fetchone()
    return str(row["valor"]) if row and row["valor"] is not None else default


def set_meta(clave: str, valor: str) -> None:
    init_db()
    with _conn() as conn:
        conn.execute(
            "INSERT OR REPLACE INTO app_meta (clave, valor) VALUES (?, ?)",
            (clave, str(valor)),
        )


def get_modo_plantillas() -> str:
    """'todas' | 'solo_usuario'"""
    v = get_meta("modo_plantillas", "todas").strip().lower()
    return v if v in {"todas", "solo_usuario"} else "todas"


def set_modo_plantillas(modo: str) -> None:
    m = (modo or "todas").strip().lower()
    if m not in {"todas", "solo_usuario"}:
        m = "todas"
    set_meta("modo_plantillas", m)


def set_empresa_activa(empresa_id: int) -> None:
    init_db()
    with _conn() as conn:
        exists = conn.execute(
            "SELECT id FROM empresas WHERE id = ?", (empresa_id,)
        ).fetchone()
        if not exists:
            raise ValueError(f"Empresa no encontrada: {empresa_id}")
        conn.execute(
            "INSERT OR REPLACE INTO app_meta (clave, valor) VALUES ('empresa_activa_id', ?)",
            (str(empresa_id),),
        )
        conn.execute(
            "INSERT OR IGNORE INTO consecutivo_empresa (empresa_id, ultimo) VALUES (?, 0)",
            (empresa_id,),
        )


def list_empresas() -> list[dict[str, Any]]:
    init_db()
    with _conn() as conn:
        rows = conn.execute(
            "SELECT id, razon_social, nit FROM empresas ORDER BY id"
        ).fetchall()
    return [dict(r) for r in rows]


def seed_empresa_suganus(conn: Optional[sqlite3.Connection] = None) -> None:
    own = conn is None
    if own:
        conn = _conn()
    clases = json.dumps(
        ["I - administrativos", "III - operarios y oficios varios"],
        ensure_ascii=False,
    )
    cur = conn.execute(
        """
        INSERT INTO empresas (
            razon_social, nit, ciiu_codigo, ciiu_descripcion, arl, clases_riesgo,
            rep_legal_nombre, rep_legal_cargo, resp_sst_nombre, resp_sst_cargo,
            resp_sst_licencia, vigia_nombre, logo_path
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            "SUGANUS S.A.S.",
            "900.340.295-9",
            "4610",
            "Comercio al por mayor a cambio de una retribucion o por contrata",
            "Positiva",
            clases,
            "Adriana Galvis Osorio",
            "Gerente Administrativa",
            "Luis Fernando Salazar Velasquez",
            "Responsable SST",
            "",
            "",
            str(LOGO_PATH) if LOGO_PATH.exists() else "",
        ),
    )
    eid = cur.lastrowid
    conn.execute(
        "INSERT OR REPLACE INTO consecutivo_empresa (empresa_id, ultimo) VALUES (?, 0)",
        (eid,),
    )
    conn.execute(
        "INSERT OR REPLACE INTO app_meta (clave, valor) VALUES ('empresa_activa_id', ?)",
        (str(eid),),
    )
    if own:
        conn.commit()
        conn.close()


def crear_empresa_nueva(data: Optional[dict[str, Any]] = None) -> int:
    """Crea empresa vacia (o con datos) y la deja activa. Historial/numeracion en 0."""
    init_db()
    data = data or {}
    clases = data.get("clases_riesgo_list") or []
    if isinstance(clases, str):
        clases = [c.strip() for c in clases.split(",") if c.strip()]
    with _conn() as conn:
        cur = conn.execute(
            """
            INSERT INTO empresas (
                razon_social, nit, ciiu_codigo, ciiu_descripcion, arl, clases_riesgo,
                rep_legal_nombre, rep_legal_cargo, resp_sst_nombre, resp_sst_cargo,
                resp_sst_licencia, vigia_nombre, logo_path
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                data.get("razon_social", ""),
                data.get("nit", ""),
                data.get("ciiu_codigo", ""),
                data.get("ciiu_descripcion", ""),
                data.get("arl", ""),
                json.dumps(clases, ensure_ascii=False),
                data.get("rep_legal_nombre", ""),
                data.get("rep_legal_cargo", ""),
                data.get("resp_sst_nombre", ""),
                data.get("resp_sst_cargo", "Responsable SST"),
                data.get("resp_sst_licencia", ""),
                data.get("vigia_nombre", ""),
                data.get("logo_path", ""),
            ),
        )
        eid = int(cur.lastrowid)
        conn.execute(
            "INSERT INTO consecutivo_empresa (empresa_id, ultimo) VALUES (?, 0)",
            (eid,),
        )
        conn.execute(
            "INSERT OR REPLACE INTO app_meta (clave, valor) VALUES ('empresa_activa_id', ?)",
            (str(eid),),
        )
    return eid


def seed_catalogo(conn: Optional[sqlite3.Connection] = None, reemplazar: bool = False) -> int:
    own = conn is None
    if own:
        conn = _conn()
        _ensure_columns(conn)

    docs = cargar_control_maestro()
    if reemplazar:
        conn.execute("DELETE FROM catalogo")
        # NO borrar historial: pertenece a cada empresa

    for d in docs:
        ref = d.get("referencia", "")
        origen = d.get("codigo_origen", "")
        referencia = f"{origen} | {ref}".strip(" |") if origen else ref
        conn.execute(
            """
            INSERT OR REPLACE INTO catalogo (
                codigo, abreviatura, nombre, formato, familia, referencia,
                version, activo, estado, ruta
            ) VALUES (?, ?, ?, ?, ?, ?, 0, 1, ?, ?)
            """,
            (
                d["codigo"],
                d["abreviatura"],
                d["nombre"],
                d["formato"],
                d["familia"],
                referencia,
                d.get("estado", "CONFORME"),
                ref,
            ),
        )
    if own:
        conn.commit()
        conn.close()
    return len(docs)


def reseed_catalogo_from_maestro() -> int:
    init_db()
    with _conn() as conn:
        return seed_catalogo(conn, reemplazar=True)


def get_empresa() -> dict[str, Any]:
    init_db()
    eid = get_empresa_activa_id()
    with _conn() as conn:
        row = conn.execute("SELECT * FROM empresas WHERE id = ?", (eid,)).fetchone()
    if not row:
        return {}
    d = dict(row)
    try:
        d["clases_riesgo_list"] = json.loads(d.get("clases_riesgo") or "[]")
    except json.JSONDecodeError:
        d["clases_riesgo_list"] = [d.get("clases_riesgo") or ""]
    return d


def save_empresa(data: dict[str, Any]) -> None:
    """Actualiza la empresa activa (no mezcla historial con otras)."""
    init_db()
    eid = get_empresa_activa_id()
    clases = data.get("clases_riesgo_list") or []
    if isinstance(clases, str):
        clases = [c.strip() for c in clases.split(",") if c.strip()]
    with _conn() as conn:
        if eid:
            conn.execute(
                """
                UPDATE empresas SET
                    razon_social=?, nit=?, ciiu_codigo=?, ciiu_descripcion=?, arl=?,
                    clases_riesgo=?, rep_legal_nombre=?, rep_legal_cargo=?,
                    resp_sst_nombre=?, resp_sst_cargo=?, resp_sst_licencia=?,
                    vigia_nombre=?, logo_path=?
                WHERE id=?
                """,
                (
                    data.get("razon_social", ""),
                    data.get("nit", ""),
                    data.get("ciiu_codigo", ""),
                    data.get("ciiu_descripcion", ""),
                    data.get("arl", ""),
                    json.dumps(clases, ensure_ascii=False),
                    data.get("rep_legal_nombre", ""),
                    data.get("rep_legal_cargo", ""),
                    data.get("resp_sst_nombre", ""),
                    data.get("resp_sst_cargo", "Responsable SST"),
                    data.get("resp_sst_licencia", ""),
                    data.get("vigia_nombre", ""),
                    data.get("logo_path", ""),
                    eid,
                ),
            )
        else:
            crear_empresa_nueva(data)


def list_catalogo(solo_activos: bool = True) -> list[dict[str, Any]]:
    init_db()
    with _conn() as conn:
        sql = "SELECT * FROM catalogo"
        if solo_activos:
            sql += " WHERE activo = 1"
        rows = conn.execute(sql).fetchall()
    docs = [dict(r) for r in rows]
    docs.sort(key=lambda d: _numero_codigo(d.get("codigo", "")))
    return docs


def get_documento(codigo: str) -> Optional[dict[str, Any]]:
    init_db()
    with _conn() as conn:
        row = conn.execute("SELECT * FROM catalogo WHERE codigo = ?", (codigo,)).fetchone()
    return dict(row) if row else None


def _ensure_consecutivo_row(conn: sqlite3.Connection, empresa_id: int) -> None:
    conn.execute(
        "INSERT OR IGNORE INTO consecutivo_empresa (empresa_id, ultimo) VALUES (?, 0)",
        (empresa_id,),
    )


def next_consecutivo() -> int:
    init_db()
    eid = get_empresa_activa_id()
    with _conn() as conn:
        _ensure_consecutivo_row(conn, eid)
        ultimo = conn.execute(
            "SELECT ultimo FROM consecutivo_empresa WHERE empresa_id = ?", (eid,)
        ).fetchone()["ultimo"]
        nuevo = int(ultimo) + 1
        conn.execute(
            "UPDATE consecutivo_empresa SET ultimo = ? WHERE empresa_id = ?",
            (nuevo, eid),
        )
    return nuevo


def _siguiente_numero_catalogo(conn: sqlite3.Connection, abreviatura: str, minimo: int = 1) -> int:
    """Siguiente numero libre del catalogo (serie global), evitando IntegrityError."""
    abr = re.sub(r"[^A-Za-z0-9]", "", (abreviatura or "FR").upper())[:12] or "FR"
    max_global = 0
    for row in conn.execute("SELECT codigo FROM catalogo"):
        codigo = row["codigo"] if isinstance(row, sqlite3.Row) else row[0]
        max_global = max(max_global, _numero_codigo(codigo))
    n = max(1, int(minimo), max_global + 1)
    while n <= 9999:
        codigo = codigo_sst(abr, n)
        existe = conn.execute(
            "SELECT 1 FROM catalogo WHERE codigo = ?", (codigo,)
        ).fetchone()
        if not existe:
            return n
        n += 1
    raise ValueError("No hay consecutivos libres para esta abreviatura.")


def add_documento_catalogo(
    abreviatura: str,
    nombre: str,
    formato: str,
    familia: str,
    referencia: str = "",
    version: int = 0,
) -> str:
    """Agrega un documento con codigo unico. Evita IntegrityError si el consecutivo esta desfasado."""
    init_db()
    abr = re.sub(r"[^A-Za-z0-9]", "", (abreviatura or "FR").upper())[:12] or "FR"
    v = max(0, int(version))
    eid = get_empresa_activa_id()
    with _conn() as conn:
        _ensure_consecutivo_row(conn, eid)
        ultimo = int(
            conn.execute(
                "SELECT ultimo FROM consecutivo_empresa WHERE empresa_id = ?",
                (eid,),
            ).fetchone()["ultimo"]
        )
        n = _siguiente_numero_catalogo(conn, abr, minimo=ultimo + 1)
        codigo = codigo_sst(abr, n)
        try:
            conn.execute(
                """
                INSERT INTO catalogo (
                    codigo, abreviatura, nombre, formato, familia, referencia, version, estado, ruta
                ) VALUES (?, ?, ?, ?, ?, ?, ?, 'CONFORME', ?)
                """,
                (codigo, abr, nombre, formato, familia, referencia, v, referencia),
            )
        except sqlite3.IntegrityError:
            # Reintento por carrera / codigo residual
            n = _siguiente_numero_catalogo(conn, abr, minimo=n + 1)
            codigo = codigo_sst(abr, n)
            conn.execute(
                """
                INSERT INTO catalogo (
                    codigo, abreviatura, nombre, formato, familia, referencia, version, estado, ruta
                ) VALUES (?, ?, ?, ?, ?, ?, ?, 'CONFORME', ?)
                """,
                (codigo, abr, nombre, formato, familia, referencia, v, referencia),
            )
        conn.execute(
            "UPDATE consecutivo_empresa SET ultimo = ? WHERE empresa_id = ?",
            (max(ultimo, n), eid),
        )
    return codigo


def bump_version(codigo: str) -> int:
    with _conn() as conn:
        row = conn.execute("SELECT version FROM catalogo WHERE codigo = ?", (codigo,)).fetchone()
        if not row:
            raise ValueError(f"Documento no encontrado: {codigo}")
        nueva = int(row["version"]) + 1
        conn.execute("UPDATE catalogo SET version = ? WHERE codigo = ?", (nueva, codigo))
    return nueva


def set_version(codigo: str, version: int) -> int:
    v = max(1, int(version))
    with _conn() as conn:
        row = conn.execute("SELECT codigo FROM catalogo WHERE codigo = ?", (codigo,)).fetchone()
        if not row:
            raise ValueError(f"Documento no encontrado: {codigo}")
        conn.execute("UPDATE catalogo SET version = ? WHERE codigo = ?", (v, codigo))
    return v


def get_version(codigo: str) -> int:
    with _conn() as conn:
        row = conn.execute("SELECT version FROM catalogo WHERE codigo = ?", (codigo,)).fetchone()
    if not row:
        raise ValueError(f"Documento no encontrado: {codigo}")
    return max(0, int(row["version"] or 0))


def registrar_historial(codigo: str, nombre: str, version: int, archivo: str, fecha: str) -> None:
    eid = get_empresa_activa_id()
    with _conn() as conn:
        conn.execute(
            """
            INSERT INTO historial (codigo, nombre, version, fecha, archivo, empresa_id)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (codigo, nombre, version, fecha, archivo, eid),
        )


def list_historial(limit: int = 200) -> list[dict[str, Any]]:
    init_db()
    eid = get_empresa_activa_id()
    with _conn() as conn:
        rows = conn.execute(
            """
            SELECT * FROM historial
            WHERE empresa_id = ?
            ORDER BY id DESC LIMIT ?
            """,
            (eid, limit),
        ).fetchall()
    return [dict(r) for r in rows]


def get_ultimo_consecutivo() -> int:
    init_db()
    eid = get_empresa_activa_id()
    with _conn() as conn:
        _ensure_consecutivo_row(conn, eid)
        row = conn.execute(
            "SELECT ultimo FROM consecutivo_empresa WHERE empresa_id = ?", (eid,)
        ).fetchone()
    return int(row["ultimo"] if row else 0)


def set_ultimo_consecutivo(ultimo: int) -> int:
    init_db()
    eid = get_empresa_activa_id()
    n = max(0, int(ultimo))
    with _conn() as conn:
        _ensure_consecutivo_row(conn, eid)
        conn.execute(
            "UPDATE consecutivo_empresa SET ultimo = ? WHERE empresa_id = ?",
            (n, eid),
        )
    return n


def max_numero_en_historial() -> int:
    init_db()
    eid = get_empresa_activa_id()
    with _conn() as conn:
        rows = conn.execute(
            "SELECT codigo FROM historial WHERE empresa_id = ?", (eid,)
        ).fetchall()
    m = 0
    for r in rows:
        m = max(m, _numero_codigo(r["codigo"]))
    return m


def get_siguiente_sugerido() -> int:
    return max(get_ultimo_consecutivo(), max_numero_en_historial()) + 1
