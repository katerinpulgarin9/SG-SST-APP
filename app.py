# -*- coding: utf-8 -*-
"""App de Gestión Documental SG-SST - Streamlit."""
from __future__ import annotations

import sys
import io
import zipfile
from datetime import date, datetime
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import streamlit as st
import pandas as pd
from PIL import Image
from generador import db
from generador.motor import estado_plantillas_catalogo, generar_documento, listar_plantillas
from generador.normativa import NORMATIVA
from generador.plantillas_index import carpeta_base_gestion_documental

st.set_page_config(
    page_title="SG-SST - Gestión documental",
    page_icon=":clipboard:",
    layout="wide",
)


def _ensure():
    db.init_db()


_ensure()





def _aplicar_estilo() -> None:
    """Tema visual claro tipo mockup SG-SST."""
    st.markdown(
        """
<style>
:root {
  --sst-azul: #4F74B8;
  --sst-azul-claro: #EAF0F9;
  --sst-azul-suave: #F5F8FC;
  --sst-amarillo-claro: #FFF8E7;
  --sst-rojo-claro: #FCEEEE;
  --sst-verde: #6FAE7B;
  --sst-verde-claro: #EAF6ED;
  --sst-texto: #2C3E50;
  --sst-muted: #6B7C8F;
  --sst-borde: #D5E0EF;
}

.stApp { background: #F7FAFD !important; }

section[data-testid="stSidebar"] {
  background: #F3F7FC !important;
  border-right: 1px solid var(--sst-borde);
}
section[data-testid="stSidebar"] > div:first-child { padding-top: 1rem; }

h1, h2, h3 { color: var(--sst-azul) !important; font-weight: 650 !important; }
.sst-page-desc { color: var(--sst-muted); font-size: 0.95rem; margin-bottom: 1rem; }
.sst-brand { text-align: center; margin-bottom: 0.6rem; }
.sst-brand-title {
  color: var(--sst-azul) !important;
  font-size: 2.75rem !important;
  line-height: 1.05 !important;
  font-weight: 800 !important;
  margin: 0.1rem 0 0.35rem 0 !important;
  letter-spacing: 0.03em !important;
}
section[data-testid="stSidebar"] .sst-brand-title,
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p.sst-brand-title,
section[data-testid="stSidebar"] h1.sst-brand-title {
  font-size: 2.75rem !important;
  font-weight: 800 !important;
  color: #4F74B8 !important;
}
.sst-brand-sub { color: var(--sst-muted); font-size: 0.78rem; line-height: 1.3; margin: 0.2rem 0 0.8rem 0; }

/* Botones de navegacion (tarjetas) */
section[data-testid="stSidebar"] div.stButton > button {
  width: 100%;
  text-align: left !important;
  justify-content: flex-start !important;
  background: #fff !important;
  border: 1px solid var(--sst-borde) !important;
  border-radius: 14px !important;
  padding: 0.85rem 1rem !important;
  margin-bottom: 0.45rem !important;
  color: var(--sst-texto) !important;
  font-weight: 600 !important;
  box-shadow: none !important;
  white-space: pre-line !important;
  min-height: 64px;
}
section[data-testid="stSidebar"] div.stButton > button:hover {
  border-color: var(--sst-azul) !important;
  background: #fff !important;
}
section[data-testid="stSidebar"] div.stButton > button[kind="primary"] {
  background: #fff !important;
  color: var(--sst-azul) !important;
  border: 2px solid var(--sst-azul) !important;
  box-shadow: 0 0 0 3px rgba(79,116,184,0.12) !important;
}
section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] p {
  margin-bottom: 0.2rem;
}

div.stButton > button[kind="primary"],
button[data-testid="baseButton-primary"] {
  background: var(--sst-verde) !important;
  border: none !important;
  color: #fff !important;
  border-radius: 10px !important;
  font-weight: 600 !important;
}
div.stButton > button {
  border-radius: 10px !important;
}

div[data-baseweb="input"] > div,
div[data-baseweb="select"] > div,
div[data-baseweb="textarea"] > div {
  border-radius: 10px !important;
  border-color: var(--sst-borde) !important;
  background: #fff !important;
}

div[data-testid="stAlert"] { border-radius: 12px !important; }
div[data-testid="stExpander"] {
  background: var(--sst-amarillo-claro);
  border: 1px solid #F0E4B8;
  border-radius: 12px;
}
div[data-testid="stDataFrame"],
div[data-testid="stDataEditor"] {
  border: 1px solid var(--sst-borde);
  border-radius: 12px;
  overflow: visible;
  background: #fff;
  min-height: 280px;
}
hr { border-color: var(--sst-borde) !important; }

div.stDownloadButton > button {
  background: var(--sst-verde-claro) !important;
  color: #2F5D3A !important;
  border: 1px solid #C5E0CC !important;
  border-radius: 10px !important;
  font-weight: 600 !important;
}

section[data-testid="stFileUploader"] {
  background: transparent !important;
  border: none !important;
  padding: 0 !important;
}

.sst-card {
  background: #fff;
  border: 1px solid var(--sst-borde);
  border-radius: 16px;
  padding: 1rem 1.1rem;
  margin-bottom: 1rem;
}
.sst-logo-box {
  background: var(--sst-amarillo-claro);
  border: 1px solid #F0E4B8;
  border-radius: 16px;
  padding: 1rem 1.1rem;
  margin-top: 0.5rem;
}
.sst-logo-empty {
  border: 1.5px dashed #D2C48A;
  border-radius: 12px;
  background: #FFFcf3;
  min-height: 140px;
  display: flex;
  align-items: center;
  justify-content: center;
  text-align: center;
  color: var(--sst-muted);
  padding: 1rem;
  font-size: 0.9rem;
}
.sst-tip {
  background: #FFF3C4;
  border-radius: 10px;
  padding: 0.7rem 0.8rem;
  color: #6A5A20;
  font-size: 0.85rem;
}
.sst-reminder {
  background: var(--sst-rojo-claro);
  border: 1px solid #F0C9C9;
  border-radius: 12px;
  padding: 0.7rem 0.85rem;
  color: #7A3B3B;
  font-size: 0.85rem;
  margin-top: 0.8rem;
}
.sst-footer {
  color: var(--sst-muted);
  font-size: 0.75rem;
  text-align: center;
  margin-top: 1rem;
}
</style>
""",
        unsafe_allow_html=True,
    )



def _es_empresa_prueba(emp: dict) -> bool:
    """Oculta empresas claramente de prueba en el selector."""
    razon = (emp.get("razon_social") or "").strip().upper()
    nit = re.sub(r"[^0-9]", "", (emp.get("nit") or "").strip())
    if not razon and not nit:
        return False
    if any(x in razon for x in ("PRUEBA", "TEST", "OTRA SAS", "DEMO")):
        return True
    if razon in {"OTRA", "OTRA S.A.S", "OTRA S.A.S."}:
        return True
    if nit in {"123", "1", "000", "000000000", "123456789"}:
        return True
    return False


def _empresas_visibles() -> list[dict]:
    return [e for e in db.list_empresas() if not _es_empresa_prueba(e)]



def _barra_lateral() -> str:
    """Menu lateral tipo tarjetas (mockup)."""
    emp = db.get_empresa()
    if _es_empresa_prueba(emp):
        visibles = _empresas_visibles()
        if visibles:
            db.set_empresa_activa(int(visibles[0]["id"]))
            emp = db.get_empresa()
            st.session_state["num_desde_pending"] = db.get_siguiente_sugerido()
            st.session_state["descargas_listas"] = []

    nav_items = [
        ("Empresa", "Configuraci\u00f3n de la empresa"),
        ("Plantillas", "Gestiona tus plantillas"),
        ("Cat\u00e1logo / Generar", "Documentos del SG-SST"),
        ("Historial", "Documentos generados"),
    ]
    if "nav_pagina" not in st.session_state:
        st.session_state["nav_pagina"] = "Empresa"

    with st.sidebar:
        st.markdown(
            '<div class="sst-brand">'
            '<h1 class="sst-brand-title" style="font-size:2.75rem!important;font-weight:800!important;color:#4F74B8!important;line-height:1.05!important;margin:0.1rem 0 0.35rem 0!important;letter-spacing:0.03em!important;">SG-SST</h1>'
            '<p class="sst-brand-sub">Sistema de Gesti\u00f3n de Seguridad y Salud en el Trabajo</p>'
            "</div>",
            unsafe_allow_html=True,
        )

        for nombre, sub in nav_items:
            activo = st.session_state["nav_pagina"] == nombre
            etiqueta = f"{'●  ' if activo else ''}{nombre}\n{sub}"
            if st.button(
                etiqueta,
                key=f"navbtn_{nombre}",
                use_container_width=True,
                type="primary" if activo else "secondary",
            ):
                st.session_state["nav_pagina"] = nombre
                st.rerun()

        st.markdown(
            '<div class="sst-reminder">Recuerda mantener la informaci\u00f3n de tu empresa actualizada.</div>',
            unsafe_allow_html=True,
        )

        with st.expander("Gu\u00eda r\u00e1pida", expanded=False):
            st.markdown(
                """1. **Empresa**: logo, NIT y responsables.
2. **Plantillas**: deja *Pack base + adjuntadas*.
3. **Cat\u00e1logo / Generar**: marca **Sel**, genera y **descarga**.
4. **Historial**: revisa lo generado por esta empresa."""
            )

        nombre = emp.get("razon_social") or "Sin nombre"
        nit = emp.get("nit") or "-"
        st.caption(f"Empresa en uso: **{nombre}** · NIT {nit}")
        st.markdown('<p class="sst-footer">SG-SST Versi\u00f3n 1.0.0</p>', unsafe_allow_html=True)

    return st.session_state["nav_pagina"]




def pagina_empresa():
    emp = db.get_empresa()
    eid = db.get_empresa_activa_id()

    head_l, head_r = st.columns([3, 2])
    with head_l:
        st.header("Configuraci\u00f3n de la empresa")
        st.markdown(
            '<p class="sst-page-desc">Completa la informaci\u00f3n de tu empresa. '
            "Estos datos se usar\u00e1n para generar los documentos del SG-SST.</p>",
            unsafe_allow_html=True,
        )
    with head_r:
        st.write("")
        guardar = st.button(
            "Guardar cambios",
            type="primary",
            use_container_width=True,
            key="btn_guardar_empresa_top",
        )

    empresas = _empresas_visibles()
    activa_id = db.get_empresa_activa_id()
    labels = {}
    for e in empresas:
        rs = (e.get("razon_social") or "").strip()
        nit_e = (e.get("nit") or "").strip() or "-"
        titulo = rs if rs else "Empresa nueva (completa los datos)"
        labels[f"{titulo} | NIT {nit_e}"] = e["id"]

    c_sel, c_new = st.columns([3, 1])
    with c_sel:
        if labels:
            actual_label = next(
                (k for k, v in labels.items() if v == activa_id),
                list(labels.keys())[0],
            )
            elegido = st.selectbox(
                "Empresa activa",
                list(labels.keys()),
                index=list(labels.keys()).index(actual_label),
            )
            if labels[elegido] != activa_id:
                db.set_empresa_activa(labels[elegido])
                st.session_state["num_desde_pending"] = db.get_siguiente_sugerido()
                st.session_state["descargas_listas"] = []
                st.rerun()
    with c_new:
        st.write("")
        if st.button("Nueva empresa", use_container_width=True):
            db.crear_empresa_nueva()
            st.session_state["num_desde_pending"] = 1
            st.session_state["descargas_listas"] = []
            st.success("Empresa nueva creada.")
            st.rerun()

    emp = db.get_empresa()
    eid = db.get_empresa_activa_id()

    st.markdown("#### Informaci\u00f3n general")
    st.markdown('<div class="sst-card">', unsafe_allow_html=True)
    r1a, r1b = st.columns(2)
    with r1a:
        razon = st.text_input("Raz\u00f3n social *", emp.get("razon_social", ""), placeholder="Ej. Empresa Ejemplo S.A.S.")
    with r1b:
        nit = st.text_input("NIT *", emp.get("nit", ""), placeholder="Ej. 900.000.000-1")
    r2a, r2b = st.columns(2)
    with r2a:
        ciiu_cod = st.text_input("CIIU (c\u00f3digo)", emp.get("ciiu_codigo", ""), placeholder="Ej. 4610")
    with r2b:
        ciiu_desc = st.text_input("Actividad econ\u00f3mica (CIIU)", emp.get("ciiu_descripcion", ""), placeholder="Descripci\u00f3n CIIU")
    r3a, r3b = st.columns(2)
    with r3a:
        arl = st.text_input("ARL", emp.get("arl", ""), placeholder="Ej. Positiva")
    with r3b:
        clases_default = ", ".join(emp.get("clases_riesgo_list") or [])
        clases = st.text_input(
            "Clase(s) de riesgo",
            clases_default,
            placeholder="Ej: I - administrativos, III - operarios",
            help="Separadas por coma",
        )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("#### Responsables")
    st.markdown('<div class="sst-card">', unsafe_allow_html=True)
    a, b = st.columns(2)
    with a:
        rep_n = st.text_input("Representante legal - nombre", emp.get("rep_legal_nombre", ""))
        rep_c = st.text_input("Representante legal - cargo", emp.get("rep_legal_cargo", ""))
    with b:
        sst_n = st.text_input("Responsable SST - nombre", emp.get("resp_sst_nombre", ""))
        sst_c = st.text_input("Responsable SST - cargo", emp.get("resp_sst_cargo", "Responsable SST"))
        sst_l = st.text_input("Licencia SST (opcional)", emp.get("resp_sst_licencia", ""))
    vigia = st.text_input(
        "Vig\u00eda / COPASST - nombre (opcional)",
        emp.get("vigia_nombre", ""),
        help="Si queda vac\u00edo, en firmas se deja l\u00ednea en blanco.",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("#### Logo de la empresa")
    st.markdown('<div class="sst-logo-box">', unsafe_allow_html=True)
    logo_path = (emp.get("logo_path") or "").strip()
    logo_file = Path(logo_path) if logo_path else None
    tiene_logo = bool(
        logo_file and logo_file.exists() and logo_file.name.startswith("logo_emp_")
    )
    if "logo_up_nonce" not in st.session_state:
        st.session_state["logo_up_nonce"] = 0

    c_prev, c_acc, c_tip = st.columns([1.2, 1.2, 1.3])
    with c_prev:
        if tiene_logo:
            st.image(str(logo_file), width=160)
        else:
            st.markdown(
                '<div class="sst-logo-empty">Sin logo.<br/>Sube el logo de tu empresa<br/>en formato PNG o JPG.</div>',
                unsafe_allow_html=True,
            )
    with c_acc:
        st.caption("Recomendado: PNG/JPG · m\u00e1x. 2 MB")
        up = st.file_uploader(
            "Subir logo",
            type=["png", "jpg", "jpeg"],
            key=f"logo_uploader_{eid}_{st.session_state['logo_up_nonce']}",
            label_visibility="collapsed",
        )
        if tiene_logo:
            if st.button("Quitar logo", key=f"btn_quitar_logo_{eid}", use_container_width=True):
                try:
                    if logo_file.exists():
                        logo_file.unlink()
                except OSError:
                    pass
                emp_data = dict(emp)
                emp_data["logo_path"] = ""
                emp_data["clases_riesgo_list"] = emp.get("clases_riesgo_list") or []
                db.save_empresa(emp_data)
                st.session_state["logo_up_nonce"] = int(st.session_state["logo_up_nonce"]) + 1
                st.success("Logo eliminado.")
                st.rerun()
        if up is not None:
            img = Image.open(up)
            db.LOGO_PATH.parent.mkdir(parents=True, exist_ok=True)
            dest = db.DATA_DIR / f"logo_emp_{eid}.png"
            img.save(dest)
            emp_data = dict(emp)
            emp_data["logo_path"] = str(dest)
            emp_data["clases_riesgo_list"] = emp.get("clases_riesgo_list") or []
            db.save_empresa(emp_data)
            st.session_state["logo_up_nonce"] = int(st.session_state["logo_up_nonce"]) + 1
            st.success("Logo guardado.")
            st.rerun()
    with c_tip:
        st.markdown(
            '<div class="sst-tip">Este logo se usar\u00e1 en los documentos generados. '
            "Puedes actualizarlo o quitarlo cuando lo necesites.</div>",
            unsafe_allow_html=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)

    guardar_abajo = st.button("Guardar configuraci\u00f3n", type="primary", key="btn_guardar_empresa_bottom")
    if guardar or guardar_abajo:
        db.save_empresa(
            {
                "razon_social": razon,
                "nit": nit,
                "ciiu_codigo": ciiu_cod,
                "ciiu_descripcion": ciiu_desc,
                "arl": arl,
                "clases_riesgo_list": [x.strip() for x in clases.split(",") if x.strip()],
                "rep_legal_nombre": rep_n,
                "rep_legal_cargo": rep_c,
                "resp_sst_nombre": sst_n,
                "resp_sst_cargo": sst_c,
                "resp_sst_licencia": sst_l,
                "vigia_nombre": vigia,
                "logo_path": (emp.get("logo_path") or "").strip(),
            }
        )
        st.success("Datos de la empresa guardados.")



def pagina_plantillas():
    st.header("2. Gestión de plantillas")
    from generador.plantillas_index import (
        USUARIO_DIR,
        cargar_mapeo_usuario,
        listar_plantillas_usuario,
        registrar_plantilla_usuario,
    )

    base = carpeta_base_gestion_documental()
    st.caption(
        "Puedes usar el pack base de la app, solo tus plantillas adjuntadas, o ambas."
    )

    modo_actual = db.get_modo_plantillas()
    modo_label = st.radio(
        "Fuente de plantillas al generar",
        [
            "Pack base + plantillas adjuntadas",
            "Solo plantillas adjuntadas (ignorar pack de la app)",
        ],
        index=0 if modo_actual == "todas" else 1,
        help=(
            "Si eliges solo adjuntadas, la app no usará las plantillas del pack "
            "GESTIÓN DOCUMENTAL SST. Si no hay coincidencia, genera desde cero."
        ),
    )
    nuevo_modo = "todas" if modo_label.startswith("Pack base") else "solo_usuario"
    if nuevo_modo != modo_actual:
        db.set_modo_plantillas(nuevo_modo)
        st.session_state.pop("_estado_pl_map", None)
        st.rerun()

    if nuevo_modo == "solo_usuario":
        st.warning(
            "Modo activo: **solo plantillas adjuntadas**. "
            "El pack de la app no se usará al generar."
        )
    elif base:
        st.success(f"Pack base disponible: `{base.name}`")
    else:
        st.warning("No se encontró la carpeta GESTIÓN DOCUMENTAL SST en PLANTILLAS/.")

    archivos = listar_plantillas()
    propias = listar_plantillas_usuario()
    if nuevo_modo == "solo_usuario":
        st.write(f"**{len(propias)} plantilla(s) adjuntada(s)** en uso.")
    else:
        st.write(
            f"**{len(archivos)} plantilla(s)** disponibles "
            f"(adjuntadas: {len(propias)} | total indexado: {len(archivos)})."
        )

    st.subheader("Anexar plantilla propia")
    st.caption(
        "Sube .docx, .doc, .xlsx o .xls (los formatos antiguos se convierten). "
        "Si lo asocias a un documento del catálogo, se usará para ese documento al generar."
    )
    docs = db.list_catalogo()
    opciones = {"(Sin asociar / coincidencia automática)": ""}
    opciones.update({f"{d['codigo']} - {d['nombre'][:60]}": d["codigo"] for d in docs})
    up = st.file_uploader(
        "Subir plantilla",
        type=["docx", "doc", "xlsx", "xls"],
        accept_multiple_files=True,
    )
    asoc = st.selectbox("Asociar a documento del catálogo (opcional)", list(opciones.keys()))
    if st.button("Guardar plantilla(s) anexada(s)", type="primary") and up:
        codigo_asoc = opciones[asoc]
        guardadas = []
        errores = []
        for f in up:
            try:
                path = registrar_plantilla_usuario(f.name, f.getvalue(), codigo_asoc)
                guardadas.append(path.name)
            except Exception as e:
                errores.append(f"{f.name}: {e}")
        if guardadas:
            st.success(f"Guardadas en `{USUARIO_DIR}`: {', '.join(guardadas)}")
        for err in errores:
            st.error(err)
        if guardadas:
            st.session_state.pop("_estado_pl_map", None)
            st.rerun()

    mapeo = cargar_mapeo_usuario()
    if mapeo:
        st.subheader("Asociaciones manuales")
        st.dataframe(
            [{"Código": k, "Plantilla": Path(v).name} for k, v in mapeo.items()],
            use_container_width=True,
            hide_index=True,
        )

    with st.expander("Ver plantillas en uso (según el modo elegido)"):
        lista_ver = propias if nuevo_modo == "solo_usuario" else archivos
        if not lista_ver:
            st.info("No hay plantillas en este modo.")
        for p in lista_ver[:80]:
            st.text(f"- {p.name}")
        if len(lista_ver) > 80:
            st.caption(f"... y {len(lista_ver) - 80} ms")

    st.subheader("Coincidencia con el catálogo")
    if nuevo_modo == "solo_usuario":
        st.info(
            "Estás en **solo plantillas adjuntadas**: el pack GESTIÓN DOCUMENTAL SST "
            "de la app **no se usa**. Por eso verás muchos **NO** aunque el pack exista. "
            "Para usar esas plantillas del pack, cambia arriba a "
            "**Pack base + plantillas adjuntadas**."
        )
    else:
        st.info(
            "Modo **pack base + adjuntadas**: al generar, si hay plantilla (SI) se usa esa; "
            "si dice NO, el documento se arma desde cero con la estructura de la app."
        )
    rows = estado_plantillas_catalogo()
    con = sum(1 for r in rows if r.get("tiene_plantilla") == "SI")
    st.caption(f"Con plantilla (modo actual): **{con}** / {len(rows)}")
    st.dataframe(rows, use_container_width=True, hide_index=True, height=360)

    st.subheader("Normativa de referencia (solo si no hay plantilla)")
    for n in NORMATIVA:
        st.markdown(f"- **{n['nombre']}**: {n['descripcion']}")


def pagina_catalogo():
    st.header("Catálogo y generación")
    if st.session_state.get("descargas_listas"):
        if st.session_state.get("flash_gen"):
            st.success(st.session_state.pop("flash_gen"))
        st.info("Hay documentos listos para descargar. Usa los botones de la sección Descargas más abajo o genera más.")
    emp = db.get_empresa()
    st.caption(
        f"{emp.get('razon_social','')} | CIIU {emp.get('ciiu_codigo','')} - {emp.get('ciiu_descripcion','')}"
    )

    c_reload, c_info = st.columns([1, 3])
    with c_reload:
        if st.button("Recargar Control Maestro"):
            n = db.reseed_catalogo_from_maestro()
            st.success(
                f"Catálogo recargado: {n} documentos "
                "(sin duplicados ni formatos de herramientas)"
            )
            st.rerun()
    with c_info:
        sugerido = db.get_siguiente_sugerido()
        st.caption(
            "Fuente: Control Maestro (referencia). Al generar se asigna "
            "`SST-SG-[ABR]-[NNN]` en el orden de generación. "
            f"| Catálogo: {len(db.list_catalogo())} | Siguiente sugerido: **{sugerido:03d}**"
        )

    c_fecha, c_ver, c_num = st.columns(3)
    with c_fecha:
        fecha = st.date_input("Fecha de los documentos", value=date.today())
    with c_ver:
        version_manual = st.number_input(
            "Versión",
            min_value=1,
            max_value=999,
            value=1,
            step=1,
            help="Version que se imprimirá en el encabezado del documento.",
        )
    with c_num:
        # Aplicar cambios de numeracion ANTES de crear el widget (regla Streamlit)
        if "num_desde_pending" in st.session_state:
            st.session_state["num_desde"] = int(st.session_state.pop("num_desde_pending"))
        if "num_desde" not in st.session_state:
            st.session_state["num_desde"] = db.get_siguiente_sugerido()
        num_desde = st.number_input(
            "Continuar numeración desde",
            min_value=1,
            max_value=9999,
            step=1,
            key="num_desde",
            help=(
                "NNN del primer documento de este lote. "
                "Si ya generaste hasta el 025, pon 26 para seguir. "
                "Los demás suben 027, 028... en el orden de generación."
            ),
        )
    modo_version = st.radio(
        "Al generar (versión)",
        ["Usar versión indicada", "Usar versión del documento", "Incrementar automáticamente"],
        horizontal=True,
        help=(
            "Indicada: la del campo Versión. "
            "Del documento: la guardada en el catálogo. "
            "Incrementar: sube 1 respecto a la guardada."
        ),
    )
    c_sync, c_reset = st.columns(2)
    with c_sync:
        if st.button("Usar siguiente sugerido"):
            st.session_state["num_desde_pending"] = db.get_siguiente_sugerido()
            st.rerun()
    with c_reset:
        if st.button("Reiniciar numeración en 001"):
            st.session_state["num_desde_pending"] = 1
            st.rerun()
    fecha_str = fecha.strftime("%d/%m/%Y")

    docs = db.list_catalogo()
    if not docs:
        st.warning("Catálogo vacío.")
        return


    c_bus, c_pl = st.columns([3, 1])
    with c_bus:
        filtro = st.text_input("Buscar en catálogo (código o nombre)", "")
    with c_pl:
        filtro_pl = st.selectbox(
            "Plantilla",
            ["Todas", "Solo SI", "Solo NO"],
            help="Filtra por si hay plantilla coincidente",
        )
    docs_filtrados = docs
    if filtro.strip():
        f = filtro.strip().lower()
        docs_filtrados = [
            d for d in docs_filtrados
            if f in d["codigo"].lower() or f in d["nombre"].lower()
        ]
    if filtro_pl != "Todas":
        if "_estado_pl_map" not in st.session_state:
            try:
                st.session_state["_estado_pl_map"] = {
                    r["codigo"]: r.get("tiene_plantilla", "NO")
                    for r in estado_plantillas_catalogo()
                }
            except Exception:
                st.session_state["_estado_pl_map"] = {}
        want = "SI" if filtro_pl == "Solo SI" else "NO"
        docs_filtrados = [
            d for d in docs_filtrados
            if st.session_state["_estado_pl_map"].get(d["codigo"], "NO") == want
        ]

    st.write(f"Mostrando **{len(docs_filtrados)}** de {len(docs)} documentos")
    st.caption("Marca el cuadrito de la columna **Sel** para elegir uno o varios documentos.")

    seleccion_prev = set(st.session_state.get("catalogo_sel_codigos", []))

    c_sel1, c_sel2, _ = st.columns([1, 1, 2])
    with c_sel1:
        marcar_filtrados = st.button("Marcar filtrados", disabled=not docs_filtrados)
    with c_sel2:
        limpiar_sel = st.button("Limpiar selección")

    def _reset_editor():
        # Forzar que la tabla refresque los cuadritos
        for k in list(st.session_state.keys()):
            if str(k).startswith("catalogo_editor"):
                del st.session_state[k]

    if marcar_filtrados:
        seleccion_prev |= {d["codigo"] for d in docs_filtrados}
        st.session_state["catalogo_sel_codigos"] = sorted(seleccion_prev)
        _reset_editor()
        st.rerun()
    if limpiar_sel:
        st.session_state["catalogo_sel_codigos"] = []
        _reset_editor()
        st.rerun()

    st.markdown("#### Lista del catálogo")
    st.caption("Desplázate en la tabla. Marca **Sel** para elegir documentos.")

    # Mapa codigo -> tiene plantilla (SI/NO), cache por sesion
    if "_estado_pl_map" not in st.session_state:
        try:
            st.session_state["_estado_pl_map"] = {
                r["codigo"]: r.get("tiene_plantilla", "NO")
                for r in estado_plantillas_catalogo()
            }
        except Exception:
            st.session_state["_estado_pl_map"] = {}
    _estado_pl = st.session_state["_estado_pl_map"]

    df = pd.DataFrame(
        [
            {
                "Sel": d["codigo"] in seleccion_prev,
                "codigo": d["codigo"],
                "nombre": d["nombre"],
                "formato": d["formato"],
                "familia": d["familia"],
                "plantilla": _estado_pl.get(d["codigo"], "?"),
                "version": int(d["version"] or 0),
                "estado": d.get("estado") or "",
            }
            for d in docs_filtrados
        ]
    )

    filtro_key = filtro.strip().lower() or "_all"
    editado = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        height=420,
        disabled=["codigo", "nombre", "formato", "familia", "plantilla", "version", "estado"],
        column_config={
            "Sel": st.column_config.CheckboxColumn(
                "Sel",
                help="Clic para seleccionar este documento",
                default=False,
                width="small",
            ),
            "codigo": st.column_config.TextColumn("Código", width="medium"),
            "nombre": st.column_config.TextColumn("Nombre", width="large"),
            "formato": st.column_config.TextColumn("Formato", width="small"),
            "familia": st.column_config.TextColumn("Familia", width="small"),
            "plantilla": st.column_config.TextColumn("Plantilla", width="small", help="SI = usa plantilla; NO = desde cero"),
            "version": st.column_config.NumberColumn("Versión", width="small"),
            "estado": st.column_config.TextColumn("Estado", width="small"),
        },
        key=f"catalogo_editor_{filtro_key}",
    )

    visibles = set(editado["codigo"].tolist()) if not editado.empty else set()
    marcados_visibles = (
        set(editado.loc[editado["Sel"].astype(bool), "codigo"].tolist())
        if not editado.empty
        else set()
    )
    seleccion = (seleccion_prev - visibles) | marcados_visibles
    st.session_state["catalogo_sel_codigos"] = sorted(seleccion)
    # Orden de generacion = orden del catalogo (tabla), no alfabetico del set
    seleccion_set = set(st.session_state["catalogo_sel_codigos"])
    codigos_sel = [d["codigo"] for d in docs if d["codigo"] in seleccion_set]
    st.write(f"**{len(codigos_sel)}** documento(s) seleccionado(s)")
    if codigos_sel:
        n0 = int(num_desde)
        st.caption(
            f"Numeración prevista: **{n0:03d}** ... **{n0 + len(codigos_sel) - 1:03d}** "
            "(en el orden de la tabla / selección)."
        )

    st.markdown("#### Generar documentos")
    st.caption("Con la selección de la tabla, elige plantilla(s) y genera o descarga.")

    from generador.plantillas_index import (
        encontrar_plantilla,
        es_plantilla_usuario,
        listar_plantillas_usuario,
        registrar_plantilla_usuario,
    )

    modo_pl_gen = st.radio(
        "Plantilla al generar",
        [
            "Automática (coincidencia / asociaciones)",
            "Elegir plantilla(s) adjuntadas o del listado",
            "Desde cero (sin plantilla)",
        ],
        horizontal=True,
        key="modo_pl_gen",
        help=(
            "Automática: usa el match de la pestaña Plantillas. "
            "Elegir: escoges una o varias plantillas y las asignas a los documentos. "
            "Desde cero: ignora plantillas."
        ),
    )

    plantilla_overrides: dict[str, str] = {}
    forzar_sin = modo_pl_gen.startswith("Desde cero")

    if modo_pl_gen.startswith("Elegir"):
        st.caption(
            "Puedes adjuntar plantillas ahora y/o marcar varias del listado. "
            "Luego asigna una plantilla a cada documento seleccionado."
        )
        up_gen = st.file_uploader(
            "Adjuntar plantilla(s) para esta generación",
            type=["docx", "doc", "xlsx", "xls"],
            accept_multiple_files=True,
            key="up_pl_gen",
        )
        if up_gen and st.button("Guardar plantilla(s) adjuntadas", key="btn_save_pl_gen"):
            guardadas = []
            errores = []
            codigo_asoc = codigos_sel[0] if len(codigos_sel) == 1 else ""
            for f in up_gen:
                try:
                    path = registrar_plantilla_usuario(f.name, f.getvalue(), codigo_asoc)
                    guardadas.append(path.name)
                except Exception as e:
                    errores.append(f"{f.name}: {e}")
            if guardadas:
                st.success("Guardadas: " + ", ".join(guardadas))
                st.session_state.pop("_estado_pl_map", None)
                st.rerun()
            for err in errores:
                st.error(err)

        todas = listar_plantillas()
        propias = listar_plantillas_usuario()
        ordenadas = []
        vistos = set()
        for p in propias + todas:
            key = str(p.resolve()) if p.exists() else str(p)
            if key in vistos:
                continue
            vistos.add(key)
            ordenadas.append(p)

        labels_pl = []
        path_by_label: dict[str, Path] = {}
        for p in ordenadas:
            fuente = "adjuntada" if es_plantilla_usuario(p) else "pack"
            lab = f"[{fuente}] {p.name}"
            labels_pl.append(lab)
            path_by_label[lab] = p

        elegidas = st.multiselect(
            "Plantillas disponibles (marca una o varias)",
            options=labels_pl,
            default=[],
            key="pl_multisel_gen",
            help="Selecciona las plantillas que quieres usar en este lote.",
        )
        pool_elegidas = [path_by_label[l] for l in elegidas if l in path_by_label]
        if not pool_elegidas and propias:
            st.info(
                "No marcaste plantillas del listado. "
                "Si ya adjuntaste alguna en Plantillas, aparecerá abajo al asignar."
            )

        pool_asig = pool_elegidas if pool_elegidas else ordenadas

        if not codigos_sel:
            st.warning("Marca documentos en la columna Sel para asignarles plantilla.")
        else:
            docs_by_cod = {d["codigo"]: d for d in docs}
            misma = st.selectbox(
                "Aplicar la misma plantilla a todos (mismo formato)",
                ["(No aplicar)"]
                + [
                    f"[{'adjuntada' if es_plantilla_usuario(p) else 'pack'}] {p.name}"
                    for p in pool_asig
                ],
                key="pl_misma_todos",
            )
            if misma != "(No aplicar)" and st.button(
                "Aplicar a seleccionados", key="btn_pl_misma"
            ):
                st.session_state["pl_map_gen"] = {c: misma for c in codigos_sel}
                st.rerun()

            mapa_prev = dict(st.session_state.get("pl_map_gen", {}))
            st.markdown("**Asignar plantilla por documento**")
            for codigo in codigos_sel:
                doc = docs_by_cod.get(codigo) or {}
                fmt = doc.get("formato") or ""
                opciones = ["(Automática)"]
                path_opts: dict[str, Path] = {}
                for p in pool_asig:
                    if fmt == "Word" and p.suffix.lower() != ".docx":
                        continue
                    if fmt == "Excel" and p.suffix.lower() != ".xlsx":
                        continue
                    lab = f"[{'adjuntada' if es_plantilla_usuario(p) else 'pack'}] {p.name}"
                    opciones.append(lab)
                    path_opts[lab] = p
                default_lab = mapa_prev.get(codigo, "(Automática)")
                if default_lab not in opciones:
                    auto_p = encontrar_plantilla(doc) if doc else None
                    if auto_p:
                        lab_auto = (
                            f"[{'adjuntada' if es_plantilla_usuario(auto_p) else 'pack'}] "
                            f"{auto_p.name}"
                        )
                        default_lab = lab_auto if lab_auto in opciones else "(Automática)"
                    else:
                        default_lab = "(Automática)"
                idx = opciones.index(default_lab) if default_lab in opciones else 0
                elegido = st.selectbox(
                    f"{codigo} — {(doc.get('nombre') or '')[:55]}",
                    opciones,
                    index=idx,
                    key=f"pl_asig_{codigo}",
                )
                mapa_prev[codigo] = elegido
                if elegido != "(Automática)" and elegido in path_opts:
                    plantilla_overrides[codigo] = str(path_opts[elegido])
            st.session_state["pl_map_gen"] = mapa_prev
            if plantilla_overrides:
                st.caption(
                    f"{len(plantilla_overrides)} documento(s) con plantilla forzada; "
                    "el resto usará coincidencia automática."
                )

    c_gen_top, c_save_top, c_all_top = st.columns([2, 2, 2])

    with c_gen_top:
        gen_btn_top = st.button(
            f"Generar seleccionados ({len(codigos_sel)})",
            type="primary",
            disabled=not codigos_sel,
            key="gen_sel_top",
            use_container_width=True,
        )
    with c_save_top:
        save_ver_btn_top = st.button(
            "Guardar versión en seleccionados",
            disabled=not codigos_sel,
            key="save_ver_top",
            use_container_width=True,
            help="Actualiza la versión sin generar.",
        )
    with c_all_top:
        gen_all_top = st.button(
            f"Generar todo el catálogo ({len(docs)})",
            key="gen_all_top",
            use_container_width=True,
        )
    codigos_sel_top = list(codigos_sel)

    def _args_version(doc_codigo: str) -> dict:
        if modo_version == "Usar versión indicada":
            return {"version": int(version_manual), "auto_bump": False}
        if modo_version == "Incrementar automáticamente":
            return {"version": None, "auto_bump": True}
        return {"version": None, "auto_bump": False}

    def _mostrar_descargas(archivos: list[dict]) -> None:
        """Muestra botones de descarga; el usuario elige dónde guardar en el navegador."""
        if not archivos:
            return
        st.subheader("Descargas")
        st.caption(
            "Los archivos no se guardan en el servidor ni en tu escritorio. "
            "Usa los botones y elige la carpeta en el cuadro de descarga del navegador."
        )
        if len(archivos) > 1:
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for a in archivos:
                    zf.writestr(a["nombre_archivo"], a["contenido"])
            st.download_button(
                label=f"Descargar ZIP ({len(archivos)} documentos)",
                data=buf.getvalue(),
                file_name=f"SG-SST_documentos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                type="primary",
                key="dl_zip_batch",
            )
        for i, a in enumerate(archivos):
            st.download_button(
                label=f"Descargar {a['nombre_archivo']} (v{a['version']})",
                data=a["contenido"],
                file_name=a["nombre_archivo"],
                mime=a["mime"],
                key=f"dl_file_{i}_{a['codigo']}",
            )

    def _generar_lista(codigos: list[str]) -> None:
        if not codigos:
            st.warning("Selecciona al menos un documento (columna Sel).")
            return
        ok, fail = 0, 0
        generados: list[dict] = []
        inicio = int(num_desde)
        progress = st.progress(0)
        status = st.empty()
        for i, codigo in enumerate(codigos):
            numero = inicio + i
            status.text(
                f"Generando {codigo} -> NNN {numero:03d} ({i + 1}/{len(codigos)})"
            )
            try:
                kwargs = dict(_args_version(codigo))
                if forzar_sin:
                    kwargs["forzar_sin_plantilla"] = True
                elif codigo in plantilla_overrides:
                    kwargs["plantilla_path"] = plantilla_overrides[codigo]
                res = generar_documento(
                    codigo,
                    fecha=fecha_str,
                    numero_asignado=numero,
                    **kwargs,
                )
                generados.append(
                    {
                        "codigo": res["codigo"],
                        "nombre_archivo": res["nombre_archivo"],
                        "contenido": res["contenido"],
                        "mime": res["mime"],
                        "version": res["version"],
                        "modo": res["modo"],
                        "numero": res["numero"],
                    }
                )
                ok += 1
            except Exception as e:
                fail += 1
                st.warning(f"{codigo}: {e}")
            progress.progress((i + 1) / len(codigos))
        status.empty()
        if generados:
            ultimo = max(g["numero"] for g in generados)
            db.set_ultimo_consecutivo(ultimo)
            st.session_state["num_desde_pending"] = ultimo + 1
        st.session_state["descargas_listas"] = generados
        st.session_state["flash_gen"] = (
            f"Listo. Exitosos: {ok} | Fallidos: {fail}. "
            f"Próximo número sugerido: {db.get_siguiente_sugerido():03d}"
        )
        st.rerun()

    # Acciones (tambien debajo de la tabla, por si se edito Sel)
    st.markdown("#### Acciones")
    c_gen, c_save_ver = st.columns(2)
    with c_gen:
        gen_btn = st.button(
            f"Generar seleccionados ({len(codigos_sel)})",
            type="primary",
            disabled=not codigos_sel,
            key="gen_sel_bottom",
            use_container_width=True,
        )
    with c_save_ver:
        save_ver_btn = st.button(
            "Guardar versión en seleccionados",
            disabled=not codigos_sel,
            key="save_ver_bottom",
            use_container_width=True,
            help="Actualiza la versión de los documentos marcados sin generarlos.",
        )

    if (save_ver_btn or save_ver_btn_top) and (codigos_sel or codigos_sel_top):
        destino = codigos_sel if (save_ver_btn and codigos_sel) else codigos_sel_top
        try:
            for codigo in destino:
                db.set_version(codigo, int(version_manual))
            st.success(
                f"Versión v{int(version_manual)} guardada en {len(destino)} documento(s)."
            )
            st.rerun()
        except Exception as e:
            st.error(f"Error: {e}")

    if gen_btn_top and codigos_sel_top:
        _generar_lista(codigos_sel_top)
    if gen_btn and codigos_sel:
        _generar_lista(codigos_sel)
    if gen_all_top:
        _generar_lista([d["codigo"] for d in docs])

    st.divider()
    st.subheader("Generar todos")
    st.warning(
        f"Esto generará los {len(docs)} documentos del catálogo "
        "(sin duplicados ni formatos de herramientas). Puede tardar varios minutos. "
        "Luego podrás descargarlos (ZIP o uno a uno)."
    )
    if st.button("Generar todo el catálogo"):
        _generar_lista([d["codigo"] for d in docs])

    if st.session_state.get("descargas_listas"):
        if st.session_state.get("flash_gen"):
            st.success(st.session_state.pop("flash_gen"))
        _mostrar_descargas(st.session_state["descargas_listas"])
        if st.button("Limpiar lista de descargas"):
            st.session_state["descargas_listas"] = []
            st.rerun()

    st.divider()
    st.subheader("Agregar documento (si no está en el catálogo)")
    st.caption(
        "Elige el tipo de documento (o escribe uno libre), dale un nombre y se asignará "
        "automáticamente un código `SST-SG-[ABR]-[NNN]` nuevo. Opcionalmente puedes generarlo al instante."
    )
    if st.session_state.get("flash_catalogo"):
        st.success(st.session_state.pop("flash_catalogo"))
    if st.session_state.get("flash_catalogo_err"):
        st.error(st.session_state.pop("flash_catalogo_err"))

    TIPOS = [
        ("POL", "politica", "Política"),
        ("PRO", "procedimiento", "Procedimiento"),
        ("PL", "plan", "Plan"),
        ("PRG", "programa", "Programa"),
        ("MT", "matriz", "Matriz"),
        ("FR", "formato", "Formato / Registro"),
        ("ACT", "acta", "Acta"),
        ("E", "evaluacion", "Evaluación"),
    ]
    tipo_labels = [f"{abr} - {label}" for abr, _, label in TIPOS] + ["Otro (escribir libre)"]
    tipo_sel = st.selectbox("Tipo de documento a generar", tipo_labels, key="nuevo_tipo_sel")

    abr_libre = ""
    familia_libre = "formato"
    if tipo_sel.startswith("Otro"):
        c_lib1, c_lib2 = st.columns(2)
        with c_lib1:
            abr_libre = st.text_input(
                "Abreviatura libre (ej. INF, MAN, INS)",
                value="",
                max_chars=12,
                help="Se usará en el código SST-SG-[ABR]-[NNN]. Solo letras/números.",
                key="nuevo_abr_libre",
            )
        with c_lib2:
            familia_libre = st.text_input(
                "Nombre del tipo / familia (libre)",
                value="",
                help="Ej: Informe, Instructivo, Circular, Manual, etc.",
                key="nuevo_fam_libre",
            )

    with st.form("nuevo_doc"):
        nombre = st.text_input("Nombre del documento")
        c_fmt, c_ver_n = st.columns(2)
        with c_fmt:
            formato = st.selectbox("Formato de salida", ["Word", "Excel"])
        with c_ver_n:
            version_nuevo = st.number_input(
                "Versión inicial",
                min_value=1,
                max_value=999,
                value=int(version_manual),
                step=1,
            )
        ref = st.text_input("Ruta / referencia (opcional)")
        c_add, c_gen = st.columns(2)
        with c_add:
            solo_agregar = st.form_submit_button("Agregar al catálogo")
        with c_gen:
            agregar_y_generar = st.form_submit_button("Agregar y generar ahora", type="primary")
        if (solo_agregar or agregar_y_generar) and nombre.strip():
            if tipo_sel.startswith("Otro"):
                abr = "".join(c for c in (abr_libre or "").upper() if c.isalnum())[:12]
                familia = (familia_libre or "otro").strip().lower() or "otro"
                if not abr:
                    st.warning("Indica una abreviatura válida (letras/números).")
            else:
                idx = tipo_labels.index(tipo_sel)
                abr, familia, _ = TIPOS[idx]

            if abr:
                codigo = db.add_documento_catalogo(
                    abr,
                    nombre.strip(),
                    formato,
                    familia,
                    ref,
                    version=int(version_nuevo),
                )
                if agregar_y_generar:
                    try:
                        res = generar_documento(
                            codigo,
                            fecha=fecha_str,
                            version=int(version_nuevo),
                            auto_bump=False,
                            numero_asignado=int(num_desde),
                        )
                        db.set_ultimo_consecutivo(int(res["numero"]))
                        st.session_state["num_desde_pending"] = int(res["numero"]) + 1
                        st.session_state["descargas_listas"] = [
                            {
                                "codigo": res["codigo"],
                                "nombre_archivo": res["nombre_archivo"],
                                "contenido": res["contenido"],
                                "mime": res["mime"],
                                "version": res["version"],
                                "modo": res["modo"],
                                "numero": res["numero"],
                            }
                        ]
                        st.session_state["flash_catalogo"] = (
                            f"Agregado y generado {res['codigo']} v{res['version']}. "
                            "Usa el botón de descarga abajo."
                        )
                    except Exception as e:
                        st.session_state["flash_catalogo_err"] = (
                            f"Agregado {codigo}, pero falló la generación: {e}"
                        )
                else:
                    st.session_state["flash_catalogo"] = (
                        f"Agregado al catlogo: {codigo} (v{int(version_nuevo)})"
                    )
                st.rerun()
        elif (solo_agregar or agregar_y_generar) and not nombre.strip():
            st.warning("Indica el nombre del documento.")


def pagina_historial():
    st.header("4. Historial y versiones")
    emp = db.get_empresa()
    st.caption(
        f"Empresa activa: **{emp.get('razon_social') or '(sin nombre)'}** "
        f"(NIT {emp.get('nit') or '-'}) - el historial es solo de esta empresa."
    )
    st.caption(
        f"Último NNN usado: **{db.get_ultimo_consecutivo():03d}** | "
        f"Siguiente sugerido: **{db.get_siguiente_sugerido():03d}**"
    )
    hist = db.list_historial()
    if not hist:
        st.info("Aún no hay generaciones registradas para esta empresa.")
        return

    busqueda = st.text_input("Buscar en historial (código o nombre)", "")
    if busqueda.strip():
        q = busqueda.strip().lower()
        hist = [
            h for h in hist
            if q in str(h.get("codigo") or "").lower()
            or q in str(h.get("nombre") or "").lower()
        ]

    vista = [
        {
            "Código": h.get("codigo"),
            "Nombre": h.get("nombre"),
            "Versión": h.get("version"),
            "Fecha": h.get("fecha"),
            "Archivo": h.get("archivo"),
            "Generado en": h.get("generado_en"),
        }
        for h in hist
    ]
    st.write(f"**{len(vista)}** registro(s)")
    st.dataframe(vista, use_container_width=True, hide_index=True)

    if vista:
        csv_buf = io.StringIO()
        pd.DataFrame(vista).to_csv(csv_buf, index=False)
        st.download_button(
            "Descargar historial (CSV)",
            data=csv_buf.getvalue().encode("utf-8-sig"),
            file_name=f"historial_sg_sst_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            key="dl_hist_csv",
        )


def main():
    _aplicar_estilo()
    pagina = _barra_lateral()
    if pagina != "Empresa":
        st.title("SG-SST - Gestión documental")
        emp = db.get_empresa()
        st.caption(
            f"Empresa activa: {emp.get('razon_social') or '-'} | NIT {emp.get('nit') or '-'}."
        )

    if pagina == "Empresa":
        pagina_empresa()
    elif pagina == "Plantillas":
        pagina_plantillas()
    elif pagina.startswith("Cat"):
        pagina_catalogo()
    else:
        pagina_historial()



if __name__ == "__main__":
    main()
