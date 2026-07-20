# -*- coding: utf-8 -*-
"""App de Gestión Documental SG-SST - Streamlit."""
from __future__ import annotations

import sys
import io
import zipfile
from datetime import date, datetime
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
    if not db.LOGO_PATH.exists():
        candidatos = [
            ROOT.parent / "SUGANUS-TRABAJOS" / "LOGO" / "LOGO SUGANUS .png",
            ROOT.parent / "SUGANUS-TRABAJOS" / "LOGO" / "LOGO SUGANUS.png",
        ]
        for c in candidatos:
            if c.exists():
                db.LOGO_PATH.parent.mkdir(parents=True, exist_ok=True)
                db.LOGO_PATH.write_bytes(c.read_bytes())
                emp = db.get_empresa()
                emp["logo_path"] = str(db.LOGO_PATH)
                db.save_empresa(emp)
                break


_ensure()


def _barra_lateral() -> None:
    """Estado rapido y guia para quien prueba la app por primera vez."""
    emp = db.get_empresa()
    with st.sidebar:
        st.markdown("### SG-SST")
        st.caption("Gestión documental")
        st.markdown(
            f"**Empresa**\n\n"
            f"{emp.get('razon_social') or '(sin nombre)'}\n\n"
            f"NIT: {emp.get('nit') or '-'}"
        )
        st.divider()
        n_docs = len(db.list_catalogo())
        st.metric("Documentos en catálogo", n_docs)
        st.metric("Siguiente NNN", f"{db.get_siguiente_sugerido():03d}")
        modo = db.get_modo_plantillas()
        st.caption(
            "Plantillas: "
            + ("pack base + adjuntadas" if modo == "todas" else "solo adjuntadas")
        )
        pack = carpeta_base_gestion_documental()
        if pack:
            st.success("Pack de plantillas disponible")
        else:
            st.warning("Sin pack base: se generará desde cero")
        st.divider()
        with st.expander("Guía rápida", expanded=False):
            st.markdown(
                "1. **Empresa**: logo, NIT y responsables.\n"
                "2. **Plantillas**: deja *Pack base + adjuntadas*.\n"
                "3. **Catálogo / Generar**: marca **Sel**, genera y **descarga**.\n"
                "4. **Historial**: revisa lo generado por esta empresa.\n\n"
                "Tip: cada empresa tiene su propia numeración."
            )
        if not (emp.get("razon_social") or "").strip():
            st.info("Completa la razón social en la pestaña Empresa.")


def pagina_empresa():
    st.header("1. Configuración de la empresa")
    st.caption(
        "Cada empresa tiene su propio historial y numeración. "
        "Al cambiar de empresa, el historial y el consecutivo cambian con ella."
    )

    empresas = db.list_empresas()
    activa_id = db.get_empresa_activa_id()
    labels = {
        f"{e['id']}: {e.get('razon_social') or '(sin nombre)'} | NIT {e.get('nit') or '-'}": e["id"]
        for e in empresas
    }
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
        if st.button("Nueva empresa"):
            db.crear_empresa_nueva()
            st.session_state["num_desde_pending"] = 1
            st.session_state["descargas_listas"] = []
            st.success("Empresa nueva creada (historial y numeración en cero).")
            st.rerun()

    emp = db.get_empresa()

    col1, col2 = st.columns([1, 2])
    with col1:
        st.subheader("Logo")
        logo = emp.get("logo_path") or str(db.LOGO_PATH)
        if Path(logo).exists():
            st.image(logo, width=180)
        elif db.LOGO_PATH.exists():
            st.image(str(db.LOGO_PATH), width=180)
        up = st.file_uploader("Subir logo (png/jpg)", type=["png", "jpg", "jpeg"])
        if up is not None:
            img = Image.open(up)
            db.LOGO_PATH.parent.mkdir(parents=True, exist_ok=True)
            # logo por empresa
            dest = db.DATA_DIR / f"logo_emp_{db.get_empresa_activa_id()}.png"
            img.save(dest)
            emp_data = dict(emp)
            emp_data["logo_path"] = str(dest)
            emp_data["clases_riesgo_list"] = emp.get("clases_riesgo_list") or []
            db.save_empresa(emp_data)
            st.success("Logo guardado para esta empresa.")
            st.rerun()

    with col2:
        razon = st.text_input("Razón social", emp.get("razon_social", ""))
        nit = st.text_input("NIT", emp.get("nit", ""))
        c1, c2 = st.columns(2)
        with c1:
            ciiu_cod = st.text_input("CIIU (código)", emp.get("ciiu_codigo", ""))
        with c2:
            ciiu_desc = st.text_input("CIIU (descripción)", emp.get("ciiu_descripcion", ""))
        arl = st.text_input("ARL", emp.get("arl", ""))
        clases_default = ", ".join(emp.get("clases_riesgo_list") or [])
        clases = st.text_input(
            "Clase(s) de riesgo (separadas por coma)",
            clases_default,
            help="Ej: I - administrativos, III - operarios",
        )

    st.subheader("Responsables")
    a, b = st.columns(2)
    with a:
        rep_n = st.text_input("Representante legal - nombre", emp.get("rep_legal_nombre", ""))
        rep_c = st.text_input("Representante legal - cargo", emp.get("rep_legal_cargo", ""))
    with b:
        sst_n = st.text_input("Responsable SST - nombre", emp.get("resp_sst_nombre", ""))
        sst_c = st.text_input("Responsable SST - cargo", emp.get("resp_sst_cargo", "Responsable SST"))
        sst_l = st.text_input("Licencia SST (opcional)", emp.get("resp_sst_licencia", ""))
    vigia = st.text_input(
        "Vigía / COPASST - nombre (opcional)",
        emp.get("vigia_nombre", ""),
        help="Si queda vacío, en firmas se deja línea en blanco.",
    )

    if st.button("Guardar configuración", type="primary"):
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
                "logo_path": emp.get("logo_path") or str(db.LOGO_PATH),
            }
        )
        st.success("Datos de la empresa activa guardados.")


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
    st.header("3. Catálogo y generación")
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
                res = generar_documento(
                    codigo,
                    fecha=fecha_str,
                    numero_asignado=numero,
                    **_args_version(codigo),
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

    c_gen, c_save_ver = st.columns(2)
    with c_gen:
        gen_btn = st.button(
            f"Generar seleccionados ({len(codigos_sel)})",
            type="primary",
            disabled=not codigos_sel,
        )
    with c_save_ver:
        save_ver_btn = st.button(
            "Guardar versión en seleccionados",
            disabled=not codigos_sel,
            help="Actualiza la versión de los documentos marcados sin generarlos.",
        )

    if save_ver_btn and codigos_sel:
        try:
            for codigo in codigos_sel:
                db.set_version(codigo, int(version_manual))
            st.success(
                f"Versión v{int(version_manual)} guardada en {len(codigos_sel)} documento(s)."
            )
            st.rerun()
        except Exception as e:
            st.error(f"Error: {e}")

    if gen_btn:
        _generar_lista(codigos_sel)

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
    _barra_lateral()
    st.title("SG-SST - Gestión documental")
    emp = db.get_empresa()
    st.caption(
        f"Empresa activa: {emp.get('razon_social') or '-'} | NIT {emp.get('nit') or '-'}. "
        "Historial y numeración son por empresa."
    )
    if "bienvenida_ok" not in st.session_state:
        st.session_state["bienvenida_ok"] = False
    if not st.session_state["bienvenida_ok"]:
        with st.expander("Bienvenida / cómo empezar", expanded=not bool(emp.get("razon_social"))):
            st.markdown(
                "Esta app genera documentos del **SG-SST** con códigos "
                "`SST-SG-[ABR]-[NNN]`.\n\n"
                "1. Configura tu **empresa** (logo y datos).\n"
                "2. Revisa **plantillas** (pack base + las tuyas).\n"
                "3. En **catálogo**, marca documentos y genera.\n"
                "4. **Descarga** los archivos (ZIP o uno a uno).\n"
            )
            if st.button("Entendido, no mostrar de nuevo"):
                st.session_state["bienvenida_ok"] = True
                st.rerun()

    tab1, tab2, tab3, tab4 = st.tabs(
        ["Empresa", "Plantillas", "Catálogo / Generar", "Historial"]
    )
    with tab1:
        pagina_empresa()
    with tab2:
        pagina_plantillas()
    with tab3:
        pagina_catalogo()
    with tab4:
        pagina_historial()


if __name__ == "__main__":
    main()
