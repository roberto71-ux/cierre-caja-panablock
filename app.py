"""
GRUPO PANABLOCK — Cierre Diario de Caja
Interfaz web con Streamlit.

Streamlit re-ejecuta este script completo en cada interacción del usuario.
st.session_state actúa como memoria entre re-ejecuciones.
"""

import io
import os
import traceback

import pandas as pd
import streamlit as st
from datetime import date

import cierre_caja_core as core
from cierre_caja_core import (
    parse_facturas, parse_recibos,
    parse_banco_general, parse_global_bank,
    generate_report, generate_pdf_report, _merge_gastos, ITEMS_PAGOS,
)

# ════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE PÁGINA
# ════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Cierre de Caja — Panablock",
    page_icon="📊",
    layout="centered",
)

CATS = [
    'Acreedores', 'Mantenimiento de Planta', 'Materia Prima',
    'Pigmento y Aditivos', 'Servicios Básicos', 'Tarimas',
    'Transporte', 'Comisiones Bancarias', 'Egresos Varios', 'Otros',
]

# ════════════════════════════════════════════════════════════════════
# CARGAR BASE DE PROVEEDORES
# (cache_resource: se carga UNA sola vez, no en cada clic del usuario)
# ════════════════════════════════════════════════════════════════════
@st.cache_resource
def load_proveedores():
    """
    Lee proveedores.xlsx y devuelve un dict  KEYWORD_EN_MAYUSCULAS → Categoría.
    Este dict le indica al clasificador a qué categoría pertenece cada proveedor.
    """
    prov_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "proveedores.xlsx")
    if not os.path.exists(prov_path):
        return {}
    try:
        df = pd.read_excel(prov_path)
        return dict(zip(df["Keyword"].str.upper(), df["Categoría"]))
    except Exception:
        return {}

# Inyectar el mapa en el módulo core para que _classify_gasto lo use
core._PROV_MAP = load_proveedores()

# ════════════════════════════════════════════════════════════════════
# HELPER: aplicar reclasificaciones del usuario a los datos de bancos
# ════════════════════════════════════════════════════════════════════
def _apply_reclassifications(bg, gb, edited_df):
    """
    Toma el DataFrame editado por el usuario (tabla de gastos "Otros")
    y actualiza las categorías dentro de bg['gastos_detalle'] y
    gb['gastos_detalle'], recalculando los totales por categoría.
    """
    corrections = {}
    for _, row in edited_df.iterrows():
        key = (row['Banco'], str(row['Descripción']), float(row['Monto']))
        corrections[key] = row['Categoría']

    def _fix_bank(data, banco_name):
        for itm in data.get('gastos_detalle', []):
            key = (banco_name, itm['desc'], itm['monto'])
            if key in corrections:
                itm['cat'] = corrections[key]
        # Reconstruir totales por categoría desde gastos_detalle
        cats = {c: 0.0 for c in ITEMS_PAGOS}
        for itm in data.get('gastos_detalle', []):
            cats[itm['cat']] = round(cats.get(itm['cat'], 0.0) + itm['monto'], 2)
        data['gastos'] = cats

    _fix_bank(bg, 'Banco General')
    _fix_bank(gb, 'Global Bank')

# ════════════════════════════════════════════════════════════════════
# ENCABEZADO
# ════════════════════════════════════════════════════════════════════
st.markdown("""
<div style="background:#1F4E79;padding:18px 22px;border-radius:8px;margin-bottom:24px">
  <h2 style="color:white;margin:0;font-family:Arial,sans-serif;font-size:22px">
      📊 GRUPO PANABLOCK
  </h2>
  <p style="color:#BDD7EE;margin:5px 0 0;font-family:Arial,sans-serif;font-size:14px">
      Generador de Cierre Diario de Caja
  </p>
</div>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════
# FECHA
# ════════════════════════════════════════════════════════════════════
fecha = st.date_input(
    "📅 Fecha del cierre",
    value=date.today(),
    format="DD/MM/YYYY",
)

st.markdown("#### Archivos de entrada")
st.caption("Carga los 4 archivos del día. Los archivos solo existen durante esta "
           "sesión — no quedan guardados en ningún servidor.")

# ════════════════════════════════════════════════════════════════════
# CARGADORES DE ARCHIVOS
# (st.file_uploader: el usuario arrastra o selecciona el archivo)
# ════════════════════════════════════════════════════════════════════
col1, col2 = st.columns(2)
with col1:
    f_fact = st.file_uploader(
        "📄 Facturas del día (PDF)",
        type=["pdf"],
        key="up_fact",
        accept_multiple_files=True,
        help="Puedes subir uno o varios PDFs — se procesan todos juntos",
    )
    f_bg = st.file_uploader(
        "🏦 Extracto Banco General (PDF)",
        type=["pdf"],
        key="up_bg",
        help="Extracto del día del Banco General",
    )
with col2:
    f_rec = st.file_uploader(
        "📄 Recibos del día (PDF)",
        type=["pdf"],
        key="up_rec",
        accept_multiple_files=True,
        help="Puedes subir uno o varios PDFs — se procesan todos juntos",
    )
    f_gb = st.file_uploader(
        "🏦 Extracto Global Bank (XLS)",
        type=["xls"],
        key="up_gb",
        help="Extracto del día de Global Bank en formato .xls",
    )

todos_cargados = all([f_fact, f_rec, f_bg])  # Global Bank es opcional

if not todos_cargados:
    archivos_faltantes = [
        nombre for nombre, f in [
            ("Facturas", f_fact), ("Recibos", f_rec),
            ("Banco General", f_bg),
        ] if not f
    ]
    if not f_gb:
        archivos_faltantes.append("Global Bank (opcional)")
    st.info(f"⬆️ Faltan: {', '.join(archivos_faltantes)}")
elif not f_gb:
    st.info("⬆️ Falta: Global Bank (opcional)")

# ── Campos opcionales: solo aparecen una vez que los 3 archivos obligatorios están cargados ──
if todos_cargados:
    # Global Bank sin extracto: el usuario confirma que no tiene el archivo
    if not f_gb:
        no_gb = st.checkbox(
            "No tengo extracto de Global Bank para hoy",
            key="no_gb_check",
        )
        if no_gb:
            st.caption("Ingresa el saldo anterior de Global Bank para completar el reporte:")
            st.number_input(
                "Saldo Anterior Global Bank ($)",
                min_value=0.0, value=0.0, step=0.01, format="%.2f",
                key="gb_saldo_ant_manual",
            )

    # Cheques en circulación: visibles para ambos bancos
    st.markdown("##### 🔲 Cheques en Circulación")
    st.caption("Ingresa el monto de cheques emitidos pendientes de cobro en cada banco (0 si no hay):")
    _col_ch1, _col_ch2 = st.columns(2)
    with _col_ch1:
        st.number_input(
            "Cheques en Circulación Banco General ($)",
            min_value=0.0, value=0.0, step=0.01, format="%.2f",
            key="bg_cheques_manual",
        )
    with _col_ch2:
        st.number_input(
            "Cheques en Circulación Global Bank ($)",
            min_value=0.0, value=0.0, step=0.01, format="%.2f",
            key="gb_cheques_manual",
        )
st.markdown("---")

# ════════════════════════════════════════════════════════════════════
# BOTÓN PROCESAR
# ════════════════════════════════════════════════════════════════════
# st.button devuelve True solo en la re-ejecución donde fue presionado.
# En todas las demás re-ejecuciones devuelve False.
procesar = st.button(
    "⚙️  Procesar archivos",
    type="primary",
    disabled=not todos_cargados,
    use_container_width=True,
)

if procesar:
    logs = []
    def log(msg):
        logs.append(msg)

    with st.spinner("Leyendo y procesando archivos…"):
        try:
            log("─" * 50)
            log("Leyendo facturas…")
            fact = parse_facturas(f_fact, log)

            log("\nLeyendo recibos…")
            rec  = parse_recibos(f_rec, log)

            log("\nLeyendo extracto Banco General…")
            bg   = parse_banco_general(f_bg, log)

            log("\nLeyendo extracto Global Bank…")
            gb   = parse_global_bank(f_gb, log)
            # Sin extracto: inyectar el saldo anterior ingresado manualmente
            if not f_gb:
                gb['saldo_anterior'] = st.session_state.get('gb_saldo_ant_manual', 0.0)
                log(f"  GB (sin extracto) saldo anterior: ${gb['saldo_anterior']:,.2f}")

            st.session_state.update({
                'fact':        fact,
                'rec':         rec,
                'bg':          bg,
                'gb':          gb,
                'logs':        logs,
                'fecha':       fecha,
                'processed':   True,
                'excel_bytes': None,   # borrar descarga anterior si la había
                'pdf_bytes':   None,
                'bg_cheques':  st.session_state.get('bg_cheques_manual', 0.0),
                'gb_cheques':  st.session_state.get('gb_cheques_manual', 0.0),
            })

        except Exception as e:
            st.error(f"❌ Error al procesar: {e}")
            with st.expander("Detalle técnico del error"):
                st.code(traceback.format_exc())
            st.stop()

# ════════════════════════════════════════════════════════════════════
# RESULTADOS  (se muestran si ya se procesaron archivos)
# ════════════════════════════════════════════════════════════════════
if st.session_state.get('processed'):
    fact = st.session_state['fact']
    rec  = st.session_state['rec']
    bg   = st.session_state['bg']
    gb   = st.session_state['gb']

    st.success("� Archivos procesados correctamente.")

    # ── Resumen rápido ──────────────────────────────────────────────
    st.markdown("#### Resumen del día")

    total_contado  = round(fact['contado_subtotal'] + fact['contado_itbms'], 2)
    total_credito  = round(fact['credito_subtotal'] + fact['credito_itbms'], 2)
    gastos_merged  = _merge_gastos(bg['gastos'], gb['gastos'])
    total_gastos   = round(sum(gastos_merged.values()), 2)

    # Fila 1: Ventas + Cobros + Gastos
    c1, c2, c3 = st.columns(3)
    c1.metric("🧾 Ventas Contado",     f"${total_contado:,.2f}")
    c2.metric("💰 Cobros a Clientes",  f"${rec['total']:,.2f}")
    c3.metric("💳 Total Gastos",        f"${total_gastos:,.2f}")

    # Fila 2: Datos de bancos
    c4, c5, c6 = st.columns(3)
    c4.metric("🏦 BG Depósitos",        f"${bg['depositos']:,.2f}")
    c5.metric("🏦 GB Depósitos",        f"${gb['depositos']:,.2f}")
    c6.metric("📈 Ventas Crédito",      f"${total_credito:,.2f}")

    # Alerta si hay notas de crédito
    ncs = fact.get('notas_credito', [])
    if ncs:
        st.warning(f"⚠️ {len(ncs)} nota(s) de crédito detectada(s) —"
                   "aparecen en la sección roja al final del reporte Excel.")

    # ── Log de procesamiento ────────────────────────────────────────
    with st.expander("📋 Ver detalle del procesamiento (log)"):
        st.code('\n'.join(st.session_state.get('logs', [])), language=None)


    # ── Botón Generar Reportes ──────────────────────────────────────
    st.markdown("---")
    generar = st.button(
        "📊  Generar Reportes (Excel y PDF)",
        type="primary",
        use_container_width=True,
    )

    if generar:
        with st.spinner("Generando reportes…"):
            fecha_str = st.session_state['fecha'].strftime('%d/%m/%Y')
            fecha_tag = st.session_state['fecha'].strftime('%d%m%Y')
            common = dict(
                fecha_str=fecha_str,
                facturas=fact,
                recibos=rec,
                bg=bg,
                gb=gb,
                output_path=None,
                cheques_bg=st.session_state.get('bg_cheques', 0.0),
                cheques_gb=st.session_state.get('gb_cheques', 0.0),
                cobros_por_forma=rec['por_forma'],
            )
            # Excel — siempre
            try:
                buf_xlsx = generate_report(**common)
                st.session_state['excel_bytes'] = buf_xlsx.getvalue()
                st.session_state['excel_fname'] = f"Cierre_Caja_{fecha_tag}_Panablock.xlsx"
            except Exception as e:
                st.error(f"❌ Error al generar Excel: {e}")
                with st.expander("Detalle técnico"):
                    st.code(traceback.format_exc())
            # PDF — independiente; si falla no bloquea el Excel
            try:
                buf_pdf = generate_pdf_report(**common)
                st.session_state['pdf_bytes'] = buf_pdf.getvalue()
                st.session_state['pdf_fname'] = f"Cierre_Caja_{fecha_tag}_Panablock.pdf"
            except Exception as e:
                st.warning(f"⚠️ PDF no disponible: {e}")

    # ── Botones de descarga ─────────────────────────────────────────
    if st.session_state.get('excel_bytes'):
        st.success("✅ Reportes listos para descargar.")
        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                label="⬇️  Descargar Excel",
                data=st.session_state['excel_bytes'],
                file_name=st.session_state['excel_fname'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with dl2:
            if st.session_state.get('pdf_bytes'):
                st.download_button(
                    label="⬇️  Descargar PDF",
                    data=st.session_state['pdf_bytes'],
                    file_name=st.session_state['pdf_fname'],
                    mime="application/pdf",
                    use_container_width=True,
                )
            else:
                st.info("PDF no disponible en este entorno.")

# ════════════════════════════════════════════════════════════════════
# PIE DE PÁGINA
# ════════════════════════════════════════════════════════════════════
st.markdown("---")
st.caption("Grupo Panablock © — Los archivos cargados en esta sesión "
           "no se almacenan en ningún servidor después de cerrar la ventana.")
