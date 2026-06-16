import os
import io
import zipfile
import datetime
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image

# Configuración del SGC y de la página
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FAVICON_PATH = os.path.join(BASE_DIR, "favicon.png")
LOGO_PATH = os.path.join(BASE_DIR, "logo_sigrama.png")

# Configurar página con favicon si existe
try:
    if os.path.exists(FAVICON_PATH):
        st.set_page_config(page_title="Control de Calidad - Recepción", layout="wide", page_icon=FAVICON_PATH)
    else:
        st.set_page_config(page_title="Control de Calidad - Recepción", layout="wide", page_icon="🔍")
except Exception:
    st.set_page_config(page_title="Control de Calidad - Recepción", layout="wide", page_icon="🔍")

# Importar funciones PDF personalizadas
import utils_pdf
import importlib
importlib.reload(utils_pdf)


# Rutas de base de datos
BD_PARAMETROS = os.path.join(BASE_DIR, "BD_Parametros_Materia_Prima.xlsx")
BD_REPORTES = os.path.join(BASE_DIR, "BD_Reportes_Incoming.xlsx")
BD_ATADOS = os.path.join(BASE_DIR, "BD_Atados_Incoming.xlsx")
PLANTILLA_PATH = os.path.join(BASE_DIR, "plantilla_incoming_calidad.xlsx")
CARPETAS_DIR = os.path.join(BASE_DIR, "carpetas_electronicas")

# Renderizado de Banner Corporativo Sigrama
BANNER_PATH = os.path.join(BASE_DIR, "banner_app.png")
if os.path.exists(BANNER_PATH):
    st.image(BANNER_PATH, use_container_width=True)
else:
    col_banner_img, col_banner_txt = st.columns([1, 4])
    with col_banner_img:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=280)
        else:
            st.subheader("INDUSTRIA SIGRAMA")
    with col_banner_txt:
        st.markdown("<h2 style='color:#D32F2F; margin-top:10px;'>SISTEMA DE CONTROL DE CALIDAD EN RECEPCIÓN (INCOMING)</h2>", unsafe_allow_html=True)
st.write("---")

# Carga de base de datos segura de forma local
def cargar_db(path, sheet=0):
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception as e:
        st.error(f"Error cargando base de datos ({os.path.basename(path)}): {e}")
        return pd.DataFrame()

def guardar_db(df, path, sheet_name="Datos_Sistema"):
    try:
        df.to_excel(path, index=False, sheet_name=sheet_name)
        return True
    except Exception as e:
        st.error(f"Error guardando base de datos ({os.path.basename(path)}): {e}")
        return False

# Inicialización de catálogos en session_state
if "BD_Parametros" not in st.session_state:
    st.session_state.BD_Parametros = cargar_db(BD_PARAMETROS, "Materia_Prima")
if "BD_Reportes" not in st.session_state:
    st.session_state.BD_Reportes = cargar_db(BD_REPORTES, "Recepciones")
if "BD_Atados" not in st.session_state:
    st.session_state.BD_Atados = cargar_db(BD_ATADOS, "Atados_Detalle")

# Asegurar compatibilidad de esquemas
if "Cantidad_Hojas" not in st.session_state.BD_Atados.columns:
    st.session_state.BD_Atados["Cantidad_Hojas"] = 0

import scipy.stats as stats
import numpy as np

def obtener_calibre(sku_str, sku_nombre=""):
    if sku_nombre:
        import re
        match = re.search(r'(?i)calibre\s+(\d+)', str(sku_nombre))
        if match:
            return f"CAL {match.group(1)}"
    for word in str(sku_str).split('-'):
        if word.isdigit():
            return f"CAL {word}"
    return "CAL 14"

def renderizar_analisis_gaussiano_atado(row, sku_info, key=None):
    esp_nom = float(row.get('Espesor_Nominal', sku_info.get('Espesor_Nominal_in', 0.0750)))
    esp_tol_min = float(row.get('Espesor_Tol_Min', sku_info.get('Espesor_Tolerancia_Min_in', -0.008)))
    esp_tol_max = float(row.get('Espesor_Tol_Max', sku_info.get('Espesor_Tolerancia_Max_in', 0.008)))
    
    lsl = esp_tol_min
    usl = esp_tol_max
    
    m1 = float(row.get('Espesor_Medido_1_in', esp_nom))
    m2 = float(row.get('Espesor_Medido_2_in', esp_nom))
    m3 = float(row.get('Espesor_Medido_3_in', esp_nom))
    
    devs = [m1 - esp_nom, m2 - esp_nom, m3 - esp_nom]
    mu = np.mean(devs)
    sigma = np.std(devs)
    if sigma < 0.0005:
        sigma = 0.0005
        
    p_below = stats.norm.cdf(lsl, loc=mu, scale=sigma)
    p_above = 1.0 - stats.norm.cdf(usl, loc=mu, scale=sigma)
    p_out_pct = (p_below + p_above) * 100.0
    
    espesor_promedio = esp_nom + mu
    dictamen_val = "ACEPTADO" if row.get('Estatus_Calidad', 'Aceptado') == 'Aceptado' else "NO ACEPTADO"
    dictamen_bg = "#e8f5e9" if dictamen_val == "ACEPTADO" else "#ffebee"
    dictamen_fg = "#2e7d32" if dictamen_val == "ACEPTADO" else "#c62828"
    dictamen_border = "#2e7d32" if dictamen_val == "ACEPTADO" else "#c62828"
    
    calibre_val = obtener_calibre(row.get('SKU', sku_info.get('SKU', '')), sku_info.get('Nombre', ''))
    tipo = row.get('Tipo_Lamina', sku_info.get('Tipo_Lamina', 'Decapada'))
    
    cabecera_html = f"""
    <div style="font-family: sans-serif; margin-top: 25px; margin-bottom: 15px;">
      <div style="display: flex; align-items: center; background-color: #f8f9fa; padding: 10px 15px; border-radius: 6px; border-left: 5px solid #1a73e8; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 12px;">
        <span style="font-size: 1.25rem; margin-right: 10px;">⚙️</span>
        <span style="font-weight: bold; color: #0d47a1; margin-right: 10px; font-size: 1rem;">{tipo}</span>
        <span style="background-color: #e8f0fe; color: #1a73e8; padding: 3px 10px; border-radius: 12px; font-size: 0.8rem; font-weight: bold; margin-right: 12px;">{calibre_val}</span>
        <span style="color: #5f6368; font-size: 0.9rem;">| Espesor Teórico: <b style="color: #1a73e8;">{esp_nom:.4f}"</b> | Tolerancia Aceptable: <b style="color: #1a73e8;">±{esp_tol_max:.4f}"</b></span>
      </div>
    """
    
    tabla_html = f"""
      <table style="width: 100%; border-collapse: collapse; text-align: left; font-size: 0.9rem; margin-bottom: 15px; border: 1px solid #e0e0e0;">
        <thead>
          <tr style="border-bottom: 1px solid #e0e0e0; background-color: #f1f3f4;">
            <th style="padding: 10px; color: #5f6368; font-weight: 600;">Rollo</th>
            <th style="padding: 10px; color: #5f6368; font-weight: 600;">Espesor Real (in)</th>
            <th style="padding: 10px; color: #5f6368; font-weight: 600;">Desviación Real (in)</th>
            <th style="padding: 10px; color: #5f6368; font-weight: 600;">% de Riesgo</th>
            <th style="padding: 10px; color: #5f6368; font-weight: 600; text-align: center;">Dictamen Final</th>
          </tr>
        </thead>
        <tbody>
          <tr style="border-bottom: 1px solid #e0e0e0;">
            <td style="padding: 12px 10px; font-weight: bold; color: #1a73e8;">{row.get('ID_Atado_Proveedor', 'Rollo')}</td>
            <td style="padding: 12px 10px; color: #1a73e8;">{espesor_promedio:.4f}"</td>
            <td style="padding: 12px 10px; color: #1a73e8;">{mu:+.4f}"</td>
            <td style="padding: 12px 10px; color: #1a73e8; font-weight: bold;">{p_out_pct:.2f}%</td>
            <td style="padding: 0px 10px; background-color: {dictamen_bg}; color: {dictamen_fg}; text-align: center; font-weight: bold; border-left: 4px solid {dictamen_border}; vertical-align: middle;">
              {dictamen_val}
            </td>
          </tr>
        </tbody>
      </table>
    </div>
    """
    
    st.markdown(cabecera_html + tabla_html, unsafe_allow_html=True)
    
    # Evaluar motivos de rechazo para mostrarlos en pantalla
    motivos_rechazo = []
    
    esp_tol_min_val = esp_nom + esp_tol_min
    esp_tol_max_val = esp_nom + esp_tol_max
    if not (esp_tol_min_val <= m1 <= esp_tol_max_val):
        motivos_rechazo.append(f"Espesor Medición 1 ({m1:.4f}\") fuera de especificación [{esp_tol_min_val:.4f}\" - {esp_tol_max_val:.4f}\"]")
    if not (esp_tol_min_val <= m2 <= esp_tol_max_val):
        motivos_rechazo.append(f"Espesor Medición 2 ({m2:.4f}\") fuera de especificación [{esp_tol_min_val:.4f}\" - {esp_tol_max_val:.4f}\"]")
    if not (esp_tol_min_val <= m3 <= esp_tol_max_val):
        motivos_rechazo.append(f"Espesor Medición 3 ({m3:.4f}\") fuera de especificación [{esp_tol_min_val:.4f}\" - {esp_tol_max_val:.4f}\"]")
        
    anc_nom = float(sku_info.get("Ancho_Nominal_in", 0))
    anc_min = anc_nom + float(sku_info.get("Ancho_Tolerancia_Min_in", 0))
    anc_max = anc_nom + float(sku_info.get("Ancho_Tolerancia_Max_in", 0))
    ancho = float(row.get("Ancho_Medido_in", anc_nom))
    if not (anc_min <= ancho <= anc_max):
        motivos_rechazo.append(f"Ancho ({ancho:.2f}\") fuera de especificación [{anc_min:.2f}\" - {anc_max:.2f}\"]")
        
    lrg_nom = float(sku_info.get("Largo_Nominal_in", 0))
    lrg_min = lrg_nom + float(sku_info.get("Largo_Tolerancia_Min_in", 0))
    lrg_max = lrg_nom + float(sku_info.get("Largo_Tolerancia_Max_in", 0))
    largo = float(row.get("Largo_Medido_in", lrg_nom))
    if not (lrg_min <= largo <= lrg_max):
        motivos_rechazo.append(f"Largo ({largo:.2f}\") fuera de especificación [{lrg_min:.2f}\" - {lrg_max:.2f}\"]")
        
    zinc_min = float(sku_info.get("Zinc_Min_oz_ft2", 0))
    zinc_val = float(row.get("Zinc_Medido_oz_ft2", 0))
    if tipo == "Galvanizada" and zinc_val < zinc_min:
        motivos_rechazo.append(f"Recubrimiento de Zinc ({zinc_val:.2f} oz/ft²) inferior al límite mínimo ({zinc_min:.2f} oz/ft²)")
        
    aceitado_val = str(row.get("Aceitado_OK", "N/A"))
    if tipo == "Decapada" and sku_info.get("Aceitado_Requerido") == "Sí" and aceitado_val.upper() != "SÍ":
        motivos_rechazo.append("Aceitado ausente o insuficiente")
        
    dureza_max = float(sku_info.get("Dureza_Max_HRB", 90))
    dureza = float(row.get("Dureza_Medida_HRB", 0))
    if dureza > dureza_max:
        motivos_rechazo.append(f"Dureza ({dureza:.1f} HRB) excede el límite máximo ({dureza_max:.1f} HRB)")
        
    peso_kg = float(row.get("Peso_Total_Kg", 0))
    if peso_kg > 2500:
        motivos_rechazo.append(f"Peso del atado ({peso_kg:.0f} Kg) excede el límite de 2.5 Toneladas (2500 Kg) según RFQ")
        
    if dictamen_val == "NO ACEPTADO" and motivos_rechazo:
        st.error("⚠️ **Motivo(s) de Rechazo:**\n" + "\n".join([f"- {m}" for m in motivos_rechazo]))
    
    x_vals = np.linspace(-0.015, 0.015, 500)
    y_vals = stats.norm.pdf(x_vals, loc=mu, scale=sigma)
    
    fig_gauss = go.Figure()
    
    fig_gauss.add_trace(go.Scatter(
        x=x_vals,
        y=y_vals,
        mode='lines',
        name=str(row.get('ID_Atado_Proveedor', 'Rollo')),
        line=dict(color='#1f77b4', width=3)
    ))
    
    fig_gauss.add_vrect(
        x0=lsl,
        x1=usl,
        fillcolor="#e8f5e9",
        opacity=0.5,
        layer="below",
        line_width=0
    )
    
    fig_gauss.add_vline(x=0, line_dash="dash", line_color="green", line_width=1.5, annotation_text="Nominal", annotation_position="top")
    fig_gauss.add_vline(x=lsl, line_dash="dot", line_color="red", line_width=1.5, annotation_text="LSL", annotation_position="top")
    fig_gauss.add_vline(x=usl, line_dash="dot", line_color="red", line_width=1.5, annotation_text="USL", annotation_position="top")
    
    fig_gauss.update_layout(
        xaxis_title="Desviación Micrométrica Real (in)",
        yaxis_title="Densidad Probabilística de Gauss",
        xaxis=dict(range=[-0.015, 0.015], gridcolor='lightgray', gridwidth=0.5),
        yaxis=dict(gridcolor='lightgray', gridwidth=0.5),
        plot_bgcolor='white',
        height=320,
        margin=dict(l=20, r=20, t=10, b=20),
        showlegend=False
    )
    
    if not key:
        id_atado = row.get('ID_Atado') or row.get('ID_Atado_Proveedor') or row.get('No_Atado') or 'default'
        key = f"plotly_gauss_{id_atado}"
    st.plotly_chart(fig_gauss, use_container_width=True, key=key)

# Navegación lateral
st.sidebar.title("🧭 Navegación")
opcion_menu = st.sidebar.radio("Seleccione un Módulo:", [
    "📊 Analíticas y Dashboard",
    "📥 Registro de Recepción (Incoming)",
    "🔍 Consulta de Historial",
    "⚙️ Catálogo de Tolerancias de SKU"
])

# Control de accesos para administración y registro en la barra lateral
st.sidebar.write("---")
st.sidebar.title("🔐 Control de Acceso")
admin_pass_input = st.sidebar.text_input("Contraseña Administrador:", type="password")
inspector_pass_input = st.sidebar.text_input("Contraseña Inspector/Registro:", type="password")

def es_admin():
    return admin_pass_input == "SigramaCalidad2026" or inspector_pass_input == "SigramaCalidad2026"

def es_inspector():
    return inspector_pass_input == "SigramaInspector2026" or admin_pass_input == "SigramaInspector2026" or es_admin()

is_admin = es_admin()
is_inspector = es_inspector()

if is_admin:
    st.sidebar.success("Modo Administrador Activo")
elif is_inspector:
    st.sidebar.success("Modo Inspector (Registro) Activo")
else:
    st.sidebar.warning("Modo Consulta Activo")

# Leyenda de Desarrollador en el panel de navegación
st.sidebar.write("---")
st.sidebar.markdown(
    "<div style='text-align: center; color: #757575; font-size: 0.8rem; font-weight: 500; font-style: italic; margin-top: 15px;'>"
    "Developed by: Jesús Morales"
    "</div>", 
    unsafe_allow_html=True
)

# =============================================================================
# MÓDULO 1: ANALÍTICAS Y DASHBOARD
# =============================================================================
if opcion_menu == "📊 Analíticas y Dashboard":
    st.title("📊 Dashboard Planta Metales - Control de Calidad en Recepción")
    
    df_rep = st.session_state.BD_Reportes
    df_atd = st.session_state.BD_Atados.copy()
    if not df_atd.empty and "Tipo_Lamina" not in df_atd.columns:
        def get_tipo_lamina(sku):
            sku_str = str(sku).upper()
            if "GALV" in sku_str:
                return "Galvanizada"
            elif "DECP" in sku_str:
                return "Decapada"
            elif "ALUM" in sku_str:
                return "Aluminio"
            return "Decapada"
        df_atd["Tipo_Lamina"] = df_atd["SKU"].apply(get_tipo_lamina)
    
    if df_rep.empty:
        st.info("No hay reportes de recepción registrados actualmente. Vaya al módulo de 'Registro de Recepción' para comenzar.")
    else:
        # --- FILTROS ---
        with st.expander("🔍 Filtros de Búsqueda y Visualización", expanded=True):
            col_f1, col_f2, col_f3, col_f4 = st.columns(4)
            with col_f1:
                fecha_inicio = st.date_input("Fecha Inicio:", datetime.date.today() - datetime.timedelta(days=90), key="dash_date_start")
                fecha_fin = st.date_input("Fecha Fin:", datetime.date.today(), key="dash_date_end")
            with col_f2:
                proveedores = ["Todos"] + sorted(list(df_rep["Proveedor"].dropna().unique().tolist()))
                prov_sel = st.selectbox("Proveedor:", proveedores, key="dash_prov")
            with col_f3:
                materiales = ["Todos"] + sorted(list(df_atd["Tipo_Lamina"].dropna().unique().tolist()))
                mat_sel = st.selectbox("Material:", materiales, key="dash_mat")
            with col_f4:
                calibres = ["Todos"] + sorted(list(df_atd["SKU"].apply(lambda x: obtener_calibre(x)).unique().tolist()))
                cal_sel = st.selectbox("Calibre:", calibres, key="dash_cal")
                
        # Filtrado de datos
        df_rep_filtered = df_rep.copy()
        df_rep_filtered["Fecha_DT"] = pd.to_datetime(df_rep_filtered["Fecha"], format="%d/%m/%Y", errors="coerce").dt.date
        df_rep_filtered = df_rep_filtered[(df_rep_filtered["Fecha_DT"] >= fecha_inicio) & (df_rep_filtered["Fecha_DT"] <= fecha_fin)]
        
        if prov_sel != "Todos":
            df_rep_filtered = df_rep_filtered[df_rep_filtered["Proveedor"] == prov_sel]
            
        folios_validos = df_rep_filtered["Folio"].unique()
        df_atd_filtered = df_atd[df_atd["Folio"].isin(folios_validos)].copy()
        
        if mat_sel != "Todos":
            df_atd_filtered = df_atd_filtered[df_atd_filtered["Tipo_Lamina"] == mat_sel]
            df_rep_filtered = df_rep_filtered[df_rep_filtered["Folio"].isin(df_atd_filtered["Folio"].unique())]
            
        if cal_sel != "Todos":
            df_atd_filtered = df_atd_filtered[df_atd_filtered["SKU"].apply(lambda x: obtener_calibre(x)) == cal_sel]
            df_rep_filtered = df_rep_filtered[df_rep_filtered["Folio"].isin(df_atd_filtered["Folio"].unique())]
            
        # Calcular KPIs
        total_lotes = len(df_rep_filtered)
        total_atados = len(df_atd_filtered)
        
        lotes_aceptados = len(df_rep_filtered[df_rep_filtered["Estatus_General"] == "Aceptado"])
        lotes_condicionados = len(df_rep_filtered[df_rep_filtered["Estatus_General"] == "Condicionado"])
        lotes_rechazados = len(df_rep_filtered[df_rep_filtered["Estatus_General"] == "Rechazado"])
        
        atados_aceptados = len(df_atd_filtered[df_atd_filtered["Estatus_Calidad"] == "Aceptado"])
        atados_rechazados = len(df_atd_filtered[df_atd_filtered["Estatus_Calidad"] == "Rechazado"])
        
        peso_total_kg = df_atd_filtered["Peso_Total_Kg"].sum()
        peso_total_lb = df_atd_filtered["Peso_Total_Lb"].sum()
        
        # Mostrar Objetivos y Resultados Clave (OKRs)
        st.markdown("<h3 style='color:#D32F2F;'>🎯 OKR 1: Conformidad y Calidad de Materia Prima</h3>", unsafe_allow_html=True)
        st.markdown("**Objetivo:** Garantizar que el 100% de la materia prima liberada cumpla con las especificaciones técnicas y dimensionales del SGC.")
        
        o1_col1, o1_col2, o1_col3 = st.columns(3)
        
        tasa_atados = (atados_aceptados / total_atados * 100) if total_atados > 0 else 100.0
        val_diff_atados = tasa_atados - 95.0
        o1_col1.metric(
            label="🔑 KR 1.1: Conformidad de Atados (Meta: ≥95%)",
            value=f"{tasa_atados:.1f}%",
            delta=f"{val_diff_atados:+.1f}% vs Meta (95%)",
            delta_color="normal"
        )
        
        tasa_lotes = (lotes_aceptados / total_lotes * 100) if total_lotes > 0 else 100.0
        val_diff_lotes = tasa_lotes - 90.0
        o1_col2.metric(
            label="🔑 KR 1.2: Aceptación Total de Lotes (Meta: ≥90%)",
            value=f"{tasa_lotes:.1f}%",
            delta=f"{val_diff_lotes:+.1f}% vs Meta (90%)",
            delta_color="normal"
        )
        
        peso_aceptado_kg = df_atd_filtered[df_atd_filtered["Estatus_Calidad"] == "Aceptado"]["Peso_Total_Kg"].sum()
        tasa_peso = (peso_aceptado_kg / peso_total_kg * 100) if peso_total_kg > 0 else 100.0
        val_diff_peso = tasa_peso - 95.0
        o1_col3.metric(
            label="🔑 KR 1.3: Eficiencia de Acero Conforme (Meta: ≥95%)",
            value=f"{tasa_peso:.1f}%",
            delta=f"{val_diff_peso:+.1f}% vs Meta (95%)",
            delta_color="normal"
        )
        
        st.markdown("<h3 style='color:#D32F2F;'>🎯 OKR 2: Eficiencia de Abastecimiento y Control</h3>", unsafe_allow_html=True)
        st.markdown("**Objetivo:** Registrar, medir e inspeccionar el volumen total de acero recibido para asegurar la continuidad de producción.")
        
        o2_col1, o2_col2, o2_col3 = st.columns(3)
        o2_col1.metric(
            label="🔑 KR 2.1: Atados Controlados",
            value=f"{total_atados} Atados",
            delta="Volumen en rango" if total_atados > 0 else "Sin ingresos"
        )
        o2_col2.metric(
            label="🔑 KR 2.2: Acero Recibido e Inspeccionado",
            value=f"{peso_total_kg:,.0f} Kg",
            delta=f"{peso_total_lb:,.0f} Lb"
        )
        o2_col3.metric(
            label="🔑 KR 2.3: Lotes Procesados",
            value=f"{total_lotes} Lotes",
            delta=f"Aceptados: {lotes_aceptados} | Recl.: {lotes_rechazados}"
        )
        
        st.write("---")
        
        # Distribución de Estatus
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.subheader("📋 Estado General de Lotes Recibidos")
            if not df_rep_filtered.empty:
                fig_pie = px.pie(
                    df_rep_filtered, 
                    names="Estatus_General", 
                    color="Estatus_General",
                    color_discrete_map={"Aceptado": "#2E7D32", "Condicionado": "#FBC02D", "Rechazado": "#C62828"},
                    hole=0.4
                )
                st.plotly_chart(fig_pie, use_container_width=True)
            else:
                st.info("Sin datos de lotes para mostrar en este rango.")
            
        with col_g2:
            st.subheader("🏢 Distribución de Volumen por Proveedor (Kg)")
            df_prov = df_atd_filtered.merge(df_rep_filtered[["Folio", "Proveedor"]], on="Folio", how="left")
            if not df_prov.empty and "Proveedor" in df_prov.columns:
                df_prov_grouped = df_prov.groupby("Proveedor")["Peso_Total_Kg"].sum().reset_index()
                fig_bar = px.bar(
                    df_prov_grouped, 
                    x="Proveedor", 
                    y="Peso_Total_Kg", 
                    color="Proveedor",
                    title="Kg Recibidos por Proveedor",
                    color_discrete_sequence=px.colors.qualitative.Set1
                )
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.info("Sin datos de peso disponibles para graficar.")
                
        # --- NUEVA SECCIÓN DE ANÁLISIS DE CALIDAD EN ATADOS Y CAUSAS DE RECHAZO ---
        st.write("---")
        st.write("### 🔍 Análisis de Calidad de Atados (Rollos)")
        col_ca1, col_ca2 = st.columns(2)
        
        # Calcular causas de rechazo de forma dinámica
        causas_list = []
        df_atd_rech = df_atd_filtered[df_atd_filtered["Estatus_Calidad"] == "Rechazado"]
        df_params = st.session_state.BD_Parametros
        
        for _, row in df_atd_rech.iterrows():
            def_vis = str(row.get("Defectos_Visuales", "Ninguno"))
            if def_vis != "Ninguno" and def_vis != "nan" and def_vis.strip() != "":
                for d in def_vis.split(","):
                    causas_list.append(d.strip())
                    
            sku = row.get("SKU")
            sku_match = df_params[df_params["SKU"] == sku]
            if not sku_match.empty:
                sku_info = sku_match.iloc[0].to_dict()
                
                esp_nom = float(sku_info.get("Espesor_Nominal_in", 0))
                esp_min = esp_nom + float(sku_info.get("Espesor_Tolerancia_Min_in", 0))
                esp_max = esp_nom + float(sku_info.get("Espesor_Tolerancia_Max_in", 0))
                
                m1 = float(row.get("Espesor_Medido_1_in", esp_nom))
                m2 = float(row.get("Espesor_Medido_2_in", esp_nom))
                m3 = float(row.get("Espesor_Medido_3_in", esp_nom))
                
                if not (esp_min <= m1 <= esp_max) or not (esp_min <= m2 <= esp_max) or not (esp_min <= m3 <= esp_max):
                    causas_list.append("Espesor fuera de tolerancia")
                
                anc_nom = float(sku_info.get("Ancho_Nominal_in", 0))
                anc_min = anc_nom + float(sku_info.get("Ancho_Tolerancia_Min_in", 0))
                anc_max = anc_nom + float(sku_info.get("Ancho_Tolerancia_Max_in", 0))
                ancho = float(row.get("Ancho_Medido_in", anc_nom))
                if not (anc_min <= ancho <= anc_max):
                    causas_list.append("Ancho fuera de tolerancia")
                    
                lrg_nom = float(sku_info.get("Largo_Nominal_in", 0))
                lrg_min = lrg_nom + float(sku_info.get("Largo_Tolerancia_Min_in", 0))
                lrg_max = lrg_nom + float(sku_info.get("Largo_Tolerancia_Max_in", 0))
                largo = float(row.get("Largo_Medido_in", lrg_nom))
                if not (lrg_min <= largo <= lrg_max):
                    causas_list.append("Largo fuera de tolerancia")
                    
                dureza_max = float(sku_info.get("Dureza_Max_HRB", 90))
                dureza = float(row.get("Dureza_Medida_HRB", 0))
                if dureza > dureza_max:
                    causas_list.append("Dureza excede límite")
                    
                peso_kg = float(row.get("Peso_Total_Kg", 0))
                if peso_kg > 2500:
                    causas_list.append("Exceso de peso (> 2,500 Kg)")
                    
                zinc_min = float(sku_info.get("Zinc_Min_oz_ft2", 0))
                zinc_val = float(row.get("Zinc_Medido_oz_ft2", 0))
                if sku_info.get("Tipo_Lamina") == "Galvanizada" and zinc_val < zinc_min:
                    causas_list.append("Zinc inferior al límite")
                    
                aceitado_val = str(row.get("Aceitado_OK", "N/A"))
                if sku_info.get("Tipo_Lamina") == "Decapada" and sku_info.get("Aceitado_Requerido") == "Sí" and aceitado_val.upper() != "SÍ":
                    causas_list.append("Falla de Aceitado")
                    
        with col_ca1:
            st.subheader("📋 Estado de Calidad en Atados Recibidos")
            if not df_atd_filtered.empty:
                fig_pie_atd = px.pie(
                    df_atd_filtered, 
                    names="Estatus_Calidad", 
                    color="Estatus_Calidad",
                    color_discrete_map={"Aceptado": "#2E7D32", "Rechazado": "#C62828"},
                    hole=0.4
                )
                st.plotly_chart(fig_pie_atd, use_container_width=True)
            else:
                st.info("Sin datos de atados para mostrar.")
                
        with col_ca2:
            st.subheader("⚠️ Motivos de Rechazo y Defectos Visuales")
            if causas_list:
                df_causas = pd.DataFrame(causas_list, columns=["Causa"])
                df_causas_grouped = df_causas.groupby("Causa").size().reset_index(name="Frecuencia")
                df_causas_grouped = df_causas_grouped.sort_values("Frecuencia", ascending=True)
                
                fig_causas = px.bar(
                    df_causas_grouped, 
                    y="Causa", 
                    x="Frecuencia", 
                    orientation="h",
                    color="Frecuencia",
                    color_continuous_scale=px.colors.sequential.OrRd,
                    labels={"Frecuencia": "Cantidad de Atados", "Causa": "Motivo/Defecto"}
                )
                fig_causas.update_layout(showlegend=False, coloraxis_showscale=False, margin=dict(l=20, r=20, t=10, b=20))
                st.plotly_chart(fig_causas, use_container_width=True)
            else:
                st.success("🎉 **100% de Conformidad:** No se registran atados rechazados ni desviaciones en los filtros seleccionados.")
                
        st.write("---")
        st.subheader("📋 Listado de Recepciones Recientes")
        
        # Calcular tasa de aceptación de atados por folio
        dict_acep = {}
        if not st.session_state.BD_Atados.empty:
            grouped = st.session_state.BD_Atados.groupby("Folio")
            for folio, group in grouped:
                total_a = len(group)
                aceptados_a = len(group[group["Estatus_Calidad"] == "Aceptado"])
                pct_a = (aceptados_a / total_a * 100) if total_a > 0 else 0.0
                dict_acep[folio] = f"{aceptados_a}/{total_a} ({pct_a:.1f}%)"
                
        df_rep_display = df_rep_filtered.copy()
        df_rep_display["Aceptación Atados"] = df_rep_display["Folio"].map(dict_acep).fillna("0/0 (0.0%)")
        df_rep_display = df_rep_display.drop(columns=["Dossier_Ruta", "Fecha_DT"], errors="ignore")
        
        # Reordenar columnas para una vista limpia y profesional
        cols_order = ["Folio", "Fecha", "Hora", "Proveedor", "Orden_Compra", "Factura_Remision", "Aceptación Atados", "Estatus_General", "Inspector"]
        cols_order = [c for c in cols_order if c in df_rep_display.columns]
        df_rep_display = df_rep_display[cols_order]
        
        st.dataframe(df_rep_display.sort_values("Folio", ascending=False), use_container_width=True, hide_index=True)
        
        # --- SECCIÓN DE EXPORTACIÓN A PDF DEL DASHBOARD ---
        st.write("---")
        st.subheader("📥 Exportar Reporte Ejecutivo")
        st.markdown("Descargue el informe en PDF con los OKRs, resultados de cumplimiento y el listado de recepciones filtradas en este periodo.")
        
        # Recopilar datos de filtros para el PDF
        filtros_pdf = {
            "fecha_inicio": fecha_inicio.strftime("%d/%m/%Y"),
            "fecha_fin": fecha_fin.strftime("%d/%m/%Y"),
            "proveedor": prov_sel,
            "material": mat_sel,
            "calibre": cal_sel
        }
        
        # Calcular variables de OKRs para pasar al PDF
        tasa_atados_pdf = (atados_aceptados / total_atados * 100) if total_atados > 0 else 100.0
        tasa_lotes_pdf = (lotes_aceptados / total_lotes * 100) if total_lotes > 0 else 100.0
        
        peso_aceptado_kg = df_atd_filtered[df_atd_filtered["Estatus_Calidad"] == "Aceptado"]["Peso_Total_Kg"].sum()
        tasa_peso_pdf = (peso_aceptado_kg / peso_total_kg * 100) if peso_total_kg > 0 else 100.0
        
        okr_data_pdf = {
            "tasa_atados": tasa_atados_pdf,
            "tasa_lotes": tasa_lotes_pdf,
            "tasa_peso": tasa_peso_pdf,
            "total_atados": total_atados,
            "peso_total_kg": peso_total_kg,
            "peso_total_lb": peso_total_lb,
            "total_lotes": total_lotes,
            "lotes_aceptados": lotes_aceptados,
            "lotes_rechazados": lotes_rechazados
        }
        
        try:
            # Crear carpeta temporal de descargas si no existe
            temp_pdf_dir = os.path.join(BASE_DIR, "carpetas_electronicas", "temp_descargas")
            os.makedirs(temp_pdf_dir, exist_ok=True)
            pdf_path_temp = os.path.join(temp_pdf_dir, f"Reporte_Dashboard_{datetime.date.today().strftime('%Y%m%d')}.pdf")
            
            # Generar el PDF
            utils_pdf.generar_pdf_reporte_dashboard(filtros_pdf, okr_data_pdf, df_rep_filtered, dict_acep, pdf_path_temp)
            
            if os.path.exists(pdf_path_temp):
                with open(pdf_path_temp, "rb") as f:
                    pdf_bytes = f.read()
                    
                st.download_button(
                    label="📥 Descargar Reporte Ejecutivo del Dashboard (PDF)",
                    data=pdf_bytes,
                    file_name=f"Reporte_Ejecutivo_Dashboard_{datetime.date.today().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key="btn_descarga_dashboard_pdf"
                )
                
                # Limpiar archivo temporal
                try:
                    os.remove(pdf_path_temp)
                except Exception:
                    pass
            else:
                st.error("No se pudo generar el archivo de reporte PDF.")
        except Exception as e:
            st.error(f"Error al generar reporte PDF: {e}")


# =============================================================================
# MÓDULO 2: REGISTRO DE RECEPCIÓN (INCOMING)
# =============================================================================
elif opcion_menu == "📥 Registro de Recepción (Incoming)":
    st.title("📥 Registro de Control de Calidad en Recepción")
    if not is_inspector:
        st.error("🔒 Área Protegida. Ingrese la contraseña de Inspector o Administrador en la barra lateral para registrar recepciones.")
        st.stop()
    st.markdown("Suba la plantilla Excel con las mediciones y los documentos adjuntos (Certificados, OC) para generar el Dosier de Calidad.")
    
    # Descargar plantilla corporativa
    if os.path.exists(PLANTILLA_PATH):
        with open(PLANTILLA_PATH, "rb") as f:
            st.download_button(
                label="📥 Descargar Formato de Plantilla Corporativa (.xlsx)",
                data=f.read(),
                file_name=f"FO-MET-30-Plantilla_Atados ({datetime.date.today().strftime('%Y%m%d')}).xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    st.write("---")
    
    # Formulario de datos generales
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        proveedor = st.text_input("🏢 Proveedor:", value="Ternium México", key="reg_proveedor")
        orden_compra = st.text_input("📋 Orden de Compra Sigrama (PO):", value="0301682984", key="reg_orden_compra")
        factura_remision = st.text_input("📄 Factura o Remisión del Proveedor:", value="TX6793212", key="reg_factura_remision")
        
    with col_f2:
        inspector = st.text_input("🔍 Inspector de Calidad:", value="Jesus Morales", key="reg_inspector")
        fecha_recepcion = st.date_input("📅 Fecha de Recepción:", datetime.date.today(), key="reg_fecha_recepcion")
        # Hora automática
        hora_recepcion = datetime.datetime.now().strftime("%I:%M %p")
        st.text_input("⏰ Hora de Ingreso:", value=hora_recepcion, disabled=True)
        
    st.write("### 📎 Carga de Documentación Física y Fotos")
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        cert_file = st.file_uploader("📑 Certificado de Calidad del Proveedor (PDF)", type=["pdf"], key="cert_file")
    with col_u2:
        oc_file = st.file_uploader("🛍️ Orden de Compra de Sigrama (PDF)", type=["pdf"], key="oc_file")
        
    st.write("### 🧮 Mediciones Dimensionales de Campo")
    mediciones_file = st.file_uploader("📥 Suba la plantilla Excel con mediciones completadas:", type=["xlsx"])
    
    if mediciones_file:
        try:
            df_med = pd.read_excel(mediciones_file, sheet_name=0)
            
            # Limpiar filas vacías
            df_med = df_med.dropna(subset=["No_Atado", "Calibre"])
            
            st.write("🔍 **Pre-visualización de Datos Cargados:**")
            st.dataframe(df_med, use_container_width=True)
            
            # --- SECCIÓN DE INSPECCIÓN VISUAL Y DE APARIENCIA POR ROLLO ---
            st.write("---")
            st.write("### 🔍 Inspección Visual y de Apariencia por Rollo")
            st.markdown("Confirme que cada rollo cumpla con los criterios visuales de apariencia del SGC. Las casillas están marcadas como **CUMPLE** por defecto.")
            
            # Asegurar columna Placa
            if "Placa" not in df_med.columns:
                df_med["Placa"] = 1
                
            unique_atados = df_med["No_Atado"].unique()
            
            for id_atd_prov in unique_atados:
                df_atd_rows = df_med[df_med["No_Atado"] == id_atd_prov]
                first_row = df_atd_rows.iloc[0]
                calibre_raw = first_row["Calibre"]
                tipo_raw = str(first_row["Galvanizado_o_Decapado"]).strip()
                
                # Obtener calibre numérico usando regex
                import re
                match_c = re.search(r'\d+', str(calibre_raw))
                if match_c:
                    calibre_num = int(match_c.group())
                else:
                    calibre_num = 14
                    
                # Determinar tipo de lámina
                tipo_lamina = "Galvanizada"
                if "decap" in tipo_raw.lower() or "dec" in tipo_raw.lower():
                    tipo_lamina = "Decapada"
                elif "alum" in tipo_raw.lower():
                    tipo_lamina = "Aluminio"
                    
                tipo_key = "GALV"
                if tipo_lamina == "Decapada":
                    tipo_key = "DECP"
                elif tipo_lamina == "Aluminio":
                    tipo_key = "ALUM"
                sku_atd = f"SKU-{tipo_key}-{calibre_num}"
                
                df_skus = st.session_state.BD_Parametros
                sku_match = df_skus[df_skus["SKU"] == sku_atd]
                nombre_sku = ""
                if not sku_match.empty:
                    tipo_lamina = sku_match.iloc[0]["Tipo_Lamina"]
                    nombre_sku = sku_match.iloc[0]["Nombre"]
                    
                calibre_val = f"CAL {calibre_num}"
                
                with st.expander(f"📋 Inspección de Apariencia: {id_atd_prov} ({tipo_lamina} - {calibre_val})", expanded=True):
                    # Dividir la cabecera en título y botón "Procesar y guardar"
                    col_title, col_btn = st.columns([3, 1])
                    with col_title:
                        st.markdown(f"<div style='font-size: 2.2rem; font-weight: bold; color: #0d47a1; margin-top: 0px;'>📋 Rollo / Atado: <span style='color: #1a73e8;'>{id_atd_prov}</span></div>", unsafe_allow_html=True)
                    with col_btn:
                        btn_guardar = st.button("💾 Procesar y Guardar", key=f"btn_save_{id_atd_prov}", use_container_width=True)
                        
                    # Inicializar almacenamiento temporal de guardados si no existe
                    if "atados_guardados" not in st.session_state:
                        st.session_state.atados_guardados = {}
                        
                    if btn_guardar:
                        # Recuperar valores live del UI
                        chk_cert_val = st.session_state.get(f"chk_cert_{id_atd_prov}", True)
                        chk_apariencia_val = st.session_state.get(f"chk_apariencia_{id_atd_prov}", True)
                        chk_empaque_val = st.session_state.get(f"chk_empaque_{id_atd_prov}", True)
                        chk_formado_val = st.session_state.get(f"chk_formado_{id_atd_prov}", True)
                        chk_rayas_val = st.session_state.get(f"chk_rayas_{id_atd_prov}", True)
                        chk_puntos_val = st.session_state.get(f"chk_puntos_{id_atd_prov}", True)
                        chk_blancas_val = st.session_state.get(f"chk_blancas_{id_atd_prov}", True)
                        chk_suciedad_val = st.session_state.get(f"chk_suciedad_{id_atd_prov}", True)
                        chk_rodillos_val = st.session_state.get(f"chk_rodillos_{id_atd_prov}", True)
                        chk_escalones_val = st.session_state.get(f"chk_escalones_{id_atd_prov}", True)
                        chk_poros_val = st.session_state.get(f"chk_poros_{id_atd_prov}", True)
                        chk_zinc_val = st.session_state.get(f"chk_zinc_{id_atd_prov}", True)
                        chk_aceitado_val = st.session_state.get(f"chk_aceitado_{id_atd_prov}", True)
                        chk_dureza_val = st.session_state.get(f"chk_dureza_{id_atd_prov}", True)
                        
                        # Evaluar criterios de apariencia comunes a todo el atado
                        motivos = []
                        if not chk_cert_val: motivos.append("Falta Certificado de Calidad")
                        if not chk_apariencia_val: motivos.append("Falla en apariencia visual")
                        if not chk_empaque_val: motivos.append("Empaque de papel VCI no conforme")
                        if not chk_formado_val: motivos.append("Marcas de línea de formado o golpes")
                        if not chk_rayas_val: motivos.append("Presencia de rayas o manchas")
                        if not chk_puntos_val: motivos.append("Presencia de puntos negros")
                        if not chk_blancas_val: motivos.append("Presencia de manchas blancas")
                        if not chk_suciedad_val: motivos.append("Suciedad o dobleces")
                        if not chk_rodillos_val: motivos.append("Marcas de rodillos")
                        if not chk_escalones_val: motivos.append("Escalones en lámina")
                        if not chk_poros_val: motivos.append("Poros en la lámina")
                        
                        if tipo_lamina == "Galvanizada" and not chk_zinc_val: motivos.append("Recubrimiento de Zinc inferior al límite")
                        if tipo_lamina == "Decapada" and not chk_aceitado_val: motivos.append("Aceitado ausente")
                        if not chk_dureza_val: motivos.append("Dureza excede el límite máximo")
                        
                        # Evaluar mediciones físicas de cada una de las placas/hojas registradas para este atado
                        for idx_sub, row_med in df_atd_rows.iterrows():
                            p_val = row_med.get("Placa", idx_sub+1)
                            try:
                                p_val = int(float(p_val))
                            except Exception:
                                pass
                            placa_lbl = f"Placa {p_val}"
                            
                            m1 = float(row_med["Espesor_Medido_1_in"])
                            m2 = float(row_med["Espesor_Medido_2_in"])
                            m3 = float(row_med["Espesor_Medido_3_in"])
                            peso_kg = float(row_med["Peso_Total_Kg"])
                            
                            # Tolerancias nominales
                            esp_nom = float(sku_match.iloc[0]["Espesor_Nominal_in"]) if not sku_match.empty else 0.075
                            esp_min = esp_nom + float(sku_match.iloc[0]["Espesor_Tolerancia_Min_in"]) if not sku_match.empty else 0.067
                            esp_max = esp_nom + float(sku_match.iloc[0]["Espesor_Tolerancia_Max_in"]) if not sku_match.empty else 0.083
                            
                            if not (esp_min <= m1 <= esp_max): motivos.append(f"{placa_lbl}: Espesor 1 ({m1:.4f}\") fuera de especificación")
                            if not (esp_min <= m2 <= esp_max): motivos.append(f"{placa_lbl}: Espesor 2 ({m2:.4f}\") fuera de especificación")
                            if not (esp_min <= m3 <= esp_max): motivos.append(f"{placa_lbl}: Espesor 3 ({m3:.4f}\") fuera de especificación")
                            if peso_kg > 2500: motivos.append(f"{placa_lbl}: Peso ({peso_kg:.0f} Kg) excede 2500 Kg")
                            
                        estatus = "Aceptado" if not motivos else "Rechazado"
                        st.session_state.atados_guardados[id_atd_prov] = {
                            "estatus": estatus,
                            "motivos": motivos,
                            "chk_cert_val": chk_cert_val,
                            "chk_apariencia_val": chk_apariencia_val,
                            "chk_empaque_val": chk_empaque_val,
                            "chk_formado_val": chk_formado_val,
                            "chk_rayas_val": chk_rayas_val,
                            "chk_puntos_val": chk_puntos_val,
                            "chk_blancas_val": chk_blancas_val,
                            "chk_suciedad_val": chk_suciedad_val,
                            "chk_rodillos_val": chk_rodillos_val,
                            "chk_escalones_val": chk_escalones_val,
                            "chk_poros_val": chk_poros_val,
                            "chk_zinc_val": chk_zinc_val,
                            "chk_aceitado_val": chk_aceitado_val,
                            "chk_dureza_val": chk_dureza_val
                        }
                        
                    # Mostrar banner informativo si ya fue guardado
                    info_guardado = st.session_state.atados_guardados.get(id_atd_prov)
                    if info_guardado:
                        estatus_saved = info_guardado["estatus"]
                        motivos_saved = info_guardado["motivos"]
                        if estatus_saved == "Aceptado":
                            st.success(f"💾 **Atado Procesado e Individualizado:** Este atado está **APROBADO** en conformidad con el SGC.")
                        else:
                            st.error(f"⚠️ **Atado Procesado e Individualizado:** Este atado está **RECHAZADO**. Motivos: {', '.join(motivos_saved)}")
                            
                    col_v1, col_v2, col_v3 = st.columns(3)
                    
                    with col_v1:
                        st.markdown("**Especificaciones y Empaque**")
                        st.checkbox("📄 Certificado de Calidad OK (ASTM, SGS, ISO 9001, CE)", value=True, key=f"chk_cert_{id_atd_prov}")
                        st.checkbox("✨ Apariencia General OK (Flor Regular / Decapado Gris)", value=True, key=f"chk_apariencia_{id_atd_prov}")
                        st.checkbox("📦 Empaque Papel VCI OK (Preparación de Atado)", value=True, key=f"chk_empaque_{id_atd_prov}")
                        st.checkbox("🚫 Sin marcas de línea de formado o golpes en lámina", value=True, key=f"chk_formado_{id_atd_prov}")
                        
                    with col_v2:
                        st.markdown("**Criterios de No Aceptación**")
                        st.checkbox("✔️ Sin Rayas o Manchas superficiales", value=True, key=f"chk_rayas_{id_atd_prov}")
                        st.checkbox("✔️ Sin Puntos Negros", value=True, key=f"chk_puntos_{id_atd_prov}")
                        st.checkbox("✔️ Sin Manchas Blancas", value=True, key=f"chk_blancas_{id_atd_prov}")
                        st.checkbox("✔️ Sin Suciedad o Dobleces en lámina", value=True, key=f"chk_suciedad_{id_atd_prov}")
                        
                    with col_v3:
                        st.markdown("**Otros Controles**")
                        st.checkbox("✔️ Sin Marcas de Rodillos", value=True, key=f"chk_rodillos_{id_atd_prov}")
                        st.checkbox("✔️ Sin Escalones en la lámina", value=True, key=f"chk_escalones_{id_atd_prov}")
                        st.checkbox("✔️ Sin Poros en la lámina", value=True, key=f"chk_poros_{id_atd_prov}")
                        
                        if tipo_lamina == "Galvanizada":
                            st.checkbox("⚡ Recubrimiento de Zinc OK (G60/G90)", value=True, key=f"chk_zinc_{id_atd_prov}")
                        if tipo_lamina == "Decapada":
                            st.checkbox("🛢️ Aceitado Conforme (Aceitado OK)", value=True, key=f"chk_aceitado_{id_atd_prov}")
                        st.checkbox("🔨 Dureza Mecánica Conforme (Dureza OK)", value=True, key=f"chk_dureza_{id_atd_prov}")
                        
                    # Carga de Evidencia Fotográfica de Defectos por Atado
                    st.write("---")
                    st.markdown("📷 **Carga de Evidencia Fotográfica**")
                    st.file_uploader(f"📸 Subir Fotos de Defectos Visuales para el Atado {id_atd_prov} (Opcional)", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key=f"defect_photos_{id_atd_prov}")
                        
                    # Pintar análisis gaussiano por cada hoja medida en pestañas (tabs)
                    if not sku_match.empty:
                        sku_info_dict = sku_match.iloc[0].to_dict()
                        st.write("---")
                        st.write("📈 **Análisis Estadístico de Espesor por Hoja (Distribución de Gauss):**")
                        
                        placa_names = []
                        for idx_sub, row_sub in df_atd_rows.iterrows():
                            p_num = row_sub.get("Placa", idx_sub+1)
                            try:
                                p_num = int(float(p_num))
                            except Exception:
                                pass
                            placa_names.append(f"Placa {p_num}")
                            
                        tabs = st.tabs(placa_names)
                        for tab, (idx_sub, row_sub) in zip(tabs, df_atd_rows.iterrows()):
                            with tab:
                                p_num = row_sub.get("Placa", idx_sub+1)
                                try:
                                    p_num = int(float(p_num))
                                except Exception:
                                    pass
                                row_compatible = row_sub.to_dict()
                                row_compatible['ID_Atado_Proveedor'] = f"{id_atd_prov} (Placa {p_num})"
                                row_compatible['Tipo_Lamina'] = tipo_lamina
                                
                                # Pre-evaluar estatus para visualización
                                m1 = float(row_compatible.get("Espesor_Medido_1_in", 0))
                                m2 = float(row_compatible.get("Espesor_Medido_2_in", 0))
                                m3 = float(row_compatible.get("Espesor_Medido_3_in", 0))
                                esp_nom = float(sku_info_dict.get("Espesor_Nominal_in", 0))
                                esp_min = esp_nom + float(sku_info_dict.get("Espesor_Tolerancia_Min_in", -0.008))
                                esp_max = esp_nom + float(sku_info_dict.get("Espesor_Tolerancia_Max_in", 0.008))
                                peso_kg = float(row_compatible.get("Peso_Total_Kg", 0))
                                
                                is_ok = (esp_min <= m1 <= esp_max) and (esp_min <= m2 <= esp_max) and (esp_min <= m3 <= esp_max) and (peso_kg <= 2500)
                                row_compatible['Estatus_Calidad'] = 'Aceptado' if is_ok else 'Rechazado'
                                
                                renderizar_analisis_gaussiano_atado(row_compatible, sku_info_dict, key=f"plotly_gauss_reg_{id_atd_prov}_P{p_num}")
            
            st.write("---")
            
            # Botón para procesar
            if st.button("💾 Salvar Todo: Procesar e Integrar Recepción a Base de Datos"):
                # Leer valores de st.session_state para evitar pérdidas en reruns
                proveedor_val = st.session_state.get("reg_proveedor", "").strip()
                orden_compra_val = st.session_state.get("reg_orden_compra", "").strip()
                factura_remision_val = st.session_state.get("reg_factura_remision", "").strip()
                inspector_val = st.session_state.get("reg_inspector", "").strip()
                fecha_recepcion_val = st.session_state.get("reg_fecha_recepcion", datetime.date.today())
                
                if not proveedor_val or not orden_compra_val or not factura_remision_val or not inspector_val:
                    st.error("❌ Error: Todos los campos generales (Proveedor, Orden de Compra, Factura y Nombre del Inspector) son obligatorios.")
                elif not cert_file or not oc_file:
                    st.error("❌ Error: Es obligatorio subir tanto el **Certificado de Molino** como la **Orden de Compra de Sigrama** para conformar el Dosier de Calidad.")
                else:
                    # Mapear variables para el resto de la ejecución
                    proveedor = proveedor_val
                    orden_compra = orden_compra_val
                    factura_remision = factura_remision_val
                    inspector = inspector_val
                    fecha_recepcion = fecha_recepcion_val
                    with st.spinner("Procesando mediciones, validando tolerancias y generando PDFs..."):
                        # 1. Determinar el siguiente Folio
                        df_reportes_master = st.session_state.BD_Reportes
                        if df_reportes_master.empty:
                            nuevo_folio = "INC-2026-0001"
                        else:
                            # Obtener el último número secuencial
                            ult_folio = df_reportes_master.sort_values("Folio")["Folio"].iloc[-1]
                            try:
                                seq = int(ult_folio.split("-")[2]) + 1
                                nuevo_folio = f"INC-2026-{seq:04d}"
                            except Exception:
                                nuevo_folio = f"INC-2026-{len(df_reportes_master)+1:04d}"
                                
                        # Crear carpeta electrónica para este folio
                        folio_folder = os.path.join(CARPETAS_DIR, nuevo_folio)
                        os.makedirs(folio_folder, exist_ok=True)
                        
                        # 2. Guardar archivos de soporte (Certificado y OC)
                        cert_path = os.path.join(folio_folder, f"Certificado_Proveedor_{nuevo_folio}.pdf")
                        oc_path = os.path.join(folio_folder, f"Orden_Compra_Sigrama_{nuevo_folio}.pdf")
                        
                        with open(cert_path, "wb") as f:
                            f.write(cert_file.read())
                        with open(oc_path, "wb") as f:
                            f.write(oc_file.read())
                            
                        # Guardar fotos de defectos por atado (agrupado por No_Atado)
                        for id_atd_prov in df_med["No_Atado"].unique():
                            atd_photos = st.session_state.get(f"defect_photos_{id_atd_prov}")
                            if atd_photos:
                                for idx_ph, ph in enumerate(atd_photos):
                                    ph_ext = os.path.splitext(ph.name)[1]
                                    ph_path = os.path.join(folio_folder, f"Foto_Defecto_{id_atd_prov}_{idx_ph+1}_{nuevo_folio}{ph_ext}")
                                    with open(ph_path, "wb") as f:
                                        f.write(ph.read())
                                    
                        # 3. Procesar atados y evaluar tolerancias
                        df_atados_lote = []
                        lote_rechazado_dimensional = False
                        lote_rechazado_visual = False
                        
                        # Obtener parámetros del primer SKU para tener un SKU principal por defecto
                        first_row = df_med.iloc[0]
                        first_calibre = first_row["Calibre"]
                        first_tipo_raw = str(first_row["Galvanizado_o_Decapado"]).strip()
                        import re
                        match_f = re.search(r'\d+', str(first_calibre))
                        if match_f:
                            f_calibre_num = int(match_f.group())
                        else:
                            f_calibre_num = 14
                        f_tipo_lamina = "Galvanizada"
                        if "decap" in first_tipo_raw.lower() or "dec" in first_tipo_raw.lower():
                            f_tipo_lamina = "Decapada"
                        elif "alum" in first_tipo_raw.lower():
                            f_tipo_lamina = "Aluminio"
                        f_tipo_key = "GALV"
                        if f_tipo_lamina == "Decapada":
                            f_tipo_key = "DECP"
                        elif f_tipo_lamina == "Aluminio":
                            f_tipo_key = "ALUM"
                        sku_principal = f"SKU-{f_tipo_key}-{f_calibre_num}"
                        
                        df_skus = st.session_state.BD_Parametros
                        sku_match_principal = df_skus[df_skus["SKU"] == sku_principal]
                        if sku_match_principal.empty:
                            st.error(f"❌ El SKU principal detectado '{sku_principal}' no está registrado en el Catálogo de Tolerancias de la empresa.")
                            st.stop()
                        sku_info = sku_match_principal.iloc[0].to_dict()
                        
                        for idx_atd, row_med in df_med.iterrows():
                            id_atd_int = f"{nuevo_folio}-A{idx_atd+1:02d}"
                            id_atd_prov = str(row_med["No_Atado"])
                            
                            # Reconstruir SKU específico para este atado
                            row_calibre = row_med["Calibre"]
                            row_tipo_raw = str(row_med["Galvanizado_o_Decapado"]).strip()
                            match_r = re.search(r'\d+', str(row_calibre))
                            if match_r:
                                r_calibre_num = int(match_r.group())
                            else:
                                r_calibre_num = 14
                            r_tipo_lamina = "Galvanizada"
                            if "decap" in row_tipo_raw.lower() or "dec" in row_tipo_raw.lower():
                                r_tipo_lamina = "Decapada"
                            elif "alum" in row_tipo_raw.lower():
                                r_tipo_lamina = "Aluminio"
                            r_tipo_key = "GALV"
                            if r_tipo_lamina == "Decapada":
                                r_tipo_key = "DECP"
                            elif r_tipo_lamina == "Aluminio":
                                r_tipo_key = "ALUM"
                            sku_atado = f"SKU-{r_tipo_key}-{r_calibre_num}"
                            
                            # Obtener info y tolerancias específicas
                            sku_match = df_skus[df_skus["SKU"] == sku_atado]
                            if sku_match.empty:
                                st.error(f"❌ El SKU '{sku_atado}' derivado para el atado {id_atd_prov} no está registrado en el catálogo.")
                                st.stop()
                            sku_info = sku_match.iloc[0].to_dict()
                            grado_acero_val = sku_info.get("Grado_Acero", "ASTM A653 CS Tipo B")
                            
                            # Límites de tolerancia
                            esp_nom = float(sku_info.get("Espesor_Nominal_in", 0))
                            esp_min = esp_nom + float(sku_info.get("Espesor_Tolerancia_Min_in", 0))
                            esp_max = esp_nom + float(sku_info.get("Espesor_Tolerancia_Max_in", 0))
                            
                            anc_nom = float(sku_info.get("Ancho_Nominal_in", 0))
                            anc_min = anc_nom + float(sku_info.get("Ancho_Tolerancia_Min_in", 0))
                            anc_max = anc_nom + float(sku_info.get("Ancho_Tolerancia_Max_in", 0))
                            
                            lrg_nom = float(sku_info.get("Largo_Nominal_in", 0))
                            lrg_min = lrg_nom + float(sku_info.get("Largo_Tolerancia_Min_in", 0))
                            lrg_max = lrg_nom + float(sku_info.get("Largo_Tolerancia_Max_in", 0))
                            
                            zinc_min = float(sku_info.get("Zinc_Min_oz_ft2", 0))
                            dureza_max = float(sku_info.get("Dureza_Max_HRB", 90))
                            
                            # Recuperar valores de UI con prioridad a los guardados individualmente
                            saved_info = st.session_state.get("atados_guardados", {}).get(id_atd_prov)
                            if saved_info:
                                chk_cert_val = saved_info.get("chk_cert_val", True)
                                chk_apariencia_val = saved_info.get("chk_apariencia_val", True)
                                chk_empaque_val = saved_info.get("chk_empaque_val", True)
                                chk_formado_val = saved_info.get("chk_formado_val", True)
                                chk_rayas_val = saved_info.get("chk_rayas_val", True)
                                chk_puntos_val = saved_info.get("chk_puntos_val", True)
                                chk_blancas_val = saved_info.get("chk_blancas_val", True)
                                chk_suciedad_val = saved_info.get("chk_suciedad_val", True)
                                chk_rodillos_val = saved_info.get("chk_rodillos_val", True)
                                chk_escalones_val = saved_info.get("chk_escalones_val", True)
                                chk_poros_val = saved_info.get("chk_poros_val", True)
                                chk_zinc_val = saved_info.get("chk_zinc_val", True)
                                chk_aceitado_val = saved_info.get("chk_aceitado_val", True)
                                chk_dureza_val = saved_info.get("chk_dureza_val", True)
                            else:
                                chk_cert_val = st.session_state.get(f"chk_cert_{id_atd_prov}", True)
                                chk_apariencia_val = st.session_state.get(f"chk_apariencia_{id_atd_prov}", True)
                                chk_empaque_val = st.session_state.get(f"chk_empaque_{id_atd_prov}", True)
                                chk_formado_val = st.session_state.get(f"chk_formado_{id_atd_prov}", True)
                                chk_rayas_val = st.session_state.get(f"chk_rayas_{id_atd_prov}", True)
                                chk_puntos_val = st.session_state.get(f"chk_puntos_{id_atd_prov}", True)
                                chk_blancas_val = st.session_state.get(f"chk_blancas_{id_atd_prov}", True)
                                chk_suciedad_val = st.session_state.get(f"chk_suciedad_{id_atd_prov}", True)
                                chk_rodillos_val = st.session_state.get(f"chk_rodillos_{id_atd_prov}", True)
                                chk_escalones_val = st.session_state.get(f"chk_escalones_{id_atd_prov}", True)
                                chk_poros_val = st.session_state.get(f"chk_poros_{id_atd_prov}", True)
                                chk_zinc_val = st.session_state.get(f"chk_zinc_{id_atd_prov}", True)
                                chk_aceitado_val = st.session_state.get(f"chk_aceitado_{id_atd_prov}", True)
                                chk_dureza_val = st.session_state.get(f"chk_dureza_{id_atd_prov}", True)
                            
                            # Validar mediciones
                            m1 = float(row_med["Espesor_Medido_1_in"])
                            m2 = float(row_med["Espesor_Medido_2_in"])
                            m3 = float(row_med["Espesor_Medido_3_in"])
                            ancho = float(row_med.get("Ancho_Medido_in", anc_nom))
                            largo = float(row_med.get("Largo_Medido_in", lrg_nom))
                            cant_hojas = int(row_med.get("Cantidad_Hojas", 0))
                            peso_kg = float(row_med["Peso_Total_Kg"])
                            peso_lb = round(peso_kg * 2.20462, 1)
                            
                            # Mapear valores medidos virtuales
                            if sku_info.get("Tipo_Lamina") == "Galvanizada":
                                zinc_val = float(sku_info.get("Zinc_Nominal_oz_ft2", 0.6)) if chk_zinc_val else (zinc_min - 0.1)
                            else:
                                zinc_val = 0.0
                                
                            dureza = (float(sku_info.get("Dureza_Max_HRB", 75)) - 5) if chk_dureza_val else (dureza_max + 5)
                            
                            if sku_info.get("Tipo_Lamina") == "Galvanizada":
                                aceitado_val = "N/A"
                            else:
                                aceitado_val = "Sí" if chk_aceitado_val else "No"
                                
                            # Evaluación de motivos de rechazo y recolección de defectos visuales
                            motivos_rechazo = []
                            visual_defects = []
                            
                            if not chk_cert_val:
                                visual_defects.append("Falta Certificado de Calidad")
                                motivos_rechazo.append("Certificado de calidad no conforme o ausente")
                            if not chk_apariencia_val:
                                visual_defects.append("Falla Apariencia General")
                                motivos_rechazo.append("Falla en apariencia visual (flor/decapado)")
                            if not chk_empaque_val:
                                visual_defects.append("Empaque Defectuoso")
                                motivos_rechazo.append("Empaque de papel VCI no conforme")
                            if not chk_formado_val:
                                visual_defects.append("Marcas de Formado/Golpes")
                                motivos_rechazo.append("Marcas de línea de formado o golpes en lámina")
                            if not chk_rayas_val:
                                visual_defects.append("Rayas/Manchas")
                                motivos_rechazo.append("Presencia de rayas o manchas")
                            if not chk_puntos_val:
                                visual_defects.append("Puntos Negros")
                                motivos_rechazo.append("Presencia de puntos negros")
                            if not chk_blancas_val:
                                visual_defects.append("Manchas Blancas")
                                motivos_rechazo.append("Presencia de manchas blancas (óxido blanco)")
                            if not chk_suciedad_val:
                                visual_defects.append("Suciedad/Dobleces")
                                motivos_rechazo.append("Suciedad o dobleces en lámina")
                            if not chk_rodillos_val:
                                visual_defects.append("Marcas de Rodillos")
                                motivos_rechazo.append("Presencia de marcas de rodillos")
                            if not chk_escalones_val:
                                visual_defects.append("Escalones en lámina")
                                motivos_rechazo.append("Escalones en lámina")
                            if not chk_poros_val:
                                visual_defects.append("Poros en la lámina")
                                motivos_rechazo.append("Presencia de poros en la lámina")
                                
                            # Zinc, dureza, aceitado
                            if sku_info.get("Tipo_Lamina") == "Galvanizada" and not chk_zinc_val:
                                motivos_rechazo.append("Recubrimiento de Zinc inferior al límite")
                            if sku_info.get("Tipo_Lamina") == "Decapada" and sku_info.get("Aceitado_Requerido") == "Sí" and not chk_aceitado_val:
                                motivos_rechazo.append("Aceitado ausente")
                            if not chk_dureza_val:
                                motivos_rechazo.append("Dureza excede el límite máximo")
                                
                            # Mediciones dimensionales
                            es_dimensional = False
                            if not (esp_min <= m1 <= esp_max): 
                                motivos_rechazo.append("Espesor 1 fuera de especificación")
                                es_dimensional = True
                            if not (esp_min <= m2 <= esp_max): 
                                motivos_rechazo.append("Espesor 2 fuera de especificación")
                                es_dimensional = True
                            if not (esp_min <= m3 <= esp_max): 
                                motivos_rechazo.append("Espesor 3 fuera de especificación")
                                es_dimensional = True
                            if not (anc_min <= ancho <= anc_max): 
                                motivos_rechazo.append("Ancho fuera de especificación")
                                es_dimensional = True
                            if not (lrg_min <= largo <= lrg_max): 
                                motivos_rechazo.append("Largo fuera de especificación")
                                es_dimensional = True
                            if peso_kg > 2500:
                                motivos_rechazo.append("Peso del atado excede el límite de 2.5 Toneladas (2500 Kg) según RFQ")
                                es_dimensional = True
                                
                            defects_str = ", ".join(visual_defects) if visual_defects else "Ninguno"
                            
                            atd_status = "Aceptado"
                            if motivos_rechazo:
                                atd_status = "Rechazado"
                                if es_dimensional:
                                    lote_rechazado_dimensional = True
                                else:
                                    lote_rechazado_visual = True
                                
                            # Obtener valor de Placa
                            placa_val = row_med.get("Placa") or row_med.get("Hoja") or row_med.get("No_Hoja") or 1
                            try:
                                placa_val = int(float(placa_val))
                            except Exception:
                                pass

                            atado_record = {
                                "ID_Atado": id_atd_int,
                                "Folio": nuevo_folio,
                                "ID_Atado_Proveedor": id_atd_prov,
                                "Placa": placa_val,
                                "SKU": sku_atado,
                                "Grado_Acero": grado_acero_val,
                                "Num_Colada": str(row_med["Num_Colada"]),
                                "Lote_Heat": str(row_med["Lote_Heat"]),
                                "Espesor_Medido_1_in": m1,
                                "Espesor_Medido_2_in": m2,
                                "Espesor_Medido_3_in": m3,
                                "Ancho_Medido_in": ancho,
                                "Largo_Medido_in": largo,
                                "Cantidad_Hojas": cant_hojas,
                                "Peso_Total_Kg": peso_kg,
                                "Peso_Total_Lb": peso_lb,
                                "Zinc_Medido_oz_ft2": zinc_val,
                                "Dureza_Medida_HRB": dureza,
                                "Aceitado_OK": aceitado_val,
                                "Defectos_Visuales": defects_str,
                                "Ubicacion_Almacen": "Almacén Metales",
                                "Estatus_Calidad": atd_status,
                                "Observaciones": str(row_med.get("Observaciones", "")),
                                "Proveedor": proveedor,
                                "Tipo_Lamina": sku_info.get("Tipo_Lamina"),
                                "Acabado": sku_info.get("Acabado_Superficial", "Liso"),
                                "Espesor_Nominal": esp_nom,
                                "Espesor_Tol_Min": sku_info.get("Espesor_Tolerancia_Min_in"),
                                "Espesor_Tol_Max": sku_info.get("Espesor_Tolerancia_Max_in"),
                                "Ancho_Tol_Min_Val": anc_min,
                                "Ancho_Tol_Max_Val": anc_max
                            }
                            df_atados_lote.append(atado_record)
                            
                        # Crear dataframe temporal
                        df_atados_temp = pd.DataFrame(df_atados_lote)
                        
                        # Generar gráficas de curvas de tolerancia
                        img_curvas_path = os.path.join(folio_folder, f"Curvas_Tolerancia_{nuevo_folio}.png")
                        utils_pdf.generar_graficas_tolerancia(df_atados_temp, sku_info, img_curvas_path)
                        
                        # Estatus general del lote
                        estatus_general = "Aceptado"
                        if lote_rechazado_dimensional:
                            estatus_general = "Rechazado"
                        elif lote_rechazado_visual:
                            estatus_general = "Condicionado"
                            
                        # Generar archivos de reportes PDF individuales
                        pdf_reporte = os.path.join(folio_folder, f"Reporte_FO-MET-31_{nuevo_folio}.pdf")
                        pdf_tecnico = os.path.join(folio_folder, f"Reporte_Tecnico_{nuevo_folio}.pdf")
                        pdf_etiquetas = os.path.join(folio_folder, f"Etiquetas_FO-MET-32_{nuevo_folio}.pdf")
                        pdf_solo_etiquetas = os.path.join(folio_folder, f"Etiquetas_Solo_FO-MET-32_{nuevo_folio}.pdf")
                        pdf_portada = os.path.join(folio_folder, f"Portada_FO-MET-33_{nuevo_folio}.pdf")
                        
                        datos_reporte = {
                            "Fecha": fecha_recepcion.strftime("%d/%m/%Y"),
                            "Proveedor": proveedor,
                            "Orden_Compra": orden_compra,
                            "Factura_Remision": factura_remision,
                            "Inspector": inspector,
                            "Estatus_General": estatus_general,
                            "Hora": hora_recepcion
                        }
                        
                        utils_pdf.generar_pdf_reporte_consolidado_fomet31(nuevo_folio, datos_reporte, df_atados_temp, sku_info, img_curvas_path, pdf_reporte)
                        utils_pdf.generar_pdf_reporte_tecnico_consolidado(nuevo_folio, datos_reporte, df_atados_temp, pdf_tecnico)
                        utils_pdf.generar_pdf_etiqueta_atado_fomet32(nuevo_folio, df_atados_temp, pdf_etiquetas)
                        utils_pdf.generar_pdf_solo_etiquetas(nuevo_folio, df_atados_temp, pdf_solo_etiquetas)
                        utils_pdf.generar_pdf_portada_dosier_fomet33(nuevo_folio, datos_reporte, df_atados_temp, pdf_portada)
                        
                        # 4. Compilar Dosier de Calidad final unificado
                        pdf_dosier = os.path.join(folio_folder, f"Dosier_Calidad_{nuevo_folio}.pdf")
                        utils_pdf.compilar_dosier_calidad(
                            pdf_dosier, 
                            pdf_portada, 
                            pdf_reporte, 
                            pdf_tecnico,
                            pdf_etiquetas, 
                            [cert_path], 
                            [oc_path]
                        )
                        
                        # 5. Registrar en Bases de Datos Maestras
                        # 5.1 Registro General de Recepciones
                        nuevo_registro_rep = pd.DataFrame([{
                            "Folio": nuevo_folio,
                            "Fecha": fecha_recepcion.strftime("%d/%m/%Y"),
                            "Hora": hora_recepcion,
                            "Proveedor": proveedor,
                            "Orden_Compra": orden_compra,
                            "Factura_Remision": factura_remision,
                            "Inspector": inspector,
                            "Estatus_General": estatus_general,
                            "Dossier_Ruta": pdf_dosier
                        }])
                        
                        st.session_state.BD_Reportes = pd.concat([st.session_state.BD_Reportes, nuevo_registro_rep], ignore_index=True)
                        guardar_db(st.session_state.BD_Reportes, BD_REPORTES, "Recepciones")
                        
                        # 5.2 Registro de atados
                        # Limpiar las columnas extras usadas para dibujar la etiqueta
                        cols_validas = [
                            "ID_Atado", "Folio", "ID_Atado_Proveedor", "Placa", "SKU", "Grado_Acero", 
                            "Num_Colada", "Lote_Heat", "Espesor_Medido_1_in", "Espesor_Medido_2_in", 
                            "Espesor_Medido_3_in", "Ancho_Medido_in", "Largo_Medido_in", "Cantidad_Hojas",
                            "Peso_Total_Kg", "Peso_Total_Lb", "Zinc_Medido_oz_ft2", 
                            "Dureza_Medida_HRB", "Aceitado_OK", "Defectos_Visuales", 
                            "Ubicacion_Almacen", "Estatus_Calidad", "Observaciones"
                        ]
                        df_atados_save = df_atados_temp[cols_validas]
                        
                        st.session_state.BD_Atados = pd.concat([st.session_state.BD_Atados, df_atados_save], ignore_index=True)
                        guardar_db(st.session_state.BD_Atados, BD_ATADOS, "Atados_Detalle")
                        
                        st.success(f"✅ ¡Ingreso procesado con éxito! Lote registrado bajo Folio: **{nuevo_folio}**")
                        
                        # Mostrar botones de descarga
                        col_dl1, col_dl2 = st.columns(2)
                        with col_dl1:
                            if os.path.exists(pdf_dosier):
                                with open(pdf_dosier, "rb") as f:
                                    st.download_button(
                                        label="📄 Descargar Dosier de Calidad Consolidado (PDF)",
                                        data=f.read(),
                                        file_name=f"Dosier_Calidad_{nuevo_folio}.pdf",
                                        mime="application/pdf",
                                        use_container_width=True
                                    )
                        with col_dl2:
                            if os.path.exists(pdf_solo_etiquetas):
                                with open(pdf_solo_etiquetas, "rb") as f:
                                    st.download_button(
                                        label="🏷️ Descargar Solo Etiquetas FO-MET-32 (PDF)",
                                        data=f.read(),
                                        file_name=f"Etiquetas_Solo_{nuevo_folio}.pdf",
                                        mime="application/pdf",
                                        use_container_width=True
                                    )
                                
                        # Dibujar las curvas de tolerancia usando Plotly
                        st.write("### 📈 Curvas de Tolerancia del Lote Recién Registrado")
                        fig_pl = go.Figure()
                        
                        # Líneas de límites
                        fig_pl.add_trace(go.Scatter(x=df_atados_temp["ID_Atado_Proveedor"], y=[esp_nom]*len(df_atados_temp), mode="lines", name=f"Nominal ({esp_nom:.4f} in)", line=dict(color="green")))
                        fig_pl.add_trace(go.Scatter(x=df_atados_temp["ID_Atado_Proveedor"], y=[esp_min]*len(df_atados_temp), mode="lines", name="LSL", line=dict(color="red", dash="dash")))
                        fig_pl.add_trace(go.Scatter(x=df_atados_temp["ID_Atado_Proveedor"], y=[esp_max]*len(df_atados_temp), mode="lines", name="USL", line=dict(color="red", dash="dash")))
                        
                        # Puntos reales
                        fig_pl.add_trace(go.Scatter(x=df_atados_temp["ID_Atado_Proveedor"], y=df_atados_temp["Espesor_Medido_1_in"], mode="markers", name="Espesor P1", marker=dict(symbol="circle", size=8)))
                        fig_pl.add_trace(go.Scatter(x=df_atados_temp["ID_Atado_Proveedor"], y=df_atados_temp["Espesor_Medido_2_in"], mode="markers", name="Espesor P2", marker=dict(symbol="triangle-up", size=8)))
                        fig_pl.add_trace(go.Scatter(x=df_atados_temp["ID_Atado_Proveedor"], y=df_atados_temp["Espesor_Medido_3_in"], mode="markers", name="Espesor P3", marker=dict(symbol="square", size=8)))
                        
                        fig_pl.update_layout(title="Control de Tolerancia de Espesor (in)", xaxis_title="ID Atado Proveedor", yaxis_title="Espesor (in)")
                        st.plotly_chart(fig_pl, use_container_width=True)
                        
                        # Dibujar el análisis gaussiano individual por atado
                        st.write("### 📊 Distribuciones Probabilísticas por Especificación Técnica")
                        for _, row_atd in df_atados_temp.iterrows():
                            id_atd_prov = str(row_atd.get("No_Atado", "default"))
                            renderizar_analisis_gaussiano_atado(row_atd, sku_info, key=f"plotly_gauss_upload_{id_atd_prov}")
                        
        except Exception as e:
            import traceback
            st.error(f"❌ Error al leer el archivo Excel: {e}")
            st.code(traceback.format_exc())

# =============================================================================
# MÓDULO 3: CONSULTA DE HISTORIAL Y EXPEDIENTES
# =============================================================================
elif opcion_menu == "🔍 Consulta de Historial":
    st.title("🔍 Consulta Histórica y Descarga de Dosiers de Calidad")
    st.markdown("Busque recepciones pasadas por fecha, proveedor o folio para descargar sus documentos y expedientes.")
    
    df_rep = st.session_state.BD_Reportes
    df_atd = st.session_state.BD_Atados
    
    if df_rep.empty:
        st.info("No hay registros históricos en el sistema.")
    else:
        # Filtros de búsqueda
        col_c1, col_c2, col_c3 = st.columns(3)
        with col_c1:
            f_folio = st.text_input("Buscar por Folio (ej. INC-2026-0001):")
            proveedores_list = ["Todos"] + df_rep["Proveedor"].dropna().unique().tolist()
            f_prov = st.selectbox("Filtrar por Proveedor:", proveedores_list)
            
        with col_c2:
            fecha_inicio = st.date_input("Fecha Inicio:", datetime.date.today() - datetime.timedelta(days=30))
            fecha_fin = st.date_input("Fecha Fin:", datetime.date.today())
            
        with col_c3:
            estatus_list = ["Todos", "Aceptado", "Rechazado", "Condicionado"]
            f_est = st.selectbox("Filtrar por Estatus de Calidad:", estatus_list)
            
        # Filtrado
        df_filtered = df_rep.copy()
        
        # Filtrado por fecha
        df_filtered["Fecha_DT"] = pd.to_datetime(df_filtered["Fecha"], format="%d/%m/%Y", errors="coerce").dt.date
        df_filtered = df_filtered[(df_filtered["Fecha_DT"] >= fecha_inicio) & (df_filtered["Fecha_DT"] <= fecha_fin)]
        
        if f_folio:
            df_filtered = df_filtered[df_filtered["Folio"].str.contains(f_folio, case=False)]
        if f_prov != "Todos":
            df_filtered = df_filtered[df_filtered["Proveedor"] == f_prov]
        if f_est != "Todos":
            df_filtered = df_filtered[df_filtered["Estatus_General"] == f_est]
            
        st.write("---")
        
        if df_filtered.empty:
            st.warning("⚠️ No se encontraron registros con los filtros seleccionados.")
        else:
            # Mostrar tabla histórica con porcentaje de aceptación de atados
            dict_acep = {}
            if not st.session_state.BD_Atados.empty:
                grouped = st.session_state.BD_Atados.groupby("Folio")
                for folio, group in grouped:
                    total_a = len(group)
                    aceptados_a = len(group[group["Estatus_Calidad"] == "Aceptado"])
                    pct_a = (aceptados_a / total_a * 100) if total_a > 0 else 0.0
                    dict_acep[folio] = f"{aceptados_a}/{total_a} ({pct_a:.1f}%)"
                    
            df_filtered_view = df_filtered.copy()
            df_filtered_view["Aceptación Atados"] = df_filtered_view["Folio"].map(dict_acep).fillna("0/0 (0.0%)")
            df_filtered_view = df_filtered_view.drop(columns=["Fecha_DT", "Dossier_Ruta"], errors="ignore")
            
            # Reordenar columnas para una vista limpia y profesional
            cols_order = ["Folio", "Fecha", "Hora", "Proveedor", "Orden_Compra", "Factura_Remision", "Aceptación Atados", "Estatus_General", "Inspector"]
            cols_order = [c for c in cols_order if c in df_filtered_view.columns]
            df_filtered_view = df_filtered_view[cols_order]
            
            st.dataframe(df_filtered_view.sort_values("Folio", ascending=False), use_container_width=True, hide_index=True)
            
            # Selector para ver detalles y descargar expediente
            st.write("### 📁 Selección de Expediente")
            folio_seleccionado = st.selectbox("Seleccione el Folio del reporte que desea descargar o revisar:", df_filtered["Folio"].tolist())
            
            if folio_seleccionado:
                info_recepcion = df_rep[df_rep["Folio"] == folio_seleccionado].iloc[0]
                df_atados_recepcion = df_atd[df_atd["Folio"] == folio_seleccionado]
                
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    st.write("**Datos de la Recepción:**")
                    st.write(f"- **Proveedor:** {info_recepcion['Proveedor']}")
                    st.write(f"- **Orden de Compra (PO):** {info_recepcion['Orden_Compra']}")
                    st.write(f"- **Factura / Remisión:** {info_recepcion['Factura_Remision']}")
                    st.write(f"- **Estatus de Liberación:** {info_recepcion['Estatus_General']}")
                    st.write(f"- **Inspector de Calidad:** {info_recepcion['Inspector']}")
                    
                with col_d2:
                    st.write("**Descargas Disponibles:**")
                    
                    # Botón para regenerar el expediente
                    if st.button("🔄 Regenerar y Actualizar Todos los PDFs de este Folio", key=f"btn_regen_{folio_seleccionado}", use_container_width=True):
                        import shutil
                        with st.spinner("Regenerando y actualizando todos los PDFs..."):
                            try:
                                folio_folder = os.path.join(CARPETAS_DIR, folio_seleccionado)
                                os.makedirs(folio_folder, exist_ok=True)
                                
                                pdf_reporte = os.path.join(folio_folder, f"Reporte_FO-MET-31_{folio_seleccionado}.pdf")
                                pdf_tecnico = os.path.join(folio_folder, f"Reporte_Tecnico_{folio_seleccionado}.pdf")
                                pdf_etiquetas = os.path.join(folio_folder, f"Etiquetas_FO-MET-32_{folio_seleccionado}.pdf")
                                pdf_solo_etiquetas = os.path.join(folio_folder, f"Etiquetas_Solo_FO-MET-32_{folio_seleccionado}.pdf")
                                pdf_portada = os.path.join(folio_folder, f"Portada_FO-MET-33_{folio_seleccionado}.pdf")
                                pdf_dosier = os.path.join(folio_folder, f"Dosier_Calidad_{folio_seleccionado}.pdf")
                                
                                cert_path = os.path.join(folio_folder, f"Certificado_Proveedor_{folio_seleccionado}.pdf")
                                oc_path = os.path.join(folio_folder, f"Orden_Compra_Sigrama_{folio_seleccionado}.pdf")
                                
                                # Si no existen Certificado u OC (por ejemplo, en folios de prueba), creamos portadas temporales como mock
                                if not os.path.exists(cert_path) or not os.path.exists(oc_path):
                                    mock_temp_pdf = os.path.join(folio_folder, f"temp_mock_support_{folio_seleccionado}.pdf")
                                    mock_datos = {
                                        "Fecha": info_recepcion["Fecha"],
                                        "Proveedor": info_recepcion["Proveedor"],
                                        "Orden_Compra": info_recepcion["Orden_Compra"],
                                        "Factura_Remision": info_recepcion["Factura_Remision"],
                                        "Inspector": info_recepcion["Inspector"],
                                        "Estatus_General": info_recepcion["Estatus_General"]
                                    }
                                    utils_pdf.generar_pdf_portada_dosier_fomet33(folio_seleccionado, mock_datos, df_atados_recepcion, mock_temp_pdf)
                                    if not os.path.exists(cert_path):
                                        shutil.copy(mock_temp_pdf, cert_path)
                                    if not os.path.exists(oc_path):
                                        shutil.copy(mock_temp_pdf, oc_path)
                                    if os.path.exists(mock_temp_pdf):
                                        os.remove(mock_temp_pdf)
                                        
                                datos_reporte = {
                                    "Fecha": info_recepcion["Fecha"],
                                    "Proveedor": info_recepcion["Proveedor"],
                                    "Orden_Compra": info_recepcion["Orden_Compra"],
                                    "Factura_Remision": info_recepcion["Factura_Remision"],
                                    "Inspector": info_recepcion["Inspector"],
                                    "Estatus_General": info_recepcion["Estatus_General"],
                                    "Hora": info_recepcion.get("Hora", "12:00")
                                }
                                
                                sku_del_lote = df_atados_recepcion["SKU"].iloc[0] if not df_atados_recepcion.empty else ""
                                df_skus = st.session_state.BD_Parametros
                                sku_match = df_skus[df_skus["SKU"] == sku_del_lote]
                                sku_info = sku_match.iloc[0].to_dict() if not sku_match.empty else {}
                                
                                img_curvas_path = os.path.join(folio_folder, f"Curvas_Tolerancia_{folio_seleccionado}.png")
                                if not sku_match.empty:
                                    utils_pdf.generar_graficas_tolerancia(df_atados_recepcion, sku_info, img_curvas_path)
                                    
                                utils_pdf.generar_pdf_reporte_consolidado_fomet31(folio_seleccionado, datos_reporte, df_atados_recepcion, sku_info, img_curvas_path, pdf_reporte)
                                utils_pdf.generar_pdf_reporte_tecnico_consolidado(folio_seleccionado, datos_reporte, df_atados_recepcion, pdf_tecnico)
                                utils_pdf.generar_pdf_etiqueta_atado_fomet32(folio_seleccionado, df_atados_recepcion, pdf_etiquetas)
                                utils_pdf.generar_pdf_solo_etiquetas(folio_seleccionado, df_atados_recepcion, pdf_solo_etiquetas)
                                utils_pdf.generar_pdf_portada_dosier_fomet33(folio_seleccionado, datos_reporte, df_atados_recepcion, pdf_portada)
                                
                                utils_pdf.compilar_dosier_calidad(
                                    pdf_dosier, 
                                    pdf_portada, 
                                    pdf_reporte, 
                                    pdf_tecnico,
                                    pdf_etiquetas, 
                                    [cert_path], 
                                    [oc_path]
                                )
                                
                                st.success("🔄 ¡Documentos del expediente regenerados y actualizados con éxito!")
                                st.rerun()
                            except Exception as ex:
                                st.error(f"❌ Error al regenerar los archivos: {ex}")

                    # Botón 1: Descargar el PDF del Dosier de Calidad
                    dossier_path = os.path.join(CARPETAS_DIR, folio_seleccionado, f"Dosier_Calidad_{folio_seleccionado}.pdf")
                    if os.path.exists(dossier_path):
                        with open(dossier_path, "rb") as f:
                            st.download_button(
                                label="📄 Descargar Dosier de Calidad (PDF)",
                                data=f.read(),
                                file_name=f"Dosier_Calidad_{folio_seleccionado}.pdf",
                                mime="application/pdf",
                                key=f"btn_dossier_{folio_seleccionado}"
                            )
                    else:
                        st.warning("⚠️ Archivo de Dosier PDF no encontrado en el servidor local.")
                        
                    # Botón 1.5: Descargar Solo Etiquetas (generar si no existe)
                    solo_etiquetas_path = os.path.join(CARPETAS_DIR, folio_seleccionado, f"Etiquetas_Solo_FO-MET-32_{folio_seleccionado}.pdf")
                    if not os.path.exists(solo_etiquetas_path) and not df_atados_recepcion.empty:
                        try:
                            os.makedirs(os.path.dirname(solo_etiquetas_path), exist_ok=True)
                            utils_pdf.generar_pdf_solo_etiquetas(folio_seleccionado, df_atados_recepcion, solo_etiquetas_path)
                        except Exception as e:
                            pass
                            
                    if os.path.exists(solo_etiquetas_path):
                        with open(solo_etiquetas_path, "rb") as f:
                            st.download_button(
                                label="🏷️ Descargar Solo Etiquetas FO-MET-32 (PDF)",
                                data=f.read(),
                                file_name=f"Etiquetas_Solo_{folio_seleccionado}.pdf",
                                mime="application/pdf",
                                key=f"btn_solo_etiquetas_{folio_seleccionado}"
                            )
                        
                    # Botón 2: Descargar todo el expediente como un archivo ZIP
                    folder_path = os.path.join(CARPETAS_DIR, folio_seleccionado)
                    if os.path.exists(folder_path):
                        # Comprimir carpeta en memoria
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                            for root, dirs, files in os.walk(folder_path):
                                for file in files:
                                    file_path = os.path.join(root, file)
                                    rel_path = os.path.relpath(file_path, folder_path)
                                    zip_file.write(file_path, rel_path)
                        zip_buffer.seek(0)
                        
                        st.download_button(
                            label="📥 Descargar Carpeta Electrónica Completa (ZIP)",
                            data=zip_buffer.getvalue(),
                            file_name=f"Expediente_Digital_{folio_seleccionado}.zip",
                            mime="application/zip",
                            key=f"btn_zip_{folio_seleccionado}"
                        )
                    else:
                        st.warning("⚠️ Carpeta del expediente digital no encontrada en el servidor.")
                
                # Mostrar tabla detallada de atados
                st.write("#### 📐 Mediciones Dimensionales de los Atados:")
                st.dataframe(df_atados_recepcion, use_container_width=True, hide_index=True)
                
                # Graficar curvas históricas de este lote
                st.write("#### 📈 Gráfica Histórica de Tolerancia de Espesor")
                sku_del_lote = df_atados_recepcion["SKU"].iloc[0] if not df_atados_recepcion.empty else ""
                df_skus = st.session_state.BD_Parametros
                sku_match = df_skus[df_skus["SKU"] == sku_del_lote]
                
                if not sku_match.empty and not df_atados_recepcion.empty:
                    sku_info = sku_match.iloc[0].to_dict()
                    esp_nom = float(sku_info.get("Espesor_Nominal_in", 0))
                    esp_min = esp_nom + float(sku_info.get("Espesor_Tolerancia_Min_in", 0))
                    esp_max = esp_nom + float(sku_info.get("Espesor_Tolerancia_Max_in", 0))
                    
                    fig_pl = go.Figure()
                    fig_pl.add_trace(go.Scatter(x=df_atados_recepcion["ID_Atado_Proveedor"], y=[esp_nom]*len(df_atados_recepcion), mode="lines", name=f"Nominal ({esp_nom:.4f} in)", line=dict(color="green")))
                    fig_pl.add_trace(go.Scatter(x=df_atados_recepcion["ID_Atado_Proveedor"], y=[esp_min]*len(df_atados_recepcion), mode="lines", name="LSL", line=dict(color="red", dash="dash")))
                    fig_pl.add_trace(go.Scatter(x=df_atados_recepcion["ID_Atado_Proveedor"], y=[esp_max]*len(df_atados_recepcion), mode="lines", name="USL", line=dict(color="red", dash="dash")))
                    
                    fig_pl.add_trace(go.Scatter(x=df_atados_recepcion["ID_Atado_Proveedor"], y=df_atados_recepcion["Espesor_Medido_1_in"], mode="markers", name="Punto 1", marker=dict(size=8)))
                    fig_pl.add_trace(go.Scatter(x=df_atados_recepcion["ID_Atado_Proveedor"], y=df_atados_recepcion["Espesor_Medido_2_in"], mode="markers", name="Punto 2", marker=dict(size=8)))
                    fig_pl.add_trace(go.Scatter(x=df_atados_recepcion["ID_Atado_Proveedor"], y=df_atados_recepcion["Espesor_Medido_3_in"], mode="markers", name="Punto 3", marker=dict(size=8)))
                    
                    fig_pl.update_layout(xaxis_title="ID Atado Proveedor", yaxis_title="Espesor (in)", margin=dict(l=20, r=20, t=40, b=20))
                    st.plotly_chart(fig_pl, use_container_width=True)
                    
                    # Dibujar el análisis gaussiano individual por atado
                    st.write("### 📊 Distribuciones Probabilísticas por Especificación Técnica")
                    for _, row_atd in df_atados_recepcion.iterrows():
                        row_with_limits = row_atd.to_dict()
                        row_with_limits['Espesor_Nominal'] = esp_nom
                        row_with_limits['Espesor_Tol_Min'] = sku_info.get("Espesor_Tolerancia_Min_in", -0.008)
                        row_with_limits['Espesor_Tol_Max'] = sku_info.get("Espesor_Tolerancia_Max_in", 0.008)
                        id_atd_prov = str(row_atd.get("ID_Atado_Proveedor", "default"))
                        renderizar_analisis_gaussiano_atado(row_with_limits, sku_info, key=f"plotly_gauss_hist_{id_atd_prov}")
                
                # --- SECCIÓN ADMINISTRATIVA: ELIMINACIÓN DE REGISTROS ---
                if is_admin:
                    st.write("---")
                    st.subheader("🗑️ Zona de Control Administrativo")
                    st.markdown("Como administrador, puede eliminar definitivamente esta recepción y todo su expediente.")
                    confirm_delete = st.checkbox("Confirmo que deseo borrar de forma permanente este lote y todos sus registros y archivos asociados.", key=f"conf_del_{folio_seleccionado}")
                    if st.button("🔴 Eliminar Recepción y Expediente Completo", type="primary", disabled=not confirm_delete, key=f"btn_del_{folio_seleccionado}"):
                        # 1. Borrar del archivo local de reportes
                        st.session_state.BD_Reportes = st.session_state.BD_Reportes[st.session_state.BD_Reportes["Folio"] != folio_seleccionado]
                        guardar_db(st.session_state.BD_Reportes, BD_REPORTES, "Recepciones")
                        
                        # 2. Borrar del archivo local de atados
                        st.session_state.BD_Atados = st.session_state.BD_Atados[st.session_state.BD_Atados["Folio"] != folio_seleccionado]
                        guardar_db(st.session_state.BD_Atados, BD_ATADOS, "Atados_Detalle")
                        
                        # 3. Borrar la carpeta física en disco
                        import shutil
                        folder_to_delete = os.path.join(CARPETAS_DIR, folio_seleccionado)
                        if os.path.exists(folder_to_delete):
                            try:
                                shutil.rmtree(folder_to_delete)
                            except Exception as ex_del:
                                st.warning(f"Se eliminaron los registros de la base de datos, pero no se pudo eliminar la carpeta física: {ex_del}")
                        
                        st.success(f"✅ Recepción y Expediente '{folio_seleccionado}' eliminados exitosamente.")
                        st.rerun()

# =============================================================================
# MÓDULO 4: CATÁLOGO DE TOLERANCIAS DE SKU
# =============================================================================
elif opcion_menu == "⚙️ Catálogo de Tolerancias de SKU":
    st.title("⚙️ Configuración y Catálogo de Parámetros de Materia Prima")
    st.markdown("Defina y modifique los valores nominales y límites de tolerancia aceptados para cada SKU en planta.")
    
    df_skus = st.session_state.BD_Parametros
    
    st.write("### 📦 SKUs Registrados en el Sistema")
    st.dataframe(df_skus, use_container_width=True, hide_index=True)
    
    st.write("---")
    
    # Solo administradores pueden agregar o modificar SKUs
    if not is_admin:
        st.error("🔒 Área de Configuración Protegida. Ingrese la contraseña de Administrador en la barra lateral para editar.")
    else:
        st.success("🔓 Acceso de Administrador Autorizado. Puede realizar cambios a los parámetros.")
        
        # Pestaña interna para agregar o modificar
        opt_edicion = st.selectbox("Seleccione Acción:", ["Agregar Nuevo Producto (SKU)", "Modificar Producto Existente"])
        
        if opt_edicion == "Agregar Nuevo Producto (SKU)":
            st.subheader("➕ Agregar Nuevo Producto de Materia Prima")
            with st.form("form_add_sku"):
                col_a1, col_a2 = st.columns(2)
                with col_a1:
                    new_sku = st.text_input("Código de SKU (Único):", placeholder="Ej. SKU-DECP-14")
                    new_nombre = st.text_input("Nombre / Descripción del Material:", placeholder="Ej. Lámina Decapada Calibre 14")
                    new_tipo = st.selectbox("Tipo de Lámina:", ["Decapada", "Galvanizada"])
                    new_grado = st.text_input("Grado de Acero:", placeholder="Ej. SAE 1008 CS")
                    new_dureza = st.number_input("Dureza Máxima Permitida (HRB):", min_value=10, max_value=120, value=75)
                    new_aceitado = st.selectbox("¿Aceitado Requerido? (Decapado):", ["Sí", "No", "N/A"])
                    
                with col_a2:
                    new_esp_nom = st.number_input("Espesor Nominal (in):", format="%.4f", value=0.0750)
                    new_esp_min = st.number_input("Tolerancia Espesor Mínima (in - negativo):", format="%.4f", value=-0.0050)
                    new_esp_max = st.number_input("Tolerancia Espesor Máxima (in - positivo):", format="%.4f", value=0.0050)
                    
                    new_anc_nom = st.number_input("Ancho Nominal (in):", format="%.2f", value=48.00)
                    new_anc_min = st.number_input("Tolerancia Ancho Mínima (in - negativo):", format="%.3f", value=-0.120)
                    new_anc_max = st.number_input("Tolerancia Ancho Máxima (in - positivo):", format="%.3f", value=0.120)
                    
                    new_lrg_nom = st.number_input("Largo Nominal (in):", format="%.2f", value=120.00)
                    new_lrg_min = st.number_input("Tolerancia Largo Mínima (in - negativo):", format="%.3f", value=-0.250)
                    new_lrg_max = st.number_input("Tolerancia Largo Máxima (in - positivo):", format="%.3f", value=0.250)
                    
                    new_zinc_nom = st.number_input("Zinc Nominal (oz/ft² - Galvanizado):", format="%.2f", value=0.60)
                    new_zinc_min = st.number_input("Zinc Mínimo Aceptable (oz/ft² - Galvanizado):", format="%.2f", value=0.50)
                    
                btn_add_sku = st.form_submit_button("💾 Guardar Nuevo SKU en Catálogo")
                
                if btn_add_sku:
                    if not new_sku or not new_nombre or not new_grado:
                        st.error("❌ Todos los campos de identificación (SKU, Nombre y Grado) son obligatorios.")
                    elif new_sku in df_skus["SKU"].values:
                        st.error(f"❌ El SKU '{new_sku}' ya existe en el catálogo.")
                    else:
                        nuevo_sku_record = {
                            "SKU": new_sku,
                            "Nombre": new_nombre,
                            "Tipo_Lamina": new_tipo,
                            "Grado_Acero": new_grado,
                            "Espesor_Nominal_in": new_esp_nom,
                            "Espesor_Tolerancia_Min_in": new_esp_min,
                            "Espesor_Tolerancia_Max_in": new_esp_max,
                            "Ancho_Nominal_in": new_anc_nom,
                            "Ancho_Tolerancia_Min_in": new_anc_min,
                            "Ancho_Tolerancia_Max_in": new_anc_max,
                            "Largo_Nominal_in": new_lrg_nom,
                            "Largo_Tolerancia_Min_in": new_lrg_min,
                            "Largo_Tolerancia_Max_in": new_lrg_max,
                            "Zinc_Nominal_oz_ft2": new_zinc_nom,
                            "Zinc_Min_oz_ft2": new_zinc_min,
                            "Aceitado_Requerido": new_aceitado,
                            "Dureza_Max_HRB": new_dureza
                        }
                        
                        st.session_state.BD_Parametros = pd.concat([st.session_state.BD_Parametros, pd.DataFrame([nuevo_sku_record])], ignore_index=True)
                        guardar_db(st.session_state.BD_Parametros, BD_PARAMETROS, "Materia_Prima")
                        st.success(f"✅ SKU '{new_sku}' agregado y sincronizado con éxito.")
                        st.rerun()
                        
        elif opt_edicion == "Modificar Producto Existente":
            st.subheader("📝 Modificar Parámetros de SKU Existente")
            sku_seleccionado = st.selectbox("Seleccione el SKU a modificar:", df_skus["SKU"].tolist())
            
            if sku_seleccionado:
                sku_sel_info = df_skus[df_skus["SKU"] == sku_seleccionado].iloc[0].to_dict()
                
                with st.form("form_edit_sku"):
                    col_e1, col_e2 = st.columns(2)
                    with col_e1:
                        st.text_input("Código de SKU (No editable):", value=sku_sel_info["SKU"], disabled=True)
                        edit_nombre = st.text_input("Nombre / Descripción del Material:", value=sku_sel_info["Nombre"])
                        edit_tipo = st.selectbox("Tipo de Lámina:", ["Decapada", "Galvanizada"], index=0 if sku_sel_info["Tipo_Lamina"] == "Decapada" else 1)
                        edit_grado = st.text_input("Grado de Acero:", value=sku_sel_info["Grado_Acero"])
                        edit_dureza = st.number_input("Dureza Máxima Permitida (HRB):", min_value=10, max_value=120, value=int(sku_sel_info["Dureza_Max_HRB"]))
                        edit_aceitado = st.selectbox("¿Aceitado Requerido? (Decapado):", ["Sí", "No", "N/A"], index=["Sí", "No", "N/A"].index(str(sku_sel_info.get("Aceitado_Requerido", "N/A"))))
                        
                    with col_e2:
                        edit_esp_nom = st.number_input("Espesor Nominal (in):", format="%.4f", value=float(sku_sel_info["Espesor_Nominal_in"]))
                        edit_esp_min = st.number_input("Tolerancia Espesor Mínima (in - negativo):", format="%.4f", value=float(sku_sel_info["Espesor_Tolerancia_Min_in"]))
                        edit_esp_max = st.number_input("Tolerancia Espesor Máxima (in - positivo):", format="%.4f", value=float(sku_sel_info["Espesor_Tolerancia_Max_in"]))
                        
                        edit_anc_nom = st.number_input("Ancho Nominal (in):", format="%.2f", value=float(sku_sel_info["Ancho_Nominal_in"]))
                        edit_anc_min = st.number_input("Tolerancia Ancho Mínima (in - negativo):", format="%.3f", value=float(sku_sel_info["Ancho_Tolerancia_Min_in"]))
                        edit_anc_max = st.number_input("Tolerancia Ancho Máxima (in - positivo):", format="%.3f", value=float(sku_sel_info["Ancho_Tolerancia_Max_in"]))
                        
                        edit_lrg_nom = st.number_input("Largo Nominal (in):", format="%.2f", value=float(sku_sel_info.get("Largo_Nominal_in", 120.00)))
                        edit_lrg_min = st.number_input("Tolerancia Largo Mínima (in - negativo):", format="%.3f", value=float(sku_sel_info.get("Largo_Tolerancia_Min_in", -0.250)))
                        edit_lrg_max = st.number_input("Tolerancia Largo Máxima (in - positivo):", format="%.3f", value=float(sku_sel_info.get("Largo_Tolerancia_Max_in", 0.250)))
                        
                        edit_zinc_nom = st.number_input("Zinc Nominal (oz/ft² - Galvanizado):", format="%.2f", value=float(sku_sel_info.get("Zinc_Nominal_oz_ft2", 0.0)))
                        edit_zinc_min = st.number_input("Zinc Mínimo Aceptable (oz/ft² - Galvanizado):", format="%.2f", value=float(sku_sel_info.get("Zinc_Min_oz_ft2", 0.0)))
                        
                    btn_edit_sku = st.form_submit_button("💾 Guardar Cambios")
                    
                    if btn_edit_sku:
                        # Actualizar en el dataframe
                        idx_match = df_skus[df_skus["SKU"] == sku_seleccionado].index[0]
                        
                        st.session_state.BD_Parametros.at[idx_match, "Nombre"] = edit_nombre
                        st.session_state.BD_Parametros.at[idx_match, "Tipo_Lamina"] = edit_tipo
                        st.session_state.BD_Parametros.at[idx_match, "Grado_Acero"] = edit_grado
                        st.session_state.BD_Parametros.at[idx_match, "Espesor_Nominal_in"] = edit_esp_nom
                        st.session_state.BD_Parametros.at[idx_match, "Espesor_Tolerancia_Min_in"] = edit_esp_min
                        st.session_state.BD_Parametros.at[idx_match, "Espesor_Tolerancia_Max_in"] = edit_esp_max
                        st.session_state.BD_Parametros.at[idx_match, "Ancho_Nominal_in"] = edit_anc_nom
                        st.session_state.BD_Parametros.at[idx_match, "Ancho_Tolerancia_Min_in"] = edit_anc_min
                        st.session_state.BD_Parametros.at[idx_match, "Ancho_Tolerancia_Max_in"] = edit_anc_max
                        st.session_state.BD_Parametros.at[idx_match, "Largo_Nominal_in"] = edit_lrg_nom
                        st.session_state.BD_Parametros.at[idx_match, "Largo_Tolerancia_Min_in"] = edit_lrg_min
                        st.session_state.BD_Parametros.at[idx_match, "Largo_Tolerancia_Max_in"] = edit_lrg_max
                        st.session_state.BD_Parametros.at[idx_match, "Zinc_Nominal_oz_ft2"] = edit_zinc_nom
                        st.session_state.BD_Parametros.at[idx_match, "Zinc_Min_oz_ft2"] = edit_zinc_min
                        st.session_state.BD_Parametros.at[idx_match, "Aceitado_Requerido"] = edit_aceitado
                        st.session_state.BD_Parametros.at[idx_match, "Dureza_Max_HRB"] = edit_dureza
                        
                        guardar_db(st.session_state.BD_Parametros, BD_PARAMETROS, "Materia_Prima")
                        st.success(f"✅ SKU '{sku_seleccionado}' modificado correctamente.")
                        st.rerun()
