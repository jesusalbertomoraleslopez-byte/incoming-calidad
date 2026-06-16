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
BASE_DIR = r"C:\Users\albertol\.gemini\antigravity\scratch\incoming_calidad"
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

# Rutas de base de datos
BD_PARAMETROS = os.path.join(BASE_DIR, "BD_Parametros_Materia_Prima.xlsx")
BD_REPORTES = os.path.join(BASE_DIR, "BD_Reportes_Incoming.xlsx")
BD_ATADOS = os.path.join(BASE_DIR, "BD_Atados_Incoming.xlsx")
PLANTILLA_PATH = os.path.join(BASE_DIR, "plantilla_incoming_calidad.xlsx")
CARPETAS_DIR = os.path.join(BASE_DIR, "carpetas_electronicas")

# Renderizado de Banner Corporativo Sigrama
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

# Navegación lateral
st.sidebar.title("🧭 Navegación")
opcion_menu = st.sidebar.radio("Seleccione un Módulo:", [
    "📊 Analíticas y Dashboard",
    "📥 Registro de Recepción (Incoming)",
    "🔍 Consulta de Historial",
    "⚙️ Catálogo de Tolerancias de SKU"
])

# Control de accesos para administración en la barra lateral
st.sidebar.write("---")
st.sidebar.title("🔐 Control de Acceso")
admin_pass_input = st.sidebar.text_input("Contraseña Administrador:", type="password")

def es_admin():
    return admin_pass_input == "SigramaCalidad2026"

is_admin = es_admin()
if is_admin:
    st.sidebar.success("Modo Administrador Activo")
else:
    st.sidebar.warning("Modo Consulta Activo")

# =============================================================================
# MÓDULO 1: ANALÍTICAS Y DASHBOARD
# =============================================================================
if opcion_menu == "📊 Analíticas y Dashboard":
    st.title("📊 Dashboard Planta Metales - Control de Calidad en Recepción")
    
    df_rep = st.session_state.BD_Reportes
    df_atd = st.session_state.BD_Atados
    
    if df_rep.empty:
        st.info("No hay reportes de recepción registrados actualmente. Vaya al módulo de 'Registro de Recepción' para comenzar.")
    else:
        # Calcular KPIs
        total_lotes = len(df_rep)
        total_atados = len(df_atd)
        
        lotes_aceptados = len(df_rep[df_rep["Estatus_General"] == "Aceptado"])
        lotes_condicionados = len(df_rep[df_rep["Estatus_General"] == "Condicionado"])
        lotes_rechazados = len(df_rep[df_rep["Estatus_General"] == "Rechazado"])
        
        peso_total_kg = df_atd["Peso_Total_Kg"].sum()
        peso_total_lb = df_atd["Peso_Total_Lb"].sum()
        
        # Mostrar Métricas Clave
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📦 Lotes Recibidos", f"{total_lotes} Lotes")
        m2.metric("🟢 Lotes Aceptados", f"{lotes_aceptados} Lotes", f"{lotes_aceptados/total_lotes*100:.1f}% del total")
        m3.metric("🔴 Lotes Rechazados", f"{lotes_rechazados} Lotes", f"-{lotes_rechazados/total_lotes*100:.1f}% Tasa de Rechazo")
        m4.metric("⚖️ Total Acero Recibido", f"{peso_total_kg:,.0f} Kg", f"{peso_total_lb:,.0f} Lb")
        
        st.write("---")
        
        # Distribución de Estatus
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.subheader("📋 Estado General de Lotes Recibidos")
            fig_pie = px.pie(
                df_rep, 
                names="Estatus_General", 
                color="Estatus_General",
                color_discrete_map={"Aceptado": "#2E7D32", "Condicionado": "#FBC02D", "Rechazado": "#C62828"},
                hole=0.4
            )
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with col_g2:
            st.subheader("🏢 Distribución de Volumen por Proveedor (Kg)")
            df_prov = df_atd.merge(df_rep[["Folio", "Proveedor"]], on="Folio", how="left")
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
                
        st.write("---")
        st.subheader("📋 Listado de Recepciones Recientes")
        st.dataframe(df_rep.sort_values("Folio", ascending=False), use_container_width=True, hide_index=True)

# =============================================================================
# MÓDULO 2: REGISTRO DE RECEPCIÓN (INCOMING)
# =============================================================================
elif opcion_menu == "📥 Registro de Recepción (Incoming)":
    st.title("📥 Registro de Control de Calidad en Recepción")
    st.markdown("Suba la plantilla Excel con las mediciones y los documentos adjuntos (Certificados, OC) para generar el Dosier de Calidad.")
    
    # Descargar plantilla corporativa
    if os.path.exists(PLANTILLA_PATH):
        with open(PLANTILLA_PATH, "rb") as f:
            st.download_button(
                label="📥 Descargar Formato de Plantilla Corporativa (.xlsx)",
                data=f.read(),
                file_name="plantilla_incoming_calidad_sigrama.xlsx",
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
        oc_file = st.file_uploader("🛍️ Orden de Compra de Sigrama (PDF)", type=["pdf"], key="oc_file")
        
    with col_u2:
        defect_photos = st.file_uploader("📸 Fotos de Defectos Visuales (Opcional)", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
        
    st.write("### 🧮 Mediciones Dimensionales de Campo")
    mediciones_file = st.file_uploader("📥 Suba la plantilla Excel con mediciones completadas:", type=["xlsx"])
    
    if mediciones_file:
        try:
            df_med = pd.read_excel(mediciones_file, sheet_name=0)
            
            # Limpiar filas vacías
            df_med = df_med.dropna(subset=["SKU", "ID_Atado_Proveedor"])
            
            st.write("🔍 **Pre-visualización de Datos Cargados:**")
            st.dataframe(df_med, use_container_width=True)
            
            st.write("---")
            
            # Botón para procesar
            if st.button("🚀 Procesar e Integrar Recepción a Base de Datos"):
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
                            
                        # Guardar fotos de defectos
                        if defect_photos:
                            for idx_ph, ph in enumerate(defect_photos):
                                ph_ext = os.path.splitext(ph.name)[1]
                                ph_path = os.path.join(folio_folder, f"Foto_Defecto_{idx_ph+1}_{nuevo_folio}{ph_ext}")
                                with open(ph_path, "wb") as f:
                                    f.write(ph.read())
                                    
                        # 3. Procesar atados y evaluar tolerancias
                        df_atados_lote = []
                        lote_rechazado = False
                        sku_principal = df_med["SKU"].iloc[0] # Tomamos el SKU del lote
                        
                        # Obtener parámetros del SKU
                        df_skus = st.session_state.BD_Parametros
                        sku_match = df_skus[df_skus["SKU"] == sku_principal]
                        
                        if sku_match.empty:
                            st.error(f"❌ El SKU '{sku_principal}' no está registrado en el Catálogo de Tolerancias de la empresa.")
                            st.stop()
                            
                        sku_info = sku_match.iloc[0].to_dict()
                        
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
                        
                        for idx_atd, row_med in df_med.iterrows():
                            id_atd_int = f"{nuevo_folio}-A{idx_atd+1:02d}"
                            
                            # Validar mediciones
                            # Validar mediciones de los 12 espesores
                            meds = [float(row_med[f"Espesor_Medido_{j}_in"]) for j in range(1, 13)]
                            ancho = float(row_med["Ancho_Medido_in"])
                            largo = float(row_med["Largo_Medido_in"])
                            cant_hojas = int(row_med.get("Cantidad_Hojas", 0))
                            peso_kg = float(row_med["Peso_Total_Kg"])
                            peso_lb = float(row_med["Peso_Total_Lb"])
                            dureza = float(row_med.get("Dureza_Medida_HRB", 0))
                            
                            # Verificar si es galvanizada
                            zinc_val = 0.0
                            if sku_info.get("Tipo_Lamina") == "Galvanizada":
                                zinc_val = float(row_med.get("Zinc_Medido_oz_ft2", 0))
                                
                            # Aceitado para decapada
                            aceitado_val = str(row_med.get("Aceitado_OK", "N/A"))
                            
                            # Evaluación
                            motivos_rechazo = []
                            for j in range(1, 13):
                                m_val = meds[j-1]
                                if not (esp_min <= m_val <= esp_max):
                                    motivos_rechazo.append(f"Espesor {j} fuera de especificación")
                            if not (anc_min <= ancho <= anc_max): motivos_rechazo.append("Ancho fuera de especificación")
                            if not (lrg_min <= largo <= lrg_max): motivos_rechazo.append("Largo fuera de especificación")
                            if sku_info.get("Tipo_Lamina") == "Galvanizada" and zinc_val < zinc_min:
                                motivos_rechazo.append("Zinc inferior al límite")
                            if sku_info.get("Tipo_Lamina") == "Decapada" and sku_info.get("Aceitado_Requerido") == "Sí" and aceitado_val.upper() != "SÍ":
                                motivos_rechazo.append("Aceitado ausente")
                            if dureza > dureza_max:
                                motivos_rechazo.append("Dureza excede el límite máximo")
                            if peso_kg > 2500:
                                motivos_rechazo.append("Peso del atado excede el límite de 2.5 Toneladas (2500 Kg) según RFQ")
                                
                            atd_status = "Aceptado"
                            if motivos_rechazo:
                                atd_status = "Rechazado"
                                lote_rechazado = True
                                
                            atado_record = {
                                "ID_Atado": id_atd_int,
                                "Folio": nuevo_folio,
                                "ID_Atado_Proveedor": str(row_med["ID_Atado_Proveedor"]),
                                "SKU": sku_principal,
                                "Grado_Acero": str(row_med["Grado_Acero"]),
                                "Num_Colada": str(row_med["Num_Colada"]),
                                "Lote_Heat": str(row_med["Lote_Heat"]),
                                "Espesor_Medido_1_in": meds[0],
                                "Espesor_Medido_2_in": meds[1],
                                "Espesor_Medido_3_in": meds[2],
                                "Espesor_Medido_4_in": meds[3],
                                "Espesor_Medido_5_in": meds[4],
                                "Espesor_Medido_6_in": meds[5],
                                "Espesor_Medido_7_in": meds[6],
                                "Espesor_Medido_8_in": meds[7],
                                "Espesor_Medido_9_in": meds[8],
                                "Espesor_Medido_10_in": meds[9],
                                "Espesor_Medido_11_in": meds[10],
                                "Espesor_Medido_12_in": meds[11],
                                "Ancho_Medido_in": ancho,
                                "Largo_Medido_in": largo,
                                "Cantidad_Hojas": cant_hojas,
                                "Peso_Total_Kg": peso_kg,
                                "Peso_Total_Lb": peso_lb,
                                "Zinc_Medido_oz_ft2": zinc_val,
                                "Dureza_Medida_HRB": dureza,
                                "Aceitado_OK": aceitado_val,
                                "Defectos_Visuales": str(row_med.get("Defectos_Visuales", "Ninguno")),
                                "Ubicacion_Almacen": str(row_med.get("Ubicacion_Almacen", "Almacén Metales")),
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
                        if lote_rechazado:
                            estatus_general = "Rechazado"
                            
                        # Generar archivos de reportes PDF individuales
                        pdf_reporte = os.path.join(folio_folder, f"Reporte_FO-MET-31_{nuevo_folio}.pdf")
                        pdf_etiquetas = os.path.join(folio_folder, f"Etiquetas_FO-MET-32_{nuevo_folio}.pdf")
                        pdf_portada = os.path.join(folio_folder, f"Portada_FO-MET-33_{nuevo_folio}.pdf")
                        
                        datos_reporte = {
                            "Fecha": fecha_recepcion.strftime("%d/%m/%Y"),
                            "Proveedor": proveedor,
                            "Orden_Compra": orden_compra,
                            "Factura_Remision": factura_remision,
                            "Inspector": inspector,
                            "Estatus_General": estatus_general
                        }
                        
                        utils_pdf.generar_pdf_reporte_consolidado_fomet31(nuevo_folio, datos_reporte, df_atados_temp, sku_info, img_curvas_path, pdf_reporte)
                        utils_pdf.generar_pdf_etiqueta_atado_fomet32(nuevo_folio, df_atados_temp, pdf_etiquetas)
                        utils_pdf.generar_pdf_portada_dosier_fomet33(nuevo_folio, datos_reporte, df_atados_temp, pdf_portada)
                        
                        # 4. Compilar Dosier de Calidad final unificado
                        pdf_dosier = os.path.join(folio_folder, f"Dosier_Calidad_{nuevo_folio}.pdf")
                        utils_pdf.compilar_dosier_calidad(
                            pdf_dosier, 
                            pdf_portada, 
                            pdf_reporte, 
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
                            "ID_Atado", "Folio", "ID_Atado_Proveedor", "SKU", "Grado_Acero", 
                            "Num_Colada", "Lote_Heat", 
                            "Espesor_Medido_1_in", "Espesor_Medido_2_in", "Espesor_Medido_3_in", 
                            "Espesor_Medido_4_in", "Espesor_Medido_5_in", "Espesor_Medido_6_in", 
                            "Espesor_Medido_7_in", "Espesor_Medido_8_in", "Espesor_Medido_9_in", 
                            "Espesor_Medido_10_in", "Espesor_Medido_11_in", "Espesor_Medido_12_in", 
                            "Ancho_Medido_in", "Largo_Medido_in", "Cantidad_Hojas",
                            "Peso_Total_Kg", "Peso_Total_Lb", "Zinc_Medido_oz_ft2", 
                            "Dureza_Medida_HRB", "Aceitado_OK", "Defectos_Visuales", 
                            "Ubicacion_Almacen", "Estatus_Calidad", "Observaciones"
                        ]
                        df_atados_save = df_atados_temp[cols_validas]
                        
                        st.session_state.BD_Atados = pd.concat([st.session_state.BD_Atados, df_atados_save], ignore_index=True)
                        guardar_db(st.session_state.BD_Atados, BD_ATADOS, "Atados_Detalle")
                        
                        st.success(f"✅ ¡Ingreso procesado con éxito! Lote registrado bajo Folio: **{nuevo_folio}**")
                        
                        # Mostrar el Dosier PDF listo
                        if os.path.exists(pdf_dosier):
                            with open(pdf_dosier, "rb") as f:
                                st.download_button(
                                    label="📄 Descargar Dosier de Calidad Consolidado (PDF)",
                                    data=f.read(),
                                    file_name=f"Dosier_Calidad_{nuevo_folio}.pdf",
                                    mime="application/pdf"
                                )
                                
                        # Dibujar las curvas de tolerancia usando Plotly
                        st.write("### 📈 Curvas de Tolerancia del Lote Recién Registrado")
                        fig_pl = go.Figure()
                        
                        # Líneas de límites
                        fig_pl.add_trace(go.Scatter(x=df_atados_temp["ID_Atado_Proveedor"], y=[esp_nom]*len(df_atados_temp), mode="lines", name=f"Nominal ({esp_nom:.4f} in)", line=dict(color="green")))
                        fig_pl.add_trace(go.Scatter(x=df_atados_temp["ID_Atado_Proveedor"], y=[esp_min]*len(df_atados_temp), mode="lines", name="LSL", line=dict(color="red", dash="dash")))
                        fig_pl.add_trace(go.Scatter(x=df_atados_temp["ID_Atado_Proveedor"], y=[esp_max]*len(df_atados_temp), mode="lines", name="USL", line=dict(color="red", dash="dash")))
                        
                        # Puntos reales
                        markers_plotly = ["circle", "triangle-up", "square", "diamond", "cross", "x", "pentagon", "hexagon", "star", "bowtie", "hourglass", "circle-open"]
                        for j in range(1, 13):
                            fig_pl.add_trace(go.Scatter(
                                x=df_atados_temp["ID_Atado_Proveedor"], 
                                y=df_atados_temp[f"Espesor_Medido_{j}_in"], 
                                mode="markers", 
                                name=f"Espesor P{j}", 
                                marker=dict(symbol=markers_plotly[j-1], size=6)
                            ))
                        
                        fig_pl.update_layout(title="Control de Tolerancia de Espesor (in)", xaxis_title="ID Atado Proveedor", yaxis_title="Espesor (in)")
                        st.plotly_chart(fig_pl, use_container_width=True)
                        
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
            # Mostrar tabla histórica
            df_filtered_view = df_filtered.drop(columns=["Fecha_DT", "Dossier_Ruta"], errors="ignore")
            st.dataframe(df_filtered_view, use_container_width=True)
            
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
                    
                    for j in range(1, 13):
                        fig_pl.add_trace(go.Scatter(
                            x=df_atados_recepcion["ID_Atado_Proveedor"], 
                            y=df_atados_recepcion[f"Espesor_Medido_{j}_in"], 
                            mode="markers", 
                            name=f"Punto {j}", 
                            marker=dict(size=6)
                        ))
                    
                    fig_pl.update_layout(xaxis_title="ID Atado Proveedor", yaxis_title="Espesor (in)", margin=dict(l=20, r=20, t=40, b=20))
                    st.plotly_chart(fig_pl, use_container_width=True)
                
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
