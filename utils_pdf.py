import os
import io
import datetime
import matplotlib
matplotlib.use('Agg')  # Usar backend no interactivo
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import scipy.stats as stats
from pypdf import PdfWriter, PdfReader

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.graphics.shapes import Drawing, Rect, String, Line, Group

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "logo_sigrama.png")

# =============================================================================
# 1. MOTOR DE GRAFICACIÓN DE CURVAS DE TOLERANCIA
# =============================================================================
def generar_graficas_tolerancia(df_atados, params, output_path):
    """
    Genera una imagen PNG con las curvas de tolerancia de espesor y ancho para el lote de atados.
    Muestra los límites USL, LSL y Nominal, junto con las mediciones reales de cada atado.
    """
    df_atados = consolidar_df_atados_para_pdf(df_atados)
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.5))
    
    # 1.1 Gráfica de Espesor
    esp_nom = float(params.get("Espesor_Nominal_in", 0))
    esp_min = esp_nom + float(params.get("Espesor_Tolerancia_Min_in", 0))
    esp_max = esp_nom + float(params.get("Espesor_Tolerancia_Max_in", 0))
    
    atado_ids = []
    for _, row in df_atados.iterrows():
        id_prov = str(row["ID_Atado_Proveedor"])
        placa = row.get("Placa")
        if placa and pd.notna(placa):
            try:
                id_prov = f"{id_prov}-P{int(float(placa))}"
            except Exception:
                id_prov = f"{id_prov}-P{placa}"
        atado_ids.append(id_prov)

    # Para cada atado, tenemos 3 mediciones de espesor
    m1 = df_atados["Espesor_Medido_1_in"].astype(float).tolist()
    m2 = df_atados["Espesor_Medido_2_in"].astype(float).tolist()
    m3 = df_atados["Espesor_Medido_3_in"].astype(float).tolist()
    
    x = range(len(atado_ids))
    
    # Dibujar límites
    ax1.axhline(y=esp_nom, color="green", linestyle="-", linewidth=1.5, label=f"Nominal ({esp_nom:.4f} in)")
    ax1.axhline(y=esp_min, color="red", linestyle="--", linewidth=1.2, label=f"LSL ({esp_min:.4f} in)")
    ax1.axhline(y=esp_max, color="red", linestyle="--", linewidth=1.2, label=f"USL ({esp_max:.4f} in)")
    
    # Graficar puntos medidos
    ax1.scatter(x, m1, color="blue", marker="o", s=40, label="Medición 1")
    ax1.scatter(x, m2, color="orange", marker="^", s=40, label="Medición 2")
    ax1.scatter(x, m3, color="purple", marker="s", s=40, label="Medición 3")
    
    # Línea que une los promedios de espesor
    promedios = [(m1[i] + m2[i] + m3[i])/3 for i in x]
    ax1.plot(x, promedios, color="gray", linestyle=":", alpha=0.7, label="Promedio")
    
    ax1.set_title("Curvas de Tolerancia - Espesor (in)", fontsize=11, fontweight="bold", pad=10)
    ax1.set_xticks(list(x))
    ax1.set_xticklabels(atado_ids, rotation=30, ha="right", fontsize=8)
    ax1.set_ylabel("Espesor (Pulgadas)", fontsize=9)
    ax1.grid(True, linestyle=":", alpha=0.5)
    ax1.legend(loc="upper right", fontsize=8)
    
    # 1.2 Gráfica de Ancho
    anc_nom = float(params.get("Ancho_Nominal_in", 0))
    anc_min = anc_nom + float(params.get("Ancho_Tolerancia_Min_in", 0))
    anc_max = anc_nom + float(params.get("Ancho_Tolerancia_Max_in", 0))
    
    ancho_med = df_atados["Ancho_Medido_in"].astype(float).tolist()
    
    ax2.axhline(y=anc_nom, color="green", linestyle="-", linewidth=1.5, label=f"Nominal ({anc_nom:.2f} in)")
    ax2.axhline(y=anc_min, color="red", linestyle="--", linewidth=1.2, label=f"LSL ({anc_min:.2f} in)")
    ax2.axhline(y=anc_max, color="red", linestyle="--", linewidth=1.2, label=f"USL ({anc_max:.2f} in)")
    
    ax2.scatter(x, ancho_med, color="darkcyan", marker="D", s=45, label="Medido")
    ax2.plot(x, ancho_med, color="darkcyan", linestyle=":", alpha=0.5)
    
    ax2.set_title("Curvas de Tolerancia - Ancho (in)", fontsize=11, fontweight="bold", pad=10)
    ax2.set_xticks(list(x))
    ax2.set_xticklabels(atado_ids, rotation=30, ha="right", fontsize=8)
    ax2.set_ylabel("Ancho (Pulgadas)", fontsize=9)
    ax2.grid(True, linestyle=":", alpha=0.5)
    ax2.legend(loc="upper right", fontsize=8)
    
    # Ajustes estéticos y guardado
    plt.tight_layout()
    plt.savefig(output_path, dpi=200)
    plt.close()
    print("Gráfica de curvas de tolerancia generada en:", output_path)

def obtener_valores_tolerancia(row):
    """
    Retorna (esp_nom, esp_tol_min, esp_tol_max) para una fila.
    Si no están presentan en la fila, los busca en BD_Parametros_Materia_Prima.xlsx usando el SKU.
    """
    esp_nom = row.get('Espesor_Nominal')
    esp_tol_min = row.get('Espesor_Tol_Min')
    esp_tol_max = row.get('Espesor_Tol_Max')
    
    if pd.isna(esp_nom) or esp_nom is None or pd.isna(esp_tol_min) or esp_tol_min is None or pd.isna(esp_tol_max) or esp_tol_max is None:
        try:
            param_path = os.path.join(BASE_DIR, "BD_Parametros_Materia_Prima.xlsx")
            if os.path.exists(param_path):
                df_params = pd.read_excel(param_path)
                sku = row.get('SKU')
                sku_match = df_params[df_params["SKU"] == sku]
                if not sku_match.empty:
                    esp_nom = float(sku_match.iloc[0]["Espesor_Nominal_in"])
                    esp_tol_min = float(sku_match.iloc[0]["Espesor_Tolerancia_Min_in"])
                    esp_tol_max = float(sku_match.iloc[0]["Espesor_Tolerancia_Max_in"])
        except Exception as e:
            print(f"Error al cargar parámetros en obtener_valores_tolerancia: {e}")
            
    if pd.isna(esp_nom) or esp_nom is None:
        esp_nom = 0.0750
    if pd.isna(esp_tol_min) or esp_tol_min is None:
        esp_tol_min = -0.008
    if pd.isna(esp_tol_max) or esp_tol_max is None:
        esp_tol_max = 0.008
        
    return float(esp_nom), float(esp_tol_min), float(esp_tol_max)

def consolidar_df_atados_para_pdf(df_atados):
    """
    Consolida un DataFrame que contiene registros individuales de placas
    (4 placas por atado) en un único registro por atado.
    """
    if df_atados is None or df_atados.empty:
        return df_atados

    # Copiar para evitar Side-Effects
    df = df_atados.copy()

    # Si la columna "ID_Atado_Proveedor" no existe pero sí "No_Atado", la creamos
    if "ID_Atado_Proveedor" not in df.columns and "No_Atado" in df.columns:
        df["ID_Atado_Proveedor"] = df["No_Atado"]

    # Si no hay identificador de atado del proveedor, no agrupamos
    if "ID_Atado_Proveedor" not in df.columns:
        return df

    # Agrupar por ID_Atado_Proveedor manteniendo el orden original de aparición
    grouped = df.groupby("ID_Atado_Proveedor", sort=False)
    consolidated_rows = []

    for prov_id, group in grouped:
        # Ordenar el grupo por Placa si existe
        if "Placa" in group.columns:
            try:
                group = group.sort_values("Placa")
            except Exception:
                pass

        first_row = group.iloc[0]

        # 1. Promediar mediciones de espesor en cada punto (P1, P2, P3) a lo largo de las placas
        avg_m1 = group["Espesor_Medido_1_in"].astype(float).mean()
        avg_m2 = group["Espesor_Medido_2_in"].astype(float).mean()
        avg_m3 = group["Espesor_Medido_3_in"].astype(float).mean()

        # 2. Promediar ancho y largo si están presentes
        avg_ancho = group["Ancho_Medido_in"].astype(float).mean() if "Ancho_Medido_in" in group.columns else first_row.get("Ancho_Medido_in", 0.0)
        avg_largo = group["Largo_Medido_in"].astype(float).mean() if "Largo_Medido_in" in group.columns else first_row.get("Largo_Medido_in", 0.0)

        # 3. Cantidad de hojas y Peso (se toma el primer valor/promedio del atado, NO la suma,
        #    ya que cada fila virtual de placa contiene duplicado el peso/hojas del atado completo)
        cant_hojas = group["Cantidad_Hojas"].iloc[0] if "Cantidad_Hojas" in group.columns else first_row.get("Cantidad_Hojas", 0)
        peso_kg = group["Peso_Total_Kg"].iloc[0] if "Peso_Total_Kg" in group.columns else first_row.get("Peso_Total_Kg", 0.0)
        peso_lb = group["Peso_Total_Lb"].iloc[0] if "Peso_Total_Lb" in group.columns else first_row.get("Peso_Total_Lb", 0.0)

        # 4. Estatus de Calidad: Aceptado solo si todas las placas del atado están aceptadas
        statuses = group["Estatus_Calidad"].astype(str).tolist()
        if "Rechazado" in statuses:
            final_status = "Rechazado"
        elif "Condicionado" in statuses:
            final_status = "Condicionado"
        else:
            final_status = "Aceptado"

        # 5. ID Interno consolidado (rango del primero al último, por ej: INC-2026-0001-A01 a A04)
        if "ID_Atado" in group.columns:
            internal_ids = group["ID_Atado"].astype(str).tolist()
            unique_ids = list(set(internal_ids))
            if len(unique_ids) == 1:
                final_internal_id = unique_ids[0]
            elif len(internal_ids) > 1:
                first_id = internal_ids[0]
                last_id = internal_ids[-1]
                if first_id == last_id:
                    final_internal_id = first_id
                else:
                    parts = first_id.split("-")
                    if len(parts) > 1:
                        first_suffix = parts[-1]
                        last_suffix = last_id.split("-")[-1]
                        base_id = "-".join(parts[:-1])
                        final_internal_id = f"{base_id}-{first_suffix} a {last_suffix}"
                    else:
                        final_internal_id = f"{first_id} a {last_id.split('-')[-1]}"
            else:
                final_internal_id = internal_ids[0]
        else:
            final_internal_id = first_row.get("ID_Atado", "")

        # 6. Recopilar todas las 12 mediciones del atado (3 puntos * 4 placas)
        #    para hacer el análisis gaussiano completo de la campana
        all_measurements = []
        for _, r in group.iterrows():
            all_measurements.extend([
                float(r["Espesor_Medido_1_in"]),
                float(r["Espesor_Medido_2_in"]),
                float(r["Espesor_Medido_3_in"])
            ])

        # Crear la fila consolidada
        row_dict = first_row.to_dict()
        row_dict.update({
            "ID_Atado": final_internal_id,
            "Espesor_Medido_1_in": avg_m1,
            "Espesor_Medido_2_in": avg_m2,
            "Espesor_Medido_3_in": avg_m3,
            "Ancho_Medido_in": avg_ancho,
            "Largo_Medido_in": avg_largo,
            "Peso_Total_Kg": peso_kg,
            "Peso_Total_Lb": peso_lb,
            "Cantidad_Hojas": cant_hojas,
            "Estatus_Calidad": final_status,
            "Placa": None,  # Limpiar la placa para que no imprima "Placa X"
            "All_Measurements": all_measurements
        })
        consolidated_rows.append(row_dict)

    return pd.DataFrame(consolidated_rows)

# =============================================================================
# 2. FORMATOS SGC: ELEMENTOS DE CABECERA Y PIE DE PÁGINA
# =============================================================================

def draw_sigrama_sgc_decorations(canvas, doc, doc_code, title_text):
    """Dibuja el encabezado y pie de página oficiales bajo la norma SGC FO-SGC-02."""
    canvas.saveState()
    
    # Franja superior roja Sigrama
    canvas.setFillColor(colors.HexColor("#D32F2F"))
    canvas.rect(36, 745, 540, 4, fill=1, stroke=0)
    
    # Logotipo
    if os.path.exists(LOGO_PATH):
        try:
            # Dibujar el logo en la parte superior izquierda de la cabecera (escala de aspect ratio horizontal)
            canvas.drawImage(LOGO_PATH, 36, 755, width=120, height=22, mask='auto')
        except Exception:
            canvas.setFont("Helvetica-Bold", 12)
            canvas.drawString(36, 755, "INDUSTRIA SIGRAMA")
    else:
        canvas.setFont("Helvetica-Bold", 12)
        canvas.drawString(36, 755, "INDUSTRIA SIGRAMA")
        
    # Marcador de Control de Calidad Superior Izquierdo Oficial (FO-MET-3X)
    canvas.setFont("Helvetica-Bold", 11)
    canvas.setFillColor(colors.HexColor("#D32F2F"))
    canvas.drawRightString(576, 765, doc_code)
    
    # Metadatos de Revisión de Control Documental
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.black)
    canvas.drawRightString(576, 753, "Revisión 01")
    
    # Título Central del Formato Oficial
    canvas.setFont("Helvetica-Bold", 10)
    canvas.drawCentredString(315, 755, title_text.upper())
    
    # Fecha más destacada
    canvas.setFont("Helvetica-Bold", 8)
    months_es = {
        1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
        7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
    }
    today = datetime.date.today()
    fecha_hoy = f"{today.day} de {months_es[today.month]} de {today.year}"
    canvas.drawString(36, 732, f"Fecha de Emisión: {fecha_hoy}")
    
    # Pie de Página Legal y Control del SGC (FO-SGC-02)
    canvas.setStrokeColor(colors.HexColor("#D32F2F"))
    canvas.setLineWidth(1)
    canvas.line(36, 45, 36, 25)
    canvas.setFont("Helvetica-Bold", 7)
    canvas.drawString(42, 37, "FO-SGC-02")
    canvas.setFont("Helvetica", 6)
    canvas.setFillColor(colors.HexColor("#424242"))
    texto_legal = "PROHIBIDA LA REPRODUCCIÓN TOTAL O PARCIAL, POR CUALQUIER MEDIO O PROCEDIMIENTO, SIN AUTORIZACIÓN DE INDUSTRIA SIGRAMA S.A. DE C.V."
    canvas.drawString(95, 37, texto_legal)
    
    canvas.restoreState()

# =============================================================================
# 3. GENERACIÓN DE REPORTES EN PDF (INDIVIDUALES)
# =============================================================================
def generar_pdf_reporte_consolidado_fomet31(folio, datos_reporte, df_atados, sku_params, img_curvas_path, output_pdf_path):
    """
    Construye el documento de calidad de recepción FO-MET-31 con las tablas de inspección
    y la gráfica de curvas de tolerancia.
    """
    df_atados = consolidar_df_atados_para_pdf(df_atados)
    doc = SimpleDocTemplate(output_pdf_path, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story = []
    styles = getSampleStyleSheet()
    
    # Estilos de texto
    style_blanco_bold = ParagraphStyle('WB_Met', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=8)
    style_normal_bold = ParagraphStyle('NB_Met', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=8)
    style_normal_text = ParagraphStyle('NT_Met', parent=styles['Normal'], fontSize=8)
    
    story.append(Spacer(1, 10))
    
    # --- PANEL 1: DATOS GENERALES DE RECEPCIÓN ---
    t_header = Table([[Paragraph("DATOS GENERALES DE RECEPCIÓN E INSPECCIÓN", style_blanco_bold)]], colWidths=[540])
    t_header.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#757575")), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(t_header)
    
    datos_panel = [
        [Paragraph("FOLIO RECEPCIÓN:", style_normal_bold), Paragraph(f"<b>{folio}</b>", style_normal_bold),
         Paragraph("FECHA RECEPCIÓN:", style_normal_bold), Paragraph(str(datos_reporte.get('Fecha', '')), style_normal_text)],
        [Paragraph("PROVEEDOR:", style_normal_bold), Paragraph(str(datos_reporte.get('Proveedor', '')), style_normal_text),
         Paragraph("ORDEN DE COMPRA:", style_normal_bold), Paragraph(str(datos_reporte.get('Orden_Compra', '')), style_normal_text)],
        [Paragraph("FACTURA/REMISION:", style_normal_bold), Paragraph(str(datos_reporte.get('Factura_Remision', '')), style_normal_text),
         Paragraph("INSPECTOR:", style_normal_bold), Paragraph(str(datos_reporte.get('Inspector', '')), style_normal_text)],
        [Paragraph("TIPO MATERIAL:", style_normal_bold), Paragraph(str(sku_params.get('Tipo_Lamina', '')), style_normal_text),
         Paragraph("ESTATUS GENERAL LOTE:", style_normal_bold), Paragraph(f"<b>{datos_reporte.get('Estatus_General', '')}</b>", style_normal_bold)]
    ]
    t_panel = Table(datos_panel, colWidths=[120, 150, 120, 150])
    t_panel.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#F5F5F5")),
        ('BACKGROUND', (2,0), (2,-1), colors.HexColor("#F5F5F5"))
    ]))
    story.append(t_panel)
    story.append(Spacer(1, 15))
    
    # --- PANEL 2: ESPECIFICACIONES Y TOLERANCIAS NOMINALES ---
    t_spec_header = Table([[Paragraph("PARÁMETROS NOMINALES DE ACEPTACIÓN (SKU: " + str(sku_params.get('SKU', '')) + ")", style_blanco_bold)]], colWidths=[540])
    t_spec_header.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#757575")), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(t_spec_header)
    
    esp_nom = float(sku_params.get('Espesor_Nominal_in', 0))
    esp_tol_min = float(sku_params.get('Espesor_Tolerancia_Min_in', 0))
    esp_tol_max = float(sku_params.get('Espesor_Tolerancia_Max_in', 0))
    
    anc_nom = float(sku_params.get('Ancho_Nominal_in', 0))
    anc_tol_min = float(sku_params.get('Ancho_Tolerancia_Min_in', 0))
    anc_tol_max = float(sku_params.get('Ancho_Tolerancia_Max_in', 0))
    
    lrg_nom = float(sku_params.get('Largo_Nominal_in', 0))
    lrg_tol_min = float(sku_params.get('Largo_Tolerancia_Min_in', 0))
    lrg_tol_max = float(sku_params.get('Largo_Tolerancia_Max_in', 0))
    
    zinc_nom = float(sku_params.get('Zinc_Nominal_oz_ft2', 0))
    zinc_min = float(sku_params.get('Zinc_Min_oz_ft2', 0))
    
    datos_specs = [
        [Paragraph("CARACTERÍSTICA", style_blanco_bold), Paragraph("VALOR NOMINAL", style_blanco_bold),
         Paragraph("TOLERANCIA MÍNIMA", style_blanco_bold), Paragraph("TOLERANCIA MÁXIMA", style_blanco_bold)],
        [Paragraph("Espesor (in)", style_normal_bold), Paragraph(f"{esp_nom:.4f}", style_normal_text),
         Paragraph(f"{esp_nom + esp_tol_min:.4f}", style_normal_text), Paragraph(f"{esp_nom + esp_tol_max:.4f}", style_normal_text)],
        [Paragraph("Ancho (in)", style_normal_bold), Paragraph(f"{anc_nom:.2f}", style_normal_text),
         Paragraph(f"{anc_nom + anc_tol_min:.2f}", style_normal_text), Paragraph(f"{anc_nom + anc_tol_max:.2f}", style_normal_text)],
        [Paragraph("Largo (in)", style_normal_bold), Paragraph(f"{lrg_nom:.2f}", style_normal_text),
         Paragraph(f"{lrg_nom + lrg_tol_min:.2f}", style_normal_text), Paragraph(f"{lrg_nom + lrg_tol_max:.2f}", style_normal_text)],
        [Paragraph("Recubrimiento Zinc (oz/ft²)", style_normal_bold), Paragraph(f"{zinc_nom:.2f}", style_normal_text),
         Paragraph(f"{zinc_min:.2f}", style_normal_text), Paragraph("-", style_normal_text)],
        [Paragraph("Aceitado Requerido", style_normal_bold), Paragraph(str(sku_params.get('Aceitado_Requerido', 'N/D')), style_normal_text),
         Paragraph("-", style_normal_text), Paragraph("-", style_normal_text)],
        [Paragraph("Dureza Máxima (HRB)", style_normal_bold), Paragraph(str(sku_params.get('Dureza_Max_HRB', 'N/D')), style_normal_text),
         Paragraph("-", style_normal_text), Paragraph("-", style_normal_text)]
    ]
    t_specs = Table(datos_specs, colWidths=[160, 120, 130, 130])
    t_specs.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E0E0E0")),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('ALIGN', (1,1), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.append(t_specs)
    story.append(Spacer(1, 15))
    
    # --- PANEL 3: TABLA DE MEDICIONES FÍSICAS POR ATADO ---
    t_med_header = Table([[Paragraph("MEDICIONES REALES Y EVALUACIÓN DE ATADOS", style_blanco_bold)]], colWidths=[540])
    t_med_header.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#757575")), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(t_med_header)
    
    tabla_mediciones = [[
        Paragraph("ID ATADO", style_blanco_bold),
        Paragraph("COLADA / HEAT", style_blanco_bold),
        Paragraph("ESPESORES (in)", style_blanco_bold),
        Paragraph("ANCHO (in)", style_blanco_bold),
        Paragraph("PESO (Kg)", style_blanco_bold),
        Paragraph("ESTATUS", style_blanco_bold)
    ]]
    
    for _, row in df_atados.iterrows():
        esp_med = f"{float(row['Espesor_Medido_1_in']):.4f} / {float(row['Espesor_Medido_2_in']):.4f} / {float(row['Espesor_Medido_3_in']):.4f}"
        
        status_val = str(row['Estatus_Calidad'])
        status_color = "#2E7D32" if status_val == "Aceptado" else "#C62828"
        status_html = f"<b><font color='{status_color}'>{status_val}</font></b>"
        
        placa = row.get("Placa")
        id_prov_lbl = str(row['ID_Atado_Proveedor'])
        if placa and pd.notna(placa):
            try:
                id_prov_lbl = f"{id_prov_lbl} - P{int(float(placa))}"
            except Exception:
                id_prov_lbl = f"{id_prov_lbl} - P{placa}"
                
        tabla_mediciones.append([
            Paragraph(f"{id_prov_lbl}<br/><font color='#616161' size='6'>{row['ID_Atado']}</font>", style_normal_text),
            Paragraph(f"{row['Num_Colada']}<br/><font color='#616161' size='6'>{row['Lote_Heat']}</font>", style_normal_text),
            Paragraph(esp_med, style_normal_text),
            Paragraph(f"{float(row['Ancho_Medido_in']):.2f}", style_normal_text),
            Paragraph(f"{int(row['Peso_Total_Kg']):,}", style_normal_text),
            Paragraph(status_html, style_normal_text)
        ])
        
    t_med = Table(tabla_mediciones, colWidths=[100, 100, 160, 60, 60, 60])
    t_med.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#757575")),
        ('ALIGN', (2,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.append(t_med)
    
    # Firmas en la primera página
    story.append(Spacer(1, 20))
    datos_firmas = [
        [Paragraph("_____________________________<br/>INSPECTOR DE CALIDAD", style_normal_bold),
         Paragraph("_____________________________<br/>AUTORIZADO POR PRODUCCIÓN", style_normal_bold)]
    ]
    t_firmas = Table(datos_firmas, colWidths=[270, 270])
    t_firmas.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'TOP')
    ]))
    story.append(t_firmas)
    
    # Segunda Página: Gráficos de Curvas de Tolerancia
    story.append(PageBreak())
    story.append(Spacer(1, 15))
    story.append(Paragraph("<b>CURVAS DE TOLERANCIA Y PROCESO DE MEDICIÓN</b>", styles['Heading2']))
    story.append(Spacer(1, 5))
    story.append(Paragraph("Este gráfico representa de forma visual la ubicación de cada medición real con respecto a las bandas de aceptación configuradas en el Sistema de Gestión de Calidad (SGC).", style_normal_text))
    story.append(Spacer(1, 10))
    
    if os.path.exists(img_curvas_path):
        story.append(Image(img_curvas_path, width=540, height=220))
        
    def decorate(canvas, doc):
        draw_sigrama_sgc_decorations(canvas, doc, "FO-MET-31", "REPORTE CONSOLIDADO DE INSPECCIÓN")
        
    doc.build(story, onFirstPage=decorate, onLaterPages=decorate)
    print("Reporte FO-MET-31 creado en:", output_pdf_path)

def generar_grafica_individual_probabilidad(row, folder_path):
    esp_nom, esp_tol_min, esp_tol_max = obtener_valores_tolerancia(row)
    
    lsl = esp_tol_min
    usl = esp_tol_max
    
    if "All_Measurements" in row and isinstance(row["All_Measurements"], list) and len(row["All_Measurements"]) > 0:
        devs = [float(x) - esp_nom for x in row["All_Measurements"]]
    else:
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
    p_out = p_below + p_above
    p_out_pct = p_out * 100.0
    
    fig, ax = plt.subplots(figsize=(7.5, 3.5))
    
    x = np.linspace(-0.015, 0.015, 500)
    y = stats.norm.pdf(x, loc=mu, scale=sigma)
    
    id_prov = row.get("ID_Atado_Proveedor", "Atado")
    placa = row.get("Placa")
    if placa and pd.notna(placa):
        try:
            id_prov = f"{id_prov} - Placa {int(float(placa))}"
        except Exception:
            id_prov = f"{id_prov} - Placa {placa}"
            
    ax.plot(x, y, label=f"{id_prov} ({p_out_pct:.1f}%)", color="#1f77b4", linewidth=2)
    
    ax.axvspan(lsl, usl, color='#e8f5e9', alpha=0.5, label='Zona de Especificación')
    ax.axhline(0, color='black', linewidth=0.5)
    ax.axvline(0, color='green', linestyle='--', linewidth=1.5, label=f'Nominal ({esp_nom:.4f}")')
    ax.axvline(lsl, color='red', linestyle=':', linewidth=1.5, label=f'LSL ({esp_nom + lsl:.4f}")')
    ax.axvline(usl, color='red', linestyle=':', linewidth=1.5, label=f'USL ({esp_nom + usl:.4f}")')
    
    tipo = row.get("Tipo_Lamina", "Decapada")
    sku = row.get("SKU", "")
    ax.set_title(f"{tipo} - {sku} - Nominal: {esp_nom:.4f}\"", fontsize=11, fontweight='bold', color='#0d47a1', pad=10)
    
    ax.set_xlim(-0.015, 0.015)
    ax.set_xlabel("Desviación Micrométrica Real (in)", fontsize=9)
    ax.set_ylabel("Densidad Probabilística de Gauss", fontsize=9)
    ax.grid(True, which='both', linestyle=':', color='lightgray', alpha=0.7)
    ax.legend(loc='upper right', fontsize=8)
    
    plt.tight_layout()
    img_path = os.path.join(folder_path, f"temp_prob_{row['ID_Atado']}.png")
    plt.savefig(img_path, dpi=200)
    plt.close()
    return img_path

def generar_pdf_etiqueta_atado_fomet32(folio, df_atados, output_pdf_path):
    """
    Genera la tarjeta de identificación de atado de materia prima FO-MET-32
    (diseño corporativo de 2 páginas por atado: Hoja 1 Datos, Hoja 2 Probabilidad).
    """
    df_atados = consolidar_df_atados_para_pdf(df_atados)
    doc = SimpleDocTemplate(output_pdf_path, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()
    
    style_tit_label = ParagraphStyle('TL_Lab', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=14)
    style_cell_header = ParagraphStyle('CH_Lab', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=8, textColor=colors.HexColor("#757575"))
    style_cell_value = ParagraphStyle('CV_Lab', parent=styles['Normal'], fontSize=9, fontName="Helvetica")
    style_cell_value_bold = ParagraphStyle('CVB_Lab', parent=styles['Normal'], fontSize=10, fontName="Helvetica-Bold")
    style_cell_value_white_bold = ParagraphStyle('CVWB_Lab', parent=styles['Normal'], fontSize=10, fontName="Helvetica-Bold", textColor=colors.white, alignment=1)
    
    # Lista de archivos temporales creados para borrar al final
    temp_files = []
    folder_path = os.path.dirname(output_pdf_path)
    
    for idx, row in df_atados.iterrows():
        # --- HOJA 1: IDENTIFICACIÓN DEL ATADO ---
        story.append(Spacer(1, 10))
        
        # 1. Cabecera de la Etiqueta
        if os.path.exists(LOGO_PATH):
            logotipo_header = Image(LOGO_PATH, width=150, height=27)
        else:
            logotipo_header = Paragraph("<b>INDUSTRIA SIGRAMA</b>", style_cell_value)
            
        header_table_data = [
            [logotipo_header, Paragraph("ATADO DE MATERIA PRIMA", style_tit_label)]
        ]
        header_table = Table(header_table_data, colWidths=[200, 340])
        header_table.setStyle(TableStyle([
            ('BACKGROUND', (1,0), (1,0), colors.HexColor("#D32F2F")),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (0,0), 'LEFT'),
            ('ALIGN', (1,0), (1,0), 'CENTER'),
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#D32F2F"))
        ]))
        story.append(header_table)
        
        # Sub-título Identificación
        story.append(Table([[Paragraph("IDENTIFICACIÓN DEL ATADO", style_blanco_bold_label(styles))]], colWidths=[540], 
                           style=TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#D32F2F")), ('ALIGN', (0,0), (-1,-1), 'CENTER')])))
        
        # 2. Bloque de Identificación del Atado (Izquierda: Tabla, Derecha: Imagen + Observaciones)
        
        # Left table (fields)
        ident_left_data = [
            [Paragraph("📅 FECHA DE RECEPCIÓN", style_cell_header), Paragraph(datetime.date.today().strftime("%d/%m/%Y"), style_cell_value_bold)],
            [Paragraph("🚛 PROVEEDOR", style_cell_header), Paragraph(str(row.get('Proveedor', 'Ternium')), style_cell_value_bold)],
            [Paragraph("📦 TIPO DE MATERIAL", style_cell_header), Paragraph(str(row.get('Tipo_Lamina', 'Decapada')), style_cell_value_bold)],
            [Paragraph("✨ ACABADO", style_cell_header), Paragraph(str(row.get('Acabado', 'Pasivado' if row.get('Tipo_Lamina')=='Galvanizada' else 'Aceitado')), style_cell_value)],
            [Paragraph("✔️ GRADO DE ACERO", style_cell_header), Paragraph(str(row['Grado_Acero']), style_cell_value_bold)],
            [Paragraph("🏷️ NÚMERO DE ATADO (INTERNO)", style_cell_header), Paragraph(f"<b>{row['ID_Atado']}</b>", style_cell_value_bold)]
        ]
        
        ident_left_table = Table(ident_left_data, colWidths=[150, 120])
        ident_left_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#D32F2F")),
            ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#F5F5F5")),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4)
        ]))
        
        # Right table (Image, OBSERVACIONES banner, text)
        observaciones_atado = str(row.get('Observaciones', '')).strip()
        if observaciones_atado == 'nan' or not observaciones_atado or observaciones_atado == 'Sin observaciones registradas.':
            observaciones_atado = "Excelente estado superficial."
            
        stack_img_path = os.path.join(BASE_DIR, "stack_metal_sheets.png")
        if os.path.exists(stack_img_path):
            stack_img = Image(stack_img_path, width=120, height=80)
            stack_img.hAlign = 'CENTER'
        else:
            stack_img = Paragraph("<b>[IMAGEN DEL ATADO]</b>", style_cell_value)
            
        ident_right_data = [
            [stack_img],
            [Paragraph("OBSERVACIONES", style_blanco_bold_label(styles))],
            [Paragraph(observaciones_atado, style_cell_value)]
        ]
        
        ident_right_table = Table(ident_right_data, colWidths=[270])
        ident_right_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#D32F2F")),
            ('BACKGROUND', (0,1), (0,1), colors.HexColor("#D32F2F")),
            ('ALIGN', (0,0), (0,0), 'CENTER'),
            ('ALIGN', (0,1), (0,1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,2), (0,2), 8),
            ('TOPPADDING', (0,2), (0,2), 8)
        ]))
        
        ident_container_data = [[ident_left_table, ident_right_table]]
        ident_container = Table(ident_container_data, colWidths=[270, 270])
        ident_container.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0)
        ]))
        story.append(ident_container)
        story.append(Spacer(1, 10))
        
        # 3. Bloque de Tres Columnas: Especificaciones, Control de Hojas, Trazabilidad
        esp_nom, esp_tol_min, esp_tol_max = obtener_valores_tolerancia(row)
        esp_min = esp_nom + esp_tol_min
        esp_max = esp_nom + esp_tol_max
        
        # Tabla de Especificaciones (Izquierda)
        spec_rows = [
            [Paragraph("ESPECIFICACIONES DEL MATERIAL", style_blanco_bold_label_small(styles)), "", "", ""],
            [Paragraph("CARACTERÍSTICA", style_cell_header), Paragraph("VALOR REAL", style_cell_header), Paragraph("TOL. MÍN", style_cell_header), Paragraph("TOL. MÁX", style_cell_header)],
            [Paragraph("Espesor Real", style_cell_header), Paragraph(f"{float(row['Espesor_Medido_1_in']):.4f}", style_cell_value_bold), Paragraph(f"{esp_min:.4f}", style_cell_value), Paragraph(f"{esp_max:.4f}", style_cell_value)],
            [Paragraph("Ancho (in)", style_cell_header), Paragraph(f"{float(row['Ancho_Medido_in']):.2f}", style_cell_value), Paragraph(f"{row.get('Ancho_Tol_Min_Val', 47.8):.2f}", style_cell_value), Paragraph(f"{row.get('Ancho_Tol_Max_Val', 48.2):.2f}", style_cell_value)],
            [Paragraph("Largo (in)", style_cell_header), Paragraph(f"{float(row['Largo_Medido_in']):.2f}", style_cell_value), Paragraph("-", style_cell_value), Paragraph("-", style_cell_value)],
            [Paragraph("Cant. Hojas", style_cell_header), Paragraph(str(int(row['Cantidad_Hojas'])), style_cell_value), Paragraph("-", style_cell_value), Paragraph("-", style_cell_value)],
            [Paragraph("Peso Total (Kg)", style_cell_header), Paragraph(f"{int(row['Peso_Total_Kg'])}", style_cell_value_bold), Paragraph("-", style_cell_value), Paragraph("-", style_cell_value)],
            [Paragraph("Peso Total (Lb)", style_cell_header), Paragraph(f"{int(row['Peso_Total_Lb'])}", style_cell_value), Paragraph("-", style_cell_value), Paragraph("-", style_cell_value)]
        ]
        spec_table = Table(spec_rows, colWidths=[75, 45, 30, 30])
        spec_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#D32F2F")),
            ('SPAN', (0,0), (3,0)),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")),
            ('BACKGROUND', (0,1), (-1,1), colors.HexColor("#F5F5F5")),
            ('ALIGN', (1,1), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 7)
        ]))
        
        # Tabla de Control de Hojas (Medio)
        control_rows = [
            [Paragraph("CONTROL DE HOJAS", style_blanco_bold_label_small(styles)), "", ""],
            [Paragraph(f"RECIBIDAS (RECEPCIÓN): <b>{int(row['Cantidad_Hojas'])}</b> HOJAS", style_cell_value_bold), "", ""],
            [Paragraph("REGISTRO DE RETIROS", style_cell_header), "", ""],
            [Paragraph("FECHA RETIRO", style_cell_header), Paragraph("CANTIDAD RETIRADA", style_cell_header), Paragraph("RESTANTE ALMACÉN", style_cell_header)],
            ["", "", ""],
            ["", "", ""],
            ["", "", ""],
            ["", "", ""]
        ]
        control_table = Table(control_rows, colWidths=[55, 55, 70])
        control_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#D32F2F")),
            ('SPAN', (0,0), (2,0)),
            ('SPAN', (0,1), (2,1)),
            ('SPAN', (0,2), (2,2)),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")),
            ('BACKGROUND', (0,2), (-1,3), colors.HexColor("#F5F5F5")),
            ('ALIGN', (0,3), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('HEIGHT', (0,4), (-1,-1), 16)
        ]))
        
        # Tabla de Trazabilidad (Derecha)
        status_value = str(row['Estatus_Calidad']).upper()
        placa = row.get("Placa")
        id_prov_lbl = str(row['ID_Atado_Proveedor'])
        if placa and pd.notna(placa):
            try:
                id_prov_lbl = f"{id_prov_lbl} - Placa {int(float(placa))}"
            except Exception:
                id_prov_lbl = f"{id_prov_lbl} - Placa {placa}"
        traz_rows = [
            [Paragraph("TRAZABILIDAD", style_blanco_bold_label_small(styles)), ""],
            [Paragraph("🔥 NÚMERO DE COLADA", style_cell_header), Paragraph(str(row['Num_Colada']), style_cell_value_bold)],
            [Paragraph("📦 LOTE / HEAT", style_cell_header), Paragraph(str(row['Lote_Heat']), style_cell_value)],
            [Paragraph("🏷️ ATADO PROVEEDOR", style_cell_header), Paragraph(id_prov_lbl, style_cell_value)],
            [Paragraph("🏢 UBICACIÓN ALMACÉN", style_cell_header), Paragraph(str(row['Ubicacion_Almacen']), style_cell_value_bold)],
            [Paragraph("✔️ ESTATUS DE CALIDAD", style_cell_header), Paragraph(status_value, style_cell_value_white_bold)]
        ]
        traz_table = Table(traz_rows, colWidths=[100, 80])
        traz_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#D32F2F")),
            ('SPAN', (0,0), (1,0)),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")),
            ('BACKGROUND', (0,1), (0,-1), colors.HexColor("#F5F5F5")),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BACKGROUND', (1, 5), (1, 5), colors.HexColor("#2E7D32" if row['Estatus_Calidad']=='Aceptado' else "#C62828")),
            ('ALIGN', (1, 5), (1, 5), 'CENTER')
        ]))
        
        # Compilar las tres tablas en paralelo
        three_cols_table = Table([[spec_table, control_table, traz_table]], colWidths=[180, 180, 180])
        three_cols_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0)
        ]))
        story.append(three_cols_table)
        story.append(Spacer(1, 10))
        
        # 4. Sección de Manejo y Conservación
        story.append(Table([[Paragraph("MANEJO Y CONSERVACIÓN", style_blanco_bold_label_small(styles))]], colWidths=[540],
                           style=TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#D32F2F")), ('ALIGN', (0,0), (-1,-1), 'CENTER')])))
        
        manejo_data = [
            [Paragraph("<b>🚜 MANEJAR CON EQUIPO ADECUADO</b><br/>Usar montacargas de capacidad apropiada.", style_cell_value),
             Paragraph("<b>☔ PROTEGER DE LA HUMEDAD</b><br/>Almacenar en área seca para evitar oxidación.", style_cell_value),
             Paragraph("<b>🚫 NO APOYAR OBJETOS SOBRE EL ATADO</b><br/>No colocar peso directo sobre el atado.", style_cell_value)]
        ]
        manejo_table = Table(manejo_data, colWidths=[180, 180, 180])
        manejo_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#D32F2F")),
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F5F5F5")),
            ('VALIGN', (0,0), (-1,-1), 'TOP')
        ]))
        story.append(manejo_table)
        
        # Generar gráfica de campana de Gauss para ponerla en la parte baja de la Hoja 1 y en la Hoja 2
        img_prob_path = generar_grafica_individual_probabilidad(row, folder_path)
        temp_files.append(img_prob_path)
        
        if os.path.exists(img_prob_path):
            story.append(Spacer(1, 8))
            img_gauss_1 = Image(img_prob_path, width=420, height=160)
            img_gauss_1.hAlign = 'CENTER'
            story.append(img_gauss_1)
        
        # --- HOJA 2: REPORTE TÉCNICO DE INGENIERÍA DE CALIDAD ---
        story.append(PageBreak())
        story.append(Spacer(1, 10))
        
        # Estilos para reporte técnico
        style_title_blue = ParagraphStyle('RepTitleBlue', parent=styles['Normal'], fontSize=15, leading=19, fontName="Helvetica-Bold", textColor=colors.HexColor("#0d47a1"))
        style_doc_info = ParagraphStyle('RepDocInfo', parent=styles['Normal'], fontSize=8, leading=11, fontName="Helvetica")
        style_section_title = ParagraphStyle('RepSecTitle', parent=styles['Normal'], fontSize=10, leading=13, fontName="Helvetica-Bold", textColor=colors.HexColor("#0d47a1"))
        style_table_header = ParagraphStyle('RepTableHeader', parent=styles['Normal'], fontSize=8, leading=10, fontName="Helvetica-Bold", textColor=colors.white, alignment=1)
        style_table_cell = ParagraphStyle('RepTableCell', parent=styles['Normal'], fontSize=8, leading=10, fontName="Helvetica", alignment=1)
        style_table_cell_bold = ParagraphStyle('RepTableCellBold', parent=styles['Normal'], fontSize=8, leading=10, fontName="Helvetica-Bold", alignment=1)
        
        # 1. Encabezado del Reporte Técnico
        story.append(Paragraph("SIGRAMA PLANTA METALES", style_title_blue))
        story.append(Spacer(1, 4))
        
        fecha_analisis_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        story.append(Paragraph(f"<b>Documento:</b> Reporte Técnico de Ingeniería de Calidad y Evaluación de Suministro", style_doc_info))
        story.append(Paragraph(f"<b>Fecha de Análisis:</b> {fecha_analisis_str}", style_doc_info))
        story.append(Paragraph(f"<b>Parámetro Comercial:</b> Desviación Ofertada por Proveedor en ±0.006\"", style_doc_info))
        story.append(Spacer(1, 10))
        
        # 2. Sección 1: Calibración del Muestreo
        story.append(Paragraph("1. Calibración del Muestreo por Unidad (Rollo por Rollo)", style_section_title))
        story.append(Spacer(1, 5))
        
        # Cálculos estadísticos para el atado
        if "All_Measurements" in row and isinstance(row["All_Measurements"], list) and len(row["All_Measurements"]) > 0:
            devs = [float(x) - esp_nom for x in row["All_Measurements"]]
        else:
            m1 = float(row.get('Espesor_Medido_1_in', esp_nom))
            m2 = float(row.get('Espesor_Medido_2_in', esp_nom))
            m3 = float(row.get('Espesor_Medido_3_in', esp_nom))
            devs = [m1 - esp_nom, m2 - esp_nom, m3 - esp_nom]
            
        mu = np.mean(devs)
        sigma = np.std(devs)
        if sigma < 0.0005:
            sigma = 0.0005
            
        esp_tol_min = float(row.get('Espesor_Tol_Min', -0.008))
        esp_tol_max = float(row.get('Espesor_Tol_Max', 0.008))
        p_below = stats.norm.cdf(esp_tol_min, loc=mu, scale=sigma)
        p_above = 1.0 - stats.norm.cdf(esp_tol_max, loc=mu, scale=sigma)
        p_out_pct = (p_below + p_above) * 100.0
        p_in_pct = 100.0 - p_out_pct
        
        # Calibre de SKU
        sku_str = str(row.get('SKU', ''))
        calibre_val = "N/A"
        for word in sku_str.split('-'):
            if word.isdigit():
                calibre_val = f"CAL {word}"
                break
        if calibre_val == "N/A":
            calibre_val = "CAL 14"
            
        espesor_promedio = esp_nom + mu
        riesgo_str = "ALTO RIESGO" if p_out_pct >= 5.0 else "RIESGO BAJO"
        
        placa = row.get("Placa")
        id_prov_lbl = str(row['ID_Atado_Proveedor'])
        if placa and pd.notna(placa):
            try:
                id_prov_lbl = f"{id_prov_lbl} - Placa {int(float(placa))}"
            except Exception:
                id_prov_lbl = f"{id_prov_lbl} - Placa {placa}"
                
        # Datos Tabla 1
        t1_data = [
            [Paragraph("Rollo / Atado", style_table_header), Paragraph("Material", style_table_header),
             Paragraph("Calibre", style_table_header), Paragraph("Espesor (in)", style_table_header),
             Paragraph("Riesgo %", style_table_header), Paragraph("Riesgo", style_table_header)],
            [Paragraph(id_prov_lbl, style_table_cell_bold),
             Paragraph(str(row.get('Tipo_Lamina', 'Decapada')), style_table_cell),
             Paragraph(calibre_val, style_table_cell),
             Paragraph(f"{espesor_promedio:.4f}\"", style_table_cell),
             Paragraph(f"{p_out_pct:.2f}%", style_table_cell_bold),
             Paragraph(riesgo_str, style_table_cell_bold)]
        ]
        t1 = Table(t1_data, colWidths=[110, 80, 70, 90, 80, 110])
        t1.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1f4e79")),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BACKGROUND', (5, 1), (5, 1), colors.HexColor("#ffcdd2" if riesgo_str == "ALTO RIESGO" else "#c8e6c9")),
            ('TEXTCOLOR', (5, 1), (5, 1), colors.HexColor("#b71c1c" if riesgo_str == "ALTO RIESGO" else "#1b5e20")),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t1)
        story.append(Spacer(1, 10))
        
        # 3. Sección 2: Análisis Estructurado y Clasificación
        story.append(Paragraph("2. Análisis Estructurado y Clasificación por Espesor Nominal", style_section_title))
        story.append(Spacer(1, 4))
        
        especificacion_text = f"<b>Especificación:</b> {row.get('Tipo_Lamina', 'Decapada')} - {calibre_val} - Espesor Teórico: {esp_nom:.3f}\" | Tolerancia Aceptable: ±{esp_tol_max:.3f}\""
        story.append(Paragraph(especificacion_text, style_doc_info))
        story.append(Spacer(1, 4))
        
        dictamen_val = "ACEPTADO" if row['Estatus_Calidad'] == 'Aceptado' else "NO ACEPTADO"
        dictamen_bg = "#c8e6c9" if dictamen_val == "ACEPTADO" else "#ffcdd2"
        dictamen_fg = "#1b5e20" if dictamen_val == "ACEPTADO" else "#b71c1c"
        
        # Datos Tabla 2
        t2_data = [
            [Paragraph("Número Rollo", style_table_header), Paragraph("Espesor Medido (in)", style_table_header),
             Paragraph("Desviación Real", style_table_header), Paragraph("Probabilidad de Fallo", style_table_header),
             Paragraph("Dictamen Final", style_table_header)],
            [Paragraph(id_prov_lbl, style_table_cell_bold),
             Paragraph(f"{espesor_promedio:.4f}\"", style_table_cell),
             Paragraph(f"{mu:+.6f}\"", style_table_cell),
             Paragraph(f"{p_out_pct:.2f}%", style_table_cell),
             Paragraph(f"<b>{dictamen_val}</b>", ParagraphStyle('DictStyle', parent=style_table_cell_bold, textColor=colors.HexColor(dictamen_fg)))]
        ]
        t2 = Table(t2_data, colWidths=[110, 110, 110, 100, 110])
        t2.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1f4e79")),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BACKGROUND', (4, 1), (4, 1), colors.HexColor(dictamen_bg)),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t2)
        story.append(Spacer(1, 10))
        
        # 4. Sección 3: Análisis de Distribución Probabilística
        story.append(Paragraph("3. Análisis de Distribución Probabilística por Especificación Técnica", style_section_title))
        story.append(Spacer(1, 4))
        
        if os.path.exists(img_prob_path):
            img_gauss_2 = Image(img_prob_path, width=540, height=200)
            img_gauss_2.hAlign = 'CENTER'
            story.append(img_gauss_2)
            
        story.append(Spacer(1, 8))
        
        # 5. Conclusión Dinámica
        if dictamen_val == "ACEPTADO":
            conclusion_text = f"<b>CONCLUSIÓN:</b> Tras evaluar las mediciones y la calibración del muestreo por unidad, se concluye que el atado/rollo <b>{id_prov_lbl}</b> CUMPLE satisfactoriamente con los criterios de inspección inicial y tolerancias dimensionales. Se clasifica con Dictamen Final <b>ACEPTADO</b> y Riesgo Bajo ({p_out_pct:.2f}%), autorizándose su liberación para producción."
        else:
            conclusion_text = f"<b>CONCLUSIÓN:</b> Tras evaluar las mediciones y la calibración del muestreo por unidad, se concluye que el atado/rollo <b>{id_prov_lbl}</b> NO CUMPLE con los límites dimensionales establecidos. Se clasifica con Dictamen Final <b>NO ACEPTADO</b> y Riesgo Alto ({p_out_pct:.2f}%), por lo que se procede al rechazo o retención del material."
            
        style_conclusion = ParagraphStyle('RepConclusion', parent=styles['Normal'], fontSize=8, leading=11, fontName="Helvetica")
        story.append(Paragraph(conclusion_text, style_conclusion))
        story.append(Spacer(1, 10))
        
        # 6. Bloque de Firma de Ing. Jesús Morales
        story.append(Table([[""]], colWidths=[540], rowHeights=[1], style=TableStyle([('LINEABOVE', (0,0), (-1,-1), 1, colors.HexColor("#757575")), ('BOTTOMPADDING', (0,0), (-1,-1), 0), ('TOPPADDING', (0,0), (-1,-1), 0)])))
        story.append(Spacer(1, 4))
        
        style_firma_name = ParagraphStyle('RepFirmaName', parent=styles['Normal'], fontSize=9, fontName="Helvetica-Bold", alignment=1)
        style_firma_job = ParagraphStyle('RepFirmaJob', parent=styles['Normal'], fontSize=8, fontName="Helvetica", alignment=1, textColor=colors.HexColor("#424242"))
        
        story.append(Paragraph("Ing. Jesús Morales", style_firma_name))
        story.append(Paragraph("Dir. Planta Metales | SIGRAMA PLANTA METALES", style_firma_job))
        
        # Agregar salto de página si hay más atados
        if idx < len(df_atados) - 1:
            story.append(PageBreak())
            
    def decorate_tag(canvas, doc):
        # Dibujar un rectángulo rojo elegante en el margen exterior para que parezca una tarjeta
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#D32F2F"))
        canvas.setLineWidth(2)
        canvas.rect(26, 26, 560, 740)
        
        # Código de formato SGC discreto en la parte inferior externa de la etiqueta
        canvas.setFont("Helvetica-Bold", 7)
        canvas.drawString(36, 15, "FO-MET-32")
        canvas.setFont("Helvetica", 6)
        canvas.setFillColor(colors.HexColor("#424242"))
        canvas.drawString(100, 15, "FORMATO DE TARJETA DE IDENTIFICACIÓN DE ATADO DE MATERIA PRIMA - SGC SIGRAMA")
        canvas.restoreState()
        
    doc.build(story, onFirstPage=decorate_tag, onLaterPages=decorate_tag)
    
    # Limpiar archivos temporales de imágenes para no saturar el disco
    for temp_file in temp_files:
        try:
            if os.path.exists(temp_file):
                os.remove(temp_file)
        except Exception as e:
            print(f"No se pudo eliminar archivo temporal {temp_file}: {e}")
            
    print("Etiquetas FO-MET-32 creadas en:", output_pdf_path)

def generar_grafica_multi_probabilidad(df_atados, output_path):
    n_atados = len(df_atados)
    fig, axs = plt.subplots(n_atados, 1, figsize=(7.5, 2.5 * n_atados))
    if n_atados == 1:
        axs = [axs]
        
    for i, (_, row) in enumerate(df_atados.iterrows()):
        ax = axs[i]
        esp_nom, esp_tol_min, esp_tol_max = obtener_valores_tolerancia(row)
        
        lsl = esp_tol_min
        usl = esp_tol_max
        
        if "All_Measurements" in row and isinstance(row["All_Measurements"], list) and len(row["All_Measurements"]) > 0:
            devs = [float(x) - esp_nom for x in row["All_Measurements"]]
        else:
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
        
        x = np.linspace(-0.015, 0.015, 500)
        y = stats.norm.pdf(x, loc=mu, scale=sigma)
        
        sku_str = str(row.get('SKU', ''))
        calibre_val = "N/A"
        for word in sku_str.split('-'):
            if word.isdigit():
                calibre_val = f"CAL {word}"
                break
        if calibre_val == "N/A":
            calibre_val = "CAL 14"
            
        tipo = row.get("Tipo_Lamina", "Decapada")
        id_prov = row.get("ID_Atado_Proveedor", "Atado")
        placa = row.get("Placa")
        if placa and pd.notna(placa):
            try:
                id_prov = f"{id_prov}-P{int(float(placa))}"
            except Exception:
                id_prov = f"{id_prov}-P{placa}"
                
        ax.plot(x, y, label=f"{id_prov} ({p_out_pct:.1f}%)", color="#1f77b4", linewidth=2)
        ax.axvspan(lsl, usl, color='#e8f5e9', alpha=0.5, label='Zona de Especificación' if i==0 else None)
        ax.axhline(0, color='black', linewidth=0.5)
        ax.axvline(0, color='green', linestyle='--', linewidth=1.5, label='Nominal' if i==0 else None)
        ax.axvline(lsl, color='red', linestyle=':', linewidth=1.5, label='LSL' if i==0 else None)
        ax.axvline(usl, color='red', linestyle=':', linewidth=1.5, label='USL' if i==0 else None)
        
        ax.set_title(f"{tipo} - {calibre_val} - Nominal: {esp_nom:.4f}\"", fontsize=9, fontweight='bold', color='#0d47a1', pad=4)
        ax.set_xlim(-0.015, 0.015)
        ax.grid(True, which='both', linestyle=':', color='lightgray', alpha=0.7)
        ax.legend(loc='upper right', fontsize=8)
        ax.tick_params(axis='both', which='major', labelsize=8)
        
    axs[-1].set_xlabel("Desviación Micrométrica Real (in)", fontsize=9)
    
    for ax in axs:
        ax.set_ylabel("Densidad", fontsize=8)
        
    plt.tight_layout()
    plt.savefig(output_path, dpi=200)
    plt.close()
    return output_path

def generar_pdf_reporte_tecnico_consolidado(folio, datos_reporte, df_atados, output_pdf_path):
    df_atados = consolidar_df_atados_para_pdf(df_atados)
    doc = SimpleDocTemplate(output_pdf_path, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()
    
    style_title_blue = ParagraphStyle('RepTitleBlueCons', parent=styles['Normal'], fontSize=15, leading=19, fontName="Helvetica-Bold", textColor=colors.HexColor("#0d47a1"))
    style_doc_info = ParagraphStyle('RepDocInfoCons', parent=styles['Normal'], fontSize=8, leading=11, fontName="Helvetica")
    style_section_title = ParagraphStyle('RepSecTitleCons', parent=styles['Normal'], fontSize=10, leading=13, fontName="Helvetica-Bold", textColor=colors.HexColor("#0d47a1"))
    style_table_header = ParagraphStyle('RepTableHeaderCons', parent=styles['Normal'], fontSize=8, leading=10, fontName="Helvetica-Bold", textColor=colors.white, alignment=1)
    style_table_cell = ParagraphStyle('RepTableCellCons', parent=styles['Normal'], fontSize=8, leading=10, fontName="Helvetica", alignment=1)
    style_table_cell_bold = ParagraphStyle('RepTableCellBoldCons', parent=styles['Normal'], fontSize=8, leading=10, fontName="Helvetica-Bold", alignment=1)
    
    # 1. Encabezado del Reporte Técnico
    story.append(Paragraph("SIGRAMA PLANTA METALES", style_title_blue))
    story.append(Spacer(1, 4))
    
    fecha_analisis_str = datos_reporte.get("Fecha", datetime.datetime.now().strftime("%d/%m/%Y"))
    hora_str = datos_reporte.get("Hora", datetime.datetime.now().strftime("%H:%M"))
    story.append(Paragraph(f"<b>Documento:</b> Reporte Técnico de Ingeniería de Calidad y Evaluación de Suministro", style_doc_info))
    story.append(Paragraph(f"<b>Fecha de Análisis:</b> {fecha_analisis_str} {hora_str}", style_doc_info))
    story.append(Paragraph(f"<b>Parámetro Comercial:</b> Desviación Ofertada por Proveedor en ±0.006\"", style_doc_info))
    story.append(Spacer(1, 10))
    
    # 2. Sección 1: Calibración del Muestreo por Unidad (Rollo por Rollo)
    story.append(Paragraph("1. Calibración del Muestreo por Unidad (Rollo por Rollo)", style_section_title))
    story.append(Spacer(1, 5))
    
    t1_data = [
        [Paragraph("Rollo", style_table_header), Paragraph("Material", style_table_header),
         Paragraph("Calibre", style_table_header), Paragraph("Espesor (in)", style_table_header),
         Paragraph("Riesgo %", style_table_header), Paragraph("Riesgo", style_table_header)]
    ]
    
    for _, row in df_atados.iterrows():
        esp_nom, esp_tol_min, esp_tol_max = obtener_valores_tolerancia(row)
        
        m1 = float(row.get('Espesor_Medido_1_in', esp_nom))
        m2 = float(row.get('Espesor_Medido_2_in', esp_nom))
        m3 = float(row.get('Espesor_Medido_3_in', esp_nom))
        
        devs = [m1 - esp_nom, m2 - esp_nom, m3 - esp_nom]
        mu = np.mean(devs)
        sigma = np.std(devs)
        if sigma < 0.0005:
            sigma = 0.0005
            
        p_below = stats.norm.cdf(esp_tol_min, loc=mu, scale=sigma)
        p_above = 1.0 - stats.norm.cdf(esp_tol_max, loc=mu, scale=sigma)
        p_out_pct = (p_below + p_above) * 100.0
        
        espesor_promedio = esp_nom + mu
        riesgo_str = "ALTO RIESGO" if p_out_pct >= 5.0 else "RIESGO BAJO"
        
        sku_str = str(row.get('SKU', ''))
        calibre_val = "N/A"
        for word in sku_str.split('-'):
            if word.isdigit():
                calibre_val = f"CAL {word}"
                break
        if calibre_val == "N/A":
            calibre_val = "CAL 14"
            
        tipo = row.get("Tipo_Lamina", "Decapada")
        
        t1_data.append([
            Paragraph(str(row['ID_Atado_Proveedor']), style_table_cell_bold),
            Paragraph(tipo, style_table_cell),
            Paragraph(calibre_val, style_table_cell),
            Paragraph(f"{espesor_promedio:.4f}\"", style_table_cell),
            Paragraph(f"{p_out_pct:.2f}%", style_table_cell_bold),
            Paragraph(riesgo_str, style_table_cell_bold)
        ])
        
    t1 = Table(t1_data, colWidths=[110, 80, 70, 90, 80, 110])
    
    t1_style = [
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1f4e79")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]
    
    for idx in range(1, len(t1_data)):
        r_val = t1_data[idx][5].text
        if "ALTO RIESGO" in r_val:
            t1_style.append(('BACKGROUND', (5, idx), (5, idx), colors.HexColor("#ffcdd2")))
            t1_style.append(('TEXTCOLOR', (5, idx), (5, idx), colors.HexColor("#b71c1c")))
        else:
            t1_style.append(('BACKGROUND', (5, idx), (5, idx), colors.HexColor("#c8e6c9")))
            t1_style.append(('TEXTCOLOR', (5, idx), (5, idx), colors.HexColor("#1b5e20")))
            
    t1.setStyle(TableStyle(t1_style))
    story.append(t1)
    story.append(Spacer(1, 15))
    
    # 3. Sección 2: Análisis Estructurado y Clasificación por Espesor Nominal
    story.append(Paragraph("2. Análisis Estructurado y Clasificación por Espesor Nominal", style_section_title))
    story.append(Spacer(1, 5))
    
    for _, row in df_atados.iterrows():
        esp_nom, esp_tol_min, esp_tol_max = obtener_valores_tolerancia(row)
        
        m1 = float(row.get('Espesor_Medido_1_in', esp_nom))
        m2 = float(row.get('Espesor_Medido_2_in', esp_nom))
        m3 = float(row.get('Espesor_Medido_3_in', esp_nom))
        
        devs = [m1 - esp_nom, m2 - esp_nom, m3 - esp_nom]
        mu = np.mean(devs)
        sigma = np.std(devs)
        if sigma < 0.0005:
            sigma = 0.0005
            
        p_below = stats.norm.cdf(esp_tol_min, loc=mu, scale=sigma)
        p_above = 1.0 - stats.norm.cdf(esp_tol_max, loc=mu, scale=sigma)
        p_out_pct = (p_below + p_above) * 100.0
        
        espesor_promedio = esp_nom + mu
        
        sku_str = str(row.get('SKU', ''))
        calibre_val = "N/A"
        for word in sku_str.split('-'):
            if word.isdigit():
                calibre_val = f"CAL {word}"
                break
        if calibre_val == "N/A":
            calibre_val = "CAL 14"
            
        tipo = row.get("Tipo_Lamina", "Decapada")
        
        especificacion_text = f"Especificación: {tipo} - {calibre_val} - Espesor Teórico: {esp_nom:.3f}\" | Tolerancia Aceptable: ±{esp_tol_max:.3f}\""
        story.append(Paragraph(especificacion_text, style_doc_info))
        story.append(Spacer(1, 4))
        
        dictamen_val = "ACEPTADO" if row.get('Estatus_Calidad', 'Aceptado') == 'Aceptado' else "NO ACEPTADO"
        dictamen_bg = "#c8e6c9" if dictamen_val == "ACEPTADO" else "#ffcdd2"
        dictamen_fg = "#1b5e20" if dictamen_val == "ACEPTADO" else "#b71c1c"
        
        t2_data = [
            [Paragraph("Número Rollo", style_table_header), Paragraph("Espesor Medido (in)", style_table_header),
             Paragraph("Desviación Real", style_table_header), Paragraph("Probabilidad de Fallo", style_table_header),
             Paragraph("Dictamen Final", style_table_header)],
            [Paragraph(str(row['ID_Atado_Proveedor']), style_table_cell_bold),
             Paragraph(f"{espesor_promedio:.4f}\"", style_table_cell),
             Paragraph(f"{mu:+.6f}\"", style_table_cell),
             Paragraph(f"{p_out_pct:.2f}%", style_table_cell),
             Paragraph(f"<b>{dictamen_val}</b>", ParagraphStyle('DictStyleCons', parent=style_table_cell_bold, textColor=colors.HexColor(dictamen_fg)))]
        ]
        t2 = Table(t2_data, colWidths=[110, 110, 110, 100, 110])
        t2.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1f4e79")),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BACKGROUND', (4, 1), (4, 1), colors.HexColor(dictamen_bg)),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t2)
        story.append(Spacer(1, 10))
        
    story.append(PageBreak())
    story.append(Spacer(1, 10))
    
    # 4. Sección 3: Análisis de Distribución Probabilística por Especificación Técnica
    story.append(Paragraph("3. Análisis de Distribución Probabilística por Especificación Técnica", style_section_title))
    story.append(Spacer(1, 5))
    
    folder_path = os.path.dirname(output_pdf_path)
    img_multi_path = os.path.join(folder_path, f"temp_multi_prob_{folio}.png")
    generar_grafica_multi_probabilidad(df_atados, img_multi_path)
    
    if os.path.exists(img_multi_path):
        img_height = 140 * len(df_atados)
        if img_height > 480:
            img_height = 480
        story.append(Image(img_multi_path, width=540, height=img_height))
        
    story.append(Spacer(1, 15))
    
    story.append(Table([[""]], colWidths=[540], rowHeights=[1], style=TableStyle([('LINEABOVE', (0,0), (-1,-1), 1, colors.HexColor("#757575")), ('BOTTOMPADDING', (0,0), (-1,-1), 0), ('TOPPADDING', (0,0), (-1,-1), 0)])))
    story.append(Spacer(1, 4))
    
    style_firma_name = ParagraphStyle('RepFirmaNameCons', parent=styles['Normal'], fontSize=9, fontName="Helvetica-Bold", alignment=1)
    style_firma_job = ParagraphStyle('RepFirmaJobCons', parent=styles['Normal'], fontSize=8, fontName="Helvetica", alignment=1, textColor=colors.HexColor("#424242"))
    
    story.append(Paragraph("Ing. Jesús Morales", style_firma_name))
    story.append(Paragraph("Dir. Planta Metales | SIGRAMA PLANTA METALES", style_firma_job))
    
    def decorate_report(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica-Bold", 7)
        canvas.drawString(36, 15, "REPORTE TÉCNICO")
        canvas.setFont("Helvetica", 6)
        canvas.setFillColor(colors.HexColor("#424242"))
        canvas.drawString(120, 15, "EVALUACIÓN DE SUMINISTRO Y ANÁLISIS GAUSSIANO - SIGRAMA PLANTA METALES")
        canvas.restoreState()
        
    doc.build(story, onFirstPage=decorate_report, onLaterPages=decorate_report)
    
    try:
        if os.path.exists(img_multi_path):
            os.remove(img_multi_path)
    except Exception as e:
        print(f"No se pudo eliminar archivo temporal {img_multi_path}: {e}")
        
    print("Reporte Técnico Consolidado Creado en:", output_pdf_path)

def generar_pdf_portada_dosier_fomet33(folio, datos_reporte, df_atados, output_pdf_path):
    """
    Genera la portada del dosier de calidad FO-MET-33.
    """
    df_atados = consolidar_df_atados_para_pdf(df_atados)
    doc = SimpleDocTemplate(output_pdf_path, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story = []
    styles = getSampleStyleSheet()
    
    style_tit = ParagraphStyle('CoverTit', parent=styles['Heading1'], fontSize=20, leading=24, alignment=1, textColor=colors.HexColor("#D32F2F"))
    style_sub = ParagraphStyle('CoverSub', parent=styles['Heading2'], fontSize=12, alignment=1, textColor=colors.HexColor("#757575"))
    style_normal_bold = ParagraphStyle('CoverNB', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=9)
    style_normal_text = ParagraphStyle('CoverNT', parent=styles['Normal'], fontSize=9)
    style_blanco_bold = ParagraphStyle('CoverWB', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=9)
    
    story.append(Spacer(1, 40))
    story.append(Paragraph("DOSIER DE CALIDAD DE RECEPCIÓN", style_tit))
    story.append(Spacer(1, 5))
    story.append(Paragraph("EXPEDIENTE DE TRAZABILIDAD Y LIBERACIÓN DE MATERIA PRIMA", style_sub))
    story.append(Spacer(1, 30))
    
    # Resumen General del Expediente
    t_header = Table([[Paragraph("RESUMEN DE CONTENIDO DEL DOSIER", style_blanco_bold)]], colWidths=[540])
    t_header.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#D32F2F")), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(t_header)
    
    # Extraer coladas y atados únicos
    coladas_lote = ", ".join(map(str, df_atados["Num_Colada"].dropna().unique().tolist()))
    total_atados = len(df_atados)
    peso_total_kg = int(df_atados["Peso_Total_Kg"].sum())
    
    datos_resumen = [
        [Paragraph("FOLIO DEL DOSIER:", style_normal_bold), Paragraph(f"<b>{folio}</b>", style_normal_bold),
         Paragraph("FECHA DE LIBERACIÓN:", style_normal_bold), Paragraph(str(datos_reporte.get('Fecha', '')), style_normal_text)],
        [Paragraph("PROVEEDOR:", style_normal_bold), Paragraph(str(datos_reporte.get('Proveedor', '')), style_normal_text),
         Paragraph("ORDEN DE COMPRA:", style_normal_bold), Paragraph(str(datos_reporte.get('Orden_Compra', '')), style_normal_text)],
        [Paragraph("CANTIDAD DE ATADOS:", style_normal_bold), Paragraph(f"{total_atados} bultos/atados", style_normal_text),
         Paragraph("PESO NETO TOTAL:", style_normal_bold), Paragraph(f"<b>{peso_total_kg:,} Kg</b>", style_normal_bold)],
        [Paragraph("HEATS / COLADAS:", style_normal_bold), Paragraph(coladas_lote, style_normal_text),
         Paragraph("INSPECTOR RESPONSABLE:", style_normal_bold), Paragraph(str(datos_reporte.get('Inspector', '')), style_normal_text)]
    ]
    t_resumen = Table(datos_resumen, colWidths=[120, 150, 130, 140])
    t_resumen.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#F5F5F5")),
        ('BACKGROUND', (2,0), (2,-1), colors.HexColor("#F5F5F5"))
    ]))
    story.append(t_resumen)
    story.append(Spacer(1, 30))
    
    # Índice del Expediente
    story.append(Paragraph("<b>CONTENIDO CERTIFICADO EN ESTE EXPEDIENTE:</b>", style_normal_bold))
    story.append(Spacer(1, 5))
    
    indice_texto = """
    1. <b>PORTADA E ÍNDICE DE DOSIER (FO-MET-33)</b> - Registro básico de trazabilidad del lote.<br/>
    2. <b>REPORTE CONSOLIDADO DE CALIDAD (FO-MET-31)</b> - Mediciones dimensionales reales y curvas de tolerancia.<br/>
    3. <b>TARJETAS DE IDENTIFICACIÓN (FO-MET-32)</b> - Tarjetas de identificación de los atados para control en almacén.<br/>
    4. <b>CERTIFICADO DE CALIDAD DE ACERÍA (MTR)</b> - Certificados de calidad originales emitidos por el molino proveedor.<br/>
    5. <b>ORDEN DE COMPRA DE SIGRAMA</b> - Documento comercial formal que solicitó el material.
    """
    story.append(Paragraph(indice_texto, style_normal_text))
    story.append(Spacer(1, 50))
    
    # Firmas formales
    datos_firmas = [
        [Paragraph("_____________________________<br/>INSPECTOR DE CALIDAD<br/>Generación y Validación", style_normal_bold),
         Paragraph("_____________________________<br/>SUPERVISOR DE CONTROL DE CALIDAD<br/>Revisión y Aprobación", style_normal_bold),
         Paragraph("_____________________________<br/>GERENCIA DE OPERACIONES/METALES<br/>Liberación de Producción", style_normal_bold)]
    ]
    t_firmas = Table(datos_firmas, colWidths=[180, 180, 180])
    t_firmas.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'TOP')
    ]))
    story.append(t_firmas)
    
    def decorate_portada(canvas, doc):
        draw_sigrama_sgc_decorations(canvas, doc, "FO-MET-33", "PORTADA DEL DOSIER DE CALIDAD")
        
    doc.build(story, onFirstPage=decorate_portada, onLaterPages=decorate_portada)
    print("Portada de Dosier FO-MET-33 creada en:", output_pdf_path)

# Helper styles
def style_blanco_bold_label(styles):
    return ParagraphStyle('WBL', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=9)

def style_blanco_bold_label_small(styles):
    return ParagraphStyle('WBLS', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=7)

# =============================================================================
# 4. FUSIÓN DINÁMICA DE PDFS CON PYPDF
# =============================================================================
def compilar_dosier_calidad(dossier_pdf_path, cover_pdf, report_pdf, tecnico_pdf, labels_pdf, certificados_paths=None, oc_paths=None):
    """
    Une todos los PDFs correspondientes para formar el Dosier de Calidad final unificado.
    """
    writer = PdfWriter()
    
    # 4.1 Añadir Portada (FO-MET-33)
    if os.path.exists(cover_pdf):
        reader = PdfReader(cover_pdf)
        for page in reader.pages:
            writer.add_page(page)
            
    # 4.2 Añadir Reporte Consolidado (FO-MET-31)
    if os.path.exists(report_pdf):
        reader = PdfReader(report_pdf)
        for page in reader.pages:
            writer.add_page(page)
            
    # 4.3 Añadir Reporte Técnico Consolidado (Evaluación de Suministro y Campanas)
    if tecnico_pdf and os.path.exists(tecnico_pdf):
        try:
            reader = PdfReader(tecnico_pdf)
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            print(f"⚠️ Error al unir reporte técnico: {e}")
            
    # 4.4 Añadir Etiquetas (FO-MET-32)
    if os.path.exists(labels_pdf):
        reader = PdfReader(labels_pdf)
        for page in reader.pages:
            writer.add_page(page)
            
    # 4.5 Añadir Certificados del Molino
    if certificados_paths:
        for path in certificados_paths:
            if path and os.path.exists(path):
                try:
                    reader = PdfReader(path)
                    for page in reader.pages:
                        writer.add_page(page)
                except Exception as e:
                    print(f"⚠️ Error al unir certificado {path}: {e}")
                    
    # 4.6 Añadir Órdenes de Compra de Sigrama
    if oc_paths:
        for path in oc_paths:
            if path and os.path.exists(path):
                try:
                    reader = PdfReader(path)
                    for page in reader.pages:
                        writer.add_page(page)
                except Exception as e:
                    print(f"⚠️ Error al unir orden de compra {path}: {e}")
                    
    # Guardar dosier unificado
    with open(dossier_pdf_path, "wb") as f_out:
        writer.write(f_out)
    print("Dosier de Calidad compilado automáticamente en:", dossier_pdf_path)


def generar_pdf_solo_etiquetas(folio, df_atados, output_pdf_path):
    """
    Genera un PDF que contiene únicamente las tarjetas de identificación de atado (FO-MET-32)
    para cada uno de los atados, una tarjeta por página (tamaño Carta).
    """
    df_atados = consolidar_df_atados_para_pdf(df_atados)
    doc = SimpleDocTemplate(output_pdf_path, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()
    
    # Lista de archivos temporales creados para borrar al final
    temp_files = []
    folder_path = os.path.dirname(output_pdf_path)
    
    style_tit_label = ParagraphStyle('TL_Lab_Solo', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=14)
    style_cell_header = ParagraphStyle('CH_Lab_Solo', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=8, textColor=colors.HexColor("#757575"))
    style_cell_value = ParagraphStyle('CV_Lab_Solo', parent=styles['Normal'], fontSize=9, fontName="Helvetica")
    style_cell_value_bold = ParagraphStyle('CVB_Lab_Solo', parent=styles['Normal'], fontSize=10, fontName="Helvetica-Bold")
    style_cell_value_white_bold = ParagraphStyle('CVWB_Lab_Solo', parent=styles['Normal'], fontSize=10, fontName="Helvetica-Bold", textColor=colors.white, alignment=1)
    
    num_atados = len(df_atados)
    for idx_row, (idx, row) in enumerate(df_atados.iterrows()):
        story.append(Spacer(1, 10))
        
        # 1. Cabecera de la Etiqueta
        if os.path.exists(LOGO_PATH):
            logotipo_header = Image(LOGO_PATH, width=150, height=27)
        else:
            logotipo_header = Paragraph("<b>INDUSTRIA SIGRAMA</b>", style_cell_value)
            
        header_table_data = [
            [logotipo_header, Paragraph("ATADO DE MATERIA PRIMA", style_tit_label)]
        ]
        header_table = Table(header_table_data, colWidths=[200, 340])
        header_table.setStyle(TableStyle([
            ('BACKGROUND', (1,0), (1,0), colors.HexColor("#D32F2F")),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (0,0), 'LEFT'),
            ('ALIGN', (1,0), (1,0), 'CENTER'),
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#D32F2F"))
        ]))
        story.append(header_table)
        
        # Sub-título Identificación
        story.append(Table([[Paragraph("IDENTIFICACIÓN DEL ATADO", style_blanco_bold_label(styles))]], colWidths=[540], 
                           style=TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#D32F2F")), ('ALIGN', (0,0), (-1,-1), 'CENTER')])))
        
        # 2. Bloque de Identificación del Atado (Izquierda: Tabla, Derecha: Imagen + Observaciones)
        
        # Left table (fields)
        ident_left_data = [
            [Paragraph("📅 FECHA DE RECEPCIÓN", style_cell_header), Paragraph(datetime.date.today().strftime("%d/%m/%Y"), style_cell_value_bold)],
            [Paragraph("🚛 PROVEEDOR", style_cell_header), Paragraph(str(row.get('Proveedor', 'Ternium')), style_cell_value_bold)],
            [Paragraph("📦 TIPO DE MATERIAL", style_cell_header), Paragraph(str(row.get('Tipo_Lamina', 'Decapada')), style_cell_value_bold)],
            [Paragraph("✨ ACABADO", style_cell_header), Paragraph(str(row.get('Acabado', 'Pasivado' if row.get('Tipo_Lamina')=='Galvanizada' else 'Aceitado')), style_cell_value)],
            [Paragraph("✔️ GRADO DE ACERO", style_cell_header), Paragraph(str(row['Grado_Acero']), style_cell_value_bold)],
            [Paragraph("🏷️ NÚMERO DE ATADO (INTERNO)", style_cell_header), Paragraph(f"<b>{row['ID_Atado']}</b>", style_cell_value_bold)]
        ]
        
        ident_left_table = Table(ident_left_data, colWidths=[150, 120])
        ident_left_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#D32F2F")),
            ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#F5F5F5")),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4)
        ]))
        
        # Right table (Image, OBSERVACIONES banner, text)
        observaciones_atado = str(row.get('Observaciones', '')).strip()
        if observaciones_atado == 'nan' or not observaciones_atado or observaciones_atado == 'Sin observaciones registradas.':
            observaciones_atado = "Excelente estado superficial."
            
        stack_img_path = os.path.join(BASE_DIR, "stack_metal_sheets.png")
        if os.path.exists(stack_img_path):
            stack_img = Image(stack_img_path, width=120, height=80)
            stack_img.hAlign = 'CENTER'
        else:
            stack_img = Paragraph("<b>[IMAGEN DEL ATADO]</b>", style_cell_value)
            
        ident_right_data = [
            [stack_img],
            [Paragraph("OBSERVACIONES", style_blanco_bold_label(styles))],
            [Paragraph(observaciones_atado, style_cell_value)]
        ]
        
        ident_right_table = Table(ident_right_data, colWidths=[270])
        ident_right_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#D32F2F")),
            ('BACKGROUND', (0,1), (0,1), colors.HexColor("#D32F2F")),
            ('ALIGN', (0,0), (0,0), 'CENTER'),
            ('ALIGN', (0,1), (0,1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,2), (0,2), 8),
            ('TOPPADDING', (0,2), (0,2), 8)
        ]))
        
        ident_container_data = [[ident_left_table, ident_right_table]]
        ident_container = Table(ident_container_data, colWidths=[270, 270])
        ident_container.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0)
        ]))
        story.append(ident_container)
        story.append(Spacer(1, 10))
        
        # 3. Bloque de Tres Columnas: Especificaciones, Control de Hojas, Trazabilidad
        esp_nom, esp_tol_min, esp_tol_max = obtener_valores_tolerancia(row)
        esp_min = esp_nom + esp_tol_min
        esp_max = esp_nom + esp_tol_max
        
        # Tabla de Especificaciones (Izquierda)
        spec_rows = [
            [Paragraph("ESPECIFICACIONES DEL MATERIAL", style_blanco_bold_label_small(styles)), "", "", ""],
            [Paragraph("CARACTERÍSTICA", style_cell_header), Paragraph("VALOR REAL", style_cell_header), Paragraph("TOL. MÍN", style_cell_header), Paragraph("TOL. MÁX", style_cell_header)],
            [Paragraph("Espesor Real", style_cell_header), Paragraph(f"{float(row['Espesor_Medido_1_in']):.4f}", style_cell_value_bold), Paragraph(f"{esp_min:.4f}", style_cell_value), Paragraph(f"{esp_max:.4f}", style_cell_value)],
            [Paragraph("Ancho (in)", style_cell_header), Paragraph(f"{float(row['Ancho_Medido_in']):.2f}", style_cell_value), Paragraph(f"{row.get('Ancho_Tol_Min_Val', 47.8):.2f}", style_cell_value), Paragraph(f"{row.get('Ancho_Tol_Max_Val', 48.2):.2f}", style_cell_value)],
            [Paragraph("Largo (in)", style_cell_header), Paragraph(f"{float(row['Largo_Medido_in']):.2f}", style_cell_value), Paragraph("-", style_cell_value), Paragraph("-", style_cell_value)],
            [Paragraph("Cant. Hojas", style_cell_header), Paragraph(str(int(row['Cantidad_Hojas'])), style_cell_value), Paragraph("-", style_cell_value), Paragraph("-", style_cell_value)],
            [Paragraph("Peso Total (Kg)", style_cell_header), Paragraph(f"{int(row['Peso_Total_Kg'])}", style_cell_value_bold), Paragraph("-", style_cell_value), Paragraph("-", style_cell_value)],
            [Paragraph("Peso Total (Lb)", style_cell_header), Paragraph(f"{int(row['Peso_Total_Lb'])}", style_cell_value), Paragraph("-", style_cell_value), Paragraph("-", style_cell_value)]
        ]
        spec_table = Table(spec_rows, colWidths=[75, 45, 30, 30])
        spec_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#D32F2F")),
            ('SPAN', (0,0), (3,0)),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")),
            ('BACKGROUND', (0,1), (-1,1), colors.HexColor("#F5F5F5")),
            ('ALIGN', (1,1), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 7)
        ]))
        
        # Tabla de Control de Hojas (Medio)
        control_rows = [
            [Paragraph("CONTROL DE HOJAS", style_blanco_bold_label_small(styles)), "", ""],
            [Paragraph(f"RECIBIDAS (RECEPCIÓN): <b>{int(row['Cantidad_Hojas'])}</b> HOJAS", style_cell_value_bold), "", ""],
            [Paragraph("REGISTRO DE RETIROS", style_cell_header), "", ""],
            [Paragraph("FECHA RETIRO", style_cell_header), Paragraph("CANTIDAD RETIRADA", style_cell_header), Paragraph("RESTANTE ALMACÉN", style_cell_header)],
            ["", "", ""],
            ["", "", ""],
            ["", "", ""],
            ["", "", ""]
        ]
        control_table = Table(control_rows, colWidths=[55, 55, 70])
        control_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#D32F2F")),
            ('SPAN', (0,0), (2,0)),
            ('SPAN', (0,1), (2,1)),
            ('SPAN', (0,2), (2,2)),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")),
            ('BACKGROUND', (0,2), (-1,3), colors.HexColor("#F5F5F5")),
            ('ALIGN', (0,3), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('HEIGHT', (0,4), (-1,-1), 16)
        ]))
        
        # Tabla de Trazabilidad (Derecha)
        status_value = str(row['Estatus_Calidad']).upper()
        placa = row.get("Placa")
        id_prov_lbl = str(row['ID_Atado_Proveedor'])
        if placa and pd.notna(placa):
            try:
                id_prov_lbl = f"{id_prov_lbl} - Placa {int(float(placa))}"
            except Exception:
                id_prov_lbl = f"{id_prov_lbl} - Placa {placa}"
        traz_rows = [
            [Paragraph("TRAZABILIDAD", style_blanco_bold_label_small(styles)), ""],
            [Paragraph("🔥 NÚMERO DE COLADA", style_cell_header), Paragraph(str(row['Num_Colada']), style_cell_value_bold)],
            [Paragraph("📦 LOTE / HEAT", style_cell_header), Paragraph(str(row['Lote_Heat']), style_cell_value)],
            [Paragraph("🏷️ ATADO PROVEEDOR", style_cell_header), Paragraph(id_prov_lbl, style_cell_value)],
            [Paragraph("🏢 UBICACIÓN ALMACÉN", style_cell_header), Paragraph(str(row['Ubicacion_Almacen']), style_cell_value_bold)],
            [Paragraph("✔️ ESTATUS DE CALIDAD", style_cell_header), Paragraph(status_value, style_cell_value_white_bold)]
        ]
        traz_table = Table(traz_rows, colWidths=[100, 80])
        traz_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#D32F2F")),
            ('SPAN', (0,0), (1,0)),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")),
            ('BACKGROUND', (0,1), (0,-1), colors.HexColor("#F5F5F5")),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BACKGROUND', (1, 5), (1, 5), colors.HexColor("#2E7D32" if row['Estatus_Calidad']=='Aceptado' else "#C62828")),
            ('ALIGN', (1, 5), (1, 5), 'CENTER')
        ]))
        
        # Compilar las tres tablas en paralelo
        three_cols_table = Table([[spec_table, control_table, traz_table]], colWidths=[180, 180, 180])
        three_cols_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0)
        ]))
        story.append(three_cols_table)
        story.append(Spacer(1, 10))
        
        # 4. Sección de Manejo y Conservación
        story.append(Table([[Paragraph("MANEJO Y CONSERVACIÓN", style_blanco_bold_label_small(styles))]], colWidths=[540],
                           style=TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#D32F2F")), ('ALIGN', (0,0), (-1,-1), 'CENTER')])))
        
        manejo_data = [
            [Paragraph("<b>🚜 MANEJAR CON EQUIPO ADECUADO</b><br/>Usar montacargas de capacidad apropiada.", style_cell_value),
             Paragraph("<b>☔ PROTEGER DE LA HUMEDAD</b><br/>Almacenar en área seca para evitar oxidación.", style_cell_value),
             Paragraph("<b>🚫 NO APOYAR OBJETOS SOBRE EL ATADO</b><br/>No colocar peso directo sobre el atado.", style_cell_value)]
        ]
        manejo_table = Table(manejo_data, colWidths=[180, 180, 180])
        manejo_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#D32F2F")),
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F5F5F5")),
            ('VALIGN', (0,0), (-1,-1), 'TOP')
        ]))
        story.append(manejo_table)
        
        # Generar gráfica de campana de Gauss para ponerla en la parte baja de la tarjeta
        img_prob_path = generar_grafica_individual_probabilidad(row, folder_path)
        temp_files.append(img_prob_path)
        
        if os.path.exists(img_prob_path):
            story.append(Spacer(1, 8))
            img_gauss = Image(img_prob_path, width=420, height=160)
            img_gauss.hAlign = 'CENTER'
            story.append(img_gauss)
        
        # Salto de página para el siguiente atado si no es el último
        if idx_row < num_atados - 1:
            story.append(PageBreak())
            
    def decorate_tag(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#D32F2F"))
        canvas.setLineWidth(2)
        canvas.rect(26, 26, 560, 740)
        
        canvas.setFont("Helvetica-Bold", 7)
        canvas.drawString(36, 15, "FO-MET-32")
        canvas.setFont("Helvetica", 6)
        canvas.setFillColor(colors.HexColor("#424242"))
        canvas.drawString(100, 15, "FORMATO DE TARJETA DE IDENTIFICACIÓN DE ATADO DE MATERIA PRIMA - SGC SIGRAMA")
        canvas.restoreState()
        
    doc.build(story, onFirstPage=decorate_tag, onLaterPages=decorate_tag)
    
    # Limpiar archivos temporales de imágenes para no saturar el disco
    for temp_file in temp_files:
        try:
            if os.path.exists(temp_file):
                os.remove(temp_file)
        except Exception as e:
            print(f"No se pudo eliminar archivo temporal {temp_file}: {e}")
            
    print("PDF de Solo Etiquetas creado en:", output_pdf_path)


def generar_pdf_reporte_dashboard(filtros, okr_data, df_rep_filtered, dict_acep, output_pdf_path):
    """
    Genera un archivo PDF con el reporte del Dashboard, incluyendo
    los filtros activos, el resumen de OKRs y la tabla de recepciones recientes.
    """
    doc = SimpleDocTemplate(output_pdf_path, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story = []
    styles = getSampleStyleSheet()
    
    # Estilos de texto
    style_blanco_bold = ParagraphStyle('WB_Dash', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=8)
    style_normal_bold = ParagraphStyle('NB_Dash', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=8)
    style_normal_text = ParagraphStyle('NT_Dash', parent=styles['Normal'], fontSize=7.5)
    style_title = ParagraphStyle('T_Dash', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=13, textColor=colors.HexColor("#D32F2F"), spaceAfter=8)
    style_section = ParagraphStyle('S_Dash', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=10, textColor=colors.HexColor("#0D47A1"), spaceBefore=10, spaceAfter=5)
    
    story.append(Spacer(1, 5))
    story.append(Paragraph("REPORTE EJECUTIVO DE CONTROL DE CALIDAD - DASHBOARD", style_title))
    story.append(Paragraph("Este informe detalla el desempeño de calidad en recepción de materia prima (Lotes y Atados de acero) evaluado mediante la estructura de OKRs establecida en el SGC.", style_normal_text))
    story.append(Spacer(1, 10))
    
    # --- PANEL: FILTROS APLICADOS ---
    t_filtros_header = Table([[Paragraph("FILTROS DE VISUALIZACIÓN ACTIVOS", style_blanco_bold)]], colWidths=[540])
    t_filtros_header.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#757575")), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(t_filtros_header)
    
    datos_filtros = [
        [Paragraph("RANGO FECHAS:", style_normal_bold), Paragraph(f"{filtros.get('fecha_inicio', 'N/D')} a {filtros.get('fecha_fin', 'N/D')}", style_normal_text),
         Paragraph("PROVEEDOR FILTRO:", style_normal_bold), Paragraph(str(filtros.get('proveedor', 'Todos')), style_normal_text)],
        [Paragraph("MATERIAL FILTRO:", style_normal_bold), Paragraph(str(filtros.get('material', 'Todos')), style_normal_text),
         Paragraph("CALIBRE FILTRO:", style_normal_bold), Paragraph(str(filtros.get('calibre', 'Todos')), style_normal_text)]
    ]
    t_filtros = Table(datos_filtros, colWidths=[120, 150, 120, 150])
    t_filtros.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#F5F5F5")),
        ('BACKGROUND', (2,0), (2,-1), colors.HexColor("#F5F5F5"))
    ]))
    story.append(t_filtros)
    story.append(Spacer(1, 10))
    
    # --- RESULTADOS DE OKRS ---
    story.append(Paragraph("🎯 CONTROL DE OBJETIVOS Y RESULTADOS CLAVE (OKRs)", style_section))
    
    # OKR 1
    t_okr1_header = Table([[Paragraph("OKR 1: CONFORMIDAD Y CALIDAD DE MATERIA PRIMA (META ≥ 95%)", style_blanco_bold)]], colWidths=[540])
    t_okr1_header.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#0D47A1")), ('ALIGN', (0,0), (-1,-1), 'LEFT')]))
    story.append(t_okr1_header)
    
    tasa_atados = okr_data.get("tasa_atados", 100.0)
    tasa_lotes = okr_data.get("tasa_lotes", 100.0)
    tasa_peso = okr_data.get("tasa_peso", 100.0)
    
    atados_val = f"{tasa_atados:.1f}%"
    atados_delta = f"{tasa_atados - 95.0:+.1f}% vs Meta (95%)"
    atados_status = "CUMPLE" if tasa_atados >= 95.0 else "FUERA META"
    atados_color = "#2E7D32" if tasa_atados >= 95.0 else "#C62828"
    
    lotes_val = f"{tasa_lotes:.1f}%"
    lotes_delta = f"{tasa_lotes - 90.0:+.1f}% vs Meta (90%)"
    lotes_status = "CUMPLE" if tasa_lotes >= 90.0 else "FUERA META"
    lotes_color = "#2E7D32" if tasa_lotes >= 90.0 else "#C62828"
    
    peso_val = f"{tasa_peso:.1f}%"
    peso_delta = f"{tasa_peso - 95.0:+.1f}% vs Meta (95%)"
    peso_status = "CUMPLE" if tasa_peso >= 95.0 else "FUERA META"
    peso_color = "#2E7D32" if tasa_peso >= 95.0 else "#C62828"
    
    datos_okr1 = [
        [Paragraph("RESULTADO CLAVE (KR)", style_normal_bold), Paragraph("VALOR REAL", style_normal_bold), Paragraph("DESVIACIÓN VS META", style_normal_bold), Paragraph("ESTADO", style_normal_bold)],
        [Paragraph("<b>KR 1.1:</b> Conformidad de Atados (Meta: ≥95%)", style_normal_text), Paragraph(atados_val, style_normal_text), Paragraph(atados_delta, style_normal_text), Paragraph(f"<b><font color='{atados_color}'>{atados_status}</font></b>", style_normal_text)],
        [Paragraph("<b>KR 1.2:</b> Aceptación de Lotes (Meta: ≥90%)", style_normal_text), Paragraph(lotes_val, style_normal_text), Paragraph(lotes_delta, style_normal_text), Paragraph(f"<b><font color='{lotes_color}'>{lotes_status}</font></b>", style_normal_text)],
        [Paragraph("<b>KR 1.3:</b> Eficiencia de Peso Conforme (Meta: ≥95%)", style_normal_text), Paragraph(peso_val, style_normal_text), Paragraph(peso_delta, style_normal_text), Paragraph(f"<b><font color='{peso_color}'>{peso_status}</font></b>", style_normal_text)]
    ]
    t_okr1 = Table(datos_okr1, colWidths=[220, 90, 130, 100])
    t_okr1.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E0E0E0")),
        ('ALIGN', (1,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.append(t_okr1)
    story.append(Spacer(1, 8))
    
    # OKR 2
    t_okr2_header = Table([[Paragraph("OKR 2: EFICIENCIA DE ABASTECIMIENTO Y CONTROL", style_blanco_bold)]], colWidths=[540])
    t_okr2_header.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#0D47A1")), ('ALIGN', (0,0), (-1,-1), 'LEFT')]))
    story.append(t_okr2_header)
    
    datos_okr2 = [
        [Paragraph("RESULTADO CLAVE (KR)", style_normal_bold), Paragraph("VALOR REGISTRADO", style_normal_bold), Paragraph("DESCRIPCIÓN / COBERTURA", style_normal_bold)],
        [Paragraph("<b>KR 2.1:</b> Atados Controlados", style_normal_text), Paragraph(f"{okr_data.get('total_atados', 0)} Atados", style_normal_text), Paragraph("Rollos inspeccionados físicamente y medidos en espesores.", style_normal_text)],
        [Paragraph("<b>KR 2.2:</b> Acero Recibido", style_normal_text), Paragraph(f"{okr_data.get('peso_total_kg', 0):,.0f} Kg", style_normal_text), Paragraph(f"Peso bruto de materia prima controlada ({okr_data.get('peso_total_lb', 0):,.0f} Lb).", style_normal_text)],
        [Paragraph("<b>KR 2.3:</b> Lotes Procesados", style_normal_text), Paragraph(f"{okr_data.get('total_lotes', 0)} Lotes", style_normal_text), Paragraph(f"Aceptados: {okr_data.get('lotes_aceptados', 0)} | Reclamados/Rechazados: {okr_data.get('lotes_rechazados', 0)}", style_normal_text)]
    ]
    t_okr2 = Table(datos_okr2, colWidths=[220, 120, 200])
    t_okr2.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E0E0E0")),
        ('ALIGN', (1,0), (1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.append(t_okr2)
    story.append(Spacer(1, 10))
    
    # --- TABLA DE RECEPCIONES RECIENTES ---
    story.append(Paragraph("📋 LISTADO DE RECEPCIONES INCLUIDAS EN EL PERIODO", style_section))
    
    tabla_recepciones = [[
        Paragraph("FOLIO", style_blanco_bold),
        Paragraph("FECHA", style_blanco_bold),
        Paragraph("PROVEEDOR", style_blanco_bold),
        Paragraph("PO / FACTURA", style_blanco_bold),
        Paragraph("ACEPTACIÓN", style_blanco_bold),
        Paragraph("ESTATUS", style_blanco_bold)
    ]]
    
    for _, row in df_rep_filtered.iterrows():
        folio_val = row["Folio"]
        fecha_val = row["Fecha"]
        prov_val = row["Proveedor"]
        po_fact = f"{row['Orden_Compra']} / {row['Factura_Remision']}"
        acep_val = dict_acep.get(folio_val, "0/0 (0.0%)")
        estatus_val = row["Estatus_General"]
        
        status_color = "#2E7D32" if estatus_val == "Aceptado" else ("#FBC02D" if estatus_val == "Condicionado" else "#C62828")
        status_html = f"<b><font color='{status_color}'>{estatus_val}</font></b>"
        
        tabla_recepciones.append([
            Paragraph(folio_val, style_normal_text),
            Paragraph(fecha_val, style_normal_text),
            Paragraph(prov_val, style_normal_text),
            Paragraph(po_fact, style_normal_text),
            Paragraph(acep_val, style_normal_text),
            Paragraph(status_html, style_normal_text)
        ])
        
    t_recepciones = Table(tabla_recepciones, colWidths=[90, 70, 130, 110, 80, 60])
    t_recepciones.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#757575")),
        ('ALIGN', (0,0), (1,-1), 'CENTER'),
        ('ALIGN', (4,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.append(t_recepciones)
    
    # Firmas al pie
    story.append(Spacer(1, 20))
    datos_firmas = [
        [Paragraph("_____________________________<br/>ELABORADO POR CALIDAD", style_normal_bold),
         Paragraph("_____________________________<br/>APROBADO POR OPERACIONES", style_normal_bold)]
    ]
    t_firmas = Table(datos_firmas, colWidths=[270, 270])
    t_firmas.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'TOP')
    ]))
    story.append(t_firmas)
    
    def decorate(canvas, doc):
        draw_sigrama_sgc_decorations(canvas, doc, "FO-SGC-03", "REPORTE DE ANALÍTICAS Y OKRS DE CALIDAD")
        
    doc.build(story, onFirstPage=decorate, onLaterPages=decorate)
    print("Reporte de Dashboard PDF creado en:", output_pdf_path)

def generar_pdf_catalogo_skus(df_skus, output_pdf_path):
    """
    Genera un reporte PDF formal en formato horizontal (Landscape)
    con el catálogo de SKUs y sus tolerancias configuradas.
    """
    from reportlab.lib.pagesizes import letter, landscape
    doc = SimpleDocTemplate(
        output_pdf_path, 
        pagesize=landscape(letter), 
        leftMargin=36, 
        rightMargin=36, 
        topMargin=80, 
        bottomMargin=50
    )
    story = []
    styles = getSampleStyleSheet()
    
    style_blanco_bold = ParagraphStyle('WB_Sku', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=8)
    style_normal_bold = ParagraphStyle('NB_Sku', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=8)
    style_normal_text = ParagraphStyle('NT_Sku', parent=styles['Normal'], fontSize=7.5, alignment=1)
    
    story.append(Spacer(1, 10))
    
    # Tabla de SKUs
    # Cabecera de la tabla
    tabla_data = [[
        Paragraph("SKU", style_blanco_bold),
        Paragraph("DESCRIPCIÓN", style_blanco_bold),
        Paragraph("TIPO LÁMINA", style_blanco_bold),
        Paragraph("GRADO ACERO", style_blanco_bold),
        Paragraph("ESPESOR NOMINAL (in)", style_blanco_bold),
        Paragraph("TOL. ESPESOR (in)", style_blanco_bold),
        Paragraph("ANCHO NOMINAL (in)", style_blanco_bold),
        Paragraph("LARGO NOMINAL (in)", style_blanco_bold),
        Paragraph("ZINC MÍN (oz/ft²)", style_blanco_bold),
        Paragraph("DUREZA MÁX (HRB)", style_blanco_bold),
        Paragraph("ACEITADO REQ.", style_blanco_bold)
    ]]
    
    for _, row in df_skus.iterrows():
        esp_nom = float(row.get('Espesor_Nominal_in', 0))
        esp_min = float(row.get('Espesor_Tolerancia_Min_in', 0))
        esp_max = float(row.get('Espesor_Tolerancia_Max_in', 0))
        
        tol_esp = f"{esp_min:+.4f} / {esp_max:+.4f}"
        
        zinc_val = float(row.get('Zinc_Min_oz_ft2', 0))
        zinc_str = f"{zinc_val:.2f}" if zinc_val > 0 else "-"
        
        dureza_val = float(row.get('Dureza_Max_HRB', 0))
        dureza_str = f"{int(dureza_val)}" if dureza_val > 0 else "-"
        
        tabla_data.append([
            Paragraph(str(row.get('SKU', '')), style_normal_text),
            Paragraph(str(row.get('Nombre', '')), style_normal_text),
            Paragraph(str(row.get('Tipo_Lamina', '')), style_normal_text),
            Paragraph(str(row.get('Grado_Acero', '')), style_normal_text),
            Paragraph(f"{esp_nom:.4f}", style_normal_text),
            Paragraph(tol_esp, style_normal_text),
            Paragraph(f"{float(row.get('Ancho_Nominal_in', 0)):.2f}", style_normal_text),
            Paragraph(f"{float(row.get('Largo_Nominal_in', 0)):.2f}", style_normal_text),
            Paragraph(zinc_str, style_normal_text),
            Paragraph(dureza_str, style_normal_text),
            Paragraph(str(row.get('Aceitado_Requerido', 'N/D')), style_normal_text)
        ])
        
    t_skus = Table(tabla_data, colWidths=[75, 120, 65, 90, 60, 65, 50, 50, 45, 50, 50])
    t_skus.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#757575")),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.append(t_skus)
    
    def decorate_landscape(canvas, doc):
        canvas.saveState()
        
        # Franja superior roja Sigrama
        canvas.setFillColor(colors.HexColor("#D32F2F"))
        canvas.rect(36, 560, 720, 4, fill=1, stroke=0)
        
        # Logotipo
        if os.path.exists(LOGO_PATH):
            try:
                canvas.drawImage(LOGO_PATH, 36, 570, width=120, height=22, mask='auto')
            except Exception:
                canvas.setFont("Helvetica-Bold", 12)
                canvas.drawString(36, 570, "INDUSTRIA SIGRAMA")
        else:
            canvas.setFont("Helvetica-Bold", 12)
            canvas.drawString(36, 570, "INDUSTRIA SIGRAMA")
            
        # Marcador superior derecho
        canvas.setFont("Helvetica-Bold", 11)
        canvas.setFillColor(colors.HexColor("#D32F2F"))
        canvas.drawRightString(756, 580, "FO-MET-34")
        
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.black)
        canvas.drawRightString(756, 568, "Revisión 01")
        
        # Título Central
        canvas.setFont("Helvetica-Bold", 12)
        canvas.drawCentredString(396, 570, "CATÁLOGO DE PARÁMETROS Y TOLERANCIAS DE MATERIA PRIMA")
        
        # Fecha
        canvas.setFont("Helvetica-Bold", 8)
        months_es = {
            1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
            7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
        }
        today = datetime.date.today()
        fecha_hoy = f"{today.day} de {months_es[today.month]} de {today.year}"
        canvas.drawString(36, 547, f"Fecha de Emisión: {fecha_hoy}")
        
        # Pie de Página Legal (Landscape)
        canvas.setStrokeColor(colors.HexColor("#D32F2F"))
        canvas.setLineWidth(1)
        canvas.line(36, 35, 36, 15)
        canvas.setFont("Helvetica-Bold", 7)
        canvas.drawString(42, 27, "FO-SGC-02")
        canvas.setFont("Helvetica", 6)
        canvas.setFillColor(colors.HexColor("#424242"))
        texto_legal = "PROHIBIDA LA REPRODUCCIÓN TOTAL O PARCIAL, POR CUALQUIER MEDIO O PROCEDIMIENTO, SIN AUTORIZACIÓN DE INDUSTRIA SIGRAMA S.A. DE C.V."
        canvas.drawString(95, 27, texto_legal)
        
        canvas.restoreState()
        
    doc.build(story, onFirstPage=decorate_landscape, onLaterPages=decorate_landscape)
    print("PDF del Catálogo de SKUs creado.")

def generar_pdf_consulta_historial(filtros, df_rep_filtered, dict_acep, output_pdf_path):
    """
    Genera un PDF con el reporte de la consulta de historial actual,
    mostrando los filtros aplicados y la tabla de registros que coinciden.
    """
    doc = SimpleDocTemplate(output_pdf_path, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story = []
    styles = getSampleStyleSheet()
    
    style_blanco_bold = ParagraphStyle('WB_Cons', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=8)
    style_normal_bold = ParagraphStyle('NB_Cons', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=8)
    style_normal_text = ParagraphStyle('NT_Cons', parent=styles['Normal'], fontSize=7.5)
    style_title = ParagraphStyle('T_Cons', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=13, textColor=colors.HexColor("#D32F2F"), spaceAfter=8)
    style_section = ParagraphStyle('S_Cons', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=10, textColor=colors.HexColor("#0D47A1"), spaceBefore=10, spaceAfter=5)
    
    story.append(Spacer(1, 5))
    story.append(Paragraph("REPORTE DE CONSULTA DE HISTORIAL DE RECEPCIÓN", style_title))
    story.append(Paragraph("Este documento contiene el listado de recepciones de materia prima filtradas según los criterios de búsqueda seleccionados en el sistema.", style_normal_text))
    story.append(Spacer(1, 10))
    
    # --- PANEL: FILTROS APLICADOS ---
    t_filtros_header = Table([[Paragraph("FILTROS DE BÚSQUEDA APLICADOS", style_blanco_bold)]], colWidths=[540])
    t_filtros_header.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#757575")), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(t_filtros_header)
    
    datos_filtros = [
        [Paragraph("RANGO FECHAS:", style_normal_bold), Paragraph(f"{filtros.get('fecha_inicio', 'Todos')} a {filtros.get('fecha_fin', 'Todos')}", style_normal_text),
         Paragraph("PROVEEDOR FILTRO:", style_normal_bold), Paragraph(str(filtros.get('proveedor', 'Todos')), style_normal_text)],
        [Paragraph("ESTATUS FILTRO:", style_normal_bold), Paragraph(str(filtros.get('estatus', 'Todos')), style_normal_text),
         Paragraph("BÚSQUEDA TEXTO:", style_normal_bold), Paragraph(str(filtros.get('buscar', 'Ninguno')), style_normal_text)]
    ]
    t_filtros = Table(datos_filtros, colWidths=[120, 150, 120, 150])
    t_filtros.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#F5F5F5")),
        ('BACKGROUND', (2,0), (2,-1), colors.HexColor("#F5F5F5"))
    ]))
    story.append(t_filtros)
    story.append(Spacer(1, 15))
    
    # --- TABLA DE RESULTADOS ---
    story.append(Paragraph(f"📋 REGISTROS ENCONTRADOS ({len(df_rep_filtered)} recepciones)", style_section))
    
    tabla_recepciones = [[
        Paragraph("FOLIO", style_blanco_bold),
        Paragraph("FECHA", style_blanco_bold),
        Paragraph("PROVEEDOR", style_blanco_bold),
        Paragraph("PO / FACTURA", style_blanco_bold),
        Paragraph("ACEPTACIÓN", style_blanco_bold),
        Paragraph("ESTATUS", style_blanco_bold)
    ]]
    
    for _, row in df_rep_filtered.iterrows():
        folio_val = row["Folio"]
        fecha_val = row["Fecha"]
        prov_val = row["Proveedor"]
        po_fact = f"{row['Orden_Compra']} / {row['Factura_Remision']}"
        acep_val = dict_acep.get(folio_val, "0/0 (0.0%)")
        estatus_val = row["Estatus_General"]
        
        status_color = "#2E7D32" if estatus_val == "Aceptado" else ("#FBC02D" if estatus_val == "Condicionado" else "#C62828")
        status_html = f"<b><font color='{status_color}'>{estatus_val}</font></b>"
        
        tabla_recepciones.append([
            Paragraph(folio_val, style_normal_text),
            Paragraph(fecha_val, style_normal_text),
            Paragraph(prov_val, style_normal_text),
            Paragraph(po_fact, style_normal_text),
            Paragraph(acep_val, style_normal_text),
            Paragraph(status_html, style_normal_text)
        ])
        
    t_recepciones = Table(tabla_recepciones, colWidths=[90, 70, 130, 110, 80, 60])
    t_recepciones.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#757575")),
        ('ALIGN', (0,0), (1,-1), 'CENTER'),
        ('ALIGN', (4,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.append(t_recepciones)
    
    # Firmas
    story.append(Spacer(1, 30))
    datos_firmas = [
        [Paragraph("_____________________________<br/>INSPECTOR DE CALIDAD", style_normal_bold),
         Paragraph("_____________________________<br/>SUPERVISOR DE CONTROL DE CALIDAD", style_normal_bold)]
    ]
    t_firmas = Table(datos_firmas, colWidths=[270, 270])
    t_firmas.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'TOP')
    ]))
    story.append(t_firmas)
    
    def decorate(canvas, doc):
        draw_sigrama_sgc_decorations(canvas, doc, "FO-MET-35", "REPORTE DE CONSULTA DE HISTORIAL")
        
    doc.build(story, onFirstPage=decorate, onLaterPages=decorate)
    print("PDF de Consulta de Historial creado.")

def generar_pdf_procedimiento_pralm01(output_pdf_path):
    """
    Genera un archivo PDF formal con el procedimiento PR-ALM-01
    adaptado para el sistema digital (12 mediciones, 4 placas, FO-MET-31/32/33, auto-push).
    """
    doc = SimpleDocTemplate(output_pdf_path, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story = []
    styles = getSampleStyleSheet()
    
    style_tit_proc = ParagraphStyle('T_Proc', parent=styles['Heading1'], fontSize=15, leading=19, fontName="Helvetica-Bold", textColor=colors.HexColor("#D32F2F"), spaceAfter=10)
    style_sec_proc = ParagraphStyle('S_Proc', parent=styles['Normal'], fontSize=10, leading=14, fontName="Helvetica-Bold", textColor=colors.HexColor("#0D47A1"), spaceBefore=10, spaceAfter=5)
    style_body_proc = ParagraphStyle('B_Proc', parent=styles['Normal'], fontSize=8.5, leading=12, fontName="Helvetica")
    style_body_proc_bold = ParagraphStyle('BB_Proc', parent=styles['Normal'], fontSize=8.5, leading=12, fontName="Helvetica-Bold")
    style_bullet_proc = ParagraphStyle('Bul_Proc', parent=styles['Normal'], fontSize=8.5, leading=12, fontName="Helvetica", leftIndent=15, bulletIndent=5)
    style_blanco_bold = ParagraphStyle('WB_Proc', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=8)
    
    story.append(Spacer(1, 10))
    story.append(Paragraph("PROCEDIMIENTO DE RECEPCIÓN DE MATERIA PRIMA", style_tit_proc))
    
    # Tabla de Metadatos
    meta_data = [
        [Paragraph("CÓDIGO:", style_body_proc_bold), Paragraph("PR-ALM-01", style_body_proc),
         Paragraph("REVISIÓN:", style_body_proc_bold), Paragraph("00 (Edición Digital)", style_body_proc)],
        [Paragraph("DEPARTAMENTO:", style_body_proc_bold), Paragraph("Almacén / Calidad", style_body_proc),
         Paragraph("SISTEMA:", style_body_proc_bold), Paragraph("SGC Digital Sigrama", style_body_proc)]
    ]
    t_meta = Table(meta_data, colWidths=[120, 150, 120, 150])
    t_meta.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#F5F5F5")),
        ('BACKGROUND', (2,0), (2,-1), colors.HexColor("#F5F5F5")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 10))
    
    # 1. OBJETIVO
    story.append(Paragraph("1. OBJETIVO", style_sec_proc))
    story.append(Paragraph(
        "Establecer de manera detallada y estricta los lineamientos para la recepción, inspección técnica, "
        "validación documental y aceptación de materia prima (lámina galvanizada y decapada) mediante la aplicación digital "
        "SGC Incoming, con el fin de asegurar que el material que ingresa a los procesos de producción de SIGRAMA cumple con "
        "las propiedades mecánicas, químicas y dimensionales necesarias para evitar productos no conformes y daños a la maquinaria.",
        style_body_proc
    ))
    
    # 2. ALCANCE
    story.append(Paragraph("2. ALCANCE", style_sec_proc))
    story.append(Paragraph(
        "Este procedimiento es aplicable a todo el personal involucrado en la cadena de suministro de SIGRAMA, iniciando desde el "
        "arribo del transporte del proveedor a la planta, pasando por la descarga, la inspección física y documental por parte de Calidad, "
        "hasta la identificación, resguardo y estiba final en el almacén de materia prima.",
        style_body_proc
    ))
    
    # 3. NORMAS DE REFERENCIA
    story.append(Paragraph("3. NORMAS DE REFERENCIA", style_sec_proc))
    story.append(Paragraph("• <b>ISO 9001:2015:</b> Sistema de Gestión de Calidad (Cláusula 8.4 Control de procesos, productos y servicios suministrados externamente).", style_bullet_proc))
    story.append(Paragraph("• <b>ASTM A653:</b> Especificación estándar para lámina de acero galvanizada por inmersión en caliente.", style_bullet_proc))
    story.append(Paragraph("• <b>ASTM A1011:</b> Especificación para acero laminado en caliente, decapado y aceitado (HRPO).", style_bullet_proc))
    story.append(Paragraph("• <b>PR-SGC-01:</b> Procedimiento para Control de Documentos (SIGRAMA).", style_bullet_proc))
    
    # 4. DEFINICIONES
    story.append(Paragraph("4. DEFINICIONES", style_sec_proc))
    story.append(Paragraph("• <b>Certificado de Molino:</b> Documento técnico expedido por el fabricante del acero que garantiza la composición química (Carbono, Manganeso, etc.) y las propiedades mecánicas (Límite elástico, tensión).", style_bullet_proc))
    story.append(Paragraph("• <b>Papel VCI (Volatile Corrosion Inhibitor):</b> Empaque industrial impregnado con químicos que inhiben la oxidación de los metales al crear una atmósfera protectora.", style_bullet_proc))
    story.append(Paragraph("• <b>G60 / G90:</b> Grado de recubrimiento de zinc. G90 ofrece una mayor resistencia a la corrosión en ambientes salinos o húmedos.", style_bullet_proc))
    story.append(Paragraph("• <b>Flor de Zinc (Spangle):</b> Patrón visible de cristales de zinc en la superficie de la lámina; su uniformidad es indicador de calidad en el proceso de inmersión.", style_bullet_proc))
    story.append(Paragraph("• <b>Atado / Bundle:</b> Unidad de empaque de láminas agrupadas por calibre y medida.", style_bullet_proc))
    story.append(Paragraph("• <b>Dosier de Calidad:</b> Expediente unificado que compila la portada (FO-MET-33), reporte de mediciones (FO-MET-31), reporte técnico estadístico de Gauss, etiquetas de almacén (FO-MET-32), certificado de calidad y orden de compra de Sigrama.", style_bullet_proc))
    
    story.append(PageBreak())
    
    # 5. DIAGRAMA DE FLUJO (TABLA ESCALONADA)
    story.append(Paragraph("5. DIAGRAMA DE FLUJO (TABLA ESCALONADA DEL PROCESO)", style_sec_proc))
    story.append(Spacer(1, 5))
    
    flow_data = [
        [Paragraph("RESPONSABLE", style_blanco_bold), Paragraph("PROCESO Ó ACTIVIDAD", style_blanco_bold), Paragraph("DOCUMENTO DE SALIDA", style_blanco_bold)],
        [Paragraph("Inspector de Calidad", style_body_proc_bold), Paragraph("Validación documental inicial (Cotejo de Certificado de Molino vs Orden de Compra de Sigrama).", style_body_proc), Paragraph("Cotejo Inicial y Remisión Firmada.", style_body_proc)],
        [Paragraph("Auxiliar de Almacén", style_body_proc_bold), Paragraph("Descarga física del material utilizando la grúa viajera y validación del peso bruto (Máximo 2.5 TON según RFQ).", style_body_proc), Paragraph("Registro de peso.", style_body_proc)],
        [Paragraph("Inspector de Calidad", style_body_proc_bold), Paragraph("Inspección dimensional micrométrica (4 placas/esquinas virtuales, 12 lecturas) e inspección visual de empaque y defectos.", style_body_proc), Paragraph("Reporte de Calidad Consolidado (FO-MET-31) y Reporte Técnico de Gauss.", style_body_proc)],
        [Paragraph("Auxiliar de Almacén", style_body_proc_bold), Paragraph("Identificación física del material en el almacén de metales y preservación con papel VCI.", style_body_proc), Paragraph("Tarjeta de Identificación (FO-MET-32) con código de barras o código interno.", style_body_proc)],
        [Paragraph("Sistema SGC Digital", style_body_proc_bold), Paragraph("Compilación del expediente digital unificado y respaldo automático a la nube mediante Token de GitHub.", style_body_proc), Paragraph("Dosier de Calidad Unificado (FO-MET-33) respaldado en GitHub.", style_body_proc)]
    ]
    t_flow = Table(flow_data, colWidths=[110, 250, 180])
    t_flow.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0D47A1")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#757575")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5)
    ]))
    story.append(t_flow)
    story.append(Spacer(1, 10))
    
    # 6. DESARROLLO DEL PROCEDIMIENTO
    story.append(Paragraph("6. DESARROLLO DEL PROCEDIMIENTO", style_sec_proc))
    
    story.append(Paragraph("6.1 Recepción y Validación Documental", style_body_proc_bold))
    story.append(Paragraph("Al arribo del material, el Inspector de Calidad debe cotejar la Remisión del proveedor contra la Orden de Compra (OC) de SIGRAMA y el Certificado de Molino.", style_body_proc))
    story.append(Paragraph("• Se debe verificar que el número de colada impreso en el atado coincida exactamente con el certificado.", style_bullet_proc))
    story.append(Paragraph("• Si el material carece de certificado o existe discrepancia en el grado de acero, el material no se descarga y se reporta inmediatamente a Compras.", style_bullet_proc))
    
    story.append(Paragraph("6.2 Maniobra de Descarga y Pesaje", style_body_proc_bold))
    story.append(Paragraph("El personal de Almacén procede a la descarga utilizando la grúa viajera.", style_body_proc))
    story.append(Paragraph("• <b>Restricción de Peso:</b> Ningún atado debe exceder las 2.5 Toneladas. En caso de que el proveedor envíe atados de mayor peso, se hará la observación y se evaluará el riesgo de maniobra; si pone en riesgo la infraestructura o seguridad, será rechazado.", style_bullet_proc))
    
    story.append(Paragraph("6.3 Inspección Técnica de Calidad", style_body_proc_bold))
    story.append(Paragraph("Una vez en piso, el Inspector de Calidad aplica los siguientes criterios:", style_body_proc))
    story.append(Paragraph("• <b>Inspección Dimensional (Digital):</b> Se realiza la medición del espesor de la lámina en 4 placas virtuales del atado. En cada placa se realizan 3 lecturas en diferentes puntos, acumulando un total de 12 lecturas de espesor por atado. La tolerancia dimensional aceptable está pre-configurada dinámicamente por SKU en el catálogo del sistema.", style_bullet_proc))
    story.append(Paragraph("• <b>Inspección Visual:</b> Se debe revisar el 100% de la cara superior y los bordes. Se rechaza material con:", style_bullet_proc))
    story.append(Paragraph("  - <b>Oxidación Blanca o Negra:</b> Presencia de humedad o falla en el pasivado.", style_bullet_proc))
    story.append(Paragraph("  - <b>Golpes o 'Escalones':</b> Deformaciones físicas que impidan el correcto nest del láser de fibra.", style_bullet_proc))
    story.append(Paragraph("  - <b>Falla en Flor:</b> Cristales de zinc irregulares o desprendimiento del recubrimiento.", style_bullet_proc))
    
    story.append(Paragraph("6.4 Identificación, Preservación y Estiba", style_body_proc_bold))
    story.append(Paragraph("• <b>Identificación:</b> El material conforme recibirá una Tarjeta de Identificación (FO-MET-32) verde con estatus 'Aceptado' con sus códigos internos, colada e histórico gaussiano. El material no conforme recibirá una etiqueta roja y se trasladará al área de segregación conforme al PR-SGC-04.", style_bullet_proc))
    story.append(Paragraph("• <b>Preservación:</b> Toda lámina debe ser resguardada sobre tarimas de madera (nunca contacto directo con suelo). Se debe colocar Papel VCI entre el material y el ambiente si el tiempo de almacenamiento previsto supera los 15 días.", style_bullet_proc))
    
    story.append(Paragraph("6.5 Cierre y Auto-Guardado en la Nube", style_body_proc_bold))
    story.append(Paragraph("El sistema recopila automáticamente los archivos PDF de portada, reporte consolidado de mediciones, reporte estadístico de Gauss y las etiquetas de identificación en un único archivo Dosier de Calidad (FO-MET-33). Una vez validado por el inspector, se realiza una sincronización en segundo plano con el repositorio web de GitHub mediante un Token seguro de acceso, permitiendo persistencia y auditoría inmediata.", style_body_proc))
    
    # 7. DOCUMENTOS RELACIONADOS
    story.append(Paragraph("7. DOCUMENTOS RELACIONADOS", style_sec_proc))
    story.append(Paragraph("• <b>FO-MET-31:</b> Reporte Consolidado de Inspección Dimensional de Materia Prima.", style_bullet_proc))
    story.append(Paragraph("• <b>FO-MET-32:</b> Tarjeta de Identificación de Atado de Materia Prima (Etiqueta).", style_bullet_proc))
    story.append(Paragraph("• <b>FO-MET-33:</b> Portada y Resumen de Contenido del Dosier de Calidad.", style_bullet_proc))
    story.append(Paragraph("• <b>PR-SGC-04:</b> Procedimiento para Control de Producto / Servicio No Conforme.", style_bullet_proc))
    story.append(Paragraph("• <b>PR-SGC-02:</b> Procedimiento para el Control de Registros.", style_bullet_proc))
    
    # 8. CONTROL DE REVISIONES
    story.append(Paragraph("8. CONTROL DE REVISIONES", style_sec_proc))
    story.append(Paragraph("• <b>Rev. 00:</b> Emisión inicial y adaptación completa para reestructuración del Sistema de Gestión de Calidad (SGC) digital SIGRAMA.", style_bullet_proc))
    
    def decorate(canvas, doc):
        draw_sigrama_sgc_decorations(canvas, doc, "PR-ALM-01", "PROCEDIMIENTO DE CONTROL DE MATERIA PRIMA")
        
    doc.build(story, onFirstPage=decorate, onLaterPages=decorate)
    print("PDF del Procedimiento PR-ALM-01 creado.")

def generar_pdf_remision_salida(datos_remision, output_pdf_path):
    """
    Genera una remisión de salida formal en PDF (FO-MET-36)
    para el despacho diario de láminas de acero.
    """
    doc = SimpleDocTemplate(output_pdf_path, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story = []
    styles = getSampleStyleSheet()
    
    style_title = ParagraphStyle('T_Rem', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=14, textColor=colors.HexColor("#D32F2F"), spaceAfter=5)
    style_subtitle = ParagraphStyle('S_Rem', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=10, textColor=colors.HexColor("#757575"), spaceAfter=15)
    style_normal_bold = ParagraphStyle('NB_Rem', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=8.5)
    style_normal_text = ParagraphStyle('NT_Rem', parent=styles['Normal'], fontName="Helvetica", fontSize=8.5, leading=12)
    style_blanco_bold = ParagraphStyle('WB_Rem', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=8.5)
    
    story.append(Spacer(1, 5))
    story.append(Paragraph("REMISIÓN DE SALIDA DE MATERIA PRIMA", style_title))
    story.append(Paragraph("Documento de registro interno para el despacho de materiales del almacén de metales.", style_subtitle))
    
    # Tabla de Datos Generales de la Remisión
    datos_meta = [
        [Paragraph("FOLIO REMISIÓN:", style_normal_bold), Paragraph(str(datos_remision.get("Folio_Salida", "REM-OUT-XXXX")), style_normal_text),
         Paragraph("FECHA DESPACHO:", style_normal_bold), Paragraph(f"{datos_remision.get('Fecha', '')} {datos_remision.get('Hora', '')}", style_normal_text)],
        [Paragraph("RESPONSABLE AUTORIZA:", style_normal_bold), Paragraph(str(datos_remision.get("Responsable", "N/D")), style_normal_text),
         Paragraph("PROYECTO / DESTINO:", style_normal_bold), Paragraph(str(datos_remision.get("Destino_Proyecto", "N/D")), style_normal_text)]
    ]
    t_meta = Table(datos_meta, colWidths=[130, 140, 130, 140])
    t_meta.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#F5F5F5")),
        ('BACKGROUND', (2,0), (2,-1), colors.HexColor("#F5F5F5")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 15))
    
    # Tabla de Detalle del Material Despachado
    story.append(Paragraph("📦 DETALLE DE MATERIAL DESPACHADO", ParagraphStyle('SH_Rem', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=10, textColor=colors.HexColor("#0D47A1"), spaceAfter=5)))
    
    headers_detalle = [
        Paragraph("ID ATADO", style_blanco_bold),
        Paragraph("SKU", style_blanco_bold),
        Paragraph("GRADO ACERO", style_blanco_bold),
        Paragraph("UBICACIÓN", style_blanco_bold),
        Paragraph("HOJAS RETIRADAS", style_blanco_bold),
        Paragraph("PESO APROX. (KG)", style_blanco_bold)
    ]
    
    tabla_detalle_data = [
        headers_detalle,
        [
            Paragraph(str(datos_remision.get("ID_Atado", "")), style_normal_text),
            Paragraph(str(datos_remision.get("SKU", "")), style_normal_text),
            Paragraph(str(datos_remision.get("Grado_Acero", "")), style_normal_text),
            Paragraph(str(datos_remision.get("Ubicacion_Almacen", "")), style_normal_text),
            Paragraph(str(datos_remision.get("Cantidad_Hojas_Despachadas", "")), style_normal_text),
            Paragraph(f"{float(datos_remision.get('Peso_Despachado_Kg', 0)):.2f}", style_normal_text)
        ]
    ]
    t_detalle = Table(tabla_detalle_data, colWidths=[110, 95, 95, 80, 80, 80])
    t_detalle.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#757575")),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 8)
    ]))
    story.append(t_detalle)
    story.append(Spacer(1, 15))
    
    # Campo de Observaciones
    story.append(Paragraph("📝 OBSERVACIONES", ParagraphStyle('SH_Obs', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=9, textColor=colors.HexColor("#0D47A1"), spaceAfter=3)))
    obs_text = str(datos_remision.get("Observaciones", "Ninguna."))
    if not obs_text.strip():
        obs_text = "Sin observaciones registradas."
    story.append(Paragraph(obs_text, style_normal_text))
    story.append(Spacer(1, 30))
    
    # Panel de Firmas
    datos_firmas = [
        [
            Paragraph("_____________________________<br/>ENTREGA (ALMACÉN DE METALES)", style_normal_bold),
            Paragraph("_____________________________<br/>RECIBE (OPERADOR/ÁREA DESTINO)", style_normal_bold)
        ],
        [Spacer(1, 15), Spacer(1, 15)],
        [
            Paragraph("_____________________________<br/>INSPECTOR DE CALIDAD (SGC)", style_normal_bold),
            Paragraph("_____________________________<br/>AUTORIZADO POR", style_normal_bold)
        ]
    ]
    t_firmas = Table(datos_firmas, colWidths=[270, 270])
    t_firmas.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'TOP')
    ]))
    story.append(t_firmas)
    
    # Decoracion canvas
    def decorate(canvas, doc):
        draw_sigrama_sgc_decorations(canvas, doc, "FO-MET-36", "REMISIÓN DE SALIDA DE MATERIAL")
        
    doc.build(story, onFirstPage=decorate, onLaterPages=decorate)
    print(f"PDF de Remisión de Salida {datos_remision.get('Folio_Salida')} creado.")




