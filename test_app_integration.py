import os
import shutil
import pandas as pd
import utils_pdf

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CARPETAS_DIR = os.path.join(BASE_DIR, "carpetas_electronicas")

print("--- INICIANDO PRUEBA DE INTEGRACION ---")

# 1. Cargar parametros de SKU muestra
db_parametros_path = os.path.join(BASE_DIR, "BD_Parametros_Materia_Prima.xlsx")
df_skus = pd.read_excel(db_parametros_path)
sku_info = df_skus[df_skus["SKU"] == "SKU-GALV-14"].iloc[0].to_dict()
print("SKU de Prueba:", sku_info["SKU"])
print("Nominal de Espesor:", sku_info["Espesor_Nominal_in"])

# 2. Crear datos de atados de prueba
df_atados_test = pd.DataFrame([
    {
        "ID_Atado": "INC-TEST-0001-A01",
        "Folio": "INC-TEST-0001",
        "ID_Atado_Proveedor": "AT-PROV-111",
        "SKU": "SKU-GALV-14",
        "Grado_Acero": "ASTM A653 CS Tipo B",
        "Num_Colada": "COL-77621",
        "Lote_Heat": "LOT-123",
        "Espesor_Medido_1_in": 0.0748,
        "Espesor_Medido_2_in": 0.0752,
        "Espesor_Medido_3_in": 0.0750,
        "Espesor_Medido_4_in": 0.0749,
        "Espesor_Medido_5_in": 0.0751,
        "Espesor_Medido_6_in": 0.0750,
        "Espesor_Medido_7_in": 0.0748,
        "Espesor_Medido_8_in": 0.0752,
        "Espesor_Medido_9_in": 0.0750,
        "Espesor_Medido_10_in": 0.0749,
        "Espesor_Medido_11_in": 0.0751,
        "Espesor_Medido_12_in": 0.0750,
        "Ancho_Medido_in": 48.05,
        "Largo_Medido_in": 120.00,
        "Cantidad_Hojas": 150,
        "Peso_Total_Kg": 2500,
        "Peso_Total_Lb": 5511.5,
        "Zinc_Medido_oz_ft2": 0.58,
        "Dureza_Medida_HRB": 68,
        "Aceitado_OK": "N/A",
        "Defectos_Visuales": "Ninguno",
        "Ubicacion_Almacen": "B-04",
        "Estatus_Calidad": "Aceptado",
        "Observaciones": "Todo en orden.",
        "Proveedor": "Ternium Mexico",
        "Tipo_Lamina": "Galvanizada",
        "Acabado": "Pasivado",
        "Espesor_Nominal": 0.0750,
        "Espesor_Tol_Min": -0.0050,
        "Espesor_Tol_Max": 0.0050,
        "Ancho_Tol_Min_Val": 47.88,
        "Ancho_Tol_Max_Val": 48.12
    },
    {
        "ID_Atado": "INC-TEST-0001-A02",
        "Folio": "INC-TEST-0001",
        "ID_Atado_Proveedor": "AT-PROV-222",
        "SKU": "SKU-GALV-14",
        "Grado_Acero": "ASTM A653 CS Tipo B",
        "Num_Colada": "COL-77621",
        "Lote_Heat": "LOT-123",
        "Espesor_Medido_1_in": 0.0782,
        "Espesor_Medido_2_in": 0.0751,
        "Espesor_Medido_3_in": 0.0753,
        "Espesor_Medido_4_in": 0.0750,
        "Espesor_Medido_5_in": 0.0752,
        "Espesor_Medido_6_in": 0.0751,
        "Espesor_Medido_7_in": 0.0753,
        "Espesor_Medido_8_in": 0.0750,
        "Espesor_Medido_9_in": 0.0752,
        "Espesor_Medido_10_in": 0.0751,
        "Espesor_Medido_11_in": 0.0753,
        "Espesor_Medido_12_in": 0.0750,
        "Ancho_Medido_in": 48.02,
        "Largo_Medido_in": 120.00,
        "Cantidad_Hojas": 148,
        "Peso_Total_Kg": 2480,
        "Peso_Total_Lb": 5467.5,
        "Zinc_Medido_oz_ft2": 0.59,
        "Dureza_Medida_HRB": 69,
        "Aceitado_OK": "N/A",
        "Defectos_Visuales": "Ninguno",
        "Ubicacion_Almacen": "B-04",
        "Estatus_Calidad": "Aceptado",
        "Observaciones": "En limites superiores.",
        "Proveedor": "Ternium Mexico",
        "Tipo_Lamina": "Galvanizada",
        "Acabado": "Pasivado",
        "Espesor_Nominal": 0.0750,
        "Espesor_Tol_Min": -0.0050,
        "Espesor_Tol_Max": 0.0050,
        "Ancho_Tol_Min_Val": 47.88,
        "Ancho_Tol_Max_Val": 48.12
    }
])

# 3. Probar graficacion de curvas
folio_test = "INC-TEST-0001"
test_folder = os.path.join(CARPETAS_DIR, folio_test)
os.makedirs(test_folder, exist_ok=True)

img_curvas_path = os.path.join(test_folder, f"Curvas_Tolerancia_{folio_test}.png")
utils_pdf.generar_graficas_tolerancia(df_atados_test, sku_info, img_curvas_path)
assert os.path.exists(img_curvas_path), "Error: No se genero la imagen de curvas."
print("[OK] Grafica de curvas generada correctamente.")

# 4. Probar generacion de PDFs individuales
pdf_reporte = os.path.join(test_folder, f"Reporte_FO-MET-31_{folio_test}.pdf")
pdf_etiquetas = os.path.join(test_folder, f"Etiquetas_FO-MET-32_{folio_test}.pdf")
pdf_portada = os.path.join(test_folder, f"Portada_FO-MET-33_{folio_test}.pdf")

datos_reporte = {
    "Fecha": "15/06/2026",
    "Proveedor": "Ternium Mexico",
    "Orden_Compra": "PO-2602-0711",
    "Factura_Remision": "TX6793212",
    "Inspector": "Jesus Morales",
    "Estatus_General": "Aceptado"
}

utils_pdf.generar_pdf_reporte_consolidado_fomet31(folio_test, datos_reporte, df_atados_test, sku_info, img_curvas_path, pdf_reporte)
assert os.path.exists(pdf_reporte), "Error: No se genero el PDF FO-MET-31."
print("[OK] PDF FO-MET-31 (Reporte consolidado) generado correctamente.")

utils_pdf.generar_pdf_etiqueta_atado_fomet32(folio_test, df_atados_test, pdf_etiquetas)
assert os.path.exists(pdf_etiquetas), "Error: No se genero el PDF FO-MET-32."
print("[OK] PDF FO-MET-32 (Etiquetas de atado) generado correctamente.")

utils_pdf.generar_pdf_portada_dosier_fomet33(folio_test, datos_reporte, df_atados_test, pdf_portada)
assert os.path.exists(pdf_portada), "Error: No se genero el PDF FO-MET-33."
print("[OK] PDF FO-MET-33 (Portada de dosier) generado correctamente.")

# 5. Crear archivos mock para Certificado y OC
cert_mock = os.path.join(test_folder, "MOCK_Certificado.pdf")
oc_mock = os.path.join(test_folder, "MOCK_OC.pdf")

# Copiar algun PDF generado para usarlo como mock
shutil.copy(pdf_portada, cert_mock)
shutil.copy(pdf_portada, oc_mock)

# 6. Probar compilacion del Dosier
pdf_dosier = os.path.join(test_folder, f"Dosier_Calidad_{folio_test}.pdf")
utils_pdf.compilar_dosier_calidad(
    pdf_dosier,
    pdf_portada,
    pdf_reporte,
    pdf_etiquetas,
    [cert_mock],
    [oc_mock]
)
assert os.path.exists(pdf_dosier), "Error: No se compilo el Dosier unificado."
print("[OK] Fusion de PDFs con pypdf completada con exito.")
print("--- PRUEBA DE INTEGRACION COMPLETADA CON EXITO ---")
