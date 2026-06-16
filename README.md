# Sistema de Control de Calidad en Recepción (Incoming) - Industria Sigrama

Este es un sistema web interactivo desarrollado con **Streamlit** para gestionar el control de calidad de materia prima (láminas y rollos de acero) en la Planta de Metales. Permite registrar mediciones micrométricas, analizar la desviación a través de curvas de Gauss y generar de forma automática los reportes técnicos y dosiers de calidad consolidados bajo la norma del SGC.

---

## 🚀 Instalación y Ejecución Local

### 1. Requisitos Previos
Asegúrese de tener Python (versión 3.9 o superior) instalado.

### 2. Instalar Dependencias
Instale todas las librerías necesarias con el siguiente comando en su terminal:
```bash
pip install -r requirements.txt
```

### 3. Ejecutar la Aplicación
Inicie el servidor local de Streamlit con:
```bash
streamlit run app.py
```
La aplicación se abrirá automáticamente en su navegador en `http://localhost:8501` (o en el puerto disponible).

---

## 📂 Estructura del Proyecto

* **`app.py`**: Interfaz de usuario de la aplicación en Streamlit y lógica de navegación.
* **`utils_pdf.py`**: Motor gráfico de PDF (ReportLab y Matplotlib) para la generación de curvas gaussianas, reportes consolidados (`FO-MET-31`), etiquetas de identificación (`FO-MET-32`), portadas (`FO-MET-33`) y compilación de dosiers.
* **`BD_Parametros_Materia_Prima.xlsx`**: Catálogo de tolerancias nominales por SKU.
* **`BD_Reportes_Incoming.xlsx`**: Base de datos de recepciones y folios de lote.
* **`BD_Atados_Incoming.xlsx`**: Base de datos detallada con las mediciones de atados individuales.
* **`plantilla_incoming_calidad.xlsx`**: Plantilla oficial de captura de mediciones descargable.
* **`carpetas_electronicas/`**: Directorio donde se guardan de forma organizada los expedientes y archivos PDF de cada lote (organizado por Folio).
* **`requirements.txt`**: Listado de dependencias para el despliegue.

---

## 🛠️ Despliegue en GitHub y Streamlit Cloud

### Paso 1: Subir el proyecto a GitHub
1. Cree un repositorio en su cuenta de GitHub (ej. `incoming-calidad`).
2. Abra su terminal en la carpeta del proyecto y ejecute los siguientes comandos:
```bash
git init
git add .
git commit -m "Initial commit: Sistema de control de calidad"
git branch -M main
git remote add origin https://github.com/SU_USUARIO/incoming-calidad.git
git push -u origin main
```

### Paso 2: Desplegar en Streamlit Community Cloud
1. Inicie sesión en [Streamlit Community Cloud](https://share.streamlit.io/).
2. Haga clic en **"New app"**.
3. Seleccione su repositorio (`incoming-calidad`), la rama (`main`), y el archivo principal (`app.py`).
4. Haga clic en **"Deploy"**. ¡Listo! Su aplicación estará en línea de forma pública o privada en segundos.

> [!WARNING]
> **Nota de Persistencia de Datos:**
> Streamlit Cloud ejecuta la aplicación en contenedores efímeros. Las modificaciones y registros escritos en los archivos Excel locales (`BD_Reportes_Incoming.xlsx`, etc.) se perderán cuando el contenedor se reinicie o reconstruya. 
> Para un entorno de producción persistente en la nube, se recomienda conectar Streamlit a una base de datos en la nube (como PostgreSQL, SQLite en un volumen persistente, o a través de la integración nativa con Google Sheets).
