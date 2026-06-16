# Manual de Usuario: Sistema de Control de Calidad en Recepción (Incoming)

Este manual de usuario proporciona las pautas y guías paso a paso para el uso del **Sistema de Control de Calidad en Recepción** de Industria Sigrama. El sistema tiene como objetivo automatizar la evaluación de conformidad del acero recibido en base a tolerancias micrométricas y registrar la información de forma estructurada en el SGC.

---

## 🔐 1. Niveles de Acceso y Contraseñas

La aplicación cuenta con tres perfiles de usuario diferentes, administrados en tiempo real desde la barra lateral izquierda en el panel **🔐 Control de Acceso**:

| Perfil de Usuario | Contraseña | Permisos y Cobertura |
| :--- | :--- | :--- |
| **Público / Consulta** | *Ninguna (Vacío)* | Solo lectura. Puede navegar por el Módulo 1 (Dashboard) y Módulo 3 (Historial), así como descargar dosiers e informes. No tiene acceso al Módulo 2 ni puede realizar modificaciones. |
| **Inspector (Registro)** | *[Consultar con Calidad / Clave de Inspector]* | Acceso de lectura más la capacidad de ingresar al **Módulo 2 (Registro)** para subir plantillas de medición de atados, cargar fotos de defectos y guardar nuevos lotes. |
| **Administrador** | *[Clave Restringida / Consultar al departamento de Calidad]* | Acceso total. Incluye permisos de Inspector más la capacidad de eliminar expedientes en Módulo 3 y agregar o modificar las tolerancias de SKUs en el Módulo 4. |

> [!TIP]
> **Intercambiabilidad de campos (Mejora de UX):**
> Para evitar errores al capturar las claves, el sistema admite escribir cualquiera de las dos contraseñas en cualquier campo (Administrador o Inspector/Registro). El sistema resolverá automáticamente el nivel de privilegios correspondiente.

---

## 📊 2. Módulo 1: Analíticas y Dashboard (OKRs)

Este módulo es el panel de inicio y es de acceso público. Permite analizar el desempeño histórico y en tiempo real de los proveedores de acero:

1. **Filtros de Búsqueda**: Defina el rango de fechas, proveedor, material (Galvanizado, Decapado, Aluminio) y calibre para actualizar las gráficas.
2. **🎯 Tablero de OKRs (Objetivos y Resultados Clave)**:
   * **OKR 1: Conformidad y Calidad de Materia Prima**
     * *KR 1.1: Conformidad de Atados (Meta ≥ 95%)*: Porcentaje de rollos individuales que cumplieron con la especificación dimensional y visual.
     * *KR 1.2: Aceptación Total de Lotes (Meta ≥ 90%)*: Porcentaje de recepciones que ingresaron sin requerir devoluciones.
     * *KR 1.3: Eficiencia de Acero Conforme (Meta ≥ 95%)*: Proporción del tonelaje (peso en Kg) de acero aceptado contra el bruto recibido.
     * *Cada KR muestra una desviación (delta en verde o rojo) contra su meta establecida.*
   * **OKR 2: Eficiencia de Abastecimiento y Control**
     * Mide la cantidad de atados, kilogramos (y libras) de acero ingresados, y número total de lotes procesados.
3. **Gráficos Estadísticos**:
   * *Estado General de Lotes*: Gráfico de pastel (Verde: Aceptado, Amarillo: Condicionado, Rojo: Rechazado).
   * *Volumen por Proveedor*: Gráfico de barras que detalla los kilogramos suministrados por cada empresa.
   * *Defectos y Frecuencias*: Histograma horizontal con las causas y frecuencia de rechazos en los atados.
4. **Listado de Recepciones Recientes**: Tabla resumida del periodo con la columna **`Aceptación Atados`** (ej. `11/12 (91.7%)`) indicando la proporción conforme por folio.
5. **📥 Exportar Reporte Ejecutivo**:
   * Al pie de la página, haga clic en **`Descargar Reporte Ejecutivo del Dashboard (PDF)`** para obtener el documento oficial consolidado de analíticas y OKRs (`FO-SGC-03`) listo para impresión o auditoría.

---

## 📥 3. Módulo 2: Registro de Recepción (Incoming)

*Requiere contraseña de **Inspector** o **Administrador** en la barra lateral.*

Este módulo permite dar de alta un lote de materia prima:

### Paso 1: Descargar y Llenar la Plantilla Excel
1. Haga clic en **`Descargar Formato de Plantilla Corporativa (.xlsx)`** para obtener el formato oficial `FO-MET-30`. El archivo se descargará con la fecha del día actual en su nombre.
2. Complete la información de las **20 columnas** en Excel. El archivo contiene un único renglón por cada atado/rollo físico. En este único renglón se capturan las 12 mediciones de espesor micrométrico de las 4 placas (Placa 1: Espesores 1 a 3; Placa 2: Espesores 4 a 6; Placa 3: Espesores 7 a 9; Placa 4: Espesores 10 a 12), además del peso total del atado, cantidad total de hojas, número de colada, lote y observaciones. No es necesario duplicar filas ni capturar códigos de SKU o columnas de Placa.
3. *Asegúrese de usar los dropdowns provistos en las celdas de las columnas de Calibre (ej. `CAL 14`) y de material (ej. `Galvanizado`, `Decapado`, `Aluminio`).*

### Paso 2: Completar Datos Generales y Documentación
1. Ingrese los datos de cabecera: Proveedor, Orden de Compra (PO), Factura/Remisión y Nombre del Inspector.
2. Suba los dos documentos de soporte obligatorios en PDF:
   * **Certificado de Calidad del Proveedor (Certificado de Molino)**.
   * **Orden de Compra de Sigrama**.

### Paso 3: Inspección Visual y Apariencia por Rollo
1. Suba el Excel completado en la sección de mediciones. El sistema agrupará las mediciones y desplegará **una sola tarjeta interactiva por cada Atado único** (evitando duplicar el checklist de apariencia visual y fotos de evidencia).
2. En cada tarjeta:
   * Desmarque las casillas de los aspectos de apariencia que presenten anomalías (por defecto todas están en "Cumple" marcadas con ✔️ y se aplican a todo el atado).
   * Si detecta defectos, suba evidencia fotográfica en el cargador del atado: `📸 Subir Fotos de Defectos Visuales`.
   * En la parte inferior de la tarjeta, use las **pestañas (tabs) de Placas** para ver las tablas detalladas de mediciones, desviaciones y dictámenes específicos de cada hoja (Placa 1, 2, 3 y 4).
   * Debajo de las pestañas, se despliega un **único Gráfico de Gauss Consolidado** que dibuja las campanas de distribución de las 4 placas juntas en una sola vista con 4 curvas de diferentes colores y su leyenda interactiva, facilitando la comparación instantánea de la dispersión micrométrica de todo el atado.
   * Presione **`💾 Procesar y Guardar`** para pre-evaluar el atado completo. El sistema evaluará los criterios visuales generales y *todas las hojas* del atado, mostrando un dictamen claro y los motivos detallados si alguna placa no cumple.

### Paso 4: Salvar Todo e Integrar
1. Una vez pre-evaluados los atados, haga clic en el botón principal **`💾 Salvar Todo: Procesar e Integrar Recepción a Base de Datos`** al final del formulario.
2. El sistema creará el folio correspondiente (ej. `INC-2026-0003`), actualizará la base de datos maestra y determinará el **Estatus General del Lote** automáticamente:
   * **`Aceptado`**: Si el 100% de los atados resultaron conformes.
   * **`Condicionado`**: Si todos los atados cumplen con sus tolerancias de espesor, ancho y largo, pero se rechazaron algunos *únicamente* por defectos visuales o aspectos de acabado (menos severos).
   * **`Rechazado`**: Si al menos un atado falló en dimensiones micrométricas críticas (espesores, ancho, largo) o por exceso de peso (>2.5 toneladas).
3. Podrá descargar el **Dosier de Calidad Consolidado (PDF)** unificado y las etiquetas de identificación individuales de forma inmediata.

---

## 🔍 4. Módulo 3: Consulta de Historial y Descarga de Dosiers

Este módulo permite consultar recepciones pasadas:

1. **Búsqueda por Filtro**: Busque por Folio, Proveedor, Rango de fechas o Estatus de Calidad.
2. **Selección de Expediente**: Seleccione un folio de la lista para ver sus datos generales y mediciones detalladas.
3. **Descargas Disponibles**:
   * **`Descargar Dosier de Calidad (PDF)`**: Descarga el archivo consolidado unificado que incluye: Portada (`FO-MET-33`), Reporte de Inspección (`FO-MET-31`), Reporte Técnico de Espesores con curvas Plotly, Tarjetas de Identificación de Atado con curva de Gauss integrada (`FO-MET-32`), Certificado del Molino y la Orden de Compra.
   * **`Descargar Solo Etiquetas FO-MET-32 (PDF)`**: Descarga un compendio con una página carta por atado que incluye la etiqueta de almacén con su campana de Gauss al pie.
   * **`Descargar Carpeta Electrónica Completa (ZIP)`**: Descarga un comprimido con todos los PDF individuales, gráficas PNG y fotografías de defectos que conforman el expediente digital del lote.
4. **Regenerar Expediente**: Haga clic en **`🔄 Regenerar y Actualizar Todos los PDFs de este Folio`** para volver a calcular y compilar los reportes de un folio histórico (por ejemplo, si se corrigieron parámetros).
5. **Zona de Control Administrativo**:
   * *Requiere contraseña de **Administrador**.*
   * Permite eliminar permanentemente de las bases de datos y del disco un lote y toda su documentación física asociada mediante la confirmación de un checkbox de seguridad.

---

## ⚙️ 5. Módulo 4: Catálogo de Tolerancias de SKU

*La consulta es pública. La edición y adición requieren clave de **Administrador**.*

Este módulo permite configurar los parámetros técnicos aceptables para cada código de SKU en planta:

1. **Visualización**: Tabla maestra de SKUs con sus especificaciones de espesor nominal, tolerancias simétricas o asimétricas, ancho, largo, zinc mínimo y dureza máxima.
2. **Agregar Nuevo Producto**: Formulario para registrar un código nuevo, su tipo de lámina, grado de acero y valores de tolerancia técnica.
3. **Modificar Producto Existente**: Permite corregir los valores nominales o tolerancias de un SKU ya registrado. Los cambios se guardan en `BD_Parametros_Materia_Prima.xlsx` de forma inmediata.
