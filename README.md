# Proyecto ETL de Análisis de Oportunidades de Negocio

## Introducción
Este proyecto realiza un análisis ETL sobre datos de oportunidades de negocio. Utilizando Python, se procesan datos desde un archivo en formato Excel, transformándolos para generar métricas clave y dashboards interactivos que responden a preguntas de negocio esenciales.

---


## Requisitos

### Herramientas y librerías necesarias:
- Python 3.x
- pandas
- openpyxl
- word2number

### Instalación de dependencias:
Ejecuta el siguiente comando para instalar las dependencias necesarias:
```bash
pip install pandas openpyxl word2number
```

---

## Estructura del Proyecto

### Carpetas:
- **`data/`**:
  - **`raw/`**: Contiene los datos originales sin procesar.
    - `BD_OPORTUNIDADES_23_24.csv`: Archivo original con datos de oportunidades.
  - **`processed/`**: Contiene los datos procesados generados por el flujo ETL.
    - `_NUEVA_BD_OPORTUNIDADES_23_24.xlsx`: Archivo procesado con hojas adicionales para análisis.
- **`src/`**:
  - `Reto_code.py`: Contiene todo el código del proyecto, incluyendo la extracción, transformación, cálculo de métricas y carga.
- **`docs/`**:
  - Carpeta destinada a toda la documentación del proyecto, como el manual de usuario y la documentación técnica.
  - `Manual de usuario.pdf`:Este manual está diseñado para ayudar a los usuarios finales a interactuar con los resultados del proyecto de análisis ETL.
  - `Reporte de Insights.pdf`:Este reporte está diseñado para destacar los resultados obtenidos durante el proyecto y las preguntas respondidas por los dashboards.
  - `Documento técnico.pdf`:Este es el complemento del README y con mas enfoque técnico.
  - `Diagrama de flujo.pdf`:Diagrama de flujo del ETL.

  - 
- **`dashboard/`**:
  - Capturas de pantalla de los dashboards y el proyecto en Power BI.
    - `Reto_dashboard.pbix`: Dashboard con los 6 pizarrones que contestan las 6 preguntas del reto.
    - `P1_Zona con mejores ventas en 2024.png`: Grafico con la zona con las mejores ventas en 2024.
    - `P2_Empresas con mayor crecimiento (2023-2024)`: Gráfico de las empresas con mayor crecimiento en los ultimos 2 años.
    - `P3_Propietarios con mayor crecimiento (2023-2024)`: Gráfico de los asesores con mayor crecimiento en los ultimos 2 años.
    - `P4_Zona con mas disminución (2023-2024)`: Gráfico con la zona que tuvo un decremento mayor en los ultimos dos años.
    - `P5_Cliente con crecimiento en la zona con menor importe`: Gráfico con el cliente que tuvo el mayor incremento en la zona con menor importe de ventas.
    - `P6_Relacion numero de participantes y zona`: Grafico con la relacion entre los participantes, las zonas y el importe de ventas.
- **Archivo principal (`README.md`)**: Documentación general del proyecto.

---

## Flujo del Proyecto

### 1. Extracción
Los datos se cargan desde un archivo CSV con la función `cargar_datos`.

### 2. Transformación
- **Limpieza de datos** (`limpiar_datos`): Manejo de valores nulos, formatos, y duplicados.
- **Normalización** (`normalizar_datos`): Conversión de importes a diferentes divisas.
- **Generación de columnas derivadas** (`generar_columnas`): Folios, rangos de importes y clasificaciones.

### 3. Cálculo de Métricas
- **Agrupaciones** (`calcular_agrupaciones`): Análisis por zonas, empresas y propietarios.
- **Crecimientos** (`calcular_crecimientos`): Comparación porcentual entre 2023 y 2024.

### 4. Carga
Se exportan los datos procesados a un archivo Excel con hojas adicionales para métricas calculadas.

---

## Instrucciones de Uso

### Pasos para ejecutar el código:
1. Asegúrte de tener el archivo `BD_OPORTUNIDADES_23_24.csv` en la carpeta `data/raw/`.
2. Modifica la ruta del archivo de entrada y salida en el script si es necesario:
   ```python
   ruta_archivo = 'data/raw/BD_OPORTUNIDADES_23_24.csv'
   ruta_salida = 'data/processed/_NUEVA_BD_OPORTUNIDADES_23_24.xlsx'
   ```
3. Ejecuta el script:
   ```bash
   python src/Reto_code.py
   ```
### Dashboards
- Abre el archivo `Reto_dashboard.pbix` en Power BI para explorar las visualizaciones.
- Capturas de pantalla disponibles en `dashboard/`.

---

## Resultados

### Archivo Excel generado:
- **Datos procesados**: Contiene los datos originales limpios y enriquecidos.
- **Crecimiento_Empresas**: Muestra el crecimiento porcentual de las empresas entre 2023 y 2024.
- **Crecimiento_Propietarios**: Muestra el crecimiento porcentual de los propietarios entre 2023 y 2024.
- **Crecimiento_Zonas**: Muestra el crecimiento porcentual de las zonas entre 2023 y 2024.

### Dashboards:
- Capturas de pantalla  de los dasboards disponibles en `dashboard/`.
- Proyecto Power BI disponible como `proyecto_power_bi.pbix`.

---

## Notas Adicionales

- **Crecimiento infinito (`inf`)**:
  Actualmente, los valores de crecimiento infinito se generan cuando los datos del año 2023 tienen un valor de 0 y existen datos positivos en 2024. Esto ocurre al calcular el crecimiento porcentual. En futuras versiones, se podría implementar una estrategia para manejar este caso, como:
  - Reemplazar `inf` con un texto explicativo como "Sin datos suficientes".
  - Asignar un valor simbólico como `100%` si las ventas pasaron de 0 en 2023 a un valor positivo en 2024.

---

## Futuras Mejoras
- Crear dashboards interactivos en Power BI o Excel para explorar métricas clave de forma visual e intuitiva.
- Implementar manejo automático de casos de crecimiento infinito.
- Optimizar el código para reducir el tiempo de ejecución en bases de datos más grandes.


## Contacto
Para dudas o consultas, por favor, contacta a:
- Correo: andreareyesmejia0@gmail.com
- Usuario de GitHub: 


