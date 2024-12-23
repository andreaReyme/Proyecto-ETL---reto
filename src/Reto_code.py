import pandas as pd
from word2number import w2n
from openpyxl import Workbook
from openpyxl.styles import NamedStyle
def cargar_datos(ruta_archivo):
    """
    Carga los datos desde un archivo CSV.
    Parametros: 
        -ruta_archivo (str): Rusta completa dl archivo CSV.
    Retorna:
        -pd.DataFrame: Datos cargados en un DataFrame de pandas.
    """
    return pd.read_csv(ruta_archivo, encoding='utf-8')

# Función: Limpiar datos
def limpiar_datos(tabla):
    """
    Realiza la limpiezada inicial de los datos.
        -Convierte los valores escritos como palabras a números.
        -Convierte las columnas de Importe y Participantes a tipo numerico
        -Reemplaza todos los valores nulos en la columna Importe a 0.
        -Corrige el formato de la columna FechaCierre
        -Unifica la columna Zona para que no haya espacios de sobra. 
        -Reemplaza los valores nulos y "sin datos" por 0's en la columna de Participantes
        -Elimina los registros duplicados.
    Parametros: 
        -tabla(pd.DataFrame): DataFrame original sin limpiar
    Retorna:
        -tabla(pd.DataFrame): DataFrame limpio y transformado
    """

    # Identificar campos escritos como palabras y convertirlos a números
    def convertir_a_numero(valor):
        try:
            if isinstance(valor, str):  
                return w2n.word_to_num(valor.lower())  
        except ValueError:
            pass  # Si falla, devolver el valor original
        return valor

    tabla['Importe'] = tabla['Importe'].apply(convertir_a_numero)
    tabla['Importe'] = pd.to_numeric(tabla['Importe'], errors='coerce')
    tabla['Importe'] = tabla['Importe'].fillna(0)
    
    tabla['FechaCierre'] = pd.to_datetime(tabla['FechaCierre'], format='%d/%m/%Y %H:%M', errors='coerce')
    tabla['FechaCierre'] = tabla['FechaCierre'].dt.date
    
    tabla['Zona'] = tabla['Zona'].fillna('Zona 6')
    tabla['Zona'] = tabla['Zona'].str.strip()
    tabla['Zona'] = tabla['Zona'].astype('category')
    
    tabla['Participantes'] = tabla['Participantes'].replace('Sin datos', 0)
    tabla['Participantes'] = pd.to_numeric(tabla['Participantes'], errors='coerce')
    tabla['Participantes'] = tabla['Participantes'].fillna(0)

    # Eliminar duplicados en 'IdOportunidad'
    #print("Número de filas duplicadas (antes):", tabla.duplicated().sum())
    #tabla = tabla.drop_duplicates(subset=['IdOportunidad'])
    tabla = tabla.drop_duplicates()  # Sin 'subset', considera toda la fila
    #print("Número de filas duplicadas (después):", tabla.duplicated().sum())

    print("Información general de la tabla:")
    print(tabla.info())
    print("\nConteo de valores nulos por columna después de la limpieza:")
    print(tabla.isnull().sum())
    return tabla

# Función: Normalizar datos
def normalizar_datos(tabla, tasas_cambio):
    """
    Normaliza los datos monetarios las fechas
        -Estandariza el formato de la fecha a tipo date.
        -Convierte el importe a distintas divisas.
    Parametros: 
        -tabla(pd.DataFrame): DataFrame limpio y transformado
        -tasas_cambio (dict): Diccionario con tasas de cambio por divisa.
    Retorna:
        -tabla(pd.DataFrame): DataFrame con importes normalizados y fechas formateadas.
    """
    # Convertir FechaCierre a datetime y normalizar la hora
    tabla['FechaCierre'] = pd.to_datetime(tabla['FechaCierre'], format='%d/%m/%Y %H:%M', errors='coerce').dt.normalize()

    tabla['TipoDivisaAjuste'] = tabla['TipoDivisaAjuste'].str.upper()
    # Crear columnas de conversión con tasas específicas
    tabla['Importe_MXN'] = tabla.apply(lambda row: row['Importe'] * tasas_cambio[row['TipoDivisaAjuste']], axis=1)
    tabla['Importe_USD'] = tabla.apply(lambda row: row['Importe_MXN'] / 20, axis=1)  # Tasa de cambio para USD
    tabla['Importe_EUR'] = tabla.apply(lambda row: row['Importe_MXN'] / 22, axis=1)  # Tasa de cambio para EUR

    print("Información general de la tabla:")
    print(tabla.info())
    return tabla

# Función: Generar columnas derivadas
def generar_columnas(tabla):
    """
    Genera columnas derivadas y clasificaciones.
        -Se crean identificadores unicos legibles.
        -Clasifica los importes por rangos.
        -Determina zonas importantes.
        -Segmenta fechas en mes, año y trimestre.
    Parametros: 
        -tabla (pd.DataFrame): DataFrame normalizado.
    Retorna:
        -tabla (pd.DataFrame): DataFrame enriquecido con columnas adicionales.
    """
    # Generar folios
    tabla['FolioOportunidad'] = tabla['IdOportunidad'].map({id_op: f"Oportunidad {i+1}" for i, id_op in enumerate(tabla['IdOportunidad'].unique())})
    tabla['FolioEmpresa'] = tabla['IdEmpresa'].map({id_emp: f"Empresa {i+1}" for i, id_emp in enumerate(tabla['IdEmpresa'].unique())})
    tabla['FolioPropietario'] = tabla['IdPropietario'].map({id_prop: f"Propietario {i+1}" for i, id_prop in enumerate(tabla['IdPropietario'].unique())})
    
    # Clasificación por rango de importe
    rangos = [0, 217000, 537000, 34000000]
    etiquetas = ['Bajo', 'Medio', 'Alto']
    tabla['RangoImporte'] = pd.cut(tabla['Importe_MXN'], bins=rangos, labels=etiquetas, include_lowest=True)
    
    # Clasificación de zonas importantes
    ingresos_por_zona = tabla.groupby('Zona', observed=True)['Importe_MXN'].sum().sort_values(ascending=False)
    zonas_importantes = ingresos_por_zona.head(3).index.tolist()
    tabla['ClasificacionZona'] = tabla['Zona'].apply(lambda x: 'Importante' if x in zonas_importantes else 'Otras')
    
    # Segmentar fechas
    tabla['AnoCierre'] = pd.to_datetime(tabla['FechaCierre']).dt.year
    tabla['MesCierre'] = pd.to_datetime(tabla['FechaCierre']).dt.month
    tabla['TrimestreCierre'] = pd.to_datetime(tabla['FechaCierre']).dt.quarter
    
    print(tabla.info())
    return tabla

# Función: Calcular agrupaciones
def calcular_agrupaciones(tabla):
    """
    Calcula métricas agregadas por zonas, empresas y propietarios.
        - Densidad de ingresos por zona.
        - Densidad de ingresos por empresa.
        - Clasificación de propietarios según ingresos totales. 
    Parametros: 
        -tabla (pd.DataFrame): DataFrame enriquecido.
    Retorna:
        Tres DataFrames con las métricas calculadas:
            - densidad_ingresos_zona
            - densidad_ingresos_empresa
            - ingresos_propietarios
    """
    
    # Densidad de ingresos por zona
    densidad_ingresos_zona = tabla.groupby('Zona', observed=True).agg(
        IngresoTotal=('Importe_MXN', 'sum'),
        Oportunidades=('IdOportunidad', 'count')
    ).reset_index()
    densidad_ingresos_zona['DensidadIngreso'] = densidad_ingresos_zona['IngresoTotal'] / densidad_ingresos_zona['Oportunidades']
    
    # Densidad de ingresos por empresa
    densidad_ingresos_empresa = tabla.groupby('FolioEmpresa').agg(
        IngresoTotal=('Importe_MXN', 'sum'),
        Oportunidades=('IdOportunidad', 'count')
    ).reset_index()
    densidad_ingresos_empresa['DensidadIngreso'] = densidad_ingresos_empresa['IngresoTotal'] / densidad_ingresos_empresa['Oportunidades']
    
    # Clasificación de propietarios
    ingresos_propietarios = tabla.groupby('FolioPropietario').agg(
        IngresoTotal=('Importe_MXN', 'sum')
    ).reset_index()
    ingresos_propietarios['Clasificacion'] = pd.qcut(
        ingresos_propietarios['IngresoTotal'], 
        q=3, 
        labels=['Bajo', 'Medio', 'Top']
    )
    
    return densidad_ingresos_zona, densidad_ingresos_empresa,ingresos_propietarios


# Función: Reordenar columnas
def reordenar_columnas(tabla):
    """
    Descripción: 
    Parametros: 
    Retorna:
    """
    """Reordena las columnas del DataFrame."""
    columnas_ordenadas = [
        'IdOportunidad', 'FolioOportunidad',
        'IdEmpresa', 'FolioEmpresa',
        'IdPropietario', 'FolioPropietario',
        'Zona', 'ClasificacionZona', 'TipoDivisaAjuste',
        'Importe', 'Importe_MXN', 'Importe_USD', 'Importe_EUR', 'RangoImporte',
        'FechaCierre', 'AnoCierre', 'MesCierre', 'TrimestreCierre', 'Participantes'
    ]
    return tabla[columnas_ordenadas]

def calcular_crecimientos(tabla):
    """
    Calcula el crecimiento porcentual anual por empresa, propietario y zona.
        - Compara importes entre 2023 y 2024.
    Parametros: 
        -tabla (pd.DataFrame): DataFrame enriquecido.
    Retorna:
        Tres DataFrames con el crecimiento calculado:
            - crecimiento_empresas
            - crecimiento_propietarios
            - crecimiento_zonas
    """

    # Calcular crecimiento anual de empresas
    crecimiento_empresas = tabla.groupby(['IdEmpresa', 'FolioEmpresa', 'AnoCierre'])['Importe_MXN'].sum().unstack(fill_value=0)
    crecimiento_empresas['Crecimiento_%'] = ((crecimiento_empresas[2024] - crecimiento_empresas[2023]) / crecimiento_empresas[2023]) * 100

    # Calcular crecimiento anual de propietarios
    crecimiento_propietarios = tabla.groupby(['IdPropietario', 'FolioPropietario', 'AnoCierre'])['Importe_MXN'].sum().unstack(fill_value=0)
    crecimiento_propietarios['Crecimiento_%'] = ((crecimiento_propietarios[2024] - crecimiento_propietarios[2023]) / crecimiento_propietarios[2023]) * 100

    # Calcular crecimiento anual por zonas
    crecimiento_zonas = tabla.groupby(['Zona', 'AnoCierre'], observed=False)['Importe_MXN'].sum().unstack(fill_value=0)
    crecimiento_zonas['Crecimiento_%'] = ((crecimiento_zonas[2024] - crecimiento_zonas[2023]) / crecimiento_zonas[2023]) * 100
    return crecimiento_empresas, crecimiento_propietarios, crecimiento_zonas

# Función: Reordenar columnas
def reordenar_columnas(tabla):
    """
    Reordena las columnas considerando las nuevas del DataFrame. 
    Parametros: 
        -tabla (pd.DataFrame): DataFrame enriquecido.
    Retorna:
        -tabla (pd.DataFrame): DataFrame enriquecido con las columnas ordenadas.
    """
    columnas_ordenadas = [
        'IdOportunidad', 'FolioOportunidad',
        'IdEmpresa', 'FolioEmpresa',
        'IdPropietario', 'FolioPropietario',
        'Zona', 'ClasificacionZona', 'TipoDivisaAjuste',
        'Importe', 'Importe_MXN', 'Importe_USD', 'Importe_EUR', 'RangoImporte',
        'FechaCierre', 'AnoCierre', 'MesCierre', 'TrimestreCierre', 'Participantes'
    ]
    return tabla[columnas_ordenadas]


def exportar_datos_con_formato(tabla, ruta_salida):
    """
    Exporta los datos a un archivo Excel.
    Parametros: 
        -tabla (pd.DataFrame): DataFrame enriquecido.
        -ruta_salida (str): Ruta del archivo Excel donde se guardarán los datos.
    Retorna:
    """
    # Crear estilos para columnas
    date_style = NamedStyle(name="date", number_format="YYYY-MM-DD")
    float_style = NamedStyle(name="float", number_format="#,##0.00")
    int_style = NamedStyle(name="int", number_format="0")

    # Crear archivo Excel
    with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
        tabla.to_excel(writer, index=False, sheet_name="Datos")
        workbook = writer.book
        worksheet = writer.sheets["Datos"]

        # Aplicar estilos de columnas basados en tipos de datos
        for col_idx, column in enumerate(tabla.columns, start=1):
            col_type = tabla[column].dtype
            if pd.api.types.is_datetime64_any_dtype(col_type):
                for row in range(2, len(tabla) + 2):  # Ajustar a las celdas de datos
                    worksheet.cell(row=row, column=col_idx).style = date_style
            elif pd.api.types.is_numeric_dtype(col_type):
                style = float_style if pd.api.types.is_float_dtype(col_type) else int_style
                for row in range(2, len(tabla) + 2):
                    worksheet.cell(row=row, column=col_idx).style = style

    print("Datos exportados con formatos explícitos para Excel.")


def agregar_hojas_crecimiento(tabla, ruta_salida,crecimiento_empresas, crecimiento_propietarios, crecimiento_zonas):
    """
    Agrega hojas para crecimientos al excel.
    Parametros: 
        -tabla (pd.DataFrame): DataFrame enriquecido.
        -ruta_salida (str): Rusta completa dl archivo CSV.
        Tres DataFrames con el crecimiento calculado:
            - crecimiento_empresas
            - crecimiento_propietarios
            - crecimiento_zonas
    Retorna:
    """
    # Agregar nuevas hojas al archivo Excel
    with pd.ExcelWriter(ruta_salida, engine='openpyxl', mode='a') as writer:
        crecimiento_empresas.to_excel(writer, sheet_name='Crecimiento_Empresas')
        crecimiento_propietarios.to_excel(writer, sheet_name='Crecimiento_Propietarios')
        crecimiento_zonas.to_excel(writer, sheet_name='Crecimiento_Zonas')

def exportar_datos_csv(tabla, ruta_salida):
    """
    Exporta los datos a un archivo CSV para evitar pérdida de formatos.
    Parametros: 
        -tabla (pd.DataFrame): DataFrame enriquecido.
        -ruta_salida (str): Ruta del archivo CSV donde se guardarán los datos.
    Retorna:
    """
    tabla.to_csv(ruta_salida, index=False, encoding='utf-8')
    print("Datos exportados exitosamente en formato CSV.")

# Bloque principal del script: organiza la ejecución del flujo ETL
# 1. Carga los datos desde un archivo CSV.
# 2. Limpia y transforma los datos.
# 3. Calcula métricas clave como densidades y crecimientos.
# 4. Exporta los resultados a un archivo Excel.
# 5. Genera hojas de crecimiento dentro del excel.
# 6. Exporta los resultados a un archivo csv.
if __name__ == "__main__":

    print("-----------------INICIANDO PROCESOS-----------")
    # Especificar rutas de entrada y salida
    ruta_archivo = r'C:\Users\andre\Downloads\Reto - Lider de datos\data\raw\BD_OPORTUNIDADES_23_24.csv'
    ruta_salida = r'C:\Users\andre\Downloads\Reto - Lider de datos\data\processed\_NUEVA_BD_OPORTUNIDADES_23_24.xlsx'
    ruta_salida_csv = r'C:\Users\andre\Downloads\Reto - Lider de datos\data\processed\_NUEVA_BD_OPORTUNIDADES_23_24.csv'
    # Tasas de cambio para normalización
    tasas_cambio = {
        'MXN': 1.0,
        'USD': 20.0,
        'EUR': 22.0,
        'GBP': 25.0
    }

    # Paso 1: Cargar los datos desde el archivo CSV
    print("Cargando datos de la base de datos BD_OPORTUNIDADES_23_24.csv .....")
    tabla = cargar_datos(ruta_archivo)
    print("Datos cargados correctamente.")
    
    # Paso 2: Limpiar los datos
    print("Limpiando datos .....")
    tabla = limpiar_datos(tabla)
    
    # Paso 3: Normalizar los datos monetarios y fechas
    print("Normalizando datos .....")
    tabla = normalizar_datos(tabla, tasas_cambio)

    # Paso 4: Generar columnas derivadas para el análisis
    print("Generando columnas derivadas .....")
    tabla = generar_columnas(tabla)
    
    # Paso 5: Calcular métricas agregadas (densidades y clasificaciones)
    print("Calculando agrupaciones .....")
    densidad_zona, densidad_empresa, densidad_propietarios = calcular_agrupaciones(tabla)

    # Ordenar los resultados de mayor a menor ingreso
    print("Generando resultados .....")
    densidad_zona = densidad_zona.sort_values(by='IngresoTotal', ascending=False)
    densidad_empresa = densidad_empresa.sort_values(by='IngresoTotal', ascending=False)
    densidad_propietarios = densidad_propietarios.sort_values(by='IngresoTotal', ascending=False)

    # Imprimir los resultados ordenados
    print("\nDensidad de ingresos por zona (ordenado de mayor a menor ingreso):")
    print(densidad_zona)

    print("\nDensidad de ingresos por empresa (ordenado de mayor a menor ingreso):")
    print(densidad_empresa)
    
    print("\nClasificación de propietarios (ordenado de mayor a menor ingreso):")
    print(densidad_propietarios)

    print("Generación de resultados terminada.")

    # Paso 6: Reorganizar columnas para una estructura más clara
    print("Reorganización de columnas .....")
    tabla = reordenar_columnas(tabla)
    
    # Paso 7: Exportar datos procesados a un archivo Excel
    print("Exportando datos al Excel _NUEVA_BD_OPORTUNIDADES_23_24.xlsx .....")
    exportar_datos_con_formato(tabla, ruta_salida)
    # Paso 8: Calcular y agregar hojas de crecimiento al archivo Excel
    print("Agregando hojas de cremientos al Excel _NUEVA_BD_OPORTUNIDADES_23_24.xlsx .....")
    crecimiento_empresas, crecimiento_propietarios, crecimiento_zonas = calcular_crecimientos(tabla)
    agregar_hojas_crecimiento(tabla, ruta_salida,crecimiento_empresas, crecimiento_propietarios, crecimiento_zonas)

    #Exportar datos a un cvs para evitar la perdida de formatos
    exportar_datos_csv(tabla, ruta_salida_csv)
    print("TODOS LOS PROCESOS FUERON EJECUTADOS CORRECTAMENTE.")