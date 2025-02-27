import pandas as pd
import os
import gc  # Importa el módulo para la recolección de basura
import time
from tqdm import tqdm

def excel_to_markdown_large_data(input_excel_path, output_md_path, chunk_size=17000):
    """
    Convierte un archivo Excel con grandes cantidades de registros a Markdown de manera eficiente.

    Args:
        input_excel_path (str): Ruta al archivo Excel de entrada.
        output_md_path (str): Ruta al archivo Markdown de salida.
        chunk_size (int): Número de filas a leer por cada fragmento (chunk).
    """
    start_time = time.time()
    try:
        # Diccionario de descripciones
        descriptions = {
            
            
            
            'AGNO': 'Año',
            'RBD': 'RBD',
            'DGV_RBD': 'Dígito RBD',
            'NOM_RBD': 'Nom Establecimiento',
            'MRUN': 'MRUN',
            'RUT_SOSTENEDOR': 'RUT Sosten',
            'P_JURIDICA': 'Perso Jurídica',
            'COD_REG_RBD': 'Cod Reg',
            'NOM_REG_RBD_A': 'Nom Reg',
            'COD_PRO_RBD': 'Cod Prov',
            'COD_COM_RBD': 'Cod Com',
            'NOM_COM_RBD': 'Nom Comuna',
            'COD_DEPROV_RBD': 'Cod Dpto Prov',
            'NOM_DEPROV_RBD': 'Nom Depto Prov',
            'COD_DEPE': 'Cod Dependencia',
            'COD_DEPE2': 'Cod Dependencia 2',
            'RURAL_RBD': 'Rural (1) / Urbano (0)',
            'LATITUD': 'Lat',
            'LONGITUD': 'Long',
            'CONVENIO_PIE': 'Convenio PIE',
            'PACE': 'Prog PACE',
            'ENS_01': 'Enseñanza 01',
            'ENS_02': 'Enseñanza 02',
            'ENS_03': 'Enseñanza 03',
            'ENS_04': 'Enseñanza 04',
            'ENS_05': 'Enseñanza 05',
            'ENS_06': 'Enseñanza 06',
            'ENS_07': 'Enseñanza 07',
            'ENS_08': 'Enseñanza 08',
            'ENS_09': 'Enseñanza 09',
            'ENS_10': 'Enseñanza 10',
            'ENS_11': 'Enseñanza 11',
            'MAT_TOTAL': 'Matrícula Total',
            'MATRICULA': 'Matriculados',
            'ESTADO_ESTAB': 'Estado Establecimiento',
            'ORI_RELIGIOSA': 'Orienta Religiosa',
            'ORI_OTRO_GLOSA': 'Glosa Otra Orienta',
            'PAGO_MATRICULA': 'Pago Matrícula',
            'PAGO_MENSUAL': 'Pago Mensual',
            'ESPE_01': 'Especialidad 01',
            'ESPE_02': 'Especialidad 02',
            'ESPE_03': 'Especialidad 03',
            'ESPE_04': 'Especialidad 04',
            'ESPE_05': 'Especialidad 05',
            'ESPE_06': 'Especialidad 06',
            'ESPE_07': 'Especialidad 07',
            'ESPE_08': 'Especialidad 08',
            'ESPE_09': 'Especialidad 09',
            'ESPE_10': 'Especialidad 10',
            'ESPE_11': 'Especialidad 11'
        }

        # Inicializar el archivo Markdown
        with open(output_md_path, "w", encoding="utf-8") as md_file:
            md_file.write("#Registros estadísticos de establecimientos educativos\n\n")

        # Contar el total de filas del archivo Excel
        print(f"Contando número de registros en el archivo {input_excel_path}...")
        xl = pd.ExcelFile(input_excel_path)
        total_rows = len(pd.read_excel(xl))
        print(f"Total de registros: {total_rows}")
        del xl #Liberar memoria
        gc.collect()

        # Iterar por fragmentos del archivo Excel
        print("Iniciando procesamiento por fragmentos...")
        for chunk_idx, chunk_start in enumerate(tqdm(range(0, total_rows, chunk_size), desc="Procesando fragmentos", unit="fragmento")):
            chunk_end = min(chunk_start + chunk_size, total_rows)
            print(f"Procesando registros {chunk_start + 1}-{chunk_end} de {total_rows}...")

            # Leer el fragmento
            df_chunk = pd.read_excel(input_excel_path, skiprows=chunk_start, nrows=chunk_end - chunk_start)

            # Procesar el fragmento
            process_dataframe_chunk(df_chunk, descriptions, output_md_path, chunk_start)

            # Liberar memoria
            del df_chunk
            gc.collect()

        end_time = time.time()
        duration = end_time - start_time
        print(f"\n✅ Conversión completa. Tiempo total: {duration:.2f} segundos")
        print(f"✅ Markdown generado correctamente en: {output_md_path}")

    except Exception as e:
        print(f"❌ Error durante la conversión: {str(e)}")

def process_dataframe_chunk(df_chunk, descriptions, output_md_path, chunk_start):
    """
    Procesa un fragmento del DataFrame y lo guarda en el archivo Markdown.

    Args:
        df_chunk (DataFrame): Fragmento del DataFrame.
        descriptions (dict): Diccionario de descripciones de las columnas.
        output_md_path (str): Ruta del archivo Markdown de salida.
        chunk_start (int): Índice de inicio del fragmento.
    """

    def safe_get(row, column):
        value = row.get(column, 'N/A')
        if pd.isna(value):
            return 'N/A'
        return value

    # Abrir el archivo Markdown en modo append
    with open(output_md_path, "a", encoding="utf-8") as md_file:
        # Iterar por cada fila del fragmento
        for idx, row in df_chunk.iterrows():
            markdown_content = f"""
## 📄 Establecimiento {chunk_start + idx + 1}

| **Campo** | **Valor** |
|-----------|-----------|
"""
            # Agregar cada campo al Markdown
            for column in df_chunk.columns:
                if column in descriptions:
                    markdown_content += f"| {descriptions[column]} | {safe_get(row, column)} |\n"

            markdown_content += "\n---\n\n"
            md_file.write(markdown_content)


# Definir la ruta de entrada y salida
input_excel_path = r"/content/20240912_Directorio_Oficial_EE_2024_20240430_WEB.xlsx"
output_md_path = r"/content/output.md"

# Ejecutar la función de conversión
excel_to_markdown_large_data(input_excel_path, output_md_path)
