# ExceltoMarkdown-LargeData
Transformación de Excel BBDD Histórica Matriculas 2004-2024 a Markdown
- El trabajo fue realizado utilizando como base el código Python creado por Carlos Morales: https://github.com/divpoliticas/UdE_Neo4J/blob/main/Deprecated/Convert%20semantic%2Bindex%20search%20in%20markdown.py
- El Excel transformado se obtuvo del portal: https://datosabiertos.mineduc.cl/
- Link portal Mineduc: https://datosabiertos.mineduc.cl/directorio-de-establecimientos-educacionales/
- Link Excel: https://datosabiertos.mineduc.cl/wp-content/uploads/2024/11/Directorio-Oficial-EE-2024-.rar

# Cabeceras procesadas
```python
- 'AGNO': 'Año'
- 'RBD': 'RBD'
- 'DGV_RBD': 'Dígito RBD'
- 'NOM_RBD': 'Nom Establecimiento'
- 'MRUN': 'MRUN'
- 'RUT_SOSTENEDOR': 'RUT Sosten'
- 'P_JURIDICA': 'Perso Jurídica'
- 'COD_REG_RBD': 'Cod Reg'
- 'NOM_REG_RBD_A': 'Nom Reg'
- 'COD_PRO_RBD': 'Cod Prov'
- 'COD_COM_RBD': 'Cod Com'
- 'NOM_COM_RBD': 'Nom Comuna'
- 'COD_DEPROV_RBD': 'Cod Dpto Prov'
- 'NOM_DEPROV_RBD': 'Nom Depto Prov'
- 'COD_DEPE': 'Cod Dependencia'
- 'COD_DEPE2': 'Cod Dependencia 2'
- 'RURAL_RBD': 'Rural (1) / Urbano (0)'
- 'LATITUD': 'Lat'
- 'LONGITUD': 'Long'
- 'CONVENIO_PIE': 'Convenio PIE'
- 'PACE': 'Prog PACE'
- 'ENS_01': 'Enseñanza 01'
- 'ENS_02': 'Enseñanza 02'
- 'ENS_03': 'Enseñanza 03'
- 'ENS_04': 'Enseñanza 04'
- 'ENS_05': 'Enseñanza 05'
- 'ENS_06': 'Enseñanza 06'
- 'ENS_07': 'Enseñanza 07'
- 'ENS_08': 'Enseñanza 08'
- 'ENS_09': 'Enseñanza 09'
- 'ENS_10': 'Enseñanza 10'
- 'ENS_11': 'Enseñanza 11'
- 'MAT_TOTAL': 'Matrícula Total'
- 'MATRICULA': 'Matriculados'
- 'ESTADO_ESTAB': 'Estado Establecimiento'
- 'ORI_RELIGIOSA': 'Orienta Religiosa'
- 'ORI_OTRO_GLOSA': 'Glosa Otra Orienta'
- 'PAGO_MATRICULA': 'Pago Matrícula'
- 'PAGO_MENSUAL': 'Pago Mensual'
- 'ESPE_01': 'Especialidad 01'
- 'ESPE_02': 'Especialidad 02'
- 'ESPE_03': 'Especialidad 03'
- 'ESPE_04': 'Especialidad 04'
- 'ESPE_05': 'Especialidad 05'
- 'ESPE_06': 'Especialidad 06'
- 'ESPE_07': 'Especialidad 07'
- 'ESPE_08': 'Especialidad 08'
- 'ESPE_09': 'Especialidad 09'
- 'ESPE_10': 'Especialidad 10'
- 'ESPE_11': 'Especialidad 11'
```

# Descripción del Código Python: Conversión de Excel a Markdown

Este documento describe en detalle cada actividad realizada en el script Python que convierte un archivo Excel en un archivo Markdown (.md). El código está diseñado para manejar grandes volúmenes de datos de manera eficiente, procesando los datos en fragmentos (chunks) y generando un archivo Markdown estructurado.

---

## 📄 Actividades Principales

### 1. **Importación de Bibliotecas**
El código comienza importando las bibliotecas necesarias para su funcionamiento:

- **`pandas`**: Para leer y manipular el archivo Excel.
- **`os`**: Para operaciones del sistema de archivos (aunque no se utiliza directamente en el código proporcionado).
- **`gc`**: Para la recolección de basura y liberación de memoria.
- **`time`**: Para medir el tiempo de ejecución del script.
- **`tqdm`**: Para mostrar una barra de progreso durante el procesamiento de los fragmentos.

```python
import pandas as pd
import os
import gc
import time
from tqdm import tqdm
```

---

### 2. **Definición de la Función Principal: `excel_to_markdown_large_data`**
Esta función es el núcleo del script y realiza las siguientes actividades:

#### Proceso:

1. **Inicialización**: Se crea un archivo Markdown y se escribe un encabezado.
2. **Conteo de Filas**: Se cuenta el número total de filas en el archivo Excel.
3. **Procesamiento por Fragmentos**: El archivo se procesa en fragmentos para evitar el desbordamiento de memoria.
4. **Escritura en Markdown**: Cada fila se convierte en una sección Markdown con una tabla que muestra los campos y sus valores.
5. **Liberación de Memoria**: Después de procesar cada fragmento, se libera la memoria utilizada.

#### **Parámetros de Entrada**
- **`input_excel_path`**: Ruta del archivo Excel de entrada.
- **`output_md_path`**: Ruta del archivo Markdown de salida.
- **`chunk_size`**: Número de filas a procesar por fragmento (por defecto: 17000).

#### **Actividades Realizadas**
1. **Inicialización del Tiempo de Ejecución**:
   - Se inicia un contador de tiempo para medir la duración total del proceso.

   ```python
   start_time = time.time()
   ```

2. **Definición del Diccionario de Descripciones**:
   - Se crea un diccionario (`descriptions`) que mapea los nombres de las columnas del Excel a descripciones más legibles. Este diccionario se utiliza para generar las tablas en el archivo Markdown.

   ```python
   descriptions = {
       'AGNO': 'Año',
       'RBD': 'RBD',
       ...
   }
   ```

3. **Inicialización del Archivo Markdown**:
   - Se crea el archivo Markdown y se escribe un encabezado inicial.

   ```python
   with open(output_md_path, "w", encoding="utf-8") as md_file:
       md_file.write("# Registros estadísticos de establecimientos educativos\n\n")
   ```

4. **Conteo de Filas del Archivo Excel**:
   - Se cuenta el número total de filas en el archivo Excel para determinar cuántos fragmentos se deben procesar.

   ```python
   xl = pd.ExcelFile(input_excel_path)
   total_rows = len(pd.read_excel(xl))
   ```

5. **Liberación de Memoria**:
   - Se libera la memoria utilizada para leer el archivo Excel.

   ```python
   del xl
   gc.collect()
   ```

6. **Procesamiento por Fragmentos**:
   - El archivo Excel se procesa en fragmentos para evitar el desbordamiento de memoria. Para cada fragmento:
     - Se lee un conjunto de filas (`chunk_size`).
     - Se procesa el fragmento utilizando la función `process_dataframe_chunk`.
     - Se libera la memoria del fragmento procesado.

   ```python
   for chunk_idx, chunk_start in enumerate(
           tqdm(range(0, total_rows, chunk_size), desc="Procesando fragmentos", unit="fragmento")):
       chunk_end = min(chunk_start + chunk_size, total_rows)
       df_chunk = pd.read_excel(input_excel_path, skiprows=chunk_start, nrows=chunk_end - chunk_start)
       process_dataframe_chunk(df_chunk, descriptions, output_md_path, chunk_start)
       del df_chunk
       gc.collect()
   ```

7. **Finalización del Proceso**:
   - Se calcula el tiempo total de ejecución y se imprime un mensaje de finalización.

   ```python
   end_time = time.time()
   duration = end_time - start_time
   print(f"\n✅ Conversión completa. Tiempo total: {duration:.2f} segundos")
   ```

---

### 3. **Definición de la Función Secundaria: `process_dataframe_chunk`**
Esta función procesa cada fragmento del DataFrame y lo convierte en contenido Markdown.

#### Proceso:

1. **Iteración por Filas**: Cada fila del fragmento se convierte en una sección Markdown.
2. **Generación de Tablas**: Los campos y valores de cada fila se presentan en una tabla Markdown.
3. **Escritura en Archivo**: El contenido generado se escribe en el archivo Markdown.

#### **Parámetros de Entrada**
- **`df_chunk`**: Fragmento del DataFrame.
- **`descriptions`**: Diccionario de descripciones de las columnas.
- **`output_md_path`**: Ruta del archivo Markdown de salida.
- **`chunk_start`**: Índice de inicio del fragmento.

#### **Actividades Realizadas**
1. **Definición de la Función `safe_get`**:
   - Esta función auxiliar se utiliza para manejar valores nulos o faltantes en el DataFrame, devolviendo 'N/A' en su lugar.

   ```python
   def safe_get(row, column):
       value = row.get(column, 'N/A')
       if pd.isna(value):
           return 'N/A'
       return value
   ```

2. **Apertura del Archivo Markdown en Modo Append**:
   - Se abre el archivo Markdown en modo de escritura adicional (`append`) para agregar el contenido generado.

   ```python
   with open(output_md_path, "a", encoding="utf-8") as md_file:
   ```

3. **Iteración por Filas del Fragmento**:
   - Para cada fila del fragmento:
     - Se genera una sección Markdown con un encabezado que indica el número de establecimiento.
     - Se crea una tabla Markdown con los campos y valores de la fila.

   ```python
   for idx, row in df_chunk.iterrows():
       markdown_content = f"""
   ## 📄 Establecimiento {chunk_start + idx + 1}

   | **Campo** | **Valor** |
   |-----------|-----------|
   """
       for column in df_chunk.columns:
           if column in descriptions:
               markdown_content += f"| {descriptions[column]} | {safe_get(row, column)} |\n"
       markdown_content += "\n---\n\n"
       md_file.write(markdown_content)
   ```

---

### 4. **Ejecución del Script**
Finalmente, el script define las rutas de entrada y salida, y ejecuta la función principal.

```python
input_excel_path = r"C:/Users/igajardo/Downloads/20240912_Directorio_Oficial_EE_2024_20240430_WEB.xlsx"
output_md_path = r"C:/Users/igajardo/Downloads/output.md"
excel_to_markdown_large_data(input_excel_path, output_md_path)
```

---

## 📌 Resumen de Actividades

1. **Importación de Bibliotecas**: Se importan las bibliotecas necesarias.
2. **Definición de la Función Principal**: Se define la función `excel_to_markdown_large_data` para manejar la conversión.
3. **Procesamiento por Fragmentos**: El archivo Excel se procesa en fragmentos para evitar problemas de memoria.
4. **Generación de Markdown**: Cada fila del Excel se convierte en una sección Markdown con una tabla.
5. **Manejo de Memoria**: Se libera memoria después de procesar cada fragmento.
6. **Ejecución del Script**: Se definen las rutas y se ejecuta la función principal.

---

## ✅ Resultado Final

El archivo Markdown generado contendrá una sección para cada fila del Excel, con una tabla que muestra los campos y sus valores correspondientes.

```markdown
# Registros estadísticos de establecimientos educativos

## 📄 Establecimiento 1

| **Campo** | **Valor** |
|-----------|-----------|
| Año       | 2024      |
| RBD       | 12345     |
| ...       | ...       |

---

## 📄 Establecimiento 2

| **Campo** | **Valor** |
|-----------|-----------|
| Año       | 2024      |
| RBD       | 67890     |
| ...       | ...       |

---
```


## 📝 Ejemplo de Uso

```python
# Definir la ruta de entrada y salida
input_excel_path = r"C:/Users/igajardo/Downloads/20240912_Directorio_Oficial_EE_2024_20240430_WEB.xlsx"
output_md_path = r"C:/Users/igajardo/Downloads/output.md"

# Ejecutar la función de conversión
excel_to_mark

