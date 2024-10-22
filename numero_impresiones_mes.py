import pdfplumber
import pandas as pd
from openpyxl import load_workbook
import os

# Función para extraer el número de impresiones
def extraer_numero_impresiones(texto):
    # Aquí busco un patrón específico en el texto (un ejemplo simplificado)
    if "Total" in texto:
        return texto.split("Total")[1].split()[0]
    return None

# Ruta a la carpeta donde están los archivos PDF (cambiar a la ruta que uses)
carpeta_pdfs = "ruta/a/tu/carpeta"

# Inicializar una lista para almacenar los resultados
resultados = []

# Procesar cada archivo PDF en la carpeta
for archivo in os.listdir(carpeta_pdfs):
    if archivo.endswith(".pdf"):
        ruta_pdf = os.path.join(carpeta_pdfs, archivo)
        with pdfplumber.open(ruta_pdf) as pdf:
            texto_completo = ""
            for pagina in pdf.pages:
                texto_completo += pagina.extract_text()

            # Extraer el número de impresiones
            numero_impresiones = extraer_numero_impresiones(texto_completo)
            if numero_impresiones:
                resultados.append((archivo, numero_impresiones))

# Convertir los resultados a un DataFrame de pandas
df = pd.DataFrame(resultados, columns=["Archivo", "Número de Impresiones"])

# Guardar los resultados en un archivo Excel
df.to_excel("resultado_impresiones.xlsx", index=False)
# Actualización menor
