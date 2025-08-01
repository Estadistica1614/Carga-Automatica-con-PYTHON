import os
import pandas as pd
import re
from PyPDF2 import PdfReader

# --- RUTAS ---
DIR_PARTES = r"C:\Users\ecastro\Desktop\PARTES"
BASE_SICPEF = r"C:\Users\ecastro\Desktop\SICPEF 2025.xlsx"
BASE_CAUSA = r"C:\Users\ecastro\Desktop\prueba_causa.xlsx"
RESULTADO = r"C:\Users\ecastro\Desktop\prueba_direcciones.xlsx"

# --- CARGAR COLUMNAS Y DATOS BASE ---
df_base = pd.read_excel(BASE_SICPEF)
columnas_finales = list(df_base.columns)

# Datos de causa ya procesados
df_causa = pd.read_excel(BASE_CAUSA)

# --- FUNCIONES ---
def leer_pdf(path_pdf):
    """Extrae texto de un PDF como string"""
    try:
        reader = PdfReader(path_pdf)
        texto = ""
        for page in reader.pages:
            texto += (page.extract_text() or "") + "\n"
        return texto
    except:
        return ""

def extraer_direcciones(texto):
    """Devuelve lista de direcciones (LUGAR n) encontradas en el PDF, tolerando saltos y espacios."""
    direcciones = []
    # Buscar bloques de LUGAR n (puede haber varios)
    bloques = re.findall(r"(LUGAR\s+\d+.*?)(?=(?:LUGAR\s+\d+|$))", texto, re.DOTALL | re.IGNORECASE)
    if not bloques:
        return [{
            "LUGAR": "-", "CALLE": "-", "LOCALIDAD": "-", 
            "PARTIDO": "-", "PROVINCIA": "-", "COORDENADAS": "-"
        }]
    for blo in bloques:
        lugar_data = {
            "LUGAR": "-", "CALLE": "-", "LOCALIDAD": "-", 
            "PARTIDO": "-", "PROVINCIA": "-", "COORDENADAS": "-"
        }
        # Número de lugar
        num = re.search(r"LUGAR\s+(\d+)", blo, re.IGNORECASE)
        if num:
            lugar_data["LUGAR"] = num.group(1).strip()

        # Patrones flexibles para capturar datos (tolerando saltos y mayúsculas)
        patrones = [
            (r"CALLE\s*[:\-]?\s*([\w\s\.,\-º°/]+)", "CALLE"),
            (r"LOCALIDAD\s*[:\-]?\s*([\w\s\.,\-º°/]+)", "LOCALIDAD"),
            (r"(?:DEPARTAMENTO|PARTIDO|COMUNA)\s*[:\-]?\s*([\w\s\.,\-º°/]+)", "PARTIDO"),
            (r"PROVINCIA\s*[:\-]?\s*([\w\s\.,\-º°/]+)", "PROVINCIA"),
            (r"COORDENADAS\s*[:\-]?\s*([\w\s\.,\-º°/]+)", "COORDENADAS")
        ]
        for patron, campo in patrones:
            m = re.search(patron, blo, re.IGNORECASE)
            if m:
                lugar_data[campo] = m.group(1).strip()

        direcciones.append(lugar_data)
    return direcciones

# --- PROCESAR PDFs ---
registros = []

for _, fila_causa in df_causa.iterrows():
    parte = str(fila_causa.get("PARTE OPERATIVO", "-"))
    if parte == "-" or not parte.strip():
        continue

    # PDF correspondiente
    pdf_path = os.path.join(DIR_PARTES, f"{parte}.pdf")
    if not os.path.exists(pdf_path):
        continue

    texto = leer_pdf(pdf_path)
    if not texto.strip():
        continue

    # Extraer direcciones con regex más flexible
    direcciones = extraer_direcciones(texto)

    for dir_data in direcciones:
        # Crear fila final combinando datos de causa y dirección
        fila = {col: "-" for col in columnas_finales}
        for col in df_causa.columns:
            if col in fila:
                fila[col] = fila_causa[col]
        for k, v in dir_data.items():
            fila[k] = v
        registros.append(fila)

# --- EXPORTAR RESULTADO ---
df_final = pd.DataFrame(registros, columns=columnas_finales)
df_final.to_excel(RESULTADO, index=False)
print(f"Proceso completado. Total filas con direcciones: {len(df_final)}")
print(f"Archivo generado en: {RESULTADO}")
