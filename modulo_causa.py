import os
import re
import pandas as pd
from PyPDF2 import PdfReader

# --- RUTAS ---
DIR_PARTES = r"C:\Users\ecastro\Desktop\PARTES"
RESULTADO = r"C:\Users\ecastro\Desktop\prueba_causa_raw.xlsx"

# --- FUNCIONES ---
def leer_pdf(path_pdf):
    """Extrae texto completo de un PDF"""
    try:
        reader = PdfReader(path_pdf)
        texto = ""
        for page in reader.pages:
            texto += (page.extract_text() or "") + "\n"
        return texto
    except:
        return ""

def extraer_datos(texto):
    """
    Extrae campos usando exactamente las expresiones regulares del flujo KNIME.
    Cada regex puede devolver None si el campo no está.
    """
    datos = {
        "PARTE_OPERATIVO": None,
        "CODIGO_DEPENDENCIA": None,
        "DEPENDENCIA": None,
        "FECHA": None,
        "HORA": None,
        "SUMARIO": None,
        "DELITO_1": None,
        "MODALIDAD_1": None,
        "DELITO_2": None,
        "MODALIDAD_2": None,
        "DELITO_3": None,
        "MODALIDAD_3": None,
        "TIPO_INTERVENCION": None,
        "JUZGADO_FISCALIA": None,
        "SECRETARIA": None,
        "CAUSA_NRO": None,
        "CARATULA": None,
        "EFECTIVOS": None,
        "MOVILES": None,
        "MOTOS": None,
        "CANES": None,
        "MORPHRAPID": None,
        "SCANNERS": None,
        "CABALLOS": None,
    }

    # Regex principales (las que agrupan PARTE OPERATIVO, CODIGO, DEPENDENCIA)
    m1 = re.search(r"PARTE OPERATIVO\s*:\s*\d+\s*-\s*PO\s*-\s*(\d+)\s*-\s*\d+\s*< >\s*CODIGO DE DEPENDENCIA:\s*(.*?)(?=<|$)< >\s*DEPENDENCIA:\s*(.*?)(?=<|$)< >", texto, re.DOTALL)
    if m1:
        datos["PARTE_OPERATIVO"], datos["CODIGO_DEPENDENCIA"], datos["DEPENDENCIA"] = m1.groups()

    # Fecha
    m2 = re.search(r"PARTE OPERATIVO.*?FECHA Y HORA:\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}|\d{2,4}[\/\-]\d{1,2}[\/\-]\d{1,2})", texto, re.DOTALL)
    if m2:
        datos["FECHA"] = m2.group(1)

    # Hora
    m3 = re.search(r"PARTE OPERATIVO.*?(\d{1,2}:\d{2})", texto, re.DOTALL)
    if m3:
        datos["HORA"] = m3.group(1)

    # Sumario
    m4 = re.search(r"PARTE OPERATIVO.*?SUMARIO:\s*(.*?)(?=<|$)<", texto, re.DOTALL)
    if m4:
        datos["SUMARIO"] = m4.group(1)

    # Delitos y modalidades
    m5 = re.search(r"DELITO 1:\s*([^<]*)<", texto)
    if m5: datos["DELITO_1"] = m5.group(1)

    m6 = re.search(r"MODALIDAD 1:\s*([^<]*)<", texto)
    if m6: datos["MODALIDAD_1"] = m6.group(1)

    # Patrones adicionales para DELITO 2-3, MODALIDAD 2-3
    patrones_extras = [
        (r"MODALIDAD 1:\s*([^<]*)<.*?DELITO 2:\s*([^<]*)<", ("MODALIDAD_1", "DELITO_2")),
        (r"MODALIDAD 1:\s*([^<]*)<.*?MODALIDAD 2:\s*([^<]*)<", ("MODALIDAD_1", "MODALIDAD_2")),
        (r"MODALIDAD 1:\s*([^<]*)<.*?DELITO 3:\s*([^<]*)<", ("MODALIDAD_1", "DELITO_3")),
        (r"MODALIDAD 1:\s*([^<]*)<.*?MODALIDAD 3:\s*([^<]*)<", ("MODALIDAD_1", "MODALIDAD_3")),
    ]
    for patron, campos in patrones_extras:
        m = re.search(patron, texto, re.DOTALL)
        if m and len(campos) == 2:
            datos[campos[1]] = m.group(2)

    # Otros campos (intervención, juzgado, etc.)
    campos_individuales = [
        (r"TIPO DE INTERVENCION:\s*([^<]*)<", "TIPO_INTERVENCION"),
        (r"JUZGADO\s*/\s*FISCALIA:\s*([^<]*)<", "JUZGADO_FISCALIA"),
        (r"SECRETARIA:\s*([^<>]*)<", "SECRETARIA"),
        (r"CAUSA NRO.:\s*([^<]*)<", "CAUSA_NRO"),
        (r"CARATULA.:\s*([^<]*)<", "CARATULA"),
    ]
    for patron, campo in campos_individuales:
        m = re.search(patron, texto, re.DOTALL)
        if m:
            datos[campo] = m.group(1)

    # Recursos (efectivos, móviles, etc.)
    recursos = re.search(r"EFECTIVOS:\s*(.*?)(?=<|$)< >\s*MOVILES:\s*(.*?)(?=<|$)< >\s*MOTOS:\s*(.*?)(?=<|$)< >\s*CANES:\s*(.*?)(?=<|$)< >\s*MORPHRAPID:\s*(.*?)(?=<|$)< >\s*SCANNERS:\s*(.*?)(?=<|$)< >\s*CABALLOS:\s*(.*?)(?=<|$)<", texto, re.DOTALL)
    if recursos:
        datos["EFECTIVOS"], datos["MOVILES"], datos["MOTOS"], datos["CANES"], datos["MORPHRAPID"], datos["SCANNERS"], datos["CABALLOS"] = recursos.groups()

    return datos

# --- PROCESAR TODOS LOS PDFs ---
registros = []

for archivo in os.listdir(DIR_PARTES):
    if not archivo.lower().endswith(".pdf"):
        continue

    texto = leer_pdf(os.path.join(DIR_PARTES, archivo))
    if not texto.strip():
        continue

    datos_pdf = extraer_datos(texto)
    datos_pdf["ARCHIVO"] = archivo  # Guardamos el nombre del archivo para referencia
    registros.append(datos_pdf)

# --- EXPORTAR RESULTADOS ---
df = pd.DataFrame(registros)
df.to_excel(RESULTADO, index=False)
print(f"Procesados {len(df)} archivos. Resultado en: {RESULTADO}")
