import os
import re
import pandas as pd
from PyPDF2 import PdfReader

# --- CONFIGURACIÓN ---
DIRECTORIO_PDFS = r"C:\Users\ecastro\Desktop\PARTES"
SALIDA_EXCEL = r"C:\Users\ecastro\Desktop\resultado_detallado_corregido.xlsx"

# --- FUNCIONES DE APOYO ---
def limpiar_dni(dni):
    return re.sub(r"\D", "", dni)

def a_mayusculas(valor):
    if isinstance(valor, str):
        return valor.strip().upper()
    return valor

def extraer_unico(patron, texto, limpiar=None):
    match = re.search(patron, texto, re.IGNORECASE | re.DOTALL)
    if match:
        dato = match.group(1).strip()
        if limpiar:
            dato = limpiar(dato)
        return dato
    return ""

def extraer_todos(patron, texto, limpiar=None):
    matches = re.findall(patron, texto, re.IGNORECASE | re.DOTALL)
    resultados = []
    for dato in matches:
        dato = dato.strip()
        if limpiar:
            dato = limpiar(dato)
        resultados.append(dato)
    return resultados

def extraer_bloques_con_lugar(tipo, texto):
    """
    Extrae bloques de un tipo (ARMA, DROGA, ELEMENTO, IMPUTADO, VEHICULO)
    y asocia el número del LUGAR más cercano antes del bloque.
    Filtra bloques sin campos clave.
    """
    bloques = re.findall(
        rf"({tipo} .*?)(?=IMPUTADO|VICTIMA|DROGA|ELEMENTO|VEHICULO|ARMA|$)",
        texto, re.IGNORECASE | re.DOTALL
    )
    resultados = []
    for bloque in bloques:
        # Filtrar si no hay ninguna pista de datos útiles
        if not re.search(r"(Tipo:|Nombres:|Incautacion:|Marca:)", bloque, re.IGNORECASE):
            continue
        # Buscar el último LUGAR N antes del bloque
        idx = texto.find(bloque)
        lugar_match = re.findall(r"LUGAR\s+(\d+)", texto[:idx], re.IGNORECASE)
        lugar_nro = lugar_match[-1] if lugar_match else "1"
        resultados.append((bloque, lugar_nro))
    return resultados

# --- TABLAS ACUMULADAS ---
cabeceras, lugares, armas, drogas, elementos, imputados, victimas, vehiculos, otros = ([] for _ in range(9))

# --- PROCESAR TODOS LOS PDF ---
for archivo in os.listdir(DIRECTORIO_PDFS):
    if not archivo.lower().endswith(".pdf"):
        continue

    ruta_pdf = os.path.join(DIRECTORIO_PDFS, archivo)
    print(f"Procesando: {archivo}")

    reader = PdfReader(ruta_pdf)
    texto = ""
    for page in reader.pages:
        texto += page.extract_text()

    # --- CABECERA ---
    fecha_hora = extraer_unico(r"Fecha y Hora:\s*([\d\-]+\s*-\s*\d{2}:\d{2})", texto)
    fecha, hora = "", ""
    if fecha_hora:
        partes = fecha_hora.split("-")
        if len(partes) >= 3:
            d, m, y = partes[0].strip(), partes[1].strip(), partes[2].strip().split()[0]
            fecha = f"{y}-{m}-{d}"  # YYYY-MM-DD
            hora = partes[-1].strip()

    cabeceras.append({
        "Archivo": archivo,
        "Parte Operativo": a_mayusculas(extraer_unico(r"Parte Operativo:\s*([\w\-]+)", texto)),
        "Código Dependencia": a_mayusculas(extraer_unico(r"Codigo de Dependencia:\s*(\d+)", texto)),
        "Dependencia": a_mayusculas(extraer_unico(r"Dependencia:\s*(.+?)\s*<", texto)),
        "Fecha": fecha,
        "Hora": hora,
        "Sumario": a_mayusculas(extraer_unico(r"Sumario:\s*(.+?)\s*<", texto)),
        "Delito": a_mayusculas(extraer_unico(r"Delito 1:\s*(.+?)\s*<", texto)),
        "Modalidad": a_mayusculas(extraer_unico(r"Modalidad 1:\s*(.+?)\s*<", texto)),
        "Tipo Intervención": a_mayusculas(extraer_unico(r"Tipo de Intervencion:\s*(.+?)\s*<", texto)),
        "Juzgado / Fiscalía": a_mayusculas(extraer_unico(r"Juzgado / Fiscalia:\s*(.+?)\s*<", texto)),
        "Secretaría": a_mayusculas(extraer_unico(r"Secretaria:\s*(.+?)\s*<", texto)),
        "Causa Nro.": a_mayusculas(extraer_unico(r"Causa Nro.:\s*(.+?)\s*<", texto)),
        "Carátula": a_mayusculas(extraer_unico(r"Caratula:\s*(.+?)\s*<", texto)),
    })

    # --- LUGARES ---
    calles = extraer_todos(r"Calle:\s*(.+?)\s*<", texto)
    localidades = extraer_todos(r"Localidad:\s*(.+?)\s*<", texto)
    departamentos = extraer_todos(r"Departamento / Partido / Comuna:\s*(.+?)\s*<", texto)
    provincias = extraer_todos(r"Provincia:\s*(.+?)\s*<", texto)
    coords = extraer_todos(r"Coordenadas:\s*([^\n<]+)", texto)
    for i in range(len(calles)):
        lugares.append({
            "Archivo": archivo,
            "Lugar Nro": i+1,
            "Calle": a_mayusculas(calles[i]) if i < len(calles) else "",
            "Localidad": a_mayusculas(localidades[i]) if i < len(localidades) else "",
            "Departamento / Comuna": a_mayusculas(departamentos[i]) if i < len(departamentos) else "",
            "Provincia": a_mayusculas(provincias[i]) if i < len(provincias) else "",
            "Coordenadas": coords[i] if i < len(coords) else "",
        })

    # --- ARMAS ---
    for bloque, lugar in extraer_bloques_con_lugar("ARMA", texto):
        tipo = a_mayusculas(extraer_unico(r"Tipo:\s*([^\n<]+)", bloque))
        detalles = a_mayusculas(extraer_unico(r"Detalles:\s*([^\n<]+)", bloque))
        marca = a_mayusculas(extraer_unico(r"Marca:\s*([^\n<]+)", bloque))
        modelo = a_mayusculas(extraer_unico(r"Modelo:\s*([^\n<]+)", bloque))
        calibre = a_mayusculas(extraer_unico(r"Calibre:\s*([^\n<]+)", bloque))
        numeracion = a_mayusculas(extraer_unico(r"Numeracion:\s*([^\n<]+)", bloque))
        secuestro = a_mayusculas(extraer_unico(r"Pedido de Secuestro:\s*([^\n<]+)", bloque))
        observaciones = a_mayusculas(extraer_unico(r"Observaciones:\s*(.+?)\s*(?=<|$)", bloque))

        if not any([tipo, detalles, marca, modelo, calibre, numeracion, secuestro, observaciones]):
            continue

        armas.append({
            "Archivo": archivo,
            "Lugar Nro": lugar,
            "Tipo": tipo,
            "Detalles": detalles,
            "Marca": marca,
            "Modelo": modelo,
            "Calibre": calibre,
            "Numeración": numeracion,
            "Pedido de Secuestro": secuestro,
            "Observaciones": observaciones,
        })

    # --- DROGAS ---
    for bloque, lugar in extraer_bloques_con_lugar("DROGA", texto):
        tipo = a_mayusculas(extraer_unico(r"Tipo:\s*([^\n<]+)", bloque))
        cantidad = extraer_unico(r"Cantidad:\s*([\d.,]+)", bloque)
        medicion = a_mayusculas(extraer_unico(r"Medicion:\s*([^\n<]+)", bloque))
        observaciones = a_mayusculas(extraer_unico(r"Observaciones:\s*(.+?)\s*(?=<|$)", bloque))

        if not any([tipo, cantidad, medicion, observaciones]):
            continue

        drogas.append({
            "Archivo": archivo,
            "Lugar Nro": lugar,
            "Tipo": tipo,
            "Cantidad": cantidad,
            "Medición": medicion,
            "Observaciones": observaciones,
        })

    # --- ELEMENTOS ---
    for bloque, lugar in extraer_bloques_con_lugar("ELEMENTO", texto):
        incautacion = a_mayusculas(extraer_unico(r"Incautacion:\s*([^\n<]+)", bloque))
        tipo = a_mayusculas(extraer_unico(r"Tipo:\s*([^\n<]+)", bloque))
        subtipo = a_mayusculas(extraer_unico(r"Subtipo:\s*([^\n<]+)", bloque))
        cantidad = extraer_unico(r"Cantidad:\s*([\d.,]+)", bloque)
        medicion = a_mayusculas(extraer_unico(r"Medicion:\s*([^\n<]+)", bloque))
        aforo = extraer_unico(r"Aforo:\$([\d.,]*)", bloque)
        observaciones = a_mayusculas(extraer_unico(r"Observaciones:\s*(.+?)\s*(?=<|$)", bloque))

        if not any([incautacion, tipo, subtipo, cantidad, medicion, aforo, observaciones]):
            continue

        elementos.append({
            "Archivo": archivo,
            "Lugar Nro": lugar,
            "Incautación": incautacion,
            "Tipo": tipo,
            "Subtipo": subtipo,
            "Cantidad": cantidad,
            "Medición": medicion,
            "Aforo": aforo,
            "Observaciones": observaciones,
        })

    # --- IMPUTADOS ---
    for bloque, lugar in extraer_bloques_con_lugar("IMPUTADO", texto):
        nombres = a_mayusculas(extraer_unico(r"Nombres:\s*([^\n<]+)", bloque))
        apellidos = a_mayusculas(extraer_unico(r"Apellidos:\s*([^\n<]+)", bloque))
        edad = extraer_unico(r"Edad:\s*(\d+)", bloque)
        genero = a_mayusculas(extraer_unico(r"Genero:\s*([^\n<]+)", bloque))
        dni = extraer_unico(r"DNI:\s*([.\d]+)", bloque, limpiar=limpiar_dni)
        nacionalidad = a_mayusculas(extraer_unico(r"Nacionalidad:\s*([^\n<]+)", bloque))
        domicilio = a_mayusculas(extraer_unico(r"Domicilio:\s*([^\n<]+)", bloque))
        situacion = a_mayusculas(extraer_unico(r"Situacion\s+Procesal\s*:\s*([^\n<]+)", bloque))
        captura = a_mayusculas(extraer_unico(r"Posee\s+Captura:\s*([^\n<]+)", bloque))
        motivo = a_mayusculas(extraer_unico(r"Motivo del Pedido de Captura:\s*([^\n<]+)", bloque))

        if not any([nombres, apellidos, edad, genero, dni, nacionalidad, domicilio, situacion, captura, motivo]):
            continue

        imputados.append({
            "Archivo": archivo,
            "Lugar Nro": lugar,
            "Nombres": nombres,
            "Apellidos": apellidos,
            "Edad": edad,
            "Género": genero,
            "DNI": dni,
            "Nacionalidad": nacionalidad,
            "Domicilio": domicilio,
            "Situación Procesal": situacion,
            "Posee Captura": captura,
            "Motivo Captura": motivo,
        })

    # --- VÍCTIMAS ---
    for bloque in re.findall(r"(VICTIMA .*?)(?=IMPUTADO|VICTIMA|DROGA|ELEMENTO|VEHICULO|ARMA|$)", texto, re.IGNORECASE | re.DOTALL):
        nombres = a_mayusculas(extraer_unico(r"Nombres:\s*([^\n<]+)", bloque))
        apellidos = a_mayusculas(extraer_unico(r"Apellidos:\s*([^\n<]+)", bloque))
        edad = extraer_unico(r"Edad:\s*(\d+)", bloque)
        genero = a_mayusculas(extraer_unico(r"Genero:\s*([^\n<]+)", bloque))
        dni = extraer_unico(r"DNI:\s*([.\d]+)", bloque, limpiar=limpiar_dni)
        nacionalidad = a_mayusculas(extraer_unico(r"Nacionalidad:\s*([^\n<]+)", bloque))
        domicilio = a_mayusculas(extraer_unico(r"Domicilio:\s*([^\n<]+)", bloque))

        if not any([nombres, apellidos, edad, genero, dni, nacionalidad, domicilio]):
            continue

        victimas.append({
            "Archivo": archivo,
            "Nombres": nombres,
            "Apellidos": apellidos,
            "Edad": edad,
            "Género": genero,
            "DNI": dni,
            "Nacionalidad": nacionalidad,
            "Domicilio": domicilio,
        })

    # --- VEHÍCULOS ---
    for bloque, lugar in extraer_bloques_con_lugar("VEHICULO", texto):
        marca = a_mayusculas(extraer_unico(r"Marca:\s*([^\n<]+)", bloque))
        modelo = a_mayusculas(extraer_unico(r"Modelo:\s*([^\n<]+)", bloque))
        dominio = a_mayusculas(extraer_unico(r"Dominio:\s*([^\n<]+)", bloque))
        tipo = a_mayusculas(extraer_unico(r"Tipo:\s*([^\n<]+)", bloque))
        detalles = a_mayusculas(extraer_unico(r"Detalles:\s*(.+?)\s*(?=<|$)", bloque))

        if not any([marca, modelo, dominio, tipo, detalles]):
            continue

        vehiculos.append({
            "Archivo": archivo,
            "Lugar Nro": lugar,
            "Marca": marca,
            "Modelo": modelo,
            "Dominio": dominio,
            "Tipo": tipo,
            "Detalles": detalles,
        })

    # --- OTROS ---
    otros.append({
        "Archivo": archivo,
        "Efectivos": extraer_unico(r"Efectivos:\s*(\d+)", texto),
        "Moviles": extraer_unico(r"Moviles:\s*(\d+)", texto),
        "Motos": extraer_unico(r"Motos:\s*(\d+)", texto),
        "Canes": extraer_unico(r"Canes:\s*(\d+)", texto),
        "Morphrapid": extraer_unico(r"Morphrapid:\s*(\d+)", texto),
        "Scanners": extraer_unico(r"Scanners:\s*(\d+)", texto),
        "Caballos": extraer_unico(r"Caballos:\s*(\d+)", texto),
    })

# --- GUARDAR EN VARIAS HOJAS ---
with pd.ExcelWriter(SALIDA_EXCEL) as writer:
    pd.DataFrame(cabeceras).to_excel(writer, sheet_name="Cabecera", index=False)
    pd.DataFrame(lugares).to_excel(writer, sheet_name="Lugares", index=False)
    pd.DataFrame(armas).to_excel(writer, sheet_name="Armas", index=False)
    pd.DataFrame(drogas).to_excel(writer, sheet_name="Drogas", index=False)
    pd.DataFrame(elementos).to_excel(writer, sheet_name="Elementos", index=False)
    pd.DataFrame(imputados).to_excel(writer, sheet_name="Imputados", index=False)
    pd.DataFrame(victimas).to_excel(writer, sheet_name="Victimas", index=False)
    pd.DataFrame(vehiculos).to_excel(writer, sheet_name="Vehiculos", index=False)
    pd.DataFrame(otros).to_excel(writer, sheet_name="Otros", index=False)

print(f"Procesamiento completo. Archivo guardado en {SALIDA_EXCEL}")
