import os
import re
import pandas as pd
from PyPDF2 import PdfReader
import warnings

warnings.filterwarnings("ignore", category=FutureWarning)  # Ocultar warning futuro

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
    bloques = re.findall(
        rf"({tipo} .*?)(?=IMPUTADO|VICTIMA|DROGA|ELEMENTO|VEHICULO|ARMA|$)",
        texto, re.IGNORECASE | re.DOTALL
    )
    resultados = []
    for bloque in bloques:
        if not re.search(r"(Tipo:|Nombres:|Incautacion:|Marca:)", bloque, re.IGNORECASE):
            continue
        idx = texto.find(bloque)
        lugar_match = re.findall(r"LUGAR\s+(\d+)", texto[:idx], re.IGNORECASE)
        lugar_nro = lugar_match[-1] if lugar_match else "1"
        resultados.append((bloque, lugar_nro))
    return resultados

def rellenar_vacios(diccionario):
    """Reemplaza valores vacíos por '-' en un diccionario (excepto números)."""
    return {k: (v if (v not in ["", None]) else "-") for k, v in diccionario.items()}

def asegurar_columnas(df, columnas, df_lugares):
    """
    Asegura que un DataFrame tenga columnas predefinidas.
    Si está vacío, crea filas usando Archivo y Lugar Nro de df_lugares,
    y rellena los demás campos con "-".
    """
    if df.empty:
        filas = []
        for _, row in df_lugares.iterrows():
            fila = {col: "-" for col in columnas}
            fila["Archivo"] = row["Archivo"]
            fila["Lugar Nro"] = str(row["Lugar Nro"])
            filas.append(fila)
        return pd.DataFrame(filas, columns=columnas)
    return df

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
            fecha = f"{y}-{m}-{d}"
            hora = partes[-1].strip()

    delito2 = a_mayusculas(extraer_unico(r"Delito 2:\s*(.+?)\s*<", texto)) or "-"
    delito3 = a_mayusculas(extraer_unico(r"Delito 3:\s*(.+?)\s*<", texto)) or "-"
    detalle_delito = a_mayusculas(extraer_unico(r"Detalle de Delito:\s*(.+?)\s*<", texto)) or "-"

    cabeceras.append({
        "Archivo": archivo,
        "Parte Operativo": a_mayusculas(extraer_unico(r"Parte Operativo:\s*([\w\-]+)", texto)),
        "Código Dependencia": a_mayusculas(extraer_unico(r"Codigo de Dependencia:\s*(\d+)", texto)),
        "Dependencia": a_mayusculas(extraer_unico(r"Dependencia:\s*(.+?)\s*<", texto)),
        "Fecha": fecha,
        "Hora": hora,
        "Sumario": a_mayusculas(extraer_unico(r"Sumario:\s*(.+?)\s*<", texto)),
        "Delito": a_mayusculas(extraer_unico(r"Delito 1:\s*(.+?)\s*<", texto)),
        "Delito 2": delito2,
        "Delito 3": delito3,
        "Detalle de Delito": detalle_delito,
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
            "Calle": a_mayusculas(calles[i]) if i < len(calles) else "-",
            "Localidad": a_mayusculas(localidades[i]) if i < len(localidades) else "-",
            "Departamento / Comuna": a_mayusculas(departamentos[i]) if i < len(departamentos) else "-",
            "Provincia": a_mayusculas(provincias[i]) if i < len(provincias) else "-",
            "Coordenadas": coords[i] if i < len(coords) else "-",
        })

    # --- ARMAS ---
    for bloque, lugar in extraer_bloques_con_lugar("ARMA", texto):
        armas.append(rellenar_vacios({
            "Archivo": archivo,
            "Lugar Nro": lugar,
            "Tipo": a_mayusculas(extraer_unico(r"Tipo:\s*([^\n<]+)", bloque)),
            "Detalles": a_mayusculas(extraer_unico(r"Detalles:\s*([^\n<]+)", bloque)),
            "Marca": a_mayusculas(extraer_unico(r"Marca:\s*([^\n<]+)", bloque)),
            "Modelo": a_mayusculas(extraer_unico(r"Modelo:\s*([^\n<]+)", bloque)),
            "Calibre": a_mayusculas(extraer_unico(r"Calibre:\s*([^\n<]+)", bloque)),
            "Numeración": a_mayusculas(extraer_unico(r"Numeracion:\s*([^\n<]+)", bloque)),
            "Pedido de Secuestro": a_mayusculas(extraer_unico(r"Pedido de Secuestro:\s*([^\n<]+)", bloque)),
            "Observaciones": a_mayusculas(extraer_unico(r"Observaciones:\s*(.+?)\s*(?=<|$)", bloque)),
            "Cantidad de Armamento": 1
        }))

    # --- DROGAS ---
    for bloque, lugar in extraer_bloques_con_lugar("DROGA", texto):
        drogas.append(rellenar_vacios({
            "Archivo": archivo,
            "Lugar Nro": lugar,
            "Tipo": a_mayusculas(extraer_unico(r"Tipo:\s*([^\n<]+)", bloque)),
            "Cantidad": extraer_unico(r"Cantidad:\s*([\d.,]+)", bloque),
            "Medición": a_mayusculas(extraer_unico(r"Medicion:\s*([^\n<]+)", bloque)),
            "Observaciones": a_mayusculas(extraer_unico(r"Observaciones:\s*(.+?)\s*(?=<|$)", bloque))
        }))

    # --- ELEMENTOS ---
    for bloque, lugar in extraer_bloques_con_lugar("ELEMENTO", texto):
        elementos.append(rellenar_vacios({
            "Archivo": archivo,
            "Lugar Nro": lugar,
            "Incautación": a_mayusculas(extraer_unico(r"Incautacion:\s*([^\n<]+)", bloque)),
            "Tipo": a_mayusculas(extraer_unico(r"Tipo:\s*([^\n<]+)", bloque)),
            "Subtipo": a_mayusculas(extraer_unico(r"Subtipo:\s*([^\n<]+)", bloque)),
            "Cantidad": extraer_unico(r"Cantidad:\s*([\d.,]+)", bloque),
            "Medición": a_mayusculas(extraer_unico(r"Medicion:\s*([^\n<]+)", bloque)),
            "Aforo": extraer_unico(r"Aforo:\$([\d.,]*)", bloque),
            "Observaciones": a_mayusculas(extraer_unico(r"Observaciones:\s*(.+?)\s*(?=<|$)", bloque))
        }))

    # --- IMPUTADOS ---
    for bloque, lugar in extraer_bloques_con_lugar("IMPUTADO", texto):
        imputados.append(rellenar_vacios({
            "Archivo": archivo,
            "Lugar Nro": lugar,
            "Nombres": a_mayusculas(extraer_unico(r"Nombres:\s*([^\n<]+)", bloque)),
            "Apellidos": a_mayusculas(extraer_unico(r"Apellidos:\s*([^\n<]+)", bloque)),
            "Edad": extraer_unico(r"Edad:\s*(\d+)", bloque),
            "Género": a_mayusculas(extraer_unico(r"Genero:\s*([^\n<]+)", bloque)),
            "DNI": extraer_unico(r"DNI:\s*([.\d]+)", bloque, limpiar=limpiar_dni),
            "Nacionalidad": a_mayusculas(extraer_unico(r"Nacionalidad:\s*([^\n<]+)", bloque)),
            "Domicilio": a_mayusculas(extraer_unico(r"Domicilio:\s*([^\n<]+)", bloque)),
            "Situación Procesal": a_mayusculas(extraer_unico(r"Situacion\s+Procesal\s*:\s*([^\n<]+)", bloque)),
            "Posee Captura": a_mayusculas(extraer_unico(r"Posee\s+Captura:\s*([^\n<]+)", bloque)),
            "Motivo Captura": a_mayusculas(extraer_unico(r"Motivo del Pedido de Captura:\s*([^\n<]+)", bloque)),
            "Alias": a_mayusculas(extraer_unico(r"Alias:\s*([^\n<]+)", bloque)) or "-",
            "Banda Criminal": a_mayusculas(extraer_unico(r"Banda Criminal:\s*([^\n<]+)", bloque)) or "-"
        }))

    # --- VÍCTIMAS ---
    for bloque, lugar in extraer_bloques_con_lugar("VICTIMA", texto):
        victimas.append(rellenar_vacios({
            "Archivo": archivo,
            "Lugar Nro": lugar,
            "Nombres": a_mayusculas(extraer_unico(r"Nombres:\s*([^\n<]+)", bloque)),
            "Apellidos": a_mayusculas(extraer_unico(r"Apellidos:\s*([^\n<]+)", bloque)),
            "Edad": extraer_unico(r"Edad:\s*(\d+)", bloque),
            "Género": a_mayusculas(extraer_unico(r"Genero:\s*([^\n<]+)", bloque)),
            "DNI": extraer_unico(r"DNI:\s*([.\d]+)", bloque, limpiar=limpiar_dni),
            "Nacionalidad": a_mayusculas(extraer_unico(r"Nacionalidad:\s*([^\n<]+)", bloque)),
            "Domicilio": a_mayusculas(extraer_unico(r"Domicilio:\s*([^\n<]+)", bloque)),
            "Cantidad de Victimas": 1
        }))

    # --- VEHÍCULOS ---
    for bloque, lugar in extraer_bloques_con_lugar("VEHICULO", texto):
        vehiculos.append(rellenar_vacios({
            "Archivo": archivo,
            "Lugar Nro": lugar,
            "Marca": a_mayusculas(extraer_unico(r"Marca:\s*([^\n<]+)", bloque)),
            "Modelo": a_mayusculas(extraer_unico(r"Modelo:\s*([^\n<]+)", bloque)),
            "Dominio": a_mayusculas(extraer_unico(r"Dominio:\s*([^\n<]+)", bloque)),
            "Tipo": a_mayusculas(extraer_unico(r"Tipo:\s*([^\n<]+)", bloque)),
            "Detalles": a_mayusculas(extraer_unico(r"Detalles:\s*(.+?)\s*(?=<|$)", bloque))
        }))

    # --- OTROS ---
    otros.append({
        "Archivo": archivo,
        "Efectivos": extraer_unico(r"Efectivos:\s*(\d+)", texto) or "-",
        "Moviles": extraer_unico(r"Moviles:\s*(\d+)", texto) or "-",
        "Motos": extraer_unico(r"Motos:\s*(\d+)", texto) or "-",
        "Canes": extraer_unico(r"Canes:\s*(\d+)", texto) or "-",
        "Morphrapid": extraer_unico(r"Morphrapid:\s*(\d+)", texto) or "-",
        "Scanners": extraer_unico(r"Scanners:\s*(\d+)", texto) or "-",
        "Caballos": extraer_unico(r"Caballos:\s*(\d+)", texto) or "-"
    })

# --- CREAR DATAFRAMES Y FORZAR FILAS VACÍAS ---
cols_arm = ["Archivo","Lugar Nro","Tipo","Detalles","Marca","Modelo","Calibre",
            "Numeración","Pedido de Secuestro","Observaciones","Cantidad de Armamento"]
cols_dro = ["Archivo","Lugar Nro","Tipo","Cantidad","Medición","Observaciones"]
cols_ele = ["Archivo","Lugar Nro","Incautación","Tipo","Subtipo","Cantidad",
            "Medición","Aforo","Observaciones"]
cols_imp = ["Archivo","Lugar Nro","Nombres","Apellidos","Edad","Género","DNI",
            "Nacionalidad","Domicilio","Situación Procesal","Posee Captura",
            "Motivo Captura","Alias","Banda Criminal"]
cols_vic = ["Archivo","Lugar Nro","Nombres","Apellidos","Edad","Género","DNI",
            "Nacionalidad","Domicilio","Cantidad de Victimas"]
cols_veh = ["Archivo","Lugar Nro","Marca","Modelo","Dominio","Tipo","Detalles"]

df_cab = pd.DataFrame(cabeceras)
df_lug = pd.DataFrame(lugares)
df_arm = asegurar_columnas(pd.DataFrame(armas), cols_arm, df_lug)
df_dro = asegurar_columnas(pd.DataFrame(drogas), cols_dro, df_lug)
df_ele = asegurar_columnas(pd.DataFrame(elementos), cols_ele, df_lug)
df_imp = asegurar_columnas(pd.DataFrame(imputados), cols_imp, df_lug)
df_vic = asegurar_columnas(pd.DataFrame(victimas), cols_vic, df_lug)
df_veh = asegurar_columnas(pd.DataFrame(vehiculos), cols_veh, df_lug)
df_otr = pd.DataFrame(otros)

# --- GUARDAR Y UNIFICAR ---
with pd.ExcelWriter(SALIDA_EXCEL) as writer:
    df_cab.to_excel(writer, sheet_name="Cabecera", index=False)
    df_lug.to_excel(writer, sheet_name="Lugares", index=False)
    df_arm.to_excel(writer, sheet_name="Armas", index=False)
    df_dro.to_excel(writer, sheet_name="Drogas", index=False)
    df_ele.to_excel(writer, sheet_name="Elementos", index=False)
    df_imp.to_excel(writer, sheet_name="Imputados", index=False)
    df_vic.to_excel(writer, sheet_name="Victimas", index=False)
    df_veh.to_excel(writer, sheet_name="Vehiculos", index=False)
    df_otr.to_excel(writer, sheet_name="Otros", index=False)

    # Convertir Lugar Nro a str
    for df in [df_lug, df_arm, df_dro, df_ele, df_imp, df_vic, df_veh]:
        if "Lugar Nro" in df.columns:
            df["Lugar Nro"] = df["Lugar Nro"].astype(str)

    # Base inicial
    unificado = df_lug.merge(df_cab, on="Archivo", how="left").merge(df_otr, on="Archivo", how="left")

    def unir(df_origen, df_apartado, nombre):
        if df_apartado.empty:
            return df_origen
        if "Lugar Nro" in df_apartado.columns:
            df_apartado["Lugar Nro"] = df_apartado["Lugar Nro"].astype(str)
        return df_origen.merge(df_apartado, on=["Archivo", "Lugar Nro"], how="left", suffixes=("", f"_{nombre}"))

    # Unir todo
    unificado = unir(unificado, df_arm, "Arma")
    unificado = unir(unificado, df_dro, "Droga")
    unificado = unir(unificado, df_ele, "Elemento")
    unificado = unir(unificado, df_imp, "Imputado")
    unificado = unir(unificado, df_vic, "Victima")
    unificado = unir(unificado, df_veh, "Vehiculo")

    # Campo Procedimiento
    unificado["Procedimiento"] = "-"
    combinaciones_vistas = set()
    for idx, row in unificado.iterrows():
        clave = (row["Archivo"], row["Lugar Nro"])
        if clave not in combinaciones_vistas:
            unificado.at[idx, "Procedimiento"] = 1
            combinaciones_vistas.add(clave)

    unificado = unificado.fillna("-").infer_objects(copy=False)
    unificado.to_excel(writer, sheet_name="Unificado", index=False)

print(f"Procesamiento completo. Archivo guardado en {SALIDA_EXCEL}")
