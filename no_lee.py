import os
import re
import pandas as pd
from PyPDF2 import PdfReader
import warnings

warnings.filterwarnings("ignore", category=FutureWarning)

# --- CONFIGURACIÓN ---
DIRECTORIO_PDFS = r"C:\Users\ecastro\Desktop\PARTES"
SALIDA_EXCEL = r"C:\Users\ecastro\Desktop\resultado_detallado_corregido.xlsx"

# --- FUNCIONES AUXILIARES ---
def limpiar_dni(dni):
    return re.sub(r"\D", "", dni)

def a_mayusculas(valor):
    return valor.strip().upper() if isinstance(valor, str) else valor

def normalizar_parte_operativo(valor):
    if not valor:
        return "-"
    match = re.search(r"\d{3,4}-PO-\d+-\d{4}", valor)
    return match.group(0) if match else valor.replace("\n", "").replace(" ", "")

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
    return [limpiar(m.strip()) if limpiar else m.strip() for m in matches]

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
    return {k: (v if (v not in ["", None]) else "-") for k, v in diccionario.items()}

def asegurar_columnas(df, columnas, df_lugares):
    if df.empty:
        filas = []
        for _, row in df_lugares.iterrows():
            fila = {col: "-" for col in columnas}
            fila["Archivo"] = row["Archivo"]
            fila["Lugar Nro"] = str(row["Lugar Nro"])
            filas.append(fila)
        return pd.DataFrame(filas, columns=columnas)
    return df

def expandir_y_combinar(*dfs, claves=("Archivo", "Lugar Nro")):
    dfs = [df.reset_index(drop=True) for df in dfs]
    claves_comunes = pd.concat([df[list(claves)].drop_duplicates() for df in dfs]).drop_duplicates()

    resultado = []
    for _, clave in claves_comunes.iterrows():
        subconjuntos = []
        max_len = 0

        for df in dfs:
            sub = df[(df[claves[0]] == clave[claves[0]]) & (df[claves[1]] == clave[claves[1]])].copy()
            if sub.empty:
                fila_vacia = {col: "-" for col in df.columns}
                fila_vacia[claves[0]] = clave[claves[0]]
                fila_vacia[claves[1]] = clave[claves[1]]
                sub = pd.DataFrame([fila_vacia], columns=df.columns)
            subconjuntos.append(sub.reset_index(drop=True))
            max_len = max(max_len, len(sub))

        subconjuntos_ext = []
        for sub in subconjuntos:
            if len(sub) < max_len:
                fila_vacia = {col: "-" for col in sub.columns}
                fila_vacia[claves[0]] = clave[claves[0]]
                fila_vacia[claves[1]] = clave[claves[1]]
                repetidas = pd.DataFrame([fila_vacia] * (max_len - len(sub)), columns=sub.columns)
                sub = pd.concat([sub, repetidas], ignore_index=True)
            subconjuntos_ext.append(sub)

        combinado = pd.concat(subconjuntos_ext, axis=1)
        combinado = combinado.loc[:, ~combinado.columns.duplicated()]
        resultado.append(combinado)

    return pd.concat(resultado, ignore_index=True)

# --- TABLAS ---
cabeceras, lugares, armas, drogas, elementos, imputados, victimas, vehiculos, otros = ([] for _ in range(9))

# --- PROCESAR PDFs ---
for archivo in os.listdir(DIRECTORIO_PDFS):
    if not archivo.lower().endswith(".pdf"):
        continue

    ruta_pdf = os.path.join(DIRECTORIO_PDFS, archivo)
    print(f"Procesando: {archivo}")

    reader = PdfReader(ruta_pdf)
    texto = "".join(page.extract_text() for page in reader.pages)

    # Normalizar texto para soportar PDFs sin "<"
    texto_norm = texto.replace("\n", " ")
    texto_norm = re.sub(r"\s+", " ", texto_norm)
    texto_norm = re.sub(r"\s*-\s*", "-", texto_norm)

    # --- Cabecera ---
    fecha = extraer_unico(r"Fecha\s*y\s*Hora\s*:\s*([0-9]{2}-[0-9]{2}-[0-9]{4})", texto_norm)
    hora = extraer_unico(r"Fecha\s*y\s*Hora\s*:[^0-9]*([0-9]{2}:[0-9]{2})", texto_norm)
    tipo_intervencion = a_mayusculas(extraer_unico(
        r"Tipo\s*de\s*Intervencion\s*:?\s*([A-ZÁÉÍÓÚÑ\s]+?)(?=<|\n|$)", texto_norm
    ))

    cabeceras.append({
        "Archivo": archivo,
        "Parte Operativo": normalizar_parte_operativo(extraer_unico(r"Parte\s*Operativo\s*:\s*([\d\s\-PO]+)", texto)),
        "Código Dependencia": a_mayusculas(extraer_unico(r"Codigo\s*de\s*Dependencia\s*:\s*(\d+)", texto)),
        "Dependencia": a_mayusculas(extraer_unico(r"Dependencia\s*:\s*(.+?)(?:<|\n|$)", texto)),
        "Fecha": fecha if fecha else "-",
        "Hora": hora if hora else "-",
        "Sumario": a_mayusculas(extraer_unico(r"Sumario\s*:\s*(.+?)(?:<|\n|$)", texto)),
        "Delito": a_mayusculas(extraer_unico(r"Delito\s*1\s*:\s*(.+?)(?:<|\n|$)", texto)),
        "Modalidad": a_mayusculas(extraer_unico(r"Modalidad\s*1\s*:\s*(.+?)(?:<|\n|$)", texto)),
        "Tipo Intervención": tipo_intervencion if tipo_intervencion else "-",
        "Juzgado / Fiscalía": a_mayusculas(extraer_unico(r"Juzgado\s*/?\s*Fiscal[ií]a\s*:?\s*([\s\S]+?)(?=<|\n|$)", texto)),
        "Secretaría": a_mayusculas(extraer_unico(r"Secretaria\s*:\s*(.+?)(?:<|\n|$)", texto)),
        "Causa Nro.": a_mayusculas(extraer_unico(r"Causa\s*(Nro\.?|Numero)\s*:?\s*(.+?)(?:<|\n|$)", texto)),
        "Carátula": a_mayusculas(extraer_unico(r"Caratula\s*:\s*(.+?)(?:<|\n|$)", texto)),
    })

    # --- LUGARES ---
    calles = extraer_todos(r"Calle\s*:\s*(.+?)(?:<|\n|$)", texto)
    localidades = extraer_todos(r"Localidad\s*:\s*(.+?)(?:<|\n|$)", texto)
    departamentos = extraer_todos(r"Departamento\s*/\s*Partido\s*/\s*Comuna\s*:\s*(.+?)(?:<|\n|$)", texto)
    provincias = extraer_todos(r"Provincia\s*:\s*(.+?)(?:<|\n|$)", texto)
    coords = extraer_todos(r"Coordenadas\s*:\s*([^\n<]+)", texto)
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
            "Archivo": archivo, "Lugar Nro": lugar,
            "Tipo": a_mayusculas(extraer_unico(r"Tipo\s*:\s*([^\n<]+)", bloque)),
            "Detalles": a_mayusculas(extraer_unico(r"Detalles\s*:\s*([^\n<]+)", bloque)),
            "Marca": a_mayusculas(extraer_unico(r"Marca\s*:\s*([^\n<]+)", bloque)),
            "Modelo": a_mayusculas(extraer_unico(r"Modelo\s*:\s*([^\n<]+)", bloque)),
            "Calibre": a_mayusculas(extraer_unico(r"Calibre\s*:\s*([^\n<]+)", bloque)),
            "Numeración": a_mayusculas(extraer_unico(r"Numeracion\s*:\s*([^\n<]+)", bloque)),
            "Pedido de Secuestro": a_mayusculas(extraer_unico(r"Pedido\s*de\s*Secuestro\s*:\s*([^\n<]+)", bloque)),
            "Observaciones": a_mayusculas(extraer_unico(r"Observaciones\s*:\s*(.+?)(?=<|\n|$)", bloque)),
            "Cantidad de Armamento": 1
        }))

    # --- DROGAS ---
    for bloque, lugar in extraer_bloques_con_lugar("DROGA", texto):
        drogas.append(rellenar_vacios({
            "Archivo": archivo, "Lugar Nro": lugar,
            "Tipo": a_mayusculas(extraer_unico(r"Tipo\s*:\s*([^\n<]+)", bloque)),
            "Cantidad": extraer_unico(r"Cantidad\s*:\s*([\d.,]+)", bloque),
            "Medición": a_mayusculas(extraer_unico(r"Medicion\s*:\s*([^\n<]+)", bloque)),
            "Observaciones": a_mayusculas(extraer_unico(r"Observaciones\s*:\s*(.+?)(?=<|\n|$)", bloque))
        }))

    # --- ELEMENTOS ---
    for bloque, lugar in extraer_bloques_con_lugar("ELEMENTO", texto):
        elementos.append(rellenar_vacios({
            "Archivo": archivo, "Lugar Nro": lugar,
            "Incautación": a_mayusculas(extraer_unico(r"Incautacion\s*:\s*([^\n<]+)", bloque)),
            "Tipo": a_mayusculas(extraer_unico(r"Tipo\s*:\s*([^\n<]+)", bloque)),
            "Subtipo": a_mayusculas(extraer_unico(r"Subtipo\s*:\s*([^\n<]+)", bloque)),
            "Cantidad": extraer_unico(r"Cantidad\s*:\s*([\d.,]+)", bloque),
            "Medición": a_mayusculas(extraer_unico(r"Medicion\s*:\s*([^\n<]+)", bloque)),
            "Aforo": extraer_unico(r"Aforo\s*:\$([\d.,]*)", bloque),
            "Observaciones": a_mayusculas(extraer_unico(r"Observaciones\s*:\s*(.+?)(?=<|\n|$)", bloque))
        }))

    # --- IMPUTADOS ---
    for bloque, lugar in extraer_bloques_con_lugar("IMPUTADO", texto):
        imputados.append(rellenar_vacios({
            "Archivo": archivo, "Lugar Nro": lugar,
            "Nombres": a_mayusculas(extraer_unico(r"Nombres\s*:\s*([^\n<]+)", bloque)),
            "Apellidos": a_mayusculas(extraer_unico(r"Apellidos\s*:\s*([^\n<]+)", bloque)),
            "Edad": extraer_unico(r"Edad\s*:\s*(\d+)", bloque),
            "Género": a_mayusculas(extraer_unico(r"Genero\s*:\s*([^\n<]+)", bloque)),
            "DNI": extraer_unico(r"DNI\s*:\s*([.\d]+)", bloque, limpiar=limpiar_dni),
            "Nacionalidad": a_mayusculas(extraer_unico(r"Nacionalidad\s*:\s*([^\n<]+)", bloque)),
            "Domicilio": a_mayusculas(extraer_unico(r"Domicilio\s*:\s*([^\n<]+)", bloque)),
            "Situación Procesal": a_mayusculas(extraer_unico(r"Situacion\s*Procesal\s*:\s*([^\n<]+)", bloque)),
            "Posee Captura": a_mayusculas(extraer_unico(r"Posee\s*Captura\s*:\s*([^\n<]+)", bloque)),
            "Motivo Captura": a_mayusculas(extraer_unico(r"Motivo\s*del\s*Pedido\s*de\s*Captura\s*:\s*([^\n<]+)", bloque)),
            "Alias": a_mayusculas(extraer_unico(r"Alias\s*:\s*([^\n<]+)", bloque)) or "-",
            "Banda Criminal": a_mayusculas(extraer_unico(r"Banda\s*Criminal\s*:\s*([^\n<]+)", bloque)) or "-"
        }))

    # --- VÍCTIMAS ---
    for bloque, lugar in extraer_bloques_con_lugar("VICTIMA", texto):
        victimas.append(rellenar_vacios({
            "Archivo": archivo, "Lugar Nro": lugar,
            "Nombres": a_mayusculas(extraer_unico(r"Nombres\s*:\s*([^\n<]+)", bloque)),
            "Apellidos": a_mayusculas(extraer_unico(r"Apellidos\s*:\s*([^\n<]+)", bloque)),
            "Edad": extraer_unico(r"Edad\s*:\s*(\d+)", bloque),
            "Género": a_mayusculas(extraer_unico(r"Genero\s*:\s*([^\n<]+)", bloque)),
            "DNI": extraer_unico(r"DNI\s*:\s*([.\d]+)", bloque, limpiar=limpiar_dni),
            "Nacionalidad": a_mayusculas(extraer_unico(r"Nacionalidad\s*:\s*([^\n<]+)", bloque)),
            "Domicilio": a_mayusculas(extraer_unico(r"Domicilio\s*:\s*([^\n<]+)", bloque)),
            "Cantidad de Victimas": 1
        }))

    # --- VEHÍCULOS ---
    for bloque, lugar in extraer_bloques_con_lugar("VEHICULO", texto):
        vehiculos.append(rellenar_vacios({
            "Archivo": archivo, "Lugar Nro": lugar,
            "Marca": a_mayusculas(extraer_unico(r"Marca\s*:\s*([^\n<]+)", bloque)),
            "Modelo": a_mayusculas(extraer_unico(r"Modelo\s*:\s*([^\n<]+)", bloque)),
            "Dominio": a_mayusculas(extraer_unico(r"Dominio\s*:\s*([^\n<]+)", bloque)),
            "Tipo": a_mayusculas(extraer_unico(r"Tipo\s*:\s*([^\n<]+)", bloque)),
            "Detalles": a_mayusculas(extraer_unico(r"Detalles\s*:\s*(.+?)(?=<|\n|$)", bloque))
        }))

    # --- OTROS ---
    otros.append({
        "Archivo": archivo,
        "Efectivos": extraer_unico(r"Efectivos\s*:\s*(\d+)", texto) or "-",
        "Moviles": extraer_unico(r"Moviles\s*:\s*(\d+)", texto) or "-",
        "Motos": extraer_unico(r"Motos\s*:\s*(\d+)", texto) or "-",
        "Canes": extraer_unico(r"Canes\s*:\s*(\d+)", texto) or "-",
        "Morphrapid": extraer_unico(r"Morphrapid\s*:\s*(\d+)", texto) or "-",
        "Scanners": extraer_unico(r"Scanners\s*:\s*(\d+)", texto) or "-",
        "Caballos": extraer_unico(r"Caballos\s*:\s*(\d+)", texto) or "-"
    })

# --- CREAR DATAFRAMES ---
cols_arm = ["Archivo","Lugar Nro","Tipo","Detalles","Marca","Modelo","Calibre",
            "Numeración","Pedido de Secuestro","Observaciones","Cantidad de Armamento"]
cols_dro = ["Archivo","Lugar Nro","Tipo","Cantidad","Medición","Observaciones"]
cols_ele = ["Archivo","Lugar Nro","Incautación","Tipo","Subtipo","Cantidad","Medición","Aforo","Observaciones"]
cols_imp = ["Archivo","Lugar Nro","Nombres","Apellidos","Edad","Género","DNI","Nacionalidad","Domicilio",
            "Situación Procesal","Posee Captura","Motivo Captura","Alias","Banda Criminal"]
cols_vic = ["Archivo","Lugar Nro","Nombres","Apellidos","Edad","Género","DNI","Nacionalidad","Domicilio","Cantidad de Victimas"]
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

# --- UNIFICAR ---
unificado_apartados = expandir_y_combinar(df_arm, df_dro, df_ele, df_imp, df_vic, df_veh)
unificado = df_lug.merge(df_cab, on="Archivo", how="left").merge(df_otr, on="Archivo", how="left")
unificado = expandir_y_combinar(unificado, unificado_apartados)

# Campo Procedimiento
unificado["Procedimiento"] = "-"
vistas = set()
for idx, row in unificado.iterrows():
    clave = (row["Archivo"], row["Lugar Nro"])
    if clave not in vistas:
        unificado.at[idx, "Procedimiento"] = 1
        vistas.add(clave)

unificado = unificado.fillna("-").infer_objects(copy=False)

# --- EXPORTAR ---
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
    unificado.to_excel(writer, sheet_name="Unificado", index=False)

print(f"Procesamiento completo. Archivo guardado en {SALIDA_EXCEL}")
