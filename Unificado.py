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

def renombrar_apartado(df, prefijo):
    if df.empty:
        return df
    nuevas = {}
    for col in df.columns:
        if col not in ["Archivo", "Lugar Nro"]:
            nuevas[col] = f"{col} {prefijo}"
    return df.rename(columns=nuevas)

# --- TABLAS ACUMULADAS ---
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

    # --- CABECERA ---
    #fecha = extraer_unico(r"Fecha y Hora:\s*([\d]{2}-[\d]{2}-[\d]{4})", texto)
    #hora = extraer_unico(r"Fecha y Hora:.*?-\s*(\d{2}:\d{2})", texto)
    #fecha = extraer_unico(r"Fecha\s*y\s*Hora\s*:\s*([0-9]{2}-[0-9]{2}-[0-9]{4})", texto_norm)
    #hora = extraer_unico(r"Fecha\s*y\s*Hora\s*:[^0-9]*([0-9]{2}:[0-9]{2})", texto_norm)
    fecha_hora = extraer_unico(r"Fecha y Hora:\s*([\d\-]+\s*-\s*\d{2}:\d{2})", texto_norm)
    fecha, hora = "", ""
    if fecha_hora:
        partes = fecha_hora.split("-")
        if len(partes) >= 3:
            d, m, y = partes[0].strip(), partes[1].strip(), partes[2].strip().split()[0]
            fecha = f"{y}-{m}-{d}"  # YYYY-MM-DD
            hora = partes[-1].strip()
    delito2 = a_mayusculas(extraer_unico(r"Delito 2:\s*(.+?)\s*<", texto)) or "-"
    delito3 = a_mayusculas(extraer_unico(r"Delito 3:\s*(.+?)\s*<", texto)) or "-"
    detalle_delito = a_mayusculas(extraer_unico(r"Detalle de Delito:\s*(.+?)\s*<", texto)) or "-"

    cabeceras.append({
        "Archivo": archivo,
        "Parte Operativo": a_mayusculas(extraer_unico(r"(.*)\.pdf", archivo)),
        "Código Dependencia": a_mayusculas(extraer_unico(r"Codigo de Dependencia:\s*(\d+)", texto)),
        "Dependencia": a_mayusculas(extraer_unico(r"(\d*)-P.*", archivo)),
        "Fecha": fecha if fecha else "-",
        "Hora": hora if hora else "-",
        "Sumario": a_mayusculas(extraer_unico(r"Sumario:\s*(.+?)\s*<", texto)),
        "Delito": a_mayusculas(extraer_unico(r"Delito 1:\s*(.+?)\s*<", texto)),
        "Delito 2": delito2,
        "Delito 3": delito3,
        "Detalle de Delito": detalle_delito,
        "Modalidad": a_mayusculas(extraer_unico(r"Modalidad 1:\s*(.+?)\s*<", texto)),
        "Tipo Intervención": a_mayusculas(extraer_unico(r"Tipo de Intervenci[oó]n:\s*([^\n<]+)", texto_norm)),
        "Juzgado / Fiscalía": a_mayusculas(extraer_unico(r"Juzgado\s*/?\s*Fiscal[ií]a\s*:?\s*([\s\S]+?)(?=<|\n|$)", texto)),
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
            "Calle": a_mayusculas(calles[i]),
            "Localidad": a_mayusculas(localidades[i]) if i < len(localidades) else "-",
            "Departamento / Comuna": a_mayusculas(departamentos[i]) if i < len(departamentos) else "-",
            "Provincia": a_mayusculas(provincias[i]) if i < len(provincias) else "-",
            "Coordenadas": coords[i] if i < len(coords) else "-",
        })

    # --- OTROS APARTADOS ---
    for bloque, lugar in extraer_bloques_con_lugar("ARMA", texto):
        armas.append(rellenar_vacios({
            "Archivo": archivo, "Lugar Nro": lugar,
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
    for bloque, lugar in extraer_bloques_con_lugar("DROGA", texto):
        drogas.append(rellenar_vacios({
            "Archivo": archivo, "Lugar Nro": lugar,
            "Tipo": a_mayusculas(extraer_unico(r"Tipo:\s*([^\n<]+)", bloque)),
            "Cantidad": extraer_unico(r"Cantidad:\s*([\d.,]+)", bloque),
            "Medición": a_mayusculas(extraer_unico(r"Medicion:\s*([^\n<]+)", bloque)),
            "Observaciones": a_mayusculas(extraer_unico(r"Observaciones:\s*(.+?)\s*(?=<|$)", bloque))
        }))
    for bloque, lugar in extraer_bloques_con_lugar("ELEMENTO", texto):
        elementos.append(rellenar_vacios({
            "Archivo": archivo, "Lugar Nro": lugar,
            "Incautación": a_mayusculas(extraer_unico(r"Incautacion:\s*([^\n<]+)", bloque)),
            "Tipo": a_mayusculas(extraer_unico(r"Tipo:\s*([^\n<]+)", bloque)),
            "Subtipo": a_mayusculas(extraer_unico(r"Subtipo:\s*([^\n<]+)", bloque)),
            "Cantidad": extraer_unico(r"Cantidad:\s*([\d.,]+)", bloque),
            "Medición": a_mayusculas(extraer_unico(r"Medicion:\s*([^\n<]+)", bloque)),
            "Aforo": extraer_unico(r"Aforo:\$([\d.,]*)", bloque),
            "Observaciones": a_mayusculas(extraer_unico(r"Observaciones:\s*(.+?)\s*(?=<|$)", bloque))
        }))
    for bloque, lugar in extraer_bloques_con_lugar("IMPUTADO", texto):
        imputados.append(rellenar_vacios({
            "Archivo": archivo, "Lugar Nro": lugar,
            "Nombres": a_mayusculas(extraer_unico(r"Nombres:\s*([^\n<]+)", bloque)),
            "Apellidos": a_mayusculas(extraer_unico(r"Apellidos:\s*([^\n<]+)", bloque)),
            "Edad": extraer_unico(r"Edad:\s*(\d+)", bloque),
            "Género": a_mayusculas(extraer_unico(r"Genero:\s*([^\n<]+)", bloque)),
            "DNI": extraer_unico(r"DNI:\s*([.\d]+)", bloque, limpiar=limpiar_dni),
            "Nacionalidad": a_mayusculas(extraer_unico(r"Nacionalidad:\s*([^\n<]+)", bloque)),
            "Domicilio": a_mayusculas(extraer_unico(r"Domicilio:\s*([^\n<]+)", bloque)),
            "Situación Procesal": a_mayusculas(extraer_unico(r"Situacion\s*Procesal\s*:\s*([\w\s]+)", bloque)),
            "Posee Captura": a_mayusculas(extraer_unico(r"Posee\s*Captura\s*:\s*([\w\s]+)", bloque)),
            "Motivo Captura": a_mayusculas(extraer_unico(r"Motivo del Pedido de Captura:\s*([^\n<]+)", bloque)),
            "Alias": a_mayusculas(extraer_unico(r"Alias:\s*([^\n<]+)", bloque)) or "-",
            "Banda Criminal": a_mayusculas(extraer_unico(r"Banda Criminal:\s*([^\n<]+)", bloque)) or "-"
        }))
    for bloque, lugar in extraer_bloques_con_lugar("VICTIMA", texto):
        victimas.append(rellenar_vacios({
            "Archivo": archivo, "Lugar Nro": lugar,
            "Nombres": a_mayusculas(extraer_unico(r"Nombres:\s*([^\n<]+)", bloque)),
            "Apellidos": a_mayusculas(extraer_unico(r"Apellidos:\s*([^\n<]+)", bloque)),
            "Edad": extraer_unico(r"Edad:\s*(\d+)", bloque),
            "Género": a_mayusculas(extraer_unico(r"Genero:\s*([^\n<]+)", bloque)),
            "DNI": extraer_unico(r"DNI:\s*([.\d]+)", bloque, limpiar=limpiar_dni),
            "Nacionalidad": a_mayusculas(extraer_unico(r"Nacionalidad:\s*([^\n<]+)", bloque)),
            "Domicilio": a_mayusculas(extraer_unico(r"Domicilio:\s*([^\n<]+)", bloque)),
            "Cantidad de Victimas": 1
        }))
    for bloque, lugar in extraer_bloques_con_lugar("VEHICULO", texto):
        vehiculos.append(rellenar_vacios({
            "Archivo": archivo, "Lugar Nro": lugar,
            "Marca": a_mayusculas(extraer_unico(r"Marca:\s*([^\n<]+)", bloque)),
            "Modelo": a_mayusculas(extraer_unico(r"Modelo:\s*([^\n<]+)", bloque)),
            "Dominio": a_mayusculas(extraer_unico(r"Dominio:\s*([^\n<]+)", bloque)),
            "Tipo": a_mayusculas(extraer_unico(r"Tipo:\s*([^\n<]+)", bloque)),
            "Detalles": a_mayusculas(extraer_unico(r"Detalles:\s*(.+?)\s*(?=<|$)", bloque))
        }))

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

# --- CREAR DATAFRAMES Y RENOMBRAR ---
df_cab = pd.DataFrame(cabeceras)
df_lug = pd.DataFrame(lugares)
df_arm = renombrar_apartado(asegurar_columnas(pd.DataFrame(armas), ["Archivo","Lugar Nro","Tipo","Detalles","Marca","Modelo","Calibre",
                                                "Numeración","Pedido de Secuestro","Observaciones","Cantidad de Armamento"], df_lug), "Arma")
df_dro = renombrar_apartado(asegurar_columnas(pd.DataFrame(drogas), ["Archivo","Lugar Nro","Tipo","Cantidad","Medición","Observaciones"], df_lug), "Droga")
df_ele = renombrar_apartado(asegurar_columnas(pd.DataFrame(elementos), ["Archivo","Lugar Nro","Incautación","Tipo","Subtipo","Cantidad",
                                                    "Medición","Aforo","Observaciones"], df_lug), "Elemento")
df_imp = renombrar_apartado(asegurar_columnas(pd.DataFrame(imputados), ["Archivo","Lugar Nro","Nombres","Apellidos","Edad","Género","DNI",
                                                    "Nacionalidad","Domicilio","Situación Procesal","Posee Captura",
                                                    "Motivo Captura","Alias","Banda Criminal"], df_lug), "Imputado")
df_vic = renombrar_apartado(asegurar_columnas(pd.DataFrame(victimas), ["Archivo","Lugar Nro","Nombres","Apellidos","Edad","Género","DNI",
                                                    "Nacionalidad","Domicilio","Cantidad de Victimas"], df_lug), "Victima")
df_veh = renombrar_apartado(asegurar_columnas(pd.DataFrame(vehiculos), ["Archivo","Lugar Nro","Marca","Modelo","Dominio","Tipo","Detalles"], df_lug), "Vehiculo")
df_otr = pd.DataFrame(otros)

# --- FORZAR 'Lugar Nro' COMO STRING ---
for df in [df_lug, df_arm, df_dro, df_ele, df_imp, df_vic, df_veh]:
    if "Lugar Nro" in df.columns:
        df["Lugar Nro"] = df["Lugar Nro"].astype(str)

# --- EXPANDIR Y UNIR ---
unificado_apartados = expandir_y_combinar(df_arm, df_dro, df_ele, df_imp, df_vic, df_veh)
if "Lugar Nro" in unificado_apartados.columns:
    unificado_apartados["Lugar Nro"] = unificado_apartados["Lugar Nro"].astype(str)

unificado = (
    df_lug.merge(df_cab, on="Archivo", how="left")
          .merge(df_otr, on="Archivo", how="left")
          .merge(unificado_apartados, on=["Archivo","Lugar Nro"], how="left")
)

# --- CAMPO PROCEDIMIENTO ---
unificado["Procedimiento"] = "-"
vistos = set()
for idx, row in unificado.iterrows():
    clave = (row["Archivo"], row["Lugar Nro"])
    if clave not in vistos:
        unificado.at[idx, "Procedimiento"] = 1
        vistos.add(clave)

# --- GUARDAR ---
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
