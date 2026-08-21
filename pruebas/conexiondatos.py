# -*- coding: utf-8 -*-
"""
Conexion a datos para abastecimiento
"""
#IMPORTACION
#Recorrer archivos

from pathlib import Path
from datetime import datetime
import pandas as pd
import re
import requests
import os
import logging

from datetime import datetime
from pathlib import Path
import getpass


_log = logging.getLogger("Codigo.ConexionDatos")


if "base" not in dir():
    base = Path(__file__).resolve().parent

_log.info("[ConexionDatos] Inicio")

#Linea_Negocio="MANGUERA Y CONEXION"
#Linea_Negocio="LUBRICACION MINERIA"
Linea_Negocio=Linea_Negocio
#Linea_Negocio="EQUIPOS TRANS. MATER"

#Variable Hoy
hoy = datetime.today()
hoy = pd.to_datetime(hoy).normalize()
#hoy = datetime(2026, 1, 1)  # 15 de diciembre de 2025

# ARCHIVOS LOCALES #
#pesca = pd.read_excel( base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Forecast"/"Datas"/"Codigos_pesca.xlsx")
#codigos_pesca = pesca["Codigo_SAP PESCA"].tolist()
_log.info("[ConexionDatos] Preparando archivos de clasificación")

# La carpeta está dentro de Pruebas Linux
base_path = base

# Carpeta específica de la línea
carpeta = base_path / "equipos_trans_mater"
archivo_valido = None

_log.info("[ConexionDatos] Buscando archivo de clasificación en: %s", carpeta)

for archivo in carpeta.glob('clasificacion_art*.xlsx'):
    # Extraer el rango MMYY-MMYY
    match = re.search(r'(\d{4})-(\d{4})', archivo.stem)
    if match:
        inicio, fin = match.groups()

        fecha_inicio = datetime.strptime(inicio, '%m%y')
        fecha_fin = datetime.strptime(fin, '%m%y')

        # Ajuste: considerar todo el mes final
        fecha_fin = fecha_fin.replace(day=28) + pd.offsets.MonthEnd(0)

        if fecha_inicio <= hoy <= fecha_fin:
            archivo_valido = archivo
            break

if archivo_valido:
    _log.info(
        "[ConexionDatos] Archivo de clasificacion: %s",
        archivo_valido.name
    )
    df_clasificacion = pd.read_excel(archivo_valido)
    
else:
    _log.info(
        "[ConexionDatos] No se encontró archivo de clasificación valido  para la fecha: %s",
        hoy.date()
    )
    raise FileNotFoundError('No se encontró un archivo válido para la fecha actual')



if Linea_Negocio in ["LUBRICACION MINERIA", "MANGUERA Y CONEXION","HIDRAULI. COMPONENTE","EQUIPOS TRANS. MATER", "FILTRACION","TRANSM. DE POTENCIA","HERRAMIENTA HIDRAULI","LUBRICACION INDUSTRI","SIST DE LUBRICACION"]:

    ## Forecast ##
    _log.info("[ConexionDatos] Buscando archivo de forecast")
    carpeta = base / "equipos_trans_mater"
    
    fecha_cierre = (
        #pd.to_datetime(hoy).replace(day=1) - pd.DateOffset(months=1)
        pd.to_datetime(hoy).replace(day=1)
    )
    # Formato esperado en el nombre del archivo: 2026-03-01
    fecha_cierre_str = fecha_cierre.strftime("%Y-%m-%d")
    # Carpeta donde buscar
    
    # Buscar primero el forecast del mes actual
    patron = f"forecast_revisado_{fecha_cierre_str}*.xlsx"
    archivos = list(carpeta.glob(patron))
    
    # Si no existe, buscar el del mes anterior
    if len(archivos) == 0:
        _log.info(
            "[ConexionDatos] No se encontró forecast de %s. Buscando mes anterior",
            fecha_cierre_str
        )
        fecha_cierre = fecha_cierre - pd.DateOffset(months=1)
        fecha_cierre_str = fecha_cierre.strftime("%Y-%m-%d")
    
        patron = f"forecast_revisado_{fecha_cierre_str}*.xlsx"
        archivos = list(carpeta.glob(patron))
    
    # Si tampoco existe, lanzar error
    if len(archivos) == 0:
        _log.info(
            "[ConexionDatos] No se encontró forecast para %s ni para el mes anterior",
            fecha_cierre_str
        )
        raise FileNotFoundError(
            f"No existe archivo de forecast revisado para {fecha_cierre_str} ni para el mes anterior."
        )
    
    archivo_valido = archivos[0]
    
    _log.info(
        "[ConexionDatos] Forecast seleccionado: %s",
        archivo_valido.name
    )

    

    df_forecast = pd.read_excel(
        archivo_valido,
        engine="openpyxl"  # recomendado para .xlsx
    )
    
    
    col_objetivo = 'Total Proy 2026'
    # Posición de la columna
    idx = df_forecast.columns.get_loc(col_objetivo)  
    # Columnas a conservar
    columnas_finales = (
        ['SAP_Origen', 'SAP'] +
        list(df_forecast.columns[idx + 1:])
    )    
    # Nos quedamos solo con ellas
    df_forecast = df_forecast[columnas_finales]
    # Elimina columnas sin nombre o llamadas "Unnamed"
    df_forecast = df_forecast.loc[:, ~(
        df_forecast.columns.isna() |
        df_forecast.columns.astype(str).str.contains(r'^Unnamed', case=False, na=False)
    )]
    df_forecast = df_forecast.drop(columns=["Comentarios"], errors="ignore")
    
    #Lineas con la columna grupo
    df_forecast=revertir_reporte_anterior_MGN(df_forecast)    

    df_forecast["Fecha"] = pd.to_datetime(df_forecast["Fecha"], errors="coerce")
    df_forecast = df_forecast[df_forecast["Fecha"] >= fecha_cierre]
    df_forecast = df_forecast[
    df_forecast["Forecast"].notna() &
    (df_forecast["Forecast"].astype(str).str.strip() != "")
    ]
    df_forecast = df_forecast.groupby(['SAP_Origen', 'Fecha'], as_index=False)['Forecast'].sum()
    # Suma total de forecast por SAP
    forecast_total_sap = (
        df_forecast
        .groupby("SAP_Origen")["Forecast"]
        .transform("sum")
    )
    
    # Eliminar SAP cuyo forecast total sea 0
    df_forecast = df_forecast[forecast_total_sap != 0]
    
       
##Resultados Estimacion de fechas ##
_log.info("[ConexionDatos] Cargando archivos de resultados")
df_provisional = pd.read_csv(
    base / "df_provisional.txt",
    sep=",",
    encoding="utf-8"
)
dfResumenLeadTimes = pd.read_csv(
    base / "dfresumenleadtimes.txt",
    sep=",",
    encoding="utf-8"
)
dfResumenLeadTimes = dfResumenLeadTimes[
    dfResumenLeadTimes["Sociedad"] == "MP"
]

FreservSAP_filtrado = pd.read_csv(
    base / "freservsap_filtrado.txt",
    sep=",",
    encoding="utf-8"
)

Reporte_Precios_Local_Imp = pd.read_csv(
    base / "df_reportedeprecios.txt",
    sep=",",
    encoding="utf-8"
)

SeguiBOSAP_f = pd.read_csv(
    base / "seguibosap_f.txt",
    sep=",",
    encoding="utf-8"
)

_log.info("[ConexionDatos] Archivos de resultados cargados")


# ARCHIVOS ONLINE COMPARTIDOS #
archivos = [
    ("https://sistemasmarco-my.sharepoint.com/:x:/g/personal/khernandez_marco_com_pe/IQCcfvvSf98BR4itSa35SsmnAYoltav3VKK7ai_Jm0_pDe8?download=1", "listadeprecios_osma.xlsx"),
    ("https://sistemasmarco-my.sharepoint.com/:x:/g/personal/khernandez_marco_com_pe/IQBHePy_AElxTJtHYQHbXVpHAfPYvV9-XsA0E59kruGTN1w?download=1", "listadeprecios_cisge.xlsx"),
    ("https://sistemasmarco-my.sharepoint.com/:x:/g/personal/aprado_marco_com_pe/IQD6n1NCji2uQqsU2MUeTrNCAfcwq9qGJ6Aj5KZp3sh_yUU?download=1", "listafinaladapconexmangueras.xlsx"),
    ("https://sistemasmarco-my.sharepoint.com/:x:/g/personal/aprado_marco_com_pe/IQCZeGyK3E6FQY5qEMLXU5XGAXqFFtyi6qqWYaajSU4mTvM?download=1", "plmpsawinnertkhoses.xlsx"),
    ("https://sistemasmarco-my.sharepoint.com/:x:/g/personal/aprado_marco_com_pe/IQCZvIE5eVyCTI8gpvxS43tGAQFhRs5Ff5BWd3biGvl6m00?download=1", "contratocsgunicon.xlsx"),
#    ("https://sistemasmarco-my.sharepoint.com/:x:/g/personal/aprado_marco_com_pe/IQC5GP9tE8JlQYJ0ZVlDkP8QASEmyy4RcBvZmjD9Gzpm-w8?download=1", "Inventario_MARCO PERUANA S.A._MARCO PERUANA S.A. - JICAMARCA.xlsx"),
#    ("https://sistemasmarco-my.sharepoint.com/:x:/g/personal/aprado_marco_com_pe/IQDVanYHZ548T5WYULoX9QtIASV1NaNbW-_Uc5r-lovsEns?download=1", "Inventario_MARCO PERUANA S.A._MARCO PERUANA S.A. - HIERBABUENA.xlsx"),
    ("https://sistemasmarco-my.sharepoint.com/:x:/g/personal/aprado_marco_com_pe/IQAna7PQPzBQSoCON4avHSzJAVsSieMjgh5z4Pj9rx630_A?download=1", "grupossku.xlsx"),
    ("https://sistemasmarco-my.sharepoint.com/:x:/g/personal/aprado_marco_com_pe/IQCD-5v8Upu3SqzWQOsNdHVoAf6CpHvMSaPv_DbE57kocgU?download=1", "pf_com.xlsx"),   
    ("https://sistemasmarco-my.sharepoint.com/:x:/g/personal/aprado_marco_com_pe/IQBTTLsBe3FLRaprQAy9y8lUAeUGlr-JA5JRUyeZjt0xOrc?download=1", "pf_com_hidra.xlsx"),
    ("https://sistemasmarco.sharepoint.com/:x:/s/PlanificaciondeInventarios/IQDiizcBRElRT4GxyivqC9gSAQNVo1rsujFRy_TGoo-N7qI?download=1", "pf_com_li.xlsx"),
    ("https://sistemasmarco.sharepoint.com/:x:/s/PlanificaciondeInventarios/IQAf_9eE4KI9QZUdBGHqXqsdAcgZz3gxyRC8q2kg0DEpj-I?download=1", "pf_com_sl.xlsx")
   ]

# Ruta donde se guardarán los archivos
carpeta_destino = base / "descargas"
ruta_excel = carpeta_destino

# Crear la carpeta si no existe
os.makedirs(carpeta_destino, exist_ok=True)
_log.info("[ConexionDatos] Descargando archivos compartidos")


for url, nombre_archivo in archivos:
    ruta_archivo = carpeta_destino / nombre_archivo

    try:
        respuesta = requests.get(url, stream=True)
        respuesta.raise_for_status()

        with open(ruta_archivo, "wb") as archivo:
            for chunk in respuesta.iter_content(chunk_size=8192):
                archivo.write(chunk)

        _log.info("[ConexionDatos] Descargado: %s", nombre_archivo)

    except requests.exceptions.RequestException as e:
        _log.error(
            "[ConexionDatos] Error descargando %s: %s",
            nombre_archivo,
            e
        )

#Relacion Componente - PF
# Leer archivos según la línea de negocio
if Linea_Negocio == "LUBRICACION MINERIA":

    dfMRP = pd.read_excel(
        ruta_excel / "PF-Com.xlsx",
        sheet_name="Hoja1"
    )

    df_KPI = pd.read_excel(
        base / "KPI.xlsx",
        sheet_name="Hoja1"
    )


elif Linea_Negocio == "HIDRAULI. COMPONENTE":

    dfMRP = pd.read_excel(
        ruta_excel / "pf_com_hidra.xlsx",
        sheet_name="Hoja1"
    )


elif Linea_Negocio == "LUBRICACION INDUSTRI":

    dfMRP = pd.read_excel(
        ruta_excel / "pf_com_li.xlsx",
        sheet_name="Hoja1"
    )


elif Linea_Negocio == "SIST DE LUBRICACION":

    dfMRP = pd.read_excel(
        ruta_excel / "pf_com_sl.xlsx",
        sheet_name="Hoja1"
    )

   
## Relacion de proyectos vigentes ##
df_proyectos = pd.read_excel(
    base / "proyectos.xlsx",
    sheet_name="Hoja1"
)

df_proyectos = df_proyectos[
    df_proyectos["Linea negocio"] == Linea_Negocio
]

df_remplazos = df_proyectos[
    df_proyectos["Grupo Remplazo"].notna()
]

if Linea_Negocio == "MANGUERA Y CONEXION":

    df_IGS = pd.read_excel(
        ruta_excel / "PL MPSA Winner TK Hoses Apr 8th 2025.xlsx",
        sheet_name="Sheet1",
        header=4
    )

    df_Osma = pd.read_excel(
        ruta_excel / "Lista de precios - Ferretería Osma.xlsx",
        sheet_name="FERRETERÍA OSMA",
        header=1
    )

    df_Osma = set(
        df_Osma["Codigo SAP MG"].dropna().unique()
    )

    df_Cisge = pd.read_excel(
        ruta_excel / "Lista de precios - Cisge.xlsx",
        sheet_name="COMERCIAL CISGE SAC 2026",
        header=1
    )

    df_Cisge = set(
        df_Cisge["Código SAP"].dropna().unique()
    )

    df_China = pd.read_excel(
        ruta_excel / "Lista Final Adap Conex Mangueras.xlsx",
        sheet_name="Hoja1"
    )

    df_China_sap = (
        df_China.loc[
            df_China["GROUP"].isna(),
            "SAP"
        ]
        .dropna()
        .unique()
    )

    df_China_Group = set(
        df_China["GROUP"]
        .dropna()
        .unique()
    )

_log.info("[ConexionDatos] Relaciones Componente - PF cargadas")


_log.info("[ConexionDatos] Cargando información UNICON")

df_Unicon_SAP_Cod = pd.read_excel(
    ruta_excel / "contratocsgunicon.xlsx",
    sheet_name="Hoja1"
)

df_Unicon_SAP_Cod=df_Unicon_SAP_Cod[["Código Articulo","SAP"]]
df_Unicon_SAP_Cod["Código Articulo"] = (df_Unicon_SAP_Cod["Código Articulo"].astype("Int64").astype("string"))
##
df_Unicon_SAP=df_Unicon_SAP_Cod.copy()
df_Unicon_SAP = df_Unicon_SAP[ ~df_Unicon_SAP["SAP"].isin([ "A20100000116", "A20100000115", "A20100000170", "A20110000205", "A21020000051", "A21020000055",  "A20030000018"])]
##
df_ASAI = pd.read_excel(
    ruta_excel / "analisi_asai_toromocho_antamina.xlsx",
    sheet_name="Hoja1"
)


df_ASAI = df_ASAI[["SAP", "Código Articulo"]]
df_ASAI = df_ASAI.drop_duplicates()
df_ASAI["Código Articulo"] = (
    pd.to_numeric(df_ASAI["Código Articulo"], errors="coerce")
    .astype("Int64")
    .astype("string")
)

df_Unicon_SAP_Cod = pd.concat([df_Unicon_SAP_Cod, df_ASAI], ignore_index=True)

## Data UNICOM ######

#df_Jicamarca = pd.read_excel(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"inventario_marco_peruana_sa_marco_peruana_jicamarca.xls", sheet_name="inventarioArticulo",header=3)
df_Jicamarca = pd.read_excel(
    base / "inventario_marco_peruana_sa_marco_peruana_jicamarca.xls",
    sheet_name="inventarioArticulo",
    header=3
)

df_Jicamarca=df_Jicamarca.iloc[:, 1:]
df_Jicamarca=df_Jicamarca[["Codigo Unicon","Stock"]]
df_Jicamarca = df_Jicamarca.dropna(subset=["Codigo Unicon"])
df_Jicamarca=df_Jicamarca.dropna(subset=["Codigo Unicon"])
df_Jicamarca["Codigo Unicon"] = (df_Jicamarca["Codigo Unicon"].astype("Int64").astype("string"))

df_HierbaBuena = pd.read_excel(base / "inventario_marco_peruana_sa_marco_peruana_hierbabuena.xls",sheet_name="inventarioArticulo",header=3)
df_HierbaBuena=df_HierbaBuena.iloc[:, 1:]
df_HierbaBuena=df_HierbaBuena[["Codigo Unicon","Stock"]]
df_HierbaBuena = df_HierbaBuena.dropna(subset=["Codigo Unicon"])
df_HierbaBuena["Codigo Unicon"] = (df_HierbaBuena["Codigo Unicon"].astype("Int64").astype("string"))

df_Antamina = pd.read_excel( base / "inventario_marco_peruana_sa_marco_peruana_cant_antamina.xls",sheet_name="inventarioArticulo", header=3)
df_Antamina=df_Antamina.iloc[:, 1:]
df_Antamina=df_Antamina[["Codigo Unicon","Stock"]]
df_Antamina = df_Antamina.dropna(subset=["Codigo Unicon"])
df_Antamina["Codigo Unicon"] = (df_Antamina["Codigo Unicon"].astype("Int64").astype("string"))

df_Toromocho = pd.read_excel(base / "inventario_marco_peruana_sa_marco_peruana_cant_toromocho.xls", sheet_name="inventarioArticulo", header=3)
df_Toromocho=df_Toromocho.iloc[:, 1:]
df_Toromocho=df_Toromocho[["Codigo Unicon","Stock"]]
df_Toromocho = df_Toromocho.dropna(subset=["Codigo Unicon"])
df_Toromocho["Codigo Unicon"] = (df_Toromocho["Codigo Unicon"].astype("Int64").astype("string"))


df_grupos_MC = pd.read_excel( ruta_excel / "grupossku.xlsx", sheet_name="Resumen")
df_ensambleETM = pd.read_excel( ruta_excel / "cant_ensamble_por_repuesto.xlsx",sheet_name="CANT ENSAM POR REPUESTO")
df_ensambleETM.columns = df_ensambleETM.columns.str.strip()
df_ensambleETM = df_ensambleETM.dropna(subset=["SAP MARCO"])
df_ensambleETM = df_ensambleETM[df_ensambleETM["SAP MARCO"].str.upper() != "NO EXISTE"]
_log.info("[ConexionDatos] Información UNICON cargada")