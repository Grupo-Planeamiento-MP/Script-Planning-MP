"""
Limpieza de LT
"""
#IMPORTACION
#Recorrer archivos
from pathlib import Path
from datetime import datetime
import pandas as pd
import re
import numpy as np

#Descarga automatica de archivos
import requests
import os
from datetime import datetime
from pathlib import Path
import getpass

### Detecta ruta del script y redirige las demas direcciones path ###
rutainicial = Path.home()
usuario = getpass.getuser()
antes, sep, despues = str(rutainicial).partition(usuario)
base = Path(antes + sep)


#Conexion a datos
dfResumenLeadTimes = pd.read_csv(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Archivos_Compartidos"/"Resultados"/"Estimacion de Fechas"/"dfResumenLeadTimes.txt",sep=',',encoding='utf-8')
dfResumenLeadTimes=dfResumenLeadTimes[dfResumenLeadTimes["Sociedad"]=="MP"]
dfResumenLeadTimes = dfResumenLeadTimes.rename(columns={'Diferencia en meses': 'LT'})
#dfResumenLeadTimes = dfResumenLeadTimes[dfResumenLeadTimes["TrnspName"] != "TERRE"]

df_LT_referencial = pd.read_excel(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Datas"/"Data General"/'LT_Referencial.xlsx', sheet_name="Hoja1")
df_LT_referencial['ItemCode'] = df_LT_referencial['ItemCode'].astype(str)
df_LT_referencial['LT_Referencial'] = pd.to_numeric(df_LT_referencial['LT_Referencial'])

#Ajuste
items_faltantes = df_articulos.loc[
    ~df_articulos["ItemCode"].isin(dfResumenLeadTimes["ItemCode"]),
    ["ItemCode"]
].drop_duplicates()

# 3. Concatenar
dfResumenLeadTimes = pd.concat([dfResumenLeadTimes, items_faltantes], ignore_index=True)

# =====================================
# 1. PREPARACIÓN
# =====================================
#Cruce
dfResumenLeadTimesMP = dfResumenLeadTimes.merge(
    df_LT_referencial[['ItemCode', 'LT_Referencial']],  # Solo seleccionamos las columnas necesarias
    left_on='ItemCode',
    right_on='ItemCode',
    how='left'
)

dfResumenLeadTimesMP = dfResumenLeadTimesMP.sort_values(["ItemCode", "Max_FechIN"], ascending=[True, False])
dfResumenLeadTimesMP["tiene_lt_ref"] = dfResumenLeadTimesMP["LT_Referencial"].notna()

# =====================================
# 2. FUNCIONES AUXILIARES
# =====================================

def quitar_outliers(grupo):
    q1 = grupo["LT"].quantile(0.25)
    q3 = grupo["LT"].quantile(0.75)
    iqr = q3 - q1
    return grupo[(grupo["LT"] >= q1 - 1.5 * iqr) &
                 (grupo["LT"] <= q3 + 1.5 * iqr)]

def marcar_outliers_serie(x):
    q1 = x.quantile(0.25)
    q3 = x.quantile(0.75)
    iqr = q3 - q1
    
    lower = q1 - 1.5 * iqr
    upper = q3 + 1.5 * iqr
    
    return ((x < lower) | (x > upper)).map({True: "SI", False: "NO"})

def ultimas_compras(df, group_col):
    df["rank"] = df.groupby(group_col).cumcount() + 1
    return df[df["rank"] <= 4]


def tipo_compra_mayor(df, group_col):
    return df.groupby(group_col)["TrnspName"]\
        .transform(lambda x: x.value_counts().idxmax())


def simulacion_lt(items_insuf):
    # Número de simulaciones
    n_sim = 4

    # Crear lista para resultados
    resultados = []

    for _, row in items_insuf.iterrows():
        item = row["ItemCode"]
        lt_ref = row["LT_Referencial"]
        
        # Parámetros de la normal
        media = lt_ref
        std = lt_ref * 0.15
        
        # Simular 4 valores
        lt_simulados = np.random.normal(loc=media, scale=std, size=n_sim)
        
        # Evitar negativos (importante)
        lt_simulados = np.clip(lt_simulados, a_min=1, a_max=None)
        
        # Crear filas (una por simulación)
        for i, lt in enumerate(lt_simulados, start=1):
            resultados.append({
                "Sociedad": "MP",
                "Tipo de Compra": "Importación",
                "PO": None,
                "ItemCode": item,
                "Line PO": None,
                "TrnspName": None,
                "Codigo de cliente/proveedor": None,
                "Tipo Producto": None,
                "Max_FechPO": None,
                "Max_FechIN": None,
                "Avg_Q_PO": None,
                "Sum_Q_EM": None,
                "LT": round(lt, 2)
            })

    # Convertir a DataFrame final
    df_simulado = pd.DataFrame(resultados)
    
    return df_simulado


# =====================================
# 3. PROCESO CON LT REFERENCIAL
# =====================================
def procesar_con_lt(df):
    df=df_con.copy()
    df = df.copy()

    # 1. FILTRO CERCANÍA
    df["es_cercano"] = abs(df["LT"] - df["LT_Referencial"]) <= 0.5
    df_cercano = df[df["es_cercano"]].copy()

    # 2. ÚLTIMAS 4 COMPRAS
    df_cercano = ultimas_compras(df_cercano, "ItemCode")

    # 3. VALIDAR HISTORIAL >= 2
    conteo = df_cercano.groupby("ItemCode")["LT"].transform("count")
    df_cercano["suficiente"] = conteo >= 1

    # 4. RESULTADO CON HISTÓRICO
    df_ok = df_cercano[df_cercano["suficiente"]]
    
    columnas = [
    "Sociedad",
    "Tipo de Compra",
    "PO",
    "ItemCode",
    "Line PO",
    "TrnspName",
    "Codigo de cliente/proveedor",
    "Tipo Producto",
    "Max_FechPO",
    "Max_FechIN",
    "Avg_Q_PO",
    "Sum_Q_EM",
    "LT"
    ]    
    LT_Historico_Referencial = df_ok[columnas]    
    LT_Historico_Referencial["Tipo LT"]="LT_Historico_Referencial"
    # 5. ITEMS SIN HISTÓRICO → SIMULACIÓN (usar LT referencial)
    items_insuf = df[~df["ItemCode"].isin(df_ok["ItemCode"])][["ItemCode", "LT_Referencial"]].drop_duplicates()
    #simulacion
    LT_Referencial = simulacion_lt(items_insuf)
    LT_Referencial["Tipo LT"]="LT_Referencial"
    # 6. UNIÓN
    resultado = pd.concat([LT_Historico_Referencial, LT_Referencial], ignore_index=True)   
    
    return resultado


# =====================================
# 4. PROCESO SIN LT REFERENCIAL
# =====================================
def procesar_sin_lt(df):
    df=df_sin.copy()
    df_item = df.copy()

    # =========================
    # 1. OUTLIERS POR ITEM
    # =========================
    df_item["Outlier_Item_Mod"] = df_item.groupby(["ItemCode", "TrnspName"])["LT"].transform(marcar_outliers_serie)
    df_item=df_item[df_item["Outlier_Item_Mod"]=="NO"]
    # =========================
    # 2. VALIDAR HISTORIAL ITEM
    # =========================
    conteo_item = df_item.groupby("ItemCode")["LT"].transform("count")
    df_item["suficiente_item"] = conteo_item >= 2

    # =========================
    # 3. ITEMS CON DATA SUFICIENTE
    # =========================
    df_ok = df_item[df_item["suficiente_item"]].copy()
    df_ok["tipo_compra_mayor"] = tipo_compra_mayor(df_ok, "ItemCode")  
    df_ok = df_ok[df_ok["TrnspName"] == df_ok["tipo_compra_mayor"]]
    df_ok = ultimas_compras(df_ok, "ItemCode")
    columnas = [
    "Sociedad",
    "Tipo de Compra",
    "PO",
    "ItemCode",
    "Line PO",
    "TrnspName",
    "Codigo de cliente/proveedor",
    "Tipo Producto",
    "Max_FechPO",
    "Max_FechIN",
    "Avg_Q_PO",
    "Sum_Q_EM",
    "LT"
    ]    
    LT_Historico = df_ok[columnas]   
    LT_Historico["Tipo LT"]="LT_Historico"
    # =========================
    # 4. ITEMS SIN DATA → FAMILIA
    # =========================
    items_insuf = df_item[~df_item["suficiente_item"]]["ItemCode"].unique()

    df_fam = df[df["ItemCode"].isin(items_insuf)].copy()

    # OUTLIERS POR FAMILIA
    df_fam["Outlier_Fam_Mod"] = df_fam.groupby(["Tipo Producto", "TrnspName"])["LT"].transform(marcar_outliers_serie)
    df_fam=df_fam[df_fam["Outlier_Fam_Mod"]=="NO"]
    
    # TIPO COMPRA MAYOR POR FAMILIA
    df_fam["tipo_compra_mayor"] = tipo_compra_mayor(df_fam, "Tipo Producto")
    df_fam = df_fam[df_fam["TrnspName"] == df_fam["tipo_compra_mayor"]]

    # ÚLTIMAS 4 POR ITEM
    df_fam = ultimas_compras(df_fam, "ItemCode")
    columnas = [
    "Sociedad",
    "Tipo de Compra",
    "PO",
    "ItemCode",
    "Line PO",
    "TrnspName",
    "Codigo de cliente/proveedor",
    "Tipo Producto",
    "Max_FechPO",
    "Max_FechIN",
    "Avg_Q_PO",
    "Sum_Q_EM",
    "LT"
    ]    
    LT_Historico_Familia = df_fam[columnas] 
    LT_Historico_Familia["Tipo LT"]="LT_Historico_Familia"

    # =========================
    # 5. ITEMS SIN NADA → NO CALCULA
    # =========================
    items_procesados = pd.concat([LT_Historico["ItemCode"], LT_Historico_Familia["ItemCode"]])

    items_no_calc = df[~df["ItemCode"].isin(items_procesados)]["ItemCode"].drop_duplicates()

    LT_NoCalcula = pd.DataFrame({
        "Sociedad": "MP",
        "Tipo de Compra": "Importación",
        "PO": np.nan,
        "ItemCode":items_no_calc,
        "Line PO": np.nan,
        "TrnspName": np.nan,
        "Codigo de cliente/proveedor": np.nan,
        "Tipo Producto": np.nan,
        "Max_FechPO": np.nan,
        "Max_FechIN": np.nan,
        "Avg_Q_PO": np.nan,
        "Sum_Q_EM": np.nan,
        #"LT": np.nan
        "LT": 1
    })
    LT_NoCalcula["Tipo LT"]="LT_NoCalcula"

    # =========================
    # 6. UNIÓN FINAL
    # =========================
    resultado = pd.concat([LT_Historico, LT_Historico_Familia, LT_NoCalcula], ignore_index=True)

    return resultado


# =====================================
# 5. APLICAR FLUJO GENERAL
# =====================================
df_con = dfResumenLeadTimesMP[dfResumenLeadTimesMP["tiene_lt_ref"]]
df_sin = dfResumenLeadTimesMP[~dfResumenLeadTimesMP["tiene_lt_ref"]]

res_con = procesar_con_lt(df_con)
res_sin = procesar_sin_lt(df_sin)

# =====================================
# 6. RESULTADO FINAL
# =====================================
df_LeadTime = pd.concat([res_con, res_sin], ignore_index=True)
df_LeadTime = df_LeadTime.rename(columns={"LT": "LT_final"})