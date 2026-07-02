import pandas as pd
#import datetime
from datetime import datetime
import numpy as np
import os
from pathlib import Path
import getpass

### Detecta ruta del script y redirige las demas direcciones path ###
rutainicial = Path.home()
usuario = getpass.getuser()
antes, sep, despues = str(rutainicial).partition(usuario)
base = Path(antes + sep)
Linea_Negocio=Linea_Negocio
print("CalculoStock0 - OK")
#exec(open(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Conexiones_a_SAP"/"KardexGeneralProv.py", encoding="utf-8-sig").read())
dfkardexorigen = pd.read_csv(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"dfkardexorigen.txt", sep="|", encoding="utf-8",on_bad_lines="skip")

dfkardexHC=dfkardexorigen[dfkardexorigen["Nombre de grupo"]==Linea_Negocio]

art=dfkardexHC['Número de artículo'].unique()
#dfkardexHC = dfkardexorigen[dfkardexorigen["ItemCode"] == 'A20090000184']
#dfkardexHC = dfkardexorigen[dfkardexorigen["Número de artículo"] == 'A20090000184']


dfkardexHC["Cantidad de entrada"] = (
    dfkardexHC["Cantidad de entrada"]
    .astype(str)  # Convertir todos los valores a string
    .str.replace(',', '', regex=False)  # Remover separadores de miles (coma)
    .replace(r'^\s*$', None, regex=True)  # Reemplazar valores vacíos o espacios con None (para que sean NaN)
    .astype(float)  # Convertir finalmente a float
)


dfkardexHC["Cantidad de salida"] = (
    dfkardexHC["Cantidad de salida"]
    .astype(str)  # Convertir todos los valores a string
    .str.replace(',', '', regex=False)  # Remover separadores de miles (coma)
    .replace(r'^\s*$', None, regex=True)  # Reemplazar valores vacíos o espacios con None (para que sean NaN)
    .astype(float)  # Convertir finalmente a float
)

#Entrada neta por artículo y fecha
dfkardexHC["Entrada neta"] = dfkardexHC["Cantidad de entrada"] - dfkardexHC["Cantidad de salida"]
dfkardexHC.rename(columns={"CreateDate": "Fecha"}, inplace=True)
#dfkardexHC['Fecha'] = pd.to_datetime(dfkardexHC['Fecha'], dayfirst=True, errors='coerce')
dfkardexHC['Fecha'] = pd.to_datetime(dfkardexHC['Fecha'], errors='coerce')
# Excluir el ítem A18130007558 en la fecha 04/03/2019
dfkardexHC = dfkardexHC[~((dfkardexHC["Número de artículo"] == "A18130007558") & (dfkardexHC["Fecha"] == pd.Timestamp("2019-03-04")))]
entrada_neta = dfkardexHC.groupby(["Número de artículo", "Fecha"])["Entrada neta"].sum()
entrada_neta = entrada_neta.reset_index()


# Inventario diario desde kardex
entrada_neta.sort_values(by=["Número de artículo", "Fecha"], ascending=[True, True],inplace=True)
entrada_neta['Inventario final'] = entrada_neta.groupby("Número de artículo")["Entrada neta"].cumsum()
# Reemplazar valores muy pequeños por cero
entrada_neta["Inventario final"] = np.where(
    np.abs(entrada_neta["Inventario final"]) < 1e-10,
    0,
    entrada_neta["Inventario final"]
)

"""

GENERAR TUPLAS ITEMCODE, FECHA MIN Y FECHA MAX
"""


# Obtener fecha mínima y máxima por artículo
rangos = entrada_neta.groupby("Número de artículo")["Fecha"].agg(["min", "max"]).reset_index()

# Generar rangos de fechas como listas
rangos["Fechas"] = rangos.apply(lambda x: pd.date_range(x["min"], x["max"], freq="D"), axis=1)

# Explode para deshacer listas y tener un registro por día
df_all = rangos[["Número de artículo", "Fechas"]].explode("Fechas").rename(columns={"Fechas": "Fecha"})


"""
GENERAR LA ENTRADA NETA FINAL
"""

entrada_neta_full = pd.merge(
    df_all,
    entrada_neta[["Número de artículo", "Fecha", "Inventario final"]],
    on=["Número de artículo", "Fecha"],
    how="left"
)

entrada_neta_full["Inventario final"] = entrada_neta_full.groupby("Número de artículo")["Inventario final"].ffill()

# entrada_neta_full["Inventario final"] = entrada_neta_full.groupby("Número de artículo")["Inventario final"].ffill().fillna(0)
# entrada_neta_full["Entrada neta"] = entrada_neta_full["Entrada neta"].fillna(0)
prueba = entrada_neta_full[entrada_neta_full["Número de artículo"] == 'A18110002183']
"""

EXPORTACION DE DATOS

"""

#output = "C://Users//AnthonyPradoCornejo//OneDrive - MARCO PERUANA SA//Escritorio//Rotacion Incubadoras-Electro-Frio//HIDRAULICA COMPONENTE//pruebakardexdiario.xlsx"

#with pd.ExcelWriter(output) as writer:
#    prueba.to_excel(writer, sheet_name="Reporte", index=False)

#Definimos la funcion
def calcular_stock_cero_por_grupo(grupo, fecha_limite_min, fecha_limite_max):
    fechas = grupo['Fecha']
    fecha_min = fechas.min()
    fecha_max = fechas.max()

    item = grupo['Número de artículo'].iloc[0]

    # Rango completamente dentro del período
    if fecha_limite_min <= fecha_max <= fecha_limite_max:
        if fecha_limite_min <= fecha_min <= fecha_limite_max:
            inicio = fecha_min
        else:
            inicio = fecha_limite_min

        calendario = pd.DataFrame({
            'Fecha': pd.date_range(start=inicio, end=fecha_limite_max, freq='D')
        })
        calendario['Número de artículo'] = item

        merged = pd.merge(
            calendario,
            grupo[['Fecha', 'Inventario final']],
            on='Fecha',
            how='left'
        ).sort_values('Fecha')

        merged['Inventario final'] = merged['Inventario final'].fillna(0)
        total_dias = len(merged)
        dias_stock_cero = (merged['Inventario final'] == 0).sum()

        return pd.Series({
            'ItemCode': item,
            'Stock Cero (%)': pd.to_datetime(0) if pd.isna(dias_stock_cero / total_dias) else dias_stock_cero / total_dias,
            'Fecha_min_rang': pd.to_datetime('1900-01-01') if pd.isna(inicio) else inicio,
            'Fecha_max_rang': pd.to_datetime('1900-01-01') if pd.isna(fecha_limite_max) else fecha_limite_max,
            'Fecha_min_art': pd.to_datetime('1900-01-01') if pd.isna(fecha_min) else fecha_min,
            'Fecha_max_art': pd.to_datetime('1900-01-01') if pd.isna(fecha_max) else fecha_max,
        })

    # Inicia dentro del rango, su fecha máxima está fuera
    elif fecha_limite_min <= fecha_min <= fecha_limite_max:
        calendario = pd.DataFrame({
            'Fecha': pd.date_range(start=fecha_min, end=fecha_limite_max, freq='D')
        })
        calendario['Número de artículo'] = item

        merged = pd.merge(
            calendario,
            grupo[['Fecha', 'Inventario final']],
            on='Fecha',
            how='left'
        ).sort_values('Fecha')

        merged['Inventario final'] = merged['Inventario final'].fillna(0)
        total_dias = len(merged)
        dias_stock_cero = (merged['Inventario final'] == 0).sum()

        return pd.Series({
            'ItemCode': item,
            'Stock Cero (%)': pd.to_datetime(0) if pd.isna(dias_stock_cero / total_dias) else dias_stock_cero / total_dias,
            'Fecha_min_rang': pd.to_datetime('1900-01-01') if pd.isna(fecha_min) else fecha_min,
            'Fecha_max_rang': pd.to_datetime('1900-01-01') if pd.isna(fecha_limite_max) else fecha_limite_max,
            'Fecha_min_art': pd.to_datetime('1900-01-01') if pd.isna(fecha_min) else fecha_min,
            'Fecha_max_art': pd.to_datetime('1900-01-01') if pd.isna(fecha_max) else fecha_max,
        })

    # Rango del articulo contiene el rango
    elif fecha_min <= fecha_limite_max and fecha_max >= fecha_limite_min:
        calendario = pd.DataFrame({
            'Fecha': pd.date_range(start=fecha_limite_min, end=fecha_limite_max, freq='D')
        })
        calendario['Número de artículo'] = item

        merged = pd.merge(
            calendario,
            grupo[['Fecha', 'Inventario final']],
            on='Fecha',
            how='left'
        ).sort_values('Fecha')

        merged['Inventario final'] = merged['Inventario final'].fillna(0)
        total_dias = len(merged)
        dias_stock_cero = (merged['Inventario final'] == 0).sum()

        return pd.Series({
            'ItemCode': item,
            'Stock Cero (%)': pd.to_datetime(0) if pd.isna(dias_stock_cero / total_dias) else dias_stock_cero / total_dias,
            'Fecha_min_rang': fecha_limite_min,
            'Fecha_max_rang': fecha_limite_max,
            'Fecha_min_art': pd.to_datetime('1900-01-01') if pd.isna(fecha_min) else fecha_min,
            'Fecha_max_art': pd.to_datetime('1900-01-01') if pd.isna(fecha_max) else fecha_max,
        })

    # Caso final: No hay cruce con el período
    else:
        return pd.Series({
            'ItemCode': item,
            'Stock Cero (%)': 1,
            'Fecha_min_rang': pd.to_datetime('1900-01-01'),
            'Fecha_max_rang': pd.to_datetime('1900-01-01'),
            'Fecha_min_art': pd.to_datetime('1900-01-01') if pd.isna(fecha_min) else fecha_min,
            'Fecha_max_art': pd.to_datetime('1900-01-01') if pd.isna(fecha_max) else fecha_max,
        })


# Fecha actual

fecha_limite_min = pd.Timestamp("2025-01-01")
fecha_limite_max = pd.Timestamp("2025-12-31")

resultado = entrada_neta_full.groupby('Número de artículo',group_keys=False).apply(
    calcular_stock_cero_por_grupo,
    fecha_limite_min=fecha_limite_min,
    fecha_limite_max=fecha_limite_max
)



resultado_Stock_Cero = resultado.reset_index(drop=True)

lista_sap_unicon = df_Unicon_SAP["SAP"].unique()
detalleUnicon = entrada_neta_full[
    (entrada_neta_full["Fecha"] >= fecha_limite_min) &
    (entrada_neta_full["Fecha"] <= fecha_limite_max)
].copy()
detalleUnicon = detalleUnicon[
    detalleUnicon["Número de artículo"].isin(lista_sap_unicon)
]


#detalleUnicon = detalleUnicon[detalleUnicon["Inventario final"] == 0]

# Ordenar
detalleUnicon = detalleUnicon.sort_values(
    by=["Número de artículo", "Fecha"]
)

detalleUnicon = detalleUnicon.reset_index(drop=True)

"""Ajuste"""
fecha_limite_min2 = pd.to_datetime("2022-12-01")
fecha_limite_max2 = pd.to_datetime("2024-11-01")
resultado2 = entrada_neta_full.groupby('Número de artículo', group_keys=False).apply(
    calcular_stock_cero_por_grupo,
    fecha_limite_min=fecha_limite_min2,
    fecha_limite_max=fecha_limite_max2
)
resultado2.rename(columns={"Stock Cero (%)": "Stock Cero (%)1"}, inplace=True)

resultado_Stock_Cero = resultado.reset_index(drop=True)
resultado_Stock_Cero2 = resultado2.reset_index(drop=True)

