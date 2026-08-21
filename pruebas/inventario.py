import pandas as pd
import numpy as np
import os
import logging
from pathlib import Path
_log = logging.getLogger("Codigo.Inventario")


### Detecta ruta del script y redirige las demas direcciones path ###
if "base" not in dir():
    base = Path(__file__).resolve().parent

_log.info("[Inventario] Inicio")



_log.info("[Inventario] Leyendo tomadeinventariov3ap.txt")

#dfInvenSAPBO= pd.read_table("C:/Users/AnthonyPradoCornejo/OneDrive - MARCO PERUANA SA/Escritorio/Rotacion Incubadoras-Electro-Frio/ASTEC/Planeamiento de Abastecimiento/Inventario/Toma de Inventario V3 AP.txt", encoding='utf-16', sep='\t', quotechar='"', low_memory=False)
#exec(open(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Conexiones_a_SAP"/"Toma de Inventario V3 AP.py", encoding="utf-8").read())
#dfInvenSAPBO = pd.read_table(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"tomadeinventariov3ap.txt", sep='\t',encoding='utf-8',engine='python')
dfInvenSAPBO = pd.read_table(
    base / "tomadeinventariov3ap.txt",
    sep="\t",
    encoding="utf-8",
    engine="python"
)

#dfInvenSAPBOAstec= dfInvenSAPBO[dfInvenSAPBO['Nombre de grupo'].isin(['LUBRICACION INDUSTRI', 'SERVICIOS LUBR INDUS'])].copy() # Solo se considera las nombres de grupo 'ASTEC' y 'SERVICIOS ASTEC'
dfInvenSAPBOAstec = dfInvenSAPBO[dfInvenSAPBO['Nombre de grupo'].isin(Lineas_Asociadas)].copy()
dfInvenSAPBOAstec['AlmacenG'] = dfInvenSAPBOAstec['AlmacenG'].astype(str)
dfInvenSAPBOAstec["Cantidad en almacén"]=dfInvenSAPBOAstec["Cantidad en almacén"].fillna(0)

# Agrupa por "Sociedad", "Número de artículo" y "primer_dia_mes" y suma la columna "Cantidad unificada"
ResumenInvSAPBOAstec = dfInvenSAPBOAstec.groupby(['Número de artículo','AlmacenG'])['Cantidad en almacén'].sum().reset_index()
#ResumenInvSAPBOAstec['Contrato'] = np.where(ResumenInvSAPBOAstec['Número de artículo'].isin(dfcontrato['SAP']), 'SI', 'NO')
#ResumenInvSAPBOAstec = ResumenInvSAPBOAstec[~((ResumenInvSAPBOAstec['Contrato'] == 'SI') & (ResumenInvSAPBOAstec['AlmacenG'].isin([800, 801])))]
#ResumenInvSAPBOAstec = ResumenInvSAPBOAstec.drop(columns=['Contrato'])
ResumenInvSAPBOAstec.columns = ['SAP', 'Almacen', 'Stock']
_log.info("[Inventario] Proceso terminado")
#------------------------------------------------------------------------------------------------------------------------------

#Consolidacion Inventario Portal + SAP
dfConsInvPortalUniconSAP= ResumenInvSAPBOAstec




