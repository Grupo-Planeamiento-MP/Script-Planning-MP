import pandas as pd
import numpy as np
import os
from pathlib import Path
import getpass

### Detecta ruta del script y redirige las demas direcciones path ###
rutainicial = Path.home()
usuario = getpass.getuser()
antes, sep, despues = str(rutainicial).partition(usuario)
base = Path(antes + sep)

#inventario YB portal unicon
# dfInvPortalUNICONyb= pd.read_excel("C:/Users/AnthonyPradoCornejo/OneDrive - MARCO PERUANA SA/Escritorio/Rotacion Incubadoras-Electro-Frio/ASTEC/Planeamiento de Abastecimiento/Inventario/Inventario_MARCO PERUANA S.A._MARCO PERUANA S.A. - HIERBABUENA.xls", header=3 ,
#     dtype={'Codigo Unicon': str})
# dfInvPortalUNICONyb = dfInvPortalUNICONyb.iloc[:, [1, 3, 12]] #seleccionar columnas deseadas
# dfInvPortalUNICONyb = dfInvPortalUNICONyb.dropna(how='all') # Eliminar filas donde todos los registros sean NaN
# dfInvPortalUNICONyb['Almacen'] = 801 # Agregar una columna "Almacen" con valor 800 para todas las filas restantes

# #inventario JIC portal unicon
# dfInvPortalUNICONjic= pd.read_excel("C:/Users/AnthonyPradoCornejo/OneDrive - MARCO PERUANA SA/Escritorio/Rotacion Incubadoras-Electro-Frio/ASTEC/Planeamiento de Abastecimiento/Inventario/Inventario_MARCO PERUANA S.A._MARCO PERUANA S.A. - JICAMARCA.xls", header=3 ,
#     dtype={'Codigo Unicon': str})
# dfInvPortalUNICONjic = dfInvPortalUNICONjic.iloc[:, [1, 3, 12]] #seleccionar columnas deseadas
# dfInvPortalUNICONjic = dfInvPortalUNICONjic.dropna(how='all') # Eliminar filas donde todos los registros sean NaN
# dfInvPortalUNICONjic['Almacen'] = 800 # Agregar una columna "Almacen" con valor 800 para todas las filas restantes

# #Consolidacion Inventario
# dfConsInvPortalUNICON = pd.concat([dfInvPortalUNICONyb, dfInvPortalUNICONjic], axis=0)

# #Conecta y guarda en un dataframe el excel del contrato de consignacion y sus parametros
# dfcontrato= pd.read_excel("C:/Users/AnthonyPradoCornejo/OneDrive - MARCO PERUANA SA/Escritorio/Rotacion Incubadoras-Electro-Frio/ASTEC/Planeamiento de Abastecimiento/ContratoCsgUNICON.xlsx",
#     dtype={'Código Articulo': str})

# dfConsInvPortalUNICON = dfConsInvPortalUNICON.merge(dfcontrato[['Código Articulo', 'SAP']],
#                                   left_on='Codigo Unicon', right_on='Código Articulo', how='left')
# dfConsInvPortalUNICON = dfConsInvPortalUNICON.drop(columns=['Código Articulo', 'Codigo Unicon', 'Item'])
# dfConsInvPortalUNICON = dfConsInvPortalUNICON[['SAP', 'Almacen', 'Stock']]

#-----------------------------------------------------------------------------------------------------------------------------

# Define la ruta del archivo Excel del Inventario de SAP BO
#dfInvenSAPBO= pd.read_table("C:/Users/AnthonyPradoCornejo/OneDrive - MARCO PERUANA SA/Escritorio/Rotacion Incubadoras-Electro-Frio/ASTEC/Planeamiento de Abastecimiento/Inventario/Toma de Inventario V3 AP.txt", encoding='utf-16', sep='\t', quotechar='"', low_memory=False)
#exec(open(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Conexiones_a_SAP"/"Toma de Inventario V3 AP.py", encoding="utf-8").read())
dfInvenSAPBO = pd.read_table(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"Toma de Inventario V3 AP.txt", sep='\t',encoding='utf-8',engine='python')

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

#------------------------------------------------------------------------------------------------------------------------------

#Consolidacion Inventario Portal + SAP
dfConsInvPortalUniconSAP= ResumenInvSAPBOAstec




