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
#Linea_Negocio="SIST DE LUBRICACION"
Linea_Negocio=Linea_Negocio

#exec(open(base/"C:/Users/planner01/MARCO PERUANA SA/Planeamiento de Inventarios - Documents/Proyectos/Python/Reporte de abastecimiento LI/Codigo/TransformCleanKardexLI.py", encoding="utf-8").read())
exec(open(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"TransformCleanKardex.py", encoding="utf-8").read())

# Define la ruta del archivo Excel
dfprevKardex= dfkardexorigen
#Conecta y guarda en un dataframe el excel del contrato de consignacion y sus parametros
#dfcontrato= pd.read_excel("C:/Users/AnthonyPradoCornejo/OneDrive - MARCO PERUANA SA/Escritorio/Rotacion Incubadoras-Electro-Frio/ASTEC/Planeamiento de Abastecimiento/ContratoCsgUNICON.xlsx")

# Solo se considera las nombres de grupo 'ASTEC' y 'SERVICIOS ASTEC'
#dfKardexASTEC= dfprevKardex[dfprevKardex['Nombre de grupo'].isin(['LUBRICACION INDUSTRI'])]
dfKardexASTEC = dfprevKardex[dfprevKardex['Nombre de grupo'].isin([Linea_Negocio])]
# Solo se consideran aquellos movimientos de inventario que representan venta , es decir salida de stock real del almacen
dfKardexASTEC = dfKardexASTEC[dfKardexASTEC['Personalizado'].isin(['Venta Directa', 'Servicio'])]
#dfKardexASTEC = dfKardexASTEC[dfKardexASTEC['Sociedad'].isin(['MP'])] #la linea solo presenta movimientos relevantes en MP

# Convierte la columna "Fecha de contabilización" a datetime
dfKardexASTEC['Fecha de contabilización'] = pd.to_datetime(dfKardexASTEC['Fecha de contabilización'], format='%d/%m/%Y', errors='coerce')

# Extrae el primer día del mes y año
dfKardexASTEC['month_year'] = dfKardexASTEC['Fecha de contabilización'].dt.to_period('M').dt.to_timestamp()

dfKardexASTEC['Codigo GET'] = dfKardexASTEC['Codigo GET'].fillna('ND')



"""

CAlCULOS DE LA FECHA DE REFERENCIA PARA EL PERIODO TOTAL
SOBRE EL CUAL SE CALCULARAN LOS PROMEDIOS DE CONSUMO
Y LA ROTACION DE LOS ARTICULOS

"""
dateactual = pd.to_datetime('today')
# Calcular el primer día del mes actual
primer_dia_mes_actual = dateactual.replace(day=1)
# Calcular el último día del mes anterior
primer_dia_mes_anterior = primer_dia_mes_actual - pd.DateOffset(months=1)
# Asignar el último día del mes anteriodo de referencia
reference_period = primer_dia_mes_actual.normalize()
end_period = reference_period - pd.DateOffset(months=24) #considerar data de 2 años para trabajar

"""
RECALCULAR LOS INDICES PARA CADA ARTICULO

A18110000030   2023-02   1
A18110000030   2023-02   0
A18110000030   2023-02   0
A18110000030   2023-02   2

"""

# Agrupa por "Sociedad", "Número de artículo" y "primer_dia_mes" y suma la columna "Cantidad unificada"
ResumenPrevKardexSAP = dfKardexASTEC.groupby(['Número de artículo','Código de almacén', 'month_year'])[['Cantidad unificada','Valor unificado']].sum().reset_index()
#df_filtrado = ResumenPrevKardexSAP[ResumenPrevKardexSAP["SAP"] == "A18110006942"]
#####
dfKardexASTEC.rename(columns={"Nombre_Cliente": "Cliente_Final"}, inplace=True)
####
ResumenPrevKardexSAPcliente = dfKardexASTEC.groupby(['Número de artículo','Cliente_Final', 'month_year'])[['Cantidad unificada','Valor unificado']].sum().reset_index()

# Generar las fechas faltantes solo para combinaciones existentes
fechas_fijas = pd.date_range(end_period, reference_period, freq='MS')

#Obtener todas las combinaciones reales de 'Número de artículo', 'Codigo GET', 'Código de almacén'
comb_existentes = ResumenPrevKardexSAP[['Número de artículo', 'Código de almacén']].drop_duplicates()
comb_existentes_c = ResumenPrevKardexSAPcliente[['Número de artículo', 'Cliente_Final']].drop_duplicates()

# 4️⃣ Crear DataFrame base combinando las combinaciones existentes con todas las fechas fijas
base_fechas = comb_existentes.assign(key=1).merge(pd.DataFrame({'month_year': fechas_fijas, 'key': 1}), on='key').drop(columns='key')
base_fechas_cliente = comb_existentes_c.assign(key=1).merge(pd.DataFrame({'month_year': fechas_fijas, 'key': 1}), on='key').drop(columns='key')

# 5️⃣ Hacer un merge con los datos reales para completar los valores faltantes con 0
ResumenPrevKardexSAP = base_fechas.merge(ResumenPrevKardexSAP, on=['Número de artículo', 'Código de almacén', 'month_year'], how='left').fillna(0)
ResumenPrevKardexSAPcliente = base_fechas_cliente.merge(ResumenPrevKardexSAPcliente, on=['Número de artículo', 'Cliente_Final', 'month_year'], how='left').fillna(0)

# 6️⃣ Renombrar columnas para claridad (opcional)
ResumenPrevKardexSAP = ResumenPrevKardexSAP.rename(columns={'Número de artículo': 'SAP', 'Cantidad unificada': 'Consumo Total', 'Valor unificado': 'Valor Total Soles'})
ResumenPrevKardexSAPcliente = ResumenPrevKardexSAPcliente.rename(columns={'Número de artículo': 'SAP', 'Cantidad unificada': 'Consumo Total', 'Valor unificado': 'Valor Total Soles'})

#Agrupa por "Sociedad", "Número de artículo" y "primer_dia_mes" y suma la columna "Cantidad unificada"
ResumenPrevKardexwhtAlmacen = ResumenPrevKardexSAP.groupby(['SAP','month_year'])[['Consumo Total', 'Valor Total Soles']].sum().reset_index()

# Consumos de este mes
dfConsumoActual = ResumenPrevKardexwhtAlmacen[(ResumenPrevKardexwhtAlmacen['month_year'] == primer_dia_mes_actual.normalize())]

#Regularizar
ResumenPrevKardexwhtAlmacen = ResumenPrevKardexwhtAlmacen[(ResumenPrevKardexwhtAlmacen['month_year'] < primer_dia_mes_actual.normalize())]
ResumenPrevKardexSAPcliente = ResumenPrevKardexSAPcliente[(ResumenPrevKardexSAPcliente['month_year'] < primer_dia_mes_actual.normalize())]

"""Clasificación ABC en base al valor de las transacciones del ultimo año"""
# Definir el rango de fechas
fecha_inicio = (reference_period - pd.DateOffset(months=11)).replace(day=1)
fecha_fin = reference_period.replace(day=1) + pd.DateOffset(months=1) - pd.Timedelta(days=1)

# Filtrar las transacciones dentro del rango
dfUltimos12M = dfKardexASTEC[(dfKardexASTEC['Fecha de contabilización'] >= fecha_inicio) & (dfKardexASTEC['Fecha de contabilización'] <= fecha_fin)]

# Resumir por 'Número de artículo' y sumar 'Valor unificado'
dfParetoABC = dfUltimos12M.groupby('Número de artículo', as_index=False)['Valor unificado'].sum()

# Asegurar que los valores negativos sean 0
dfParetoABC['Valor unificado'] = dfParetoABC['Valor unificado'].clip(lower=0)

# Ordenar en orden descendente
dfParetoABC = dfParetoABC.sort_values(by='Valor unificado', ascending=False)

# Calcular el total acumulado y el porcentaje acumulado
dfParetoABC['Porcentaje acumulado'] = dfParetoABC['Valor unificado'].cumsum() / dfParetoABC['Valor unificado'].sum()

# Asignar clasificación basada en el diagrama de Pareto
dfParetoABC['Clasificación Pareto'] = 'C'
dfParetoABC.loc[dfParetoABC['Porcentaje acumulado'] <= 0.8, 'Clasificación Pareto'] = 'A'
dfParetoABC.loc[(dfParetoABC['Porcentaje acumulado'] > 0.8) & (dfParetoABC['Porcentaje acumulado'] <= 0.95), 'Clasificación Pareto'] = 'B'
