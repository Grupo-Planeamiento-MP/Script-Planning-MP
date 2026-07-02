## LIBRERIAS ##
import pandas as pd
import numpy as np
import os
from pathlib import Path
import getpass
import requests
from calendar import monthrange
from datetime import datetime
from openpyxl.utils import get_column_letter
import calendar

## GLOBALES ##
### Detecta ruta del script ###
rutainicial = Path.home()
usuario = getpass.getuser()
antes, sep, despues = str(rutainicial).partition(usuario)
base = Path(antes + sep)


Linea_Negocio="EQUIPOS TRANS. MATER"
Lineas_Asociadas = ["EQUIPOS TRANS. MATER"]



## SCRIPTS ##
exec(open(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"AnalisisConsumosSAP.py", encoding="utf-8").read()) #ResumenPrevKardexwhtAlmacen
exec(open(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"CalculoStockSeguridadLTvariabilityV2.py", encoding="utf-8").read()) #calcular_stock_seguridad_ltvar_opt  identificar_outliers
exec(open(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"FuncionesAbastecimiento.py", encoding="utf-8").read())  #agregar_consumo_promedio ,agregar_consumo_prom_compo, agregar_promedio_forecast_3m_mineria, agregar_fixedcost, agregar_valor_consumo, calcular_totaltransito_comp
exec(open(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"Inventario.py", encoding="utf-8").read()) #ResumenInvSAPBOAstec
exec(open(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"ConexionDatos.py", encoding="utf-8").read()) #dfResumenLeadTimes , df_ensambleETM, #Reporte_Precios_Local_Imp.txt
exec(open(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"CalculoStock0.py", encoding="utf-8").read())
dfActiveCode = pd.read_table(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"dfActiveCode.txt", sep='\t',encoding='utf-8',engine='python')
dfActiveCode = dfActiveCode.query("`Linea de Negocio` in @Lineas_Asociadas")
#dfActiveCode = dfActiveCode[dfActiveCode['Codigo_SAP'].isin(['A18110000283'])]

###
linea_reemplazo = ["EQUIPOS TRANS. MATER"]
###

"""AJUSTES REMPLAZOS"""
# Resumir cuadro de settings
#dfMRP_filter = dfMRP[['Producto Final', 'Componente', 'Q']]
# Obtener item vigente por grupo
vigentes = (
    df_remplazos[df_remplazos["Estado de Remplazo"] == "VIGENTE"]
    [["Grupo Remplazo", "Item"]]
    .rename(columns={"Item": "Componente"})
    #.rename(columns={"Item": "Producto Final"})
)
# Obtener items no vigentes por grupo
no_vigentes = (
    df_remplazos[df_remplazos["Estado de Remplazo"] != "VIGENTE"]
    [["Grupo Remplazo", "Item","Q"]]
    .rename(columns={"Item": "Producto Final"})
    #.rename(columns={"Item": "Componente"})
)
# Cruce para asignar a cada no vigente su vigente del mismo grupo
nuevas_filas = no_vigentes.merge(
    vigentes,
    on="Grupo Remplazo",
    how="left"
)

# Quedarse solo con las columnas requeridas
nuevas_filas = nuevas_filas[["Producto Final", "Componente", "Q"]]
###
if Linea_Negocio in linea_reemplazo:
    dfMRP_filter = nuevas_filas.copy()
else:

    dfMRP_filter = dfMRP[["Producto Final", "Componente", "Q"]].copy()
    dfMRP_filter = pd.concat([dfMRP_filter, nuevas_filas],ignore_index=True)
###
# Agregar las nuevas filas a relacionPF_Comp
#dfMRP_filter = pd.concat([dfMRP_filter, nuevas_filas], ignore_index=True)
###ajuste remplazos##
dfMRP_filter.set_index('Producto Final', inplace=True)

""" TRANSFORMAR PF A COMP FORECAST"""
df_forecast = forecast_comp(df_forecast,dfMRP_filter) ###

"""AGREGAR LOS DATOS MAS RELEVANTES AL RESUMEN DE ARTICULOS DE LA LINEA"""
dfActiveCodeMIN= dfActiveCode[['Codigo_SAP', 'UM']]
dfActiveCodeMIN.set_index('Codigo_SAP', inplace=True)####
dfBOMfinal = pd.merge(dfActiveCodeMIN, dfMRP_filter, left_index=True, right_index=True, how='left')
dfBOMfinal.reset_index(inplace=True)
dfBOMfinal.rename(columns={'index': 'Codigo_SAP'}, inplace=True)
dfBOMfinal['Componente'] = dfBOMfinal['Componente'].fillna(dfBOMfinal['Codigo_SAP'])
dfBOMfinal['Q'] = dfBOMfinal['Q'].fillna(1)

"""CONSUMO DE COMPONENTES"""
# Fusionar BOM con el consumo de productos finales usando 'Producto Final' = 'SAP'
df_componente_consumo = dfBOMfinal.merge(ResumenPrevKardexwhtAlmacen, left_on='Codigo_SAP', right_on='SAP')
# Multiplicar el consumo de cada producto final por la cantidad requerida de cada componente
df_componente_consumo['Consumo Total'] = df_componente_consumo['Consumo Total'] * df_componente_consumo['Q']
# Agrupar por 'Componente' y 'month_year' sumando los consumos
df_componente_consumo = df_componente_consumo.groupby(['Componente', 'month_year'])['Consumo Total'].sum().reset_index()
df_componente_consumo = df_componente_consumo.rename(columns={"Componente": "SAP"})
#df_componente_consumo: tiene todos componenetes con conversion y el resto , no contiene los productos finales 

"""AJUSTE LISTA DE ARTICULOS"""
columnas_mantener = [
    "Codigo_SAP",
    "Codigo_GET",
    "Codigo_Barra",
    "Categoria",
    "Ferreteria",
    "Marca",
    "Descripcion",
    "UM",
    "Costo S/.",
    "Comprometido",
    "En OV",
    "En OF",
    "En Transito",
    "Linea de Negocio"
]
dfActiveCode = dfActiveCode[columnas_mantener]
dfActiveCode.drop(columns = ['En Transito', 'Codigo_Barra', 'Ferreteria', 'Linea de Negocio'], axis=1, inplace=True)

"""LIMPIEZA LEADTIMES"""
df_articulos = dfActiveCode.rename(columns={"Codigo_SAP": "ItemCode"})
exec(open(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Reporte para lineas"/"Codigo"/"Abastecimiento"/"LimpiezaTiemposTransitov2.py", encoding="utf-8").read())

"""LIMPIEZA CONSUMOS"""
df_consumos_componentes , df_history , df_consumos_standar = limpieza_consumos_setting(ResumenPrevKardexwhtAlmacen,df_componente_consumo,dfMRP_filter,reference_period)    
df_consumos_componentes_so = df_consumos_componentes[df_consumos_componentes['Outlier'] == 'NO']
df_consumos_standar_so = df_consumos_standar[df_consumos_standar['Outlier'] == 'NO']

"""COMPARACION DE PRECIOS"""
df_resultado_ltvar = agregar_compar_precios(dfActiveCode, Reporte_Precios_Local_Imp, 'MP')

"""CALCULAR COMPROMETIDO POR COMPONENTE"""
df_resultado_ltvar = calcular_comprometido_comp(df_resultado_ltvar, dfBOMfinal)

"""AGREGAR LOS GRUPOS"""
df_resultado_ltvar = asignar_grupo_y_jerarquia(df_resultado_ltvar, dfBOMfinal)

"""SS y ROP"""
df_resultado_ltvar = calcular_stock_seguridad_ltvar(df_resultado_ltvar, df_consumos_componentes_so, df_LeadTime, reference_period)
###
multiplo_ensamble = dict(zip(df_ensambleETM["SAP MARCO"],df_ensambleETM["CANT. ENSAM"]))
###
"""
if Linea_Negocio == "EQUIPOS TRANS. MATER":
    for idx, row in df_resultado_ltvar.iterrows():
        sap = row.get("Codigo_SAP")
        if sap in multiplo_ensamble:
            multiplo = pd.to_numeric(multiplo_ensamble[sap], errors="coerce")
            if pd.notna(multiplo) and multiplo > 0:
                prom = row.get("Prom_Cons_LT_ltvar")
                if pd.notna(prom):
                    # ===================== 95% =====================
                    if pd.notna(row.get("SS_95_ltvar")):
                        # Ajustar SS
                        ss95 = np.ceil(row["SS_95_ltvar"] / multiplo) * multiplo
                        # Calcular P95 preliminar
                        p95 = prom + ss95
                        # Ajustar P95 al múltiplo
                        p95_ajustado = np.ceil(p95 / multiplo) * multiplo
                                           
                        df_resultado_ltvar.at[idx, "SS_95_ltvar"] = int(ss95)
                        df_resultado_ltvar.at[idx, "P95_ltvar"] = int(p95_ajustado)
                    # ===================== 98% =====================
                    if pd.notna(row.get("SS_98_ltvar")):
                        ss98 = np.ceil(row["SS_98_ltvar"] / multiplo) * multiplo
                        p98 = prom + ss98
                        p98_ajustado = np.ceil(p98 / multiplo) * multiplo          
                        df_resultado_ltvar.at[idx, "SS_98_ltvar"] = int(ss98)
                        df_resultado_ltvar.at[idx, "P98_ltvar"] = int(p98_ajustado)
"""
###

"""CLASIFICACION"""
if Linea_Negocio in ["EQUIPOS TRANS. MATER"]:
    columnas = ["Codigo_SAP", "Clasificación General", "Contratos"]
else:
    columnas = ["Codigo_SAP", "Clasificación General"]
###
#columnas = ["Codigo_SAP", "Clasificación General"]
df_resultado_ltvar = df_resultado_ltvar.merge(df_clasificacion[columnas],how="left")
df_resultado_ltvar = clasificacion_comp(df_resultado_ltvar,dfMRP_filter,df_clasificacion)


"""AGREGAR COBERTURA STOCK DE SEGURIDAD Y PUNTO MAXIMO"""
#df_resultado_ltvar = agregar_coberturas_y_ajustes(df_resultado_ltvar)

"""AGREGAR LA ROTACION"""
df_resultado_ltvar = agregar_categoria_rotacion(df_resultado_ltvar, ResumenPrevKardexwhtAlmacen, reference_period)

"""AGREGAR LA ROTACION COMPONENTE"""
df_resultado_ltvar = agregar_categoria_rotacion_comp(df_resultado_ltvar, df_componente_consumo, reference_period)

"""AGREGAR EL INVENTARIO DE MINERIA"""
#df_resultado_ltvar = agregar_stock_disponible_comp(df_resultado_ltvar, dfConsInvPortalUniconSAP, dfBOMfinal)
if Linea_Negocio in ["EQUIPOS TRANS. MATER"]:
    df_resultado_ltvar = agregar_stock_disponible_comp(df_resultado_ltvar, dfConsInvPortalUniconSAP, dfBOMfinal)

    def agregar_stock_disponibleAST(dfActiveCode, dfConsInvPortalUniconSAP):
        dfConsInvPortalUniconSAP["Almacen"] = dfConsInvPortalUniconSAP["Almacen"].astype(str)
        callaoycd = {'100','101','102','105','106','110','114','113','400','495','395','500','502','503'}
        jic = {'800'}
        hb = {'801'}
        toromocho = {'802'}
        antamina = {'803'}
        df_callao = dfConsInvPortalUniconSAP[dfConsInvPortalUniconSAP["Almacen"].isin(callaoycd)]
        df_jic = dfConsInvPortalUniconSAP[dfConsInvPortalUniconSAP["Almacen"].isin(jic)]
        df_hb = dfConsInvPortalUniconSAP[dfConsInvPortalUniconSAP["Almacen"].isin(hb)]
        df_toro = dfConsInvPortalUniconSAP[dfConsInvPortalUniconSAP["Almacen"].isin(toromocho)]
        df_anta = dfConsInvPortalUniconSAP[dfConsInvPortalUniconSAP["Almacen"].isin(antamina)]
        stock_callao = df_callao.groupby("SAP")["Stock"].sum().rename("StockCallao")
        stock_jic = df_jic.groupby("SAP")["Stock"].sum().rename("StockJIC")
        stock_hb = df_hb.groupby("SAP")["Stock"].sum().rename("StockHB")
        stock_por_sap = pd.concat([stock_callao, stock_jic, stock_hb], axis=1).reset_index()
        stock_por_sap.rename(columns={"SAP": "Codigo_SAP"}, inplace=True)
        dfActiveCode = dfActiveCode.merge(stock_por_sap, on="Codigo_SAP", how="left")
        columnas_base = ["StockCallao","StockJIC","StockHB"]
        dfActiveCode[columnas_base] = dfActiveCode[columnas_base].fillna(0)
        if df_toro.empty:
            dfActiveCode["StockToromocho"] = 0
        else:
                stock_toro = df_toro.groupby("SAP")["Stock"].sum()
                dfActiveCode = dfActiveCode.merge(stock_toro.rename("StockToromocho"),left_on="Codigo_SAP",right_index=True,how="left").fillna({"StockToromocho":0})
        if df_anta.empty:
            dfActiveCode["StockAntamina"] = 0
        else:
                stock_anta = df_anta.groupby("SAP")["Stock"].sum()
                dfActiveCode = dfActiveCode.merge(stock_anta.rename("StockAntamina"),left_on="Codigo_SAP",right_index=True,how="left").fillna({"StockAntamina":0})    
        return dfActiveCode

    df_resultado_ltvar = agregar_stock_disponibleAST(df_resultado_ltvar, dfConsInvPortalUniconSAP)
    df_resultado_ltvar = df_resultado_ltvar.merge(df_Unicon_SAP_Cod[['SAP', 'Código Articulo']],left_on='Codigo_SAP', right_on='SAP', how='left')
    df_resultado_ltvar = df_resultado_ltvar.drop(columns=["SAP"])
    #Jicamarca
    df_resultado_ltvar = df_resultado_ltvar.merge(df_Jicamarca[['Codigo Unicon', 'Stock']],left_on='Código Articulo', right_on='Codigo Unicon', how='left')
    df_resultado_ltvar = df_resultado_ltvar.drop(columns=["Codigo Unicon"])
    df_resultado_ltvar = df_resultado_ltvar.rename(columns={"Stock": "StockJIC_Portal"})  
    #HierbaBuena
    df_resultado_ltvar = df_resultado_ltvar.merge(df_HierbaBuena[['Codigo Unicon', 'Stock']], left_on='Código Articulo', right_on='Codigo Unicon', how='left')
    df_resultado_ltvar = df_resultado_ltvar.drop(columns=["Codigo Unicon"])
    df_resultado_ltvar = df_resultado_ltvar.rename(columns={"Stock": "StockHB_Portal"})
    # Toromocho
    df_resultado_ltvar = df_resultado_ltvar.merge(df_Toromocho[['Codigo Unicon', 'Stock']], left_on='Código Articulo',right_on='Codigo Unicon', how='left')
    df_resultado_ltvar = df_resultado_ltvar.drop(columns=["Codigo Unicon"])
    df_resultado_ltvar = df_resultado_ltvar.rename(columns={"Stock": "StockToromocho_Portal"})
    # Antamina
    df_resultado_ltvar = df_resultado_ltvar.merge(df_Antamina[['Codigo Unicon', 'Stock']], left_on='Código Articulo',right_on='Codigo Unicon', how='left')
    df_resultado_ltvar = df_resultado_ltvar.drop(columns=["Codigo Unicon"])
    df_resultado_ltvar = df_resultado_ltvar.rename(columns={"Stock": "StockAntamina_Portal"})
    
    df_resultado_ltvar[["StockJIC_Portal","StockHB_Portal", "StockToromocho_Portal","StockAntamina_Portal"]] = df_resultado_ltvar[[ "StockJIC_Portal","StockHB_Portal","StockToromocho_Portal", "StockAntamina_Portal"]].fillna(0)
    cols = ["Stock Disponible", "StockJIC", "StockHB",  "StockToromocho", "StockAntamina","StockJIC_Portal", "StockHB_Portal", "StockToromocho_Portal","StockAntamina_Portal"]
    for c in cols:
        df_resultado_ltvar[c] = df_resultado_ltvar[c].astype(float)      
    mask = df_resultado_ltvar["Código Articulo"].notna() & (df_resultado_ltvar["Código Articulo"] != "")
    df_resultado_ltvar.loc[mask, "Stock Disponible"] = df_resultado_ltvar.loc[mask, "Stock Disponible"] - df_resultado_ltvar.loc[mask, "StockJIC"] - df_resultado_ltvar.loc[mask, "StockHB"]  - df_resultado_ltvar.loc[mask, "StockToromocho"] - df_resultado_ltvar.loc[mask, "StockAntamina"] + df_resultado_ltvar.loc[mask, "StockJIC_Portal"] + df_resultado_ltvar.loc[mask, "StockHB_Portal"]+  df_resultado_ltvar.loc[mask, "StockToromocho_Portal"]+ df_resultado_ltvar.loc[mask, "StockAntamina_Portal"] 
    columnas_a_eliminar = ["StockCallao","StockJIC","StockHB",  "StockToromocho","StockAntamina", "StockJIC_Portal","StockHB_Portal", "StockToromocho_Portal", "StockAntamina_Portal", "Código Articulo"]
    df_resultado_ltvar = df_resultado_ltvar.drop(columns=columnas_a_eliminar) 
    
    dfBOMfinal["Q"] = pd.to_numeric(dfBOMfinal["Q"], errors="coerce")
    stock_actualizado = df_resultado_ltvar[ ["Codigo_SAP", "Stock Disponible", "Stock no Disponible"]].copy()
    stock_equivalente = stock_actualizado.merge( dfBOMfinal, on="Codigo_SAP",how="inner")
    stock_equivalente["Stock Disponible-C"] = ( stock_equivalente["Stock Disponible"]* stock_equivalente["Q"])
    stock_equivalente["Stock no Disponible-C"] = (stock_equivalente["Stock no Disponible"]* stock_equivalente["Q"])
    stock_componentes = stock_equivalente.groupby("Componente")[["Stock Disponible-C", "Stock no Disponible-C"]].sum().reset_index()
    stock_componentes.rename(columns={"Componente": "Codigo_SAP"}, inplace=True)
    df_resultado_ltvar = df_resultado_ltvar.drop(  columns=["Stock Disponible-C", "Stock no Disponible-C"],  errors="ignore")
    df_resultado_ltvar = df_resultado_ltvar.merge( stock_componentes, on="Codigo_SAP", how="left")
    df_resultado_ltvar[[ "Stock Disponible-C", "Stock no Disponible-C"]] = df_resultado_ltvar[[ "Stock Disponible-C", "Stock no Disponible-C"]].fillna(0)

      
else:
    df_resultado_ltvar = agregar_stock_disponible_comp(df_resultado_ltvar, dfConsInvPortalUniconSAP, dfBOMfinal)

"""AGREGAR CONSUMO PROMEDIO"""
df_resultado_ltvar = agregar_consumo_promedio(df_resultado_ltvar, df_consumos_standar_so,reference_period)

"""AGREGAR CONSUMO PROMEDIO POR COMPONENTE"""
df_resultado_ltvar = agregar_consumo_prom_compo(df_resultado_ltvar, df_consumos_componentes_so, reference_period)

"""AGREGAR CONSUMO PROMEDIO POR CLIENTE"""
df_resultado_ltvar = agregar_consumo_promedio_cliente(df_resultado_ltvar, ResumenPrevKardexSAPcliente, reference_period)

"""AGREGAR CLASIFICACION DE PARETO"""
# Realizar el merge para incluir la clasificación Pareto en df_resultado
df_resultado_ltvar = df_resultado_ltvar.merge(dfParetoABC[['Número de artículo', 'Clasificación Pareto']], 
                                  left_on='Codigo_SAP', 
                                  right_on='Número de artículo', 
                                  how='left')
# Renombrar la columna agregada
df_resultado_ltvar.rename(columns={'Clasificación Pareto': 'Pareto'}, inplace=True)
# Rellenar con 'C' si no se encontró clasificación
df_resultado_ltvar['Pareto'].fillna('C', inplace=True)
# Eliminar la columna 'Número de artículo' generada en el merge
df_resultado_ltvar.drop(columns=['Número de artículo'], inplace=True)

"""AGREGAR VALOR DE CONSUMO"""
df_resultado_ltvar = agregar_valor_consumo(df_resultado_ltvar, ResumenPrevKardexwhtAlmacen, reference_period)
    
"""AGREGAR EL TRANSITO TOTAL"""
#df_resultado_ltvar = calcular_totaltransito(df_resultado_ltvar, FreservSAP_filtrado, SeguiBOSAP_f)

"""AGREGAR EL TRANSITO TOTAL POR COMPONENTE"""
df_resultado_ltvar = calcular_totaltransito_comp(df_resultado_ltvar, FreservSAP_filtrado, SeguiBOSAP_f, dfBOMfinal)
# Rotacion ficticia para asignacion de valores
df_resultado_ltvar['Rotacion_F'] = df_resultado_ltvar['Categoria Rot Comp'].replace(0, pd.NA).fillna(df_resultado_ltvar['Categoria Rot'])
 
"""AGREGAR EL TRANSITO POR FECHAS"""
hoy = pd.Timestamp.today().normalize()
df_resultado_ltvar = generar_transito_por_fechas(df_resultado_ltvar, FreservSAP_filtrado, SeguiBOSAP_f)
df_resultado_ltvar["Fecha_LLegada"] = (df_resultado_ltvar["Fecha_LLegada"].fillna(hoy))
df_resultado_ltvar["Fecha_LLegada"] = pd.to_datetime(df_resultado_ltvar["Fecha_LLegada"], errors="coerce")
df_resultado_ltvar["Fecha_LLegada"] = df_resultado_ltvar["Fecha_LLegada"].dt.normalize()

"""AGREGAR STOCK 0"""
if Linea_Negocio == "EQUIPOS TRANS. MATER":
    exec(open(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"CalculoStock0ETM.py", encoding="utf-8").read())

#Agregar Stock Cero
df_resultado_ltvar = pd.merge(df_resultado_ltvar, resultado_Stock_Cero[['ItemCode', 'Stock Cero (%)']], left_on='Codigo_SAP', right_on='ItemCode', how='left')
df_resultado_ltvar = df_resultado_ltvar.drop('ItemCode', axis=1)
df_resultado_ltvar['Stock Cero (%)'] = df_resultado_ltvar['Stock Cero (%)'].fillna(1)

"""LIMPIEZA CAMPOS NUMERICOS"""
columnas_a_limpieza = [
 "Comprometido",
 "Comprometido - C",
 "Stock no Disponible", 
 "Stock Disponible", 
 "Stock no Disponible-C", 
 "Stock Disponible-C",
 "Transito Total",
 "Transito Total - C",
 "Consumo Promedio 12M",
 "Consumo Promedio Comp 12M",
 "SS_95_ltvar",
 "SS_98_ltvar",
 "Prom_Cons_LT_ltvar",
 "P95_ltvar",
 "P98_ltvar",
 "Promedio_LT"
] 


df_resultado_ltvar[columnas_a_limpieza] = df_resultado_ltvar[columnas_a_limpieza].apply(
    pd.to_numeric, errors="coerce"
)
                  
#estandarizar nombres
df_resultado_ltvar = df_resultado_ltvar.rename(columns={
    "Clasificación General": "Clasificación",
    "L/I": "Tipo de Abastecimiento"
})

"""AGREGAR POSICION DE INVENTARIO """
#df_resultado_ltvar["Posicion de Inventario (PI)"] = (df_resultado_ltvar["Stock Disponible-C"] - df_resultado_ltvar["Comprometido - C"] + df_resultado_ltvar["Transito Total - C"])
df_resultado_ltvar["Posicion de Inventario (PI)"] = (df_resultado_ltvar["Stock Disponible-C"]+ df_resultado_ltvar["Transito Total - C"])

"""AGREGAR COBERTURA STOCK"""
df_resultado_ltvar["Cobertura Stock Fisico"]=df_resultado_ltvar["Stock Disponible-C"]/df_resultado_ltvar["Consumo Promedio Comp 12M"]

"""AGREGAR COBERTURA STOCK + TRANSITO"""
df_resultado_ltvar["Cobertura Stock + Transito"]=df_resultado_ltvar["Posicion de Inventario (PI)"]/df_resultado_ltvar["Consumo Promedio Comp 12M"]

"""AGREGAR NIVEL DE SERVICIO"""
df_resultado_ltvar["Nivel de Servicio"] = np.where(   
    # REGRESAR df_resultado_ltvar["Categoria Rot Comp"] == "ALTA",
    df_resultado_ltvar["Clasificación"] == "Estrategico",
    ">=95%",
    np.where(
        #Regresar df_resultado_ltvar["Categoria Rot Comp"] == "MEDIA",
        df_resultado_ltvar["Clasificación"] == "Revision Periodica",
        ">=90%",
        ""
    )
)

"""AGREGAR STOCK MINIMO"""
df_resultado_ltvar["Stock Minimo"] = np.where(
    df_resultado_ltvar["Clasificación"] == "Estrategico",
    #regresar df_resultado_ltvar["Categoria Rot Comp"] == "ALTA",
    df_resultado_ltvar["SS_98_ltvar"],
    np.where(
        df_resultado_ltvar["Clasificación"].isin(
            ["Revision Periodica", "Gestion Simplificada"]
            #regresar ["MEDIA"]
        ),
        df_resultado_ltvar["SS_95_ltvar"],
        np.nan
    )
)

"""AGREGAR COBERTURA STOCK MINIMO"""
#Evitar errores con division entre 0
df_resultado_ltvar["Cobertura (Stock Minimo)"] = 0
mask = df_resultado_ltvar["Consumo Promedio Comp 12M"] != 0
df_resultado_ltvar.loc[mask, "Cobertura (Stock Minimo)"] = (
    df_resultado_ltvar.loc[mask, "Stock Minimo"] /
    df_resultado_ltvar.loc[mask, "Consumo Promedio Comp 12M"]
)


"""AGREGAR PUNTO DE REORDEN"""
df_resultado_ltvar["Punto de reorden"] = np.where(
    df_resultado_ltvar["Clasificación"] == "Estrategico",
    df_resultado_ltvar["SS_98_ltvar"] + df_resultado_ltvar["Promedio_LT"]*df_resultado_ltvar["Consumo Promedio Comp 12M"] ,
    np.where(
        df_resultado_ltvar["Clasificación"].isin(
            ["Revision Periodica", "Gestion Simplificada"]
        ),
        df_resultado_ltvar["SS_95_ltvar"] + df_resultado_ltvar["Promedio_LT"]*df_resultado_ltvar["Consumo Promedio Comp 12M"] ,
        np.nan
    )
)

"""AGREGAR COBERTURA PUNTO DE REORDEN"""

df_resultado_ltvar["Cobertura (Punto de Reorden)"]=df_resultado_ltvar["Punto de reorden"]/df_resultado_ltvar["Consumo Promedio Comp 12M"]

"""AGREGAR CONSUMO DEL MES ACTUAL"""
df_resultado_ltvar = pd.merge(
df_resultado_ltvar, 
dfConsumoActual[['SAP', 'Consumo Total']], 
left_on='Codigo_SAP', 
right_on='SAP', 
how='left')
df_resultado_ltvar = df_resultado_ltvar.drop('SAP', axis=1)
df_resultado_ltvar = df_resultado_ltvar.rename(columns={"Consumo Total": "Consumo Mes Actual"})
df_resultado_ltvar['Consumo Mes Actual'] = df_resultado_ltvar['Consumo Mes Actual'].fillna(0)

"""AGREGAR CONSUMO DEL MES ACTUAL - COMPONENTES"""
df_resultado_ltvar=calcular_Consumo_Mes_Actual_comp(df_resultado_ltvar, dfBOMfinal)
cobertura_condicion="Cobertura Stock Fisico"
if Linea_Negocio == "LUBRICACION MINERIA":
   
    """AJUSTE UMBRAL DE SS MINIMO 3"""   
    mask = (df_resultado_ltvar["Cobertura (Stock Minimo)"] > 0) & \
       (df_resultado_ltvar["Cobertura (Stock Minimo)"] <= 3)
    # Columna diferencia
    df_resultado_ltvar.loc[mask, "Diferencia"] = 3 - df_resultado_ltvar.loc[mask, "Cobertura (Stock Minimo)"]  
    # Reemplazar Cobertura (Stock Minimo) = 3
    df_resultado_ltvar.loc[mask, "Cobertura (Stock Minimo)"] = 3    
    # Recalcular Stock Minimo
    df_resultado_ltvar.loc[mask, "Stock Minimo"] = (
        3 * df_resultado_ltvar.loc[mask, "Consumo Promedio Comp 12M"]
    )    
    # Ajustar Cobertura (Punto de Reorden)
    df_resultado_ltvar.loc[mask, "Cobertura (Punto de Reorden)"] = (
        df_resultado_ltvar.loc[mask, "Cobertura (Punto de Reorden)"] +
        df_resultado_ltvar.loc[mask, "Diferencia"]
    )
    # Recalcular Punto de Reorden
    df_resultado_ltvar.loc[mask, "Punto de reorden"] = (
        df_resultado_ltvar.loc[mask, "Cobertura (Punto de Reorden)"] *
        df_resultado_ltvar.loc[mask, "Consumo Promedio Comp 12M"]
    )
    # Eliminar columna Diferencia
    df_resultado_ltvar.drop(columns=["Diferencia"], inplace=True)
    
    """AGREGAR KPIS ULTIMO MES """
    df_resultado_ltvar=KPI_forecast(df_resultado_ltvar,df_KPI)

""" AGREGAR COBERTURA STOCK FISICO - FORECAST""" 
#df_forecast = df_forecast[df_forecast['SAP_Origen'].isin(['A18110001845'])]
df_resultado_ltvar=cobertura_forecast(df_resultado_ltvar, df_forecast, primer_dia_mes_actual,"Stock Disponible-C")
df_resultado_ltvar = df_resultado_ltvar.rename(columns={"Meses_Forecast": "Cobertura Stock Fisico - Forecast"})

""" AGREGAR COBERTURA Posicion de inventario - FORECAST"""   
df_resultado_ltvar=cobertura_forecast(df_resultado_ltvar, df_forecast, primer_dia_mes_actual,"Posicion de Inventario (PI)")
df_resultado_ltvar = df_resultado_ltvar.rename(columns={"Meses_Forecast": "Cobertura Stock + Transito - Forecast"})

"""AGREGAR  PUNTO DE REORDEN - FORECAST"""
df_resultado_ltvar=unidades_forecast(df_resultado_ltvar,df_forecast,"Cobertura (Punto de Reorden)",primer_dia_mes_actual)
df_resultado_ltvar = df_resultado_ltvar.rename(columns={"Unidades forecast": "Punto de reorden - Forecast"})

df_resultado_ltvar = ajustar_multiplo_ensamble(df_resultado_ltvar, "Punto de reorden - Forecast", multiplo_ensamble)


"""AGREGAR PUNTO DE REORDEN - FORECAST Ajustado"""
df_resultado_ltvar = unidades_despues_LT(df_resultado_ltvar, df_forecast,"Cobertura (Punto de Reorden)")
df_resultado_ltvar = df_resultado_ltvar.rename(columns={"Unidades_despues_LT": "Punto de reorden - Forecast Ajustado"})

df_resultado_ltvar = ajustar_multiplo_ensamble(df_resultado_ltvar,"Punto de reorden - Forecast Ajustado", multiplo_ensamble)


"""AGREGAR  STOCK MINIMO - FORECAST"""
df_resultado_ltvar=unidades_forecast(df_resultado_ltvar,df_forecast,"Cobertura (Stock Minimo)",primer_dia_mes_actual)
df_resultado_ltvar = df_resultado_ltvar.rename(columns={"Unidades forecast": "Stock Minimo - Forecast"}) 

df_resultado_ltvar = ajustar_multiplo_ensamble(df_resultado_ltvar,"Stock Minimo - Forecast", multiplo_ensamble)


"""AGREGAR SUGERENCIA DE COMPRA"""
df_resultado_ltvar = unidades_despues_LT(df_resultado_ltvar, df_forecast,"Promedio_LT")
df_resultado_ltvar.rename(columns={"Unidades_despues_LT": "Unidades_segundo_LT"}, inplace=True)
df_resultado_ltvar = sugerencia_compra_forecast_DOS(df_resultado_ltvar,df_forecast)

#df_forecast = df_forecast[df_forecast["SAP_Origen"] == "A18110000283"]
"""AGREGAR COBERTURA DE COMPRA"""
df_resultado_ltvar=cobertura_despues_LT(df_resultado_ltvar, df_forecast,"Sugerencia de Compra")
df_resultado_ltvar = df_resultado_ltvar.rename(columns={"Meses_Forecast": "Cobertura (Compra) - Forecast"})
cobertura_condicion="Cobertura Stock Fisico - Forecast"

"""AGREGAR ALERTA DE COMPRA"""
###esto con orden de compra
hoy = pd.Timestamp.today().normalize()
df_resultado_ltvar["Meses de Llegada"] = ((df_resultado_ltvar["Fecha_LLegada"] - hoy ).dt.days.div(30).round(2))

mask_sin_oc = df_resultado_ltvar["Transito Total - C"] <= 0 #mascara booleana T/F x fila
df_resultado_ltvar.loc[mask_sin_oc, "Meses de Llegada"] = (df_resultado_ltvar.loc[mask_sin_oc, "Promedio_LT"]) # x cada True redefine Meses de llegada
#alerta

df_resultado_ltvar["Dif_Alert"] = df_resultado_ltvar[cobertura_condicion]-df_resultado_ltvar["Meses de Llegada"]
df_resultado_ltvar["Alerta"] =  np.where(               
    df_resultado_ltvar["Tipo de Abastecimiento"]=="IMPORTADO",
        np.select(
        [
                df_resultado_ltvar["Dif_Alert"] > 1.5,#tiempo de reaccion - ss 
                df_resultado_ltvar["Dif_Alert"] <= 0,
                (df_resultado_ltvar["Dif_Alert"] > 0) & (df_resultado_ltvar["Dif_Alert"] <= 1.5)
            ],
            [
                "Sin Riesgo",
                "Quiebre de stock",
                "Riesgo Alto"
            ],
            default="-"
        ),
        #Local
        np.select(
        [
                df_resultado_ltvar["Dif_Alert"] > 0.23,
                df_resultado_ltvar["Dif_Alert"] <= 0,
                (df_resultado_ltvar["Dif_Alert"] > 0) & (df_resultado_ltvar["Dif_Alert"] <= 0.23)
            ],
            [
                "Sin Riesgo",
                "Quiebre de stock",
                "Riesgo Alto"
            ],
            default="-"
        )
    
    )

"""ORDENAR EL DATAFRAME FINAL"""
# Clasificacion ajustes
df_resultado_ltvar["Clasificación"] = df_resultado_ltvar["Clasificación"].replace(0, "Compra Calzada")

# Orden de prioridad de categorías
orden_categoria = {
    "Estrategico": 1,
    "Revision Periodica": 2,
    "Gestion Simplificada": 3,
    "Compra Calzada": 4
}

df_resultado_ltvar["Orden Categoria"] = df_resultado_ltvar["Clasificación"].map(orden_categoria)

df_resultado_ltvar = df_resultado_ltvar.sort_values(
    by=["Orden Categoria", "Valor de Consumo"],
    ascending=[True, False]
)

if Linea_Negocio in ["LUBRICACION MINERIA","HIDRAULI. COMPONENTE","EQUIPOS TRANS. MATER"]:
    """REPORTE SIMULACION IN/OUT"""
    ## Base de articulos a considerar ##
    #Sin productos finales
    dfMRP_filter = dfMRP_filter.reset_index()
    df_base_simulacion = df_resultado_ltvar[~df_resultado_ltvar["Codigo_SAP"].isin(dfMRP_filter["Producto Final"])]
    # Solo clasificacion elegida
    df_base_simulacion=df_base_simulacion[df_resultado_ltvar["Clasificación"]=="Estrategico"]
    df_base_simulacion=df_base_simulacion[["Codigo_SAP","Clasificación","Costo S/.","Stock Disponible-C","Consumo Mes Actual-C"]]
    
    ## Estimacion fechas ##
    df1=FreservSAP_filtrado[["Número de artículo","Restante","ETA Marco"]]
    df1['ETA Marco'] = pd.to_datetime(df1['ETA Marco'])
    df1 = df1.rename(columns={'Restante': 'Cantidad'})
    df2=SeguiBOSAP_f[["Número de artículo","Pendiente","ETA Marco"]]
    df2['ETA Marco'] = pd.to_datetime(df2['ETA Marco'])
    df2 = df2.rename(columns={'Pendiente': 'Cantidad'})
    df_ingresos = pd.concat([df1, df2], ignore_index=True)
    df_ingresos['Mes_año'] = pd.to_datetime(df_ingresos['ETA Marco']).dt.to_period('M').dt.to_timestamp()
    # Considerar componentes
    # Unir BOM con transito total
    df_ingresos_f = df_ingresos.merge(
        dfMRP_filter,
        left_on="Número de artículo",
        right_on="Producto Final",
        how="left"
    )
    # Rellenar 
    df_ingresos_f['Componente'] = (df_ingresos_f['Componente'].fillna(df_ingresos_f['Número de artículo']))
    df_ingresos_f['Q'] = df_ingresos_f['Q'].fillna(1)
    
    # Calcular comp
    df_ingresos_f["Cantidad - C"] = df_ingresos_f["Cantidad"] * df_ingresos_f["Q"]
    
    df_ingresos_f = (
        df_ingresos_f.groupby(['Mes_año', 'Componente'], as_index=False)
        [['Cantidad - C']]
        .sum()
    )
    df_ingresos_f = df_ingresos_f.rename(columns={'Componente': 'Número de artículo'})
    
    df_base_simulacion = df_base_simulacion.rename(columns={
        "Costo S/.": "Costo",
        "Stock Disponible-C": "Stock",
        "Clasificación": "Clasificacion",
        'Consumo Mes Actual-C':"Consumo_Mes_Actual"
    })
    ## Forecast ##
    #df_forecast #condicionamos que el primer de acuerdo al consumo mensual actual
    df_forecast = pd.merge(
    df_forecast, 
    df_base_simulacion[['Codigo_SAP', 'Consumo_Mes_Actual']], 
    left_on='SAP_Origen', 
    right_on='Codigo_SAP', 
    how='left')
    df_forecast = df_forecast.drop('Codigo_SAP', axis=1)
    mask = df_forecast['Fecha'] == primer_dia_mes_actual   
    df_forecast.loc[mask, 'Forecast'] = np.where(
        df_forecast.loc[mask, 'Consumo_Mes_Actual'] < df_forecast.loc[mask, 'Forecast'],
        df_forecast.loc[mask, 'Forecast'] - df_forecast.loc[mask, 'Consumo_Mes_Actual'],
        0
    )  
    df_forecast = df_forecast.sort_values(["SAP_Origen", "Fecha"]) 
    df_forecast = df_forecast.drop('Consumo_Mes_Actual', axis=1)
    ##Simulaciones##
    df_final=generar_reporte_simulacion(df_base_simulacion,df_ingresos_f,df_forecast)
    dfMRP_filter.set_index('Producto Final', inplace=True)

"""REPORTE PARA SETTING"""
# Producto → lista de componentes
mapa_prod_comp = (dfMRP_filter.reset_index().groupby('Producto Final')['Componente'].apply(lambda x: set(x.dropna().astype(str))).to_dict())
productos = list(mapa_prod_comp.keys())
visitados = set()
grupos = []

for prod in productos:
    if prod in visitados:
        continue

    grupo_actual = [prod]
    visitados.add(prod)

    comps_base = mapa_prod_comp[prod]

    for otro in productos:
        if otro in visitados:
            continue

        interseccion = comps_base.intersection(mapa_prod_comp[otro])
        union = comps_base.union(mapa_prod_comp[otro])

        similitud = len(interseccion) / len(union)

        if similitud >= 0.5:  # 
            grupo_actual.append(otro)
            visitados.add(otro)

    grupos.append(grupo_actual)
    
grupos = sorted(grupos, key=lambda g: len(g), reverse=True)
df_temp = df_resultado_ltvar.set_index("Codigo_SAP")

filas_setting = []
componentes_usados = set()

for grupo in grupos:

    comps_grupo = set()
    for prod in grupo:
        comps_grupo.update(mapa_prod_comp.get(prod, set()))

    
    # 1. productos
    for prod in grupo:
        if prod in df_temp.index:
            fila = df_temp.loc[[prod]].copy()
            fila["Producto Final"] = ""
            filas_setting.append(fila)

    # 2. componentes 
    for comp in sorted(comps_grupo):

        if comp in componentes_usados:
            continue

        if comp in df_temp.index:
            fila = df_temp.loc[[comp]].copy()

            productos_asociados = [
                p for p, comps in mapa_prod_comp.items()
                if comp in comps
            ]

            fila["Producto Final"] = " // ".join(productos_asociados)
           

            filas_setting.append(fila)
            componentes_usados.add(comp)
   
    
df_setting = pd.concat(filas_setting).reset_index()    
df_setting = ajustar_clasificacion(df_setting)

df_resultado_ltvar = df_resultado_ltvar.drop(columns=["Orden Categoria"], errors="ignore")
df_resultado_ltvar = df_resultado_ltvar.drop(columns=["Rotacion_F"], errors="ignore")

productos_setting = set(dfMRP_filter.index)
componentes_setting = set(dfMRP_filter["Componente"].dropna().astype(str))
codigos_setting = productos_setting.union(componentes_setting)
df_reporte = df_resultado_ltvar[ ~df_resultado_ltvar["Codigo_SAP"].astype(str).isin(codigos_setting)].copy()
df_reporte = df_reporte.drop(columns=["Producto Final"], errors="ignore")

"""AGREGAR COLUMNA TIPO"""
# Crear conjunto con los SAP de reemplazos
sap_remplazos = set(df_remplazos["Item"].astype(str))
# Crear columna Tipo en df_reporte
df_setting["Tipo"] = np.where(
    df_setting["Codigo_SAP"].astype(str).isin(sap_remplazos),
    "Remplazo",
    "Setting"
)

"""DEFINIR ORDEN DE LAS COLUMNAS"""
orden_columnas = [
    "Tipo",
    "Producto Final",
    "Codigo_SAP",
    "Codigo_GET",
    "Categoria",
    "Marca",
    "Descripcion",
    "UM",
    "Costo S/.",
    "Valor de Consumo",
    "Clasificación",
    "Clasificacion Ajustada",
]

if Linea_Negocio in ["EQUIPOS TRANS. MATER"]:
    orden_columnas.append("Contratos")

orden_columnas += [
    "Categoria Rot",
    "Categoria Rot Comp",
    "N° Meses",
    "Promedio_LT",
    "Observacion LT",
    "Tipo de Abastecimiento",
    "Comprometido",
    "Comprometido - C",
    "En OV",
    "En OF",
    "Stock no Disponible",
    "Stock Disponible",
    "Stock no Disponible-C",
    "Stock Disponible-C",
    "Transito Total",
    "Transito Total - C",
    "Transito por Fechas",
    "% Precision",
    "% MAPE",
    "RMSE",
    "BIAS",
    "Posicion de Inventario (PI)",
    "Consumo Promedio 12M",
    "Consumo Promedio Comp 12M",
    "Cobertura Stock Fisico",
    "Cobertura Stock Fisico - Forecast",
    "Cobertura Stock + Transito",
    "Cobertura Stock + Transito - Forecast",
    "Nivel de Servicio",
    "Stock Cero (%)",
    "Consumo Mes Actual",
    "Stock Minimo",
    "Stock Minimo - Forecast",
    "Cobertura (Stock Minimo)",
    "Punto de reorden",
    "Punto de reorden - Forecast",
    "Cobertura (Punto de Reorden)",
    "Sugerencia de Compra",
    "Cobertura (Compra) - Forecast",
    "Comparacion Precio",
    "Alerta",
    "Consumo Prom 12M Cliente",
    "CV_LT",
    "CV_Consumo"
    ]

if Linea_Negocio == "LUBRICACION INDUSTRI":
    columnas_eliminar = [
        "% Precision",
        "% MAPE",
        "RMSE",
        "BIAS",
    ]
    
    orden_columnas = [
        col for col in orden_columnas
        if col not in columnas_eliminar
    ]

if Linea_Negocio == "HIDRAULI. COMPONENTE":
    columnas_eliminar = [
        "% Precision",
        "% MAPE",
        "RMSE",
        "BIAS"
    ]
    
    orden_columnas = [
        col for col in orden_columnas
        if col not in columnas_eliminar
    ]
    

######
if Linea_Negocio == "EQUIPOS TRANS. MATER":
    columnas_eliminar = ["% Precision","% MAPE", "RMSE", "BIAS"]
    orden_columnas = [
        col for col in orden_columnas
        if col not in columnas_eliminar
    ]
    codigos_vsi = ["A20030000004","A20030000005","A20030000006","A20030000007","A20030000008", "A20030000009","A20030000011", "A20030000010", "A20030000023", "A20030000012", "A20030000022"]
    df_resultado_ltvar["Comentario"] = np.where( df_resultado_ltvar["Codigo_SAP"].isin(codigos_vsi), "VSIxAyS", "" )
    df_resultado_ltvar["Comentario"] = df_resultado_ltvar.apply(
        lambda row: (
            f"{row['Comentario']} // CANT ENSAMBLADA: {multiplo_ensamble[row['Codigo_SAP']]}"
            if row["Codigo_SAP"] in multiplo_ensamble and row["Comentario"] != ""
            else (
                f"CANT ENSAMBLADA: {multiplo_ensamble[row['Codigo_SAP']]}"
                if row["Codigo_SAP"] in multiplo_ensamble
                else row["Comentario"]
            )
        ),axis=1)
    df_setting = df_setting.merge(df_resultado_ltvar[["Codigo_SAP", "Comentario"]], on="Codigo_SAP",how="left")
    df_reporte = df_reporte.merge( df_resultado_ltvar[["Codigo_SAP", "Comentario"]], on="Codigo_SAP",how="left")
    if "Comentario" not in orden_columnas:
        orden_columnas.append("Comentario") 
############

df_setting = df_setting[orden_columnas]
df_reporte = df_reporte.reindex(columns=[col for col in orden_columnas if col not in ["Producto Final", "Clasificacion Ajustada"]])
df_reporte = df_reporte.drop(columns=["Categoria Rot Comp", "Consumo Promedio Comp 12M", "Stock Disponible-C","Stock no Disponible-C" ,"Transito Total - C"], errors="ignore")

# Crear diccionario: Producto Final -> Componente
dic_relacion = nuevas_filas.set_index('Producto Final')['Componente'].to_dict()
# Reemplazar solo los productos que existan en el diccionario
df_setting['Producto Final'] = df_setting['Producto Final'].replace(dic_relacion)

if Linea_Negocio == "HIDRAULI. COMPONENTE": 
    df_pesca = df_reporte[df_reporte["Codigo_SAP"].astype(str).isin(codigos_pesca)].copy()
    df_reporte = df_reporte[~df_reporte["Codigo_SAP"].astype(str).isin(codigos_pesca)].copy()
    df_pesca=df_pesca.drop(columns=["Tipo"])

"""EXPORTACION DE DATOS"""
mes_actual = pd.Timestamp.today()
mes_anio = mes_actual.strftime("%m%y")

df_reporte = df_reporte.drop(columns=["Tipo"])

def excel_formula_es_to_en(formula):

    reemplazos = {
        "SI.ERROR": "IFERROR",
        "SI(": "IF(",
        "SUMA": "SUM",
        "PROMEDIO": "AVERAGE",
        "COINCIDIR": "MATCH",
        "INDICE": "INDEX",
        "VERDADERO": "TRUE",
        "FIN.MES": "EOMONTH",
        "DIA": "DAY",
        "HOY": "TODAY"
    }

    for es, en in reemplazos.items():
        formula = formula.replace(es, en)

    formula = formula.replace(";", ",")

    return formula


hojas = {
    "Reporte": df_reporte,
    "Setting-Remplazos": df_setting,
    "Consumo Historico 24M": df_history,
    "LeadTime": dfResumenLeadTimes
}

if Linea_Negocio == "HIDRAULI. COMPONENTE":
    hojas["Pesca"] = df_pesca

if Linea_Negocio == "EQUIPOS TRANS. MATER":
    hojas["Detalle"] = detalleUnicon
    
if Linea_Negocio in ["LUBRICACION MINERIA","HIDRAULI. COMPONENTE","EQUIPOS TRANS. MATER"]:
    hojas["Simulaciones"] = df_final


output = (base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/ f"Reporte_Abastecimiento_{Linea_Negocio}_{mes_anio}_prueba.xlsx")

#archivo_excel = f"Reporte_Abastecimiento_{Linea_Negocio}_{mes_anio}.xlsx"


with pd.ExcelWriter(output, engine='xlsxwriter') as writer:

    workbook = writer.book

    # =====================================================
    # IMPORTANTE
    # =====================================================

    workbook.set_calc_mode('auto')

    # =====================================================
    # FORMATOS
    # =====================================================

    formato_encabezado = workbook.add_format({
        'bold': True,
        'bg_color': '#D9F2FF',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    formato_celdas = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    formato_fecha = workbook.add_format({
        'num_format': 'dd/mm/yyyy',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    formato_numerico = workbook.add_format({
        'num_format': '0.00',
        'border': 1,
        'align': 'right',
        'valign': 'vcenter'
    })

    # =====================================================
    # LOOP HOJAS
    # =====================================================

    for nombre_hoja, df in hojas.items():

        # =================================================
        # HOJAS NORMALES
        # =================================================

        if nombre_hoja != "Simulaciones":

            df_redondeado = df.copy()

            for col in df_redondeado.select_dtypes(include=[np.number]).columns:
                df_redondeado[col] = df_redondeado[col].round(1)

            df_redondeado.to_excel(
                writer,
                sheet_name=nombre_hoja,
                index=False
            )

            worksheet = writer.sheets[nombre_hoja]

            # =============================================
            # HEADERS
            # =============================================

            for col_num, value in enumerate(df_redondeado.columns.values):

                worksheet.write(
                    0,
                    col_num,
                    value,
                    formato_encabezado
                )

            # =============================================
            # ANCHOS
            # =============================================

            for col_num, col_name in enumerate(df_redondeado.columns):

                try:

                    max_len = max(
                        df_redondeado[col_name]
                        .astype(str)
                        .map(len)
                        .max(),
                        len(col_name)
                    ) + 2

                except:

                    max_len = len(col_name) + 2

                # NUMERICOS
                if np.issubdtype(
                    df_redondeado[col_name].dtype,
                    np.number
                ):

                    worksheet.set_column(
                        col_num,
                        col_num,
                        max_len,
                        formato_numerico
                    )

                # FECHAS
                elif np.issubdtype(
                    df_redondeado[col_name].dtype,
                    np.datetime64
                ):

                    worksheet.set_column(
                        col_num,
                        col_num,
                        max_len,
                        formato_fecha
                    )

                # TEXTO
                else:

                    worksheet.set_column(
                        col_num,
                        col_num,
                        max_len,
                        formato_celdas
                    )

        # =================================================
        # HOJA SIMULACIONES
        # =================================================

        else:

            worksheet = workbook.add_worksheet(nombre_hoja)

            writer.sheets[nombre_hoja] = worksheet

            # =============================================
            # ANCHOS
            # =============================================

            worksheet.set_column(0, 0, 24)
            worksheet.set_column(1, 300, 14)


            # =============================================
            # TOTAL ROWS
            # =============================================

            total_rows = len(df)

            # =============================================
            # WRITE MANUAL
            # =============================================

            for row_num in range(total_rows):

                valor = df.iloc[row_num, 0]

                # =========================================
                # HEADERS
                # =========================================

                if valor in [
                    "Codigo_SAP",
                    "Clasificación",
                    "Costo S/.",
                    "Stock Disponible"
                ]:

                    worksheet.write(
                        row_num,
                        0,
                        valor,
                        formato_encabezado
                    )

                    if len(df.columns) > 1:

                        val = df.iloc[row_num, 1]

                        worksheet.write(
                            row_num,
                            1,
                            val,
                            formato_celdas
                        )

                # =========================================
                # FECHAS
                # =========================================

                elif valor == "":

                    if row_num + 1 < total_rows:

                        siguiente = df.iloc[row_num + 1, 0]

                        if siguiente == "SI":

                            for col_num in range(1, len(df.columns)):

                                val = df.iloc[row_num, col_num]

                                if pd.notna(val) and val != "":

                                    worksheet.write(
                                        row_num,
                                        col_num,
                                        val,
                                        formato_fecha
                                    )

                # =========================================
                # TABLA
                # =========================================

                elif valor in [
                    "SI",
                    "IN",
                    "OUT",
                    "SF",
                    "COB",
                    "ACUM OUT"
                ]:

                    for col_num in range(len(df.columns)):

                        val = df.iloc[row_num, col_num]
                    
                        # =================================
                        # FORMULAS
                        # =================================
                    
                        if isinstance(val, str) and val.startswith("="):
                    
                            val = excel_formula_es_to_en(val)
                    
                            worksheet.write_array_formula(
                                row_num,
                                col_num,
                                row_num,
                                col_num,
                                "{" + val + "}",
                                formato_numerico
                            )
                    
                        # =================================
                        # NUMEROS
                        # =================================
                    
                        elif isinstance(val, (int, float)):
                    
                            worksheet.write_number(
                                row_num,
                                col_num,
                                val,
                                formato_numerico
                            )
                    
                        # =================================
                        # TEXTO
                        # =================================
                    
                        else:
                    
                            worksheet.write(
                                row_num,
                                col_num,
                                val,
                                formato_celdas
                            )

    print("Archivo Excel creado con éxito en:", output)