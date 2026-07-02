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

dfkardexorigen = pd.read_csv(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Archivos_Compartidos"/"Querys automatizados"/"dfkardexorigen.txt", sep="|", encoding="utf-8")
dfkardexorigen['Fecha de contabilización'] = pd.to_datetime(dfkardexorigen['Fecha de contabilización'])

# Eliminar columnas no deseadas
dfkardexorigen = dfkardexorigen.drop(columns=["Número de operación", "Status", "Clave de documento creada"])
dfkardexorigen["Cantidad de entrada"] = dfkardexorigen["Cantidad de entrada"].fillna(0)
dfkardexorigen["Cantidad de salida"] = dfkardexorigen["Cantidad de salida"].fillna(0)
# Cambiar tipos de columnas
dfkardexorigen["Cargado a"] = dfkardexorigen["Cargado a"].astype(str)
dfkardexorigen["Codigo GET"] = dfkardexorigen["Codigo GET"].astype(str)

# Agregar columna personalizada "Codigo_Conc"
dfkardexorigen["Codigo_Conc"] = "MP" + dfkardexorigen["Número de artículo"]

# Cambiar tipo de columna "Codigo_Conc" y "Código de almacén"
dfkardexorigen["Codigo_Conc"] = dfkardexorigen["Codigo_Conc"].astype(str)
dfkardexorigen["Código de almacén"] = dfkardexorigen["Código de almacén"].astype(str)


# Unir con Maestro Almacenes (2)
df_maestro_almacenes = pd.read_table(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"Almacenes.txt", encoding='utf-8', sep='\t', quotechar='"', low_memory=False)
dfkardexorigen = dfkardexorigen.merge(df_maestro_almacenes[['Código de almacén', 'Nombre de almacén', 'TIPO']], on="Código de almacén", how="left")

# Unir con Maestro Articulos
df_maestro_articulos = pd.read_table(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"Maestro Art.txt", encoding='utf-8', sep='\t', quotechar='"', low_memory=False)
dfkardexorigen = dfkardexorigen.merge(df_maestro_articulos[['Codigo Concateando', 'Codigo Unico']], left_on="Codigo_Conc", right_on="Codigo Concateando", how="left")

# Agregar columna personalizada "Control"
dfkardexorigen["Control"] = np.where(dfkardexorigen["Codigo Unico"].isnull(), dfkardexorigen["Codigo_Conc"], dfkardexorigen["Codigo Unico"])

# Renombrar columnas
dfkardexorigen = dfkardexorigen.rename(columns={"Control": "Maestro Articulos.Codigo Unico"})

# Unir con Maestro_Soc
df_maestro_soc = pd.read_table(base/"MARCO PERUANA SA"/"Planeamiento de Inventarios - Documents"/"Proyectos"/"Python"/"Pruebas Linux"/"MaestroSoc.txt", encoding='utf-8', sep='\t', quotechar='"', low_memory=False)
dfkardexorigen = dfkardexorigen.merge(df_maestro_soc[['Sublineas', 'Unidad de Negocios', 'Linea de Negocio']], left_on="Nombre de grupo", right_on="Sublineas", how="left")

# Filtrar filas no deseadas
dfkardexorigen = dfkardexorigen [(dfkardexorigen ["Código de almacén"] != "397") & 
        (~dfkardexorigen ["Tipo"].isin(["Devolucion por Compra", "Factura de  Proveedores", "NC de Proveedores", "Precio de Entrega", "Transferencia por Inventario"]))]

# Agregar columna condicional "Tipo de articulo"
dfkardexorigen["Tipo de articulo"] = np.where(dfkardexorigen["Número de artículo"].str.startswith("A"), "Articulo", "Otros")

# Filtrar filas con comentarios no deseados
comentarios_excluir = ["cold import - error en moneda", "REGULARIZACION DE FECHAS DE LA OF 2019000191", "SALDO", "Saldo", "Toma de inventario 2019-12"]
dfkardexorigen = dfkardexorigen[~dfkardexorigen["Comentarios"].str.contains('|'.join(comentarios_excluir), case=False, na=False)]

# Reemplazar errores y valores nulos
dfkardexorigen["Comentarios.1"] = dfkardexorigen["Comentarios.1"].fillna("Sindato").replace({"": "Sindato"})

# Filtrar filas con más criterios
dfkardexorigen = dfkardexorigen[(~dfkardexorigen["Comentarios.1"].str.contains("AJUSTE|INV SUF 2021|FALTANTES 2020|FALTANTE INV 20-2|INV FALTANTES 20|INV FALTANTES 20-2|AUDITORIA|quema", case=False)) & 
        (~dfkardexorigen["Referencia base"].isin([2144000508, 223403346, 223300254, 224800113]))]

#----------------------------------------------------------------------------------------------------------------------------

##############################################################################################################################
##############################################################################################################################

"""
CONDICIONES NUMERADAS
para categorizar operaciones de salidas de stock validas para el calculo de la rotacion y analisis de reposicion de stock
contemplando aquellas operaciones que las netean como documentos de devoluciones, notas de credito, recibos, etc


#1. Sin son operaciones de venta o devolucion por venta se categorizan como "Venta Directa"
#2. Si es una emision para produccion y el codigo padre de la OF comienza en S o MP se categoriza como "Servicio"
#3. Si es una emision para produccion y el codido padre de la OF no es S o MP se categoriza como "Consumo" dado que es un setting
#4. Si es un recibo de produccion y el codigo padre es S o MP y el articulo devuelto es diferente al padre entonces se categoriza como "Servicio"
    .Este caso es una devolucion de materiales no usados en la OF por lo tanto deben netear la emision de produccion
#5. Es el caso 4 pero cuando el articulo devuelto es igual al padre. En este caso lo que esta ingresando al stock es el resultado de la fabricacion
    o servicio y por lo tanto debe quedar en blanco, haciendo referencia a una operacion valida de entrada al stock
#6. Si es un recibo de produccion y el codigo padre no es ni S o MP, entonces se trata de un setting. La operacion de transformacion
    produce codigos nuevos y es considerada una operacion valida de entrada al stock. Se categoriza en None
#7 y #8. Las operaciones de salida de mercancias y otros se categorizan en "Consumo". Aunque esto no es del todo preciso
         no impacta en el analisis de demanda ya que aqui solo se consideran operaciones de Venta Directa y Servicio.


"""


# Condiciones
cond1 = dfkardexorigen["Comentarios"].isin(["Emisión para producción", "Issue for Production"])
cond1_servicio = cond1 & dfkardexorigen["ItemPrincipalOF"].str.startswith(("S", "MP"), na=False)
cond1_consumo = cond1 & ~dfkardexorigen["ItemPrincipalOF"].str.startswith(("S", "MP"), na=False)

cond2 = dfkardexorigen["Comentarios"].isin(["Recibo de producción", "Receipt from Production"])
cond2_item_start = dfkardexorigen["ItemPrincipalOFReciboProd"].str.startswith(("S", "MP"), na=False)
cond2_equal = cond2 & cond2_item_start & (dfkardexorigen["Número de artículo"] == dfkardexorigen["ItemPrincipalOFReciboProd"])
cond2_diff = cond2 & cond2_item_start & (dfkardexorigen["Número de artículo"] != dfkardexorigen["ItemPrincipalOFReciboProd"])
cond2_else = cond2 & ~cond2_item_start

# Definimos condiciones y resultados
conditions = [
    dfkardexorigen["Clase de operación"].isin([13, 14, 15, 16]), #1
    cond1_servicio, #2
    cond1_consumo, #3
    cond2_diff, #4
    cond2_equal, #5
    cond2_else, #6
    dfkardexorigen["Comentarios"] == "Salida de mercancías", #7
    dfkardexorigen["Comentarios"].str.startswith("Regularizaion", na=False), #8
]

choices = [
    "Venta Directa",   # Clase de operación
    "Servicio",        # Emisión para producción / Issue for Production y comienza con S o MP
    "Consumo",         # Emisión para producción / Issue for Production y NO comienza con S o MP
    "Servicio",        # Recibo de producción y ≠ ItemPrincipalOFReciboProd pero sí comienza con S o MP
    None,                # Recibo de producción y = ItemPrincipalOFReciboProd y sí comienza con S o MP
    None,                # Recibo de producción pero NO comienza con S o MP
    "Consumo",         # Salida de mercancías
    "Consumo",         # Regularizaion
]

# Asignar columna personalizada
dfkardexorigen["Personalizado"] = np.select(conditions, choices, default=None)

#----------------------------------------------------------------------------------------------------------------------------


dfkardexorigen["Cantidad de entrada"] = (
    dfkardexorigen["Cantidad de entrada"]
    .astype(str)  # Convertir todos los valores a string
    .str.replace(',', '', regex=False)  # Remover separadores de miles (coma)
    .replace(r'^\s*$', None, regex=True)  # Reemplazar valores vacíos o espacios con None (para que sean NaN)
    .astype(float)  # Convertir finalmente a float
)


dfkardexorigen["Cantidad de salida"] = (
    dfkardexorigen["Cantidad de salida"]
    .astype(str)  # Convertir todos los valores a string
    .str.replace(',', '', regex=False)  # Remover separadores de miles (coma)
    .replace(r'^\s*$', None, regex=True)  # Reemplazar valores vacíos o espacios con None (para que sean NaN)
    .astype(float)  # Convertir finalmente a float
)

# # Agregar columna "Cantidad unificada"
# dfkardexorigen["Cantidad unificada"] = np.where(dfkardexorigen["Tipo"].isin(["Devolucion de Venta", "NC de Clientes", "Recibo de producción", "Receipt from Production"]),
#                                     -dfkardexorigen["Cantidad de entrada"],
#                                     dfkardexorigen["Cantidad de entrada"] + dfkardexorigen["Cantidad de salida"])

# Definir las condiciones
conditions = [
    dfkardexorigen["Tipo"].isin(["Devolucion de Venta", "NC de Clientes"]),
    dfkardexorigen["Comentarios"].isin(["Recibo de producción", "Receipt from Production"]),
]

# Definir los valores correspondientes a las condiciones
choices = [
    -dfkardexorigen["Cantidad de entrada"],  # Para los casos de 'Tipo'
    -dfkardexorigen["Cantidad de entrada"],  # Para los casos de 'Comentarios'
]

# Si no se cumple ninguna condición, calcular 'Cantidad de entrada' + 'Cantidad de salida'
default = dfkardexorigen["Cantidad de entrada"] + dfkardexorigen["Cantidad de salida"]

# Aplicar np.select para asignar los valores
dfkardexorigen["Cantidad unificada"] = np.select(conditions, choices, default=default)





# Convertir a tipo numérico
dfkardexorigen["Cantidad unificada"] = pd.to_numeric(dfkardexorigen["Cantidad unificada"], errors='coerce')




dfkardexorigen["Valor de transacción"] = (
    dfkardexorigen["Valor de transacción"]
    .astype(str)  # Convertir todos los valores a string
    .str.replace(',', '', regex=False)  # Remover separadores de miles (coma)
    .replace(r'^\s*$', None, regex=True)  # Reemplazar valores vacíos o espacios con None (para que sean NaN)
    .astype(float)  # Convertir finalmente a float
)
    

    
# dfkardexorigen["Ingreso bruto línea"] = (
#     dfkardexorigen["Ingreso bruto línea"]
#     .astype(str)  # Convertir todos los valores a string
#     .str.replace(',', '', regex=False)  # Remover separadores de miles (coma)
#     .replace(r'^\s*$', None, regex=True)  # Reemplazar valores vacíos o espacios con None (para que sean NaN)
#     .astype(float)  # Convertir finalmente a float
# ) 
    

# dfkardexorigen["GrssProfit Devol"] = (
#     dfkardexorigen["GrssProfit Devol"]
#     .astype(str)  # Convertir todos los valores a string
#     .str.replace(',', '', regex=False)  # Remover separadores de miles (coma)
#     .replace(r'^\s*$', None, regex=True)  # Reemplazar valores vacíos o espacios con None (para que sean NaN)
#     .astype(float)  # Convertir finalmente a float
# ) 


# Agregar columna "Valor unificado"
dfkardexorigen["Valor unificado"] = -dfkardexorigen["Valor de transacción"]

# Convertir a tipo numérico
dfkardexorigen["Valor unificado"] = pd.to_numeric(dfkardexorigen["Valor unificado"], errors='coerce')


