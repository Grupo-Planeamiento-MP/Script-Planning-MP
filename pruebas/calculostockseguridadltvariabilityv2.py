import pandas as pd
import numpy as np

def calcular_stock_seguridad_ltvar(
    df_resultado_ltvar,
    df_consumos,
    DfLeadTimes,
    reference_period,
    num_simulaciones=500
):

    lt_por_item = dict(tuple(DfLeadTimes.groupby('ItemCode')))
    consumo_por_item = dict(tuple(df_consumos.groupby('SAP')))
    resultados = []
    

    for row in df_resultado_ltvar.itertuples(index=True):
        print(row.Codigo_SAP)
        codigo_sap = row.Codigo_SAP  
        df_lt = lt_por_item.get(codigo_sap)
        df_cons = consumo_por_item.get(codigo_sap)
        
        
        obs=df_lt['Tipo LT'].iloc[0]
        tipo = df_lt['Tipo de Compra'].iloc[0]
        
        lead_times = df_lt['LT_final'].dropna().values
        catLI = "IMPORTADO" if tipo == 'Importación' else "LOCAL"        
        promedio_lt = np.mean(lead_times)
        ###
        cv_lt = np.std(lead_times, ddof=1) / promedio_lt if promedio_lt > 0 else np.nan
        ###
        # Condicion si no existe LT
        if (
                df_lt is None
                or df_lt.empty
                or df_lt['LT_final'].sum() == 0
            ):
            df_lt = pd.DataFrame()
            obs="sin LT"
            resultados.append((np.nan, np.nan, np.nan, np.nan, np.nan, 0, "LOCAL",obs,np.nan,np.nan))
            continue
        
        # Condicion si no existe Consumos
        if (
                df_cons is None
                or df_cons.empty
                or df_cons['Consumo Total'].sum() == 0
            ):
            resultados.append(("Sin consumos", np.nan, np.nan, np.nan, np.nan, promedio_lt, catLI,obs,np.nan,np.nan))
            continue
        
       

        consumo_rango = np.zeros(12)
        consumo_rango[:len(df_cons)] = df_cons['Consumo Total'].values
        
        media_cons = np.mean(consumo_rango)
        std_cons = np.std(consumo_rango, ddof=1)
        cv_cons = std_cons / media_cons if media_cons > 0 else np.nan 
        
        # Simulacion
        lt_sim = np.random.choice(lead_times, size=num_simulaciones)
        entero = np.floor(lt_sim).astype(int)
        frac = lt_sim - entero

        suma_entero = np.array([
            np.random.choice(consumo_rango, size=n, replace=True).sum()
            for n in entero
        ])

        muestra_frac = np.random.choice(consumo_rango, size=num_simulaciones) * frac
        simulaciones = suma_entero + muestra_frac # 

        p95, p98 = np.percentile(simulaciones, [95, 98])
        prom = simulaciones.mean()

        resultados.append((max(p95 - prom, 0),max(p98 - prom, 0), prom, p95, p98, promedio_lt, catLI,obs, round(cv_lt, 1), round(cv_cons, 1)))#cv_lt,cv_cons

    # Resultados
    cols = ['SS_95_ltvar','SS_98_ltvar','Prom_Cons_LT_ltvar','P95_ltvar','P98_ltvar','Promedio_LT','L/I',"Observacion LT", "CV_LT","CV_Consumo"  ]#"CV_LT","CV_Consumo"
    df_resultado_ltvar[cols] = pd.DataFrame(resultados, columns=cols)

    return df_resultado_ltvar