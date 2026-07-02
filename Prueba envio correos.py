import pandas as pd
from datetime import datetime
import numpy as np
import os
from decimal import Decimal
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

#Ejecucion de reportes de Hidraulica
data = {
    "Numero": [101, 102, 103, 104, 105, 106, 107, 108],
    "Estado": ["Activo", "Inactivo", "Pendiente", "Activo", 
               "Suspendido", "Activo", "Inactivo", "Pendiente"]
}
df_resultado_ltvar = pd.DataFrame(data)

#Declaracion de funcion de envio de correo

def enviar_correo(df_resultado_ltvar, nombre_linea ,nombre_linea_descripcion):

    #Envio de correo
    df_resultado_ltvar = df_resultado_ltvar.replace([np.nan, np.inf, -np.inf], "")
    # Guardar archivo adjunto (Le damos formato al excel adjnto)
    archivo_excel = f'Reporte de abastecimiento {nombre_linea}.xlsx'
    #df_resultado_ltvar.to_excel(archivo_excel, index=False)
    # Crear el Excel con formato
    with pd.ExcelWriter(archivo_excel, engine='xlsxwriter') as writer:
        df_resultado_ltvar.to_excel(writer, index=False, sheet_name=f'Reporte {nombre_linea}')
        workbook = writer.book
        worksheet = writer.sheets[f'Reporte {nombre_linea}']

        # Formatos
        formato_encabezado = workbook.add_format({
            'bold': True,
            'bg_color': '#D9F2FF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        formato_celdas = workbook.add_format({
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })

        formato_fecha = workbook.add_format({
            'num_format': 'dd/mm/yyyy',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        # Aplicar formato a encabezados y ajustar ancho
        for col_num, col_name in enumerate(df_resultado_ltvar.columns):
            worksheet.write(0, col_num, col_name, formato_encabezado)
            max_len = max(
                df_resultado_ltvar[col_name].astype(str).map(len).max(),
                len(col_name)
            ) + 2
            worksheet.set_column(col_num, col_num, max_len)

        # Aplicar formato a celdas con tipo adecuado
        for fila in range(1, len(df_resultado_ltvar) + 1):
            for col_num, col_name in enumerate(df_resultado_ltvar.columns):
                valor = df_resultado_ltvar.iloc[fila - 1, col_num]
                       
                # Si la columna es tipo datetime, aplicar formato de fecha
                if pd.api.types.is_datetime64_any_dtype(df_resultado_ltvar[col_name]):
                    worksheet.write_datetime(fila, col_num, valor,formato_fecha)
                else:
                    worksheet.write(fila, col_num, valor, formato_celdas)

    # Obtener la fecha actual
    hoy = datetime.today()
    # Obtener el número de semana ISO
    numero_semana = hoy.isocalendar()[1]
    numero_semana=str(numero_semana)

    # ------------ CONFIGURAR CORREO ------------
    print("Se ejecuto el envio correo")
    remitente = "rhuamani@marco.com.pe"
    clave = "gSDvbs%#$"
    #destinatario = "VentasGHTP@marco.com.pe"
    #destinatario = ["planeamiento.inventarios@marco.com.pe","aprado@marco.com.pe","acampos@marco.com.pe","rhuamani@marco.com.pe"]
    destinatario = ["rhuamani@marco.com.pe","rhuamani@marco.com.pe"]
    asunto = "SEM "+ numero_semana + ": Reporte PRUEBA "+nombre_linea
    # Copia a otras personas
    #cc = ["greynac@marco.com.pe","planeamiento.inventarios@marco.com.pe","aprado@marco.com.pe"]  # Ajusta los correos según tu necesidad
    #cc = ["planeamiento.inventarios@marco.com.pe","aprado@marco.com.pe"]  # Ajusta los correos según tu necesidad

    # Crear el mensaje
    mensaje = MIMEMultipart()
    mensaje["From"] = remitente
    #mensaje["To"] = destinatario
    mensaje["To"] = ", ".join(destinatario)
    mensaje["Subject"] = asunto
    #mensaje["Cc"] = ", ".join(cc)


    # Cuerpo HTML con la tabla
    cuerpo_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        body {{
          font-family: Arial, sans-serif;
          line-height: 1.5;
          color: #333333;
        }}
        table {{
          border-collapse: collapse;
          width: 100%;
        }}
        th, td {{
          border: 1px solid #dddddd;
          padding: 8px;
          text-align: left;
        }}
        th {{
          background-color: #e6f7ff;
        }}
      </style>
    </head>
    <body>
      <p>Estimados(as),</p>
    
      <p>
        Adjunto envío el <strong>reporte de abastecimiento de {nombre_linea_descripcion} </strong>, correspondiente al periodo más reciente.
        En este documento podrán revisar el estado actual de abastecimiento, niveles de inventario y cualquier posible desviación detectada.
      </p>
    
      <p>
        Saludos cordiales,<br>
        <strong>Planeamiento de Inventarios</strong>
      </p>
    </body>
    </html>
    """
    # {tabla_html} para insertar tabla en cuerpo de correo
    mensaje.attach(MIMEText(cuerpo_html, "html"))

    # Adjuntar archivo Excel
    with open(archivo_excel, "rb") as adjunto:
        parte_adjunto = MIMEApplication(adjunto.read(), _subtype="xlsx")
        parte_adjunto.add_header('Content-Disposition', 'attachment', filename=archivo_excel)
        mensaje.attach(parte_adjunto)

    # ------------ ENVIAR CORREO ------------
    try:
        servidor = smtplib.SMTP("smtp.office365.com", 587)
        servidor.starttls()
        servidor.login(remitente, clave)
        servidor.sendmail(remitente, destinatario, mensaje.as_string())
        servidor.quit()
        print("Correo enviado correctamente.")
    except Exception as e:
        print(f" Error al enviar el correo: {e}")
    # No se retorna nada
    return

#Ejecucion de correos
#Hidraulica componentes
enviar_correo(df_resultado_ltvar, "FRIO" , "FRIO")

