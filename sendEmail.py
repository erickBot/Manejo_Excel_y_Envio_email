from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import openpyxl
import time
import os
from os import remove
from datetime import datetime, date, time, timedelta
import calendar

mylist = []

def run():
    enviarEmail = False

    name_file = "Lista_de_Sensores_HGAS.xlsx"
    file = openpyxl.load_workbook(name_file)
    sheet1 = file.get_sheet_by_name('LEL')
    sheet2 = file.get_sheet_by_name('H2S')

    porVencer_lel = sheet1['I24'].value
    vencidos_lel  = sheet1['I25'].value
    vigentes_lel  = sheet1['I26'].value
    porVencer_h2s = sheet2['I24'].value
    vencidos_h2s  = sheet2['I25'].value
    vigentes_h2s  = sheet2['I26'].value

    mylist.append(porVencer_lel)
    mylist.append(vencidos_lel)
    mylist.append(vigentes_lel)
    mylist.append(porVencer_h2s)
    mylist.append(vencidos_h2s)
    mylist.append(vigentes_h2s)

    if porVencer_lel > 0 or vencidos_lel > 0:
        enviarEmail = True

    if porVencer_h2s > 0 or vencidos_h2s > 0:
        enviarEmail = True

    if (enviarEmail):
        procesaDestinatarios()

    file.close()

def procesaDestinatarios():

    destinatarios = {
        "user1": "erickpasache0@gmail.com",
        "user2": "epasache_28@hotmail.com"
    }

    for i in destinatarios:
        sendEmail(destinatarios[i])

def sendEmail(destinatarios):
    #crea la instancia del objeto de mensaje
    msg = MIMEMultipart()
    message = "Hola a todos \n\nSe envia lista actualizada de sensores HGAS. Hay " + str(mylist[0]) + " sensores LEL por vencer, " + str(mylist[1]) + " sensores LEL vencidos, " + str(mylist[2]) + " sensores LEL vigentes,  " + str(mylist[3]) + " sensores H2S por vencer, " + str(mylist[4]) + " sensores H2S vencidos, y " + str(mylist[5])+ " sensores H2S vigentes.\n\nSoy un Bot."
    ruta_adjunto = "Lista_de_Sensores_HGAS.xlsx"
    nombre_adjunto = "Lista_de_Sensores_HGAS.xlsx"
    #configura los parametros del mensaje
    password = "99e12438cf"
    msg['From'] ="soyunbot2817@gmail.com"
    msg['To'] = destinatarios
    msg['Subject'] = "Lista de Sensores HGAS por expirar en Talara"
    #agrega el cuerpo del mensaje
    msg.attach(MIMEText(message, 'plain'))
    # Abrimos el archivo que vamos a adjuntar
    archivo_adjunto = open(ruta_adjunto, 'rb')
    # Creamos un objeto MIME base
    adjunto_MIME = MIMEBase('application', 'octet-stream')
    # Y le cargamos el archivo adjunto
    adjunto_MIME.set_payload((archivo_adjunto).read())
    # Codificamos el objeto en BASE64
    encoders.encode_base64(adjunto_MIME)
    # Agregamos una cabecera al objeto
    adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_adjunto)
    # Y finalmente lo agregamos al mensaje
    msg.attach(adjunto_MIME)
    #crear servidor
    server = smtplib.SMTP('smtp.gmail.com: 587')
    server.starttls()
    #Ingresa credenciales para enviar email
    server.login(msg['From'], password)
    #envia el mensaje al servidor
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.quit()
    print ("Envio de email exitoso a %s:" % (msg['To']))

if __name__ == "__main__":
    run()
