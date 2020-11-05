from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import openpyxl
import time

def run():
    cont = 0
    name_file = "Lista_Sensores_HGAS.xlsx"
    file = openpyxl.load_workbook(name_file, data_only=True)
    sheet = file.get_sheet_by_name('LEL')
    cell_all = sheet['I4':'I19']
    for row in cell_all:
        for cell in row:
            if cell.value < 30 :
                cont = cont + 1

    if cont > 0 :
        print(cont)
        file.save(name_file)
        procesaDestinatarios()

def procesaDestinatarios():
    
    destinatarios = {
        "user1": "erick.pasache@pason.com",
        "user2": "epasache_28@hotmail.com"
    }

    for i in ["user1", "user2"]:
        sendEmail(destinatarios[i])

def sendEmail(destinatarios):
    #crea la instancia del objeto de mensaje
    msg = MIMEMultipart()
    message = "Hola a todos \n\nSe envia lista actualizada de sensores HGAS que estan a punto de expirar su fecha de calibracion.\n\nNo responder soy un Bot."
    ruta_adjunto = "Lista_Sensores_HGAS.xlsx"
    nombre_adjunto = "Lista_Sensores_HGAS.xlsx"
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