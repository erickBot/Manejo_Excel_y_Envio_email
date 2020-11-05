from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib

def run():
    destinatarios = {
        "user1": "erick.pasache@pason.com",
        "user2": "epasache_28@hotmail.com",
        "user3": "johnnel.panduro@pason.com"
    }

    for i in ["user1", "user2", "user3"]:
        sendEmail(destinatarios[i])

def sendEmail(destinatarios):
    #crea la instancia del objeto de mensaje
    msg = MIMEMultipart()
    message = "Test send email XD"
    ruta_adjunto = "inventarioJunio2020.xlsx"
    nombre_adjunto = "ArchivoAdjunto"
    #configura los parametros del mensaje
    password = "epas2817epas"
    msg['From'] ="erickpasache0@gmail.com"
    msg['To'] = destinatarios
    msg['Subject'] = "Subscription"
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