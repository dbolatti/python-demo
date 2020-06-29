import smtplib, ssl, csv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

sender_email = "no-reply@domain.com"
password = input("Type your password and press enter:")

# Create the plain-text and HTML version of your message
text = """\
Estimado: {nombres} {apellido} 
Su nueva cuenta de Google For Education es {cuenta} 
Clave temporal: ********* 
Acceder vía: https://gsuite.google.com/dashboard
"""

html = """\
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html xmlns=\"http://www.w3.org/1999/xhtml\"> <head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <title>UTN FRRe</title><meta name="viewport" content="width=device-width, initial-scale=1.0"/>
        </head>
        <body style="margin: 0; padding: 0;">
            <table align="center" border="0" cellpadding="0" cellspacing="0" width="600">
            <tr>
                <td style="font-size: 0; line-height: 0;" height="10">&nbsp;</td>
            </tr>
                <td><h1>UTN FRRe- Google para Centros Educativos</h1></td>
            </tr>
            <tr>
                <td style="font-size: 0; line-height: 0;" height="10">&nbsp;</td>
            </tr>
                <td>
                    <table border="1" cellpadding="0" cellspacing="0" width="100%">
                        <tr>
                            <td><h2>ALTA de cuenta</h2></td>
                        </tr>
                        <tr>
                            <td style="padding: 20px 0 30px 0;"><b>Estimada/o: </b> {nombres} {apellido} su nueva cuenta de Google For Education es: {cuenta}</td>
                        </tr>
                        <tr>
                            <td>Clave temporal: ********* (se solicita el cambio en el primer inicio de sesion) <br> URL de acceso: https://gsuite.google.com/dashboard</td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td style="font-size: 0; line-height: 0;" height="10">&nbsp;</td>
            </tr>
                <td><table border="0" cellpadding="0" cellspacing="0">
                        <tr>
                            <td><a href="http://www.domain.com"><img src="http://www.domain.com/www.png" alt="img" border="0" /></a></td>
                            <td style="font-size: 0; line-height: 0;" width="20">&nbsp;</td>
                            <td></td>
                        </tr>
                    </table>
                </td>
            </tr>
            </table>
        </body>
    </html>
"""

# Create secure connection with server and send email
context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(sender_email, password)
    with open("C://Descargas/A_GSuite/Usuarios-GSuite-Para_Diego.csv") as file:
        reader = csv.reader(file)
        next(reader)  # Skip header row
        for linea in reader:
            message = MIMEMultipart("alternative")
            message["Subject"] = "Nueva cuenta de Google For Education."
            message["From"] = sender_email
            message["To"] = linea[4]

            html.format(nombres=linea[0],apellido=linea[1],cuenta=linea[2])
            text.format(nombres=linea[0],apellido=linea[1],cuenta=linea[2])

            # Turn these into plain/html MIMEText objects
            part1 = MIMEText(text, "plain")
            part2 = MIMEText(html, "html")

            # Add HTML/plain-text parts to MIMEMultipart message
            # The email client will try to render the last part first
            message.attach(part1)
            message.attach(part2)

            print('Enviando a: ', linea)
            #print('Mensaje: ', message.as_string().format(nombres=linea[0],apellido=linea[1],cuenta=linea[2]))
            server.sendmail(sender_email, linea[4], message.as_string().format(nombres=linea[0],apellido=linea[1],cuenta=linea[2]))

file.close()
print('************************ PROCESO TERMINADO ************************')
