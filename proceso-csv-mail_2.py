import smtplib, ssl, csv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


sender_email = "no-reply@gfe.frre.utn.edu.ar"
password = input("Ingrese contraseña [no-reply@gfe.frre.utn.edu.ar]:")

# Create the plain-text and HTML version of your message
text = """\
Estimado: {nombres} {apellido} 
Su nueva cuenta de Goole For Education es {cuenta} 
Clave temporal: Temporal123 
Acceder vía: https://gsuite.google.com/dashboard
Otros enlaces:
- Mail: http://mail.google.com/a/gfe.frre.utn.edu.ar
- Classroom: https://classroom.google.com/a/gfe.frre.utn.edu.ar
- Meet: https://meet.google.com
- Capacitacion en G-Suite: https://teachercenter.withgoogle.com/training
"""

html = """\
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html xmlns=\"http://www.w3.org/1999/xhtml\"> <head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <title>UTN FRRe</title><meta name="viewport" content="width=device-width, initial-scale=1.0"/>
        </head>
        <body style="margin: 0; padding: 0;">
            <table align="center" border="0" cellpadding="1" cellspacing="0" width="600">
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
                            <td style="padding: 20px 0 30px 0;"><b>Estimada/o: </b> {nombres} {apellido} su nueva cuenta de Goole para Centros Educativos es: {cuenta}</td>
                        </tr>
                        <tr>
                            <td>Clave temporal: Temporal123 (se solicita el cambio en el primer inicio de sesion) <br> URL de acceso: <a href="https://gsuite.google.com/dashboard">https://gsuite.google.com/dashboard</a></td>
                        </tr>
                        <tr>
                            <td><h2>Otros enlaces</h2></td>
                        </tr>
                        <tr>
                            <td>Mail: <a href="http://mail.google.com/a/gfe.frre.utn.edu.ar">http://mail.google.com/a/gfe.frre.utn.edu.ar</a></td>
                        </tr>
                        <tr>
                            <td>Classroom: <a href="https://classroom.google.com/a/gfe.frre.utn.edu.ar">https://classroom.google.com/a/gfe.frre.utn.edu.ar</a></td>
                        </tr>
                        <tr>
                            <td>Meet: <a href="https://meet.google.com">https://meet.google.com</a></td>
                        </tr>                        
                        <tr>
                            <td>Capacitacion en G-Suite: <a href="https://teachercenter.withgoogle.com/training">https://teachercenter.withgoogle.com/training</a></td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td style="font-size: 0; line-height: 0;" height="10">&nbsp;</td>
            </tr>
                <td><table border="0" cellpadding="0" cellspacing="0">
                        <tr>
                            <td><a href="http://www.frre.utn.edu.ar"><img src="http://www.frre.utn.edu.ar/public/themes/utn/images/UTN_FRRE-www.png" alt="FRRe" border="0" /></a></td>
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
    with open("A_GSuite/Usuarios-GSuite-Para_PowerShell.csv") as file:
        #Creo un diccionario con los datos del csv
        reader = csv.DictReader(file)

        #En este caso no se usa, porque se utiliza DICCIONARIO
        #next(reader)  # Skip header row

        for linea in reader:
            message = MIMEMultipart("alternative")
            message["Subject"] = "Nueva cuenta de Google For Education."
            message["From"] = sender_email
            message["To"] = linea['mail']

            html.format(nombres=linea['nombres'],apellido=linea['apellido'],cuenta=linea['cuenta'])
            text.format(nombres=linea['nombres'],apellido=linea['apellido'],cuenta=linea['cuenta'])

            # Turn these into plain/html MIMEText objects
            part1 = MIMEText(text, "plain")
            part2 = MIMEText(html, "html")

            # Add HTML/plain-text parts to MIMEMultipart message
            # The email client will try to render the last part first
            message.attach(part1)
            message.attach(part2)

            print('Enviando a: ', linea)
            #log --- print('Mensaje: ', message.as_string().format(nombres=linea['nombres'],apellido=linea['apellido'],cuenta=linea['cuenta']))
            server.sendmail(sender_email, linea['mail'], message.as_string().format(nombres=linea['nombres'],apellido=linea['apellido'],cuenta=linea['cuenta']))

file.close()
print('************************ PROCESO TERMINADO ************************')
