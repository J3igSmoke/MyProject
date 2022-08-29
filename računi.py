from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import openpyxl
import smtplib
import qrcode
from PIL import Image
import datetime

racun = openpyxl.load_workbook("POLOZNICA.xlsx")

html_email = ("""
<html>

<head>
    <style>
        .banner-color {
            background-color: #eb681f;
        }

        .title-color {
            color: #0066cc;
        }

        .button-color {
            background-color: #0066cc;
        }

        @media screen and (min-width: 500px) {
            .banner-color {
                background-color: #0066cc;
            }

            .title-color {
                color: #eb681f;
            }

            .button-color {
                background-color: #eb681f;
            }
        }

        .main {
            background-color: #ececec;
            padding: 20px;
            margin: 0 auto;
            font-weight: 200;
        }
    </style>
</head>

<body>
    <div class="main">
        <div>
        <h2> Spoštovani! </h2>
            <h3>V prilogi je opomin za članarino.</h3>
            <p>
                Prosimo za nakazilo ali za plačilo v gotovini, ki jo lahko oddate blagajniku Branku ob sredah  ob 17.00 uri. <br> 
                <br> 
                Veseli bomo za čim večjo podporno članstvo. Članarina za podporne člane je 1 € ali več po lastni presoji. Vsak član šteje.<br> 
                <br> 
                QR koda deluje samo na nekaterih E-bankah. <br> 
                V primeru, da hitri vnos ne deluje, podatke vnesite ročno.<br> 
                <br> 
                V primeru, da ste članarino že plačali, položnice ni potrebno plačati.
            </p>
            <p>
                Lep pozdrav! <br> 
                Klemen in Branko.
            </p>
        </div>
    </div>
</body>

</html>
""")
a=1
zac_stevilka=int(input("Enter starting number: "))
koc_stevilka=int(input("Enter ending number: "))
List_podatek= racun.sheetnames
podatek1=racun["Clanarina"]
osnutek=racun["Poloznica"]

def send_mail(send_from, send_to, subject, message, file):
    msg = MIMEMultipart("alternative")
    msg['Subject'] = subject
    msg.attach(MIMEText(message.encode('utf-8'), 'html', 'UTF-8'))

    part = MIMEBase('application', "octet-stream")
    with open(file, 'rb') as file:
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition',
                    'attachment; filename=Poloznica.pdf')
    msg.attach(part)

    server_ssl.sendmail(send_from, send_to, msg.as_string())


USER = "email"
PASSWORD = "password"

server_ssl = smtplib.SMTP_SSL("smtp.gmail.com", 465)
server_ssl.ehlo()
server_ssl.login(USER, PASSWORD)
k=0
for i in podatek1:
   if(zac_stevilka<=a and koc_stevilka>=a):

        imeOsebe = podatek1["A" + str(a)].value
        NaslovOsebe = podatek1["B" + str(a)].value
        obcinaOsebe = podatek1["C" + str(a)].value
        cena = podatek1["E" + str(a)].value
        mailNaslov= podatek1["F" + str(a)].value
        osnutek["F33"] = imeOsebe
        osnutek["F34"] = NaslovOsebe
        osnutek["F35"] = obcinaOsebe
        osnutek["G41"] = cena
        datum = datetime.datetime.strptime(str(datetime.datetime.now().date()), "%Y-%m-%d").strftime("%d.%m.%Y")
        kodacount = ("UPNQR\r\n\r\n\r\n\r\n\r\n"+imeOsebe +"\r\n"+NaslovOsebe+"\r\n"+obcinaOsebe +"\r\n0000000"+str(cena)+"00\r\n\r\n\r\nOTHR\r\nPlačilo članarine 2021/22\r\n"+ str(datum)+"\r\nSI56070000000193411\r\nSI005555\r\nSTRELSKO DRUŠTVO LOTRIČ ŽELEZNIKI\r\nČešnjica 10\r\n4228 Železniki\r\n")
        koda = qrcode.make("UPNQR\r\n\r\n\r\n\r\n\r\n"+imeOsebe +"\r\n"+NaslovOsebe+"\r\n"+obcinaOsebe +"\r\n0000000"+str(cena)+"00\r\n\r\n\r\nOTHR\r\nPlačilo članarine 2021/22\r\n"+ str(datum)+"\r\nSI56070000000193411\r\nSI005555\r\nSTRELSKO DRUŠTVO LOTRIČ ŽELEZNIKI\r\nČešnjica 10\r\n4228 Železniki\r\n"+ str(len(kodacount)))
        koda.save("poloznica.jpg")
        #koda.save("poloznica2.jpg")
        image = Image.open('poloznica.jpg')
        new_image = image.resize((162, 162))
        new_image.save('poloznica.jpg')

        qr = openpyxl.drawing.image.Image('poloznica.jpg')
        qr.anchor = 'A43'
        osnutek.add_image(qr)

        racun.save("POLOZNICA.xlsx")

        izpis = "Položnica " + imeOsebe
        print(izpis)
        racun.save("POLOZNICA.xlsx")
        racun.close()
        from win32com import client

        # Open Microsoft Excel
        excel = client.Dispatch("Excel.Application")

        # Read Excel File
        sheets = excel.Workbooks.Open("C:/Users/Rok/PycharmProjects/Certifikati/POLOZNICA.xlsx")
        work_sheets = sheets.Worksheets[0]
        imeOsebe = str(imeOsebe).replace(" ","_")
        # Convert into PDF File
        try:
            work_sheets.ExportAsFixedFormat(0, "C:/Users/Rok/PycharmProjects/Certifikati/" + "POLOŽNICA_" +imeOsebe )
        except:
            print("invoice created")
        sheets.Close(True)
        try:
            send_mail(USER, mailNaslov, "Poloznica", html_email,  "POLOŽNICA_" +imeOsebe+".pdf")
            print("invoice PDF sent via email")
        except:
            print("email sending error")



        a += 1
   else:
       a += 1

server_ssl.close()