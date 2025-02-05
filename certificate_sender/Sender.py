import smtplib
import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

'''Excel dosyası Book1 dosyası formatında, proje konumunda olmalı
 Pdf'i oluşturmak için aynı excel dosyası ile ms word mail merge özelliği kullanılabilir'''

# SMTP Ayarları
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_ADDRESS = "yourname@gmail.com"  # Kendi e-posta adresiniz
EMAIL_PASSWORD = "app_password"  # Gmail uygulama şifresi

# Sertifikaların bulunduğu PDF dosyasının yolu
PDF_FILE_PATH = "/Users/user1/Desktop/certificates.pdf"

# Excel Dosyasını Oku
data = pd.read_excel("Book1.xlsx")  # Excel dosyanızın adı ve yolu

# PDF dosyasını oku
reader = PdfReader(PDF_FILE_PATH)

# Herkese Sertifikalarını Gönder
for index, row in data.iterrows():
    recipient_email = row['email']  # Excel'deki Email sütunu
    recipient_name = row['name']   # Excel'deki Name sütunu
    page_number = row['page'] - 1  # PDF sayfası sıfırdan başladığı için -1 yapıyoruz

    # Null kontrolü: Email sütunu boşsa atlama yap
    if pd.isna(recipient_email) or recipient_email.strip() == "":
        print(f"Skipping email for {recipient_name} because email is missing.")
        continue

    # PDF'den ilgili sayfayı al
    writer = PdfWriter()
    writer.add_page(reader.pages[page_number])

    # Yeni PDF dosyasını kaydet (geçici dosya)
    temp_pdf_path = f"temp_{recipient_name.replace(' ', '_')}_certificate.pdf"
    with open(temp_pdf_path, "wb") as temp_pdf:
        writer.write(temp_pdf)

    # E-posta Gövdesi
    subject = "Yapay Zeka Etkinliği Sertifikanız"
    body = f'Merhaba {recipient_name},\n\nKatılımın için teşekkürler. Katılım sertifikanı ekte bulabilirsin. \nYapacağımız diğer etkinliklerde görüşmek üzere.❤️'

    # E-posta Oluştur
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Sertifikayı Ekle (Yeni oluşturulan PDF dosyasını ekle)
    with open(temp_pdf_path, 'rb') as attachment:
        part = MIMEBase('application', 'pdf')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename="certificate.pdf"'
        )
        msg.attach(part)

    # E-posta Gönder
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, recipient_email, msg.as_string())
        print(f"E-mail sent to {recipient_name} at {recipient_email}")
    except Exception as e:
        print(f"Error sending email to {recipient_email}: {e}")

    # Geçici PDF dosyasını sil
    os.remove(temp_pdf_path)
