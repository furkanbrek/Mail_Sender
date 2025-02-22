import smtplib
import os
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

class EmailSender:
    def __init__(self):
        # SMTP Ayarları
        self.smtp_server = "smtp.gmail.com"
        self.smtp_port = 587
        self.email_address = os.getenv("EMAIL_ADDRESS")
        self.email_password = os.getenv("EMAIL_PASSWORD")
        
        if not self.email_address or not self.email_password:
            raise ValueError("Email ve şifre bilgileri .env dosyasında bulunamadı!")

    def send_mass_email(self, excel_path, attachment_path=None, subject=None, body_template=None):
        """
        Toplu email gönderme fonksiyonu
        
        :param excel_path: Excel dosyasının yolu
        :param attachment_path: Opsiyonel ek dosya yolu
        :param subject: Email konusu
        :param body_template: Email içeriği şablonu
        """
        try:
            # Excel dosyasını oku
            data = pd.read_excel(excel_path)
            
            # Gerekli sütunların varlığını kontrol et
            required_columns = ['email', 'name']
            if not all(col in data.columns for col in required_columns):
                raise ValueError("Excel dosyasında gerekli sütunlar eksik! Gerekli sütunlar: email, name")

            # SMTP sunucusuna bağlan
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.email_address, self.email_password)

                # Her alıcı için email gönder
                for index, row in data.iterrows():
                    if pd.isna(row['email']) or str(row['email']).strip() == "":
                        print(f"UYARI: {row['name']} için email adresi bulunamadı, atlanıyor.")
                        continue

                    try:
                        # Email mesajını oluştur
                        msg = self._create_email_message(
                            recipient_email=row['email'],
                            recipient_name=row['name'],
                            subject=subject or "Bilgilendirme Mesajı",
                            body_template=body_template,
                            attachment_path=attachment_path
                        )

                        # Emaili gönder
                        server.sendmail(self.email_address, row['email'], msg.as_string())
                        print(f"✓ Email başarıyla gönderildi: {row['name']} ({row['email']})")
                    
                    except Exception as e:
                        print(f"✗ {row['email']} adresine email gönderilirken hata oluştu: {str(e)}")

        except Exception as e:
            print(f"Toplu email gönderme işlemi sırasında hata oluştu: {str(e)}")

    def _create_email_message(self, recipient_email, recipient_name, subject, body_template, attachment_path=None):
        """Email mesajını oluştur"""
        msg = MIMEMultipart()
        msg['From'] = self.email_address
        msg['To'] = recipient_email
        msg['Subject'] = subject

        # Varsayılan mesaj şablonu
        default_body = f"""
Sayın {recipient_name},

Bu bir toplu email bilgilendirme mesajıdır.

Saygılarımızla,
"""
        # Mesaj içeriğini ekle
        body = body_template.format(name=recipient_name) if body_template else default_body
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        # Ek dosya varsa ekle
        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{os.path.basename(attachment_path)}"'
                )
                msg.attach(part)

        return msg

if __name__ == "__main__":
    # Kullanım örneği
    sender = EmailSender()
    
    # Özel mesaj şablonu
    message_template = """
İyi günler,


Ben Furkan Berk, Işık Üniversitesi'nde Yazılım Mühendisliği öğrencisiyim. Aynı zamanda IEEE Computer Society Türkiye'de öğrenci temsilciliği yapıyorum.  Sizi bu sene 8.si düzenlenecek olan IEEE Türkiye Computer Society Conference (CSCON’25) etkinliğinde stant açmak için davet ediyoruz. 21-22-23 Mart tarihlerinde Gazi Üniversitesi ev sahipliğinde gerçekleşecektir. 21 Mart Konferans, 22-23 Mart Eğitim ve 23 Mart Yarışma modülü olarak 3 ana kısımdan oluşuyor.


Etkinliğimizin ilk gününde 1000+ kişilik konferans gerçekleştirmeyi ikinci ve üçüncü gününde 300 kişilik kitleye paralel eğitimler düzenlemek istiyoruz. Sponsorlarımızı tüm duyuru yollarımızda duyuracağımızı stant açma ve etkinlik içerisinde oturum alma gibi haklardan yararlanabileceğinizi söylemek isterim. Detayları Sponsorluk dosyamızdan da inceleyebilirsiniz.


IEEE Türkiye 110 üniversitedeki IEEE öğrenci kulübünün öğrencileri , akademisyenlerin ve profesyonellerin ortak buluştuğu bir mühendislik derneğidir. Etkinliklerimizde okullar arasında etkileşimi arttırmayı ve mühendis adaylarında yazılım kültürünü oluşturmayı hedefliyoruz. Katkılarınız bizim için çok kıymetli olur. Etkinliğimizin detaylarını Ekteki dosyamızdan  da inceleyebilirsiniz.


Saygılarımla,
"""
    
    # Toplu email gönder
    sender.send_mass_email(
        excel_path="/email.xlsx",
        attachment_path="/CSCON Tanıtım Dosyası.pdf",  # Opsiyonel
        subject="CSCON'25",
        body_template=message_template
    )
