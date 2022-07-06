import smtplib, ssl
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header
from email.utils import formataddr
from utils_.config import EMAIL_ADDRESS, EMAIL_PASSWORD
import os


class AutoMailing:
    email_address = EMAIL_ADDRESS
    email_password = EMAIL_PASSWORD

    def __init__(self, receiver=None, cc=None, subject=None, html=None, attachments=None,
                 img: dict = None, reply_to='lih@glovis.net'):
        self.receiver = receiver
        self.cc = cc
        self.subject = subject
        self.attachments = attachments
        self.img = img
        self.reply_to = reply_to
        self.msg = None
        self.html = html

    def initiate(self):
        self.msg = MIMEMultipart()
        self.msg['From'] = formataddr((str(Header(f'이일희 매니저', 'utf-8')), f'lih1500252@naver.com'))
        self.msg['To'] = self.receiver
        self.msg['Cc'] = self.cc
        self.msg['Reply-To'] = self.reply_to
        self.msg['Subject'] = self.subject
        self.print_contents()

    def print_contents(self):
        print('\n<<Email Information>>')
        print('From :', self.email_address)
        print('To :', self.receiver)
        print('Cc :', self.cc)
        print('Reply-To :', self.reply_to)
        print('Subject :', self.subject)

    def attach_file(self):
        if self.attachments is None:
            pass
        else:
            if type(self.attachments) is str:
                self.attachments = [self.attachments]
            for n, file in enumerate(self.attachments):
                print(f'Attachment_{n+1} : {file}')
                with open(file, 'rb') as f:
                    file_data = MIMEApplication(f.read(), Name=os.path.basename(file))
                    self.msg.attach(file_data)

    def attache_img(self):
        """ https://stackoverflow.com/questions/48272100/embed-an-image-in-html-for-automatic-outlook365-email-send """
        if self.img is None:
            pass

        else:
            for n, cid in enumerate(self.img.keys()):
                with open(self.img[cid], 'rb') as file:
                    img = MIMEImage(file.read())
                    img.add_header('Content-ID', f'<{cid}>')
                    self.msg.attach(img)
                    print(f'Image_{n+1} : {self.img[cid]}')

    def send(self):
        self.initiate()
        self.attach_file()
        self.attache_img()

        self.msg.attach(MIMEText(self.html, _subtype='html'))

        confirm = input("\n정말로 메일을 송부합니까?(y/n, 기본값 n) : ").lower()

        if confirm == "y":
            with smtplib.SMTP_SSL('smtp.naver.com', 465) as smtp:
                smtp.login(self.email_address, self.email_password)
                smtp.send_message(self.msg)
                print('메일이 송부되었습니다!\n')
        else:
            print("메일이 송부되지 않았습니다.\n")


if __name__ == "__main__":
    from modules.mail.html.monthly_html import index_html
    import pandas as pd
    import os

    os.chdir(os.pardir)
    os.chdir(os.pardir)

    receiver = "lih@glovis.net,asahi8769@gmail.com"
    cc = "asahii8769@gmail.com"
    subject = "테스트용 메일입니다."
    attachments = "spawn/(202109)KD종합품질지수(GQMS기준)_2021-10-31_20_58_41.xlsx"

    yyyymm = 202109

    ovs_plot1 = os.path.abspath(os.path.join("spawn", "plots", f"{yyyymm}_plot1.png"))
    footer = os.path.abspath(os.path.join("images", "mail_footer.png"))
    img = {"ovs_plot1": ovs_plot1, "footer": footer}

    df = pd.read_excel("spawn/(202109)KD종합품질지수(GQMS기준)_2021-10-31_20_58_41.xlsx",
                       sheet_name='해외품질지수(GQMS+공급망)', index_col=[1, 2, 3, 4], skiprows=3)

    download_date = '2021-01-05'
    html, _ = index_html(yyyymm, df, download_date)

    mail = AutoMailing(receiver=receiver, cc=cc, subject=subject, html=html, attachments=attachments, img=img)
    mail.send()