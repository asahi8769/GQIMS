import sendgrid
import os
from sendgrid.helpers.mail import Mail, Email, To, Content, HtmlContent, Attachment, Cc, Personalization


def send_mail(to_emails, ccs, subject, html_content, reply_to):
    sg = sendgrid.SendGridAPIClient(api_key=os.environ.get('SENDGRID_KEY'))
    from_email = Email("lih@glovis.net")  # Change to your verified sender

    receivers = []
    for to in to_emails:
        receivers.append(To(to))
    for cc in ccs:
        receivers.append(Cc(cc))

    receivers1 = [To("asahi8769@gmail.com"), To("lih@glovis.net"), Cc("lih1500252@naver.com"), Cc("ladfih1500252@naver.com")]

    mail = Mail(
        from_email=from_email,
        to_emails=receivers,
        subject='subject',
        plain_text_content="")

    mail.add_attachment("1.txt")

    mail_json = mail.get()
    response = sg.client.mail.send.post(request_body=mail_json)

    print(response.status_code, response.body, response.headers)


if __name__ =="__main__":

    to_emails = ['asahi8769@gmail.com', "lih@glovis.net"]
    ccs = ['lih1500252@naver.com', "ladfih1500252@naver.com"]
    subject = "Sending with SendGrid is Fun"
    html_content = "<h1>yes hello</h1>"

    send_mail(to_emails, ccs, subject, html_content, "lih@glovis.net")







