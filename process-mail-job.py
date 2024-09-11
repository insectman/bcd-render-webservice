
import os
import smtplib
import email
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from imapclient import IMAPClient
from barcode import EAN13
from barcode.writer import ImageWriter
import py7zr

from email.header import decode_header

# app = Flask(__name__)

YANDEX_LOGIN = os.environ.get('YANDEX_LOGIN')
YANDEX_MAIL = os.environ.get('YANDEX_MAIL')
YANDEX_PASS = os.environ.get('YANDEX_PASS')
DIAP_MAIL = os.environ.get('DIAP_MAIL')

if not YANDEX_LOGIN:
    from dotenv import load_dotenv
    load_dotenv()
    YANDEX_LOGIN = os.getenv('YANDEX_LOGIN')
    YANDEX_MAIL = os.getenv('YANDEX_MAIL')
    YANDEX_PASS = os.getenv('YANDEX_PASS')
    DIAP_MAIL = os.getenv('DIAP_MAIL')

print('YANDEX_LOGIN', YANDEX_LOGIN)
print('YANDEX_MAIL', YANDEX_MAIL)
print('YANDEX_PASS', YANDEX_PASS)
print('DIAP_MAIL', '>'+DIAP_MAIL+'<')

writer=ImageWriter()

print('writer type', type(writer))

def process_mail():
    print('processing mail')
    with IMAPClient(host="imap.ya.ru") as client:
        client.login(YANDEX_LOGIN, YANDEX_PASS)
        client.select_folder('INBOX')

        messages = client.search(['FROM', DIAP_MAIL])
        # fetch selectors are passed as a simple list of strings.
        response = client.fetch(messages, ['FLAGS', 'RFC822', 'ENVELOPE'])


        for dir in ['input', 'output', 'output/archives', 'output/imgs']:
            if os.path.exists(dir):
                os.chmod(dir, 0o777)

        if not os.path.exists('input'):
            os.makedirs('input')
        if not os.path.exists('output'):
            os.makedirs('output')
        for f in os.listdir('input'):
            os.remove(os.path.join('input', f))
        if not os.path.exists('output/archives'):
            os.makedirs('output/archives')
        if not os.path.exists('output/imgs'):
            os.makedirs('output/imgs')
        for f in os.listdir('output/imgs'):
            os.remove(os.path.join('output/imgs', f))
        for f in os.listdir('output/archives'):
            os.remove(os.path.join('output/archives', f))

        for dir in ['input', 'output', 'output/archives', 'output/imgs']:
            os.chmod(dir, 0o777)


        # `response` is keyed by message id and contains parsed,
        # converted response items.
        for message_id, message_data in response.items():
            email_message = email.message_from_bytes(message_data[b'RFC822'])
            print(message_id, email_message.get('From'), email_message.get('Subject'))
            for part in email_message.walk():
              if part.get_content_type() == 'text/plain':
                payload = part.get_payload(decode=True)
                filename = part.get_filename()
                # print(filename, payload)

                if filename:
                    filename = decode_header(filename)[0][0].decode(decode_header(filename)[0][1])

                    with open('input/' + filename, 'wb') as f:
                        f.write(payload)
                    with open('input/' + filename, 'r') as f:
                        for line in f:
                            line = line.strip()
                            if not line:
                                continue
                            with open("output/imgs/" + line + ".jpg", "wb") as f:
                                EAN13(str(line), writer).write(f)
                    with py7zr.SevenZipFile('output/archives/' + filename + ".7z", 'w') as archive:
                        archive.writeall("output/imgs/", '')

                    print("Saved to file: %s" % filename)

            client.move(message_id, 'auto_completed')

            # compose a new email, attach all files in output/archives to it and send to 'insectman.yandex.ru'
            msg = MIMEMultipart()
            msg['From'] = YANDEX_MAIL
            msg['To'] = YANDEX_MAIL
            msg['Subject'] = 'Re: ' + email_message.get('Subject')

            for file_name in os.listdir('output/archives'):
                with open('output/archives/' + file_name, "rb") as f:
                    part = MIMEApplication(
                        f.read(),
                        Name=file_name
                    )
                part['Content-Disposition'] = 'attachment; filename="%s"' % file_name
                msg.attach(part)

            smtp = smtplib.SMTP_SSL('smtp.yandex.com.tr', 465)
            smtp.login(YANDEX_MAIL, YANDEX_PASS)
            # smtp.sendmail(YANDEX_MAIL, DIAP_MAIL, msg.as_string())
            print('sendmail:start')
            smtp.sendmail(YANDEX_MAIL, YANDEX_MAIL, msg.as_string())
            print('sendmail:done')
            smtp.quit()

        print("Number of messages: %d" % len(response.items()))


process_mail()

