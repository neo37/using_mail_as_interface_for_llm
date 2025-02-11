import imaplib
import email
import time
import smtplib
from dataclasses import dataclass
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT
# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT
# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT
# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT
# --- Отправка писем, чё бы нам не скучно было ---

# DRAFT# DRAFT
# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT# DRAFT
@dataclass
class EventData:
    subject: str
    start_time: datetime
    end_time: datetime
    description: str
    location: str

class YandexEmailService:
    def __init__(self, smtp_server: str, smtp_port: int, sender_email: str, sender_password: str):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.sender_email = sender_email
        self.sender_password = sender_password

    def send_email(self, recipients_emails: list, subject: str, body: str, ics_content: str = None):
        # Собираем список получателей
        recipients = ','.join(recipients_emails) if len(recipients_emails) > 1 else recipients_emails[0]

        msg = MIMEMultipart()
        msg['From'] = self.sender_email
        msg['To'] = recipients
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Если есть приглашение в календарь, то прикрепляем его
        if ics_content:
            ics = MIMEBase('text', 'calendar', method='REQUEST', name='invite.ics')
            ics.set_payload(ics_content.encode('utf-8'))
            encoders.encode_base64(ics)
            ics.add_header('Content-Type', 'text/calendar; method=REQUEST; charset=UTF-8; component=VEVENT')
            ics.add_header('Content-Disposition', 'attachment; filename="invite.ics"')
            msg.attach(ics)

        server = None
        try:
            # Подключаемся к серверу и шлём письмо
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()  # начинаем защищённое соединение
            server.login(self.sender_email, self.sender_password)
            server.send_message(msg)
            print("Письмо успешно отправлено!")
        except Exception as e:
            print(f"Ошибка при отправке письма: {e}")
        finally:
            if server:
                server.quit()

    def build_ics_content(self, event: EventData) -> str:
        # Собираем содержимое ICS-файла
        ics_content = [
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            "PRODID:-//Python Calendar Event//EN",
            "BEGIN:VEVENT",
            f"SUMMARY:{event.subject}",
            f"DTSTART:{event.start_time.strftime('%Y%m%dT%H%M%SZ')}",
            f"DTEND:{event.end_time.strftime('%Y%m%dT%H%M%SZ')}",
            f"DESCRIPTION:{event.description}",
            f"LOCATION:{event.location}",
            "END:VEVENT",
            "END:VCALENDAR"
        ]
        return "\r\n".join(ics_content)

# --- Получение писем, чтобы не пропустить ничего интересного ---
class YandexEmailReceiver:
    def __init__(self, imap_server: str, username: str, password: str, mailbox: str = 'inbox'):
        self.imap_server = imap_server
        self.username = username
        self.password = password
        self.mailbox = mailbox

    def check_for_excel_attachment(self):
        try:
            # Подключаемся к почтовому серверу IMAP (все по-старому, только через SSL)
            mail = imaplib.IMAP4_SSL(self.imap_server)
            mail.login(self.username, self.password)
            mail.select(self.mailbox)  # выбираем папку "inbox"

            # Ищем все непрочитанные письма
            result, data = mail.search(None, 'UNSEEN')
            if result != 'OK':
                print("Чё-то пошло не так при поиске писем.")
                return

            email_ids = data[0].split()
            for e_id in email_ids:
                result, msg_data = mail.fetch(e_id, '(RFC822)')
                if result != 'OK':
                    print(f"Письмо с id {e_id.decode()} не получилось скачать, пропускаем...")
                    continue

                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                # Бежим по всем частям письма и ищем вложения
                for part in msg.walk():
                    content_disposition = part.get("Content-Disposition", "")
                    if "attachment" in content_disposition.lower():
                        filename = part.get_filename()
                        # Если нашли Excel, то кричим "ура"
                        if filename and filename.lower().endswith(('.xls', '.xlsx')):
                            print("ура, Excel-файл обнаружен!")
            mail.logout()
        except Exception as e:
            print(f"Упс, ошибка при проверке почты: {e}")

# --- Основной блок: запускаем всё и валим в бесконечный цикл ---
if __name__ == '__main__':
    # Параметры для отправки почты через Yandex
    smtp_server = 'smtp.yandex.ru'
    smtp_port = 587
    sender_email = 'your_yandex_email@yandex.ru'   # вставь сюда свой email
    sender_password = 'your_password'              # и свой пароль (или специальный пароль)

    email_service = YandexEmailService(smtp_server, smtp_port, sender_email, sender_password)

    # Пример отправки приглашения (если надо, раскомментируй и заполни данные)
    # event = EventData(
    #     subject="Встреча",
    #     start_time=datetime(2025, 2, 15, 12, 0, 0),
    #     end_time=datetime(2025, 2, 15, 13, 0, 0),
    #     description="Обсудим важные дела",
    #     location="Офис"
    # )
    # ics_content = email_service.build_ics_content(event)
    # email_service.send_email(['recipient@example.com'], "Приглашение на встречу", "Привет! Вот приглашение.", ics_content)

    # Параметры для проверки входящих писем через IMAP (Yandex)
    imap_server = 'imap.yandex.ru'
    receiver_email = sender_email   # для простоты, тот же email
    receiver_password = sender_password

    email_receiver = YandexEmailReceiver(imap_server, receiver_email, receiver_password)

    # Запускаем проверку почты каждые 60 секунд
    print("Запуск проверки почты, не отключай меня!")
    while True:
        email_receiver.check_for_excel_attachment()
        time.sleep(60)  # спим минутку и снова в бой
