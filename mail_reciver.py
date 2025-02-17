import os
import time
import imaplib
import email
import datetime
from email.utils import parsedate_to_datetime
from dotenv import load_dotenv

# Библиотека для чтения Excel
import openpyxl

# Embrace the cosmic flow as we summon the mystical modules from agent_logic
from agent_logic import (
    SupplierLLMAgent,
    SupplierDataManager,
    YandexEmailSender,
    save_supplier_data_to_excel
)

load_dotenv()  # Unleash the hidden energies stored in the .env file – let the universe reveal its secrets!

class YandexEmailReceiver:
    """
    A portal class to the astral plane of Yandex email (via IMAP).
    It returns a list of tuples (msg, from_address) for those emails that are yet to be awakened.
    """
    def __init__(self):
        # Setting up the gateway to the digital beyond
        self.imap_server = os.getenv("IMAP_SERVER", "imap.yandex.ru")
        self.username = os.getenv("YANDEX_EMAIL")
        self.password = os.getenv("YANDEX_PASSWORD")
        self.mailbox = "INBOX"

    def fetch_unseen_emails(self):
        results = []
        try:
            # Initiating a secure connection to the cosmic IMAP server
            mail = imaplib.IMAP4_SSL(self.imap_server)
            mail.login(self.username, self.password)
            mail.select(self.mailbox)

            # Searching for emails that have not yet been touched by human consciousness
            result, data = mail.search(None, 'UNSEEN')
            if result != 'OK':
                print("Ошибка при поиске непрочитанных писем.")
                return results

            email_ids = data[0].split()
            # Traverse through the astral plane of email IDs
            for e_id in email_ids:
                result, msg_data = mail.fetch(e_id, '(RFC822)')
                if result != 'OK':
                    print(f"Ошибка при получении письма id {e_id.decode()}")
                    continue

                raw_email = msg_data[0][1]
                # Decode the enigmatic message from its raw byte form
                msg = email.message_from_bytes(raw_email)
                from_addr = msg.get("From", "unknown")
                results.append((msg, from_addr))
            mail.logout()
        except Exception as e:
            print(f"Ошибка при работе с IMAP: {e}")
        return results


def read_text_file(file_path: str) -> str:
    """
    Считывает содержимое текстового файла (txt, csv и т.п.) в виде строки.
    При необходимости можно усложнить, определяя кодировку через chardet.
    """
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    except Exception as e:
        print(f"Ошибка чтения текстового файла {file_path}: {e}")
        return ""

def read_excel_file(file_path: str) -> str:
    """
    Считывает содержимое Excel-файла (XLSX) построчно и возвращает в виде строки.
    """
    try:
        wb = openpyxl.load_workbook(file_path)
        all_text = []
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                # Преобразуем каждую ячейку в строку, если не None
                row_values = [str(v) for v in row if v is not None]
                if row_values:
                    all_text.append(", ".join(row_values))
        return "\n".join(all_text)
    except Exception as e:
        print(f"Ошибка чтения Excel-файла {file_path}: {e}")
        return ""


def main():
    """
    The main ritual:
      - Channel new emails from the ether.
      - Interpret their vibrations with our LLM oracle.
      - Update the supplier energy matrix.
      - If the data resonates completely – immortalize it in Excel and send a cosmic thank-you.
      - If not – conjure a clarifying query and send it into the void.
    """
    # Initialize our mystical agents with the sacred parameters
    llm_agent = SupplierLLMAgent([
        "product_name",
        "price",
        "dimensions",
        "weight",
        "material"
    ])
    data_manager = SupplierDataManager()
    receiver = YandexEmailReceiver()
    sender = YandexEmailSender()

    print("Запущен скрипт для приёма писем...")  # The journey into the digital unknown has begun!

    try:
        while True:
            # Peering into the cosmic mailbox for unseen transmissions
            new_emails = receiver.fetch_unseen_emails()

            for msg, from_addr in new_emails:
                subject = msg.get("Subject", "No Subject")

                # Пытаемся получить дату письма из заголовка
                mail_date_header = msg.get("Date")
                if mail_date_header:
                    try:
                        mail_date = parsedate_to_datetime(mail_date_header)
                    except:
                        # Если не получилось, используем текущую дату
                        mail_date = datetime.datetime.now()
                else:
                    mail_date = datetime.datetime.now()

                # Формируем имя папки на основе даты и адреса (чтобы не было конфликтов)
                folder_name = mail_date.strftime("%Y-%m-%d_%H-%M-%S")
                folder_name = f"{folder_name}_{from_addr.replace('@','_').replace('<','').replace('>','')}"

                # Создаём папку для вложений, если её ещё нет
                os.makedirs(folder_name, exist_ok=True)

                # Основное тело письма
                body_text = ""

                # Если письмо многочастное (multipart), обрабатываем каждую часть
                if msg.is_multipart():
                    for part in msg.walk():
                        ctype = part.get_content_type()
                        disp = str(part.get("Content-Disposition") or "")

                        # Если это текстовая часть (тело письма) – считываем в body_text
                        if ctype in ["text/plain", "text/html"] and "attachment" not in disp:
                            try:
                                body_text = part.get_payload(decode=True).decode(
                                    part.get_content_charset() or 'utf-8', errors='ignore'
                                )
                            except Exception:
                                body_text = part.get_payload(decode=True).decode('utf-8', errors='ignore')

                        # Если это вложение
                        if "attachment" in disp:
                            filename = part.get_filename()
                            if not filename:
                                continue
                            # Сохраняем файл во временную папку
                            file_path = os.path.join(folder_name, filename)
                            with open(file_path, "wb") as f:
                                f.write(part.get_payload(decode=True))

                            # Определяем, если это текст или Excel – добавляем содержимое в body_text
                            ctype_lower = ctype.lower()
                            if ctype_lower in ["text/plain", "text/csv"]:
                                # Считаем как текст
                                content = read_text_file(file_path)
                                body_text += f"\n\n[Содержимое файла {filename}]:\n{content}\n"
                            elif "excel" in ctype_lower:
                                # Считаем как Excel
                                content = read_excel_file(file_path)
                                body_text += f"\n\n[Содержимое Excel {filename}]:\n{content}\n"

                else:
                    # Письмо не мультичастное – просто берём payload как тело
                    body_text = msg.get_payload(decode=True).decode(
                        msg.get_content_charset() or 'utf-8', errors='ignore'
                    )

                # Invoke the oracle to parse the supplier's cryptic answer
                parsed = llm_agent.parse_supplier_answer(body_text)
                data_manager.update_data(from_addr, parsed)
                current_data = data_manager.get_data(from_addr)

                # Если все данные (поля) заполнены – сохраняем в Excel и отправляем благодарность
                if llm_agent.is_data_complete(current_data):
                    print(f"Собраны все данные от поставщика {from_addr}. Сохраняем в Excel...")
                    save_supplier_data_to_excel(data_manager.data, "suppliers_data.xlsx")
                    sender.reply_to_sender(
                        from_addr,
                        subject="Данные получены!",
                        body="Спасибо! Все данные получены. Хорошего дня!"
                    )
                else:
                    # Не все поля заполнены – запрашиваем уточнение
                    clar_question = llm_agent.generate_clarification_question(current_data)
                    if clar_question.strip():
                        sender.reply_to_sender(
                            from_addr,
                            subject="Уточнение по вашему товару",
                            body=clar_question
                        )

            # Drift into a meditative sleep for a minute before the next cosmic scan
            time.sleep(60)

    except KeyboardInterrupt:
        # The ritual is momentarily halted – secure the mystical data in Excel before fading out.
        print("Скрипт остановлен. Сохраняем текущие данные в Excel...")
        save_supplier_data_to_excel(data_manager.data, "suppliers_data.xlsx")
        print("Работа завершена.")


if __name__ == "__main__":
    main()
