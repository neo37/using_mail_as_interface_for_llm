import os
import time
import imaplib
import email
from dotenv import load_dotenv

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
                body_text = ""
                if msg.is_multipart():
                    # Dive into the many-layered dimensions of the message
                    for part in msg.walk():
                        ctype = part.get_content_type()
                        if ctype in ["text/plain", "text/html"]:
                            try:
                                # Decode the message part like deciphering ancient glyphs
                                body_text = part.get_payload(decode=True).decode(
                                    part.get_content_charset() or 'utf-8'
                                )
                            except Exception:
                                body_text = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                            break
                else:
                    # For messages of a singular essence, decode directly
                    body_text = msg.get_payload(decode=True).decode(
                        msg.get_content_charset() or 'utf-8', errors='ignore'
                    )

                # Invoke the oracle to parse the supplier's cryptic answer
                parsed = llm_agent.parse_supplier_answer(body_text)
                data_manager.update_data(from_addr, parsed)
                current_data = data_manager.get_data(from_addr)

                if llm_agent.is_data_complete(current_data):
                    # A full spectrum of data has been achieved – time to crystallize it in Excel!
                    print(f"Собраны все данные от поставщика {from_addr}. Сохраняем в Excel...")
                    save_supplier_data_to_excel(data_manager.data, "suppliers_data.xlsx")
                    sender.reply_to_sender(
                        from_addr,
                        subject="Данные получены!",
                        body="Спасибо! Все данные получены. Хорошего дня!"
                    )
                else:
                    # The data is still hazy – conjure up a clarifying question to clear the mystic fog
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
