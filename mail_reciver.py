import os
import time
import imaplib
import email
from dotenv import load_dotenv

# Импортируем всё необходимое из agent_logic
from agent_logic import (
    SupplierLLMAgent, 
    SupplierDataManager, 
    YandexEmailSender, 
    save_supplier_data_to_excel
)

load_dotenv()  # Подгружаем переменные окружения

class YandexEmailReceiver:
    """
    Простой класс для чтения писем из Яндекс-Почты (IMAP).
    Возвращает список (email_message, from_address) для непрочитанных писем.
    """
    def __init__(self):
        self.imap_server = os.getenv("IMAP_SERVER", "imap.yandex.ru")
        self.username = os.getenv("YANDEX_EMAIL")
        self.password = os.getenv("YANDEX_PASSWORD")
        self.mailbox = "INBOX"

    def fetch_unseen_emails(self):
        """
        Возвращает [(msg, from_addr), ...] для всех непрочитанных (UNSEEN) писем.
        """
        results = []
        try:
            mail = imaplib.IMAP4_SSL(self.imap_server)
            mail.login(self.username, self.password)
            mail.select(self.mailbox)

            result, data = mail.search(None, 'UNSEEN')
            if result != 'OK':
                print("Ошибка при поиске непрочитанных писем.")
                return results

            email_ids = data[0].split()
            for e_id in email_ids:
                result, msg_data = mail.fetch(e_id, '(RFC822)')
                if result != 'OK':
                    print(f"Ошибка при получении письма id {e_id.decode()}")
                    continue

                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                from_addr = msg.get("From", "unknown")
                results.append((msg, from_addr))

            mail.logout()
        except Exception as e:
            print(f"Ошибка при работе с IMAP: {e}")

        return results


def main():
    """
    Основная логика: бесконечно читаем новые письма, парсим их, 
    сохраняем данные о товарах и при необходимости отправляем уточняющие письма.
    """
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

    print("Запущен скрипт для приёма писем...")

    try:
        while True:
            # Ищем новые письма (непрочитанные)
            new_emails = receiver.fetch_unseen_emails()

            for msg, from_addr in new_emails:
                # Получаем тему письма (пригодится для отправки ответа)
                subject = msg.get("Subject", "No Subject")

                # Извлекаем текст письма. Если multipart, ищем text/plain или text/html
                body_text = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        ctype = part.get_content_type()
                        if ctype in ["text/plain", "text/html"]:
                            try:
                                body_text = part.get_payload(decode=True).decode(
                                    part.get_content_charset() or 'utf-8'
                                )
                            except Exception:
                                body_text = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                            break
                else:
                    body_text = msg.get_payload(decode=True).decode(msg.get_content_charset() or 'utf-8', errors='ignore')

                # 1. Парсим письмо поставщика через LLM
                parsed = llm_agent.parse_supplier_answer(body_text)
                # 2. Обновляем данные
                data_manager.update_data(from_addr, parsed)
                current_data = data_manager.get_data(from_addr)

                # 3. Проверяем, все ли поля собраны
                if llm_agent.is_data_complete(current_data):
                    print(f"Собраны все данные от поставщика {from_addr}. Сохраняем в Excel...")
                    save_supplier_data_to_excel(data_manager.data, "suppliers_data.xlsx")
                    # Отправляем письмо поставщику, что всё ок
                    sender.reply_to_sender(
                        from_addr,
                        subject="Данные получены!",
                        body="Спасибо! Все данные получены. Хорошего дня!"
                    )
                else:
                    # Если чего-то не хватает — генерируем уточняющий вопрос
                    clar_question = llm_agent.generate_clarification_question(current_data)
                    if clar_question.strip():
                        sender.reply_to_sender(
                            from_addr,
                            subject="Уточнение по вашему товару",
                            body=clar_question
                        )

            # Ждём 60 секунд и повторяем
            time.sleep(60)

    except KeyboardInterrupt:
        print("Скрипт остановлен. Сохраняем текущие данные в Excel...")
        save_supplier_data_to_excel(data_manager.data, "suppliers_data.xlsx")
        print("Работа завершена.")


if __name__ == "__main__":
    main()
