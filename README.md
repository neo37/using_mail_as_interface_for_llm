Here's a non-formal but precise instruction on how to run the project:

---

1. **Make sure you have Python 3.7+ installed**  
   If you haven't installed it yet, grab it from [python.org](https://www.python.org/downloads/) and install it.

2. **Clone the repository or download the project files**  
   Your project folder should look something like this:
   ```
   your_project/
     ├─ .env
     ├─ agent_logic.py
     ├─ mail_receiver.py
     ├─ requirements.txt
   ```

3. **Create and configure the `.env` file**  
   In the project root (right next to `agent_logic.py` and `mail_receiver.py`), create a file called `.env` and put your credentials in it. Replace the placeholder values with your own:
   ```
   OPENAI_API_KEY=your_openai_api_key
   YANDEX_EMAIL=your_yandex_email@yandex.ru
   YANDEX_PASSWORD=your_yandex_password
   IMAP_SERVER=imap.yandex.ru
   SMTP_SERVER=smtp.yandex.ru
   SMTP_PORT=587
   ```
   This file stores all your secret stuff so that your credentials don’t end up in the code.

4. **Install the dependencies**  
   If you have a `requirements.txt` file, open your terminal in the project folder and run:
   ```bash
   pip install -r requirements.txt
   ```
   Otherwise, install the needed packages manually:
   ```bash
   pip install openai python-dotenv openpyxl
   ```

5. **Run the project**  
   Now that everything’s set up, run the mail receiver script from your terminal:
   ```bash
   python mail_receiver.py
   ```
   The script will start checking for new emails every 60 seconds. When an email from a supplier comes in, the LLM agent will try to extract the required data, send a clarification question (if needed), and save the final data into a file named `suppliers_data.xlsx`.

6. **Stop the script**  
   To stop the script, simply press `Ctrl+C` in the terminal. All the data collected so far will be saved in the Excel file.

---

If something goes wrong, check the terminal for error messages. Good luck, and may your suppliers reply promptly!))))))))))

===================================
Вот тебе неформальная, но точная инструкция по запуску проекта:

---

1. **Проверь, что у тебя установлен Python 3.7+**  
   Если еще не установил – скачай с [python.org](https://www.python.org/downloads/) и установи.

2. **Склонируй репозиторий или скачай файлы проекта**  
   Структура должна выглядеть примерно так:
   ```
   your_project/
     ├─ .env
     ├─ agent_logic.py
     ├─ mail_receiver.py
     ├─ requirements.txt
   ```
   
3. **Создай и настрой файл `.env`**  
   В корне проекта (рядом с `agent_logic.py` и `mail_receiver.py`) создай файл `.env` и вставь туда свои данные (замени значения на свои):
   ```
   OPENAI_API_KEY=your_openai_api_key
   YANDEX_EMAIL=your_yandex_email@yandex.ru
   YANDEX_PASSWORD=your_yandex_password
   IMAP_SERVER=imap.yandex.ru
   SMTP_SERVER=smtp.yandex.ru
   SMTP_PORT=587
   ```
   Этот файл хранит все секретные штуки, чтобы их не было в коде.

4. **Установи зависимости**  
   Если у тебя есть файл `requirements.txt`, открой терминал в папке проекта и выполни:
   ```bash
   pip install -r requirements.txt
   ```
   Если нет – установи вручную нужные пакеты:
   ```bash
   pip install openai python-dotenv openpyxl
   ```

5. **Запусти проект**  
   Теперь, когда все готово, в терминале запусти файл, который занимается приемом почты:
   ```bash
   python mail_receiver.py
   ```
   Скрипт начнет проверять входящие письма (каждые 60 секунд). Если придет письмо от поставщика, LLM-агент попробует извлечь данные, отправит уточняющий вопрос (если нужно) и сохранит итоговые данные в `suppliers_data.xlsx`.

6. **Остановка скрипта**  
   Чтобы остановить работу, просто нажми `Ctrl+C` в терминале. Все накопленные данные при этом сохранятся в Excel-файл.

---

Если что-то пойдет не так, загляни в консоль – там появятся сообщения об ошибках. Удачи, и пусть твои поставщики отвечают оперативно!))))))))))))))))))

============================================================================================================================================================



- **`agent_logic.py`**: здесь сосредоточены все классы и функции, связанные с LLM-логикой (SupplierLLMAgent), хранением данных (SupplierDataManager), отправкой писем (YandexEmailSender) и сохранением в Excel (save_supplier_data_to_excel).  
- **`mail_receiver.py`**: тут мы занимаемся приёмом почты (через IMAP), обращаемся к методам из `agent_logic.py`, и запускаем бесконечный цикл, обрабатывающий новые письма.

### Структура проекта

```
your_project/
  ├─ .env
  ├─ agent_logic.py
  ├─ mail_receiver.py
  ├─ requirements.txt (опционально)
  └─ ...
```

## `agent_logic.py`
- Переменные окружения загружаются через dotenv (load_dotenv()), инициализируя openai.api_key.
- В файле описаны три основных класса:
1. SupplierLLMAgent: методы для парсинга текста поставщика и генерации уточняющих вопросов.
2. SupplierDataManager: хранение данных по каждому поставщику.
3. YandexEmailSender: отправка писем по SMTP (логин, пароль, сервер берутся из .env).
4. В конце есть функция save_supplier_data_to_excel, которая сохраняет весь накопленный словарь data (email → поля товара) в Excel.

## `mail_receiver.py`

- Загружается `.env` ещё раз (`load_dotenv()`), чтобы переменные окружения были доступны.
- **YandexEmailReceiver** работает через `IMAP4_SSL` и ищет письма с флагом `UNSEEN`.  
- В `main()` мы:
  1. Создаём объекты (LLM-агент, хранитель данных, приёмник писем, отправитель писем).
  2. В бесконечном цикле получаем список новых писем, обрабатываем их.  
  3. С помощью `llm_agent.parse_supplier_answer` извлекаем данные.  
  4. Сохраняем в `data_manager`.  
  5. Если данные **все** есть → сохраняем в Excel, пишем поставщику «всё ок».  
  6. Если не все → генерируем уточняющий вопрос и отсылаем обратно.

# Инструкция по запуску

1. **Установите зависимости**  
   Создайте (опционально) файл `requirements.txt` со списком пакетов:
   ```
   openai
   python-dotenv
   openpyxl
   ```
   Затем выполните:
   ```bash
   pip install -r requirements.txt
   ```
   или установите пакеты вручную:
   ```bash
   pip install openai python-dotenv openpyxl
   ```

2. **Создайте файл `.env`** в корне проекта (рядом с `mail_receiver.py` и `agent_logic.py`). Пример содержимого:
   ```bash
   OPENAI_API_KEY=your_openai_api_key
   YANDEX_EMAIL=your_yandex_email@yandex.ru
   YANDEX_PASSWORD=your_yandex_password
   IMAP_SERVER=imap.yandex.ru
   SMTP_SERVER=smtp.yandex.ru
   SMTP_PORT=587
   ```
   Подставьте реальные значения!

3. **Запустите скрипт `mail_receiver.py`:**
   ```bash
   python mail_receiver.py
   ```
   Скрипт будет каждые 60 секунд проверять входящие письма в Яндексе. Если придёт письмо от поставщика, LLM-агент постарается извлечь нужные параметры товара и при необходимости отправит уточняющие вопросы в ответ.

4. **Остановить скрипт** можно, нажав `Ctrl + C`. При остановке текущие данные будут сохранены в `suppliers_data.xlsx`.

---

## Дополнительные рекомендации

- **Продвинутые настройки:** 
  - Если нужно, вы можете дополнить `.env` другими переменными (например, `EXCEL_FILENAME=...` и т.д.).
  - Можно добавить логирование (через `logging`) вместо простых `print()`, чтобы писать логи в файл.
- **Развёртывание на сервере:**
  - Запустите скрипт в фоне (через `tmux`, `screen`, `nohup` или systemd). Например:
    ```bash
    nohup python mail_receiver.py > log.txt 2>&1 &
    ```
- **Хранение контекста:** 
  - Если вам важно вести **многократную** переписку с одним поставщиком, продумайте идентификацию цепочки (ID письма, «In-Reply-To» и т.д.), а также хранение истории вопросов-ответов (не только итоговых полей).
- **Безопасность:** 
  - Файл `.env` с паролями **не** должен попадать в публичные репозитории. Добавьте его в `.gitignore`.

Таким образом, у вас получатся два файла (`agent_logic.py` и `mail_receiver.py`), которые удобно поддерживать и расширять, а все ключи и логины будут лежать в `.env`, не светясь в репозитории. Удачи!
