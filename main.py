import csv
import re
from bs4 import BeautifulSoup
from imap_tools import MailBox
from tqdm import tqdm
from datetime import datetime
from backoffice import backoffice
import config

def parse_emails(data_rows: list = None) -> list:
    """
    Парсит и обрабатывает электронные письма из указанного почтового ящика.

    Args:
        data_rows (list, optional): Список для хранения данных из писем. По умолчанию None, в этом случае инициализируется пустым списком.

    Returns:
        list: Список с данными, извлеченными из писем.
    """
    if data_rows is None:
        data_rows = []

    with MailBox(config.mail_server).login(config.mail_login, config.mail_password) as mailbox:
        # print(mailbox.folder.list())
        mailbox.folder.set(config.mailbox)
        for msg in tqdm(mailbox.fetch()):
            if msg.subject == "Siberian Wellness: новая регистрация в вашей команде!":
                soup = BeautifulSoup(msg.html, 'html.parser')
                texts = [text.get_text(strip=True) for text in soup.find_all(["h1", "h2", "p", "a"])]
                data_row = [str(msg.date)[:10]] + texts
                data_rows.append(data_row)
    return data_rows

def extract_tier(text_list: list) -> str:
    """
    Извлекает уровень (тип) регистрации из списка текстов.

    Args:
        text_list (list): Список строк с текстом письма.

    Returns:
        str: Уровень регистрации, если найден, иначе None.
    """
    text = ' '.join(text_list)
    if "Бизнес-Партнер" in text:
        return "Бизнес-Партнер"
    elif "Привилегированный клиент" in text:
        return "Привилегированный клиент"
    return None

def save_to_csv(data_registration: list, my_team) -> None:
    """
    Сохраняет данные регистрации в CSV файл.

    Args:
        data_registration (list): Список данных регистрации.
        my_team (pandas.DataFrame): DataFrame с информацией о команде.

    Returns:
        None
    """
    csv_filename = f'{config.mail_login.split("@")[0]}.csv'
    with open(csv_filename, 'w', newline='', encoding='utf-8-sig') as csv_file:
        writer = csv.writer(csv_file, delimiter=';')
        writer.writerow(['Дата', 'Имя', 'Телефон', 'Почта', 'Регистрационный номер', 'Тип', 'Примечание', 'WhatsApp'])

        for row in tqdm(data_registration):
            index = next(
                (idx for idx, item in enumerate(row) if "Регистрационный номер:" in item or "Номер Соглашения:" in item),
                None
            )
            if index is not None:
                registration_numbers = row[index].replace("Регистрационный номер:", "").replace("Номер Соглашения:", "")
                data = str(row)

                phone_match = re.search(r'\d{11}|375\d{9}', data)
                phones = phone_match.group() if phone_match else "---"

                email_match = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', data)
                emails = email_match[0] if email_match else "---"

                name = row[index - 1]
                tier = extract_tier(row)
                info = ""

                if my_team["ФИО"].tolist().count(name) > 1:
                    info = "Повторная регистрация"

                if my_team["E-mail"].tolist().count(emails) > 1:
                    info = "Повторная регистрация"

                if my_team["Телефон"].tolist().count(int(phones)) > 1:
                    info = "Повторная регистрация"

                if not my_team["Регистрационный номер"].tolist().count(int(registration_numbers)):
                    info += " Закрыт"
                else:
                    noo_value = my_team.loc[my_team["Регистрационный номер"] == int(registration_numbers), "НОО"].values[0]
                    if bool(int(noo_value)):
                        info += f" {noo_value}"

                if datetime.strptime(row[0], "%Y-%m-%d").date() == datetime.now().date():
                    info = ""

                whatsapp = f'=ГИПЕРССЫЛКА("https://web.whatsapp.com/send/?phone={int(phones)}";"Написать")'

                new_list = [row[0], name, phones, emails, registration_numbers, tier, info, whatsapp]
                writer.writerow(new_list)

if __name__ == '__main__':
    my_team = backoffice()
    new_registration = parse_emails()
    save_to_csv(new_registration, my_team)
