import csv
import re
from datetime import datetime

from bs4 import BeautifulSoup
from imap_tools import MailBox
from tqdm import tqdm

import config
from backoffice import backoffice


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

def find_index(row):
    """
    Находит индекс первого элемента в списке row, который содержит подстроку
    "Регистрационный номер:" или "Номер Соглашения:". Возвращает индекс или None, если не найдено.
    """
    search_terms = ("Регистрационный номер:", "Номер Соглашения:")

    for idx, item in enumerate(row):
        if any(term in item for term in search_terms):
            return idx
    return False

def process_message(msg):
    email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
    phone_pattern = re.compile(r'\d{11}|375\d{9}')

    if msg.subject == "Siberian Wellness: новая регистрация в вашей команде!":
        soup = BeautifulSoup(msg.html, 'lxml')
        strings = [text.get_text(strip=True) for text in soup.find_all(["h1", "h2", "p", "a"])]
        index = find_index(strings)

        # Оптимизация прохода по строкам
        emails = []
        phones = []

        for string in strings:
            email_match = email_pattern.search(string)
            if email_match:
                emails.append(email_match.group(0))

            phone_match = phone_pattern.search(string)
            if phone_match:
                phones.append(phone_match.group(0))

        if isinstance(index, int):
            data_row = {
                'Дата': str(msg.date)[:10],
                'Имя': strings[index-1],
                'Регистрационный номер': strings[index].split(':')[-1].strip(),
                'Телефон': phones[0] if phones else None,
                'Почта': emails[0] if emails else None,
                'Тип': extract_tier(strings),
                'Примечание': ""
            }
            return data_row
    return None


def fetch_emails(mail_server=config.mail_server,
                 mail_login=config.mail_login,
                 mail_password=config.mail_password,
                 mailbox_name=config.mailbox):
    with MailBox(mail_server).login(mail_login, mail_password) as mailbox:
        mailbox.folder.set(mailbox_name)
        messages = list(tqdm(mailbox.fetch(limit=1000), desc="fetch_emails"))

    return messages


def save_to_csv(data_list, csv_filename):
    headers = ['Дата', 'Имя', 'Телефон', 'Почта', 'Регистрационный номер', 'Тип', 'Примечание']

    with open(csv_filename, mode='w', newline='', encoding='utf-8-sig') as file:
        writer = csv.DictWriter(file, fieldnames=headers, delimiter=';')
        writer.writeheader()
        for data in data_list:
            writer.writerow(data)

def check_duplicates(data):
    unique = dict()
    for index, element in enumerate(tqdm(data, desc="check_duplicates")):
        for key in ['Имя', 'Телефон', 'Почта']:
            try:
                if element[key] and element[key] in unique:
                    element['Примечание'] = "Повторная регистрация"
                    data[unique[element[key]]]['Примечание'] = f"Первая регистрация"
                    break
            finally:
                unique[element[key]] = index
    return data

def process_messages(messages):
    processed_messages = []
    for msg in tqdm(messages, desc="process_messages"):
        result = process_message(msg)
        if result:
            processed_messages.append(result)
    return processed_messages


def update_my_team(list_of_dicts, df):
    for item in tqdm(list_of_dicts, desc="update_my_team"):
        reg_number = int(item["Регистрационный номер"])
        if reg_number in df["Регистрационный номер"].values:
            noo_value = df.loc[df["Регистрационный номер"] == reg_number, "НОО"].values[0]
            if bool(int(noo_value)):
                item["Примечание"] += f" {noo_value}"
        else:
            item["Примечание"] += " Закрыто"
        item["Примечание"] = item["Примечание"].replace('.', ',').strip()
    return list_of_dicts


def save_to_html(data_list, html_filename):
    headers = ['Дата', 'Имя', 'Телефон', 'Почта', 'Регистрационный номер', 'Тип', 'Примечание', 'Сообщение']

    html_content = """
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Данные</title>
        <style>
            table {
                border-collapse: collapse;
                width: 100%;
            }
            th, td {
                border: 1px solid black;
                padding: 8px;
                text-align: left;
            }
            th {
                background-color: #f2f2f2;
            }
        </style>
    </head>
    <body>
        <h1>Данные</h1>
        <table>
            <thead>
                <tr>
    """

    for header in headers:
        html_content += f"<th>{header}</th>"

    html_content += """
                </tr>
            </thead>
            <tbody>
    """

    for data in data_list:
        if data['Сообщение']:
            html_content += "<tr>"
            for header in headers:
                html_content += f"<td>{data.get(header, '')}</td>"
            html_content += "</tr>"

    html_content += """
            </tbody>
        </table>
    </body>
    </html>
    """

    with open(html_filename, 'w', encoding='utf-8') as file:
        file.write(html_content)

def add_button_to_closed(data_list):
    for data in data_list:
        if data.get('Примечание') in ['Закрыто', 'Повторная регистрация Закрыто']:
            text = f"{data['Имя'].split()[-1].strip()}, " \
                   f"добрый день! Ранее Вы регистрировались на сайте Siberian Wellness (Сибирское Здоровье). " \
                   f"Ваш аккаунт был удален в связи с отсутствием активности в течение продолжительного времени." \
                   f"Если вы захотите снова стать частью Siberian Wellness – просто зарегистрируйтесь по ссылке - " \
                   f"это бесплатно. " \
                   f"https://ru.siberianhealth.com/ru/shop/user/registration/PRIVILEGED_CLIENT/?referral=2596572021"

            url = f"https://api.whatsapp.com/send/?phone={data['Телефон']}&text={text}&type=phone_number&app_absent=1"

            data['Сообщение'] = f'<button>' \
                                f'<a href="{url}"' \
                                f'target="_blank">Отправить</a>' \
                                f'</button>'


if __name__ == '__main__':
    messages = fetch_emails()

    data_list = process_messages(messages)

    check_duplicates(data_list)

    my_team = backoffice()

    update_my_team(data_list, my_team)

    save_to_csv(data_list, f'{config.mail_login.split("@")[0]}.csv')

    add_button_to_closed(data_list)

    save_to_html(data_list, f'{config.mail_login.split("@")[0]}.html')