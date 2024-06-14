import datetime
import io
import traceback
from io import BytesIO
import requests
import pandas as pd
from anticaptchaofficial.imagecaptcha import *

import config


class myDict(dict):
    def __getattr__(self, attr):
        if attr in self:
            return self[attr]
        else:
            return None


def bypass_captcha(session, captcha_key=config.captcha_key):
    solver = imagecaptcha()
    # solver.set_verbose(0)  # Установите уровень отладки на 0, чтобы отключить сообщения работы
    while True:
        try:
            solver.set_key(captcha_key)
            img = session.get("https://ru.siberianhealth.com/ru/captcha/default/")
            captcha_content = BytesIO(img.content)
            captcha_text = solver.solve_and_return_solution(
                file_path=None, body=captcha_content.read()
            )
            return captcha_text
        except:
            pass


def auth(
    user,
    url="https://ru.siberianhealth.com/ru/backoffice-new/?newStyle=yes",
    url_ajax="https://ru.siberianhealth.com/ru/controller/ajax/",
):
    payload = {
        "login": f"{user.number}",
        "pass": f"{user.password}",
        "url": url,
        "_controller": "Backoffice_Auth/submit",
        "_url": "https://ru.siberianhealth.com/ru/backoffice/auth/?url=https://ru.siberianhealth.com/ru/backoffice-new/?newStyle=yes",
    }

    with requests.Session() as session:
        while True:
            response = session.post(url=url_ajax, data=payload, allow_redirects=True)
            response_json = response.json()
            if response_json["result"]["status"] == "Denied":
                payload["captcha"] = bypass_captcha(session)
            else:
                break
        if (
            response_json["result"]["success"]
            and "Стать Бизнес-Партнером" not in session.get(url=url).text
        ):
            return session
        elif "Стать Бизнес-Партнером" in session.get(url=url).text:
            return f"{user.number} нужно стать Бизнес-Партнером"
        else:
            return f"{user.number} {response_json['result']['status']}"


def get_current_period(format):
    return datetime.now().strftime(format)


def download_csv_data(
    id, session, url_ajax="https://ru.siberianhealth.com/ru/controller/ajax/"
):
    data = {
        "filters[page]": "1",
        "filters[perPage]": "20",
        "filters[period]": f'{get_current_period("%m.%Y")}',
        "filters[fromCache]": "0",
        "filters[contract]": "",
        "filters[search]": "",
        "filters[group]": "0",
        "filters[minLO]": "",
        "filters[maxLO]": "",
        "filters[qualificationOpen]": "0",
        "filters[qualificationClosed]": "0",
        "filters[newbies]": "",
        "filters[type]": "all",
        "filters[sort][field]": "lo",
        "filters[sort][direction]": "ASC",
        "filters[club200]": "0",
        "filters[club500]": "0",
        "filters[club1000]": "0",
        "filters[firstLine]": "1",
        "filters[specialDiscount]": "",
        "filters[specialGift]": "",
        "_controller": "Backoffice_Report_Inf/download_prepare",
        "_contract": f"{id}",
        "_url": "https://ru.siberianhealth.com/ru/backoffice/report/inf/",
    }

    hash = session.post(url_ajax, data=data, allow_redirects=True)
    hash = hash.json()["result"]["hash"]
    report = session.get(
        f"https://ru.siberianhealth.com/ru/backoffice/report/inf/download/{hash}/"
    )

    # with open('Моя команда.xls', 'wb') as file:
    #     file.write(report.content)

    df = pd.read_excel(io.BytesIO(report.content))
    return df


def process_user_data(user):
    try:
        session = auth(user)
        if isinstance(session, requests.sessions.Session):
            dataTeam = download_csv_data(user.number, session)
            # Выход из цикла, если все прошло успешно
            return dataTeam
        else:
            return session

    except Exception as e:
        print("Произошла ошибка:", e)
        traceback.print_exc()
        time.sleep(60)


def backoffice(user=config.user):
    try:
        user_data = process_user_data(myDict(user))
        return user_data

    except Exception as e:
        # Блок обработки исключений
        print("Произошла ошибка:", e)
        traceback.print_exc()
        time.sleep(60)
