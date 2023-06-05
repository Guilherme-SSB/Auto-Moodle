import os
import platform
from email.mime.multipart import MIMEMultipart

import win32com.client as win32
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

import google_helper as gh


def clean_screen():
    """
    Limpa o terminal.
    """
    os.system('cls || clear')


def iniciate_chromedriver() -> webdriver:
    options = webdriver.ChromeOptions()
    options.add_argument("--incognito")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    os.system('cls || clear')
    return driver


def iniciate_chromedriver():
    if platform.system() == 'Windows':
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        # options.add_argument('--headless')
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    else:
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--ignore-ssl-errors')
        chrome_prefs = {}
        chrome_options.experimental_options["prefs"] = chrome_prefs
        chrome_prefs["profile.default_content_settings"] = {"images": 2}
        driver = webdriver.Chrome(options=chrome_options)

    driver.maximize_window()
    return driver


def find_element_by_xpath(driver: webdriver, xpath: str):
    """
    Busca um elemento pelo xpath, com uma tolerância de 5 segundos.
    
    Parâmetros
    ----------
    driver: objeto webdriver do Chrome.
    xpath: string com o xpath do elemento.
    
    Retorna
    -------
    O elemento que foi encontrado.
    """
    return WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, xpath)))


def check_exists_by_xpath(driver, xpath) -> bool:
    """
    Se o elemento buscado existir, retorna True, caso contrário, retorna False.
    
    Parâmetros
    ----------
    driver: objeto webdriver do Chrome.
    xpath: string com o xpath do elemento.

    Retorna
    -------
    True se o elemento existir, False caso contrário.
    """
    try:
        find_element_by_xpath(driver, xpath)
    except TimeoutException:
        return False
    return True


def mandar_email(to, subject, message):
    msg = MIMEMultipart()
    msg['To'] = to
    msg['Subject'] = subject
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = msg['To'] 
    mail.Subject = msg['Subject']
    mail.HtmlBody = message
    mail.Display(False)
    mail.Send()

def create_google_events(df_tarefas_to_do, df_tarefas_done):
    svc = gh.create_service()

    events_todo = []
    for index, row in df_tarefas_to_do.iterrows():
        event = {
            'summary': row['NOME'],
            'location': 'Moodle',
            'description': f'Matéria: {row["MATÉRIA"]}\nStatus: {row["STATUS"]}\nLink: {row["LINK"]}',
            'start': {
                'dateTime': row['DATA ENTREGA'].strftime('%Y-%m-%dT%H:%M:%S'),
                'timeZone': 'America/Sao_Paulo',
            },
            'end': {
                'dateTime': row['DATA ENTREGA'].strftime('%Y-%m-%dT%H:%M:%S'),
                'timeZone': 'America/Sao_Paulo',
            },
            'reminders': {
                'useDefault': False,
                'overrides': [
                    {'method': 'popup', 'minutes': 180},
                ],
            },
            'colorId': 11,
        }
        events_todo.append(event)
    
    for event in events_todo:
        existing_events = svc.events().list(calendarId='primary', q=event['summary']).execute()['items']

        if len(existing_events) == 0:
            try:
                ev = svc.events().insert(calendarId='primary', body=event).execute()
                print(f'Event created: {ev.get("htmlLink")}')
            except Exception as e:
                print(e)
                print(f'Failed to create event: {event["summary"]}')
        else:
            try:
                ev = svc.events().update(calendarId='primary', eventId=existing_events[0]['id'], body=event).execute()
                print(f'Event updated: {event["summary"]}')
            except Exception as e:
                print(e)
                print(f'Failed to update event: {event["summary"]}')

    events_done = []
    for index, row in df_tarefas_done.iterrows():
        event = {
            'summary': row['NOME'],
            'location': 'Moodle',
            'description': f'Matéria: {row["MATÉRIA"]}\nStatus: {row["STATUS"]}\nLink: {row["LINK"]}',
            'start': {
                'dateTime': row['DATA ENTREGA'].strftime('%Y-%m-%dT%H:%M:%S'),
                'timeZone': 'America/Sao_Paulo',
            },
            'end': {
                'dateTime': row['DATA ENTREGA'].strftime('%Y-%m-%dT%H:%M:%S'),
                'timeZone': 'America/Sao_Paulo',
            },
            'reminders': {
                'useDefault': False,
                'overrides': [],
            },
            'colorId': 10,
        }
        events_done.append(event)

    for event in events_done:
        existing_events = svc.events().list(calendarId='primary', q=event['summary']).execute()['items']

        if len(existing_events) > 0:
            try:
                ev = svc.events().update(calendarId='primary', eventId=existing_events[0]['id'], body=event).execute()
                print(f'Event updated: {event["summary"]}')
            except Exception as e:
                print(e)
                print(f'Failed to update event: {event["summary"]}')