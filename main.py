import logging
import re
from pathlib import Path
import requests
from bs4 import BeautifulSoup
import cursor


class WrongPasswordError(Exception):
    pass


logging.basicConfig(filename='log_file.txt',
                    filemode='w',
                    format='%(message)s',
                    level='DEBUG')
headers = {'User-Agent': 'Mozilla/5.0'}

login_url = 'http://www.edumich.gob.mx/sigem_tel/index/1/'
page_with_tables = "http://www.edumich.gob.mx/sigem_tel/sisat_registro_2223/" \
                   "3c2a3acc3ebccb7e62352756b14fd812b6913fe1/I2223/"

Path(f"tables").mkdir(parents=True, exist_ok=True)

cursor.hide()


class School:
    table_count = int(0)
    table_count_check = int(0)

    def __init__(self, credentials):
        self.login, self.password = credentials.split(':')
        self.name = self.login.split('_')[0]

    def get_tables(self):
        with requests.session() as session:
            # print(f"Fetching tables for the school {self.name}...")
            response = session.post(login_url, data={'inputEmail': '',
                                                     'inputPassword': self.password, 'grabar': 'si'})
            response.raise_for_status()
            # Here we got to check if we logged in successfully, else raise WrongPasswordError
            # with open(f'{self.name}_main.html', 'w') as main_file:
            #     main_file.write(response.text)

            response = session.get(page_with_tables, headers=headers)
            response.raise_for_status()
            # with open(f'{self.name}_tables.html', 'w') as tables_file:
            #     tables_file.write(response.text)

            # Now we need to get urls of tables, they always have titles "Concentrado de Información"
            soup = BeautifulSoup(response.content, 'html.parser')
            number_of_tables = len(soup.find_all(title='Concentrado de Información'))
            print(f"Found {number_of_tables} tables for the school {self.name}.")
            School.table_count += len(soup.find_all(title='Concentrado de Información'))
            for count, table_page_element in enumerate(soup.find_all(title="Concentrado de Información")):
                table_url = table_page_element.get('href')

                table_page = session.get(table_url, headers=headers)
                table_page.raise_for_status()

                table_page_soup = BeautifulSoup(table_page.content, 'html.parser')
                download_table_url = table_page_soup.find(class_="btn btn-warning btn-sm").get('href')
                print(f"Downloading table {count + 1} of {number_of_tables}...", end='')
                print('\r', end='')
                # print(f"Downloading table {count + 1} from {download_table_url}...")
                table_file = session.get(download_table_url, allow_redirects=True, stream=True)
                file_name = re.search(f'^.*="(.+)"$', table_file.headers['Content-Disposition']).group(1)
                # print(f"Saving table {count + 1} in the file {self.name}_{file_name}...")
                with open(f"tables/{self.name}_{file_name}", 'wb') as file:
                    file.write(table_file.content)
                    School.table_count_check += 1


print("Press Ctrl + C to exit.")

list_of_logins = list()
with open("logins.txt", 'r') as file:
    for line in file:
        list_of_logins.append(line.strip())

for login in list_of_logins:
    School(login).get_tables()

print(f"\nFinished! Saved {School.table_count_check} of {School.table_count} tables.")