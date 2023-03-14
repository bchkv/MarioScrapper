import logging
import re
from pathlib import Path
import requests
from bs4 import BeautifulSoup
import cursor
import pandas
import os
import atexit


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
atexit.register(cursor.show)

# For each xlsx file we need to memorise: school, file name, course year, school group
tables_data = dict()
faulty_tables = list()


class School:
    table_count = int(0)
    table_count_check = int(0)

    def __init__(self, credentials):
        self.login, self.password = credentials.split(':')
        self.name = self.login.split('_')[0]

    def get_tables(self):
        with requests.session() as session:
            response = session.post(login_url, data={'inputEmail': '',
                                                     'inputPassword': self.password, 'grabar': 'si'})
            response.raise_for_status()
            if response.url == login_url:
                print(f"Could not login to the school {self.login} with the password '{self.password}'!")
                print("Check credentials for that school and start again!")
                exit(1)

            response = session.get(page_with_tables, headers=headers)
            response.raise_for_status()

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
                table_file = session.get(download_table_url, allow_redirects=True, stream=True)

                # If the file is not an Excel-table, show error message and continue with the next file
                correct_content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                if table_file.headers['Content-Type'] != correct_content_type:
                    print(f'''Something wrong with the file!\n
                              Log in manually in the school {self.name} and then check the link: {download_table_url}\n
                              You can download this file manually and put it in the folder 'Tables''')
                    file_name = f"Unknown table {count + 1} from {self.name}"
                    faulty_tables.append(file_name)
                    tables_data[file_name] = {'school': self.name, 'course': None, 'group': None,
                                              'download_url': download_table_url}
                    continue

                file_name = re.search(f'^.*="(.+)"$', table_file.headers['Content-Disposition']).group(1)
                with open(f"tables/{self.name}_{file_name}", 'wb') as file:
                    file.write(table_file.content)
                    School.table_count_check += 1

                class_data = re.search(r'"(\d)([a-zA-Z])"', file_name)
                if class_data:
                    course_year = class_data.group(1)
                    group = class_data.group(2)
                else:
                    course_year = None
                    group = None
                    faulty_tables.append(file_name)

                tables_data[file_name] = {'school': self.name, 'course': course_year, 'group': group,
                                          'download_url': download_table_url}


def proceed_tables():
    pass

# os.system('clear')
print("Press Ctrl + C to quit the program\n")

list_of_logins = list()
with open("logins.txt", 'r') as file:
    for line in file:
        list_of_logins.append(line.strip())

for login in list_of_logins:
    School(login).get_tables()

print(f"\nFinished! Saved {School.table_count_check} of {School.table_count} tables.\n")

for table in faulty_tables:
    print(f'Something went wrong with the table {table} from the school {tables_data[table]["school"]}.\n'
          f"Please, check it manually: log in into the school {tables_data[table]['school']} and visit "
          f"{tables_data[table]['download_url']} to download file\n")

if faulty_tables:
    while True:
        print('You had a problem with one or more table files. You can resolve it manually or ignore it.\n'
              'Print "y" once you are ready: ', end='')
        char = input()
        if char == 'y':
            break

proceed_tables()
