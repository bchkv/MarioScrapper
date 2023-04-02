import requests
import re
from bs4 import BeautifulSoup
from pathlib import Path
from globals import tables_data, faulty_tables

headers = {'User-Agent': 'Mozilla/5.0'}
login_url = 'http://www.edumich.gob.mx/sigem_tel/index/1/'
page_with_tables = "http://www.edumich.gob.mx/sigem_tel/sisat_registro_2223/" \
                   "3c2a3acc3ebccb7e62352756b14fd812b6913fe1/I2223/"


class School:
    table_count = int(0)
    table_count_check = int(0)

    def __init__(self, credentials):
        self.login, self.password = credentials.split(':')
        self.name = self.login.split('_')[0]

    def _login_to_school(self, session):
        response = session.post(login_url, data={'inputEmail': '',
                                                 'inputPassword': self.password, 'grabar': 'si'})
        response.raise_for_status()
        if response.url == login_url:
            print(f"Could not login to the school {self.login} with the password '{self.password}'!")
            print("Check credentials for that school and start again!")
            exit(1)

    def _get_tables_urls(self, session):
        response = session.get(page_with_tables, headers=headers)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        tables_page_elements = soup.find_all(title="Concentrado de Información")
        return [table_page_element.get('href') for table_page_element in tables_page_elements]

    def _download_table(self, session, table_url, count, number_of_tables):
        table_page = session.get(table_url, headers=headers)
        table_page.raise_for_status()

        table_page_soup = BeautifulSoup(table_page.content, 'html.parser')
        download_table_url = table_page_soup.find(class_="btn btn-warning btn-sm").get('href')

        print(f"Descargando tabla {count + 1} de {number_of_tables}...", end='')
        print('\r', end='')
        table_file = session.get(download_table_url, allow_redirects=True, stream=True)

        correct_content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        if table_file.headers['Content-Type'] != correct_content_type:
            print(f'''Something wrong with the file!\n
                      Log in manually in the school {self.name} and then check the link: {download_table_url}\n
                      You can download this file manually and put it in the folder 'Tables''')
            file_name = f"Unknown table {count + 1} from {self.name}"
            faulty_tables.append(file_name)
            tables_data[file_name] = {'school': self.name, 'course': None, 'group': None,
                                      'download_url': download_table_url}
            return file_name

        file_name = re.search(f'^.*="(.+)"$', table_file.headers['Content-Disposition']).group(1)
        file_path = Path(f"tables/{self.name}_{file_name}")
        with open(file_path, 'wb') as file:
            # Download the file in chunks and update the progress bar
            for chunk in table_file.iter_content(chunk_size=1024):
                if chunk:
                    file.write(chunk)
        School.table_count_check += 1

        return file_name

    def get_tables(self):
        """
        Downloads and saves tables for the school.
        """
        with requests.session() as session:
            self._login_to_school(session)
            tables_urls = self._get_tables_urls(session)
            number_of_tables = len(tables_urls)
            print(f"Encontradas {number_of_tables} tablas para la escuela '{self.name}'.")

            # Iterate through the table URLs and call the _download_table method
            for count, table_url in enumerate(tables_urls):
                file_name = self._download_table(session, table_url, count, number_of_tables)
                class_data = re.search(r'"(\d)([a-zA-Z])"', file_name)
                if class_data:
                    course_year = class_data.group(1)
                    group = class_data.group(2)
                else:
                    course_year = None
                    group = None
                    faulty_tables.append(file_name)

                tables_data[file_name] = {'school': self.name, 'course': course_year, 'group': group,
                                          'download_url': table_url}
