import pandas as pd
from pathlib import Path
from globals import faulty_tables
import re
from globals import tables_data
from schools import School


def categorize_tables(tables_names_list):
    """
    Categorizes tables by school name, year, and class letter.

    Args:
        tables_names_list (list): A list of table names.

    Returns:
        dict: A dictionary with categorized tables.
    """
    tables_dict = dict()

    for table_name in tables_names_list:
        if table_name == ".DS_Store" or not re.search(r"^.*\.xlsx$", table_name):
            continue

        if check_table(table_name):
            school_name, year, letter = extract_table_info(table_name)
            add_table_to_dict(tables_dict, school_name, year, letter, table_name)
        else:
            add_faulty_table(tables_dict, table_name)

    return tables_dict


def extract_table_info(table_name):
    """
    Extracts school name, year, and class letter from a table name using regex.

    Args:
        table_name (str): The name of the table.

    Returns:
        tuple: A tuple containing school_name, year, and letter.
    """
    school_name = re.search(r"(^[0-9a-zA-Z]*)_", table_name).group(1)
    year = re.search(r'"([0-9])[A-Za-z]"', table_name).group(1)
    letter = re.search(r'"[0-9]([A-Za-z])"', table_name).group(1)
    return school_name, year, letter


def add_table_to_dict(tables_dict, school_name, year, letter, table_name):
    """
    Adds a table to the tables dictionary with the corresponding keys.

    Args:
        tables_dict (dict): The dictionary to add the table to.
        school_name (str): The name of the school.
        year (str): The year of the table.
        letter (str): The class letter of the table.
        table_name (str): The name of the table.

    Returns:
        None
    """
    if school_name not in tables_dict:
        tables_dict[school_name] = dict()
    if year not in tables_dict[school_name]:
        tables_dict[school_name][year] = dict()
    tables_dict[school_name][year][letter] = table_name


def add_faulty_table(tables_dict, table_name):
    """
    Adds a faulty table to the tables dictionary.

    Args:
        tables_dict (dict): The dictionary to add the faulty table to.
        table_name (str): The name of the faulty table.

    Returns:
        None
    """
    if 'faulty_tables' not in tables_dict:
        tables_dict['faulty_tables'] = list()
    tables_dict['faulty_tables'].append(table_name)


def proceed_tables(_min, _max):
    """
    Processes tables in the folder and categorizes them by school name, year, and class letter.
    It also colors cells in the row 'Average' with values within the range [_min, _max].

    Args:
        _min (float): The minimum value of the range to color cells.
        _max (float): The maximum value of the range to color cells.

    Returns:
        None
    """

    tables_names_list = [str(table_path) for table_path in Path("tables").iterdir()]
    tables_names_list = [table_name.replace("tables/", "") for table_name in tables_names_list]
    tables_names_list.sort()

    tables_dict = categorize_tables(tables_names_list)

    # ... rest of the function implementation ...


def download_tables():
    """
    Downloads tables from the website.

    Returns:
        None
    """
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


def check_table(table_name):
    """
    Checks if a table with the given table_name can be opened.

    Args:
        table_name (str): The name of the table to check.

    Returns:
        bool: True if the table can be opened, False otherwise.
    """
    try:
        pd.read_excel("tables/" + table_name)
    except ValueError:
        print(f"The table {table_name} seems to be faulty!")
        faulty_tables.append(table_name)
        return False
    return True
