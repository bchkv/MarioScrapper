import logging
from pathlib import Path
import cursor
import os
import atexit
from schools import School
from tables import check_table, proceed_tables, download_tables

logging.basicConfig(filename='log_file.txt',
                    filemode='w',
                    format='%(message)s',
                    level='DEBUG')

if __name__ == "__main__":
    Path(f"tables").mkdir(parents=True, exist_ok=True)
    cursor.hide()
    atexit.register(cursor.show)
    # os.system('clear')
    print("Press Ctrl + C to quit the program\n")
    answer = input("Do you want to download tables (d) or just proceed them (p)? ")
    if answer == 'd':
        download_tables()
    elif answer not in ['d', 'p']:
        print("Wrong input!")
        exit(0)
    min_inut, max_input = input("Enter min and max values to color row with average scores: ").split(' ')
    print(f"Coloring values between {min_inut} and {max_input}...")
    proceed_tables(int(min_inut), int(max_input))
