import logging
import sys
from pathlib import Path
import cursor
import os
import atexit
from tables import process_tables, download_tables
from utils import set_up_scripts

logging.basicConfig(filename='log_file.txt',
                    filemode='w',
                    format='%(message)s',
                    level='DEBUG')

if __name__ == "__main__":
    set_up_scripts()

    # Mode can be "update" or "process"
    if len(sys.argv) > 2:
        raise ValueError("Only one argument is allowed.")
    elif len(sys.argv) == 2:
        if sys.argv[1] in ["update", "process"]:
            mode = sys.argv[1]
        else:
            raise ValueError("Only 'update' or 'process' arguments are allowed")
    else:
        mode = "update"

    # Create 'tables' and 'data' directories if not exist
    Path(f"tables").mkdir(parents=True, exist_ok=True)
    Path(f"data").mkdir(parents=True, exist_ok=True)
    cursor.hide()
    # Run 'cursor.show' when the script exits
    atexit.register(cursor.show)
    os.system('clear')
    print("Press Ctrl + C to quit the program\n")
    if mode == 'update':
        download_tables()

    min_input, max_input = input("Ingrese valores mínimo y máximo: ").split(' ')
    print(f"Coloreando valores entre {min_input} y {max_input}...")
    process_tables(int(min_input), int(max_input))
