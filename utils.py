import os

desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

update_script = '''
    #!/bin/bash
    cd ~/Desktop/MarioProject
    clear
    python3 main.py --update
    '''
process_script = '''
    #!/bin/bash
    cd ~/Desktop/MarioProject
    clear
    python3 main.py --process
    '''


def set_up_scripts():
    """
        Sets up all scripts to run the program
    :return:
        None
    """
    create_script(update_script)
    create_script(process_script)


def create_script(script_text, name):
    """
        Creates script for updating tables in the Desktop.
    :arg
            script_text (str): The script itself
            name (str): Name of the script file *.sh
    :return:
        None
    """

    script_path = os.path.join(desktop_path, name)

    # If file exists, do nothing
    if os.path.exists(script_path):
        return

    with open(script_path, 'w') as script_file:
        script_file.write(script_text)

    # Sets the executable permissions
    os.chmod(script_path, 0o755)
