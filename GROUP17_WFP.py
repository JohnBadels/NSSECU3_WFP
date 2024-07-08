import os
import subprocess
import csv
import json
from datetime import datetime

def display_menu():
    print("\n---[ S h e l l  P a r s e r ]---")
    print("[ 1 ] Shortcut Parser")
    print("[ 2 ] Jump List Parser")
    print("[ 0 ] Exit")
    choice = input("Select: ")
    return choice

def list_shortcut_files():
    shortcut_files = []

    # Get the user's Desktop path
    desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    start_menu_paths = [
        os.path.join(os.environ['PROGRAMDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
        os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs')
    ]
    taskbar_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Internet Explorer', 'Quick Launch', 'User Pinned', 'TaskBar')
    recent_items_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Recent')
    
    # List shortcut files on Desktop
    if os.path.exists(desktop_path):
        print("\nShortcut files on Desktop:")
        for item in os.listdir(desktop_path):
            if item.endswith('.lnk'):
                file_path = os.path.join(desktop_path, item)
                print(file_path)
                shortcut_files.append(file_path)
    else:
        print("Desktop path does not exist.")
    
    # List shortcut files in Start Menu paths
    for path in start_menu_paths:
        if os.path.exists(path):
            print(f"\nShortcut files in Start Menu ({path}):")
            for root, dirs, files in os.walk(path):
                for file in files:
                    if file.endswith('.lnk'):
                        file_path = os.path.join(root, file)
                        print(file_path)
                        shortcut_files.append(file_path)
        else:
            print(f"Start Menu path does not exist: {path}")
    
    # List shortcut files in Taskbar
    if os.path.exists(taskbar_path):
        print("\nShortcut files in Taskbar:")
        for item in os.listdir(taskbar_path):
            if item.endswith('.lnk'):
                file_path = os.path.join(taskbar_path, item)
                print(file_path)
                shortcut_files.append(file_path)
    else:
        print("Taskbar path does not exist.")
    
    # List shortcut files in Recent Items
    if os.path.exists(recent_items_path):
        print("\nShortcut files in Recent Items:")
        for item in os.listdir(recent_items_path):
            if item.endswith('.lnk'):
                file_path = os.path.join(recent_items_path, item)
                print(file_path)
                shortcut_files.append(file_path)
    else:
        print("Recent Items path does not exist.")
    
    return shortcut_files

def parse_shortcut():
    shortcut_files = list_shortcut_files()
    file_path = input("\nEnter the path of the .LNK file: ")
    
    if not os.path.exists(file_path):
        print("File does not exist.")
        return

    try:
        lecmd_path = "C:\\Users\\geony\\OneDrive\\Desktop\\net6\\LECmd.exe"
        output_dir = os.path.join(os.path.dirname(__file__), 'parsed_data')
        os.makedirs(output_dir, exist_ok=True)

        subprocess.run([lecmd_path, '-f', file_path, '--csv', output_dir], check=True)
        
        csv_file_path = find_latest_csv(output_dir)
        if csv_file_path:
            print(f"CSV file saved: {csv_file_path}")
            convert_csv_to_json(csv_file_path)
        else:
            print("No CSV file generated.")

    except Exception as e:
        print(f"An error occurred while parsing the shortcut: {e}")

def list_jump_list_files():
    jump_list_files = []

    jump_list_paths = [
        os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Recent', 'AutomaticDestinations'),
        os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Recent', 'CustomDestinations')
    ]

    for path in jump_list_paths:
        if os.path.exists(path):
            print(f"\nJump List files in ({path}):")
            for item in os.listdir(path):
                file_path = os.path.join(path, item)
                print(file_path)
                jump_list_files.append(file_path)
        else:
            print(f"Jump List path does not exist: {path}")

    return jump_list_files

def parse_jump_list():
    jump_list_files = list_jump_list_files()
    file_path = input("\nEnter the path of the Jump List file: ")

    if not os.path.exists(file_path):
        print("File does not exist.")
        return

    try:
        jlecmd_path = "C:\\Users\\geony\\OneDrive\\Desktop\\net6\\JLECmd.exe"
        output_dir = os.path.join(os.path.dirname(__file__), 'parsed_data')
        os.makedirs(output_dir, exist_ok=True)

        subprocess.run([jlecmd_path, '-f', file_path, '--csv', output_dir], check=True)
        
        csv_file_path = find_latest_csv(output_dir)
        if csv_file_path:
            print(f"CSV file saved: {csv_file_path}")
            convert_csv_to_json(csv_file_path)
        else:
            print("No CSV file generated.")

    except Exception as e:
        print(f"An error occurred while parsing the jump list: {e}")

def find_latest_csv(directory):
    csv_files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.csv')]
    if not csv_files:
        return None
    latest_file = max(csv_files, key=os.path.getctime)
    return latest_file

def convert_csv_to_json(csv_file_path):
    json_file_path = csv_file_path.replace('.csv', '.json')

    try:
        with open(csv_file_path, 'r', encoding='utf-8') as csv_file:
            reader = csv.DictReader(csv_file)
            rows = list(reader)
        
        with open(json_file_path, 'w', encoding='utf-8') as json_file:
            json.dump(rows, json_file, indent=4)

        print(f"JSON file saved: {json_file_path}")

    except Exception as e:
        print(f"An error occurred while converting CSV to JSON: {e}")

def main():
    while True:
        choice = display_menu()
        if choice == '1':
            parse_shortcut()
        elif choice == '2':
            parse_jump_list()
        elif choice == '0':
            print("Goodbye!")
            break
        else:
            print("Invalid choice. Please try again.")

if __name__ == "__main__":
    main()
