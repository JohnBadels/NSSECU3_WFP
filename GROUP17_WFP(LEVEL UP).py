import os
import subprocess
import csv
import json
from datetime import datetime

def display_menu():
    print("\n---[ S h e l l  P a r s e r ]---")
    print("[ 1 ] Shortcut Parser")
    print("[ 2 ] Jump List Parser")
    print("[ 3 ] Prefetch Parser")
    print("[ 4 ] Shellbags Parser")
    print("[ 0 ] Exit")
    choice = input("Select: ")
    return choice

def list_files(folder_paths, extension):
    files = []
    for path in folder_paths:
        if os.path.exists(path):
            print(f"\nFiles in ({path}):")
            for root, _, filenames in os.walk(path):
                for filename in filenames:
                    if filename.endswith(extension):
                        file_path = os.path.join(root, filename)
                        print(file_path)
                        files.append(file_path)
        else:
            print(f"Path does not exist: {path}")
    return files

def list_shortcut_files():
    desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    start_menu_paths = [
        os.path.join(os.environ['PROGRAMDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
        os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs')
    ]
    taskbar_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Internet Explorer', 'Quick Launch', 'User Pinned', 'TaskBar')
    recent_items_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Recent')
    
    return list_files([desktop_path] + start_menu_paths + [taskbar_path, recent_items_path], '.lnk')

def list_jump_list_files():
    jump_list_paths = [
        os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Recent', 'AutomaticDestinations'),
        os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Recent', 'CustomDestinations')
    ]
    return list_files(jump_list_paths, '')

def list_prefetch_files():
    prefetch_path = os.path.join(os.environ['SYSTEMROOT'], 'Prefetch')
    return list_files([prefetch_path], '.pf')

def list_shellbags_files():
    users_dir = os.path.join(os.environ['SYSTEMDRIVE'] + '\\', 'Users')

    ntuser_files = []
    usrclass_files = []

    for user_folder in os.listdir(users_dir):
        user_path = os.path.join(users_dir, user_folder)
        if os.path.isdir(user_path):
            ntuser_path = os.path.join(user_path, 'NTUSER.DAT')
            if os.path.exists(ntuser_path):
                ntuser_files.append(ntuser_path)
            
            usrclass_path = os.path.join(user_path, 'AppData', 'Local', 'Microsoft', 'Windows', 'UsrClass.dat')
            if os.path.exists(usrclass_path):
                usrclass_files.append(usrclass_path)
    
    return ntuser_files, usrclass_files

def parse_file(tool_path, file_path, is_shellbags=False):
    output_dir = os.path.join(os.path.dirname(__file__), 'parsed_data')
    os.makedirs(output_dir, exist_ok=True)

    try:
        if is_shellbags:
            command = [tool_path, '-d', os.path.dirname(file_path), '--csv', output_dir]
        else:
            command = [tool_path, '-f', file_path, '--csv', output_dir]

        result = subprocess.run(command, check=True, capture_output=True, text=True)
        print(result.stdout)
        print(result.stderr)

        csv_file_path = find_latest_csv(output_dir)
        if csv_file_path:
            print(f"CSV file saved: {csv_file_path}")
            convert_csv_to_json(csv_file_path)
        else:
            print("No CSV file generated.")

    except Exception as e:
        print(f"An error occurred while parsing the file: {e}")

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
    tool_paths = {
        '1': "C:\\Users\\geony\\OneDrive\\Desktop\\net6\\LECmd.exe",
        '2': "C:\\Users\\geony\\OneDrive\\Desktop\\net6\\JLECmd.exe",
        '3': "C:\\Users\\geony\\OneDrive\\Desktop\\net6\\PECmd.exe",
        '4': "C:\\Users\\geony\\OneDrive\\Desktop\\net6\\SBECmd.exe"
    }

    list_functions = {
        '1': list_shortcut_files,
        '2': list_jump_list_files,
        '3': list_prefetch_files,
        '4': list_shellbags_files
    }

    while True:
        choice = display_menu()
        if choice in tool_paths:
            if choice == '4':
                ntuser_files, usrclass_files = list_functions[choice]()
                print("\nNTUSER.DAT files found:")
                for file in ntuser_files:
                    print(file)
                print("\nUsrClass.dat files found:")
                for file in usrclass_files:
                    print(file)
                
                file_path = input("\nEnter the path of the NTUSER.DAT or UsrClass.dat file (or press Enter to return to main menu): ")
                if file_path == "":
                    continue

                if not os.path.exists(file_path):
                    print("File does not exist.")
                    continue

                parse_file(tool_paths[choice], file_path, is_shellbags=True)
            else:
                files = list_functions[choice]()
                file_path = input("\nEnter the path of the file or directory (or press Enter to return to main menu): ")
                if file_path == "":
                    continue

                if not os.path.exists(file_path):
                    print("File or directory does not exist.")
                    continue

                parse_file(tool_paths[choice], file_path)
        elif choice == '0':
            print("Goodbye!")
            break
        else:
            print("Invalid choice. Please try again.")

        input("\nPress Enter to return to the main menu...")

if __name__ == "__main__":
    main()
