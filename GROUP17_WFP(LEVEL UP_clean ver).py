import os
import subprocess
import csv
import json
from datetime import datetime
import time
import sys

def display_menu():
    """
    Display the main menu of the shell parser program.
    """
    print("\n---[ S h e l l  P a r s e r ]---")
    print("[ 1 ] Shortcut Parser")
    print("[ 2 ] Jump List Parser")
    print("[ 3 ] Prefetch Parser")
    print("[ 4 ] Shellbags Parser")
    print("[ 0 ] Exit")
    choice = input("Select: ")
    return choice

def list_files(folder_paths, extension):
    """
    List files with a specific extension from given folders.

    Args:
    - folder_paths (list): List of folder paths to search.
    - extension (str): File extension to filter by.

    Returns:
    - files (dict): Dictionary of file paths categorized by folder.
    """
    files = {}
    for path in folder_paths:
        if os.path.exists(path):
            files[path] = []
            for root, _, filenames in os.walk(path):
                for filename in filenames:
                    if filename.endswith(extension):
                        file_path = os.path.join(root, filename)
                        files[path].append(file_path)
        else:
            print(f"Path does not exist: {path}")
    return files

def list_shortcut_files():
    """
    List shortcut files (.lnk) from various system locations.
    
    Returns:
    - files (dict): Dictionary of shortcut file paths categorized by folder.
    """
    desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    start_menu_paths = [
        os.path.join(os.environ['PROGRAMDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
        os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs')
    ]
    taskbar_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Internet Explorer', 'Quick Launch', 'User Pinned', 'TaskBar')
    recent_items_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Recent')
    
    return list_files([desktop_path] + start_menu_paths + [taskbar_path, recent_items_path], '.lnk')

def list_jump_list_files():
    """
    List jump list files (no extension) from recent items locations.
    
    Returns:
    - files (dict): Dictionary of jump list file paths categorized by folder.
    """
    jump_list_paths = [
        os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Recent', 'AutomaticDestinations'),
        os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Recent', 'CustomDestinations')
    ]
    return list_files(jump_list_paths, '')

def list_prefetch_files():
    """
    List prefetch files (.pf) from the Windows prefetch folder.
    
    Returns:
    - files (dict): Dictionary of prefetch file paths categorized by folder.
    """
    prefetch_path = os.path.join(os.environ['SYSTEMROOT'], 'Prefetch')
    return list_files([prefetch_path], '.pf')

def list_shellbags_files():
    """
    List NTUSER.DAT and UsrClass.dat files from user directories for Shellbags parsing.
    
    Returns:
    - ntuser_files (dict): Dictionary of NTUSER.DAT file paths categorized by folder.
    - usrclass_files (dict): Dictionary of UsrClass.dat file paths categorized by folder.
    """
    users_dir = os.path.join(os.environ['SYSTEMDRIVE'] + '\\', 'Users')

    ntuser_files = {}
    usrclass_files = {}

    for user_folder in os.listdir(users_dir):
        user_path = os.path.join(users_dir, user_folder)
        if os.path.isdir(user_path):
            ntuser_path = os.path.join(user_path, 'NTUSER.DAT')
            if os.path.exists(ntuser_path):
                ntuser_files[user_folder] = ntuser_files.get(user_folder, []) + [ntuser_path]
            
            usrclass_path = os.path.join(user_path, 'AppData', 'Local', 'Microsoft', 'Windows', 'UsrClass.dat')
            if os.path.exists(usrclass_path):
                usrclass_files[user_folder] = usrclass_files.get(user_folder, []) + [usrclass_path]
    
    return ntuser_files, usrclass_files

def display_files(files):
    """
    Display list of files found, categorized by their directories.

    Args:
    - files (dict): Dictionary of file paths categorized by folder.
    """
    for folder, file_list in files.items():
        print(f"\nFiles found in {folder}:")
        for file in file_list:
            print(file)

def parse_file(tool_path, file_path, is_shellbags=False, is_live=False):
    """
    Parse a file using the specified tool and save results to CSV and JSON.

    Args:
    - tool_path (str): Path to the parsing tool executable.
    - file_path (str): Path to the file to be parsed.
    - is_shellbags (bool): Flag indicating if parsing Shellbags.
    - is_live (bool): Flag indicating if live analysis mode.
    """
    output_dir = os.path.join(os.path.dirname(__file__), 'parsed_data')
    os.makedirs(output_dir, exist_ok=True)

    while True:
        try:
            if is_shellbags:
                if is_live:
                    command = [tool_path, '-l', '--csv', output_dir]
                else:
                    command = [tool_path, '-d', os.path.dirname(file_path), '--csv', output_dir]
            else:
                command = [tool_path, '-f', file_path, '--csv', output_dir]

            print(f"Analyzing file: {file_path}")
            result = subprocess.run(command, check=True, text=True)
            print(result.stdout)
            
            csv_file_path = find_latest_csv(output_dir)
            if csv_file_path:
                print(f"CSV file saved: {csv_file_path}")
                convert_csv_to_json(csv_file_path)
            else:
                print("No CSV file generated.")

            break  # Exit the loop if successful

        except FileNotFoundError:
            if not is_live:
                print(f"File not found: {file_path}")
                file_path = input("Please enter a valid file path or press Enter to cancel: ").strip()
                if not file_path:
                    break  # Cancel if user presses Enter without input
            else:
                raise  # Reraise the error during live analysis
        except subprocess.CalledProcessError as e:
            print(f"An error occurred while parsing the file: {e}")
            break  # Exit the loop on other subprocess errors

def find_latest_csv(directory):
    """
    Find the latest CSV file in a directory.

    Args:
    - directory (str): Directory to search for CSV files.

    Returns:
    - latest_file (str): Path to the latest CSV file found.
    """
    csv_files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.csv')]
    if not csv_files:
        return None
    latest_file = max(csv_files, key=os.path.getctime)
    return latest_file

def convert_csv_to_json(csv_file_path):
    """
    Convert a CSV file to JSON format and save.

    Args:
    - csv_file_path (str): Path to the CSV file.
    """
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

def live_analysis(tool_path, directories, extension, duration=60):
    """
    Perform live analysis of files using the specified tool for a fixed duration.

    Args:
    - tool_path (str): Path to the parsing tool executable.
    - directories (list): List of directories to monitor.
    - extension (str): File extension filter.
    - duration (int): Duration to run the live analysis (in seconds).
    """
    seen_files = set()
    end_time = time.time() + duration

    try:
        while time.time() < end_time:
            current_files = set()
            for folder, file_list in list_files(directories, extension).items():
                current_files.update(file_list)
            new_files = current_files - seen_files
            for file_path in new_files:
                try:
                    parse_file(tool_path, file_path, is_shellbags=False, is_live=True)
                except FileNotFoundError:
                    print(f"File not found during live analysis: {file_path}")
                except Exception as e:
                    print(f"An error occurred during live analysis: {e}")
            seen_files = current_files
            time.sleep(10)  # Adjust the sleep time as needed

    except KeyboardInterrupt:
        pass  # Handle Ctrl+C to exit 

    print("\nLive analysis completed.")


def get_live_choice():
    """
    Prompt user to choose whether to perform live analysis.

    Returns:
    - is_live (bool): True if user chooses live analysis, False otherwise.
    """
    while True:
        choice = input("Do you want to perform live analysis? (y/n): ").strip().lower()
        if choice in {'y', 'n'}:
            return choice == 'y'
        print("Invalid choice. Please enter 'y' or 'n'.")

def main():
    """
    Main function to run the shell parser program.
    """
    tool_paths = {
        '1': "C:\\Path\\To\\LECmd.exe", # shortcut parser (LECmd)
        '2': "C:\\Path\\To\\JLECmd.exe", # jump list parser (JLECmd)
        '3': "C:\\Path\\To\\PECmd.exe", # prefetch parser (PECmd)
        '4': "C:\\Path\\To\\SBECmd.exe" #shellbags parser (SBECmd)
    } # make sure to change the path for each parser tools (use double backslash "\\" to separate directories in a file path)

    list_functions = {
        '1': list_shortcut_files,
        '2': list_jump_list_files,
        '3': list_prefetch_files,
        '4': list_shellbags_files
    }

    directories = {
        '1': [
            os.path.join(os.environ['USERPROFILE'], 'Desktop'),
            os.path.join(os.environ['PROGRAMDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
            os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
            os.path.join(os.environ['APPDATA'], 'Microsoft', 'Internet Explorer', 'Quick Launch', 'User Pinned', 'TaskBar'),
            os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Recent')
        ],
        '2': [
            os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Recent', 'AutomaticDestinations'),
            os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Recent', 'CustomDestinations')
        ],
        '3': [
            os.path.join(os.environ['SYSTEMROOT'], 'Prefetch')
        ]
    }

    extensions = {
        '1': '.lnk',
        '2': '',
        '3': '.pf'
    }

    while True:
        choice = display_menu()
        if choice in tool_paths:
            is_live = get_live_choice()
            
            if is_live:
                if choice == '4':
                    try:
                        parse_file(tool_paths[choice], None, is_shellbags=True, is_live=True)
                        print("Live analysis completed.")
                        input("Press Enter to return to the main menu...")
                    except Exception as e:
                        print(f"An error occurred during live analysis: {e}")
                        input("Press Enter to return to the main menu...")
                else:
                    try:
                        live_analysis(tool_paths[choice], directories[choice], extensions[choice])
                    except Exception as e:
                        print(f"An error occurred during live analysis: {e}")
                        input("Press Enter to return to the main menu...")
            else:
                if choice == '4':
                    ntuser_files, usrclass_files = list_shellbags_files()
                    print("\nNTUSER.DAT files found:")
                    display_files(ntuser_files)
                    print("\nUsrClass.dat files found:")
                    display_files(usrclass_files)
                else:
                    files = list_functions[choice]()
                    display_files(files)
                
                while True:
                    file_path = input("\nEnter the path of the file or directory (or press Enter to return to main menu): ").strip()
                    if not file_path:
                        break  # Exit if user presses Enter without input

                    if not os.path.exists(file_path):
                        print("File or directory does not exist.")
                    else:
                        parse_file(tool_paths[choice], file_path, is_shellbags=(choice == '4'), is_live=False)
                        break

        elif choice == '0':
            break
        else:
            print("Invalid selection. Please choose a valid option.")

if __name__ == "__main__":
    main()