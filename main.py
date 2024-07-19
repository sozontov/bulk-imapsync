import pandas as pd
import os
import colorama
import textwrap
import pyfiglet
import subprocess
from tabulate import tabulate

def readcmd(cmd):
    result = subprocess.run(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    return result.stdout

def read_excel(action):
    try:
        df = pd.read_excel('working.xlsx')
        accounts = df.to_dict(orient='records')
        if action == 'showaccounts':
            show_accounts(accounts)
        elif action == 'checklogin':
            checklogin(accounts)
        elif action == 'sync':
            sync(accounts)
        elif action == 'showlogs':
            show_logs(accounts)
    except Exception as e:
        print(f"Error: {e}")

def show_accounts(accounts):
    response = [["Index", "From", "To"]]
    for index, account in enumerate(accounts, start=1):
        response.append([index, account["from_user"], account["to_user"]])
    print(colorama.Fore.LIGHTYELLOW_EX)
    print(tabulate(response))
    print(colorama.Fore.RESET)

def checklogin(accounts):
    response = [["Index", "From", "To", "Status", "Message"]]
    for index, account in enumerate(accounts, start=1):
        try:
            print(f"Checking {account['from_user']} -> {account['to_user']} ...")
            cmd = [
                "imapsync",
                "--host1", account['from_host'],
                "--user1", account['from_user'],
                "--password1", account['from_password'],
                "--host2", account['to_host'],
                "--user2", account['to_user'],
                "--password2", account['to_password'],
                "--justlogin"
            ]
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            if "Exiting with return value 0" in result.stdout:
                response.append([index, account["from_user"], account["to_user"], "success", ""])
            else:
                error_message = [line for line in result.stdout.split('\n') if "failure:" in line]
                response.append([index, account["from_user"], account["to_user"], "error", '\n'.join(error_message)])
        except Exception as e:
            response.append([index, account["from_user"], account["to_user"], "error", str(e)])
    print(colorama.Fore.LIGHTYELLOW_EX)
    print(tabulate(response))
    print(colorama.Fore.RESET)

def sync(accounts):
    response = [["Index", "From", "To", "Status", "Message"]]
    try:
        from_index = int(input("Sync from index: "))
        to_index = int(input("Sync to index: "))
        if not (1 <= from_index <= len(accounts)) or not (1 <= to_index <= len(accounts)):
            print(f'Index must be between 1 and {len(accounts)}')
            return
    except ValueError:
        print('Invalid input')
        return
    
    for index, account in enumerate(accounts, start=1):
        if from_index <= index <= to_index:
            try:
                print(f"Syncing {account['from_user']} -> {account['to_user']} ...")
                sync_cmd = [
                    "imapsync",
                    "--addheader",
                    "--automap",
                    "--host1", account['from_host'],
                    "--user1", account['from_user'],
                    "--password1", account['from_password'],
                    "--host2", account['to_host'],
                    "--user2", account['to_user'],
                    "--password2", account['to_password']
                ]
                subprocess.run(sync_cmd)
            except Exception as e:
                response.append([index, account["from_user"], account["to_user"], "error", str(e)])
    print(colorama.Fore.LIGHTYELLOW_EX)
    print(tabulate(response))
    print(colorama.Fore.RESET)

def show_logs(accounts):
    try:
        log_index = int(input("Input index to show logs: "))
        if not (1 <= log_index <= len(accounts)):
            print(f'Index must be between 1 and {len(accounts)}')
            return
    except ValueError:
        print('Invalid input')
        return

    for index, account in enumerate(accounts, start=1):
        if log_index == index:
            try:
                while True:
                    print(f"Show log file {account['from_user']} -> {account['to_user']} ...")
                    file_list = os.listdir("LOG_imapsync")
                    response = [["Index", "Date", "File Name"]]
                    fileIndex = 0
                    for file in file_list:
                        file_parts = file.split('_')
                        if len(file_parts) < 9:
                            continue
                        file_date = f"{file_parts[2]}/{file_parts[1]}/{file_parts[0]} {file_parts[3]}:{file_parts[4]}:{file_parts[5]}"
                        fromto = f"{file_parts[7]} {file_parts[8]}"
                        if fromto == f"{account['from_user']} {account['to_user']}.txt":
                            fileIndex += 1
                            response.append([fileIndex, file_date, file])
                    print(colorama.Fore.LIGHTYELLOW_EX)
                    print(tabulate(response))
                    print(colorama.Fore.RESET)

                    log_file_index = input("Input index logs file (Press 'a' to break): ")
                    if log_file_index.lower() == 'a':
                        break
                    try:
                        log_file_index = int(log_file_index)
                        log_file = next(file for file in response if file[0] == log_file_index)
                        print(f'Tail 1000 lines of file {log_file[2]}')
                        os.system('clear')
                        os.system(f'tail -1000 LOG_imapsync/{log_file[2]}')
                        input("Press any key to exit! ")
                    except (ValueError, StopIteration):
                        print('Invalid log file index')
            except Exception as e:
                print(e)

def main():
    print(pyfiglet.figlet_format("BULK IMAPSYNC"))
    print("""
    0. Exit
    1. Show accounts
    2. Check login
    3. Sync accounts
    4. Show log file
    """)
    menu_selected = input("Select tool: ")
    if menu_selected == '0':
        print("Bye!")
        return False
    elif menu_selected == '1':
        read_excel('showaccounts')
    elif menu_selected == '2':
        read_excel('checklogin')
    elif menu_selected == '3':
        read_excel('sync')
    elif menu_selected == '4':
        read_excel('showlogs')
    return True

while main():
    pass
