import os
import platform
import re
import pandas as pd
import PyPDF2
from docx import Document
import win32api  # For Windows drive detection
import pyfiglet
from pathlib import Path
from collections import deque
import time, threading

# Define constants
ERROR_LIMIT = 5
INTERVAL = 120  # 2 minutes in seconds

# Initialize error counter
error_counter = 0
error_queue = deque(maxlen=ERROR_LIMIT)

# Define colors for terminal output
red = "\033[1;31m"
green = "\033[1;32m"
cyan = "\033[1;36m"
reset = "\033[0m"

def clear():
    if platform.system() == "Windows":
        os.system("cls")
    else:
        os.system("clear")

def print_divider():
    print(f"{cyan}{'='*50}{reset}")

def print_error_divider():
    print(f"{red}{'='*50}{reset}")

def extract_emails(text):
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    return re.findall(email_pattern, text)

def process_text_file(file_path):
    emails = []
    try:
        with open(file_path, 'r', errors='ignore') as file:
            text = file.read()
            emails.extend(extract_emails(text))
    except Exception as e:
        handle_error(f"Error reading {file_path}: {e}")
    return emails

def process_spreadsheet(file_path):
    emails = []
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        for column in df.columns:
            if df[column].dtype == object:
                emails.extend(extract_emails('\n'.join(df[column].astype(str))))
    except Exception as e:
        handle_error(f"Error reading {file_path}: {e}")
    return emails

def process_pdf(file_path):
    emails = []
    try:
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    emails.extend(extract_emails(text))
    except PyPDF2.errors.PdfReadError as e:
        handle_error(f"Error reading PDF (likely corrupt or invalid): {file_path}: {e}")
    except Exception as e:
        handle_error(f"Error reading {file_path}: {e}")
    return emails

def process_docx(file_path):
    emails = []
    try:
        doc = Document(file_path)
        for para in doc.paragraphs:
            text = para.text
            emails.extend(extract_emails(text))
    except Exception as e:
        handle_error(f"Error reading DOCX (likely corrupt or invalid): {file_path}: {e}")
    return emails

def search_emails(start_path):
    all_emails = set()
    for root, _, files in os.walk(start_path):
        for file in files:
            file_path = os.path.join(root, file)
            if file.lower().endswith(('.txt', '.csv', '.xlsx', '.pdf', '.docx')):
                if file.lower().endswith('.txt') or file.lower().endswith('.csv'):
                    emails = process_text_file(file_path)
                elif file.lower().endswith('.xlsx'):
                    emails = process_spreadsheet(file_path)
                elif file.lower().endswith('.pdf'):
                    emails = process_pdf(file_path)
                elif file.lower().endswith('.docx'):
                    emails = process_docx(file_path)

                # Print found emails for testing
                if emails:
                    print(f"Found emails in {file_path}:")
                    for email in emails:
                        print(f"{cyan}{email}{reset}")

                all_emails.update(emails)
    return all_emails

def first_banner():
    clear()
    salfi = pyfiglet.Figlet(font="slant")
    banner = salfi.renderText("Cyber Fantics")
    print(f"""{red}{banner}
{green}============================
{red}[+] {green}Coded By @SalfiHacker
{red}[+] {green}Gift From Cyber Fantics
{red}[+] {green}It Search {cyan}For Email in Whole {cyan}Drives{green} in Each And Every {cyan}File
{green}============================""")

def handle_error(message):
    global error_counter
    error_queue.append(message)
    if len(error_queue) >= ERROR_LIMIT:
        first_banner()
        print_error_divider()
        print("\n".join(error_queue))
        print_error_divider()
        error_queue.clear()

def save_emails_periodically():
    desktop = Path.home() / 'Desktop'
    output_file = desktop / 'email.txt'

    # Get all available drives on Windows
    try:
        drives = win32api.GetLogicalDriveStrings().split('\000')[:-1]
    except Exception as e:
        handle_error(f"Error accessing drives: {e}")
        return

    while True:
        all_emails = set()

        for drive in drives:
            try:
                all_emails.update(search_emails(drive))
            except Exception as e:
                handle_error(f"Error accessing drive {drive}: {e}")

        # Save emails to a file
        with open(output_file, 'w') as file:
            for email in sorted(all_emails):
                file.write(email + '\n')

        print(f"{green}[+] Emails have been saved to {output_file}{reset}")
        time.sleep(INTERVAL)

def main():
    first_banner()
    
    user_input = input(f"This script will search for email addresses and save them to a file.\n Do you want to proceed? (yes/no): {cyan}").strip().lower()
    if user_input != 'yes':
        print("Exiting the program.")
        return
    
    # Start background tasks using threads
    threading.Thread(target=save_emails_periodically, daemon=True).start()

    # Keep the main thread running
    while True:
        time.sleep(1)

if __name__ == "__main__":
    main()
