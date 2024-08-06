# Email Finder
Welcome to the Email Finder project!

## Description
The Email Finder script is designed to search for email addresses in various types of files located on your drives. It supports files with extensions such as `.txt`, `.csv`, `.xlsx`, `.pdf`, and `.docx`. The script scans specified directories, extracts email addresses from these files, and saves them to a text file on your Desktop.

## Features
- **Searches for emails:** Finds email addresses in text files, spreadsheets, PDFs, and DOCX files.
- **Saves results:** Writes the found email addresses to a text file named `email.txt` on your Desktop.
- **Error Handling:** Logs errors encountered during the scanning process and prints them in the terminal.

## Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/cyberfantics/email-founder.git
   cd email-founder
   ```

2. **Install required packages:**
   You need Python 3.x and `pip` installed on your system. Install the required packages using:

   ```bash
   pip install pandas PyPDF2 python-docx pywin32
   ```

## Usage

1. **Run the script:**
   ```bash
   python main.py
   ```

2. **Follow the prompts:**
   The script will prompt you to confirm whether you want to proceed. Type `yes` to start the email search.

## File Structure
- `main.py`: The main script that performs the email search.
- `README.md`: This file.

## Contributing
If you wish to contribute to this project, please fork the repository, make your changes, and submit a pull request. 


## Author
Syed Mansoor ul Hassan Bukhari
