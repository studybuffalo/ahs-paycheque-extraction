"""Utility classes, functions and variables for the application."""
import logging
import os
from pathlib import Path

from dotenv import find_dotenv, load_dotenv


def generate_config():
    """Generates the configuration details for app."""
    load_dotenv(find_dotenv(filename='config.env'))

    config = {
        'pdf_dir_path': Path(os.getenv('PDF_DIR_PATH')),
        'excel_file_path': Path(os.getenv('EXCEL_FILE_PATH')),
        'excel_backup_path': Path(os.getenv('EXCEL_BACKUP_PATH')),
        'log_level': int(os.getenv('LOG_LEVEL', '20')),
        'save_coordinates': os.getenv('SAVE_COORDINATES', False) == 'True',
    }

    return config

def setup_logging(config):
    """Setups logging for the app."""
    log = logging.getLogger('ahs-paycheque-extraction')
    console_format = logging.Formatter(
        '{levelname:<8} | {message}',
        style='{',
    )
    console_handler = logging.StreamHandler()
    console_handler.setLevel(config['log_level'])
    console_handler.setFormatter(console_format)

    log.setLevel(config['log_level'])
    log.addHandler(console_handler)

    log.debug(f'Logger setup with logging level {log.level}')

    return log

def display_files(files, start_index):
    """Display a paginated list of files."""
    end_index = start_index + 5

    for i, file in enumerate(files[start_index:end_index], start=start_index + 1):
        print(f'{i:<2} | {file}')

    return start_index, end_index

def identify_pdf(config, log):
    """Identifies which PDF to extract."""
    # Collects all PDF files in the provided directory
    pdf_files = [f for f in os.listdir(config['pdf_dir_path']) if f.endswith('.pdf')]

    # Sorts files in descending order
    pdf_files.sort(reverse=True)  # Sort files in descending order

    # Notifies user that no PDF files are found
    if not pdf_files:
        print('No PDF files found in directory. Confirm the correct directory is listed in the config file.')
        return

    start_index = 0

    while True:
        # User command options
        print('Enter the number to select the desired file. You may also enter "n" for the next page, "p" for previous page, or "q" to quit.')
        current_start, current_end = display_files(pdf_files, start_index)
        command = input('Selection: ').strip().lower()

        if command.isdigit():
            # Select file by number
            selection = int(command)
            if current_start < selection <= current_end:
                file = Path(config['pdf_dir_path'], pdf_files[selection - 1])
                log.info(f'File selected for extraction: {file}')

                return file
            else:
                print('Invalid selection, please try again.')

        elif command == 'n':
            # Go to the next page
            if current_end < len(pdf_files):
                start_index = current_end
            else:
                print('This is the last page, there are no further files.')

        elif command == 'p':
            # Go to the previous page
            start_index = max(0, len(pdf_files) - start_index)

            if start_index == 0:
                print('This is the first page, there are no earlier files.')

        elif command == 'q':
            # Quit the program
            return None
        else:
            print('Invalid command, please try again.')
