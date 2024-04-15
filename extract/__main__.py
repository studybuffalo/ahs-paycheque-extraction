"""Extracts details from AHS paycheque PDF."""
import os
from pathlib import Path
import shutil

from utils import extract_data, generate_config, setup_logging, save_data, move_pdf


def main():
    """Main function to run application."""
    # Setup Config and Logging details
    config = generate_config()
    log = setup_logging(config)

    # Iterate through each file in the folder
    log.info('Collecting files for extraction')
    pdf_files = [
        Path(config['pdf_extract_path'], file) for file in os.listdir(config['pdf_extract_path']) if file.endswith('.pdf')
    ]

    for file in pdf_files:
        log.info(f'Extracting data from {file}')

        # Extract Data from selected PDF
        data = extract_data(file, config, log)

        # Save PDF data
        save_data(data.data, config, log)

    # Move PDF to the configured path
    log.info('Moving files to configured directory')

    for file in pdf_files:
        move_pdf(file, config, log)

if __name__ == '__main__':
    main()
