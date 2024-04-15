"""Utility classes, functions and variables for the application."""
import logging
import os
from pathlib import Path
import shutil

from dotenv import find_dotenv, load_dotenv


def generate_config():
    """Generates the configuration details for app."""
    load_dotenv(find_dotenv(filename='config.env'))

    config = {
        'pdf_extract_path': Path(os.getenv('PDF_EXTRACT_PATH')),
        'pdf_move_path': Path(os.getenv('PDF_MOVE_PATH')),
        'data_path': Path(os.getenv('DATA_PATH')),
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

def move_pdf(file, config, log):
    """Moves PDF to the configured path."""
    move_path = Path(config['pdf_move_path'], file.name)

    try:
        shutil.move(file, move_path)
        log.info(f'  PDF moved to {move_path}')
    except IOError:
        try:
            os.remove(move_path)
            shutil.move(file, move_path)
            log.info(f'  PDF moved to {move_path} (existing file overwritten)')
        except PermissionError as e:
            log.info(f'  Unable to move PDF to {move_path}: {e}')
