"""Utility classes, functions and variables for the application."""
import os
from pathlib import Path

from dotenv import find_dotenv, load_dotenv


def generate_config():
    """Generates the configuration details for app."""
    load_dotenv(find_dotenv(filename='config.env'))

    config = {
        'pdf_file_path': Path(os.getenv('PDF_FILE_PATH'))
    }

    return config
