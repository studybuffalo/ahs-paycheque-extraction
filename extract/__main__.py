"""Extracts details from AHS paycheque PDF."""
from utils import (
    extract_data, generate_config, identify_pdf, setup_logging, save_data,
)


def main():
    """Main function to run application."""
    # Setup Config and Logging details
    config = generate_config()
    log = setup_logging(config)

    # Identify which PDF to extract data from
    pdf_path = identify_pdf(config, log)

    if pdf_path is None:
        return

    # Extract Data from selected PDF
    data = extract_data(pdf_path, config, log)

    # Save PDF data to Excel file
    save_data(data.data, config, log)

if __name__ == '__main__':
    main()
