"""Saves extracted data to an Excel file."""
import csv
import os
from pathlib import Path


def confirm_or_create_save_directories(config, log):
    """Confirms the required save directories exist and creats them if needed."""
    log.info(f'  Confirming or creating directories to save extracted data: {config['data_path']}')

    required_directories = [
        'Pay Cheque Details',
        'Baseline Details',
        'Tax Data',
        'Hours and Earnings',
        'Taxes',
        'Before-Tax Deductions',
        'After-Tax Deductions',
        'Employer Paid Benefits',
        'Gross and Net Pay',
        'Vacation',
        'Bank Balances',
        'Advance Outstanding',
        'Direct Deposit Distribution',
        'Net Pay Distribution',
        'Message'
    ]

    for directory in required_directories:
        directory_path = Path(config['data_path'], directory)

        if not directory_path.exists():
            directory_path.mkdir(exist_ok=True)

def save_data(data, config, log):
    """Saves extracted data to Excel file."""
    log.info('Saving Data')

    confirm_or_create_save_directories(config, log)

    # Dictionary mapping extracted data to required output details
    dict_map = {
        'paycheque_details': {
            'folder_name': 'Pay Cheque Details',
            'headers': [
                'Pay Begin Date',
                'Pay End Date',
                'Advice Number',
                'Advice Date',
            ]
        },
        'baseline_details': {
            'folder_name': 'Baseline Details',
            'headers': [
                'Pay Begin Date',
                'Pay End Date',
                'Advice Number',
                'Advice Date',
                'Employee ID',
                'Department',
                'Location',
                'Job Title',
                'Pay Rate',
            ]
        },
        'tax_data': {
            'folder_name': 'Tax Data',
            'headers': [
                'Pay Begin Date',
                'Pay End Date',
                'Advice Number',
                'Advice Date',
                'Federal - Net Claim Amount',
                'Federal - Special Letters',
                'Federal - Additional Percent',
                'Federal - Additional Amount',
                'Alberta - Net Claim Amount',
                'Alberta - Special Letters',
                'Alberta - Additional Percent',
                'Alberta - Additional Amount',
            ]
        },
        'hours_and_earnings': {
            'folder_name': 'Hours and Earnings',
            'headers': [
                'Pay Begin Date',
                'Pay End Date',
                'Advice Number',
                'Advice Date',
                'Description	Current - Rate',
                'Current - Rate',
                'Current - Hours',
                'Current - Earnings',
                'YTD - Hours',
                'YTD - Earnings',
            ]
        },
        'taxes': {
            'folder_name': 'Taxes',
            'headers': [
                'Pay Begin Date',
                'Pay End Date',
                'Advice Number',
                'Advice Date',
                'Description',
                'Current',
                'YTD',
            ]
        },
        'before_tax_deductions': {
            'folder_name': 'Before-Tax Deductions',
            'headers': [
                'Pay Begin Date',
                'Pay End Date',
                'Advice Number',
                'Advice Date',
                'Description',
                'Current',
                'YTD',
            ]
        },
        'after_tax_deductions': {
            'folder_name': 'After-Tax Deductions',
            'headers': [
                'Pay Begin Date',
                'Pay End Date',
                'Advice Number',
                'Advice Date',
                'Description',
                'Current',
                'YTD',
            ]
        },
        'employer_paid_benefits': {
            'folder_name': 'Employer Paid Benefits',
            'headers': [
                'Pay Begin Date',
                'Pay End Date',
                'Advice Number',
                'Advice Date',
                'Description',
                'Current',
                'YTD',
            ]
        },
        'gross_and_net': {
            'folder_name': 'Gross and Net Pay',
            'headers': [
                'Pay Begin Date',
                'Pay End Date',
                'Advice Number',
                'Advice Date',
                'Current - Total Gross',
                'Current - CIT Taxable Gross',
                'Current - Total Taxes',
                'Current - Total Deductions',
                'Current - Net Pay',
                'YTD - Total Gross',
                'YTD - CIT Taxable Gross',
                'YTD - Total Taxes',
                'YTD - Total Deductions',
                'YTD - Net Pay',
            ]
        },
        'vacation': {
            'folder_name': 'Vacation',
            'headers': [
                'Pay Begin Date',
                'Pay End Date',
                'Advice Number',
                'Advice Date',
                'Current',
                'Supplemental',
                'Next Year',

            ]
        },
        'bank_balances': {
            'folder_name': 'Bank Balances',
            'headers': [
                'Pay Begin Date',
                'Pay End Date',
                'Advice Number',
                'Advice Date',
                'YTD OT Bank',
                'YTD Sick Bank',
                'YTD Stat Bank',
                'YTD Float Bank',

            ]
        },
        'advance_outstanding': {
            'folder_name': 'Advance Outstanding',
            'headers': [
                'Pay Begin Date',
                'Pay End Date',
                'Advice Number',
                'Advice Date',
                'OS/Advance',
            ]
        },
        'direct_deposit_distribution': {
            'folder_name': 'Direct Deposit Distribution',
            'headers': [
                'Pay Begin Date',
                'Pay End Date',
                'Advice Number',
                'Advice Date',
                'Account Type',
                'Deposit Amount',
            ]
        },
        'net_pay_distribution': {
            'folder_name': 'Net Pay Distribution',
            'headers': [
                'Pay Begin Date',
                'Pay End Date',
                'Advice Number',
                'Advice Date',
                'Advice Number Reference',
                'Amount',
            ]
        },
        'message': {
            'folder_name': 'Message',
            'headers': [
                'Pay Begin Date',
                'Pay End Date',
                'Advice Number',
                'Advice Date',
                'Message',
            ]
        },
    }

    # Collect the paycheque details; these are applied to every table
    paycheque_details = data['paycheque_details'][0]
    date_start = paycheque_details[0]['value'].strftime('%Y-%m-%d')
    date_end = paycheque_details[1]['value'].strftime('%Y-%m-%d')

    for key, item in dict_map.items():
        csv_name = f'{item["folder_name"]} - {date_start} to {date_end}'
        csv_path = Path(config['data_path'], item['folder_name'], f'{csv_name}.csv')

        log.info(f'  Saving data to: {csv_path}')

        # Remove existing CSV if it exists
        if csv_path.exists():
            os.remove(csv_path)

        with open(csv_path, 'w', newline='') as file:
            writer = csv.writer(file)

            # Write the header row
            writer.writerow(item['headers'])

            # Organize the data for saving
            save_data = []

            for row in data[key]:
                row_data = []

                # Add Paycheque details to all other groups of data
                if key == 'paycheque_details':
                    updated_row = row
                else:
                    updated_row = paycheque_details + row

                for cell in updated_row:
                    row_data.append(cell['value'])

                save_data.append(row_data)

            # Write the data
            writer.writerows(save_data)
