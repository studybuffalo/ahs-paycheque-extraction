"""Extracts and parses content from the PDF."""
from pprint import pprint
from time import time

import fitz

class Coordinates:
    """Holds PDF coordinates."""
    def _generate_rect(self):
        """Generats a PyMuPDF rect object."""
        return fitz.Rect(self.left, self.top, self.right, self.bottom)

    def __init__(self, rect):
        self.left = rect[0]
        self.top = rect[1]
        self.right = rect[2]
        self.bottom = rect[3]
        self.rect = self._generate_rect()

    def __str__(self):
        """String representation of the object"""
        return f'x1 = {self.left}; y1 = {self.top}; x2 = {self.right}; y2 = {self.bottom}'

class PaychequeData:
    """Class to extract and hold PDF data."""
    def _identify_pay_advice_coordinate(self):
        """Identifies coordinates to extract pay advice data."""
         # Identify the initial text anchors
        anchors = {}

        instances = self.page.search_for('Pay Begin Date:')
        anchors['pay_begin_date'] = Coordinates(instances[0])

        instances = self.page.search_for('Pay End Date:')
        anchors['pay_end_date'] = Coordinates(instances[0])

        instances = self.page.search_for('Advice #:')
        anchors['advice_number'] = Coordinates(instances[0])

        instances = self.page.search_for('Advice Date:')
        anchors['advice_date'] = Coordinates(instances[0])

        # Calculate the relevant extraction coordinates
        right_coord_1 = anchors['advice_number'].left - 5
        right_coord_2 = self.page_coordinates.right - 5

        extract_coords = {
            'pay_begin_date': Coordinates([
                anchors['pay_begin_date'].right + 2,
                anchors['pay_begin_date'].top + 2,
                right_coord_1,
                anchors['pay_begin_date'].bottom - 2,
            ]),
            'pay_end_date': Coordinates([
                anchors['pay_end_date'].right + 2,
                anchors['pay_end_date'].top + 2,
                right_coord_1,
                anchors['pay_end_date'].bottom - 2,
            ]),
            'advice_number': Coordinates([
                anchors['advice_number'].right + 2,
                anchors['advice_number'].top + 2,
                right_coord_2,
                anchors['advice_number'].bottom - 2,
            ]),
            'advice_date': Coordinates([
                anchors['advice_date'].right + 2,
                anchors['advice_date'].top + 2,
                right_coord_2,
                anchors['advice_date'].bottom - 2,
            ]),
        }

        return extract_coords

    def _identify_demographic_coordinates(self):
        """Identifies coordinates to extract demographic data."""
        # Identify the initial text anchors
        anchors = {}

        instances = self.page.search_for('Employee ID:')
        anchors['employee_id'] = Coordinates(instances[0])

        instances = self.page.search_for('Department:')
        anchors['department'] = Coordinates(instances[0])

        instances = self.page.search_for('Location:')
        anchors['location'] = Coordinates(instances[0])

        instances = self.page.search_for('Job Title:')
        anchors['job_title'] = Coordinates(instances[0])

        instances = self.page.search_for('Pay Rate:')
        anchors['pay_rate'] = Coordinates(instances[0])

        instances = self.page.search_for('TAX DATA:')
        anchors['tax_data'] = Coordinates(instances[0])

        # Calculate the relevant extraction coordinates
        left_coord = anchors['employee_id'].right + 2
        right_coord = anchors['tax_data'].left - 5

        extract_coords = {
            'employee_id': Coordinates([
                left_coord,
                anchors['employee_id'].top + 2,
                right_coord,
                anchors['employee_id'].bottom - 2,
            ]),
            'department': Coordinates([
                left_coord,
                anchors['department'].top + 2,
                right_coord,
                anchors['department'].bottom - 2,
            ]),
            'location': Coordinates([
                left_coord,
                anchors['location'].top + 2,
                right_coord,
                anchors['location'].bottom - 2,
            ]),
            'job_title': Coordinates([
                left_coord,
                anchors['job_title'].top + 2,
                right_coord,
                anchors['job_title'].bottom - 2,
            ]),
            'pay_rate': Coordinates([
                left_coord,
                anchors['pay_rate'].top + 2,
                right_coord,
                anchors['pay_rate'].bottom - 2,
            ]),
        }

        return extract_coords

    def _identify_coordinates(self):
        """Identifies the coordinates to extract data."""
        pay_advice_coords = self._identify_pay_advice_coordinate()
        demographic_coords = self._identify_demographic_coordinates()

        return {
            **pay_advice_coords,
            **demographic_coords,
        }

    def _extract_pdf_data(self):
        """Extracts PDF data."""
        data = {
            'pay_date_start': None,
            'pay_date_end': None,
            'advice_date': None,
            'advice_number': None,
            'employee_id': None,
            'department': None,
            'location': None,
            'job_title': None,
            'pay_rate': None,
            'tax_data': {
                'federal': {
                    'net_claim_amount': None,
                    'special_letters': None,
                    'additional_percent': None,
                    'additional_amount': None,
                },
                'alberta': {
                    'net_claim_amount': None,
                    'special_letters': None,
                    'additional_percent': None,
                    'additional_amount': None,
                },
            },
            'earnings': [],
            'taxes': {
                'cit': {
                    'current': None,
                    'ytd': None,
                },
                'cpp': {
                    'current': None,
                    'ytd': None,
                },
                'ei': {
                    'current': None,
                    'ytd': None,
                },
            },
        }

        return data

    def draw_extract_coords(self):
        """Draws boxes around the extract coordinates for debugging."""
        for _, coords in self.extract_coordinates.items():
            self.page.draw_rect(coords.rect, color=fitz.pdfcolor['red'])

        self.pdf.save(f'debug_{int(time())}.pdf')

    def __init__(self, pdf):
        self.pdf = pdf
        self.page = self.pdf.load_page(0)
        self.page_coordinates = Coordinates(self.page.rect)
        self.extract_coordinates = self._identify_coordinates()
        self.data = self._extract_pdf_data()

def extract_data(config):
    pdf = fitz.open(config['pdf_file_path'])
    data = PaychequeData(pdf)
    data.draw_extract_coords()

    return data

