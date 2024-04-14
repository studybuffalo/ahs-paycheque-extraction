"""Extracts and parses content from the PDF."""
from datetime import datetime
from decimal import Decimal
import re
from time import time

import fitz

class Coordinates:
    """Holds PDF coordinates."""
    def _generate_rect(self):
        """Generates a PyMuPDF rect object."""
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
    def _parse_coordinates(self, instances, selection=None):
        """Parses Rect object and stores coordinates.

            Parameters:
                instances (list): a list of PyMuPDF Rect objects.
                selection (str): an additional qualifer to specify which
                    instance to use if multiple instances are present.
        """
        instance = None

        if len(instances) > 1 and selection is None:
            print('Warning: Multiple Instances Found')

        if selection is None:
            instance = instances[0]

        elif selection == 'left':
            left_min = None
            min_index = None

            for index, instance in enumerate(instances):
                if index == 0:
                    left_min = instance[0]
                    min_index = index

                if instance[0] < left_min:
                    min_index = index

            instance = instances[min_index]

        elif selection == 'right':
            right_max = None
            right_index = None

            for index, instance in enumerate(instances):
                if index == 0:
                    right_max = instance[2]
                    right_index = index

                if instance[2] > right_max:
                    right_index = index

            instance = instances[right_index]

        elif selection == 'top':
            top_min = None
            top_index = None

            for index, instance in enumerate(instances):
                if index == 0:
                    top_min = instance[1]
                    top_index = index

                if instance[2] < top_min:
                    right_index = top_index

            instance = instances[top_index]

        elif selection == 'bottom':
            bottom_max = None
            bottom_index = None

            for index, instance in enumerate(instances):
                if index == 0:
                    bottom_max = instance[2]
                    bottom_index = index

                if instance[2] > bottom_max:
                    bottom_index = index

            instance = instances[bottom_index]

        return Coordinates(instance)

    def _identify_margins(self):
        """Identifies the page margins using the left and right-most objects."""
        self.log.info('Identify page margins')
        min_x = None
        max_x = None

        drawings = self.page.get_drawings()

        for index, drawing in enumerate(drawings):
            for item in drawing['items']:
                if item[0] == 'l':  # A line item
                    x_values = [item[1][0], item[2][0]]  # Start and end x coordinates of the line
                elif item[0] == 're':  # A rectangle item
                    x_values = [item[1][1], item[1][3]]  # Rectangle x coordinates
                else:
                    continue

                # Identifies the min and max values for all collected coordinates
                min_item_x = min(x_values)
                max_item_x = max(x_values)

                # Handles the first instance where no min or max exists
                if index == 0:
                    min_x = min_item_x
                    max_x = max_item_x
                    continue

                # Identify if this is a new minimum
                if min_item_x < min_x:
                    min_x = min_item_x

                # Identify if this is a new maximum
                if max_item_x > max_x:
                    max_x = max_item_x

        return min_x, max_x

    def _identify_pay_advice_coordinate(self):
        """Identifies coordinates to extract pay advice data."""
         # Identify the initial text anchors
        anchors = {}

        instances = self.page.search_for('Pay Begin Date:')
        anchors['pay_begin_date'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Pay End Date:')
        anchors['pay_end_date'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Advice #:')
        anchors['advice_number'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Advice Date:')
        anchors['advice_date'] = self._parse_coordinates(instances)

        # Calculate the relevant extraction coordinates
        right_coord_1 = anchors['advice_number'].left - 5
        right_coord_2 = self.right_margin

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
        anchors['employee_id'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Department:')
        anchors['department'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Location:')
        anchors['location'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Job Title:')
        anchors['job_title'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Pay Rate:')
        anchors['pay_rate'] = self._parse_coordinates(instances)

        instances = self.page.search_for('TAX DATA:')
        anchors['tax_data'] = self._parse_coordinates(instances)

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

    def _identify_tax_data_coordinates(self):
        """Identifies coordinates to extract tax data."""
         # Identify the initial text anchors
        anchors = {}

        instances = self.page.search_for('Quebec')
        anchors['quebec'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Net Claim Amount:')
        anchors['net_claim_amount'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Special Letters:')
        anchors['special_letters'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Addl. Percent:')
        anchors['additional_percent'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Addl. Amount:')
        anchors['additional_amount'] = self._parse_coordinates(instances)

        # Calculate the relevant extraction coordinates
        quebec_left = anchors['quebec'].left - 5
        quebec_right = anchors['quebec'].right + 5
        page_right = self.right_margin

        extract_coords = {
            'tax_data_federal_net_claim_amount': Coordinates([
                anchors['net_claim_amount'].right + 2,
                anchors['net_claim_amount'].top + 2,
                quebec_left,
                anchors['net_claim_amount'].bottom - 2,
            ]),
            'tax_data_federal_special_letters': Coordinates([
                anchors['special_letters'].right + 2,
                anchors['special_letters'].top + 2,
                quebec_left,
                anchors['special_letters'].bottom - 2,
            ]),
            'tax_data_federal_additional_percent': Coordinates([
                anchors['additional_percent'].right + 2,
                anchors['additional_percent'].top + 2,
                quebec_left,
                anchors['additional_percent'].bottom - 2,
            ]),
            'tax_data_federal_additional_amount': Coordinates([
                anchors['additional_amount'].right + 2,
                anchors['additional_amount'].top + 2,
                quebec_left,
                anchors['additional_amount'].bottom - 2,
            ]),
            'tax_data_alberta_net_claim_amount': Coordinates([
                quebec_right,
                anchors['net_claim_amount'].top + 2,
                page_right,
                anchors['net_claim_amount'].bottom - 2,
            ]),
            'tax_data_alberta_special_letters': Coordinates([
                quebec_right,
                anchors['special_letters'].top + 2,
                page_right,
                anchors['special_letters'].bottom - 2,
            ]),
            'tax_data_alberta_additional_percent': Coordinates([
                quebec_right,
                anchors['additional_percent'].top + 2,
                page_right,
                anchors['additional_percent'].bottom - 2,
            ]),
            'tax_data_alberta_additional_amount': Coordinates([
                quebec_right,
                anchors['additional_amount'].top + 2,
                page_right,
                anchors['additional_amount'].bottom - 2,
            ]),
        }

        return extract_coords

    def _identify_hours_coordinates(self):
        """Identifies coordinates to extract hours and earnings data."""
        # Identify the primary anchors to refine the search area
        primary_anchors = {}

        instances = self.page.search_for('HOURS AND EARNINGS')
        primary_anchors['hours_and_earnings'] = self._parse_coordinates(instances)

        instances = self.page.search_for('BEFORE-TAX DEDUCTIONS')
        primary_anchors['before_tax_deductions'] = self._parse_coordinates(instances)

        instances = self.page.search_for('TAX DATA:')
        primary_anchors['tax_data'] = self._parse_coordinates(instances)

        # Calculate the relevant extraction coordinates
        page_left = self.page_coordinates.left + 5

        # Construct the primary search area
        search_area = Coordinates([
                page_left,
                primary_anchors['hours_and_earnings'].bottom - 2,
                primary_anchors['tax_data'].left - 5,
                primary_anchors['before_tax_deductions'].top - 5,
            ])

        # Identify the text anchors
        anchors = {}

        instances = self.page.search_for('Description', clip=search_area.rect)
        anchors['description'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Rate', clip=search_area.rect)
        anchors['rate_current'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Hours', clip=search_area.rect)
        anchors['hours_current'] = self._parse_coordinates(instances, 'left')

        instances = self.page.search_for('Earnings', clip=search_area.rect)
        anchors['earnings_current'] = self._parse_coordinates(instances, 'left')

        instances = self.page.search_for('Hours', clip=search_area.rect)
        anchors['hours_ytd'] = self._parse_coordinates(instances, 'right')

        instances = self.page.search_for('Earnings', clip=search_area.rect)
        anchors['earnings_ytd'] = self._parse_coordinates(instances, 'right')

        instances = self.page.search_for('TOTAL:', clip=search_area.rect)
        anchors['total'] = self._parse_coordinates(instances, 'bottom')

        # Identify a secondary search area for rows
        row_search_area = Coordinates([
                anchors['earnings_ytd'].left,
                anchors['earnings_ytd'].bottom,
                anchors['earnings_ytd'].right,
                primary_anchors['before_tax_deductions'].top - 5,
            ])

        instances = self.page.search_for('.', clip=row_search_area.rect)
        row_anchors = [self._parse_coordinates([instance]) for instance in instances]

        # Pop off the last row, as this will be the "total" row
        row_anchors.pop()

        # Collect a list of coordinates for each row entry
        extract_coords = []

        for row in row_anchors:
            extract_coords.append({
                'description': Coordinates([
                    anchors['description'].left - 1,
                    row.top + 2,
                    anchors['rate_current'].left - 30,
                    row.bottom - 2,
                ]),
                'rate_current': Coordinates([
                    anchors['rate_current'].left - 20,
                    row.top + 2,
                    anchors['rate_current'].right + 2,
                    row.bottom - 2,
                ]),
                'hours_current': Coordinates([
                    anchors['rate_current'].right + 4,
                    row.top + 2,
                    anchors['hours_current'].right + 2,
                    row.bottom - 2,
                ]),
                'earnings_current': Coordinates([
                    anchors['hours_current'].right + 4,
                    row.top + 2,
                    anchors['earnings_current'].right + 2,
                    row.bottom - 2,
                ]),
                'hours_ytd': Coordinates([
                    anchors['earnings_current'].right + 4,
                    row.top + 2,
                    anchors['hours_ytd'].right + 2,
                    row.bottom - 2,
                ]),
                'earnings_ytd': Coordinates([
                    anchors['hours_ytd'].right + 4,
                    row.top + 2,
                    anchors['earnings_ytd'].right + 2,
                    row.bottom - 2,
                ]),
            })

        total = {
            'hours_current': Coordinates([
                anchors['rate_current'].right + 4,
                anchors['total'].top + 2,
                anchors['hours_current'].right + 2,
                anchors['total'].bottom - 2,
            ]),
            'earnings_current': Coordinates([
                anchors['hours_current'].right + 4,
                anchors['total'].top + 2,
                anchors['earnings_current'].right + 2,
                anchors['total'].bottom - 2,
            ]),
            'hours_ytd': Coordinates([
                anchors['earnings_current'].right + 4,
                anchors['total'].top + 2,
                anchors['hours_ytd'].right + 2,
                anchors['total'].bottom - 2,
            ]),
            'earnings_ytd': Coordinates([
                anchors['hours_ytd'].right + 4,
                anchors['total'].top + 2,
                anchors['earnings_ytd'].right + 2,
                anchors['total'].bottom - 2,
            ]),
        }

        return extract_coords, total

    def _identify_taxes_coordinates(self):
        """Identifies coordinates to extract taxes data."""
         # Identify the primary anchors to refine the search area
        primary_anchors = {}

        instances = self.page.search_for('TAXES')
        primary_anchors['taxes'] = self._parse_coordinates(instances, 'top')

        instances = self.page.search_for('EMPLOYER PAID BENEFITS')
        primary_anchors['employer_paid_benefits'] = self._parse_coordinates(instances)

        instances = self.page.search_for('TAX DATA:')
        primary_anchors['tax_data'] = self._parse_coordinates(instances)

        # Calculate the relevant extraction coordinates
        page_right = self.page_coordinates.right - 5

        # Construct the primary search area
        search_area = Coordinates([
                primary_anchors['tax_data'].left - 2,
                primary_anchors['taxes'].bottom - 2,
                page_right,
                primary_anchors['employer_paid_benefits'].top - 5,
            ])

        # Identify the text anchors
        anchors = {}

        instances = self.page.search_for('Description', clip=search_area.rect)
        anchors['description'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Current', clip=search_area.rect)
        anchors['current'] = self._parse_coordinates(instances)

        instances = self.page.search_for('YTD', clip=search_area.rect)
        anchors['ytd'] = self._parse_coordinates(instances)

        instances = self.page.search_for('TOTAL:', clip=search_area.rect)
        anchors['total'] = self._parse_coordinates(instances, 'bottom')

        # Identify a secondary search area for rows
        row_search_area = Coordinates([
            anchors['current'].right + 4,
            anchors['ytd'].bottom,
            anchors['ytd'].right,
            primary_anchors['employer_paid_benefits'].top - 5,
        ])

        instances = self.page.search_for('.', clip=row_search_area.rect)
        row_anchors = [self._parse_coordinates([instance]) for instance in instances]

        # Pop off the last row, as this will be the "total" row
        row_anchors.pop()

        # Collect a list of coordinates for each row entry
        extract_coords = []

        for row in row_anchors:
            extract_coords.append({
                'description': Coordinates([
                    anchors['description'].left - 1,
                    row.top + 2,
                    anchors['current'].left - 30,
                    row.bottom - 2,
                ]),
                'current': Coordinates([
                    anchors['current'].left - 20,
                    row.top + 2,
                    anchors['current'].right + 2,
                    row.bottom - 2,
                ]),
                'ytd': Coordinates([
                    anchors['current'].right + 4,
                    row.top + 2,
                    anchors['ytd'].right + 2,
                    row.bottom - 2,
                ]),
            })

        total = {
            'current': Coordinates([
                anchors['current'].left - 20,
                anchors['total'].top + 2,
                anchors['current'].right + 2,
                anchors['total'].bottom - 2,
            ]),
            'ytd': Coordinates([
                anchors['current'].right + 4,
                anchors['total'].top + 2,
                anchors['ytd'].right + 2,
                anchors['total'].bottom - 2,
            ]),
        }

        return extract_coords, total

    def _identify_before_tax_coordinates(self):
        """Identifies coordinates to extract before-tax deductions data."""
        # Identify the primary anchors to refine the search area
        primary_anchors = {}

        instances = self.page.search_for('BEFORE-TAX DEDUCTIONS')
        primary_anchors['before_tax_deductions'] = self._parse_coordinates(instances)

        instances = self.page.search_for('CIT TAXABLE GROSS')
        primary_anchors['cit_taxable_gross'] = self._parse_coordinates(instances)

        page_left = self.left_margin

        # These three columns are equally sized, so can just adjust for page
        # margins and divide by 3
        page_right = self.left_margin + ((self.right_margin - self.left_margin) / 3)

        # Construct the primary search area
        search_area = Coordinates([
                page_left,
                primary_anchors['before_tax_deductions'].bottom + 2,
                page_right,
                primary_anchors['cit_taxable_gross'].top - 5,
            ])

        # Identify the text anchors
        anchors = {}

        instances = self.page.search_for('Description', clip=search_area.rect)
        anchors['description'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Current', clip=search_area.rect)
        anchors['current'] = self._parse_coordinates(instances)

        instances = self.page.search_for('YTD', clip=search_area.rect)
        anchors['ytd'] = self._parse_coordinates(instances)

        instances = self.page.search_for('TOTAL:', clip=search_area.rect)
        anchors['total'] = self._parse_coordinates(instances, 'bottom')

        # Identify a secondary search area for rows
        row_search_area = Coordinates([
            anchors['current'].right + 4,
            anchors['ytd'].bottom,
            anchors['ytd'].right,
            primary_anchors['cit_taxable_gross'].top - 5,
        ])

        instances = self.page.search_for('.', clip=row_search_area.rect)
        row_anchors = [self._parse_coordinates([instance]) for instance in instances]

        # Pop off the last row, as this will be the "total" row
        row_anchors.pop()

        # Collect a list of coordinates for each row entry
        extract_coords = []

        for row in row_anchors:
            extract_coords.append({
                'description': Coordinates([
                    anchors['description'].left - 1,
                    row.top + 2,
                    anchors['current'].left - 30,
                    row.bottom - 2,
                ]),
                'current': Coordinates([
                    anchors['current'].left - 20,
                    row.top + 2,
                    anchors['current'].right + 2,
                    row.bottom - 2,
                ]),
                'ytd': Coordinates([
                    anchors['current'].right + 4,
                    row.top + 2,
                    anchors['ytd'].right + 2,
                    row.bottom - 2,
                ]),
            })

        total = {
            'description': 'Total',
            'current': Coordinates([
                anchors['current'].right - 20,
                anchors['total'].top + 2,
                anchors['current'].right + 2,
                anchors['total'].bottom - 2,
            ]),
            'ytd': Coordinates([
                anchors['current'].right + 4,
                anchors['total'].top + 2,
                anchors['ytd'].right + 2,
                anchors['total'].bottom - 2,
            ]),
        }

        return extract_coords, total

    def _identify_after_tax_coordinates(self):
        """Identifies coordinates to extract after-tax deductions data."""
        # Identify the primary anchors to refine the search area
        primary_anchors = {}

        instances = self.page.search_for('AFTER-TAX DEDUCTIONS')
        primary_anchors['after_tax_deductions'] = self._parse_coordinates(instances)

        instances = self.page.search_for('CIT TAXABLE GROSS')
        primary_anchors['cit_taxable_gross'] = self._parse_coordinates(instances)

        # These three columns are equally sized, so can just adjust for page
        # margins and divide by 3
        page_left = self.left_margin + ((self.right_margin - self.left_margin) / 3)
        page_right = page_left * 2

        # Construct the primary search area
        search_area = Coordinates([
                page_left,
                primary_anchors['after_tax_deductions'].bottom + 2,
                page_right,
                primary_anchors['cit_taxable_gross'].top - 5,
            ])

        # Identify the text anchors
        anchors = {}

        instances = self.page.search_for('Description', clip=search_area.rect)
        anchors['description'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Current', clip=search_area.rect)
        anchors['current'] = self._parse_coordinates(instances)

        instances = self.page.search_for('YTD', clip=search_area.rect)
        anchors['ytd'] = self._parse_coordinates(instances)

        instances = self.page.search_for('TOTAL:', clip=search_area.rect)
        anchors['total'] = self._parse_coordinates(instances, 'bottom')

        # Identify a secondary search area for rows
        row_search_area = Coordinates([
            anchors['current'].right + 4,
            anchors['ytd'].bottom,
            anchors['ytd'].right,
            primary_anchors['cit_taxable_gross'].top - 5,
        ])

        instances = self.page.search_for('.', clip=row_search_area.rect)
        row_anchors = [self._parse_coordinates([instance]) for instance in instances]

        # Pop off the last row, as this will be the "total" row
        row_anchors.pop()

        # Collect a list of coordinates for each row entry
        extract_coords = []

        for row in row_anchors:
            extract_coords.append({
                'description': Coordinates([
                    anchors['description'].left - 1,
                    row.top + 2,
                    anchors['current'].left - 30,
                    row.bottom - 2,
                ]),
                'current': Coordinates([
                    anchors['current'].left - 20,
                    row.top + 2,
                    anchors['current'].right + 2,
                    row.bottom - 2,
                ]),
                'ytd': Coordinates([
                    anchors['current'].right + 4,
                    row.top + 2,
                    anchors['ytd'].right + 2,
                    row.bottom - 2,
                ]),
            })

        total = {
            'description': 'Total',
            'current': Coordinates([
                anchors['current'].right - 20,
                anchors['total'].top + 2,
                anchors['current'].right + 2,
                anchors['total'].bottom - 2,
            ]),
            'ytd': Coordinates([
                anchors['current'].right + 4,
                anchors['total'].top + 2,
                anchors['ytd'].right + 2,
                anchors['total'].bottom - 2,
            ]),
        }

        return extract_coords, total

    def _identify_employer_benefits_coordinates(self):
        """Identifies coordinates to extract employer-paid benefits data."""
        # Identify the primary anchors to refine the search area
        primary_anchors = {}

        instances = self.page.search_for('EMPLOYER PAID BENEFITS')
        primary_anchors['employer_paid_benefits'] = self._parse_coordinates(instances)

        instances = self.page.search_for('CIT TAXABLE GROSS')
        primary_anchors['cit_taxable_gross'] = self._parse_coordinates(instances)

        instances = self.page.search_for('TAX DATA:')
        primary_anchors['tax_data'] = self._parse_coordinates(instances)

        # Construct the primary search area
        search_area = Coordinates([
            primary_anchors['tax_data'].left - 2,
            primary_anchors['employer_paid_benefits'].bottom + 2,
            self.right_margin,
            primary_anchors['cit_taxable_gross'].top - 5,
        ])

        # Identify the text anchors
        anchors = {}

        instances = self.page.search_for('Description', clip=search_area.rect)
        anchors['description'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Current', clip=search_area.rect)
        anchors['current'] = self._parse_coordinates(instances)

        instances = self.page.search_for('YTD', clip=search_area.rect)
        anchors['ytd'] = self._parse_coordinates(instances)

        instances = self.page.search_for('TOTAL:', clip=search_area.rect)
        anchors['total'] = self._parse_coordinates(instances, 'bottom')

        # Identify a secondary search area for rows
        row_search_area = Coordinates([
            anchors['current'].right + 4,
            anchors['ytd'].bottom,
            anchors['ytd'].right,
            primary_anchors['cit_taxable_gross'].top - 5,
        ])

        instances = self.page.search_for('.', clip=row_search_area.rect)
        row_anchors = [self._parse_coordinates([instance]) for instance in instances]

        # Pop off the last row, as this will be the "total" row
        row_anchors.pop()

        # Collect a list of coordinates for each row entry
        extract_coords = []

        for row in row_anchors:
            extract_coords.append({
                'description': Coordinates([
                    anchors['description'].left - 1,
                    row.top + 2,
                    anchors['current'].left - 30,
                    row.bottom - 2,
                ]),
                'current': Coordinates([
                    anchors['current'].left - 20,
                    row.top + 2,
                    anchors['current'].right + 2,
                    row.bottom - 2,
                ]),
                'ytd': Coordinates([
                    anchors['current'].right + 4,
                    row.top + 2,
                    anchors['ytd'].right + 2,
                    row.bottom - 2,
                ]),
            })

        total = {
            'current': Coordinates([
                anchors['current'].right - 20,
                anchors['total'].top + 2,
                anchors['current'].right + 2,
                anchors['total'].bottom - 2,
            ]),
            'ytd': Coordinates([
                anchors['current'].right + 4,
                anchors['total'].top + 2,
                anchors['ytd'].right + 2,
                anchors['total'].bottom - 2,
            ]),
        }

        return extract_coords, total

    def _identify_gross_and_net_coordinates(self):
        """Identifies coordinates to extract gross and net data."""
        # Identify the primary anchors to refine the search area
        primary_anchors = {}

        instances = self.page.search_for('CIT TAXABLE GROSS')
        primary_anchors['cit_taxable_gross'] = self._parse_coordinates(instances)

        instances = self.page.search_for('DIRECT DEPOSIT DISTRIBUTION')
        primary_anchors['direct_deposit_distribution'] = self._parse_coordinates(instances)

        # Construct the primary search area
        search_area = Coordinates([
            self.left_margin,
            primary_anchors['cit_taxable_gross'].top,
            self.right_margin,
            primary_anchors['direct_deposit_distribution'].top - 5,
        ])

        # Identify the text anchors
        anchors = {}

        instances = self.page.search_for('TOTAL GROSS', clip=search_area.rect)
        anchors['total_gross'] = self._parse_coordinates(instances)

        anchors['cit_taxable_gross'] = primary_anchors['cit_taxable_gross']

        instances = self.page.search_for('TOTAL TAXES', clip=search_area.rect)
        anchors['total_taxes'] = self._parse_coordinates(instances)

        instances = self.page.search_for('TOTAL DEDUCTIONS', clip=search_area.rect)
        anchors['total_deductions'] = self._parse_coordinates(instances)

        instances = self.page.search_for('NET PAY', clip=search_area.rect)
        anchors['net_pay'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Current:', clip=search_area.rect)
        anchors['current'] = self._parse_coordinates(instances)

        instances = self.page.search_for('YTD:', clip=search_area.rect)
        anchors['ytd'] = self._parse_coordinates(instances)

        # Collect a list of coordinates for each row entry
        extract_coords = {
            'gross_and_net': {
                'current': {
                    'total_gross': Coordinates([
                        anchors['current'].right + 2,
                        anchors['current'].top + 2,
                        anchors['total_gross'].right + 2,
                        anchors['current'].bottom - 2,
                    ]),
                    'cit_taxable_gross': Coordinates([
                        anchors['total_gross'].right + 4,
                        anchors['current'].top + 2,
                        anchors['cit_taxable_gross'].right + 2,
                        anchors['current'].bottom - 2,
                    ]),
                    'total_taxes': Coordinates([
                        anchors['cit_taxable_gross'].right + 4,
                        anchors['current'].top + 2,
                        anchors['total_taxes'].right + 2,
                        anchors['current'].bottom - 2,
                    ]),
                    'total_deductions': Coordinates([
                        anchors['total_taxes'].right + 4,
                        anchors['current'].top + 2,
                        anchors['total_deductions'].right + 2,
                        anchors['current'].bottom - 2,
                    ]),
                    'net_pay': Coordinates([
                        anchors['total_deductions'].right + 4,
                        anchors['current'].top + 2,
                        anchors['net_pay'].right + 2,
                        anchors['current'].bottom - 2,
                    ]),
                },
                'ytd': {
                    'total_gross': Coordinates([
                        anchors['ytd'].right + 2,
                        anchors['ytd'].top + 2,
                        anchors['total_gross'].right + 2,
                        anchors['ytd'].bottom - 2,
                    ]),
                    'cit_taxable_gross': Coordinates([
                        anchors['total_gross'].right + 4,
                        anchors['ytd'].top + 2,
                        anchors['cit_taxable_gross'].right + 2,
                        anchors['ytd'].bottom - 2,
                    ]),
                    'total_taxes': Coordinates([
                        anchors['cit_taxable_gross'].right + 4,
                        anchors['ytd'].top + 2,
                        anchors['total_taxes'].right + 2,
                        anchors['ytd'].bottom - 2,
                    ]),
                    'total_deductions': Coordinates([
                        anchors['total_taxes'].right + 4,
                        anchors['ytd'].top + 2,
                        anchors['total_deductions'].right + 2,
                        anchors['ytd'].bottom - 2,
                    ]),
                    'net_pay': Coordinates([
                    anchors['total_deductions'].right + 4,
                    anchors['ytd'].top + 2,
                    anchors['net_pay'].right + 2,
                    anchors['ytd'].bottom - 2,
                ]),
                },
            }
        }

        return extract_coords

    def _identify_vacation_coordinates(self):
        """Identifies coordinates to extract vacation accrual data."""
        # Identify the primary anchors to refine the search area
        primary_anchors = {}

        instances = self.page.search_for('Vacation Accrual')
        primary_anchors['vacation_accrual'] = self._parse_coordinates(instances)

        instances = self.page.search_for('YTD Bank Balances')
        primary_anchors['ytd_bank_balances'] = self._parse_coordinates(instances)

        instances = self.page.search_for('TOTAL')
        primary_anchors['total'] = self._parse_coordinates(instances, 'bottom')

        # Construct the primary search area
        search_area = Coordinates([
            self.left_margin,
            primary_anchors['vacation_accrual'].bottom - 2,
            primary_anchors['ytd_bank_balances'].left - 4,
            primary_anchors['total'].top - 5,
        ])

        # Identify the text anchors
        anchors = {}

        instances = self.page.search_for('Current:', clip=search_area.rect)
        anchors['current'] = self._parse_coordinates(instances)

        instances = self.page.search_for('supplemental', clip=search_area.rect)
        anchors['supplemental'] = self._parse_coordinates(instances)

        # Collect a list of coordinates for each row entry
        extract_coords = {
                'current': Coordinates([
                    anchors['current'].right + 2,
                    anchors['current'].top + 2,
                    primary_anchors['ytd_bank_balances'].left - 4,
                    anchors['current'].bottom - 2,
                ]),
                'supplemental': Coordinates([
                    anchors['supplemental'].right + 4,
                    anchors['supplemental'].top + 2,
                    primary_anchors['ytd_bank_balances'].left - 4,
                    anchors['supplemental'].bottom - 2,
                ])
        }

        # Try to collect "Next Year" vacation if present
        try:
            instances = self.page.search_for('Next Year:', clip=search_area.rect)
            anchors['next_year'] = self._parse_coordinates(instances)

            extract_coords['next_year'] = Coordinates([
                anchors['next_year'].right + 4,
                anchors['next_year'].top + 2,
                primary_anchors['ytd_bank_balances'].left - 4,
                anchors['next_year'].bottom - 2,
            ])
        except IndexError:
            extract_coords['next_year'] = ''

        return {'vacation': extract_coords}

    def _identify_bank_balances_coords(self):
        """Identifies coordinates to extract bank balances data."""
        # Identify the text anchors
        anchors = {}

        instances = self.page.search_for('YTD OT Bank')
        anchors['ytd_ot_bank'] = self._parse_coordinates(instances)

        instances = self.page.search_for('YTD Sick Bank')
        anchors['ytd_sick_bank'] = self._parse_coordinates(instances)

        instances = self.page.search_for('YTD Stat Bank')
        anchors['ytd_stat_bank'] = self._parse_coordinates(instances)

        instances = self.page.search_for('YTD Float Bank')
        anchors['ytd_float_bank'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Advance Outstanding')
        anchors['advance_outstanding'] = self._parse_coordinates(instances)

        # Collect a list of coordinates for each row entry
        extract_coords = {
            'ytd_ot_bank': Coordinates([
                anchors['ytd_ot_bank'].right + 4,
                anchors['ytd_ot_bank'].top + 2,
                anchors['advance_outstanding'].left - 4,
                anchors['ytd_ot_bank'].bottom - 2,
            ]),
            'ytd_sick_bank': Coordinates([
                anchors['ytd_sick_bank'].right + 4,
                anchors['ytd_sick_bank'].top + 2,
                anchors['advance_outstanding'].left - 4,
                anchors['ytd_sick_bank'].bottom - 2,
            ]),
            'ytd_stat_bank': Coordinates([
                anchors['ytd_stat_bank'].right + 4,
                anchors['ytd_stat_bank'].top + 2,
                anchors['advance_outstanding'].left - 4,
                anchors['ytd_stat_bank'].bottom - 2,
            ]),
            'ytd_float_bank': Coordinates([
                anchors['ytd_float_bank'].right + 4,
                anchors['ytd_float_bank'].top + 2,
                anchors['advance_outstanding'].left - 4,
                anchors['ytd_float_bank'].bottom - 2,
            ]),
        }

        return {'bank_balances': extract_coords}

    def _identify_advance_outstanding_coordinates(self):
        """Identifies coordinates to extract advance outstanding data."""
        # Identify the text anchors
        anchors = {}

        instances = self.page.search_for('OS/Advance')
        anchors['os_advance'] = self._parse_coordinates(instances)

        instances = self.page.search_for('DIRECT DEPOSIT DISTRIBUTION')
        anchors['direct_deposit_distribution'] = self._parse_coordinates(instances)

        # Collect a list of coordinates for each row entry
        extract_coords = {
            'os_advance': Coordinates([
                anchors['os_advance'].right + 4,
                anchors['os_advance'].top + 2,
                anchors['direct_deposit_distribution'].left - 4,
                anchors['os_advance'].bottom - 2,
            ])
        }

        return {'advance_outstanding': extract_coords}

    def _identify_direct_deposit_coordinates(self):
        """Identifies coordinates to extract direct deposit distribution data."""
        # Identify the primary anchors to refine the search area
        primary_anchors = {}

        instances = self.page.search_for('DIRECT DEPOSIT DISTRIBUTION')
        primary_anchors['direct_deposit_distribution'] = self._parse_coordinates(instances)

        instances = self.page.search_for('NET PAY DISTRIBUTION')
        primary_anchors['net_pay_distribution'] = self._parse_coordinates(instances)

        instances = self.page.search_for('TOTAL')
        primary_anchors['total'] = self._parse_coordinates(instances, 'bottom')

        # Construct the primary search area
        search_area = Coordinates([
            primary_anchors['direct_deposit_distribution'].left - 2,
            primary_anchors['direct_deposit_distribution'].bottom - 2,
            primary_anchors['net_pay_distribution'].left - 4,
            primary_anchors['total'].top - 5,
        ])

        # Identify the text anchors
        anchors = {}

        instances = self.page.search_for('Account Type', clip=search_area.rect)
        anchors['account_type'] = self._parse_coordinates(instances)

        instances = self.page.search_for('Deposit Amount', clip=search_area.rect)
        anchors['deposit_amount'] = self._parse_coordinates(instances)

        # Identify a secondary search area for rows
        row_search_area = Coordinates([
            anchors['deposit_amount'].left,
            anchors['deposit_amount'].bottom,
            anchors['deposit_amount'].right,
            primary_anchors['total'].top - 5,
        ])

        instances = self.page.search_for('.', clip=row_search_area.rect)
        row_anchors = [self._parse_coordinates([instance]) for instance in instances]

        # Collect a list of coordinates for each row entry
        extract_coords = []

        for row in row_anchors:
            extract_coords.append({
                'account_type': Coordinates([
                    anchors['account_type'].left - 1,
                    row.top + 2,
                    anchors['deposit_amount'].left - 4,
                    row.bottom - 2,
                ]),
                'deposit_amount': Coordinates([
                    anchors['deposit_amount'].left,
                    row.top + 2,
                    anchors['deposit_amount'].right,
                    row.bottom - 2,
                ])
            })

        # Add a total row
        total = {
            'account_type': 'Total',
            'deposit_amount': Coordinates([
                anchors['deposit_amount'].left,
                primary_anchors['total'].top + 2,
                anchors['deposit_amount'].right,
                primary_anchors['total'].bottom - 2,
            ])
        }

        return extract_coords, total

    def _identify_net_pay_distribution_coordinates(self):
        """Identifies coordinates to extract net pay distribution data."""
        # Identify the anchors to refine the search area
        anchors = {}

        instances = self.page.search_for('NET PAY DISTRIBUTION')
        anchors['net_pay_distribution'] = self._parse_coordinates(instances)

        instances = self.page.search_for('TOTAL')
        anchors['total'] = self._parse_coordinates(instances, 'bottom')

        # Construct the primary search area
        search_area = Coordinates([
            anchors['net_pay_distribution'].left - 2,
            anchors['net_pay_distribution'].bottom - 2,
            self.right_margin,
            anchors['total'].top - 5,
        ])

        # Conduct second search for rows
        instances = self.page.search_for('Advice', clip=search_area.rect)
        row_anchors = [self._parse_coordinates([instance]) for instance in instances]

        # Collect a list of coordinates for each row entry
        extract_coords = []

        for row in row_anchors:
            extract_coords.append({
                'advice_number': Coordinates([
                    anchors['net_pay_distribution'].left - 1,
                    row.top + 2,
                    anchors['net_pay_distribution'].right + 6,
                    row.bottom - 2,
                ]),
                'amount': Coordinates([
                    anchors['net_pay_distribution'].right + 10,
                    row.top + 2,
                    self.right_margin,
                    row.bottom - 2,
                ]),
            })

        # Add a total row
        total = {
            'advice_number': 'Total',
            'amount': Coordinates([
                anchors['total'].right + 10,
                anchors['total'].top + 2,
                self.right_margin,
                anchors['total'].bottom - 2,
            ])
        }

        return extract_coords, total

    def _identify_message_coordinates(self):
        """Identifies coordinates to extract message data."""
        # Identify the anchors to refine the search area
        anchors = {}

        instances = self.page.search_for('TOTAL')
        anchors['total'] = self._parse_coordinates(instances, 'bottom')

        # Get the lowest text on the page
        blocks = self.page.get_text("dict")['blocks']

        lowest_y = 0

        # Iterate through text blocks to find the lowest (maximum y1 value) text block
        for block in blocks:
            if block['type'] == 0:  # Type 0 indicates text
                bbox = block['bbox']  # Bounding box of the text block: (x0, y0, x1, y1)
                y1 = bbox[3]  # y1 is the bottom edge of the text block

                if y1 > lowest_y:
                    lowest_y = y1

        # Construct the primary search area
        search_area = Coordinates([
            self.left_margin,
            anchors['total'].bottom + 5,
            self.right_margin,
            lowest_y,
        ])

        instances = self.page.search_for('MESSAGE', clip=search_area.rect)
        anchors['message'] = self._parse_coordinates(instances, 'top')

        # Collect a list of coordinates for each row entry
        extract_coords = {
            'message': Coordinates([
                anchors['message'].left,
                anchors['message'].top,
                self.right_margin,
                lowest_y,
            ])
        }

        return extract_coords

    def _identify_coordinates(self):
        """Identifies the coordinates to extract data."""
        self.log.info('Identifying coordinates of data')

        pay_advice_coords = self._identify_pay_advice_coordinate()
        demographic_coords = self._identify_demographic_coordinates()
        tax_data_coords = self._identify_tax_data_coordinates()
        hours_coords, hours_total_coords = self._identify_hours_coordinates()
        taxes_coords, taxes_total_coords = self._identify_taxes_coordinates()
        before_tax_coords, before_tax_total_coords = self._identify_before_tax_coordinates()
        after_tax_coords, after_tax_total_coords = self._identify_after_tax_coordinates()
        employer_benefits_coords, employer_benefits_total_coords = self._identify_employer_benefits_coordinates()
        gross_and_net_coords = self._identify_gross_and_net_coordinates()
        vacation_coords = self._identify_vacation_coordinates()
        bank_balances_coords = self._identify_bank_balances_coords()
        advance_outstanding_coords = self._identify_advance_outstanding_coordinates()
        direct_deposit_coords, direct_deposit_total_coords = self._identify_direct_deposit_coordinates()
        pay_distribution_coords, pay_distribution_total_coords = self._identify_net_pay_distribution_coordinates()
        message_coords = self._identify_message_coordinates()

        return {
            **pay_advice_coords,
            **demographic_coords,
            **tax_data_coords,
            'hours_and_earnings': hours_coords,
            'taxes': taxes_coords,
            'before_tax_deductions': before_tax_coords,
            'after_tax_deductions': after_tax_coords,
            'employer_paid_benefits': employer_benefits_coords,
            **gross_and_net_coords,
            **vacation_coords,
            **bank_balances_coords,
            **advance_outstanding_coords,
            'direct_deposit_distribution': direct_deposit_coords,
            'net_pay_distribution': pay_distribution_coords,
            **message_coords,
            'totals': {
                'hours_and_earnings': hours_total_coords,
                'taxes': taxes_total_coords,
                'before_tax_deductions': before_tax_total_coords,
                'after_tax_deductions': after_tax_total_coords,
                'employer_paid_benefits': employer_benefits_total_coords,
                'direct_deposit_distribution': direct_deposit_total_coords,
                'net_pay_distribution': pay_distribution_total_coords,
            }
        }

    def _extract_from_pdf(self, coords, name, data_type='text'):
        """Extracts text from provided coordinates."""
        self.log.debug(f'    Extracting "{name}" ({data_type})')

        value = self.page.get_textbox(coords.rect).strip()

        self.log.debug(f'    Extracted value before formatting: {value}')

        if data_type == 'date':
            value = datetime.strptime(value, "%m/%d/%Y")
            value = value.date()
        elif data_type == 'currency' or data_type == 'number':
            # Strip out any non-number characters
            value = re.sub(r'[^\d.]+', '', value)

            # If value is blank, sub in "0"
            if value == '':
                value = '0'

            # Convert to decimal for future handling
            value = Decimal(value)

        self.log.debug(f'    Extracted value after formatting: {value}')

        return {
            'name': name,
            'value': value,
            'data_type': data_type
        }

    def _validate_extracted_data(self, data_list, total_dict):
        """Validates data list by confirming extracted data equals total."""
        for total_index, total_value in enumerate(total_dict):
            # Skip over any None values (as there is nothing to validate)
            if total_value is None:
                continue

            total = 0
            extracted_total = total_value['value']

            for item in data_list:
                total += item[total_index]['value']

            if total != total_value['value']:
                self.log.warning(
                    f'    {total_value['name']}: Calculated total ({total}) not equal to extracted total {extracted_total}')
            else:
                self.log.debug(
                    f'    {total_value['name']}: Calculated and extracted totals match ({extracted_total})'
                )

    def _extract_paycheque_details(self):
        """Extracts paychque details."""
        self.log.info('  Extracting paycheque details data')
        coords = self.extract_coordinates

        extract_data = [
            self._extract_from_pdf(
                coords['pay_begin_date'], 'Pay Begin Date', 'date'
            ),
            self._extract_from_pdf(
                coords['pay_end_date'], 'Pay End Date', 'date'
            ),
            self._extract_from_pdf(
                coords['advice_number'], 'Advice Number'
            ),
            self._extract_from_pdf(
                coords['advice_date'], 'Advice Date', 'date'
            ),
        ]

        return [extract_data]

    def _extract_baseline_details(self):
        """Extracts baseline details."""
        self.log.info('  Extracting baseline details data')
        coords = self.extract_coordinates

        extract_data = [
            self._extract_from_pdf(
                coords['employee_id'], 'Employee ID'
            ),
            self._extract_from_pdf(
                coords['department'], 'Department'
            ),
            self._extract_from_pdf(
                coords['location'], 'Location'
            ),
            self._extract_from_pdf(
                coords['job_title'], 'Job Title'
            ),
            self._extract_from_pdf(
                coords['pay_rate'], 'Pay Rate', 'currency'
            ),
        ]

        return [extract_data]

    def _extract_tax_data(self):
        """Extracts tax data."""
        self.log.info('  Extracting tax data')
        coords = self.extract_coordinates

        extract_data = [
            self._extract_from_pdf(
                coords['tax_data_federal_net_claim_amount'],
                'Tax Data - Federal - Net Claim Amount',
                'currency',
            ),
            self._extract_from_pdf(
                coords['tax_data_federal_special_letters'],
                'Tax Data - Federal - Special Letters',
                'currency',
            ),
            self._extract_from_pdf(
                coords['tax_data_federal_additional_percent'],
                'Tax Data - Federal - Additional Percent',
                'number',
            ),
            self._extract_from_pdf(
                coords['tax_data_federal_additional_amount'],
                'Tax Data - Federal - Additional Amount',
                'currency',
            ),
            self._extract_from_pdf(
                coords['tax_data_alberta_net_claim_amount'],
                'Tax Data - alberta - Net Claim Amount',
                'currency',
            ),
            self._extract_from_pdf(
                coords['tax_data_alberta_special_letters'],
                'Tax Data - alberta - Special Letters',
                'currency',
            ),
            self._extract_from_pdf(
                coords['tax_data_alberta_additional_percent'],
                'Tax Data - alberta - Additional Percent',
                'number',
            ),
            self._extract_from_pdf(
                coords['tax_data_alberta_additional_amount'],
                'Tax Data - alberta - Additional Amount',
                'currency',
            ),
        ]

        return [extract_data]

    def _extract_hours_and_earnings(self):
        """Extracts data for hours and earnings."""
        self.log.info('  Extracting Hours and Earnings data')
        coords = self.extract_coordinates

        # Organize Hours and Earnings Data
        extract_data = []

        for index, item in enumerate(coords['hours_and_earnings']):
            self.log.debug(f'    Extracting from row {index}')

            description = self._extract_from_pdf(
                item['description'], 'Description',
            )
            rate_current = self._extract_from_pdf(
                item['rate_current'], 'Rate - Current', 'currency'
            )
            hours_current = self._extract_from_pdf(
                item['hours_current'], 'Hours - Current', 'number'
            )
            earnings_current = self._extract_from_pdf(
                item['earnings_current'], 'Earnings - Current', 'currency'
            )
            hours_ytd = self._extract_from_pdf(
                item['hours_ytd'], 'Hours - YTD', 'number'
            )
            earnings_ytd = self._extract_from_pdf(
                item['earnings_ytd'], 'Earnings - YTD', 'currency'
            )

            extract_data.append([
                description,
                rate_current,
                hours_current,
                earnings_current,
                hours_ytd,
                earnings_ytd,
            ])

        # Extract Total data
        total = [
            None,
            None,
            self._extract_from_pdf(
                coords['totals']['hours_and_earnings']['hours_current'], 'Hours - Current', 'number'
            ),
            self._extract_from_pdf(
                coords['totals']['hours_and_earnings']['earnings_current'], 'Earnings - Current', 'number'
            ),
            self._extract_from_pdf(
                coords['totals']['hours_and_earnings']['hours_ytd'], 'Hours - YTD', 'number'
            ),
            self._extract_from_pdf(
                coords['totals']['hours_and_earnings']['earnings_ytd'], 'Earnings - YTD', 'number'
            ),
        ]

        # Validate the extracted data
        self.log.info('    Validating Hours and Earnings data')
        self._validate_extracted_data(extract_data, total)

        return extract_data

    def _extract_taxes(self):
        """Extracts data for taxes."""
        self.log.info('  Extracting Taxes')

        coords = self.extract_coordinates
        extract_data = []

        for index, item in enumerate(coords['taxes']):
            self.log.debug(f'    Extracting from row {index}')

            description = self._extract_from_pdf(
                item['description'], 'Description',
            )
            current = self._extract_from_pdf(
                item['current'], 'Current', 'currency'
            )
            ytd = self._extract_from_pdf(
                item['ytd'], 'YTD', 'currency'
            )

            extract_data.append([description, current, ytd])

        # Extract Total data
        total = [
            None,
            self._extract_from_pdf(
                coords['totals']['taxes']['current'], 'Current', 'currency'
            ),
            self._extract_from_pdf(
                coords['totals']['taxes']['ytd'], 'YTD', 'currency'
            ),
        ]

        # Validate the extracted data
        self.log.info('    Validating Taxes data')
        self._validate_extracted_data(extract_data, total)

        return extract_data

    def _extract_before_tax_deductions(self):
        """Extracts data for Before-Tax Deductions."""
        self.log.info('  Extracting Before-Tax Deductions')

        coords = self.extract_coordinates
        extract_data = []

        for index, item in enumerate(coords['before_tax_deductions']):
            self.log.debug(f'    Extracting from row {index}')

            description = self._extract_from_pdf(
                item['description'], 'Description',
            )
            current = self._extract_from_pdf(
                item['current'], 'Current', 'currency'
            )
            ytd = self._extract_from_pdf(
                item['ytd'], 'YTD', 'currency'
            )

            extract_data.append([description, current, ytd])

        # Extract Total data
        total = [
            None,
            self._extract_from_pdf(
                coords['totals']['before_tax_deductions']['current'], 'Current', 'currency'
            ),
            self._extract_from_pdf(
                coords['totals']['before_tax_deductions']['ytd'], 'YTD', 'currency'
            ),
        ]

        # Validate the extracted data
        self.log.info('    Validating Before-Tax Deductions data')
        self._validate_extracted_data(extract_data, total)

        return extract_data

    def _extract_after_tax_deductions(self):
        """Extracts data for After-Tax Deductions."""
        self.log.info('  Extracting After-Tax Deductions')

        coords = self.extract_coordinates
        extract_data = []

        for index, item in enumerate(coords['after_tax_deductions']):
            self.log.debug(f'    Extracting from row {index}')

            description = self._extract_from_pdf(
                item['description'], 'Description',
            )
            current = self._extract_from_pdf(
                item['current'], 'Current', 'currency'
            )
            ytd = self._extract_from_pdf(
                item['ytd'], 'YTD', 'currency'
            )

            extract_data.append([description, current, ytd])

        # Extract Total data
        total = [
            None,
            self._extract_from_pdf(
                coords['totals']['after_tax_deductions']['current'], 'Current', 'currency'
            ),
            self._extract_from_pdf(
                coords['totals']['after_tax_deductions']['ytd'], 'YTD', 'currency'
            ),
        ]

        # Validate the extracted data
        self.log.info('    Validating After-Tax Deductions data')
        self._validate_extracted_data(extract_data, total)

        return extract_data

    def _extract_employer_paid_benefits(self):
        """Extracts data for Employer Paid Benefits."""
        self.log.info('  Extracting Employer Paid Benefits')

        coords = self.extract_coordinates
        extract_data = []

        for index, item in enumerate(coords['employer_paid_benefits']):
            self.log.debug(f'    Extracting from row {index}')

            description = self._extract_from_pdf(
                item['description'], 'Description',
            )
            current = self._extract_from_pdf(
                item['current'], 'Current', 'currency'
            )
            ytd = self._extract_from_pdf(
                item['ytd'], 'YTD', 'currency'
            )

            extract_data.append([description, current, ytd])

        # Extract Total data
        total = [
            None,
            self._extract_from_pdf(
                coords['totals']['employer_paid_benefits']['current'], 'Current', 'currency'
            ),
            self._extract_from_pdf(
                coords['totals']['employer_paid_benefits']['ytd'], 'YTD', 'currency'
            ),
        ]

        # Validate the extracted data
        self.log.info('    Validating Employer Paid Benefits data')
        self._validate_extracted_data(extract_data, total)

        return extract_data

    def _extract_gross_and_net(self):
        """Extracts data for Gross and Net Pay."""
        self.log.info('  Extracting Gross and Net Pay')

        coords = self.extract_coordinates

        extract_data = [
            self._extract_from_pdf(
                coords['gross_and_net']['current']['total_gross'],
                'Current - Total Gross',
                'currency',
            ),
            self._extract_from_pdf(
                coords['gross_and_net']['current']['cit_taxable_gross'],
                'Current - CIT Taxable Gross',
                'currency',
            ),
            self._extract_from_pdf(
                coords['gross_and_net']['current']['total_taxes'],
                'Current - Total Taxes',
                'currency',
            ),
            self._extract_from_pdf(
                coords['gross_and_net']['current']['total_deductions'],
                'Current - Total Deductions',
                'currency',
            ),
            self._extract_from_pdf(
                coords['gross_and_net']['current']['net_pay'],
                'Current - Net Pay',
                'currency',
            ),
            self._extract_from_pdf(
                coords['gross_and_net']['ytd']['total_gross'],
                'YTD - Total Gross',
                'currency',
            ),
            self._extract_from_pdf(
                coords['gross_and_net']['ytd']['cit_taxable_gross'],
                'YTD - CIT Taxable Gross',
                'currency',
            ),
            self._extract_from_pdf(
                coords['gross_and_net']['ytd']['total_taxes'],
                'YTD - Total Taxes',
                'currency',
            ),
            self._extract_from_pdf(
                coords['gross_and_net']['ytd']['total_deductions'],
                'YTD - Total Deductions',
                'currency',
            ),
            self._extract_from_pdf(
                coords['gross_and_net']['ytd']['net_pay'],
                'YTD - Net Pay',
                'currency',
            ),
        ]

        return [extract_data]

    def _extract_vacation(self):
        """Extracts data for Vacation."""
        self.log.info('  Extracting Vacation')

        coords = self.extract_coordinates

        extract_data = [
            self._extract_from_pdf(
                coords['vacation']['current'],
                'Current',
                'number',
            ),
            self._extract_from_pdf(
                coords['vacation']['supplemental'],
                'Supplemental',
                'number',
            ),
        ]

        # Handle next-year vacation
        try:
            extract_data.append(self._extract_from_pdf(
                coords['vacation']['next_year'],
                'Next Year',
                'number',
            ))
        except AttributeError:
            extract_data.append(
                {'name': 'Next Year', 'value': 0, 'data_type': 'number'}
            )

        return [extract_data]

    def _extract_bank_balances(self):
        """Extracts data for Bank Balances."""
        self.log.info('  Extracting Bank Balances')

        coords = self.extract_coordinates

        extract_data = [
            self._extract_from_pdf(
                coords['bank_balances']['ytd_ot_bank'],
                'YTD OT Bank',
                'number',
            ),
            self._extract_from_pdf(
                coords['bank_balances']['ytd_sick_bank'],
                'YTD Sick Bank',
                'number',
            ),
            self._extract_from_pdf(
                coords['bank_balances']['ytd_stat_bank'],
                'YTD Stat Bank',
                'number',
            ),
            self._extract_from_pdf(
                coords['bank_balances']['ytd_float_bank'],
                'YTD Float Bank',
                'number',
            ),
        ]

        return [extract_data]

    def _extract_advance_outstanding(self):
        """Extracts data for Advance Outstanding."""
        self.log.info('  Extracting Advance Outstanding')

        coords = self.extract_coordinates

        extract_data = [self._extract_from_pdf(
            coords['advance_outstanding']['os_advance'], 'OS/Advance', 'currency',
        )]

        return [extract_data]

    def _extract_direct_deposit_distribution(self):
        """Extracts data for Direct Deposit Distribution."""
        self.log.info('  Extracting Direct Deposit Distribution')

        coords = self.extract_coordinates
        extract_data = []

        for index, item in enumerate(coords['direct_deposit_distribution']):
            self.log.debug(f'    Extracting from row {index}')

            account_type = self._extract_from_pdf(
                item['account_type'], 'Account Type',
            )
            deposit_amount = self._extract_from_pdf(
                item['deposit_amount'], 'Deposit Amount', 'currency'
            )

            extract_data.append([account_type, deposit_amount])

        # Extract Total data
        total = [
            None,
            self._extract_from_pdf(
                coords['totals']['direct_deposit_distribution']['deposit_amount'], 'Deposit Amount', 'currency'
            ),
        ]

        # Validate the extracted data
        self.log.info('    Validating Direct Deposit Distribution data')
        self._validate_extracted_data(extract_data, total)

        return extract_data

    def _extract_net_pay_distribution(self):
        """Extracts data for Net Pay Distribution."""
        self.log.info('  Extracting Net Pay Distribution')

        coords = self.extract_coordinates
        extract_data = []

        for index, item in enumerate(coords['net_pay_distribution']):
            self.log.debug(f'    Extracting from row {index}')

            advice_number = self._extract_from_pdf(
                item['advice_number'], 'Advice Number',
            )
            amount = self._extract_from_pdf(
                item['amount'], 'Amount', 'currency'
            )

            extract_data.append([advice_number, amount])

        # Extract Total data
        total = [
            None,
            self._extract_from_pdf(
                coords['totals']['net_pay_distribution']['amount'], 'Amount', 'currency'
            ),
        ]

        # Validate the extracted data
        self.log.info('    Validating Net Pay Distribution data')
        self._validate_extracted_data(extract_data, total)
        return extract_data

    def _extract_message(self):
        """Extracts data for Message."""
        self.log.info('  Extracting Message')

        coords = self.extract_coordinates

        extract_data = self._extract_from_pdf(coords['message'], 'Message')

        # Remove the "MESSAGE:" label
        extract_data['value'] = extract_data['value'].replace('MESSAGE:', '').strip()
        self.log.debug(f'    Extracted value after second formatting: {extract_data["value"]}')

        return [[extract_data]]

    def _extract_data(self):
        """Extracts data from all collected coordinates."""
        self.log.info('Extracting data from PDF')

        paycheque_details = self._extract_paycheque_details()
        baseline_details = self._extract_baseline_details()
        tax_data = self._extract_tax_data()
        hours_and_earnings = self._extract_hours_and_earnings()
        taxes = self._extract_taxes()
        before_tax_deductions = self._extract_before_tax_deductions()
        after_tax_deductions = self._extract_after_tax_deductions()
        employer_paid_benefits = self._extract_employer_paid_benefits()
        gross_and_net = self._extract_gross_and_net()
        vacation = self._extract_vacation()
        bank_balances = self._extract_bank_balances()
        advance_outstanding = self._extract_advance_outstanding()
        direct_deposit_distribution = self._extract_direct_deposit_distribution()
        net_pay_distribution = self._extract_net_pay_distribution()
        message = self._extract_message()

        data = {
            'paycheque_details': paycheque_details,
            'baseline_details': baseline_details,
            'tax_data': tax_data,
            'hours_and_earnings': hours_and_earnings,
            'taxes': taxes,
            'before_tax_deductions': before_tax_deductions,
            'after_tax_deductions': after_tax_deductions,
            'employer_paid_benefits': employer_paid_benefits,
            'gross_and_net': gross_and_net,
            'vacation': vacation,
            'bank_balances': bank_balances,
            'advance_outstanding': advance_outstanding,
            'direct_deposit_distribution': direct_deposit_distribution,
            'net_pay_distribution': net_pay_distribution,
            'message': message,
        }

        return data

    def _obtain_and_draw_coords(self, coords):
        """Function to isolate down to a Coordinates object."""
        if isinstance(coords, str):
            return
        elif isinstance(coords, Coordinates):
            self.page.draw_rect(coords.rect, color=fitz.pdfcolor['red'])
        elif isinstance(coords, list):
            for list_item in coords:
                self._obtain_and_draw_coords(list_item)
        elif isinstance(coords, dict):
            for _, dict_item in coords.items():
                self._obtain_and_draw_coords(dict_item)
        else:
            print(f'Unable to draw coordinates: {coords}')

    def draw_extract_coords(self):
        """Draws boxes around the extract coordinates for debugging."""
        for _, coords in self.extract_coordinates.items():
            self._obtain_and_draw_coords(coords)

        self.pdf.save(f'debug_{int(time())}.pdf')

    def __init__(self, pdf, log):
        self.pdf = pdf
        self.log = log
        self.page = self.pdf.load_page(0)
        self.page_coordinates = Coordinates(self.page.rect)
        left_margin, right_margin = self._identify_margins()
        self.left_margin = left_margin
        self.right_margin = right_margin
        self.extract_coordinates = self._identify_coordinates()
        self.data = self._extract_data()

def extract_data(pdf_path, config, log):
    pdf = fitz.open(pdf_path)
    data = PaychequeData(pdf, log)

    if config['save_coordinates']:
        data.draw_extract_coords()

    return data
