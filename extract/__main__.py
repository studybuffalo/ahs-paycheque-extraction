"""Extracts details from AHS paycheque PDF."""
from pprint import pprint

from utils import extract_data, generate_config


CONFIG = generate_config()
DATA = extract_data(CONFIG)

pprint(DATA)
