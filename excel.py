# import os
# import re

# import numpy as np
# import pandas as pd
# import requests
import xlwings as xw

import functions
import hide

# wb = xw.Book.caller()
# ws = wb.sheets.active

def main():
    pass

def fill_formula():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    functions.fill_formula(ws)