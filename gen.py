
# functions for generating the intermediate 
# files based on an Excel document 

import openpyxl 
import os 
import sys 


class ScanResults:

    def __init__(self):
        self.formulas = dict()

    def record_formula(self, formula_str, in_cell, in_sheet):
        if in_sheet not in self.formulas:
            self.formulas[in_sheet] = dict()
        
        self.formulas[in_sheet][in_cell] = formula_str


def scan_excel(input_excel):
    
    wb = openpyxl.load_workbook(input_excel)






def dump_scanned(obj):   # input ScanResults to this one 
    print(obj)

