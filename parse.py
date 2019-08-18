# -*- coding: utf-8 -*-

import sys
import os
from pathlib import Path
from openpyxl import load_workbook

# constants
curr_encoding = 'windows-1251'
path_input = 'input'
path_output = 'output'
table_ext = '.xlsx'
typo_f_name = 'table_tipology'
typo_splitter = '|'
typo_row_start = 2
typo_col_name = 1
typo_col_pattern = 2

def err(str):
    print(''.join(['Error. ', str]))
    sys.exit()

#open temlate of typology
typo_path = Path(path_input, typo_f_name).with_suffix(table_ext)

if not typo_path.is_file():
    err(''.join(['File "', str(typo_path), '" not exists.']))

typo_wb = load_workbook(typo_path)

typo_patterns = {}

row_end = typo_row_start
while (typo_wb.active.cell(row = row_end, column = typo_col_name).value != None or typo_wb.active.cell(row = row_end, column = typo_col_pattern).value != None):
    
    if typo_wb.active.cell(row = row_end, column = typo_col_name).value == None:
        err(''.join(['Typology name is nulled. Str [', str(row_end), '] of "', str(typo_path), '".']))

    if typo_wb.active.cell(row = row_end, column = typo_col_pattern).value == None:
        err(''.join(['Typology pattern is nulled. Str [', str(row_end), '] of "', str(typo_path), '".']))

    name = str(typo_wb.active.cell(row = row_end, column = typo_col_name).value).strip().capitalize()
    if typo_patterns.get(name) != None:
        err(''.join(['Typology name "', name, '" is repeated. Str [', str(row_end), '] of "', str(typo_path), '".']))

    patterns = str(typo_wb.active.cell(row = row_end, column = typo_col_pattern).value).split(typo_splitter)
    patterns = list(map(lambda x: str(x).strip().lower(), patterns))
    typo_patterns[name] = patterns.copy()

    row_end += 1
row_end -= 1

if row_end < 1:
    err('Typology wocabulary is empty in file "', str(typo_path), '".')

print('Typology file info:')
print('\tFile name: "', typo_path, '";')
print('\tActive sheet name: "', typo_wb.active.title, '";')
print('\tCount of typology name: ', str(row_end - (typo_row_start - 1)), '.')
print('')

#test output
for key in typo_patterns:
    print(key, '=', ' | '.join(typo_patterns[key]))

print('')
print('OK')
