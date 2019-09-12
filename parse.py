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
lists_f_name = 'lists'
log_f_name = 'log.txt'
typo_splitter = '|'

def err(str):
    print(''.join(['Error. ', str]))
    sys.exit()

#convert xmlx to dict
def exel_to_dict(_path : Path):
    
    result = dict()

    if not _path.is_file():
        err(''.join(['File "', str(_path), '" not exists.']))

    wb = load_workbook(_path)

    for sht in wb.sheetnames:
        # работаем отдельно с каждым листом
        result.update({sht: []})
        ws = wb[sht]

        # инициализируем поля
        keys = []
        for c in range(1, ws.max_column + 1):
            if ws.cell(1, c).value != '':
                keys.append(ws.cell(1, c).value)
            else:
                break
        # собираем данные из строк
        for r in range(2, ws.max_row + 1):
            values = [ws.cell(row=r,column=i).value for i in range(1,len(keys)+1)]
            if ''.join(str(i) for i in values) != '':
                result[sht].append(dict(zip(keys, values)))
            else:
                break
    return result

#open temlate of typology
typo_path = Path(path_input, typo_f_name).with_suffix(table_ext)
typo_obj = exel_to_dict(typo_path)

#convert typology objects
for i in typo_obj['table_tipology']:
    i['variants'] = list(map(lambda x: str(x).strip().lower(), i['variants'].split(typo_splitter)))

#open main lists
lists_path = Path(path_input, lists_f_name).with_suffix(table_ext)
lists_obj = exel_to_dict(lists_path)

#return typology from string
def get_typo(_str : str):
    _str = _str.strip().lower()
    for i in typo_obj['table_tipology']:
        for j in i['variants']:
            if _str.find(j) != -1:
                return i['alias']
    return ''

#add typology
for year, rows in lists_obj.items():
    prev_typo = None
    for i in rows:
        if i['description'] != None:
            prev_typo = get_typo(str(i['description']))
        if prev_typo == None:
            print('Lists error! page {p}, number {n} description is invalid!')
        else:
            i['type'] = prev_typo

log_path = Path(path_output, log_f_name)

#for example
with open(log_path, 'w') as f:
    print(typo_obj, file = f)
    print(lists_obj, file = f)

print('complete!')
