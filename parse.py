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

#convert xmlx to dict
def exel_to_dict(_path : Path, _dict : dict):
    if not _path.is_file():
        err(''.join(['File "', str(_path), '" not exists.']))

    wb = load_workbook(typo_path)

    for sht in wb.sheetnames:
        # работаем отдельно с каждым листом
        _dict.update({sht: []})
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
                _dict[sht].append(dict(zip(keys, values)))
            else:
                break
    return

#open temlate of typology
typo_path = Path(path_input, typo_f_name).with_suffix(table_ext)

typo_obj = dict()

exel_to_dict(typo_path, typo_obj)

#for example
with open('test.txt', 'w') as f:
    print(typo_obj, file = f)

os.system('pause')
exit(0)
