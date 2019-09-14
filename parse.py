# -*- coding: utf-8 -*-

import sys
import os
import random
from timeit import default_timer
from pathlib import Path
from openpyxl import load_workbook
from openpyxl import Workbook

time_start = default_timer()

# constants
curr_encoding = 'windows-1251'
path_input = 'input'
table_ext = '.xlsx'
typo_f_name = 'table_tipology'
lists_f_name = 'lists'
out_f_name = 'lists_out'
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
        print('{}..'.format(str(sht)))
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

print('typology parsing..')

#open temlate of typology
typo_path = Path(path_input, typo_f_name).with_suffix(table_ext)
typo_obj = exel_to_dict(typo_path)

#convert typology objects
for i in typo_obj['table_tipology']:
    i['variants'] = list(map(lambda x: str(x).strip().lower(), i['variants'].split(typo_splitter)))

print('lists parsing..')

#open main lists
lists_path = Path(path_input, lists_f_name).with_suffix(table_ext)
lists_obj = exel_to_dict(lists_path)

#return typology from string
def get_typo(_str : str):
    _str = _str.strip().lower()
    for i in typo_obj['table_tipology']:
        for j in i['variants']:
            result = True
            for k in j.split():
                if len(k) > 2 and _str.find(k) < 0:
                    result = False
                    break
            if result == True:
                return i['alias']
    return None

# get horizon
def get_horizon(_str : str):
    _str = _str.strip().lower()
    for i in typo_obj['horizon']:
        if _str == str(i['horizon']).strip().lower():
            return [int(i['min']), int(i['max'])]
    return get_horizon('default')

# random generator init
random.seed(1024)

# open logfile
logfile = open(Path(path_input, log_f_name), 'w', encoding = curr_encoding)

print('typologies and coordinates adding..')

# add typology and coordinates
for year, rows in lists_obj.items():
    print('{}..'.format(str(year)))

    q_ltrs = 'абвгдежзиклмнопрстуфхцчшщъыьэюя'
    prev_ltr = None
    prev_num = None
    prev_hor = None
    prev_typo = None

    description_str = None

    for n, i in enumerate(rows):
        first_row = bool(n == 0)
        err_str = 'Lists error! page {p}, row {n} '.format(p = str(year), n = str(n))
        
        # set typology
        if i['description'] != None:
            prev_typo = get_typo(str(i['description']))
            description_str = i['description']
        if prev_typo == None:
            if i['description'] == None and not first_row: continue
            print('{}description is invalid! "{}"'.format(err_str, str(i['description'])), file = logfile)
        else:
            i['type'] = prev_typo
            i['description'] = description_str
        
        # add quad letter
        if i['quad_letter'] != None:
            ql_st = str(i['quad_letter']).strip().lower()
            if len(ql_st) < 1:
                prev_ltr = None
            else:
                if q_ltrs.find(ql_st) < 0: ql_st = ql_st[::-1] # slice string
                prev_ltr = [q_ltrs.find(ql_st)]
                prev_ltr.append(prev_ltr[0] + len(ql_st) - 1)
                if prev_ltr[0] < 0 or prev_ltr[1] < 0:
                    prev_ltr = None
        if prev_ltr == None:
            if i['quad_letter'] == None and not first_row: continue
            print('{}quad letter is invalid! "{}"'.format(err_str, str(i['quad_letter'])), file = logfile)

        # add quad number
        if i['quad_num'] != None:
            qn_ls = str(i['quad_num']).split('-')
            if 1 <= len(qn_ls) <= 2:
                try:
                    prev_num = [int(str(qn_ls[0]).strip().lower())]
                    if len(qn_ls) < 2:
                        prev_num.append(prev_num[0])
                    else:
                        prev_num.append(int(str(qn_ls[1]).strip().lower()))
                    if (prev_num[0] > prev_num[1]):
                        prev_num[1], prev_num[0] = prev_num[0], prev_num[1]
                except:
                    prev_num = None
            else:
                prev_num = None
        if prev_num == None:
            if i['quad_num'] == None and not first_row: continue
            print('{}quad number is invalid! "{}"'.format(err_str, str(i['quad_num'])), file = logfile)

        # add horizon
        if i['horizon'] != None:
            prev_hor = get_horizon(str(i['horizon']))
        if prev_hor == None:
            if i['horizon'] == None and not first_row: continue
            print('{}horizon is invalid! "{}"'.format(err_str, str(i['horizon'])), file = logfile)

        # set coord as is (relative quad a0), cm
        if prev_ltr != None and prev_num != None and prev_hor != None:
            i['coord'] = '{x}:{y}:{z}'.format(
                x = str(random.randint(prev_ltr[0] * 100, prev_ltr[1] * 100 + 100)),
                y = str(random.randint(prev_num[0] * 100, prev_num[1] * 100 + 100)),
                z = str(0 - random.randint(prev_hor[0], prev_hor[1])))

logfile.close()

print('saving results..')

# create new lists and save
wr_wb = Workbook()

for year, rows in lists_obj.items():
    print('{}..'.format(str(year)))
    sheet = wr_wb.create_sheet(title = str(year))
    sheet.append(list(rows[0].keys()))
    for r in rows:
        sheet.append(list(r.values()))

wr_wb.save(Path(path_input, out_f_name).with_suffix(table_ext))

print('complete!')
print('{} sec'.format(str(default_timer() - time_start)))
