#!/usr/bin/python

from ast import arg
from cmath import sin
import sys
import os
from enum import Enum
from chardet import detect
from datetime import datetime
from dateutil.relativedelta import relativedelta
from calendar import monthrange
import xlrd
import xlwings
import csv
import string

def main(argv):
    if not argv:
        main_help()
    
    match argv[0]:
        case '--learn':
            main_learn()
        case '--banking':
            main_banking(argv[1:])
        case '--transactions':
            main_transactions(argv[1:])
        case '--simple':
            main_simple(argv[1:])
        case '--62':
            main_sixtytwo(argv[1:])
        case '--dzo':
            main_dzo(argv[1:])
        case _:
            print('Unknown procedure')
            print()
            main_help()

def main_learn():
    class Person:
        def __init__(self, name, job, since):
            self.name = name
            self.job = job
            self.since = since

        def __str__(self):
            return "name=" + self.name + "; job=" + self.job + "; since=" + str(self.since)

    nastya = Person('Настя', 'Юрист', 2021)
    lilya = Person('Лилия', 'Главбух', 2020)
    print(nastya)
    print(lilya)

def main_help():
    print('Use for comparing Jivo, Alfa, and YooKassa:')
    print('python helper.py --banking jivo.xlsx alfa.csv yookassa.csv')
    print()
    print('Use for comparing internal documents about realisation:')
    print('python helper.py --transactions primary.xls copy.xls')
    print()
    print('Use for comparing internal documents by 62:')
    print('python helper.py --62 62_1.xls 62_2.xls')
    print()
    print('Use for comparing internal documents by simple:')
    print('python helper.py --simple doc.xls')
    print()
    print('Use for generating the DZO block:')
    print('python helper.py --dzo ru doc.xls payments.csv')
    print()

class Common_FileMeta:
    def __init__(self, kind, name, encoding):
        self.kind = kind
        self.name = name
        self.encoding = encoding

class Common_FileKind(Enum):
    UNKNOWN = 0
    JIVO = 1
    ALFABANK = 2
    YOOKASSA = 3

def common_recognize_file(name):
    with open(name, 'rb') as file:
        data = file.read()
        encoding = detect(data)['encoding'] 

    if name.endswith('.xlsx'):
        workbook = xlrd.open_workbook(name)
        worksheet = workbook.sheet_by_index(0)
        marker = worksheet.cell_value(0, 0)

        if marker.find('Есть файлы') > -1:
            return Common_FileMeta(Common_FileKind.JIVO, name, encoding)
        else :
            return Common_FileMeta(Common_FileKind.UNKNOWN, name, encoding)
    elif name.endswith('.csv'):
        with open(name, encoding=encoding) as file:
            document = csv.reader(file, delimiter=';')
            marker = list(next(document))[0]

            if marker.find('Наименование предприятия') > -1:
                return Common_FileMeta(Common_FileKind.ALFABANK, name, encoding)
            elif marker.find('ЮKassa') > -1:
                return Common_FileMeta(Common_FileKind.YOOKASSA, name, encoding)
            else:
                return Common_FileMeta(Common_FileKind.UNKNOWN, name, encoding)
    else:
        return None

def common_calc_date_diff(date_fst, date_snd):
    delta = relativedelta(date_snd, date_fst)
    return delta.years * 12 + delta.months

def common_excel_formula(lang, formula):
    packages = {
        'ru': [
            ('DATEVALUE', 'ДАТАЗНАЧ'),
            ('INDIRECT', 'ДВССЫЛ'),
            ('DATEDIF', 'РАЗНДАТ'),
            ('EOMONTH', 'КОНМЕСЯЦА'),
            ('ISBLANK', 'ЕПУСТО'),
            ('ADDRESS', 'АДРЕС'),
            ('COLUMN', 'СТОЛБЕЦ'),
            ('ROW', 'СТРОКА'),
            ('DAY', 'ДЕНЬ'),
            ('SUM', 'СУММ'),
            ('IF', 'ЕСЛИ'),
        ]
    }

    package = packages.get(lang, list())
    for (original_name, need_name) in package:
        formula = formula.replace(original_name, need_name)

    return formula

def common_excel_comment(lang, comment):
    signatures = {
        'ru': 'Ч'
    }

    signature = signatures.get(lang, 'N')
    return '%s("%s") + ' % (signature, comment)

def main_banking(argv):
    jivo_meta = None
    alfabank_meta = None
    yookassa_meta = None

    for name in argv:
        meta = common_recognize_file(name)

        if not meta:
            print('Not readable:', name)
        elif meta.kind == Common_FileKind.JIVO:
            print('Jivo:', name)
            jivo_meta = meta
        elif meta.kind == Common_FileKind.ALFABANK:
            print('AlfaBank:', name)
            alfabank_meta = meta
        elif meta.kind == Common_FileKind.YOOKASSA:
            print('YooKassa:', name)
            yookassa_meta = meta
        else:
            print('Not recognized:', name)
    
    print()
    
    alfabank_orders = main_banking_find_alfabank_orders(alfabank_meta)
    yookassa_orders = main_banking_find_yookassa_orders(yookassa_meta)
    main_banking_compare_orders(jivo_meta, alfabank_orders, yookassa_orders)

def main_banking_find_alfabank_orders(meta):
    with open(meta.name, encoding=meta.encoding) as file:
        document = csv.reader(file, delimiter=';')

        orders = set()
        for row in document:
            id = row[15].split('.')[0]
            if not id:
                continue

            orders.add(int(id))

        return orders

def main_banking_find_yookassa_orders(meta):
    with open(meta.name, encoding=meta.encoding) as file:
        document = csv.reader(file, delimiter=';')

        orders = set()
        for row in document:
            id = row[7]
            if not id:
                continue
            elif id == 'Описание заказа':
                continue
            
            status = row[3]
            if not status == 'Оплачен':
                continue

            orders.add(int(id))

        return orders

def main_banking_compare_orders(meta, alfabank_expected_orders, yookassa_expected_orders):
    workbook = xlrd.open_workbook(meta.name)
    worksheet = workbook.sheet_by_index(0)

    alfabank_found_orders = set()
    yookassa_found_orders = set()
    total_diff_num = 0

    for i in range(1, worksheet.nrows):
        id = worksheet.cell_value(i, 2)
        if not id:
            continue

        extra = worksheet.cell_value(i, 13)
        if extra.find('ALFA-BANK') > -1:
            alfabank_found_orders.add(int(id))
        elif extra.find('YANDEX-JS') > -1:
            yookassa_found_orders.add(int(id))
    
    alfabank_ours_diff = alfabank_expected_orders.difference(alfabank_found_orders)
    total_diff_num += len(alfabank_ours_diff)
    if alfabank_ours_diff:
        print('AlfaBank, these orders were found within BANK system only:')
        for id in sorted(alfabank_ours_diff):
            print('-', id)
        print()
    
    alfabank_theirs_diff = alfabank_found_orders.difference(alfabank_expected_orders)
    total_diff_num += len(alfabank_theirs_diff)
    if alfabank_theirs_diff:
        print('AlfaBank, these orders were found within JIVO system only:')
        for id in sorted(alfabank_theirs_diff):
            print('-', id)
        print()

    yookassa_ours_diff = yookassa_expected_orders.difference(yookassa_found_orders)
    total_diff_num += len(yookassa_ours_diff)
    if yookassa_ours_diff:
        print('YooKassa, these orders were found within BANK system only:')
        for id in sorted(yookassa_ours_diff):
            print('-', id)
        print()
    
    yookassa_theirs_diff = yookassa_found_orders.difference(yookassa_expected_orders)
    total_diff_num += len(yookassa_theirs_diff)
    if yookassa_theirs_diff:
        print('YooKassa, these orders were found within JIVO system only:')
        for id in sorted(yookassa_theirs_diff):
            print('-', id)
        print()

    if not total_diff_num:
        print('No issues found')
        print()

def main_transactions(argv):
    first_orders = main_transactions_find_orders(argv[0])
    second_orders = main_transactions_find_orders(argv[1])

    first_diff = first_orders.difference(second_orders)
    second_diff = second_orders.difference(first_orders)
    total_diff = first_diff.union(second_diff)

    if total_diff:
        print('Please check these orders:')
        for id in sorted(total_diff):
            print('-', id.lstrip('0'))
        print()
    else:
        print('No issues found')
        print()

def main_transactions_find_orders(name):
    workbook = xlrd.open_workbook(name)
    worksheet = workbook.sheet_by_index(0)

    orders = set()

    for i in range(1, worksheet.nrows):
        id = worksheet.cell_value(i, 2)
        if not id:
            continue

        if id.startswith('^'):
            continue

        orders.add(id)
    
    return orders

def main_simple(argv):
    first_values = main_simple_find_values(argv[0], 0)
    second_values = main_simple_find_values(argv[0], 1)

    first_diff = first_values.difference(second_values)
    second_diff = second_values.difference(first_values)
    total_diff = first_diff.union(second_diff)

    if total_diff:
        print('Please check these values:')
        for value in sorted(total_diff):
            print('-', value)
        print()
    else:
        print('No difference found')
        print()

def main_simple_find_values(name, column):
    workbook = xlrd.open_workbook(name)
    worksheet = workbook.sheet_by_index(0)

    values = set()

    for i in range(1, worksheet.nrows):
        value = worksheet.cell_value(i, column)
        if not value:
            continue

        values.add(value)
    
    return values

def main_sixtytwo(argv):
    first_customers = main_sixtytwo_find_customers(argv[0])
    second_customers = main_sixtytwo_find_customers(argv[1])

    ids = set(first_customers.keys()).union(second_customers.keys())
    for id in ids:
        (first_name, first_info) = first_customers.get(id, (None, None))
        (second_name, second_info) = second_customers.get(id, (None, None))

        if first_info == second_info:
            continue
        elif not first_info:
            print('Exclusive in %s: "%s"' % (os.path.basename(argv[1]), second_name))
        elif not second_info:
            print('Exclusive in %s: "%s"' % (os.path.basename(argv[0]), first_name))
        elif first_name != second_name:
            print('Differ by names: "%s" & "%s"' % (first_name, second_name))
        else:
            print('Differ by amounts: "%s"' % (first_name,))

def main_sixtytwo_find_customers(name):
    workbook = xlrd.open_workbook(name)
    worksheet = workbook.sheet_by_index(0)

    customers = dict()

    for ir in range(9, worksheet.nrows - 1):
        name = worksheet.cell_value(ir, 0)
        inn = worksheet.cell_value(ir, 2)
        id = inn if inn else name
        # print(id)

        info = str()
        for ic in range(3, worksheet.row_len(ir)):
            info += str(worksheet.cell_value(ir, ic)) + ";"
        # print(info)

        customers[id] = (name, info)
    
    return customers

class MainBanking_Dzo:
    def __init__(self, activated, money, since, till):
        self.activated = activated
        self.money = money
        self.since = since
        self.till = till
    
    def __repr__(self):
        return "'%d: %s:%s @%s'" % (self.money, self.since, self.till, self.activated)

def main_dzo(argv):
    lang = argv[0]

    transactions = main_dzo_read_source(argv[2])
    # print(transactions)

    workbook = xlwings.Book(argv[1])
    for sheet in workbook.sheets:
        if sheet.name != 'сбербизнессофт':
            continue
        else:
            worksheet = sheet
    
    if not worksheet:
        print('Worksheet not found')
        return

    (previous_cell, initial_cell) = main_dzo_find_initial_cell(worksheet)
    match input('Going to place the data starting the line #%d: (y)es or (n)o? ' % (initial_cell.row,)):
        case 'y' | 'yes' | 'д' | 'да' | '1':
            print()
            (group_since, group_till) = main_dzo_migrate_from_source(worksheet, initial_cell, transactions)
            date_anchor = main_dzo_ensure_date_headers(worksheet, group_since, group_till)
            main_dzo_fill_matrix(lang, worksheet, previous_cell, initial_cell, transactions, date_anchor)
        case _:
            return

def main_dzo_read_source(name):
    transactions = dict()

    meta = common_recognize_file(name)
    with open(meta.name, encoding=meta.encoding) as file:
        source = csv.reader(file, delimiter=';')
        for row in source:
            if row[2] != 'SBS':
                continue
            
            activated = datetime.strptime(row[1], '%Y-%m-%d').date()
            since = datetime.strptime(row[4], '%Y-%m-%d').date()
            till = datetime.strptime(row[5], '%Y-%m-%d').date()
            transactions[int(row[0])] = MainBanking_Dzo(activated, float(row[3]), since, till)

    return transactions

def main_dzo_find_initial_cell(worksheet):
    for ir in range(worksheet.used_range.last_cell.row, 0, -1):
        cell = worksheet.range((ir, 1))
        if cell.value == 'ИТОГО':
            return (worksheet.range("%d:%d" % (ir, ir)), cell.offset(3, 0))
    
    return (None, worksheet.range((3, 1)))

def main_dzo_migrate_from_source(worksheet, initial_cell, transactions):
    items = sorted(transactions.items())
    items_num = len(items)

    group_since = datetime.strptime('9999-01-01 00:00:00', '%Y-%m-%d %H:%M:%S').date()
    group_till = datetime.strptime('0001-01-01 00:00:00', '%Y-%m-%d %H:%M:%S').date()

    for (index, (id, meta)) in enumerate(items):
        group_since = min(group_since, meta.since)
        group_till = max(group_till, meta.till)

        worksheet.range((initial_cell.row + index, 1)).value = meta.activated.strftime('%d.%m.%Y')
        worksheet.range((initial_cell.row + index, 2)).value = id
        worksheet.range((initial_cell.row + index, 3)).value = meta.money
        worksheet.range((initial_cell.row + index, 4)).value = meta.since.strftime('%d.%m.%Y')
        worksheet.range((initial_cell.row + index, 5)).value = meta.till.strftime('%d.%m.%Y')
    
    return (group_since, group_till)

def main_dzo_ensure_date_headers(worksheet, since, till):
    first_date_cell = worksheet.range('G1')

    if not first_date_cell.value:
        first_date_cell.value = since.strftime('%Y-%m-%d %H:%M:%S')
    
    if first_date_cell.value.toordinal() > since.toordinal():
        print('First cell is too late')
        return
    
    month_diff = common_calc_date_diff(first_date_cell.value, till)
    for ic in range(first_date_cell.column, first_date_cell.column + month_diff):
        cell = worksheet.range((1, ic))
        if not cell.value:
            date = first_date_cell.value + relativedelta(months = ic)
            cell.value = date.strftime('%Y-%m-%d %H:%M:%S')
    
    return first_date_cell

def main_dzo_fill_matrix(lang, worksheet, previous_cell, initial_cell, transactions, date_anchor):
    row_max = initial_cell.row + len(transactions)
    column_max = date_anchor.column

    def _formula_for_matrix_cell():
        cell_map = {
            'cell_money': 'INDIRECT(ADDRESS(ROW(); 3; 3))',
            'cell_since': 'INDIRECT(ADDRESS(ROW(); 4; 3))',
            'cell_till': 'INDIRECT(ADDRESS(ROW(); 5; 3))',
            'cell_month': 'INDIRECT(ADDRESS(1; COLUMN(); 2))',
        }

        condition_map = {
            'condition_before': '%(cell_month)s < EOMONTH(%(cell_since)s; -1)' % cell_map,
            'condition_first': '%(cell_month)s <= %(cell_since)s' % cell_map,
            'condition_last': '%(cell_till)s <= EOMONTH(%(cell_month)s; 0)' % cell_map,
            'condition_after': '%(cell_month)s > EOMONTH(%(cell_till)s; 0)' % cell_map,
        }

        value_map = {
            'value_money': common_excel_comment(lang, "Сумма счёта:") + '%(cell_money)s' % cell_map,
            'value_duration': common_excel_comment(lang, "Колво дней в периоде:") + 'DATEDIF(%(cell_since)s; %(cell_till)s; "d") + 1' % cell_map,
            'value_before': common_excel_comment(lang, "Период ещё не начался:") + '0' % cell_map,
            'value_first': common_excel_comment(lang, "Первый месяц периода, частичное присутствие:") + 'DAY(EOMONTH(%(cell_month)s; 0)) - DAY(%(cell_since)s) + 1' % cell_map,
            'value_middle': common_excel_comment(lang, "Промежуточный месяц периода, полное присутствие:") + 'DAY(EOMONTH(%(cell_month)s; 0))' % cell_map,
            'value_last': common_excel_comment(lang, "Последний месяц периода, частичное присутствие:") + 'DAY(%(cell_till)s)' % cell_map,
            'value_after': common_excel_comment(lang, "Период уже закончился:") + '0' % cell_map,
        }

        formula_map = {
            'formula_month_usage': common_excel_comment(lang, "Колво дней периода в месяце:") + 'IF(%(condition_before)s; %(value_before)s; IF(%(condition_after)s; %(value_after)s; IF(%(condition_first)s; %(value_first)s; IF(%(condition_last)s; %(value_last)s; %(value_middle)s))))' % {**condition_map, **value_map}
        }

        return common_excel_formula(lang, '= (%(value_money)s) / (%(value_duration)s) * (%(formula_month_usage)s)' % {**value_map, **formula_map})

    def _formula_for_inner_result():
        cell_map = {
            'cell_first': 'INDIRECT(ADDRESS(%d; COLUMN(); 2))' % (initial_cell.row),
            'cell_last': 'INDIRECT(ADDRESS(%d; COLUMN(); 2))' % (initial_cell.row + len(transactions) - 1),
        }

        return common_excel_formula(lang, '= SUM(%(cell_first)s:%(cell_last)s)' % cell_map)

    def _formula_for_outer_result():
        cell_map = {
            'cell_month': 'INDIRECT(ADDRESS(1; COLUMN(); 2))',
            'cell_previous': 'INDIRECT(ADDRESS(%d; COLUMN(); 2))' % (initial_cell.row - 3),
            'cell_current': 'INDIRECT(ADDRESS(%d; COLUMN(); 2))' % (initial_cell.row + len(transactions)),
        }

        return common_excel_formula(lang, '= IF(ISBLANK(%(cell_month)s); 0; SUM(%(cell_previous)s; %(cell_current)s))' % cell_map)

    print('Formula for matrix cells:')
    print(_formula_for_matrix_cell())
    print()
    print('Formula for inner result:')
    print(_formula_for_inner_result())
    print()
    print('Formula for outer result:')
    print(_formula_for_outer_result())
    print()

    worksheet.range((row_max, 1)).value = 'Итого'

    finish_cell = initial_cell.offset(len(transactions) + 2, 0)
    previous_cell.copy(finish_cell)
    finish_cell.offset(0, 0).value = 'ИТОГО ???'
    for ic in range(date_anchor.column - 1, worksheet.used_range.last_cell.column):
        finish_cell.offset(0, ic).value = None

    for ir in range(initial_cell.row, row_max):
        id = worksheet.range((ir, 2)).value
        meta = transactions[id]

        offset = common_calc_date_diff(date_anchor.value, meta.since)
        months = common_calc_date_diff(meta.since, meta.till) + 1

        for ic in range(0, months):
            column = date_anchor.column + offset + ic
            column_max = max(column_max, column)
            worksheet.range((ir, column)).value = _formula_for_matrix_cell()
        
    for ic in ([2] + [*range(date_anchor.column, column_max)]):
        initial_cell.offset(len(transactions), ic).value = _formula_for_inner_result()
    
    for ic in range(date_anchor.column - 1, worksheet.used_range.last_cell.column):
        finish_cell.offset(0, ic).value = _formula_for_outer_result()

if __name__ == '__main__':
    main(sys.argv[1:])
