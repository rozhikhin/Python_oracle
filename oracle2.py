import argparse
import cx_Oracle
from openpyxl import Workbook

def exec_query(query, cursor):
    print(query)
    query = query.lower()
    if query.startswith('select'):
        cursor.execute(query)
        res = cursor.fetchall()
    else:
        cursor.execute(query)
        res = cursor.execute("COMMIT")
    print(res)
    return res


parser = argparse.ArgumentParser()
parser.add_argument("file1", help='file1')
parser.add_argument("file2", help='file2')
args = parser.parse_args()
file1 = args.file1
file2 = args.file2


params = dict()
# Читаем данные для подключения из файла
with open('settings.ini', 'r')  as ini:
    lines = ini.readlines()
    for line in lines:
        param, value = line.split('=')
        param, value = param.strip(), value.strip()
        params[param] = value


# Имя Файла excel с результатами
excel_file_name = 'result.xlsx'
file_log = 'log.txt'

# Подключаемся к базе
con = cx_Oracle.connect(params['db_user'], params['db_password'], '{}:{}/{}'.format(params['db_host'], params['port'], params['database_name']))
cur = con.cursor()



all_queries = []


# Читаем файл с запросами
with open(file1, 'r') as f:
    for line in f:
        queries = line.split(';')
        for q in queries:
            all_queries.append(q)

# В цикле выполняем запросы
with open(file_log, 'a+') as log:
    for query in all_queries:
        if len(query) < 3:
            continue
        try:
            exec_query(query, cur)
            log.write('Success')
        except cx_Oracle.DatabaseError as e:
            # если ошибка, то запрос и ошибку записываем в лог
            error, = e.args
            str_err = '{}, {}'.format(error.code, error.message)
            log.write(str_err)


# Создаем файл Excel
wb = Workbook()
ws = wb.active

with open(file2, 'r') as f2:
    for line in f2:
        try:
            res = exec_query(line, cur)
            for s in res:
                ws.append(s)
        except cx_Oracle.DatabaseError as e:
            # если ошибка, то запрос и ошибку записываем в лог
            error, = e.args
            str_err = '{}, {}'.format(error.code, error.message)
            print(str_err)

# сохраняем файл excel
wb.save(excel_file_name)

# закрываем базу
cur.close()
con.close()




