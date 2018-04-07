# -*- coding: utf-8 -*-
# Заполняем неправильными, отсутствующими СНИЛСАми

import sys
import datetime
import time
import csv
from mysql.connector import MySQLConnection, Error

from lib import read_config, lenl, s_minus, s, l, filter_rus_sp, filter_rus_minus

TABLE = 'aug_out'

def checksum(snils_dig):                         # Вычисляем 2 последних цифры СНИЛС по первым 9-ти
    snils = '{0:09d}'.format(snils_dig)

    def snils_csum(snils):
        k = range(9, 0, -1)
        pairs = zip(k, [int(x) for x in snils.replace('-', '').replace(' ', '')])
        return sum([k * v for k, v in pairs])

    csum = snils_csum(snils)

    while csum > 101:
        csum %= 101
    if csum in (100, 101):
        csum = 0

    return csum


dbconfig = read_config(filename='gen_snils.ini', section='mysql')
dbconn = MySQLConnection(**dbconfig)


dbcursor = dbconn.cursor()
dbcursor.execute('SELECT min(number_dub) FROM ' + TABLE + ' WHERE number_dub != 0 AND (mobile0 IS NOT NULL '
                  'OR mobile1 IS NOT NULL OR mobile2 IS NOT NULL);')
rows = dbcursor.fetchall()
start_snils = int('{0:011d}'.format(rows[0][0])[:-2])
dbcursor = dbconn.cursor()
dbcursor.execute('SELECT count(*) FROM ' + TABLE + ' WHERE number_dub = 0 AND (mobile0 IS NOT NULL '
                 'OR mobile1 IS NOT NULL OR mobile2 IS NOT NULL);')
rows = dbcursor.fetchall()
count_snils = rows[0][0]

while count_snils > 0:
    start_snils -= 1
    checksum_snils = checksum(start_snils)
    cached_snils = []
    for i in range(0,99):
        if i != checksum_snils:
            full_snils = start_snils * 100 + i
            dbcursor = dbconn.cursor()
            dbcursor.execute('SELECT `number` FROM clients WHERE `number` = %s', (full_snils,))
            rows = dbcursor.fetchall()
            if len(rows) == 0:
                cached_snils.append((full_snils,))
                count_snils -= 1
    dbcursor = dbconn.cursor()
    dbcursor.executemany('UPDATE ' + TABLE + ' SET number_dub = %s WHERE number_dub = 0 AND (mobile0 IS NOT NULL '
                         'OR mobile1 IS NOT NULL OR mobile2 IS NOT NULL) LIMIT 1;', cached_snils)
    dbconn.commit()
    print(datetime.datetime.now().strftime("%H:%M:%S"), start_snils)
dbconn.close()
