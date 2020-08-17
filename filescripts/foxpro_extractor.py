import pandas as pd
import sys
import csv
import os
import re
from os.path import join, dirname
from dotenv import load_dotenv
from dbfread import DBF

def getExcel():
    data = fix_anomaly(getData())
    writeExcel()

def writeExcel():
    pass

def getData():
    dotenv_path = join(dirname(__file__), '.env')
    load_dotenv(dotenv_path)

    DBF_NAME = os.environ.get('DBF_NAME')
    CST_NAME = os.environ.get('CST_NAME')

    data = pd.read_csv(CST_NAME, header=None)
    
    string_dump = ''
    for i in range(data.shape[0]):
        string_dump += data.loc[i].values[0]
    splitted_dump = re.split('(SEKRET|PP-WEL|PLTHN)', string_dump)

    records = {'SEKRET':[], 'PP-WEL':[], 'PLTHN':[]}
    for i in range(len(splitted_dump)):
        if splitted_dump[i] in records.keys():
            records[splitted_dump[i]].append(splitted_dump[i-1])

    entries = []
    for key in records.keys():
        if '' in records[key]:
            records[key].remove('')
        for el in records[key]:
            entry = re.split('\s{2,}', el)
            if '' in entry:
                entry.remove('')
            entry.append(key)
            entries.append(entry)
    return entries
    # dbf = DBF(DBF_NAME)

    # with open(CST_NAME, 'r') as csv_file:
    #     reader = csv.reader(csv_file)
    #     for row in reader:
    #         print(row)    
        # writer = csv.writer(file)
        # writer.writerow(dbf.field_names)
        # for record in dbf:
        #     writer.writerow(list(record.values()))

def fix_anomaly(data):    
    cnt = 0
    cnt1 = 0
    a = 0
    # Check anomaly data
    anomalies = []
    for i in range(len(data)):    
        if len(data[i]) > 6:
            if len(data[i]) <= 7:
                anomalies.append(i)
                a += 1
            else:
                print(data[i], i)
            cnt += 1
        elif len(data[i]) < 6:
            cnt1 += 1
    # Fix anomaly data
    for i in range(len(data)):
        if len(data[i]) == 7:
            tmp = data[i].pop(2) + ' ' + data[i].pop(2)
            data[i].insert(2, tmp)

    # re-check
    # print('CHECKED')
    # for i in anomalies:
    #     print(data[i])

    print(a)
    print('total <6 = ', cnt1)
    print('total >6 = ', cnt)
    print('all = ', len(data))