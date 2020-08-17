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
    writeExcel(data)

def writeExcel(data):
    with open('restored_data.csv', mode='w', newline='') as restored_data:
        writer = csv.writer(restored_data)
        writer.writerow(['No. Voucher', 'Tanggal+Account', 'Keterangan','Debet', 'Kredit', 'Divisi'])
        anomalies = []
        for el in data:
            if len(el) > 6:
                anomalies.append(el)
            else:
                writer.writerow(el)
        for el in anomalies:
            writer.writerow(el)

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

def fix_anomaly(data):    
    cnt = 0
    cnt1 = 0
    # Check anomaly data
    anomalies_less = []
    anomalies_more = []
    for i in range(len(data)):    
        if len(data[i]) > 6:
            if len(data[i]) <= 7:
                anomalies_more.append(i)
            cnt += 1
        elif len(data[i]) < 6:
            # print(data[i], i)
            anomalies_less.append(i)
            cnt1 += 1
    # Fix anomaly data
    for i in range(len(data)):
        if len(data[i]) == 7:
            tmp = data[i].pop(2) + ' ' + data[i].pop(2)
            data[i].insert(2, tmp)
        if len(data[i]) < 6:
            tmp = data[i].pop(0).split()
            for j in tmp[::-1]:
                data[i].insert(0, j)
            if len(data[i]) > 6 or (len(data[i]) == 6 and data[i][0] == 'AC'):
                tmp = data[i].pop(0) + '-' + data[i].pop(0)
                data[i].insert(0, tmp)
            if len(data[i]) < 6:
                tmp = re.split('(KK|KM)', data[i].pop(0))
                if len(tmp) > 2:
                    tmp2 = [tmp[0] + tmp[1], tmp[2]]
                for j in tmp2[::-1]:
                    data[i].insert(0, j)
    # re-check
    # print('CHECKED')
    anomalies_unique = []
    for i in range(len(data)):
        if len(data[i]) != 6:
            anomalies_unique.append(i)
    
    return data