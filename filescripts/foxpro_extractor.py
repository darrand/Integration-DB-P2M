import pandas as pd
import sys
import os
import re
import datetime as dt
import csv
import xlwt
import openpyxl
import xlrd
from xlutils.copy import copy
from xlwt import Workbook
from os.path import join, dirname
from dotenv import load_dotenv
from dbfread import DBF

def getExcel():
    '''
    Pengambilan data untuk ditulis ke bentuk excel(xls) dan csv
    '''
    data = fix_anomaly(getData())
    writeExcel(data)

def writeExcel(data):
    '''
    Penulisan entry ke bentuk excel (xls) dan csv
    dengan format biasa untuk csv dan format sesuai journal untuk xls
    '''
    wb = Workbook()
    sheet = wb.add_sheet('DTJUR')
    header = ['No. Voucher', 'Tanggal','Account', 'Keterangan','Debet', 'Kredit', 'Divisi']
    data.insert(0, header)
    anomalies = []

    for i in range(len(data)):
        for j in range(len(data[i])):
            if len(data[i]) > 7:
                anomalies.append(data[i])
            else:
                sheet.write(i, j, data[i][j])

    wb.save('restored_data.xls')

    with open('restored_data.csv', mode='w', newline='') as restored_data:
        writer = csv.writer(restored_data)
        anomalies = []
        for el in data:
                writer.writerow(el)

    rb = xlrd.open_workbook(filename='template.xls')
    wb = copy(rb)

    s = wb.get_sheet(0)
    data.pop(0)
    for i in range(len(data)):
        entry = ['','Universitas Indonesia','','UKK-FT-UP2M IDR',str(data[i][1]),"'10407", "'00000000","'71",'','','',"'000","'000"
            ,data[i][4],data[i][5],str(data[i][1].strftime('%b'))+'-'+str(data[i][1].year)[2:], 'UKK-FT UP2M ' + data[i][0], data[i][3],'',data[i][3]
            ,'','','','','','','','J','']
        for j in range(len(entry)):
            s.write(i,j,entry[j])
    wb.save('template.xls')

def getData():
    '''
    Mengambil data dari parameter yang didefinisikan .env
    dengan slicing string sesuai dengan divisi jurnal (sekaligus penulisan salah yang ada),
    mengembalikan list dengan isi [data], divisi
    '''
    dotenv_path = join(dirname(__file__), '.env')
    load_dotenv(dotenv_path)

    DBF_NAME = os.environ.get('DBF_NAME')
    CST_NAME = os.environ.get('CST_NAME')

    data = pd.read_csv(CST_NAME, header=None)
    
    string_dump = ''
    for i in range(data.shape[0]):
        string_dump += data.loc[i].values[0]
    splitted_dump = re.split('(SEKRET|SEKRE|SKRET|PP-WEL|P-WEL|PTHN|PLTHN|PP-OTO|PP-BM)', string_dump)

    records = {'SEKRET':[], 'PP-WEL':[], 'PLTHN':[], 'PP-OTO':[], 'PP-BM':[]}
    for i in range(len(splitted_dump)):
        if splitted_dump[i] == 'SEKRE':
            splitted_dump[i] = 'SEKRET'
        if splitted_dump[i] == 'SKRET':
            splitted_dump[i] = 'SEKRET'
        if splitted_dump[i] == 'P-WEL':
            splitted_dump[i] = 'PP-WEL'
        if splitted_dump[i] == 'PTHN':
            splitted_dump[i] = 'PLTHN'
            
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
    '''
    Ditemukan banyak anomali data dari fungsi getData()
    data normal memiliki panjang 6, data anomali memiliki panjang tidak 6
    fungsi ini digunakan sebagai "pembersih" untuk data tersebut
    fungsi ini memisahkan dan memfilter data sesuai dengan panjangnya
    Setelah itu untuk setiap data pada posisi ke 2 terdapat atribut tanggal yang tergabung dengan atribut lain
    atribut tersebut dipisah dan diletakkan di posisi ke 2 dan ke 3 
    '''
    cnt = 0
    cnt1 = 0
    # Check anomaly data
    anomalies_less = []
    anomalies_more = []
    for i in range(len(data)):    
        # Pengecekan data > 6
        if len(data[i]) > 6:
            if len(data[i]) <= 7:
                anomalies_more.append(i)
            cnt += 1
        # Pengecekan data < 6
        elif len(data[i]) < 6:
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
    anomalies_unique = []
    for i in range(len(data)):
        if len(data[i]) != 6:
            anomalies_unique.append(i)
    # split date
    error_date = []
    for i in range(len(data)):
        if i in anomalies_unique:
            continue
        else:
            tmp = data[i].pop(1)
            try:
                date = dt.datetime.strptime(tmp[0:8], '%Y%m%d').date()
            except ValueError:
                error_date.append(i)
                data[i].insert(1, tmp)
                continue
            account = tmp[8:]
            data[i].insert(1, account)
            data[i].insert(1, date)
    # case khusus karena terdapat data yang salah tanggal
    for i in range(len(error_date)):
        tmp = ''
        if i == 0:
            data[error_date[i]][1] = '20141216120.103'
            tmp = data[error_date[i]].pop(1)
        elif i == 1: 
            data[error_date[i]][1] = '20171214110.101'
            tmp = data[error_date[i]].pop(1)
        elif i == 2:
            data[error_date[i]][1] = '20190109110.101'
            tmp = data[error_date[i]].pop(1)
        elif i == 4:
            data[error_date[i]][1] = '20160423400.105'
            tmp = data[error_date[i]].pop(1)
        else:
            tmp = data[error_date[i]].pop(0).split()
            data[error_date[i]].insert(0, tmp.pop(0))
            tmp = tmp[0]
        date = dt.datetime.strptime(tmp[0:8], '%Y%m%d').date()
        account = tmp[8:]
        data[error_date[i]].insert(1, account)
        data[error_date[i]].insert(1, date)
    # Case khusus untuk beberapa data yang tidak ter split normal
    for i in range(len(anomalies_unique)):
        if i == 0:
            broken_data = data.pop(anomalies_unique[i])
            tmp = []
            tmp_wrapper = []
            cnt = 0
            for j in range(len(broken_data)):
                tmp.append(broken_data[j])
                cnt += 1
                if cnt % 5 == 0:
                    tmp.append('SEKRET')
                    tmp_wrapper.append(tmp)
                    tmp = []
            for k in tmp_wrapper:
                split_this = k.pop(1)
                date = dt.datetime.strptime(split_this[0:8], "%Y%m%d").date()
                acc = split_this[8:]
                k.insert(1, acc)
                k.insert(1, date)
                data.insert(anomalies_unique[i], k)
    # Pembenaran tanggal untuk case khusus diatas
    cnt = 0
    for i in data:
        if len(i) > 7:
            if cnt < 4:
                i.pop(0)
                i.pop(0)
                tmp = i.pop(1)
                date = dt.datetime.strptime(tmp[0:8], "%Y%m%d").date()
                acc = tmp[8:]
                i.insert(1, acc)
                i.insert(1, date)
            elif cnt == 4:
                i.pop(0)
                i.pop(0)
                i.pop(0)
                i.pop(0)
                tmp = i.pop(1)
                date = dt.datetime.strptime(tmp[0:8], "%Y%m%d").date()
                acc = tmp[8:]
                i.insert(1, acc)
                i.insert(1, date)
            else:
                tmp = i.pop(3) + ' ' + i.pop(3)
                i.insert(3, tmp)
            cnt += 1
    return data