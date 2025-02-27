import mysql.connector as mysqldb
import sys
import csv
import os
import pandas as pd
from os.path import join, dirname
from dotenv import load_dotenv

dotenv_path = join(dirname(__file__), '.env')
load_dotenv(dotenv_path)
HOST = os.environ.get('HOST')
DB = os.environ.get('DB')
USER = os.environ.get('USER')
PASS = os.environ.get('PASS')

def getExcel(name, header=True):
    excel = name
    records = None
    if header:
        records = pd.read_csv(excel)
    else:
        records = pd.read_csv(excel, header=None)
    return records

def master_peserta():
    data = getExcel('master_peserta.csv')
    inject(data, 'master_peserta')

def inject(records, query):    
    connect = mysqldb.connect(host=HOST, database=DB, user=USER, password=PASS)
    cursor = connect.cursor()
    data = records
    headers = tuple(data.columns.values)

    tmp_query = 'INSERT INTO ' + query + ' ' + str(headers).replace("'", "`") + ' VALUES '
    for i in range(data.shape[0]):
        new_entry = data.loc[i]
        tmp_query += '\n' + str(tuple(new_entry)).replace('nan', 'NULL') + ','
    
    query = tmp_query[:len(tmp_query)-1] + ';'
    cursor.execute(query)
    connect.commit()
    print(cursor.rowcount, "rows inserted.")

    if (connect.is_connected()):
        connect.close()
        cursor.close()
