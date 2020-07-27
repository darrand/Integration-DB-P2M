import mysql.connector as mysqldb
import sys
import csv
import os
from os.path import join, dirname
from dotenv import load_dotenv

def getExcel():
    dotenv_path = join(dirname(__file__), '.env')
    load_dotenv(dotenv_path)

    HOST = os.environ.get('HOST')
    DB = os.environ.get('DB')
    USER = os.environ.get('USER')
    PASS = os.environ.get('PASS')
    TABLE = 'indonesia'

    connect = mysqldb.connect(host=HOST, database=DB, user=USER, password=PASS)
    query = 'SELECT * FROM ' + TABLE
    cursor = connect.cursor()
    cursor.execute(query)
    records = cursor.fetchall()
    print(
        records
    )

    if (connect.is_connected()):
        connect.close()
        cursor.close()
