import mysql.connector as mysqldb
import sys
import csv
import os
import pandas as pd
from os.path import join, dirname
from dotenv import load_dotenv

def getExcel():
    dotenv_path = join(dirname(__file__), '.env')
    load_dotenv(dotenv_path)

    HOST = os.environ.get('HOST')
    DB = os.environ.get('DB')
    USER = os.environ.get('USER')
    PASS = os.environ.get('PASS')
    
    connect = mysqldb.connect(host=HOST, database=DB, user=USER, password=PASS)
    cursor = connect.cursor()
    # GETTING THE TABLES    
    table_query = 'SHOW TABLES'
    cursor.execute(table_query)
    tables = [ i[0] for i in cursor.fetchall()] 
    
    # Records of data from mysql db
    records = {}
    for table in tables:
       query = 'SELECT * FROM ' + table
       cursor.execute(query)
       records[table] = pd.DataFrame(cursor.fetchall(), columns=cursor.column_names)
       loc = 'csv/'+table+'.csv'
       # Print to csv
       records[table].to_csv(loc, index=False, header=True)

    if (connect.is_connected()):
        connect.close()
        cursor.close()
