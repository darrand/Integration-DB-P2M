import pandas as pd
import sys
import csv
import os
from os.path import join, dirname
from dotenv import load_dotenv
from dbfread import DBF

dotenv_path = join(dirname(__file__), '.env')
load_dotenv(dotenv_path)

DBF_NAME = os.environ.get('DBF_NAME')
print(DBF_NAME)

dbf = DBF(DBF_NAME)

with open('db1.csv', 'w', newline='') as file:    
    writer = csv.writer(file)
    writer.writerow(dbf.field_names)
    for record in dbf:
        writer.writerow(list(record.values()))