import pandas as pd
import sys
import csv
import os
import re
from os.path import join, dirname
from dotenv import load_dotenv
from dbfread import DBF

def getExcel():
    dotenv_path = join(dirname(__file__), '.env')
    load_dotenv(dotenv_path)

    DBF_NAME = os.environ.get('DBF_NAME')
    CST_NAME = os.environ.get('CST_NAME')

    data = pd.read_csv(CST_NAME, header=None)
    parsed = []
    for i in range(data.shape[0]):
        example = data.loc[i].values[0] 
        outer = example.split("*")
        outer.pop(0)
        
        # if i > 136:
        #     print(outer)
        for el in outer:
            inner = re.split("\s{2,}|SEKRET[\\\w\d]{2,}|SEKRET ", el)
            parsed.append(inner)
    for i in parsed:
        print(len(i))
        print(i)
    # dbf = DBF(DBF_NAME)

    # with open(CST_NAME, 'r') as csv_file:
    #     reader = csv.reader(csv_file)
    #     for row in reader:
    #         print(row)    
        # writer = csv.writer(file)
        # writer.writerow(dbf.field_names)
        # for record in dbf:
        #     writer.writerow(list(record.values()))