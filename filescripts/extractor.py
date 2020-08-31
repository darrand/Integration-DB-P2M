import subprocess
import sys
import os
import foxpro_extractor
import datetime
import time
import tkinter as tk
import tkinter.ttk as ttk
from tkcalendar import Calendar, DateEntry

def getDate(date1, date2):
    dates = [date1.get_date(), date2.get_date()]
    return(dates)

def statusCallBack(label):
    if os.path.isfile('./output.xls'):
        label.configure(text="Sukses, file yang dibuat berjudul \"output.xls\"")
    else:
        label.configure(text="Gagal")

def UI():
    master = tk.Tk()
    master.title('Excel Generator GL P2M')
    master.geometry('400x200')
    topframe = tk.Frame(master)
    bottomframe = tk.Frame(master)
    
    l1 = tk.Label(topframe, text='Pilih Tanggal Awal: ', font=("Arial", 15))
    l1.grid(row = 0, column= 0, padx=10, pady=5)
    
    cal1 = DateEntry(topframe)
    cal1.grid(row= 0, column= 1, padx=5, pady=5)

    l2 = tk.Label(topframe, text='Pilih Tanggal Akhir:', font=("Arial", 15))
    l2.grid(row= 1, column= 0, padx=10, pady=5)
    
    cal2 = DateEntry(topframe)
    cal2.grid(row= 1, column= 1, padx=5, pady=5)

    l3 = tk.Label(bottomframe, text='')
    l3.grid(row= 3, column=0)

    b = ttk.Button(bottomframe, text='Buat Excel',command=lambda : [foxpro_extractor.getExcel(getDate(cal1, cal2)),statusCallBack(l3)])
    b.grid(row=2, column=0, padx=10)
    topframe.grid(row= 0, column= 0,padx=50, pady=10)
    bottomframe.grid(row= 1, column= 0,padx=50, pady=10)
    master.mainloop()

if __name__ == "__main__":
    UI()