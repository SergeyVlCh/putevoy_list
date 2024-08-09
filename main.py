from tkinter import *
from tkinter.ttk import Combobox
import sqlite3
from openpyxl import load_workbook
from openpyxl.styles import Font
import os
import zipfile
from datetime import datetime
import pytz

def get_dispatchers():
    conn = sqlite3.connect('putevoy_list.db')
    cur = conn.cursor()
    cur.execute("SELECT fio FROM dispetcher")
    dispatchers = [row[0] for row in cur.fetchall()]
    conn.close()
    return dispatchers

def get_mehaniks():
  conn = sqlite3.connect('putevoy_list.db')
  cur = conn.cursor()
  cur.execute("SELECT fio FROM mehanik")
  mehaniks = [row[0] for row in cur.fetchall()]
  conn.close()
  return mehaniks

def get_mediks():
  conn = sqlite3.connect('putevoy_list.db')
  cur = conn.cursor()
  cur.execute("SELECT fio FROM medsestra")
  mediks = [row[0] for row in cur.fetchall()]
  conn.close()
  return mediks


window = Tk()
window.title("Путевые листы")
window.geometry('640x480')

vyb_disp = Label(window, text="Диспетчер")  
vyb_disp.grid(column=0, row=0)

vyb_meh = Label(window, text="Механик")  
vyb_meh.grid(column=0, row=1)

vyb_med = Label(window, text="Медработник")  
vyb_med.grid(column=0, row=2)

dispatchers = get_dispatchers()
mehaniks = get_mehaniks()
mediks = get_mediks()


combo_disp = Combobox(window)  
combo_disp['values'] = ['Выберите диспетчера'] + dispatchers  
combo_disp.current(0)  # установите вариант по умолчанию  
combo_disp.grid(column=1, row=0)

combo_meh = Combobox(window)  
combo_meh['values'] = ['Выберите механика'] + mehaniks
combo_meh.current(0)  # установите вариант по умолчанию  
combo_meh.grid(column=1, row=1)

combo_med = Combobox(window)  
combo_med['values'] = ['Выберите медработника'] + mediks
combo_med.current(0)  # установите вариант по умолчанию  
combo_med.grid(column=1, row=2)


window.mainloop()