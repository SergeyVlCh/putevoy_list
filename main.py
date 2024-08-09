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

def get_autos():
    conn = sqlite3.connect('putevoy_list.db')
    cur = conn.cursor()
    cur.execute("SELECT marka, gosnomer FROM avto")
    autos = cur.fetchall()
    conn.close()
    return autos

def get_drivers():
    conn = sqlite3.connect('putevoy_list.db')
    cur = conn.cursor()
    cur.execute("SELECT fio FROM voditel")
    drivers = [row[0] for row in cur.fetchall()]
    conn.close()
    return drivers

def load_selections():
    selections = {}
    if os.path.exists('selections.txt'):
        with open('selections.txt', 'r') as file:
            for line in file:
                gosnomer, driver = line.strip().split(',')
                selections[gosnomer] = driver
    return selections

def save_selections(selections):
    with open('selections.txt', 'w') as file:
        for gosnomer, driver in selections.items():
            file.write(f"{gosnomer},{driver}\n")

window = Tk()
window.title("Путевые листы")
window.geometry('640x480')

# Лейблы для диспетчера, механика и медработника
vyb_disp = Label(window, text="Диспетчер")  
vyb_disp.grid(column=0, row=0)

vyb_meh = Label(window, text="Механик")  
vyb_meh.grid(column=0, row=1)

vyb_med = Label(window, text="Медработник")  
vyb_med.grid(column=0, row=2)

dispatchers = get_dispatchers()
mehaniks = get_mehaniks()
mediks = get_mediks()

# Комбо-боксы для диспетчера, механика и медработника
combo_disp = Combobox(window)  
combo_disp['values'] = ['Выберите диспетчера'] + dispatchers  
combo_disp.current(0)  
combo_disp.grid(column=1, row=0)

combo_meh = Combobox(window)  
combo_meh['values'] = ['Выберите механика'] + mehaniks
combo_meh.current(0)  
combo_meh.grid(column=1, row=1)

combo_med = Combobox(window)  
combo_med['values'] = ['Выберите медработника'] + mediks
combo_med.current(0)  
combo_med.grid(column=1, row=2)

# Создаем фрейм с прокруткой для списка авто и водителей
frame = Frame(window)
frame.grid(column=0, row=3, columnspan=2, pady=10, padx=10, sticky='nsew')

# Создаем полосу прокрутки
scrollbar = Scrollbar(frame, orient=VERTICAL)
scrollbar.pack(side=RIGHT, fill=Y)

# Создаем холст для размещения списка
canvas = Canvas(frame, yscrollcommand=scrollbar.set)
canvas.pack(side=LEFT, fill=BOTH, expand=True)

scrollbar.config(command=canvas.yview)

# Создаем внутренний фрейм для размещения элементов списка
inner_frame = Frame(canvas)
canvas.create_window((0, 0), window=inner_frame, anchor='nw')

# Получаем список авто и водителей
autos = get_autos()
drivers = get_drivers()

# Загружаем сохраненные выборы водителей для каждого авто
selections = load_selections()

# Переменные для хранения чекбоксов и комбо-боксов
check_vars = {}
combo_boxes = {}

# Создаем виджеты для каждого авто
for marka, gosnomer in autos:
    row_frame = Frame(inner_frame)
    row_frame.pack(fill='x', pady=5)

    # Создаем чекбокс для каждого авто
    var = BooleanVar()
    check_vars[gosnomer] = var
    check = Checkbutton(row_frame, text=f"{marka} ({gosnomer})", variable=var)
    check.pack(side='left')

    # Создаем выпадающий список для выбора водителя
    combo_driver = Combobox(row_frame, values=drivers)
    combo_driver.set(selections.get(gosnomer, 'Выберите водителя'))  # Устанавливаем сохраненное значение или вариант по умолчанию
    combo_driver.pack(side='right', padx=10)
    combo_boxes[gosnomer] = combo_driver

# Обновляем размеры холста
def on_frame_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

inner_frame.bind("<Configure>", on_frame_configure)

# Функция для сохранения выбранных водителей
def save_selected_drivers():
    for gosnomer, combo in combo_boxes.items():
        selections[gosnomer] = combo.get()
    save_selections(selections)

# Кнопка для сохранения выбранных водителей
save_button = Button(window, text="Сохранить выбор", command=save_selected_drivers)
save_button.grid(column=0, row=4, columnspan=2, pady=10)

window.mainloop()