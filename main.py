import sys
import sqlite3
import requests
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QComboBox, QPushButton, QMessageBox
from openpyxl import load_workbook
from openpyxl.styles import Font
import os
import zipfile
from datetime import datetime
import pytz

files_to_download = []
file_counter = 1

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Формирование и скачивание файла')

        layout = QVBoxLayout()

        self.dispatcher_label = QLabel('Диспетчер')
        self.dispatcher_combo = QComboBox()
        self.load_dispatchers()

        self.mechanic_label = QLabel('Механик')
        self.mechanic_combo = QComboBox()
        self.load_mechanics()

        self.medic_label = QLabel('Медсестра')
        self.medic_combo = QComboBox()
        self.load_medics()

        self.car_label = QLabel('Автомобиль')
        self.car_combo = QComboBox()
        self.load_cars()

        self.driver_label = QLabel('Водитель')
        self.driver_combo = QComboBox()
        self.load_drivers()

        self.generate_button = QPushButton('Сформировать')
        self.generate_button.clicked.connect(self.generate_file)

        self.download_button = QPushButton('Скачать')
        self.download_button.clicked.connect(self.download_file)

        layout.addWidget(self.dispatcher_label)
        layout.addWidget(self.dispatcher_combo)
        layout.addWidget(self.mechanic_label)
        layout.addWidget(self.mechanic_combo)
        layout.addWidget(self.medic_label)
        layout.addWidget(self.medic_combo)
        layout.addWidget(self.car_label)
        layout.addWidget(self.car_combo)
        layout.addWidget(self.driver_label)
        layout.addWidget(self.driver_combo)
        layout.addWidget(self.generate_button)
        layout.addWidget(self.download_button)

        self.setLayout(layout)

    def load_dispatchers(self):
        conn = sqlite3.connect('putevoy_list.db')
        cursor = conn.cursor()
        cursor.execute("SELECT fio FROM dispetcher")
        dispatchers = cursor.fetchall()
        conn.close()

        for dispatcher in dispatchers:
            self.dispatcher_combo.addItem(dispatcher[0])

    def load_mechanics(self):
        conn = sqlite3.connect('putevoy_list.db')
        cursor = conn.cursor()
        cursor.execute("SELECT fio FROM mehanik")
        mechanics = cursor.fetchall()
        conn.close()

        for mechanic in mechanics:
            self.mechanic_combo.addItem(mechanic[0])

    def load_medics(self):
        conn = sqlite3.connect('putevoy_list.db')
        cursor = conn.cursor()
        cursor.execute("SELECT fio FROM medsestra")
        medics = cursor.fetchall()
        conn.close()

        for medic in medics:
            self.medic_combo.addItem(medic[0])

    def load_cars(self):
        conn = sqlite3.connect('putevoy_list.db')
        cursor = conn.cursor()
        cursor.execute("SELECT id, marka, gosnomer FROM avto")
        cars = cursor.fetchall()
        conn.close()

        for car in cars:
            self.car_combo.addItem(f'{car[0]}, {car[1]}, {car[2]}')

    def load_drivers(self):
        conn = sqlite3.connect('putevoy_list.db')
        cursor = conn.cursor()
        cursor.execute("SELECT fio FROM voditel")
        drivers = cursor.fetchall()
        conn.close()

        for driver in drivers:
            self.driver_combo.addItem(driver[0])

        def generate_file(self):
            global file_counter

            selected_car = self.car_combo.currentText()
            driver = self.driver_combo.currentText()
            dispatcher = self.dispatcher_combo.currentText()
            mechanic = self.mechanic_combo.currentText()
            medic = self.medic_combo.currentText()

            car_id, marka, gosnomer = selected_car.split(', ')

            conn = sqlite3.connect('putevoy_list.db')
            cursor = conn.cursor()
            cursor.execute("SELECT nomer_vu, klass, snils FROM voditel WHERE fio = ?", (driver,))
            driver_details = cursor.fetchone()
            conn.close()

            wb = load_workbook('pl.xlsx')
            sheet = wb.active

            bold_black_font = Font(bold=True, color="000000", size=12)

            moscow_tz = pytz.timezone('Europe/Moscow')
            moscow_time = datetime.now(moscow_tz)
            date_str = moscow_time.strftime("%d%m%y")
            day_str = moscow_time.strftime("%d")
            month_str = moscow_time.strftime("%m")
            year_str = moscow_time.strftime("%y")

            cells_to_bold = [
                'BF5', 'BO5', 'CN5', 'DA3', 'S15', 'AA16', 'I17', 'M19', 'AS19', 
                'M21', 'EN42', 'DQ48', 'GT16', 'EN40', 'GA48', 'GT17', 'V46', 
                'BV39', 'X59', 'DP59', 'AP59', 'EH59', 'AU59', 'EM59', 'BS59', 'FL59'
            ]

            for cell in cells_to_bold:
                sheet[cell].font = bold_black_font

            sheet['BF5'] = day_str
            sheet['BO5'] = month_str
            sheet['CN5'] = year_str
            sheet['DA3'] = date_str + str(file_counter)
            sheet['S15'] = marka
            sheet['AA16'] = gosnomer
            sheet['I17'] = driver
            sheet['M19'] = driver_details[0]  # Номер ВУ
            sheet['AS19'] = driver_details[1]  # Класс
            sheet['M21'] = driver_details[2]  # СНИЛС
            sheet['EN42'] = driver
            sheet['DQ48'] = driver
            sheet['GT16'] = mechanic
            sheet['EN40'] = mechanic
            sheet['GA48'] = mechanic
            sheet['GT17'] = dispatcher
            sheet['V46'] = dispatcher
            sheet['BV39'] = medic
            sheet['X59'] = date_str + str(file_counter)
            sheet['DP59'] = date_str + str(file_counter)
            sheet['AP59'] = day_str
            sheet['EH59'] = day_str
            sheet['AU59'] = month_str
            sheet['EM59'] = month_str
            sheet['BS59'] = year_str
            sheet['FL59'] = year_str

            file_name = f'{moscow_time.strftime("%d-%m-%Y")}-{gosnomer}.xlsx'
            wb.save(file_name)
            files_to_download.append(file_name)

            file_counter += 1

            QMessageBox.information(self, "Файл сформирован", f"Файл {file_name} успешно сформирован.")

        def download_file(self):
            with zipfile.ZipFile('files.zip', 'w') as zipf:
                for file in files_to_download:
                    zipf.write(file)

            QMessageBox.information(self, "Файл скачан", "Архив files.zip успешно создан.")

            os.remove('files.zip')
            for file in files_to_download:
                os.remove(file)
            files_to_download.clear()

        if __name__ == '__main__':
            app = QApplication(sys.argv)
            mainWin = MainWindow()
            mainWin.show()
            sys.exit(app.exec_())