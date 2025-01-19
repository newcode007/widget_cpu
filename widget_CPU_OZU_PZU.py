import math
import os
import sys
import time
from tkinter import filedialog

import pandas as pd
import openpyxl
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QAction, QVBoxLayout
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QTimer
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

global file_name_bd
global my_tuple
global data
global DISK

import psutil

my_tuple = ()


class Cpu_mem(QtWidgets.QLabel):

    def __init__(self, parent=None):
        global DISK
        DISK= "C:"
        super(Cpu_mem, self).__init__()
        self.setParent(parent)
        self.setWindowTitle("CPU_MEM_WIDGET")
        self.setObjectName("Form")
        self.resize(False, False)
        self.setMinimumHeight(800)
        self.setMinimumWidth(800)
        self.setMaximumHeight(800)
        self.setMaximumWidth(800)
        self.setFixedWidth(800)
        self.setFixedHeight(800)
        self.setStyleSheet(
            "QWidget {font: bold 17px; border-style: solid; border-width: 0px; border-radius: 8px; border-color: #11EB0D;}"
            "QtWidgets.QPushButton {font: bold 17px; border-style: solid; border-width: 1px; border-radius: 4px; border-color: #11EB0D;}")
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)

        self.cent = QtWidgets.QWidget(self)
        self.exit_button = QtWidgets.QPushButton(self.cent)
        self.exit_button.setObjectName("x")
        self.exit_button.setGeometry(QtCore.QRect(750, 5, 30, 30))
        self.exit_button.setText("x")
        self.exit_button.clicked.connect(QApplication.instance().quit)

        self.hide_b = QtWidgets.QPushButton(self.cent)
        self.hide_b.setObjectName("h")
        self.hide_b.setGeometry(QtCore.QRect(700, 5, 30 ,30))
        self.hide_b.setText("_")
        self.hide_b.clicked.connect(self.hide)

        self.cpu = QtWidgets.QLabel(self.cent)
        self.cpu.setObjectName("cpu")
        self.cpu.setGeometry(QtCore.QRect(10, 10, 95, 30))
        self.cpu.setText("cpu")

        self.memory = QtWidgets.QLabel(self.cent)
        self.memory.setObjectName("memory")
        self.memory.setGeometry(QtCore.QRect(110, 10, 130, 30))
        self.memory.setText(f'ОЗУ: {psutil.virtual_memory().total / (1024**3):.2f} Гб')

        self.swap = QtWidgets.QLabel(self.cent)
        self.swap.setObjectName("swap_memory")
        self.swap.setGeometry(QtCore.QRect(245, 10, 140, 30))
        self.swap.setText(f'ПЗУ: {psutil.disk_usage(DISK).free/(1024*1024*1024):.2f} Гб')

        self.time_lag = QtWidgets.QLineEdit(self.cent)
        self.time_lag.setObjectName("time_lag")
        self.time_lag.setGeometry(QtCore.QRect(380, 10, 150, 30))
        self.time_lag.setText('Интервал, мс')

        self.save_file = QtWidgets.QPushButton(self.cent)
        self.save_file.setObjectName("save_file")
        self.save_file.setGeometry(QtCore.QRect(535, 10, 80, 30))
        self.save_file.setText("Файл БД")
        self.save_file.clicked.connect(self.open_file)


        self.cpu_mem_b_s = QtWidgets.QPushButton(self.cent)
        self.cpu_mem_b_s.setObjectName("cpu_mem_b_s")
        self.cpu_mem_b_s.setGeometry(QtCore.QRect(620, 10, 80, 30))
        self.cpu_mem_b_s.setText("Start")
        self.cpu_mem_b_s.clicked.connect(self.mon)

        self.cpu_mem_b_st = QtWidgets.QPushButton(self.cent)
        self.cpu_mem_b_st.setObjectName("cpu_mem_b_st")
        self.cpu_mem_b_st.setGeometry(QtCore.QRect(620, 10, 80, 30))
        self.cpu_mem_b_st.setText("Stop")
        self.cpu_mem_b_st.clicked.connect(self.save)
        self.cpu_mem_b_st.hide()

        self.fr = QtWidgets.QFrame(self.cent)
        self.fr.setGeometry(0, 50, 790, 660)
        self.m_layout = QVBoxLayout(self.fr)

        self.fig = Figure(figsize=(2, 2), dpi=100)
        self.canvas = FigureCanvas(self.fig)
        self.axes = self.fig.add_subplot(111)
        self.m_layout.addWidget(self.canvas, alignment=Qt.AlignTop)

        self.axes.set_xlabel('Time')
        self.axes.set_ylabel('Загрузка ЦП, %')
        self.axes.set_title("График загрузки ЦП")
        self.canvas.draw

        self.fig1 = Figure(figsize=(2, 2), dpi=100)
        self.canvas1 = FigureCanvas(self.fig1)
        self.axes1 = self.fig1.add_subplot(111)
        self.m_layout.addWidget(self.canvas1, alignment=Qt.AlignTop)

        self.axes1.set_xlabel('Time')
        self.axes1.set_ylabel('Загрузка оп. памяти, %')
        self.axes1.set_title("График загрузки оперативной памяти")
        self.canvas.draw

        self.fig2 = Figure(figsize=(2, 2), dpi=100)
        self.canvas2 = FigureCanvas(self.fig2)
        self.axes2 = self.fig2.add_subplot(111)
        self.m_layout.addWidget(self.canvas2, alignment=Qt.AlignTop)

        self.axes2.set_xlabel('Time')
        self.axes2.set_ylabel('Загрузка долг. памяти, Гб')
        self.axes2.set_title("График загрузки долговременной памяти")
        self.canvas.draw

        self.tray_icon = QSystemTrayIcon(self)
        pixmap = QtGui.QPixmap(":/i.png")

        self.tray_icon.setIcon(QIcon(pixmap))

        #Контекстное меню
        self.tray_menu = QMenu(self)
        self.hide_action = QAction("Свернуть", self)
        self.restore_action = QAction("Развернуть", self)
        self.exit_action = QAction("Выход", self)

        #Подключение слотов для действий контекстного меню
        self.hide_action.triggered.connect(self.hide)
        self.restore_action.triggered.connect(self.showNormal)
        self.exit_action.triggered.connect(QApplication.instance().quit)

        #Добавление действий в контекстное меню
        self.tray_menu.addAction(self.hide_action)
        self.tray_menu.addAction(self.restore_action)
        self.tray_menu.addAction(self.exit_action)

        #Установка контекстного меню для трея
        self.tray_icon.setContextMenu(self.tray_menu)

        #Показ трея и установка обработчика для двойного щелчка
        self.tray_icon.show()
        self.tray_icon.activated.connect(self.tray_icon_clicked)

        self.show()

    def get_cpu_load(self):
        return psutil.cpu_percent(interval=1)

    def get_memory_usage(self):
        mem = psutil.virtual_memory()
        return mem.percent

    def get_swap_memory(self):
        swap = round((psutil.disk_usage(DISK).free/(1024*1024*1024)), 2)
        return swap


    def mausePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.draggable_position = event.globalPos() - self.frameGeometry().topLeft()

    def mauseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            self.move(event.globalPos() - self.draggable_position)
            event.accept()

    def tray_icon_clicked(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.showNormal()

    def add_element(self):
        global my_tuple
        global dat
        self.cpu_load = self.get_cpu_load()
        self.memory_usage = self.get_memory_usage()
        self.swap_memory = self.get_swap_memory()
        data = str(f'{self.cpu_load}; ') + str(f'{self.get_memory_usage()}; ') + str(f'{self.get_swap_memory()}; ') + time.strftime("%Y-%m-%d %H:%M:%S")
        dat = time.strftime("%H:%M:%S")
        my_tuple += (data,) #Добавляем элемент в кортеж
        self.cpu.setText(f'cpu: {self.cpu_load} %') #Добавляем элемент в лэйбл загрузки CPU
        self.memory.setText(f'mem.: {self.memory_usage} %') # --- в оперативную память
        self.swap.setText(f'sw.mem.: {self.swap_memory} Гб') # --- в долговременную память
        self.axes.plot(dat, self.cpu_load, marker='*', color='black')
        self.axes1.plot(dat, self.memory_usage, marker='*', color='black')
        self.axes2.plot(dat, self.swap_memory, marker='*', color='black')
        self.canvas.draw()
        self.canvas1.draw()
        self.canvas2.draw()

    def mon(self):
        global my_tuple
        global data
        global ttt

        self.add_element() #Вызываем функцию добавления элемента
        self.timer = QTimer()
        self.timer.timeout.connect(self.mon)
        try:
            ttt = int(self.time_lag.text())
        except:
            ttt = 1000
        self.timer.start(ttt) #Обновлять график каждую секунду
        self.cpu_mem_b_s.hide()
        self.cpu_mem_b_st.show()

    def open_file(self):
        global file_name_bd
        return os.system('start EXCEL.EXE "' + file_name_bd + '"')

    def save(self):
        self.timer.stop()
        da = my_tuple
        print(da)
        df = pd.DataFrame(da)
        print(df)
        headers = ['CPU_Usage' 'Memory_Usage' 'Swap_memory' 'Time']
        print(df)
        df.columns = headers
        df[['CPU_Usage', 'Memory_Usage', 'Swap_memory', 'Time']] = df['CPU_Usage' 'Memory_Usage' 'Swap_memory' 'Time'].str.split(';', expand=True)

        file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')])

        if file_path:
            #Сохраняем DataFrame в выбранный файл
            df[['CPU_Usage', 'Memory_Usage', 'Swap_memory', 'Time']].to_excel(file_path, index=False)
            print(f'DataFrame сохранен в файл {file_path}')
            global file_name_bd
            file_name_bd= file_path
        else:
            print("Файл не выбран")
        self.cpu_mem_b_s.show()
        self.cpu_mem_b_st.hide()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    cpu_mem = Cpu_mem()
    cpu_mem.show()
    sys.exit(app.exec_())

