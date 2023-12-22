import sys
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QApplication,QFileDialog,QMessageBox,QVBoxLayout,QFrame,QLabel,QWidget
from PySide6.QtCore import QFile, QIODevice,Qt
from PySide6.QtGui import QImage,QPixmap
import os
from openpyxl import load_workbook
import xlrd
import re
import matplotlib as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import numpy as np
import openpyxl
from openpyxl.chart import BarChart, Reference
from datetime import datetime
from cal030 import calcular030
from cal020 import calcular020
from cal040 import calcular040
from cal050A import calcular050A

path_cal20 = ''
path_cal30 = ''
path_cal40 = ''
path_cal50 = ''

def selectPath(_condition):
    global path_cal20
    global path_cal30
    global path_cal40
    global path_cal50

    options = QFileDialog.Options()

    directory_dialog = QFileDialog()
    directory_dialog.setFileMode(QFileDialog.Directory)

    directory_path = directory_dialog.getExistingDirectory(None, "Select Directory", "", options=options)
    ruta = directory_path
    nombre_archivo = os.path.basename(ruta)
    if _condition == 1:
        window.path_20.setText(nombre_archivo)
        path_cal20 = directory_path
    elif _condition == 2:
        window.path_30.setText(nombre_archivo)
        path_cal30 = directory_path
    elif _condition == 3:
        window.path_40.setText(nombre_archivo)
        path_cal40 = directory_path
    else:
        window.path_50.setText(nombre_archivo)
        path_cal50 = directory_path


def procesarDb():
    option20 = window.radio20.isChecked()
    option30 = window.radio30.isChecked()
    option40 = window.radio40.isChecked()
    option50 = window.radio50.isChecked()
    date = datetime.now()
    
    name_file = ''
    if option20:
        print('procesamos cal 20')
        name_file = f' cal20 {date.day}-{date.month}-{date.year} {date.hour}-{date.minute}-{date.second}'
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(None, "Guardar Excel", f"{name_file}.xlsx", "Excel Files (*.xlsx);;All Files (*)", options=options)
        calcular020(path_cal20,file_name)

    elif option30:
        print('procesamos cal 30')
        name_file = f' cal30 {date.day}-{date.month}-{date.year} {date.hour}-{date.minute}-{date.second}'
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(None, "Guardar Excel", f"{name_file}.xlsx", "Excel Files (*.xlsx);;All Files (*)", options=options)
        calcular030(path_cal30,file_name)

    elif option40:
        print('procesamos cal 40')
        name_file = f' cal40 {date.day}-{date.month}-{date.year} {date.hour}-{date.minute}-{date.second}'
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(None, "Guardar Excel", f"{name_file}.xlsx", "Excel Files (*.xlsx);;All Files (*)", options=options)
        calcular040(path_cal40,file_name)

    else:
        print('procesamos cal 50')
        name_file = f' cal50 {date.day}-{date.month}-{date.year} {date.hour}-{date.minute}-{date.second}'
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(None, "Guardar Excel", f"{name_file}.xlsx", "Excel Files (*.xlsx);;All Files (*)", options=options)
        calcular050A(path_cal50,file_name)
        print(file_name)
    print(name_file)
   
   


if __name__ == "__main__":
    app = QApplication(sys.argv)

    ui_file_name = "mainwindow.ui"
    ui_file = QFile(ui_file_name)
    if not ui_file.open(QIODevice.ReadOnly):
        print(f"Cannot open {ui_file_name}: {ui_file.errorString()}")
        sys.exit(-1)
    loader = QUiLoader()
    window = loader.load(ui_file)
    window.btn20.clicked.connect(lambda: selectPath(1))
    window.btn30.clicked.connect(lambda: selectPath(2))
    window.btn40.clicked.connect(lambda: selectPath(3))
    window.btn50.clicked.connect(lambda: selectPath(4))
    window.process.clicked.connect(lambda: procesarDb())

    ui_file.close()
    if not window:
        print(loader.errorString())
        sys.exit(-1)
    window.show()

    sys.exit(app.exec())