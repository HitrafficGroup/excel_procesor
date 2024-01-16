import sys
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QApplication,QFileDialog
from PySide6.QtCore import QFile, QIODevice
import os
from datetime import datetime
from cal020 import calcular020
from cal030 import calcular030
from cal040 import calcular040
from cal050 import calcular50
from calcular_porcentaje import calcularPorcent
path_cal20 = ''
path_cal30 = ''
path_cal40 = ''
path_cal50 = ''
path_consolided = ''

data_cal20 = []
data_cal30 = []
data_cal40 = []
data_cal50 = []


'''
    INTERFAZ GRAFICA
'''
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
    global data_cal20
    global data_cal30
    global data_cal40
    global data_cal50
    option20 = window.radio20.isChecked()
    option30 = window.radio30.isChecked()
    option40 = window.radio40.isChecked()
    date = datetime.now()
    name_file = ''
    if option20:
        name_file = f' cal20 {date.day}-{date.month}-{date.year} {date.hour}-{date.minute}-{date.second}'
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(None, "Guardar Excel", f"{name_file}.xlsx", "Excel Files (*.xlsx);;All Files (*)", options=options)
        data_cal20 = calcular020(path_cal20,file_name)

    elif option30:
        name_file = f' cal30 {date.day}-{date.month}-{date.year} {date.hour}-{date.minute}-{date.second}'
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(None, "Guardar Excel", f"{name_file}.xlsx", "Excel Files (*.xlsx);;All Files (*)", options=options)
        data_cal30 = calcular030(path_cal30,file_name)

    elif option40:
        name_file = f' cal40 {date.day}-{date.month}-{date.year} {date.hour}-{date.minute}-{date.second}'
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(None, "Guardar Excel", f"{name_file}.xlsx", "Excel Files (*.xlsx);;All Files (*)", options=options)
        data_cal40 = calcular040(path_cal40,file_name)

    else:
        name_file = f' cal50 {date.day}-{date.month}-{date.year} {date.hour}-{date.minute}-{date.second}'
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(None, "Guardar Excel", f"{name_file}.xlsx", "Excel Files (*.xlsx);;All Files (*)", options=options)
        data_cal50 = calcular50(path_cal50,file_name)
  
def seleccionarExcel():
    global path_consolided
    fname = QFileDialog.getOpenFileName()
    ruta = fname[0]
    nombre_archivo = os.path.basename(ruta)
    if fname[0][-5:] == ".xlsx":
        path_consolided = fname[0]
        window.txtPorcent.setText(nombre_archivo)

def generarPorcentajedeMuestra():
    date = datetime.now()
    porcent = window.porcentaje.text()
    op2 = window.op2.isChecked()
    option20 = window.radio20.isChecked()
    option30 = window.radio30.isChecked()
    option40 = window.radio40.isChecked()
    name_file = f'porcentaje {porcent}% {date.day}-{date.month}-{date.year} {date.hour}-{date.minute}-{date.second}'
    options = QFileDialog.Options()
    file_name, _ = QFileDialog.getSaveFileName(None, "Guardar Excel", f"{name_file}.xlsx", "Excel Files (*.xlsx);;All Files (*)", options=options)
    global path_consolided
    aux_porcentaje = 0.8
    if option20:
        data_analisys = data_cal20
    elif option30:
        data_analisys = data_cal30
    elif option40:
        data_analisys = data_cal40
    else:
        data_analisys = data_cal50

    if op2:
        if window.porcentaje.text() != '':
            aux_porcentaje = window.porcentaje.text()
            try:
                aux_porcentaje = round(int(aux_porcentaje)/100,2)

            except:
                aux_porcentaje = 0.8
        else:
            aux_porcentaje = 0.8
        calcularPorcent(data_analisys,aux_porcentaje,file_name,1)
    else:
        calcularPorcent(data_analisys,aux_porcentaje,file_name,2)
    
   

    
def detectarOpcion(op):
    global data_cal20
    global data_cal30
    global data_cal40
    global data_cal50
    if op == 1:
        if len(data_cal20)==0:
            window.btnporcent.setEnabled(False)
        else:
            window.btnporcent.setEnabled(True)
    if op == 2:
        if len(data_cal30)==0:
            window.btnporcent.setEnabled(False)
        else:
            window.btnporcent.setEnabled(True)
    if op == 3:
        if len(data_cal40)==0:
            window.btnporcent.setEnabled(False)
        else:
            window.btnporcent.setEnabled(True)
    else:
        if len(data_cal50)==0:
            window.btnporcent.setEnabled(False)
        else:
            window.btnporcent.setEnabled(True)


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
   
    window.btnporcent.clicked.connect(lambda: generarPorcentajedeMuestra())
    window.radio20.clicked.connect(lambda: detectarOpcion(1))
    window.radio30.clicked.connect(lambda: detectarOpcion(2))
    window.radio40.clicked.connect(lambda: detectarOpcion(3))
    window.radio50.clicked.connect(lambda: detectarOpcion(4))

    ui_file.close()
    if not window:
        print(loader.errorString())
        sys.exit(-1)
    window.show()

    sys.exit(app.exec())