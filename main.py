import sys
from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QMessageBox, QLineEdit
from PyQt5.QtGui import QIcon
import openpyxl
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment, NamedStyle
import time
import pandas as pd
import tkinter 
from tkinter import *
from tkinter import filedialog
from openpyxl import load_workbook
from openpyxl.workbook.protection import WorkbookProtection
from giaodien import Ui_MainWindow
import code_excel 
from code_excel  import main_code, got_data_from
import tkinter
#tao duong dan :  
open_path=''
save_path=''
CSDL_path=''
BM_path=''
DA_name=''
def split_str(name):
    try: 

        extension = name.split('.')
        return extension[1]

    except:
        tbao_chuoi = QMessageBox()
        tbao_chuoi.setWindowTitle('THÔNG BÁO ')
        tbao_chuoi.setText('SAI ĐỊNH DẠNG FILE HOẶC CHƯA CHỌN FILE !')
        tbao_chuoi.setWindowIcon(QIcon('IMAGES\logo.ico'))
        x = tbao_chuoi.exec_()
class MainWindow :
    def __init__(self):
        global DA_name
        self.main_win = QMainWindow()
        self.uic = Ui_MainWindow()
        self.uic.setupUi(self.main_win)    
        
        self.uic.open_button.clicked.connect(self.open_push)
        self.uic.save_button.clicked.connect(self.save_push)
        self.uic.export_button.clicked.connect(self.export_push)
        self.uic.CSDL_button.clicked.connect(self.CSDL_push)
        self.uic.BM_button.clicked.connect(self.BM_push)
                

    def show(self):
        self.main_win.show()
    def open_push(self):
        global open_path
        root = Tk()
        root.withdraw()
        filename_open = filedialog.askopenfilename()        
        open_path=filename_open
        self.uic.open_text.setText(open_path)
        print(open_path)
    def save_push(self):
        global save_path
        filename_save = filedialog.askdirectory()
        save_path=filename_save
        self.uic.save_text.setText(save_path)
        print(save_path)
    def CSDL_push(self):
        global CSDL_path
        filename_CSDL = filedialog.askopenfilename()        
        CSDL_path=filename_CSDL
        self.uic.CSDL_text.setText(CSDL_path)
        print(CSDL_path)
    def BM_push(self):
        global BM_path
        filename_BM = filedialog.askopenfilename()        
        BM_path=filename_BM
        self.uic.BM_text.setText(BM_path)
        print(BM_path)
        

    def export_push(self):
        DA_name = self.uic.DA_text.text()
        print('ten du an la :'+DA_name)

        copen_path = split_str(open_path)       
        cCSDL_path = split_str(CSDL_path)
        cBM_path = split_str(BM_path)

        list_file=['xlsx','xls','xlsm', 'xlsb']
        list_check=[copen_path, cCSDL_path, cBM_path ]
        dkien = True
        for check in list_check:
            if(check in list_file):
                pass
            else:
                dkien = False
        if (dkien == False):
            tbao_type = QMessageBox()
            tbao_type.setWindowTitle('LỖI')
            tbao_type.setText('CHỈ ĐỌC ĐƯỢC FILE EXCEL!')
            tbao_type.setWindowIcon(QIcon('IMAGES\logo.ico'))
            x = tbao_type.exec_()
        
        if(open_path !='' and save_path!='' and CSDL_path!='' and BM_path!=''):    
            #xu ly loi file 
            try:
                got_data_from(open_path, CSDL_path, BM_path)
            except: 
                tbao3 = QMessageBox()
                tbao3.setWindowTitle('LỖI')
                tbao3.setText('KHÔNG THỂ XUẤT FILE !!! HÃY XEM LẠI FILE EXCEL HOẶC DỮ LIỆU TRONG FILE ĐÃ PHÙ HỢP CHƯA')
                tbao3.setWindowIcon(QIcon('IMAGES\logo.ico'))
                x = tbao3.exec_()  
                
            try:
                print(save_path)
                main_code(save_path, DA_name)
                tbao = QMessageBox()
                tbao.setWindowTitle('THÔNG BÁO ')
                tbao.setText('XUẤT THÀNH CÔNG, ĐÃ LƯU VÀO THƯ MỤC : '+save_path)
                tbao.setWindowIcon(QIcon('IMAGES\logo.ico'))
                x = tbao.exec_()
            except:
                tbao4 = QMessageBox()
                tbao4.setWindowTitle('THÔNG BÁO')
                tbao4.setText('CÓ LỖI XẢY RA TRONG KHI XUẤT!')
                tbao4.setWindowIcon(QIcon('IMAGES\logo.ico'))
                x = tbao4.exec_()

        else:
            tbao2 = QMessageBox()
            tbao2.setWindowTitle('LỖI')
            tbao2.setText('BẠN CHƯA CHỌN FILE EXCEL HOẶC THƯ MỤC LƯU')
            tbao2.setWindowIcon(QIcon('IMAGES\logo.ico'))
            x = tbao2.exec_()
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec())



