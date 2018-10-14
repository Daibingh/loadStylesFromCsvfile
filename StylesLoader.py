# -*- coding: utf-8 -*-

from MainWindow import Ui_MainWindow
from demo import newStyle
from PyQt5.QtWidgets import QMainWindow, QTableWidgetItem, QDesktopWidget, QApplication, QMessageBox
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtGui import QCloseEvent
import win32com.client as win32
import pywintypes
import csv
import PyQt5.sip
from PyQt5.QtCore import QThread
import pythoncom


shared_list = []

class Work(QThread):

    done = pyqtSignal()
    error = pyqtSignal()

    def __int__(self):
        super().__init__()

    def loadStyles(self):
        l = shared_list
        app = win32.gencache.EnsureDispatch('Word.Application')
        try:
            doc = app.ActiveDocument
        except pywintypes.com_error as e:
            print('com_error:', e)
            raise e
        else:
            for style in l:
                if style['Bold'] == 'FALSE' or style['Bold'] == 'False' or style['Bold'] == 'false':
                    style['Bold'] = False
                elif style['Bold'] == 'TRUE' or style['Bold'] == 'True' or style['Bold'] == 'true':
                    style['Bold'] = True
                if style['Italic'] == 'FALSE' or style['Italic'] == 'False' or style['Italic'] == 'false':
                    style['Italic'] = False
                elif style['Italic'] == 'TRUE' or style['Italic'] == 'True' or style['Italic'] == 'true':
                    style['Italic'] = True
                style['OutlineLevel'] = int(style['OutlineLevel'])
                style['LeftIndent'] = int(style['LeftIndent'])
                style['RightIndent'] = int(style['RightIndent'])
                style['FirstLineIndent'] = int(style['FirstLineIndent'])
                style['LineUnitBefore'] = int(style['LineUnitBefore'])
                style['LineUnitAfter'] = int(style['LineUnitAfter'])
                style['LineSpacing'] = float(style['LineSpacing'])
                newStyle(app, doc, style)

    def run(self):
        pythoncom.CoInitialize()
        try:
            self.loadStyles()
        except pywintypes.com_error as e:
            self.error_state = str(e)
            self.error.emit()
        else:
            self.done.emit()

class StylesLoader(QMainWindow):
    # name_default = QTableWidgetItem('新样式')
    # chinesefont_default = QTableWidgetItem('宋体')
    # westernfont_default = QTableWidgetItem('Times New Roman')
    # fontsize_default = QTableWidgetItem('小四')
    # bold_default = QTableWidgetItem('False')
    # italic_default = QTableWidgetItem('False')
    # alignment_default = QTableWidgetItem('左')
    # outlinelevel_default = QTableWidgetItem('10')
    default_value = ['新样式', '宋体', 'Times New Roman', '小四', 'False', 'False',
                     '左', '10', '0', '0', '2', '0', '0', '1.25', '']
    keys = ['Name', 'ChineseFont', 'WesternFont', 'FontSize', 'Bold', 'Italic', 'Alignment',
            'OutlineLevel', 'LeftIndent', 'RightIndent', 'FirstLineIndent', 'LineUnitBefore',
            'LineUnitAfter', 'LineSpacing', 'Shortcut']
    num = 0
    file = 'styles.csv'

    def __init__(self):
        super().__init__()
        self.worker = Work()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.resize(920, 600)
        self.center()
        try:
            self.showStyles()
        except FileNotFoundError as e:
            QMessageBox.information(self, '提示', '没有找到styles.csv')

        self.ui.pushButton.clicked.connect(self.addStyle)
        self.ui.pushButton_4.clicked.connect(self.load)
        self.ui.pushButton_2.clicked.connect(self.deleteItem)
        self.ui.pushButton_3.clicked.connect(self.deleteAll)
        self.worker.done.connect(self.ready)
        self.worker.error.connect(self.hold_error)

    def hold_error(self):
        QMessageBox.warning(self, '警告', self.worker.error_state)
        self.ready()

    def load(self):
        global shared_list
        self.ui.pushButton.setEnabled(False)
        self.ui.pushButton_2.setEnabled(False)
        self.ui.pushButton_3.setEnabled(False)
        self.ui.pushButton_4.setEnabled(False)
        self.ui.tableWidget.setEnabled(False)
        shared_list = self.getStylesList()
        self.worker.start()

    def ready(self):
        self.ui.pushButton.setEnabled(True)
        self.ui.pushButton_2.setEnabled(True)
        self.ui.pushButton_3.setEnabled(True)
        self.ui.pushButton_4.setEnabled(True)
        self.ui.tableWidget.setEnabled(True)

    def addStyle(self):
        i = 0
        # print('hello')
        self.ui.tableWidget.insertRow(0)
        for v in self.default_value:
            self.ui.tableWidget.setItem(0, i, QTableWidgetItem(v))
            i += 1
        self.num += 1

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.setGeometry(qr)

    def getStylesList(self):
        l = []
        for i in range(self.num):
            j = 0
            d = dict()
            for j in range(15):
                d[self.keys[j]] = self.ui.tableWidget.item(i, j).text()
            l.append(d)
        # print(len(l))
        # print(l)
        return l

    def showStyles(self):
        with open(self.file, 'r', newline='') as f:
            dictReader = csv.DictReader(f)
            for style in dictReader:
                self.ui.tableWidget.insertRow(self.num)
                for j in range(15):
                    self.ui.tableWidget.setItem(self.num, j, QTableWidgetItem(style[self.keys[j]]))
                self.num += 1

    def saveStyles(self):
        with open(self.file, 'w', newline='') as f:
            dictWriter = csv.DictWriter(f, fieldnames=self.keys)
            dictWriter.writeheader()
            for i in range(self.num):
                d = dict()
                for j in range(15):
                    d[self.keys[j]] = self.ui.tableWidget.item(i, j).text()
                dictWriter.writerow(d)
        print('save successfully!')

    def closeEvent(self, event):
        reply = QMessageBox.question(self, '关闭提示', '保存当前更改？',
                                     QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel, QMessageBox.Yes)
        if reply == QMessageBox.Cancel:
            event.ignore()
        elif reply == QMessageBox.No:
            event.accept()
            # QApplication.quit()
        else:
            self.saveStyles()
            event.accept()
            # QApplication.quit()


    def deleteItem(self):
        self.ui.tableWidget.removeRow(self.ui.tableWidget.currentRow())
        self.num -= 1

    def deleteAll(self):
        self.ui.tableWidget.clearContents()
        self.num = 0