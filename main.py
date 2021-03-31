import sys
import openpyxl
from PyQt5 import uic
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import QWidget, QApplication, QPushButton, QLineEdit, QLabel, QTableWidget, QFileDialog, \
    QMessageBox, QAbstractItemView, QHeaderView, QShortcut, QTabWidget
from openpyxl import Workbook
from PyQt5.QtCore import Qt, QSize

main_ui = uic.loadUiType("cpb.ui")[0]


class SheetTable(QTableWidget):
    def __init__(self):
        # TableWidget Config
        super().__init__(0,3)
        self.setAcceptDrops(True)
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        self.setAlternatingRowColors(True)
        self.setFocusPolicy(Qt.StrongFocus)
        self.verticalHeader().setDefaultSectionSize(220)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.setHorizontalHeaderLabels(['작업 전 사진', '전주 번호', '작업 후 사진'])
        self.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)

        # copyShortcut = QShortcut(QKeySequence.Copy, self)
        # pasteShortcut = QShortcut(QKeySequence.Paste, self)




class MainWindow(QWidget, main_ui):
    name_label: QLabel
    name_lineedit: QLineEdit
    company_label: QLabel
    company_lineedit: QLineEdit
    date_label: QLabel
    date_lineedit: QLineEdit
    manager_label: QLabel
    manager_lineedit: QLineEdit
    load_button: QPushButton
    save_button: QPushButton
    sheet_tabwidget: QTabWidget
    wb: Workbook

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.disabled()

        self.load_button.clicked.connect(self.load_excel)

    def load_excel(self):
        fname, _ = QFileDialog.getOpenFileName(self, '엑셀 파일 선택', './', "Excel File (*.xlsx)")

        if not fname:
            return
        else:
            self.wb = openpyxl.load_workbook(fname)
            for sheetname in self.wb.sheetnames:
                new_tab = SheetTable()
                self.sheet_tabwidget.addTab(new_tab, sheetname)
            sheet = self.wb['공정별사진대장']
            name: str = sheet['B2'].value
            company: str = sheet['B3'].value
            date: str = sheet['Q2'].value.split(':')[-1]
            manager: str = sheet['Q3'].value.split(':')[-1]

            self.name_lineedit.setText(name)
            self.company_lineedit.setText(company)
            self.date_lineedit.setText(date)
            self.manager_lineedit.setText(manager)

            self.enabled()


    def disabled(self):
        self.name_label.setEnabled(False)
        self.name_lineedit.setEnabled(False)
        self.company_label.setEnabled(False)
        self.company_lineedit.setEnabled(False)
        self.date_label.setEnabled(False)
        self.date_lineedit.setEnabled(False)
        self.manager_label.setEnabled(False)
        self.manager_lineedit.setEnabled(False)
        self.save_button.setEnabled(False)
        self.sheet_tabwidget.setEnabled(False)
        # self.sheet_tablewidget.setEnabled(False)

    def enabled(self):
        self.name_label.setEnabled(True)
        self.name_lineedit.setEnabled(True)
        self.company_label.setEnabled(True)
        self.company_lineedit.setEnabled(True)
        self.date_label.setEnabled(True)
        self.date_lineedit.setEnabled(True)
        self.manager_label.setEnabled(True)
        self.manager_lineedit.setEnabled(True)
        self.save_button.setEnabled(True)
        self.sheet_tabwidget.setEnabled(True)
        # self.sheet_tablewidget.setEnabled(True)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    app.exec_()