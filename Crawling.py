import requests
from bs4 import BeautifulSoup
import openpyxl

from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import sys
import icon


def write_excel_template(filename, sheetname, listdata):
    excel_file = openpyxl.Workbook()
    excel_sheet = excel_file.active
    excel_sheet.column_dimensions['A'].width = 100
    excel_sheet.column_dimensions['B'].width = 20

    if sheetname != '':
        excel_sheet.title = sheetname

    for item in listdata:
        excel_sheet.append(item)
    excel_file.save(filename)
    excel_file.close()


product_lists = list()


class CrawlingGUI(QMainWindow, QWidget):

    def __init__(self):
        super(CrawlingGUI, self).__init__()
        self.ui()

    def ui(self):
        self.ch = 'Select'
        # ======== 라벨
        lab_1 = QLabel(self)
        lab_1.setText('웹사이트 주소')
        lab_1.move(217, 2)

        lab_2 = QLabel(self)
        lab_2.setText('크롤링 CSS')
        lab_2.move(220, 60)

        # ======== 라디오 버튼
        self.radio_1 = QRadioButton(self)
        self.radio_1.setText('Select')
        self.radio_1.move(170, 115)
        self.radio_1.setChecked(True)

        self.radio_2 = QRadioButton(self)
        self.radio_2.setText('Select One')
        self.radio_2.move(270, 115)

        # ======== 버튼 체크
        self.radio_1.clicked.connect(self.css_check)
        self.radio_2.clicked.connect(self.css_check)

        # ======== 텍스트
        self.textvi_1 = QPlainTextEdit(self)
        self.textvi_1.move(50, 30)
        self.textvi_1.resize(400, 25)

        self.textvi_2 = QPlainTextEdit(self)
        self.textvi_2.move(50, 90)
        self.textvi_2.resize(400, 25)

        self.textvi_3 = QPlainTextEdit(self)
        self.textvi_3.move(50, 200)
        self.textvi_3.resize(400, 290)
        self.textvi_3.setReadOnly(True)

        # ======== 버튼
        button_1 = QPushButton(self)
        button_1.move(80, 150)
        button_1.resize(150, 40)
        button_1.setText('크롤링 시작')
        button_1.clicked.connect(self.crawling_start)

        button_2 = QPushButton(self)
        button_2.move(270, 150)
        button_2.resize(150, 40)
        button_2.setText('엑셀 저장')
        button_2.clicked.connect(self.xlsm_create)

        self.setWindowTitle('Crawling Program')
        self.resize(500, 500)
        self.setWindowIcon(QIcon(':/icon/kwonpu2.png'))
        self.show()

    def crawling_start(self):
        self.textvi_3.clear()
        self.les = []
        res = requests.get(self.textvi_1.toPlainText())
        soup = BeautifulSoup(res.content, 'html.parser')

        titles = soup.select(self.textvi_2.toPlainText())
        for title in titles:
            self.textvi_3.appendPlainText(title.get_text())
            self.les.append(title.get_text().split())
            print(title.get_text())

    def css_check(self):
        if self.radio_1.isChecked():
            self.ch = 'select'
            print(self.ch)
        if self.radio_2.isChecked():
            self.ch = 'select_one'
            print(self.ch)
        return self.ch

    def xlsm_create(self):
        write_excel_template('../크롤링 결과.xlsx', '크롤링', self.les)
        QMessageBox.about(self, '엑셀 저장', '엑셀 저장 성공')


def my_exception_hook(exctype, value, traceback):
    sys._excepthook(exctype, value, traceback)


if __name__ == '__main__':
    loopexit = QApplication(sys.argv)
    instance = CrawlingGUI()
    sys._excepthook = sys.excepthook
    sys.excepthook = my_exception_hook
    loopexit.exec_()
