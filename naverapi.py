import requests
import openpyxl

from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import sys
import icon

class NaverAPI(QMainWindow, QWidget):

    def __init__(self):
        super(NaverAPI, self).__init__()
        self.ui()

    def ui(self):
        self.ch = 'https://openapi.naver.com/v1/search/webkr.json?query='
        # 라벨
        lab_1 = QLabel(self)
        lab_1.setText('Client ID :')
        lab_1.move(39, 20)

        lab_2 = QLabel(self)
        lab_2.setText('Client Secret :')
        lab_2.move(13, 58)

        lab_3 = QLabel(self)
        lab_3.setText('검색할 키워드 :')
        lab_3.move(10, 98)

        # 텍스트
        self.text_1 = QPlainTextEdit(self)
        self.text_1.move(100, 22)
        self.text_1.resize(370, 25)

        self.text_2 = QPlainTextEdit(self)
        self.text_2.move(100, 60)
        self.text_2.resize(370, 25)

        self.text_3 = QPlainTextEdit(self)
        self.text_3.move(100, 100)
        self.text_3.resize(370, 25)

        self.text_4 = QTextEdit(self)
        self.text_4.move(50, 200)
        self.text_4.resize(400, 270)
        self.text_4.setReadOnly(True)

        # ======== 라디오 버튼
        self.radio_1 = QRadioButton(self)
        self.radio_1.setText('웹문서')
        self.radio_1.move(10, 130)
        self.radio_1.setChecked(True)

        self.radio_2 = QRadioButton(self)
        self.radio_2.setText('쇼핑')
        self.radio_2.move(10, 150)

        self.radio_3 = QRadioButton(self)
        self.radio_3.setText('뉴스')
        self.radio_3.move(10, 170)

        # ======== 버튼 체크
        self.radio_1.clicked.connect(self.choice_check)
        self.radio_2.clicked.connect(self.choice_check)
        self.radio_3.clicked.connect(self.choice_check)

        # ======== 버튼
        button_1 = QPushButton(self)
        button_1.move(80, 140)
        button_1.resize(150, 40)
        button_1.setText('크롤링 시작')
        button_1.clicked.connect(self.crawling_start)

        button_2 = QPushButton(self)
        button_2.move(270, 140)
        button_2.resize(150, 40)
        button_2.setText('엑셀 저장')
        button_2.clicked.connect(self.xlsx_save)

        self.setWindowTitle('Naver API : Search')
        self.resize(500, 500)
        self.setWindowIcon(QIcon(':/icon/kwonpu2.png'))
        # self.setGeometry(600, 260, 500, 600)
        self.show()

    def crawling_start(self):
        self.text_4.clear()
        client_id = self.text_1.toPlainText()
        client_secret = self.text_2.toPlainText()
        # client_id = 'ytsOC0NZVEksBGJh1wS0'
        # client_secret = 'GR2vTkcxWX'
        keyword = self.text_3.toPlainText()

        for index in range(10):
            naver_open_api = self.ch + keyword + 'display=100'
            print(naver_open_api)
            header_params = {"X-Naver-Client-Id":client_id, "X-Naver-Client-Secret":client_secret}
            res = requests.get(naver_open_api, headers=header_params)

            if res.status_code == 200:
                data = res.json()
                for item in data['items']:
                    self.text_4.append(str(item['title']))
                    self.text_4.append(str(item['link']))
                    self.text_4.append('\n')
            else:
                QMessageBox.about(self, 'Error', '입력 데이터를 다시 확인해주세요.')
        QMessageBox.about(self, '결과', '크롤링 성공')

    def xlsx_save(self):
        client_id = self.text_1.toPlainText()
        client_secret = self.text_2.toPlainText()
        keyword = self.text_3.toPlainText()

        xlsx_file = openpyxl.Workbook()
        xlsx_sheet = xlsx_file.active
        xlsx_sheet.column_dimensions['A'].width = 100
        xlsx_sheet.column_dimensions['B'].widht = 100

        for index in range(10):
            naver_open_api = self.ch + keyword
            print(naver_open_api)
            header_params = {"X-Naver-Client-Id": client_id, "X-Naver-Client-Secret": client_secret}
            res = requests.get(naver_open_api, headers=header_params)

            if res.status_code == 200:
                data = res.json()
                for item in data['items']:
                    xlsx_sheet.append([item['title'], item['link']])
            else:
                QMessageBox.about(self, 'Error', '입력 데이터를 다시 확인해주세요.')

            xlsx_file.save('크롤링 결과.xlsx')
            xlsx_file.close()
        QMessageBox.about(self, '결과', '엑셀 저장 성공')

    def choice_check(self):
        if self.radio_1.isChecked():
            self.ch = 'https://openapi.naver.com/v1/search/webkr.json?query='
            print(self.ch) # 웹문서
        if self.radio_2.isChecked():
            self.ch = 'https://openapi.naver.com/v1/search/shop.json?query='
            print(self.ch) # 쇼핑
        if self.radio_3.isChecked():
            self.ch = 'https://openapi.naver.com/v1/search/news.json?query='
            print(self.ch) # 뉴스
        return self.ch

def my_exception_hook(exctype, value, traceback):
    sys._excepthook(exctype, value, traceback)


if __name__ == '__main__':
    loopexit = QApplication(sys.argv)
    instance = NaverAPI()
    sys._excepthook = sys.excepthook
    sys.excepthook = my_exception_hook
    loopexit.exec_()
