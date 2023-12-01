import sys, datetime
from preprocess import CAPDExcel
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QDateEdit, QVBoxLayout, QHBoxLayout, QSizePolicy, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt, QDate


class DialogApp(QWidget):
    def __init__(self) -> None:
        super().__init__()

        self.setWindowTitle('腹膜透析紀錄表生成器')
        self.resize(250, 180)

        today = QDate.currentDate()

        self.startDateLabel = QLabel('起始日期: ')
        self.startDateLabel.setVisible(False)

        self.startDateEdit = QDateEdit(today)
        self.startDateEdit.setCalendarPopup(True)
        self.startDateEdit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.startDateEdit.setVisible(False)

        self.endDateLabel = QLabel('結束日期: ')
        self.endDateLabel.setVisible(False)

        self.endDateEdit = QDateEdit(today)
        self.endDateEdit.setCalendarPopup(True)
        self.endDateEdit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.endDateEdit.setVisible(False)

        self.uploadBtn = QPushButton('上傳透析紀錄')
        self.uploadBtn.clicked.connect(self.toggleWidgets)
        self.uploadfile = ''

        content_layout = QVBoxLayout()
        content_layout.addWidget(self.uploadBtn)
        content_layout.addWidget(self.startDateLabel)
        content_layout.addWidget(self.startDateEdit)
        content_layout.addSpacing(10)
        content_layout.addWidget(self.endDateLabel)
        content_layout.addWidget(self.endDateEdit)
        content_layout.addSpacing(10)



        self.prevBtn = QPushButton('< 上一頁')
        self.prevBtn.clicked.connect(self.backtoUploadPage)
        self.prevBtn.setVisible(False)

        self.enterBtn = QPushButton('產生透析日誌')
        self.enterBtn.clicked.connect(self.generateCAPDrecord)
        self.enterBtn.setVisible(False)
        
        footer_layout = QHBoxLayout()
        footer_layout.addWidget(self.prevBtn)
        footer_layout.addWidget(self.enterBtn)

        self.warnBox = QMessageBox()

        layout = QVBoxLayout()
        layout.addLayout(content_layout)
        layout.addLayout(footer_layout)
        self.setLayout(layout)


    def toggleWidgets(self):
        if(self.uploadBtn.isHidden()):
            self.uploadBtn.show()
            self.startDateLabel.hide()
            self.startDateEdit.hide()
            self.endDateLabel.hide()
            self.endDateEdit.hide()
            self.prevBtn.hide()
            self.enterBtn.hide()            
        else:
            uploadfile, _ = QFileDialog.getOpenFileName(self, '上傳檔案', '/', "Excel Files (*.xlsx)")
            if uploadfile:
                self.uploadfile = uploadfile
                self.uploadBtn.hide()
                self.startDateLabel.show()
                self.startDateEdit.show()
                self.endDateLabel.show()
                self.endDateEdit.show()
                self.prevBtn.show()
                self.enterBtn.show()


    def backtoUploadPage(self):
        self.uploadBtn.hide()
        self.toggleWidgets()

    def generateCAPDrecord(self):
        startDate = self.startDateEdit.date().toPyDate()
        endDate = self.endDateEdit.date().toPyDate() + datetime.timedelta(days=1)

        startDateStr = self.startDateEdit.date().toString('yyyyMMdd')
        endDateStr = self.endDateEdit.date().toString('yyyyMMdd')

        downloadFile, _ = QFileDialog.getSaveFileName(self, '下載檔案', f'/透析日誌{startDateStr}~{endDateStr}.xlsx', "Excel Files (*.xlsx)")
        if downloadFile:
            try:
                CAPDExcel(self.uploadfile, downloadFile, startDate, endDate).getExcel()
            except:
                self.warnBox.warning(self, '警告', '<center>上傳資料格式有誤，<br>或是沒有該段時間區間的資料。</center>')
                



if __name__ == '__main__':
    app = QApplication(sys.argv)

    window = DialogApp()
    window.show()

    app.exec()