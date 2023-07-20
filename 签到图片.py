from PyQt5.Qt import *
from PyQt5 import uic
import sys
import requests
import os
import shutil
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


class My_Ui(QWidget):
    def __init__(self):
        super().__init__()
        self.ui = uic.loadUi("主窗口.ui", self)
        self.ui.select.clicked.connect(self.selectFileDialog)
        self.ui.start.clicked.connect(self.mkDir)

    def selectFileDialog(self):
        result = QFileDialog.getOpenFileName(self, "选择一个excel文件", "./")
        self.ui.FileName.setText(result[0])

    def downloadMessage(self, flag):  # 下载提示, 0表示下载完成，1表示开始下载
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("提示")
        if flag == 1:
            msg.setText("已开始下载")
        elif flag == 0:
            msg.setText("下载完成")
        elif flag == 2:
            msg.setIcon(QMessageBox.Warning)
            msg.setText("该文件中不包含图片链接，请重新选择")
        msg.exec()

    def deleteExistDir(self, dirName):  # 删除已存在的图片文件夹
        msg = QMessageBox()
        msg.setWindowTitle("提示")
        msg.setIcon(QMessageBox.Question)
        msg.setText("图片文件夹已存在")
        msg.setInformativeText("是否删除？")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        result = msg.exec()
        if result == QMessageBox.Ok:
            shutil.rmtree(dirName)

    def mkDir(self):  # 创建图片文件夹
        save_folder = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') + "\图片"
        try:
            os.makedirs(save_folder, exist_ok=False)
            self.download(save_folder)
        except OSError as e:
            self.deleteExistDir(save_folder)

    def download(self, save_folder):  # 下载函数
        excelPath = self.ui.FileName.text()
        sheet_name = "已签到"
        try:
            workbook = load_workbook(excelPath)
            sheet = workbook[sheet_name]
            link_list = []
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.hyperlink:
                        cell.value = cell.hyperlink.target
                        link_list.append(cell.hyperlink.target)

            if link_list:
                self.downloadMessage(1)
                self.ui.start.setEnabled(False)

            for i, url in enumerate(link_list):
                filename = f"{i}.jpg"
                save_path = os.path.join(save_folder, filename)
                response = requests.get(url, stream=True)
                response.raise_for_status()
                with open(save_path, 'wb') as file:
                    for chunk in response.iter_content(chunk_size=256000):
                        file.write(chunk)

            self.downloadMessage(0)
            self.ui.start.setEnabled(True)
        except InvalidFileException:
            self.downloadMessage(2)
            shutil.rmtree(save_folder)


app = QApplication(sys.argv)
my = My_Ui()
my.show()
sys.exit(app.exec_())
