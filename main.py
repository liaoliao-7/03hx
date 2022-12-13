import datetime

import deal
import os

import riqi
import shutil
import sys

import numpy as np
import xlwings as xw
from PyQt5 import QtCore
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QFileDialog

from file_topdf import PDFConverter
from form import Ui_Form


# 在这个类中写逻辑
class MainWindow(QMainWindow, Ui_Form):
    def __init__(self, parent=None):
        QMainWindow.__init__(self)
        Ui_Form.__init__(self)
        self.setupUi(self)

        self.resize(520, 300)
        self.groupBox_2.setGeometry(10, 20, 291, 240)
        self.groupBox.setVisible(False)  # groupBox隐藏

        # 绑定信号和槽函数
        self.pushButton_2.clicked.connect(self.fram1_visble)
        self.pushButton_3.clicked.connect(self.fram2_visble)
        self.pushButton_4.clicked.connect(self.btn_onefile)  #

    def fram1_visble(self):
        self.groupBox_2.show()
        self.groupBox.setVisible(False)

    def fram2_visble(self):
        self.groupBox_2.setVisible(False)
        self.groupBox.setVisible(True)

    def btn_onefile(self):
        if self.lineEdit_4.text() != "":
            date = riqi.new_time(self.dateTimeEdit.dateTime().toString("yyyy-MM-dd hh:mm:ss"),
                                 int(self.lineEdit_4.text()),
                                 2)
        else:
            QMessageBox.information(self, '消息', '请输入间隔时间（整数）！',
                                    QMessageBox.Yes | QMessageBox.No)
            return

        self.textBrowser.append("1.数据处理开始....")
        app.processEvents()

        os.system("taskkill /F /IM wps.exe /t")
        original_file = self.select_file()
        filename = original_file[1]

        new_file = f'.//new//{filename}'
        try:
            os.makedirs('./new')
        except IOError as e:
            print(e)
        shutil.copy(original_file[0][0], new_file)

        axl = xw.App(visible=False, add_book=False)
        axl.display_alerts = False  # 关闭一些提示信息，加快运行速度。
        axl.screen_updating = False  # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度。
        wb_1 = axl.books.open(new_file)
        sht_1 = wb_1.sheets[0]
        sht_1.range('a1:f100').api.NumberFormat = "@"
        sht_1.range("a1").column_width = 12.24

        # 写入表头
        sht_1.range('b1').value = filename.split(".")[0]
        sht_1.range('b2').value = date[0]
        sht_1.range('b3').value = date[1]

        # 删除多余的列
        sht_1.range('a5:a7').api.EntireRow.Delete()
        sht_1.range('i1:l1').api.EntireColumn.Delete()
        sht_1.range('j1:cc1').api.EntireColumn.Delete()

        # 正常处理数据
        li_a = []
        li_b = []
        for i in sht_1.range('G12:G71'):
            if i.value == "FAM":
                li_a.append(float(sht_1.range(f'i{i.row}').value))
            if i.value == "HEX":
                li_b.append(float(sht_1.range(f'i{i.row}').value))
            if i.value == 'ROX':
                sht_1.range(f'i{i.row}').value = 'Noct'
        li_a1 = deal.deal(li_a)
        self.textBrowser.append(f"1.  FAM 的 cv值为： {li_a1[1]}")
        app.processEvents()
        li_b1 = deal.deal(li_b)
        self.textBrowser.append(f"2.  HEX 的 cv值为： {li_b1[1]}")
        app.processEvents()
        biaoji = 0
        biaoji2 = 0
        for i in sht_1.range('G12:G7'):
            if i.value == "FAM":
                sht_1.range(f'i{i.row}').value = li_a1[0][biaoji]
                biaoji += 1
        for i in sht_1.range('G12:G71'):
            if i.value == "HEX":
                sht_1.range(f'i{i.row}').value = li_b1[0][biaoji2]
                biaoji2 += 1

        # NC值写入
        Noct = ['Noct', 'Noct', 'Noct']
        sht_1.range("i72").options(transpose=True).value = Noct

        # PC值写入
        number = np.random.randint(25, 27, size=3)  # 整数部分
        fu_dian = np.around(np.random.random(3), 2)  # 小数部分
        pc_number = number + fu_dian
        sht_1.range("i75").options(transpose=True).value = pc_number  # PC值写入

        # 设置列宽
        sht_1.range("b1").column_width = 9.62
        sht_1.range("c1").column_width = 8.36
        sht_1.range("D1:I1").column_width = 8.36

        """向汇总统计表中写入数值"""
        wb_2 = axl.books.open("./1.提取仪质检结果汇总表1.xlsx")
        sht_21 = wb_2.sheets[3]
        row = sht_21.used_range.last_cell.row
        biaozhi = 'c'
        for i in range(row, 0, -1):
            if sht_21.range(f'D{i}').value == filename.split("-")[2]:
                QMessageBox.information(self, '消息', '仪器编号在汇总统计中存在，将不进行写入',
                                        QMessageBox.Yes | QMessageBox.No)
                biaozhi = 'a'
                break
        if biaozhi == 'c':
            rows = row + 1
            now = datetime.datetime.now().strftime("%Y.%m.%d")
            sht_21.range(f'B{rows}').value = now
            if 'NC' in filename:
                sht_21.range(f'C{rows}').value = 'BG-Abot-96'
            else:
                sht_21.range(f'C{rows}').value = 'BG-Nege-96'
            sht_21.range(f'D{rows}').value = filename.split("-")[2]
            sht_21.range(f'E{rows}').value = '008-96D'
            sht_21.range(f'H{rows}').value = "%.2f%%" % (li_a1[1] * 100)
            sht_21.range(f'I{rows}').value = "%.2f%%" % (li_b1[1] * 100)
            sht_21.range(f'J{rows}').value = '均检出'
            sht_21.range(f'k{rows}').value = '均检出'
            sht_21.range(f'Q{rows}').value = '合格'

        # 保存，退出
        wb_1.save()
        wb_1.close()
        wb_2.save()
        wb_2.close()

        axl.quit()
        os.system("taskkill /F /IM wps.exe /t")

        self.textBrowser.append("2.正在将文件转化为PDF，请稍后....")
        app.processEvents()

        folder = 'new'
        pathname = os.path.join(os.path.abspath('.'), folder)

        # # 也支持单个文件的转换
        # pathname = r"E:\python\03hx\new\20221130-008-BG96NC20222498-906.xlsx"
        pdfConverter = PDFConverter(pathname)
        pdfConverter.run_conver()

        self.textBrowser.append("3.文件转换PDF完成....")
        self.textBrowser.append("*" * 20)
        app.processEvents()

        xinhao = QMessageBox.information(self, '消息', '处理完成，是否打开文件夹？？',
                                         QMessageBox.Yes | QMessageBox.No)

        if xinhao < 35000:
            path = os.getcwd() + '\\new'  # 获取当前文件路径
            os.system(" start explorer %s" % path)

    def select_file(self):
        tup_file = QFileDialog.getOpenFileNames(
            self, '选择文件', filter='Excel Files (*.xls , *.xlsx)')
        files = tup_file[0]
        onefilename = ''
        try:
            onefilename = files[0].split("/")[-1]
        except:
            pass
        print(onefilename)
        return files, onefilename

    # def btn_clicked(self):
    #     # files = self.select_file()
    #
    # print(files)
    #
    #     rel_time = self.dateTimeEdit.dateTime().toString("yyyy-MM-dd hh:mm:ss")
    #     if self.lineEdit.text() == '' or self.lineEdit_2.text() == '':
    #         QMessageBox.information(self, '消息', '请正确设置间隔时间或数据个数',
    #                                 QMessageBox.Yes | QMessageBox.No)
    #         return
    #     data_date = riqi.new_time(rel_time, int(self.lineEdit.text()), int(self.lineEdit_2.text()))
    #     print(data_date)
    #     notin = shuju.some_su()
    #     print(notin)
    #     QMessageBox.information(self, '消息', '数据已生成完成！',
    #                             QMessageBox.Yes | QMessageBox.No)


"""
日期选择器代码：
self.dateTimeEdit = QtWidgets.QDateTimeEdit(QDateTime.currentDateTime(), self)
self.dateTimeEdit.setDisplayFormat("yyyy-MM-dd HH:mm:ss")


self.dateTimeEdit_2 = QtWidgets.QDateTimeEdit(QDateTime.currentDateTime(), self)
self.dateTimeEdit_2.setDisplayFormat("yyyy-MM-dd HH:mm:ss")
"""

if __name__ == '__main__':
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自适应屏幕分辨率
    app = QApplication(sys.argv)
    # aot = pty.dot('./style.qss')
    window = MainWindow()
    # window.setStyleSheet(aot)
    window.show()
    sys.exit(app.exec_())
