# -*- coding: utf-8 -*-
import configparser
import datetime
import email
import os
import poplib
import re
import shutil
import sys
import time
from email.header import decode_header
from email.parser import Parser
import pandas as pd
import xlwings as xw
from threading import Thread

from PyQt5 import QtCore
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QFileDialog
from zlsjzdh import Ui_mainWindow


# 在这个类中写逻辑
class MainWindow(QMainWindow, Ui_mainWindow):
    def __init__(self, parent=None):
        QMainWindow.__init__(self)
        Ui_mainWindow.__init__(self)
        self.setupUi(self)

        # 绑定信号和槽函数
        self.btn_selectFile_0.clicked.connect(self.btn_selectFile0_clicked)  # 选择合并文件
        self.btn_mergeData.clicked.connect(self.btn_mergeData_Clicked)  # 合并

        self.btn_selectSDMFile.clicked.connect(self.btn_selectSDMFile_clicked)  # 选择SDM文件
        self.btn_generateSDMData.clicked.connect(self.btn_generateSDMData_Clicked)  # sdm数据生成

        self.btn_SelectITMFile.clicked.connect(self.btn_selectITMFile_clicked)  # 选择ITM文件
        self.btn_generateITMData.clicked.connect(self.btn_generateITMData_Clicked)

        self.btn_autoSDM.clicked.connect(self.autoSDM_clicked)  # SDM数据自动化
        self.btn_autoITM.clicked.connect(self.autoITM_clicked)  # ITM数据自动化

        # 读取配置文件
        self.__config = configparser.ConfigParser()
        self.__config.read("Address.ini", encoding="utf-8")
        self.__kaixiang_ad = self.__config.get('adress_jichu', 'kaixiang_ad')
        self.__fenxi_ad = self.__config.get('adress_jichu', 'fenxi_ad')

        # 创建改变状态标志（需在文档中解决）
        self.flag = True
        self.flag_1 = True
    def fun(self):
        self.flag = False
        for i in range(1000000000):
            print("hl")

    def btn_selectFile0_clicked(self):
        try:
            tup_file = QFileDialog.getOpenFileNames(
                self, '选择文件', filter='Excel Files (*.xls , *.xlsx)')
            fpath_0 = tup_file[0]
            self.lineEdit_0.setText(fpath_0[0])
            self.lineEdit_1.setText(fpath_0[1])
            self.lineEdit_2.setText(fpath_0[2])
            self.lineEdit_3.setText(fpath_0[3])
        except:
            pass

    def btn_mergeData_Clicked(self):
        text_0 = self.lineEdit_0.text()
        text_1 = self.lineEdit_1.text()
        text_2 = self.lineEdit_2.text()
        text_3 = self.lineEdit_3.text()
        if len(text_0) == 0 or len(text_1) == 0:
            QMessageBox.information(self, '消息', '请先选择文件！',
                                    QMessageBox.Yes | QMessageBox.No)
            return
        self.__btn_mergeData_Clicked(text_0, text_1, text_2, text_3)

    def __btn_mergeData_Clicked(self, fpath_0, fpath_1, fpath_2, fpath_3):
        global wb_3, wb_4, u_1, tt
        X = []
        Y = []
        for i in range(26):
            X.append(chr(97 + i))
            Y.append(i + 1)
        dit = {x: y for x, y in zip(X, Y)}
        print(dit)

        dot = ['a', 'b']
        dit_2 = {}
        d_t = 27
        for i in dot:
            for y in dit.keys():
                dit_2[i + y] = d_t
                d_t += 1
        dit.update(dit_2)

        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False  # 关闭一些提示信息，加快运行速度。
        app.screen_updating = False  # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度。
        wb_1 = app.books.open(fpath_0)
        wb_2 = app.books.open(fpath_1)
        sht_1 = wb_1.sheets[1]
        sht_2 = wb_2.sheets[1]
        a_1 = sht_1.range('a1:bz90')
        for cell in a_1:
            if cell.value == '田胜春':
                tt = cell.address
                break
        dr = re.search(r'\w+', tt).group()
        dt = dit[dr.lower()]
        di = dt + 1
        for key, value in dit.items():
            if value == di:
                u_1 = key
        a_2 = sht_1.range(f'{u_1}3:{u_1}90').value
        for i in range(len(a_2)):
            if a_2[i] is None:
                sht_1.range(i + 3, di).value = sht_2.range(i + 3, di).value
        wb_1.save()
        wb_1.close()
        try:
            wb_1 = app.books.open(fpath_0)
            sht_1 = wb_1.sheets[1]
            a_2 = sht_1.range(f'{u_1}3:{u_1}90').value
            wb_3 = app.books.open(fpath_2)
            sht_3 = wb_3.sheets[1]
            for i in range(len(a_2)):
                if a_2[i] is None:
                    sht_1.range(i + 3, di).value = sht_3.range(i + 3, di).value
            wb_1.save()
            wb_1.close()
        except:
            pass
        try:
            wb_1 = app.books.open(fpath_0)
            sht_1 = wb_1.sheets[1]
            a_2 = sht_1.range(f'{u_1}3:{u_1}90').value
            wb_4 = app.books.open(fpath_3)
            sht_4 = wb_4.sheets[1]
            for i in range(len(a_2)):
                if a_2[i] is None:
                    sht_1.range(i + 3, di).value = sht_4.range(i + 3, di).value
            wb_1.save()
            wb_1.close()
        except:
            pass

        """关闭后台打开的wb_2,wb_3,wb_4(如果其被打开)"""
        try:
            wb_2.close()
            wb_3.close()
            wb_4.close()
        except:
            pass
        time.sleep(1)
        """打开合并完成的目标文件"""
        app = xw.App(visible=True, add_book=False)
        app.books.open(fpath_0)

    def btn_generateSDMData_Clicked(self):
        text = self.lineEdit_5.text()
        if len(text) == 0:
            QMessageBox.information(self, '消息', '请先选择文件！',
                                    QMessageBox.Yes | QMessageBox.No)
            return  # 此处加返回语句，不加将继续向下运行
        try:
            self.__generateData(text)
        except:
            QMessageBox.information(self, '消息', '选择正确的文件！',
                                    QMessageBox.Yes | QMessageBox.No)

    def btn_selectSDMFile_clicked(self):
        tup_file = QFileDialog.getOpenFileName(
            self, '选择文件', filter='Excel Files (*.xls , *.xlsx)')
        fpath = tup_file[0]
        self.lineEdit_5.setText(fpath)
        """将文件地址写入配置文件"""
        # self.__config.set( 'adress_jichu', 'kaixiang_ad', fpath )
        # with open( 'Address.ini', 'w+', encoding='utf-8' ) as f:
        #     self.__config.write( f )
        # self.__kaixiang_ad = self.__config.get( 'adress_jichu', 'kaixiang_ad' )

    def __generateData(self, fpath):
        # 在默认的条件下，只读取第一个表格  header=0 指定第一行作为列标签
        data = pd.read_excel(fpath, header=0)

        condition_1 = ~data.loc[:, '设备类型'].astype(str).str.startswith('A2000')
        condition_2 = ~data.loc[:, '父设备'].astype(str).str.startswith('802100')
        condition_3 = ~data.loc[:, '开箱异常'].isna()

        install_data = data.loc[condition_1 & condition_2, :]
        open_data = data.loc[condition_1 & condition_2, :]
        exception_data = data.loc[condition_1 & condition_2 & condition_3, :]

        # 计算装机数
        int_install_num = len(install_data)
        # 计算开箱数
        int_open_num = 0
        ls_temp = list(open_data['设备类型'])
        for s in ls_temp:
            int_open_num += int(s[-1]) + 1
        # 计算异常数
        int_exception_num = len(exception_data)
        print(f'装机数为{int_install_num}')
        print(f'开箱数为{int_open_num}')
        print(f'异常数为{int_exception_num}')
        print('--------------------')
        li_object = []

        for i in range(len(exception_data)):
            dic_obj = {'id': str(exception_data.iloc[i, 1]), 'exception': str(exception_data.iloc[i, 6]),
                       'reason': str(exception_data.iloc[i, 7])}
            li_object.append(dic_obj)
        print(li_object)

        book = xw.Book(self.__kaixiang_ad)
        open_sht = book.sheets['2021开箱统计']
        question_sht = book.sheets['问题统计']

        # 在开箱统计中写入 异常数、装机数、开箱数
        temp = 'BCDEFGHIJKLM'
        for char in temp:
            cell = str(char) + str(3)
            ans = open_sht.range(cell).value is None
            if ans == True:
                open_sht.range(char + str(3)).value = int_exception_num
                open_sht.range(char + str(5)).value = int_install_num
                open_sht.range(char + str(6)).value = int_open_num
                break

        # 在问题统计sheet中写入序列号、异常、异常原因   应该从sheet['问题统计']的最后一行开始写入
        int_new_line_row = question_sht.used_range.last_cell.row + 1
        for i in range(len(li_object)):
            # 写入序列号
            question_sht.range(
                str('C') + str(int_new_line_row)).value = li_object[i]['id']
            question_sht.range(
                str('D') +
                str(int_new_line_row)).value = li_object[i]['exception']
            question_sht.range(
                str('E') +
                str(int_new_line_row)).value = li_object[i]['reason']
            int_new_line_row += 1

    def btn_selectITMFile_clicked(self):
        tup_file = QFileDialog.getOpenFileName(
            self, '选择文件', filter='Excel Files (*.xls , *.xlsx)')
        fpath_ITM = tup_file[0]
        self.lineEdit_ITM.setText(fpath_ITM)

    def btn_generateITMData_Clicked(self):
        text = self.lineEdit_ITM.text()
        if len(text) == 0:
            QMessageBox.information(self, '消息', '请先选择文件！',
                                    QMessageBox.Yes | QMessageBox.No)
            return
        try:
            self.__itm_generateData(text, self.comboBox.currentText())
        except:
            QMessageBox.information(self, '消息', '请先选择正确的文件！',
                                    QMessageBox.Yes | QMessageBox.No)

    def __itm_generateData(self, path, month):
        """备份文件"""
        # shutil.copy('I:/python project/yuebao project4 - 副本.xls', 'F:/sjzdh./bak') #完完成后替换成 path_x.get()

        global bb, a, c, al, bl, cl
        """关闭excel进程"""
        # os.system('taskkill /im EXCEL.EXE /F')

        """打开原始表格，复制2个处理表格"""
        app = xw.App(visible=False, add_book=False)  # visible值为False为后台运行
        app.display_alerts = False  # 关闭一些提示信息，加快运行速度。
        app.screen_updating = True  # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度。
        wb_1 = app.books.open(path)  # 完成后替换成 path_x.get()
        wb_1.sheets[0].api.Copy(Before=wb_1.sheets[-1].api)
        sht_1 = wb_1.sheets[-2]
        sht_1.name = '01'
        wb_1.sheets[-1].api.Copy(Before=wb_1.sheets[-1].api)
        sht_2 = wb_1.sheets[-2]
        sht_2.name = '02'
        wb_1.sheets.api.Add(Before=wb_1.sheets[-1].api)
        sht_3 = wb_1.sheets[-2]
        sht_3.name = '03'

        """计算在用仪器台数"""
        rng = sht_1.range('a1:aj3')
        for cell in rng:
            if cell.value == '启用状态':
                a = cell.address
                break
        b = a[1:2]
        row_1 = sht_1.range(f'{b}4:{b}200')
        for cell in row_1:
            if cell.value is None:
                c = cell.address
                break
        d = c[-2:]
        sht_1.range(f'{c}').formula = f'=COUNTif({b}1:{b}{int(d) - 1},"整体启用")'  # 在用条数公式

        """计算再用模块数，不含生化/包含生化"""
        for cell in rng:
            if cell.value == '线上 免疫数':
                al = cell.address
            elif cell.value == '线上 生化数':
                bl = cell.address
            elif cell.value == '线体模块数':
                cl = cell.address
        b_1 = al[1:2]
        c_1 = str(b_1) + str(d)
        c_2 = str(b_1) + str(int(d) + 1)
        c_3 = str(b_1) + str(int(d) + 2)
        b_11 = bl[1:2]
        c_11 = str(b_11) + str(d)
        d_10 = cl[1:2]
        d_11 = str(d_10) + str(d)
        sht_1.range(c_1).formula = f'= sumifs(O4:O{int(d) - 1},{b}4:{b}{int(d) - 1},"整体启用")'  # 在线免疫数
        sht_1.range(c_11).formula = f'= sumifs(p4:p{int(d) - 1},{b}4:{b}{int(d) - 1},"整体启用")'  # 在线生化数
        sht_1.range(d_11).formula = f'= sumifs(N4:N{int(d) - 1},{b}4:{b}{int(d) - 1},"整体启用")'  # 在线模块数(流水线模块)
        a_1 = sht_1.range(f'{c}').value  # 在用仪器条数值
        d_1 = sht_1.range(c_1).value  # 在线免疫数
        f_1 = sht_1.range(d_11).value  # 在线模块数（流水线模块）
        sht_1.range(c_2).formula = f'= {d_1}*2+4*{a_1}+{f_1}'  # 不带生化b1模块数
        sht_1.range(c_3).formula = f'= {sht_1.range(c_2).value}+{d_1}+{sht_1.range(c_11).value}'  # 带生化b1模块数(含免疫、生化)
        dott = sht_1.range(c_3).value

        """打印整线模块数（含免疫+生化）"""
        print('整线模块数（含免疫+生化）：' + str(dott))

        """计算在用测试量"""
        rge = sht_2.range('cb2:cy200')
        for cell in rge:
            if cell.value == str(month):
                bb = cell.address
                break
        cc = bb[1:3]
        ddt = sht_2.range(str(cc) + '2:' + str(cc) + '200')
        at = []
        for cell in ddt:
            if isinstance(cell.value, float):
                at.append(cell.value)
        dt = sum(at)

        """重新生成数据表格"""
        a_5 = ['在用仪器数', '在用模块数（不含生化）', '样本数', '在用模块数（含生化）']
        b_5 = [a_1, sht_1.range(c_2).value, dt, dott]
        sht_3.range('A1').options(transpose=True).value = a_5
        sht_3.range('B1').options(transpose=True).value = b_5
        sht_3.range('A:A').api.WrapText = True  # 设置自动换行

        wb_1.sheets['01'].api.Visible = False  # 隐藏 sheet 01
        wb_1.sheets['02'].api.Visible = False  # 隐藏 sheet 02

        """在故障分析表格输入值"""
        wb_2 = app.books.open(self.__fenxi_ad)  # 不同的电脑需替换不同的地址
        sht_a = wb_2.sheets[1]
        rgn = sht_a.range("a4").expand('right')
        ab = len(rgn) + 1
        ak = {4: 1, 5: 2, 6: 3}
        for k, v in ak.items():
            sht_a.range(k, ab).value = sht_3.range(v, 2).value  # 分别生成3个相关数据
        wb_2.save()
        app.kill()

        """重新打开分析表格"""
        app = xw.App(visible=True, add_book=False)  # visible值为False为后台运行
        app.books.open(self.__fenxi_ad)  # 不同的电脑需替换不同的地址

    def btn_autoSDM_clicked(self):
        self.__new_file()
        self.__server_lj('免疫模块机装机及开箱异常')
        for filename in os.walk('./fujian/'):
            print(filename)
            path = './fujian/' + str(filename[2][0])
            print(path)
            self.__generateData(path)

        self.btn_autoSDM.setEnabled(True)
        self.btn_autoSDM.setStyleSheet('''QPushButton#btn_autoSDM{
                               background-color:rgb(255, 255, 127);}QPushButton#btn_autoSDM:hover{
                               background-color:red;}''')
        self.flag = True

    def autoSDM_clicked(self):
        if self.flag==True:
            self.flag=False
            autosdm = Thread(target=self.btn_autoSDM_clicked)
            autosdm.start()
            self.btn_autoSDM.setEnabled(False)
            self.btn_autoSDM.setStyleSheet('''QPushButton#btn_autoSDM{
                                               background-color:red;}''')
    def btn_autoITM_clicked(self):
        self.__new_file()
        self.__server_lj('流水线B1客户信息表')
        today = datetime.datetime.today()
        month = str(today.month) + '月'
        for filename in os.walk('./fujian/'):
            print(filename)
            path = './fujian/' + str(filename[2][0])
            self.__itm_generateData(path, month)

        self.btn_autoITM.setEnabled(True)
        self.btn_autoITM.setStyleSheet('''QPushButton#btn_autoITM{
                               background-color:rgb(255, 255, 127);}QPushButton#btn_autoITM:hover{
                               background-color:red;}''')
        self.flag_1 = True

    def autoITM_clicked(self):
        if self.flag_1 == True:
            self.flag_1 = False
            autoitm = Thread(target=self.btn_autoITM_clicked)
            autoitm.start()
            self.btn_autoITM.setEnabled(False)
            self.btn_autoITM.setStyleSheet('''QPushButton#btn_autoITM{
                       background-color:red;}''')

    def __new_file(self):  # 创建新文件夹

        b = os.getcwd()  # 获取当前文件路径（即打开文件.py路径）
        c = b + '/fujian/'  # 获取当前文件路径# 清空筛选后日志文件夹内容
        try:
            shutil.rmtree(c)
            time.sleep(0.5)
            os.mkdir(c)  # 新建文件夹
        except:
            pass
        return c

    def __decode_str(self, msg):  # 字符编码转换
        value, charset = decode_header(msg)[0]
        if charset:
            value = value.decode(charset)
        return value

    def __server_lj(self, dot):
        with open('config.txt', 'r') as f1:
            config = f1.readlines()
        for i in range(0, len(config)):
            config[i] = config[i].rstrip('\n')
        # print(config)
        # POP3服务器、用户名、密码
        host = config[0]  # pop.163.com
        username = config[1]  # 用户名
        password = config[2]  # 密码

        # 连接到POP3服务器
        server = poplib.POP3(host)

        # 身份验证
        server.user(username)
        server.pass_(password)  # 参数是你的邮箱密码，如果出现poplib.error_proto: b'-ERR login fail'，就用开启POP3服务时拿到的授权码

        # stat()返回邮件数量和占用空间:
        print('Messages: %s. Size: %s' % server.stat())

        # 可以查看返回的列表类似[b'1 82923', b'2 2184', ...]
        resp, mails, octets = server.list()
        # 倒序遍历邮件
        index = len(mails)
        for i in range(index, 0, -1):
            time.sleep(0.5)
            # for i in range(1, index + 1):# 顺序遍历邮件
            print(i)
            resp, lines, octets = server.retr(i)
            # lines存储了邮件的原始文本的每一行,
            # 邮件的原始文本:
            msg_content = b'\r\n'.join(lines).decode('GBK')  # 使用utf-8编码会错误，更换为gbk则不会。
            # 解析邮件:
            msg = Parser().parsestr(msg_content)
            Subject = self.__decode_str(msg.get('Subject'))  # 使主题名可读
            print(Subject)
            if dot in Subject:
                self.__get_att(msg)
            if os.listdir('./fujian/'):
                break

    def __get_att(self, msg):  # 解析邮件，下载附件
        attachment_files = []
        for part in msg.walk():
            file_name = part.get_filename()  # 获取附件名称类型
            if file_name:
                h = email.header.Header(file_name)
                dh = email.header.decode_header(h)  # 对附件名称进行解码
                filename = dh[0][0]
                if dh[0][1]:
                    filename = self.__decode_str(str(filename, dh[0][1]))  # 将附件名称可读化
                    # print( filename )
                    # filename = filename.encode("utf-8")
                data = part.get_payload(decode=True)  # 下载附件
                att_file = open('./fujian/' + filename, 'wb')  # 在指定目录下创建文件，注意二进制文件需要用wb模式打开
                attachment_files.append(filename)
                att_file.write(data)  # 保存附件
                att_file.close()


class pty():
    def __init__(self):
        pass

    @staticmethod
    def dot(path):
        with open(path, 'r', ) as t:
            return t.read()


if __name__ == '__main__':
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)  # 自适应屏幕分辨率
    app = QApplication(sys.argv)
    aot = pty.dot('./style.qss')
    window = MainWindow()
    window.setStyleSheet(aot)
    window.show()
    sys.exit(app.exec_())
