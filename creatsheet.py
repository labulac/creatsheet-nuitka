from typing import NewType
from PyQt5.QtWidgets import QMainWindow, QApplication, QGraphicsOpacityEffect
from PyQt5.QtCore import *
from need.Ui_creatsheet import Ui_MainWindow
import sys, requests, threading, time, openpyxl, shutil, os, configparser, subprocess


def get_resource_path(relative_path):  # 利用此函数实现资源路径的定位
    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS  # 获取临时资源
    else:
        base_path = os.path.abspath(".")  # 获取当前路径
    return os.path.join(base_path, relative_path)  # 绝对路径


MUBAN_PATH = get_resource_path(os.path.join("resources", "muban.xlsx"))
CONF = get_resource_path(os.path.join("resources", "labulac.conf"))


class UI(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(UI, self).__init__()
        self.setupUi(self)

        self.ms = MySignals()
        self.ms.update_object_text.connect(self.update_gui_text)

        _translate = QCoreApplication.translate

        self.label_5.setText(
            _translate(
                "MainWindow",
                "<html><head/><body><p><span style=\" color:#ff007f;\">新建账单小助手v2.6</span></p></body></html>"
            ))
        self.currentversion = "26"
        self.download_finish = '0'
        self.NETWORK = True

        self.pushButton_3.clicked.connect(self.quit)
        self.pushButton.clicked.connect(self.create_new_sheet)
        self.pushButton_2.clicked.connect(self.write_date)

        self.progressBar.setGraphicsEffect(self.op(0))

        # 不显示标题栏并置顶
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)

        # 灰度显示最大化与关闭按钮
        #self.setWindowFlags(Qt.WindowMinimizeButtonHint)

        yiyan_update_thread = threading.Thread(target=self.yiyan_update)
        yiyan_update_thread.start()

    def op(self, str):
        op = QGraphicsOpacityEffect()
        op.setOpacity(str)
        return op

    # 下面三个函数是重写鼠标事件的

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.m_flag = True
            self.m_Position = event.globalPos() - self.pos()  #获取鼠标相对窗口的位置
            event.accept()

    def mouseMoveEvent(self, QMouseEvent):
        if Qt.LeftButton and self.m_flag:
            self.move(QMouseEvent.globalPos() - self.m_Position)  #更改窗口位置
            QMouseEvent.accept()

    def mouseReleaseEvent(self, QMouseEvent):
        self.m_flag = False

    def conf(self):
        # 检查配置是否存在
        if os.path.isfile("upgrade.bat"):
            os.remove("upgrade.bat")

        try:
            requests.get(
                'https://sc.ftqq.com/SCU126653T812824e9c91dc2707f0f712c5cc598bd5faf9a749f235.send?text=creatsheet启动啦~'
            )
            #github_net = 'https://cdn.jsdelivr.net/gh/labulac/creatsheet@main/creatsheet_info.js'
            github_net = 'https://raw.fastgit.org/labulac/creatsheet/main/creatsheet_info.js'
            github_conf = requests.get(github_net).text

            if github_conf != "":
                with open('D:/labulac.conf', 'w') as f:
                    f.write(github_conf)
                print('更新最新配置完成！')
        except:
            print('检查最新配置失败，网络异常！')

            self.NETWORK = False

            if os.path.exists('D:/labulac.conf'):
                print('配置存在，飘过~~')
            else:
                shutil.copy(CONF, 'D:/labulac.conf')
                print('默认配置已经生成！')

        cf = configparser.ConfigParser()
        cf.read("D:/labulac.conf", encoding="utf-8")

        self.yuan = MUBAN_PATH
        self.xian = cf.get("dir", "xian")
        self.newversion = cf.get("update", "newversion")
        self.downloadurl = cf.get("update", "downloadurl")
        self.yiyan_url = cf.get("yiyan", "yiyan")

    def update_gui_text(self, fb, text):
        fb.setText(str(text))

    def yiyan_update(self):

        # 获取配置文件
        self.conf()
        print('network=' + str(self.NETWORK))

        if self.NETWORK != False:

            self.yiyan_change_thread = yiyan_change_thread(self.yiyan_url)
            self.yiyan_change_thread.yiyan_change_signal.connect(
                self.yiyan_changed_value)
            self.yiyan_change_thread.start()

            self.check_update()

        else:
            print("一言网络错误")
            yiyan_text = ("开心每一天！")

            self.ms.update_object_text.emit(self.yiyan_label, str(yiyan_text))

    def check_update(self):

        if self.currentversion < self.newversion:
            self.checkupdate = True
        else:
            self.checkupdate = False
            print("没有发现新的版本")

        if self.checkupdate == True:
            print("发现新的版本！！！")
            self.progressBar.setGraphicsEffect(self.op(1))
            self.ms.update_object_text.emit(self.label_4, str("下载新版本中："))

            the_url = self.downloadurl
            the_filesize = requests.get(the_url,
                                        stream=True).headers['Content-Length']
            the_filepath = "creatsheet1.exe"
            the_fileobj = open(the_filepath, 'wb')
            # 创建下载线程
            self.downloadThread = downloadThread(the_url,
                                                 the_filesize,
                                                 the_fileobj,
                                                 buffer=10240)
            self.downloadThread.download_proess_signal.connect(
                self.set_progressbar_value)
            self.downloadThread.start()

    def quit(self):
        if self.download_finish == '1':
            self.WriteRestartCmd("creatsheet.exe")

        sys.exit()

    def yiyan_changed_value(self, yiyan_text):

        self.ms.update_object_text.emit(self.yiyan_label, str(yiyan_text))

    # 设置进度条
    def set_progressbar_value(self, value):
        self.progressBar.setValue(value)
        if value == 100:
            self.download_finish = '1'
            self.ms.update_object_text.emit(self.pushButton_3, str('重启并更新'))
            self.pushButton_3.setStyleSheet("color: red")
            self.ms.update_object_text.emit(self.label_4, str("下载完毕，重启试试吧！"))
            return

    def WriteRestartCmd(self, exe_name):
        b = open("upgrade.bat", 'w')
        TempList = "@echo off\n"
        TempList += "if not exist " + exe_name + " exit \n"
        TempList += "timeout /nobreak /t 3\n"
        TempList += "del " + os.path.realpath(sys.argv[0]) + "\n"
        TempList += "rename creatsheet1.exe creatsheet.exe\n"
        TempList += "start " + exe_name
        b.write(TempList)
        b.close()
        subprocess.Popen("upgrade.bat")

    def create_new_sheet(self):
        if self.lineEdit.text() != '':
            try:
                if os.path.exists(self.xian + self.lineEdit.text() +
                                  ".xlsx") is True:
                    self.ms.update_object_text.emit(self.pushButton,
                                                    str("账单名称已存在"))
                else:
                    shutil.copy(self.yuan,
                                self.xian + self.lineEdit.text() + ".xlsx")
                    self.pushButton.setEnabled(False)
                    self.ms.update_object_text.emit(self.pushButton,
                                                    str("已创建"))
                self.pushButton_2.setEnabled(True)
            except:
                print("路径不正确!")
                print(self.comboBox_4.currentText() + '年' +
                      self.comboBox_5.currentText() + '月' +
                      self.comboBox_6.currentText() + '日')
        else:
            print("账单名称为空！")

    def write_date(self):
        cal2 = self.comboBox_4.currentText(
        ) + '年' + self.comboBox_5.currentText(
        ) + '月' + self.comboBox_6.currentText() + '日'

        cal1 = self.comboBox_3.currentText(
        ) + '年' + self.comboBox_2.currentText(
        ) + '月' + self.comboBox_1.currentText() + '日'

        try:

            wb = openpyxl.load_workbook(self.xian + self.lineEdit.text() +
                                        ".xlsx")
            sheet = wb['Sheet1']

            sheet['A2'].value = cal1
            sheet['C2'].value = cal2
            wb.save(self.xian + self.lineEdit.text() + ".xlsx")
            self.pushButton_2.setEnabled(False)
            self.ms.update_object_text.emit(self.pushButton_2, str("已写入完成"))
        except:
            print("写入失败")


# 自定义信号源
class MySignals(QObject):
    update_object_text = pyqtSignal(QObject, str)


# 一言语句更新线程
class yiyan_change_thread(QThread):
    yiyan_change_signal = pyqtSignal(str)

    def __init__(self, url):
        super(yiyan_change_thread, self).__init__()

        self.ms = MySignals()
        self.yiyan_url = url

    def run(self):

        while True:
            yiyan_text = requests.get(self.yiyan_url).text
            print(yiyan_text)
            self.yiyan_change_signal.emit(str(yiyan_text))
            time.sleep(3)

            


#下载线程
class downloadThread(QThread):
    download_proess_signal = pyqtSignal(int)  #创建信号

    def __init__(self, url, filesize, fileobj, buffer):
        super(downloadThread, self).__init__()
        self.url = url
        self.filesize = filesize
        self.fileobj = fileobj
        self.buffer = buffer

    def run(self):
        try:
            rsp = requests.get(self.url, stream=True)  #流下载模式
            offset = 0
            for chunk in rsp.iter_content(chunk_size=self.buffer):
                if not chunk: break
                self.fileobj.seek(offset)  #设置指针位置
                self.fileobj.write(chunk)  #写入文件
                offset = offset + len(chunk)
                proess = offset / int(self.filesize) * 100
                self.download_proess_signal.emit(int(proess))  #发送信号

            self.fileobj.close()  #关闭文件

            self.exit(0)  #关闭线程

        except Exception as e:
            self.download_finish = '0'
            requests.get(
                'https://sc.ftqq.com/SCU126653T812824e9c91dc2707f0f712c5cc598bd5faf9a749f235.send?text=checksheet下载失败啦~'
            )
            print(e)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = UI()
    window.show()
    sys.exit(app.exec_())