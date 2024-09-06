import sys
import os
# 将ui目录添加到系统路径中
sys.path.append(os.path.join(os.getcwd(), "ui"))
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication
from ui.ExcelFixWindow import ExcelFixWindow
from utils import glo

if __name__ == '__main__':
    app = QApplication([])  # 创建应用程序实例
    app.setWindowIcon(QIcon('images/swimmingliu.ico'))  # 设置应用程序图标

    # 为整个应用程序设置样式表，去除所有QFrame的边框
    app.setStyleSheet("QFrame { border: none; }")

    # 创建窗口实例
    excelfix_window = ExcelFixWindow()

    # 初始化全局变量管理器，并设置值
    glo._init()  # 初始化全局变量空间
    glo.set_value('excelfix_window', excelfix_window)  # 存储randy_window窗口实例

    # 从全局变量管理器中获取窗口实例
    excelfix_window_glo = glo.get_value('excelfix_window')

    # 显示yoloshow窗口
    excelfix_window_glo.show()
    app.exec()  # 启动应用程序的事件循环