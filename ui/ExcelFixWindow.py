import json
import os
import shutil
import time

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QMainWindow, QApplication, QFileDialog
from PySide6.QtGui import QMouseEvent, QGuiApplication
from qfluentwidgets import InfoBar, InfoBarPosition

from fixfont import ExcelProcessorThread
from ui.excelfix import Ui_MainWindow
from utils import glo
from utils.customGrips import CustomGrip
from PySide6.QtCore import Qt, QPropertyAnimation


class ExcelFixWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_workpath = os.getcwd()
        self.animation_window = None
        self.excel_input_path = None
        self.excel_output_path = None
        self.excel_input_name = None
        self.excel_output_name = None
        self.excel_result_path = None
        # --- 拖动窗口 改变窗口大小 --- #
        self.center()  # 窗口居中
        self.left_grip = CustomGrip(self, Qt.LeftEdge, True)
        self.right_grip = CustomGrip(self, Qt.RightEdge, True)
        self.top_grip = CustomGrip(self, Qt.TopEdge, True)
        self.bottom_grip = CustomGrip(self, Qt.BottomEdge, True)
        self.setAcceptDrops(True)  # ==> 设置窗口支持拖动（必须设置）
        # --- 拖动窗口 改变窗口大小 --- #

        # --- 加载UI --- #
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        # --- 加载UI --- #

        # --- 导入 导出 --- #
        self.ui.import_excel.clicked.connect(self.importExcel)
        self.ui.export_excel.clicked.connect(self.exportExcel)
        # --- 导入 导出 --- #

        # --- 修复字体 --- #
        self.fixfont_thread = ExcelProcessorThread()
        self.ui.fix_font.clicked.connect(self.fixFont)
        # --- 修复字体 --- #

    # 导入Excel
    def importExcel(self):
        # 获取上次选择文件的路径
        config_file = f'{self.current_workpath}/config/file.json'
        config = json.load(open(config_file, 'r', encoding='utf-8'))
        file_path = config['file_path']
        if not os.path.exists(file_path):
            file_path = os.getcwd()
        file, _ = QFileDialog.getOpenFileName(
            self,  # 父窗口对象
            "导入Excel文件",  # 标题
            file_path,  # 默认打开路径为当前路径
            "文件类型 (*.xls *.xlsx)"  # 选择类型过滤项，过滤内容在括号中
        )
        if file:
            self.excel_input_path = file
            glo.set_value('excel_input_path', self.excel_input_path)
            self.excel_input_name = os.path.basename(self.excel_input_path)
            # 获取self.excel_input_name 的文件后缀
            if self.excel_input_name.endswith('.xls'):
                self.excel_output_name = self.excel_input_name.replace(".xls", "_fixed.xls")
            else:
                self.excel_output_name = self.excel_input_name.replace(".xlsx", "_fixed.xlsx")
            print(self.excel_output_name)
            self.excel_result_path = os.path.join(self.current_workpath, 'result', self.excel_output_name)
            self.showStatus('导入Excel文件：{}'.format(os.path.basename(self.excel_input_path)))
            config['file_path'] = os.path.dirname(self.excel_input_path)
            config_json = json.dumps(config, ensure_ascii=False, indent=2)
            with open(config_file, 'w', encoding='utf-8') as f:
                f.write(config_json)

    # 导出Excel
    def exportExcel(self):
        config_file = f'{self.current_workpath}/config/save.json'
        config = json.load(open(config_file, 'r', encoding='utf-8'))
        save_path = config['save_path']
        if not os.path.exists(save_path):
            save_path = os.getcwd()
        if not self.excel_result_path:
            self.showStatus('请先导入Excel文件，并进行修复！')
            return
        self.excel_output_path, _ = QFileDialog.getSaveFileName(
            self,  # 父窗口对象
            "导出Excel文件",  # 标题
            os.path.join(save_path, str(self.excel_output_name)),  # 起始目录
            "文件类型 (*.xls *.xlsx)"  # 选择类型过滤项，过滤内容在括号中
        )
        if self.excel_output_path:
            try:
                if os.path.exists(self.excel_result_path):
                    shutil.copy(self.excel_result_path, self.excel_output_path)
                    self.showStatus('成功导出Excel： {}'.format(self.excel_output_path))
                else:
                    self.showStatus('请修复成功之后，再进行导出！')
            except Exception as err:
                self.showStatus(f"导出Excel失败: {err}")
        config['save_path'] = self.excel_output_path
        config_json = json.dumps(config, ensure_ascii=False, indent=2)
        with open(config_file, 'w', encoding='utf-8') as f:
            f.write(config_json)

    # 展示修复结果
    def resultInfo(self, x):
        if "成功" in x:
            self.showStatus('Excel已修复成功')
            self.createSuccessInfoBar("修复结果", "Excel已修复成功，请导出Excel！")
        else:
            self.showStatus('Excel修复失败')
            self.createErrorInfoBar("修复结果", x)

    # Excel 修复字体
    def fixFont(self):
        if self.excel_input_path is None:
            self.showStatus('请先导入Excel文件！')
            return
        self.showStatus('正在修复Excel字体...')
        self.fixfont_thread.set_path(self.excel_input_path, self.excel_result_path)
        self.fixfont_thread.send_result.connect(lambda x: self.resultInfo(x))
        self.fixfont_thread.start()

    # 显示通知消息
    def showStatus(self, msg):
        self.ui.message_box.setText(msg)

    # --- 拖动窗口 改变窗口大小 窗口居中 --- #
    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.LeftButton:
            self.mouse_start_pt = event.globalPosition().toPoint()
            self.window_pos = self.frameGeometry().topLeft()
            self.drag = True

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        if self.drag:
            distance = event.globalPosition().toPoint() - self.mouse_start_pt
            self.move(self.window_pos + distance)

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.LeftButton:
            self.drag = False

    def center(self):
        # PyQt6获取屏幕参数
        screen = QGuiApplication.primaryScreen().size()
        size = self.geometry()
        self.move((screen.width() - size.width()) / 2,
                  (screen.height() - size.height()) / 2 - 10)

    def resizeEvent(self, event):
        # Update Size Grips
        self.resizeGrip()

    def resizeGrip(self):
        self.left_grip.setGeometry(0, 10, 10, self.height())
        self.right_grip.setGeometry(self.width() - 10, 10, 10, self.height())
        self.top_grip.setGeometry(0, 0, self.width(), 10)
        self.bottom_grip.setGeometry(0, self.height() - 10, self.width(), 10)

    # --- 拖动窗口 改变窗口大小 窗口居中 --- #

    # --- InfoBar --- #
    def createErrorInfoBar(self, title, content, duration=2000):
        return InfoBar.error(
            title=title,
            content=content,
            orient=Qt.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP,
            duration=duration,  # won't disappear automatically
            parent=self
        )

    def createSuccessInfoBar(self, text, content, duration=2000):
        # convenient class mothod
        return InfoBar.success(
            title=text,
            content=content,
            orient=Qt.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP,
            # position='Custom',   # NOTE: use custom info bar manager
            duration=duration,
            parent=self
        )

    # --- InfoBar --- #

    # --- 关闭窗口 自动删除Result缓存 --- #
    def closeEvent(self, event):
        if not self.animation_window:
            self.animation_window = QPropertyAnimation(self, b"windowOpacity")
            self.animation_window.setStartValue(1)
            self.animation_window.setEndValue(0)
            self.animation_window.setDuration(500)
            self.animation_window.start()
            self.animation_window.finished.connect(self.close)
            try:
                if os.path.exists(f"{self.current_workpath}/result/"):
                    shutil.rmtree(f"{self.current_workpath}/result/")
            except Exception:
                time.sleep(1)
                if os.path.exists(f"{self.current_workpath}/result/"):
                    shutil.rmtree(f"{self.current_workpath}/result/")
            os.makedirs(f"{self.current_workpath}/result/", exist_ok=True)
            event.ignore()
    # --- 关闭窗口 自动删除Result缓存 --- #


# 测试
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
