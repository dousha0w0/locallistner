import datetime
import sys

import PyQt5
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QTextEdit, QPushButton, QWidget
from PyQt5.QtCore import QThread, pyqtSlot
import os
import time
import shutil
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import win32print
import win32api


def print_file(file_path, printer):
    win32print.SetDefaultPrinter(printer)
    win32api.ShellExecute(0, "print", file_path, None, ".", 0)


class FileHandler(FileSystemEventHandler):
    def __init__(self, monitor_directories, file_patterns, printers, move_directory, log_signal):
        self.monitor_directories = monitor_directories
        self.file_patterns = file_patterns
        self.printers = printers
        self.move_directory = move_directory
        self.log_signal = log_signal

    def on_created(self, event):
        for index, directory in enumerate(self.monitor_directories):
            if event.src_path.startswith(directory):
                for file_pattern in self.file_patterns[index]:
                    if event.src_path.endswith(file_pattern):
                        print_file(event.src_path, self.printers[index])
                        self.move_file(event.src_path)
                        self.write_log(event.src_path)

    # 移动文件
    def move_file(self, file_path):
        shutil.move(file_path, os.path.join(self.move_directory, os.path.basename(file_path)))

    # 记录日志
    def write_log(self, file_path):
        log_message = f'[{time.strftime("%Y-%m-%d %H:%M:%S")}] Printed and moved file: {file_path}\n'
        self.log_signal.emit(log_message)
        filename = datetime.datetime.today().now().strftime('%Y-%m-%d') + '.txt'
        if not os.path.isfile(filename):
            with open(filename, 'w') as f:
                f.write(f'{log_message}\n')
        else:
            print('file exist')
            with open(filename, 'a') as f:
                f.write(f'{log_message}\n')


class MonitorThread(QThread):
    log_signal = PyQt5.QtCore.pyqtSignal(str)

    def __init__(self, mn_func):
        super().__init__()
        self.monitor_function = mn_func

    def run(self):
        self.monitor_function(self.log_signal)


def monitor_function(log_signal):
    # 配置监控目录、文件模式、打印机名称
    monitor_directories = ["E:\\新建文件夹\\Microsoft Print to PDF", "E:\\新建文件夹\\Monitor2"]
    file_patterns = [[".pdf", ".docx", ".jpg", ".png"]]
    printers = ["Microsoft Print to PDF"]

    # 配置移动目录
    move_directory = "E:\\新建文件夹\\Processed"

    # 确保移动目录存在
    if not os.path.exists(move_directory):
        os.makedirs(move_directory)

    # 创建文件处理器并配置监视器
    event_handler = FileHandler(monitor_directories, file_patterns, printers, move_directory, log_signal)
    observer = Observer()

    for directory in monitor_directories:
        observer.schedule(event_handler, directory, recursive=True)

    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # 设置主窗口属性
        self.setWindowTitle("文件监控器")
        self.setGeometry(100, 100, 800, 600)

        # 创建日志显示框、开始监控按钮和停止监控按钮
        self.log_text_edit = QTextEdit(self)
        self.log_text_edit.setReadOnly(True)

        self.start_button = QPushButton("开始监控", self)
        self.start_button.clicked.connect(self.start_monitor)

        self.stop_button = QPushButton("停止监控", self)
        self.stop_button.clicked.connect(self.stop_monitor)

        # 创建布局并设置为主窗口的中心窗口
        layout = QVBoxLayout()
        layout.addWidget(self.log_text_edit)
        layout.addWidget(self.start_button)
        layout.addWidget(self.stop_button)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # 初始化监控线程
        self.monitor_thread = MonitorThread(monitor_function)

    def start_monitor(self):
        self.monitor_thread.start()
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)

    def stop_monitor(self):
        self.monitor_thread.terminate()
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)

    @pyqtSlot(str)
    def update_log(self, log_message):
        self.log_text_edit.append(log_message)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.monitor_thread.log_signal.connect(main_window.update_log)
    main_window.show()
    sys.exit(app.exec_())
