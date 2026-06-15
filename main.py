# -*- coding: utf-8 -*-
"""
ZURU Shipment Assistant - PySide6 GUI
选择接单表 + 选择出货文件夹 + 开始处理
"""

import os
import sys
import logging
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QTextEdit, QGroupBox,
    QFileDialog, QMessageBox, QProgressBar, QDateEdit,
)
from PySide6.QtCore import Qt, QThread, Signal, QDate
from PySide6.QtGui import QFont

from shipment_processor import process

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S',
)


class ProcessThread(QThread):
    """后台处理线程"""
    log_signal = Signal(str)
    finished_signal = Signal(str, dict)   # (output_path, stats)
    error_signal = Signal(str)

    def __init__(self, order_path, folder_path, mark_date=''):
        super().__init__()
        self.order_path = order_path
        self.folder_path = folder_path
        self.mark_date = mark_date

    def run(self):
        try:
            result, info = process(
                self.order_path,
                self.folder_path,
                mark_date=self.mark_date,
                log_callback=self.log_signal.emit,
            )
            if result is None:
                self.error_signal.emit(str(info))
            else:
                self.finished_signal.emit(result, info)
        except Exception as e:
            self.error_signal.emit(str(e))


class ShipmentApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("ZURU Shipment Assistant - 出货助手")
        self.setMinimumSize(750, 600)
        self.resize(820, 650)

        self.order_path = ""
        self.folder_path = ""
        self.worker = None
        self._output_path = None

        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(15, 10, 15, 10)
        layout.setSpacing(10)

        # ---- 标题 ----
        title = QLabel("ZURU 出货助手")
        title.setFont(QFont("Microsoft YaHei", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        subtitle = QLabel("自动读取出货资料 → 匹配接单表 → 扣减库存 → 生成新文件")
        subtitle.setFont(QFont("Microsoft YaHei", 9))
        subtitle.setStyleSheet("color: #888;")
        subtitle.setAlignment(Qt.AlignCenter)
        layout.addWidget(subtitle)

        # ---- 接单表 ----
        order_group = QGroupBox("接单表")
        order_layout = QHBoxLayout(order_group)
        self.order_edit = QLineEdit()
        self.order_edit.setReadOnly(True)
        self.order_edit.setPlaceholderText("请选择接单表 Excel 文件...")
        order_layout.addWidget(self.order_edit, 1)
        btn_order = QPushButton("选择接单表")
        btn_order.setFixedWidth(120)
        btn_order.clicked.connect(self._select_order)
        order_layout.addWidget(btn_order)
        layout.addWidget(order_group)

        # ---- 出货文件夹 ----
        folder_group = QGroupBox("出货资料文件夹")
        folder_layout = QHBoxLayout(folder_group)
        self.folder_edit = QLineEdit()
        self.folder_edit.setReadOnly(True)
        self.folder_edit.setPlaceholderText("请选择出货资料文件夹...")
        folder_layout.addWidget(self.folder_edit, 1)
        btn_folder = QPushButton("选择文件夹")
        btn_folder.setFixedWidth(120)
        btn_folder.clicked.connect(self._select_folder)
        folder_layout.addWidget(btn_folder)
        layout.addWidget(folder_group)

        # ---- 备注日期 ----
        date_group = QGroupBox("备注日期（统一填入接单表备注列）")
        date_layout = QHBoxLayout(date_group)
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setDisplayFormat("M月d日")
        self.date_edit.setFont(QFont("Microsoft YaHei", 12))
        self.date_edit.setFixedHeight(36)
        date_layout.addWidget(self.date_edit, 1)

        btn_today = QPushButton("今天")
        btn_today.setFixedWidth(60)
        btn_today.clicked.connect(
            lambda: self.date_edit.setDate(QDate.currentDate()))
        date_layout.addWidget(btn_today)
        layout.addWidget(date_group)

        # ---- 进度条 ----
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(22)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("就绪")
        self.progress_bar.setValue(0)
        self.progress_bar.setStyleSheet(
            "QProgressBar { border: 1px solid #ccc; border-radius: 3px; "
            "text-align: center; }"
            "QProgressBar::chunk { background-color: #4CAF50; }"
        )
        layout.addWidget(self.progress_bar)

        # ---- 开始处理 ----
        self.start_btn = QPushButton("开始处理")
        self.start_btn.setFont(QFont("Microsoft YaHei", 13, QFont.Bold))
        self.start_btn.setFixedHeight(48)
        self.start_btn.setStyleSheet(
            "QPushButton { background-color: #4CAF50; color: white; "
            "border-radius: 4px; }"
            "QPushButton:hover { background-color: #45a049; }"
            "QPushButton:disabled { background-color: #999; }"
        )
        self.start_btn.setCursor(Qt.PointingHandCursor)
        self.start_btn.clicked.connect(self._start)
        layout.addWidget(self.start_btn)

        # ---- 日志 ----
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        self.log_text.setStyleSheet(
            "QTextEdit { background-color: #1e1e1e; color: #d4d4d4; }")
        log_layout.addWidget(self.log_text)
        layout.addWidget(log_group, 1)

        # ---- 输出文件行 ----
        out_row = QHBoxLayout()
        self.output_label = QLabel("输出文件: 尚未生成")
        self.output_label.setFont(QFont("Microsoft YaHei", 9))
        self.output_label.setStyleSheet("color: #666;")
        out_row.addWidget(self.output_label, 1)

        self.open_btn = QPushButton("打开输出文件")
        self.open_btn.setFixedWidth(130)
        self.open_btn.setEnabled(False)
        self.open_btn.clicked.connect(self._open_output)
        out_row.addWidget(self.open_btn)
        layout.addLayout(out_row)

    # ---- 选择 ----

    def _select_order(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择接单表", "",
            "Excel文件 (*.xlsx *.xls);;所有文件 (*.*)")
        if path:
            self.order_path = path
            self.order_edit.setText(path)

    def _select_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "选择出货资料文件夹", "")
        if folder:
            self.folder_path = folder
            self.folder_edit.setText(folder)

    # ---- 日志 ----

    def _log(self, msg):
        self.log_text.append(msg)
        sb = self.log_text.verticalScrollBar()
        sb.setValue(sb.maximum())

    # ---- 处理 ----

    def _start(self):
        if self.worker and self.worker.isRunning():
            return
        if not self.order_path:
            QMessageBox.warning(self, "提示", "请先选择接单表")
            return
        if not self.folder_path:
            QMessageBox.warning(self, "提示", "请先选择出货资料文件夹")
            return
        if not os.path.isfile(self.order_path):
            QMessageBox.critical(self, "错误",
                                 f"接单表不存在:\n{self.order_path}")
            return
        if not os.path.isdir(self.folder_path):
            QMessageBox.critical(self, "错误",
                                 f"文件夹不存在:\n{self.folder_path}")
            return

        # 重置
        self.log_text.clear()
        self._output_path = None
        self.open_btn.setEnabled(False)
        self.output_label.setText("输出文件: 处理中...")
        self.output_label.setStyleSheet("color: #666;")
        self.start_btn.setEnabled(False)
        self.start_btn.setText("处理中...")
        self.progress_bar.setMaximum(0)
        self.progress_bar.setFormat("处理中...")

        # 格式化日期为 "X月X日"
        qd = self.date_edit.date()
        mark_date = f"{qd.month()}月{qd.day()}日"

        self.worker = ProcessThread(self.order_path, self.folder_path,
                                    mark_date)
        self.worker.log_signal.connect(self._log)
        self.worker.finished_signal.connect(self._on_success)
        self.worker.error_signal.connect(self._on_error)
        self.worker.finished.connect(self._reset_ui)
        self.worker.start()

    def _on_success(self, output_path, stats):
        self._output_path = output_path
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("处理完成！")

        self.output_label.setText(
            f"输出文件: {os.path.basename(output_path)}")
        self.output_label.setStyleSheet("color: #2e7d32; font-weight: bold;")
        self.open_btn.setEnabled(True)

        sub_skipped = stats.get('sub_skipped', 0)
        msg_text = (
            f"接单表已更新！\n\n"
            f"成功匹配: {stats['processed']} 组\n"
            f"子行跳过: {sub_skipped} 组（正常）\n"
            f"未找到匹配: {stats['not_found']} 组\n"
            f"标记出货行: {stats['rows_marked']} 行\n\n"
            f"输出文件:\n{output_path}")

        report = stats.get('error_report')
        if report:
            msg_text += (
                f"\n\n--- 库存异常 ---\n"
                f"发现库存不足/未匹配项，已生成错误报告:\n"
                f"{os.path.basename(report)}")

        has_errors = stats['not_found'] > 0
        if has_errors:
            QMessageBox.warning(self, "处理完成（有异常）", msg_text)
        else:
            QMessageBox.information(self, "处理完成", msg_text)

    def _on_error(self, msg):
        self._log(f"\n处理失败: {msg}")
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("处理失败")
        self.output_label.setText("输出文件: 处理失败")
        self.output_label.setStyleSheet("color: #c62828;")
        QMessageBox.critical(self, "处理失败", msg)

    def _open_output(self):
        if self._output_path and os.path.isfile(self._output_path):
            os.startfile(self._output_path)

    def _reset_ui(self):
        self.worker = None
        self.start_btn.setEnabled(True)
        self.start_btn.setText("开始处理")
        if self.progress_bar.maximum() == 0:
            self.progress_bar.setMaximum(100)
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("就绪")


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ShipmentApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
