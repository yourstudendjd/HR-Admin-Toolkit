# -*- coding: utf-8 -*-
"""考勤与宿舍管理工具集 — 统一 PyQt5 应用程序

包含两大功能模块:
  1. 考勤汇总 — 打卡记录汇总 + 迟到判定 + 乐捐金额
  2. 宿舍分摊 — 水电住宿费按人按天分摊
"""

import os
import sys
from datetime import datetime
from calendar import monthrange

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QProgressBar, QTextEdit, QFileDialog,
    QMessageBox, QFrame, QCheckBox, QTabWidget, QComboBox, QLineEdit,
    QGridLayout, QGroupBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QPalette, QColor, QDragEnterEvent, QDropEvent

from attendance_processor import AttendanceProcessor
from dormitory_processor import process_allocation, parse_billing_month_from_filename

DEFAULT_EXCLUDE_NAMES = ["李诚维", "李天龙"]


# ══════════════════════════════════════════════════════════════════════════════
#  通用组件
# ══════════════════════════════════════════════════════════════════════════════

class DropZone(QFrame):
    """拖放区域组件"""
    file_dropped = pyqtSignal(str)

    def __init__(self, icon_text, prompt_text, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setMinimumHeight(90)
        self._icon = icon_text
        self._prompt = prompt_text
        self._normal_style = """
            DropZone {
                border: 2px dashed #8899AA;
                border-radius: 8px;
                background-color: #F5F7FA;
            }
            DropZone:hover {
                border-color: #4472C4;
                background-color: #EEF1F8;
            }
        """
        self._active_style = """
            DropZone {
                border: 2px solid #4472C4;
                border-radius: 8px;
                background-color: #DCE6F5;
            }
        """
        self.init_ui()

    def init_ui(self):
        self.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        self.setStyleSheet(self._normal_style)
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(4)

        icon = QLabel(self._icon)
        icon.setFont(QFont("Segoe UI Emoji", 18))
        icon.setAlignment(Qt.AlignCenter)
        layout.addWidget(icon)

        text = QLabel(self._prompt)
        text.setFont(QFont("Microsoft YaHei", 9))
        text.setAlignment(Qt.AlignCenter)
        text.setStyleSheet("color: #666;")
        layout.addWidget(text)

    def mousePressEvent(self, event):
        self.file_dropped.emit("__browse__")

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith(('.xlsx', '.xls')):
                    event.acceptProposedAction()
                    self.setStyleSheet(self._active_style)
                    return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self._normal_style)

    def dropEvent(self, event: QDropEvent):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(('.xlsx', '.xls')):
                self.file_dropped.emit(path)
                break
        self.setStyleSheet(self._normal_style)


# ══════════════════════════════════════════════════════════════════════════════
#  工作线程
# ══════════════════════════════════════════════════════════════════════════════

class AttendanceWorker(QThread):
    progress = pyqtSignal(int, int, str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, input_path, output_path, exclude_names, shift_path=None):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
        self.exclude_names = exclude_names
        self.shift_path = shift_path

    def run(self):
        try:
            processor = AttendanceProcessor(
                exclude_names=self.exclude_names,
                progress_callback=lambda c, t, m: self.progress.emit(c, t, m)
            )
            result = processor.process(
                self.input_path, self.output_path,
                shift_file_path=self.shift_path
            )
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


class DormitoryWorker(QThread):
    progress = pyqtSignal(int, int, str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, filepath, billing_year, billing_month, treat_prev_month=False):
        super().__init__()
        self.filepath = filepath
        self.billing_year = billing_year
        self.billing_month = billing_month
        self.treat_prev_month = treat_prev_month

    def run(self):
        try:
            result = process_allocation(
                self.filepath,
                self.billing_year,
                self.billing_month,
                treat_end_of_prev_month_as_living=self.treat_prev_month,
                progress_callback=lambda c, t, m: self.progress.emit(c, t, m)
            )
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


# ══════════════════════════════════════════════════════════════════════════════
#  标签页 1: 考勤汇总
# ══════════════════════════════════════════════════════════════════════════════

class AttendanceTab(QWidget):
    def __init__(self, log_writer):
        super().__init__()
        self.log = log_writer
        self.worker = None
        self.last_output_dir = None
        self.current_input = None
        self.current_shift = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 12, 20, 12)
        layout.setSpacing(8)

        # 打卡文件
        section1 = QLabel("① 考勤打卡文件")
        section1.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        section1.setStyleSheet("color: #333;")
        layout.addWidget(section1)

        self.drop_zone = DropZone("📂", "将 XLSX 打卡文件拖拽到此处\n或点击选择文件")
        self.drop_zone.file_dropped.connect(self.on_file_selected)
        layout.addWidget(self.drop_zone)

        self.file_label = QLabel("未选择打卡文件")
        self.file_label.setFont(QFont("Microsoft YaHei", 8))
        self.file_label.setStyleSheet("color: #999; padding: 1px 4px;")
        self.file_label.setWordWrap(True)
        layout.addWidget(self.file_label)

        self.select_btn = QPushButton("  选择打卡文件...")
        self.select_btn.setFont(QFont("Microsoft YaHei", 9))
        self.select_btn.setStyleSheet("""
            QPushButton {
                background-color: #4472C4; color: white; border: none;
                border-radius: 5px; padding: 6px 14px;
            }
            QPushButton:hover { background-color: #365A9E; }
        """)
        self.select_btn.clicked.connect(self.browse_punch_file)
        layout.addWidget(self.select_btn)

        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: #E0E0E0;")
        layout.addWidget(sep)

        # 班次表
        section2 = QLabel("② 班次表文件（可选，用于迟到判定）")
        section2.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        section2.setStyleSheet("color: #333;")
        layout.addWidget(section2)

        self.shift_zone = DropZone("📋", "将班次 XLSX 文件拖拽到此处\n或点击选择文件")
        self.shift_zone.file_dropped.connect(self.on_shift_selected)
        layout.addWidget(self.shift_zone)

        self.shift_label = QLabel("未选择班次表（将跳过迟到判定）")
        self.shift_label.setFont(QFont("Microsoft YaHei", 8))
        self.shift_label.setStyleSheet("color: #999; padding: 1px 4px;")
        self.shift_label.setWordWrap(True)
        layout.addWidget(self.shift_label)

        self.shift_btn = QPushButton("  选择班次表...")
        self.shift_btn.setFont(QFont("Microsoft YaHei", 9))
        self.shift_btn.setStyleSheet("""
            QPushButton {
                background-color: #6C757D; color: white; border: none;
                border-radius: 5px; padding: 6px 14px;
            }
            QPushButton:hover { background-color: #5A6268; }
        """)
        self.shift_btn.clicked.connect(self.browse_shift_file)
        layout.addWidget(self.shift_btn)

        self.late_check = QCheckBox("启用迟到检测（需加载班次表）")
        self.late_check.setFont(QFont("Microsoft YaHei", 9))
        self.late_check.setChecked(True)
        self.late_check.setEnabled(False)
        layout.addWidget(self.late_check)

        # 按钮
        btn_layout = QHBoxLayout()
        self.process_btn = QPushButton("  开始处理考勤")
        self.process_btn.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        self.process_btn.setEnabled(False)
        self.process_btn.setMinimumHeight(36)
        self.process_btn.setStyleSheet("""
            QPushButton {
                background-color: #2B8C3C; color: white; border: none;
                border-radius: 6px; padding: 8px 32px;
            }
            QPushButton:hover { background-color: #237030; }
            QPushButton:disabled { background-color: #C0C0C0; color: #F0F0F0; }
        """)
        self.process_btn.clicked.connect(self.start_processing)
        btn_layout.addWidget(self.process_btn)

        self.open_btn = QPushButton("  打开输出文件夹")
        self.open_btn.setFont(QFont("Microsoft YaHei", 9))
        self.open_btn.setEnabled(False)
        self.open_btn.setStyleSheet("""
            QPushButton {
                background-color: #E8ECF2; color: #333; border: 1px solid #C0C0C0;
                border-radius: 5px; padding: 6px 14px;
            }
            QPushButton:hover { background-color: #D4DAE4; }
            QPushButton:disabled { color: #C0C0C0; }
        """)
        self.open_btn.clicked.connect(self.open_output_folder)
        btn_layout.addWidget(self.open_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximum(100)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #D0D0D0; border-radius: 5px; text-align: center;
                background-color: #F0F0F0; height: 20px;
                font-family: "Microsoft YaHei"; font-size: 9px;
            }
            QProgressBar::chunk {
                background-color: #4472C4; border-radius: 4px;
            }
        """)
        layout.addWidget(self.progress_bar)

        self.status_label = QLabel("就绪")
        self.status_label.setFont(QFont("Microsoft YaHei", 9))
        self.status_label.setStyleSheet("color: #888;")
        layout.addWidget(self.status_label)
        layout.addStretch()

    def browse_punch_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择考勤打卡文件", os.path.expanduser("~"),
            "Excel 文件 (*.xlsx *.xls);;所有文件 (*.*)")
        if path:
            self.on_file_selected(path)

    def browse_shift_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择班次表文件", os.path.expanduser("~"),
            "Excel 文件 (*.xlsx *.xls);;所有文件 (*.*)")
        if path:
            self.on_shift_selected(path)

    def on_file_selected(self, path):
        if path == "__browse__":
            self.browse_punch_file(); return
        self.current_input = path
        self.file_label.setText(f"已选择: {os.path.basename(path)}")
        self.file_label.setStyleSheet("color: #2B579A; padding: 1px 4px;")
        self._update_process_btn()
        self.log(f"[考勤] 打卡文件: {path}")

    def on_shift_selected(self, path):
        if path == "__browse__":
            self.browse_shift_file(); return
        self.current_shift = path
        self.shift_label.setText(f"已选择: {os.path.basename(path)}")
        self.shift_label.setStyleSheet("color: #2B579A; padding: 1px 4px;")
        self.late_check.setEnabled(True)
        self.late_check.setChecked(True)
        self._update_process_btn()
        self.log(f"[考勤] 班次表: {path}")

    def _update_process_btn(self):
        self.process_btn.setEnabled(self.current_input is not None)

    def start_processing(self):
        if not self.current_input:
            return
        input_path = self.current_input
        base = os.path.splitext(os.path.basename(input_path))[0]
        input_dir = os.path.dirname(input_path)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(input_dir, f"{base}_汇总_{ts}.xlsx")
        self.last_output_dir = input_dir

        shift_path = None
        if self.late_check.isChecked() and self.current_shift:
            shift_path = self.current_shift

        self.process_btn.setEnabled(False)
        self.select_btn.setEnabled(False)
        self.shift_btn.setEnabled(False)
        self.open_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_label.setText("处理中...")
        self.status_label.setStyleSheet("color: #2B579A; font-weight: bold;")
        self.log("=" * 40)
        self.log("[考勤] 开始处理...")
        if shift_path:
            self.log("[考勤] 迟到检测: 已启用")
        else:
            self.log("[考勤] 迟到检测: 未启用（无班次表）")

        self.worker = AttendanceWorker(input_path, output_path, DEFAULT_EXCLUDE_NAMES, shift_path)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_progress(self, current, total, message):
        self.progress_bar.setValue(current)
        self.status_label.setText(message)
        if message:
            self.log(f"[考勤] {message}")

    def on_finished(self, result):
        self.progress_bar.setValue(100)
        self.status_label.setText("处理完成！")
        self.status_label.setStyleSheet("color: #2B8C3C; font-weight: bold;")
        self.process_btn.setEnabled(True)
        self.select_btn.setEnabled(True)
        self.shift_btn.setEnabled(True)
        self.open_btn.setEnabled(True)

        self.log("-" * 40)
        self.log("[考勤] 处理完成！")
        self.log(f"  总记录数: {result['total_records']}")
        self.log(f"  有效记录: {result['valid_records']}")
        self.log(f"  已排除: {result['excluded_count']}")
        self.log(f"  员工数: {result['employee_count']}")
        self.log(f"  日期数: {result['date_count']} ({result['date_range']})")
        if result.get('late_count', 0) > 0:
            self.log(f"  迟到人次: {result['late_count']}")

        msg_parts = [
            f"员工数: {result['employee_count']}",
            f"日期范围: {result['date_range']}",
        ]
        if result.get('late_count', 0) > 0:
            msg_parts.append(f"迟到人次: {result['late_count']}")
        msg_parts.append(f"\n输出: {os.path.basename(result['output_path'])}")
        QMessageBox.information(self, "完成", "考勤汇总表已生成！\n\n" + "\n".join(msg_parts))

    def on_error(self, msg):
        self.progress_bar.setValue(0)
        self.status_label.setText("出错！")
        self.status_label.setStyleSheet("color: #CC0000; font-weight: bold;")
        self.process_btn.setEnabled(True)
        self.select_btn.setEnabled(True)
        self.shift_btn.setEnabled(True)
        self.log(f"[考勤] 错误: {msg}")
        QMessageBox.critical(self, "错误", f"处理失败:\n{msg}")

    def open_output_folder(self):
        if self.last_output_dir and os.path.isdir(self.last_output_dir):
            os.startfile(self.last_output_dir)


# ══════════════════════════════════════════════════════════════════════════════
#  标签页 2: 宿舍分摊
# ══════════════════════════════════════════════════════════════════════════════

class DormitoryTab(QWidget):
    def __init__(self, log_writer):
        super().__init__()
        self.log = log_writer
        self.worker = None
        self.last_output_dir = None
        self.current_file = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 12, 20, 12)
        layout.setSpacing(8)

        # 文件选择
        section1 = QLabel("① 宿舍费用 Excel 文件")
        section1.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        section1.setStyleSheet("color: #333;")
        layout.addWidget(section1)

        self.drop_zone = DropZone("📂", "将宿舍费用 XLSX 文件拖拽到此处\n或点击选择文件")
        self.drop_zone.file_dropped.connect(self.on_file_selected)
        layout.addWidget(self.drop_zone)

        self.file_label = QLabel("未选择文件")
        self.file_label.setFont(QFont("Microsoft YaHei", 8))
        self.file_label.setStyleSheet("color: #999; padding: 1px 4px;")
        self.file_label.setWordWrap(True)
        layout.addWidget(self.file_label)

        self.select_btn = QPushButton("  选择宿舍费用文件...")
        self.select_btn.setFont(QFont("Microsoft YaHei", 9))
        self.select_btn.setStyleSheet("""
            QPushButton {
                background-color: #4472C4; color: white; border: none;
                border-radius: 5px; padding: 6px 14px;
            }
            QPushButton:hover { background-color: #365A9E; }
        """)
        self.select_btn.clicked.connect(self.browse_file)
        layout.addWidget(self.select_btn)

        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: #E0E0E0;")
        layout.addWidget(sep)

        # 计费月份
        section2 = QLabel("② 计费月份")
        section2.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        section2.setStyleSheet("color: #333;")
        layout.addWidget(section2)

        month_layout = QHBoxLayout()
        month_layout.setSpacing(10)

        now = datetime.now()
        month_layout.addWidget(QLabel("年份:"))
        self.year_combo = QComboBox()
        self.year_combo.addItems([str(y) for y in range(2020, 2031)])
        self.year_combo.setCurrentText(str(now.year))
        self.year_combo.setMinimumWidth(80)
        month_layout.addWidget(self.year_combo)

        month_layout.addWidget(QLabel("月份:"))
        self.month_combo = QComboBox()
        self.month_combo.addItems([str(m) for m in range(1, 13)])
        self.month_combo.setCurrentText(str(now.month))
        self.month_combo.setMinimumWidth(60)
        month_layout.addWidget(self.month_combo)

        month_layout.addStretch()
        layout.addLayout(month_layout)

        # 选项
        self.treat_prev_check = QCheckBox("将上月最后一天的截止日期视为仍在住（适用于批量填充场景）")
        self.treat_prev_check.setFont(QFont("Microsoft YaHei", 9))
        layout.addWidget(self.treat_prev_check)

        # 按钮
        btn_layout = QHBoxLayout()
        self.process_btn = QPushButton("  开始分摊计算")
        self.process_btn.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        self.process_btn.setEnabled(False)
        self.process_btn.setMinimumHeight(36)
        self.process_btn.setStyleSheet("""
            QPushButton {
                background-color: #2B8C3C; color: white; border: none;
                border-radius: 6px; padding: 8px 32px;
            }
            QPushButton:hover { background-color: #237030; }
            QPushButton:disabled { background-color: #C0C0C0; color: #F0F0F0; }
        """)
        self.process_btn.clicked.connect(self.start_processing)
        btn_layout.addWidget(self.process_btn)

        self.open_btn = QPushButton("  打开输出文件夹")
        self.open_btn.setFont(QFont("Microsoft YaHei", 9))
        self.open_btn.setEnabled(False)
        self.open_btn.setStyleSheet("""
            QPushButton {
                background-color: #E8ECF2; color: #333; border: 1px solid #C0C0C0;
                border-radius: 5px; padding: 6px 14px;
            }
            QPushButton:hover { background-color: #D4DAE4; }
            QPushButton:disabled { color: #C0C0C0; }
        """)
        self.open_btn.clicked.connect(self.open_output_folder)
        btn_layout.addWidget(self.open_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximum(100)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #D0D0D0; border-radius: 5px; text-align: center;
                background-color: #F0F0F0; height: 20px;
                font-family: "Microsoft YaHei"; font-size: 9px;
            }
            QProgressBar::chunk {
                background-color: #4472C4; border-radius: 4px;
            }
        """)
        layout.addWidget(self.progress_bar)

        self.status_label = QLabel("就绪")
        self.status_label.setFont(QFont("Microsoft YaHei", 9))
        self.status_label.setStyleSheet("color: #888;")
        layout.addWidget(self.status_label)
        layout.addStretch()

    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择宿舍费用文件", os.path.expanduser("~"),
            "Excel 文件 (*.xlsx *.xls);;所有文件 (*.*)")
        if path:
            self.on_file_selected(path)

    def on_file_selected(self, path):
        if path == "__browse__":
            self.browse_file(); return
        self.current_file = path
        self.file_label.setText(f"已选择: {os.path.basename(path)}")
        self.file_label.setStyleSheet("color: #2B579A; padding: 1px 4px;")
        self.process_btn.setEnabled(True)

        y, m = parse_billing_month_from_filename(path)
        if y and m:
            self.year_combo.setCurrentText(str(y))
            self.month_combo.setCurrentText(str(m))
            self.status_label.setText(f"已自动识别计费月份: {y}-{m:02d}")

        self.log(f"[宿舍] 文件: {path}")

    def start_processing(self):
        if not self.current_file:
            return
        try:
            y = int(self.year_combo.currentText())
            m = int(self.month_combo.currentText())
        except ValueError:
            QMessageBox.critical(self, "错误", "请输入有效的计费年份和月份")
            return

        treat_flag = self.treat_prev_check.isChecked()
        self.last_output_dir = os.path.dirname(self.current_file) or os.getcwd()

        self.process_btn.setEnabled(False)
        self.select_btn.setEnabled(False)
        self.open_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_label.setText("处理中...")
        self.status_label.setStyleSheet("color: #2B579A; font-weight: bold;")
        self.log("=" * 40)
        self.log(f"[宿舍] 开始分摊计算, 计费月份: {y}-{m:02d}")

        self.worker = DormitoryWorker(self.current_file, y, m, treat_flag)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_progress(self, current, total, message):
        self.progress_bar.setValue(current)
        self.status_label.setText(message)
        if message:
            self.log(f"[宿舍] {message}")

    def on_finished(self, result):
        self.progress_bar.setValue(100)
        self.status_label.setText("处理完成！")
        self.status_label.setStyleSheet("color: #2B8C3C; font-weight: bold;")
        self.process_btn.setEnabled(True)
        self.select_btn.setEnabled(True)
        self.open_btn.setEnabled(True)

        self.log("-" * 40)
        self.log("[宿舍] 分摊计算完成！")
        self.log(f"  计费月份: {result['billing_period']}")
        self.log(f"  涉及住客: {result['person_count']} 人")
        self.log(f"  涉及房间: {result['room_count']} 间")
        self.log(f"  总应付金额: {result['total_amount']:.2f} 元")
        self.log(f"  输出: {os.path.basename(result['output_path'])}")

        summary = (
            f"计费月份: {result['billing_period']}\n"
            f"涉及住客: {result['person_count']} 人\n"
            f"涉及房间: {result['room_count']} 间\n"
            f"总应付金额: {result['total_amount']:.2f} 元\n\n"
            f"结果已保存至:\n{result['output_path']}"
        )
        if result.get('errors'):
            summary += f"\n\n警告 ({len(result['errors'])} 条):\n"
            for e in result['errors'][:5]:
                summary += f"  - {e}\n"
            if len(result['errors']) > 5:
                summary += f"  ... 及其他 {len(result['errors']) - 5} 条"
        QMessageBox.information(self, "完成", summary)

    def on_error(self, msg):
        self.progress_bar.setValue(0)
        self.status_label.setText("出错！")
        self.status_label.setStyleSheet("color: #CC0000; font-weight: bold;")
        self.process_btn.setEnabled(True)
        self.select_btn.setEnabled(True)
        self.log(f"[宿舍] 错误: {msg}")
        QMessageBox.critical(self, "错误", f"处理失败:\n{msg}")

    def open_output_folder(self):
        if self.last_output_dir and os.path.isdir(self.last_output_dir):
            os.startfile(self.last_output_dir)


# ══════════════════════════════════════════════════════════════════════════════
#  主窗口
# ══════════════════════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("考勤与宿舍管理工具集")
        self.setMinimumSize(600, 680)
        self.resize(620, 720)

        palette = self.palette()
        palette.setColor(QPalette.Window, QColor("#FFFFFF"))
        self.setPalette(palette)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(0)

        # 标题
        title = QLabel("考勤与宿舍管理工具集")
        title.setFont(QFont("Microsoft YaHei", 14, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #2B579A; padding: 4px 0;")
        layout.addWidget(title)

        # 标签页
        self.tabs = QTabWidget()
        self.tabs.setFont(QFont("Microsoft YaHei", 10))
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #D0D0D0; border-radius: 4px;
                background-color: #FFFFFF;
            }
            QTabBar::tab {
                padding: 8px 24px; border: 1px solid #D0D0D0;
                border-bottom: none; border-top-left-radius: 5px;
                border-top-right-radius: 5px; margin-right: 2px;
                background-color: #F0F2F5;
            }
            QTabBar::tab:selected {
                background-color: #FFFFFF; font-weight: bold; color: #2B579A;
            }
            QTabBar::tab:hover:!selected {
                background-color: #E8ECF2;
            }
        """)
        layout.addWidget(self.tabs)

        # 日志区域
        log_group = QGroupBox("处理日志")
        log_group.setFont(QFont("Microsoft YaHei", 9))
        log_layout = QVBoxLayout(log_group)
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setMaximumHeight(130)
        self.log_output.setFont(QFont("Consolas", 8))
        self.log_output.setStyleSheet("""
            QTextEdit {
                background-color: #FAFBFC; border: 1px solid #D0D0D0;
                border-radius: 4px; padding: 4px;
            }
        """)
        log_layout.addWidget(self.log_output)
        layout.addWidget(log_group)

        # 添加标签页
        self.attendance_tab = AttendanceTab(self._log)
        self.dormitory_tab = DormitoryTab(self._log)
        self.tabs.addTab(self.attendance_tab, "📋 考勤汇总")
        self.tabs.addTab(self.dormitory_tab, "🏠 宿舍分摊")

        self.setStyleSheet("QMainWindow { background-color: #FFFFFF; }")

    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_output.append(f"[{ts}] {msg}")
        scrollbar = self.log_output.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("AttendanceSuite")
    app.setStyle("Fusion")

    font = app.font()
    font.setFamily("Microsoft YaHei")
    font.setPointSize(9)
    app.setFont(font)

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
