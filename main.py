# -*- coding: utf-8 -*-
"""HR 管理工具集 — 统一 PyQt5 应用程序

包含四大功能模块:
  1. 考勤汇总 — 打卡记录汇总 + 迟到判定 + 乐捐金额
  2. 宿舍分摊 — 水电住宿费按人按天分摊
  3. 出入库流水 — 读取出入库流水 XLSX → 按部门/存货清洗分类汇总
  4. 数据清洗 — 客户对账单去重分类，按客户型号+加工类型+加工要求组合
"""

import os
import sys
from datetime import datetime
from calendar import monthrange

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QProgressBar, QTextEdit, QFileDialog,
    QMessageBox, QFrame, QCheckBox, QTabWidget, QComboBox, QLineEdit,
    QGridLayout, QGroupBox, QTableWidget, QTableWidgetItem,
    QAbstractItemView, QHeaderView
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QPalette, QColor, QDragEnterEvent, QDropEvent

from attendance_processor import AttendanceProcessor
from dormitory_processor import process_allocation, parse_billing_month_from_filename
from inventory_processor import process_inventory
from data_cleaner_processor import DataCleanerProcessor

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


class InventoryWorker(QThread):
    progress = pyqtSignal(int, int, str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, input_path, output_path):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path

    def run(self):
        try:
            result = process_inventory(
                self.input_path, self.output_path,
                progress_callback=lambda c, t, m: self.progress.emit(c, t, m)
            )
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


# ══════════════════════════════════════════════════════════════════════════════
#  标签页 3: 出入库流水清洗
# ══════════════════════════════════════════════════════════════════════════════

class InventoryTab(QWidget):
    def __init__(self, log_writer):
        super().__init__()
        self.log = log_writer
        self.worker = None
        self.last_output_path = None
        self.current_input = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 12, 20, 12)
        layout.setSpacing(8)

        section1 = QLabel("① 出入库流水 Excel 文件")
        section1.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        section1.setStyleSheet("color: #333;")
        layout.addWidget(section1)

        self.drop_zone = DropZone("📂", "将出入库流水 XLSX 文件拖拽到此处\n或点击选择文件")
        self.drop_zone.file_dropped.connect(self.on_file_selected)
        layout.addWidget(self.drop_zone)

        self.file_label = QLabel("未选择文件")
        self.file_label.setFont(QFont("Microsoft YaHei", 8))
        self.file_label.setStyleSheet("color: #999; padding: 1px 4px;")
        self.file_label.setWordWrap(True)
        layout.addWidget(self.file_label)

        self.select_btn = QPushButton("  选择出入库流水文件...")
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

        section2 = QLabel("② 输出设置")
        section2.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        section2.setStyleSheet("color: #333;")
        layout.addWidget(section2)

        out_layout = QHBoxLayout()
        out_layout.addWidget(QLabel("输出路径:"))
        self.output_edit = QLineEdit()
        self.output_edit.setReadOnly(True)
        self.output_edit.setPlaceholderText("自动生成（与输入文件同目录）")
        out_layout.addWidget(self.output_edit)
        self.output_browse_btn = QPushButton("选择...")
        self.output_browse_btn.clicked.connect(self.browse_output)
        out_layout.addWidget(self.output_browse_btn)
        layout.addLayout(out_layout)

        layout.addSpacing(6)

        # 按钮
        btn_layout = QHBoxLayout()
        self.process_btn = QPushButton("  开始清洗")
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

        self.open_btn = QPushButton("  打开输出文件")
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
        self.open_btn.clicked.connect(self.open_output)
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
            self, "选择出入库流水文件", os.path.expanduser("~"),
            "Excel 文件 (*.xlsx *.xls);;所有文件 (*.*)")
        if path:
            self.on_file_selected(path)

    def browse_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "保存输出文件", os.path.expanduser("~"),
            "Excel 文件 (*.xlsx);;所有文件 (*.*)")
        if path:
            self.output_edit.setText(path)

    def on_file_selected(self, path):
        if path == "__browse__":
            self.browse_file(); return
        self.current_input = path
        self.file_label.setText(f"已选择: {os.path.basename(path)}")
        self.file_label.setStyleSheet("color: #2B579A; padding: 1px 4px;")
        self.process_btn.setEnabled(True)
        self.log(f"[出入库] 文件: {path}")

    def start_processing(self):
        if not self.current_input:
            return
        input_path = self.current_input
        output_path = self.output_edit.text().strip() or None

        self.process_btn.setEnabled(False)
        self.select_btn.setEnabled(False)
        self.output_browse_btn.setEnabled(False)
        self.open_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_label.setText("处理中...")
        self.status_label.setStyleSheet("color: #2B579A; font-weight: bold;")
        self.log("=" * 40)
        self.log("[出入库] 开始清洗...")

        self.worker = InventoryWorker(input_path, output_path)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_progress(self, current, total, message):
        pct = int(current / total * 100) if total else 0
        self.progress_bar.setValue(pct)
        self.status_label.setText(message)
        if message:
            self.log(f"[出入库] {message}")

    def on_finished(self, result):
        self.progress_bar.setValue(100)
        self.status_label.setText("处理完成！")
        self.status_label.setStyleSheet("color: #2B8C3C; font-weight: bold;")
        self.process_btn.setEnabled(True)
        self.select_btn.setEnabled(True)
        self.output_browse_btn.setEnabled(True)
        self.open_btn.setEnabled(True)
        self.last_output_path = result['output_path']
        self.output_edit.setText(result['output_path'])

        self.log("-" * 40)
        self.log("[出入库] 清洗完成！")
        self.log(f"  筛选前: {result['raw_count']} 条 → 材料出库单: {result['filtered_count']} 条")
        self.log(f"  部门映射: {result['mapped_count']} 条匹配, {result['other_count']} 条归入'其他'")
        self.log(f"  存货种类: {result['inventory_count']} 种")
        self.log(f"  出库金额合计: {result['total_amount']:,.2f} 元")
        dept_info = result.get('dept_info', {})
        for dn in ['VCP', '曝光', '蚀刻', '品质', '其他']:
            if dn in dept_info:
                d = dept_info[dn]
                if d['count'] > 0:
                    self.log(f"  [{dn}] {d['count']} 条, 金额 {d['amount']:,.2f}")
                else:
                    self.log(f"  [{dn}] 无数据")
        self.log(f"  输出: {os.path.basename(result['output_path'])}")

        QMessageBox.information(
            self, "完成",
            f"出入库流水清洗完成！\n\n"
            f"材料出库单: {result['filtered_count']} 条\n"
            f"存货种类: {result['inventory_count']} 种\n"
            f"出库金额合计: {result['total_amount']:,.2f} 元\n\n"
            f"输出文件:\n{result['output_path']}"
        )

    def on_error(self, msg):
        self.progress_bar.setValue(0)
        self.status_label.setText("出错！")
        self.status_label.setStyleSheet("color: #CC0000; font-weight: bold;")
        self.process_btn.setEnabled(True)
        self.select_btn.setEnabled(True)
        self.output_browse_btn.setEnabled(True)
        self.log(f"[出入库] 错误: {msg}")
        QMessageBox.critical(self, "错误", f"处理失败:\n{msg}")

    def open_output(self):
        if self.last_output_path and os.path.exists(self.last_output_path):
            os.startfile(self.last_output_path)


# ══════════════════════════════════════════════════════════════════════════════
#  工作线程: 数据清洗
# ══════════════════════════════════════════════════════════════════════════════

#  工作线程: 数据清洗（支持批量多文件 → 一个工作簿多工作表）
# ══════════════════════════════════════════════════════════════════════════════

class DataCleanerWorker(QThread):
    progress = pyqtSignal(int, int, str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, file_tasks, output_path):
        super().__init__()
        self.file_tasks = file_tasks
        self.output_path = output_path

    def run(self):
        try:
            processor = DataCleanerProcessor(
                progress_callback=lambda c, t, m: self.progress.emit(c, t, m)
            )
            sheet_names, results, warn = processor.clean_and_merge(
                self.file_tasks, self.output_path
            )
            result = {
                'sheet_names': sheet_names,
                'results': results,
                'output_path': self.output_path,
                'warning': warn or '',
                'total_files': len(self.file_tasks),
                'success_count': sum(1 for r in results if r.get('状态') == '成功'),
            }
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


#  标签页 4: 数据清洗与分类
# ══════════════════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════════════════
#  标签页 4: 数据清洗与分类（批量多文件）
# ══════════════════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════════════════
#  标签页 4: 对账数据清洗（批量多文件，每文件独立模糊匹配列名）
# ══════════════════════════════════════════════════════════════════════════════

class DataCleanerTab(QWidget):
    """对账数据清洗标签页 — 支持批量添加多个 Excel 文件，
    每个文件独立模糊匹配列名，分别清洗后合并为一个工作簿（多工作表）。"""

    def __init__(self, log_writer):
        super().__init__()
        self.log = log_writer
        self.worker = None
        self.last_output_path = None
        self.file_tasks = []          # 每个元素: {path, header_row, mapping, strategy, columns, row_count, status}
        self.current_edit_row = -1    # 当前正在编辑映射的文件行
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 12, 20, 12)
        layout.setSpacing(8)

        # === ① 文件列表 ===
        section1 = QLabel("① 对账文件列表（可添加多个 Excel 文件）")
        section1.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        section1.setStyleSheet("color: #333;")
        layout.addWidget(section1)

        self.file_table = QTableWidget()
        self.file_table.setColumnCount(5)
        self.file_table.setHorizontalHeaderLabels(["文件名", "表头行", "行数", "匹配状态", ""])
        self.file_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.file_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Fixed)
        self.file_table.setColumnWidth(1, 70)
        self.file_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Fixed)
        self.file_table.setColumnWidth(2, 70)
        self.file_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Fixed)
        self.file_table.setColumnWidth(3, 100)
        self.file_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Fixed)
        self.file_table.setColumnWidth(4, 80)
        self.file_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.file_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.file_table.setMinimumHeight(150)
        self.file_table.setFont(QFont("Microsoft YaHei", 9))
        self.file_table.setStyleSheet("""
            QTableWidget {
                background-color: #FAFBFC; border: 1px solid #D0D0D0;
                border-radius: 4px; gridline-color: #E8E8E8;
            }
            QHeaderView::section {
                background-color: #F0F2F5; border: 1px solid #D0D0D0;
                padding: 4px; font-weight: bold; color: #333;
            }
        """)
        self.file_table.itemSelectionChanged.connect(self._on_file_selected)
        layout.addWidget(self.file_table)

        # 添加/删除按钮行
        file_btn_layout = QHBoxLayout()
        self.add_btn = QPushButton("  + 添加文件")
        self.add_btn.setFont(QFont("Microsoft YaHei", 9))
        self.add_btn.setStyleSheet("""
            QPushButton {
                background-color: #4472C4; color: white; border: none;
                border-radius: 5px; padding: 6px 14px;
            }
            QPushButton:hover { background-color: #365A9E; }
        """)
        self.add_btn.clicked.connect(self.add_files)
        file_btn_layout.addWidget(self.add_btn)

        self.remove_btn = QPushButton("  - 移除选中")
        self.remove_btn.setFont(QFont("Microsoft YaHei", 9))
        self.remove_btn.setEnabled(False)
        self.remove_btn.setStyleSheet("""
            QPushButton {
                background-color: #E8ECF2; color: #333; border: 1px solid #C0C0C0;
                border-radius: 5px; padding: 6px 14px;
            }
            QPushButton:hover { background-color: #D4DAE4; }
            QPushButton:disabled { color: #C0C0C0; }
        """)
        self.remove_btn.clicked.connect(self.remove_selected)
        file_btn_layout.addWidget(self.remove_btn)

        self.clear_btn = QPushButton("  清空列表")
        self.clear_btn.setFont(QFont("Microsoft YaHei", 9))
        self.clear_btn.setEnabled(False)
        self.clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #E8ECF2; color: #333; border: 1px solid #C0C0C0;
                border-radius: 5px; padding: 6px 14px;
            }
            QPushButton:hover { background-color: #D4DAE4; }
            QPushButton:disabled { color: #C0C0C0; }
        """)
        self.clear_btn.clicked.connect(self.clear_all)
        file_btn_layout.addWidget(self.clear_btn)
        file_btn_layout.addStretch()
        layout.addLayout(file_btn_layout)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: #E0E0E0;")
        layout.addWidget(sep)

        # === ② 选中文件的列映射（每文件独立） ===
        section2 = QLabel("② 选中文件的列映射（每个文件独立匹配）")
        section2.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        section2.setStyleSheet("color: #333;")
        layout.addWidget(section2)

        self.current_file_label = QLabel("请先添加文件，然后在列表中选中一个文件进行列映射")
        self.current_file_label.setFont(QFont("Microsoft YaHei", 8))
        self.current_file_label.setStyleSheet("color: #999;")
        self.current_file_label.setWordWrap(True)
        layout.addWidget(self.current_file_label)

        mapping_group = QGroupBox()
        mapping_group.setStyleSheet("QGroupBox { border: 1px solid #D0D0D0; border-radius: 5px; margin-top: 6px; padding-top: 8px; }")
        mg_layout = QGridLayout(mapping_group)
        mg_layout.setContentsMargins(12, 8, 12, 8)
        mg_layout.setSpacing(6)

        self.mapping_combos = {}
        field_labels = [
            ('客户型号', '对账单中表示客户产品型号的列'),
            ('单价', '对账单中表示单价的列'),
            ('加工类型', '对账单中表示加工工艺/类型的列'),
            ('规格', '对账单中表示产品规格的列（不参与去重组合）'),
            ('加工要求', '对账单中表示加工要求的列'),
        ]
        for idx, (field, tip) in enumerate(field_labels):
            lbl = QLabel(f"{field}:")
            lbl.setFont(QFont("Microsoft YaHei", 9))
            lbl.setToolTip(tip)
            mg_layout.addWidget(lbl, idx, 0)
            combo = QComboBox()
            combo.setMinimumWidth(200)
            combo.addItem("（未匹配）")
            combo.currentTextChanged.connect(self._on_mapping_combo_changed)
            mg_layout.addWidget(combo, idx, 1)
            self.mapping_combos[field] = combo

        layout.addWidget(mapping_group)

        # === ③ 全局设置 ===
        sep3 = QFrame()
        sep3.setFrameShape(QFrame.HLine)
        sep3.setStyleSheet("color: #E0E0E0;")
        layout.addWidget(sep3)

        section3 = QLabel("③ 全局设置")
        section3.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
        section3.setStyleSheet("color: #333;")
        layout.addWidget(section3)

        # 表头行 + 单价策略 同行
        global_layout = QHBoxLayout()
        global_layout.addWidget(QLabel("表头行号:"))
        self.header_combo = QComboBox()
        self.header_combo.addItems([str(i) for i in range(25)])
        self.header_combo.setCurrentText("7")
        self.header_combo.setMinimumWidth(60)
        self.header_combo.currentTextChanged.connect(self._update_all_header_rows)
        global_layout.addWidget(self.header_combo)

        global_layout.addSpacing(20)
        global_layout.addWidget(QLabel("单价冲突策略:"))
        self.strategy_combo = QComboBox()
        self.strategy_combo.addItems([
            '取首次出现的值',
            '取平均值',
            '取最大值',
            '提示冲突报错',
        ])
        self.strategy_combo.setCurrentText('取首次出现的值')
        self.strategy_combo.setMinimumWidth(160)
        global_layout.addWidget(self.strategy_combo)
        global_layout.addStretch()
        layout.addLayout(global_layout)

        # === ④ 操作按钮 ===
        sep4 = QFrame()
        sep4.setFrameShape(QFrame.HLine)
        sep4.setStyleSheet("color: #E0E0E0;")
        layout.addWidget(sep4)

        btn_layout = QHBoxLayout()
        self.process_btn = QPushButton("  开始批量对账清洗")
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

        self.status_label = QLabel("请添加对账文件以开始")
        self.status_label.setFont(QFont("Microsoft YaHei", 9))
        self.status_label.setStyleSheet("color: #888;")
        layout.addWidget(self.status_label)
        layout.addStretch()

    # ------------------------------------------------------------------
    # 文件列表管理
    # ------------------------------------------------------------------
    def add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择对账文件", os.path.expanduser("~"),
            "Excel 文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        if not paths:
            return

        new_tasks = []
        for path in paths:
            if any(t['path'] == path for t in self.file_tasks):
                self.log(f"[对账清洗] 跳过（已存在）: {os.path.basename(path)}")
                continue
            task = {
                'path': path,
                'header_row': int(self.header_combo.currentText()),
                'mapping': {},
                'strategy': 'first',
                'columns': [],
                'row_count': 0,
                'status': '待处理',
            }
            new_tasks.append(task)

        if not new_tasks:
            return

        # 逐个加载文件信息并独立模糊匹配列名
        for task in new_tasks:
            self._load_and_match_file(task)

        self.file_tasks.extend(new_tasks)
        self._refresh_file_table()

        # 自动选中第一个新添加的文件
        if len(new_tasks) == 1 or self.current_edit_row < 0:
            self.file_table.selectRow(len(self.file_tasks) - len(new_tasks))
            # 触发选中事件
            self._on_file_selected()

        self._update_process_btn()
        self.log(f"[对账清洗] 已添加 {len(new_tasks)} 个文件，共 {len(self.file_tasks)} 个")

    def _load_and_match_file(self, task):
        """加载文件信息：先自动检测表头行，再模糊匹配列名（每文件独立）"""
        processor = DataCleanerProcessor()

        # 第一步：自动检测表头行（每文件独立检测，不依赖全局设置）
        detected_header = processor.auto_detect_header_row(task['path'])
        task['header_row'] = detected_header

        # 第二步：用检测到的表头行加载文件
        columns, preview, total_rows, error = processor.load_file(task['path'], detected_header)

        if error:
            task['status'] = '加载失败'
            task['columns'] = []
            task['row_count'] = 0
            self.log(f"[对账清洗] 加载失败: {os.path.basename(task['path'])} - {error}")
            return

        task['columns'] = columns
        task['row_count'] = total_rows

        # 第三步：用模糊匹配自动识别列映射（每文件独立匹配）
        mapping = DataCleanerProcessor.fuzzy_match_columns(columns)
        task['mapping'] = mapping

        matched_count = len(mapping)
        if matched_count == 5:
            task['status'] = '已匹配'
        elif matched_count > 0:
            task['status'] = f'部分匹配({matched_count}/5)'
        else:
            task['status'] = '未匹配'

        task['strategy'] = DataCleanerProcessor.PRICE_STRATEGIES.get(
            self.strategy_combo.currentText(), 'first'
        )

        self.log(f"[对账清洗] {os.path.basename(task['path'])}: "
                 f"表头行={detected_header}, {total_rows}行, {len(columns)}列, 匹配{matched_count}/5个字段")

    def remove_selected(self):
        row = self.file_table.currentRow()
        if row < 0 or row >= len(self.file_tasks):
            return
        removed = self.file_tasks.pop(row)
        self._refresh_file_table()
        if self.current_edit_row == row:
            self.current_edit_row = -1
            self._clear_mapping_combos()
        elif self.current_edit_row > row:
            self.current_edit_row -= 1
        self._update_process_btn()
        self.log(f"[对账清洗] 已移除: {os.path.basename(removed['path'])}")

    def clear_all(self):
        if not self.file_tasks:
            return
        self.file_tasks.clear()
        self._refresh_file_table()
        self.current_edit_row = -1
        self._clear_mapping_combos()
        self._update_process_btn()
        self.log("[对账清洗] 已清空文件列表")

    def _on_file_selected(self):
        """文件列表选中变更 → 加载该文件的列映射到编辑区"""
        row = self.file_table.currentRow()
        self.remove_btn.setEnabled(row >= 0)
        self.clear_btn.setEnabled(len(self.file_tasks) > 0)

        if row < 0 or row >= len(self.file_tasks):
            self.current_edit_row = -1
            self._clear_mapping_combos()
            self.current_file_label.setText("请先添加文件，然后在列表中选中一个文件进行列映射")
            return

        self.current_edit_row = row
        task = self.file_tasks[row]
        self.current_file_label.setText(
            f"当前编辑: {os.path.basename(task['path'])}  "
            f"（{task['row_count']} 行, {len(task['columns'])} 列）"
        )
        self._load_mapping_to_combos(task)

    def _clear_mapping_combos(self):
        """清空列映射下拉框"""
        for combo in self.mapping_combos.values():
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("（未匹配）")
            combo.blockSignals(False)
        self.current_file_label.setText("请先添加文件，然后在列表中选中一个文件进行列映射")

    def _load_mapping_to_combos(self, task):
        """将指定任务的列映射加载到下拉框"""
        columns = task['columns']
        mapping = task['mapping']
        col_options = ['（未匹配）'] + columns

        for field, combo in self.mapping_combos.items():
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(col_options)
            if field in mapping:
                combo.setCurrentText(mapping[field])
            else:
                combo.setCurrentText('（未匹配）')
            combo.blockSignals(False)

    def _on_mapping_combo_changed(self):
        """列映射下拉框变更 → 保存到当前选中文件的任务"""
        if self.current_edit_row < 0 or self.current_edit_row >= len(self.file_tasks):
            return

        task = self.file_tasks[self.current_edit_row]
        mapping = self._get_current_mapping()
        task['mapping'] = mapping

        # 更新状态
        matched = len(mapping)
        if matched == 5:
            task['status'] = '已匹配'
        elif matched > 0:
            task['status'] = f'部分匹配({matched}/5)'
        else:
            task['status'] = '未匹配'

        self._refresh_file_table()
        # 重新选中当前行
        self.file_table.selectRow(self.current_edit_row)
        self._update_process_btn()

    def _get_current_mapping(self):
        """从下拉框读取当前显示的列映射"""
        mapping = {}
        for field, combo in self.mapping_combos.items():
            val = combo.currentText().strip()
            if val and val != '（未匹配）':
                mapping[field] = val
        return mapping

    def _update_process_btn(self):
        """检查：有文件 且 每个文件都匹配了5个字段"""
        if not self.file_tasks:
            self.process_btn.setEnabled(False)
            return
        all_ready = all(len(t.get('mapping', {})) == 5 for t in self.file_tasks)
        self.process_btn.setEnabled(all_ready)
        if all_ready:
            self.status_label.setText(f"就绪 — {len(self.file_tasks)} 个文件均已匹配，可开始清洗")
            self.status_label.setStyleSheet("color: #2B8C3C;")
        else:
            ready = sum(1 for t in self.file_tasks if len(t.get('mapping', {})) == 5)
            self.status_label.setText(f"待处理 — {ready}/{len(self.file_tasks)} 个文件已匹配")
            self.status_label.setStyleSheet("color: #888;")

    # ------------------------------------------------------------------
    # 全局设置
    # ------------------------------------------------------------------
    def _update_all_header_rows(self):
        val = int(self.header_combo.currentText())
        for task in self.file_tasks:
            task['header_row'] = val
        self._refresh_file_table()

    # ------------------------------------------------------------------
    # 文件列表表格
    # ------------------------------------------------------------------
    def _refresh_file_table(self):
        self.file_table.setRowCount(len(self.file_tasks))
        for i, task in enumerate(self.file_tasks):
            # 文件名
            name_item = QTableWidgetItem(os.path.basename(task['path']))
            name_item.setToolTip(task['path'])
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            self.file_table.setItem(i, 0, name_item)

            # 表头行
            h_item = QTableWidgetItem(str(task['header_row']))
            h_item.setFlags(h_item.flags() & ~Qt.ItemIsEditable)
            h_item.setTextAlignment(Qt.AlignCenter)
            self.file_table.setItem(i, 1, h_item)

            # 行数
            r_item = QTableWidgetItem(str(task['row_count']) if task['row_count'] else '-')
            r_item.setFlags(r_item.flags() & ~Qt.ItemIsEditable)
            r_item.setTextAlignment(Qt.AlignCenter)
            self.file_table.setItem(i, 2, r_item)

            # 匹配状态
            matched = len(task.get('mapping', {}))
            status_text = f"{matched}/5"
            s_item = QTableWidgetItem(status_text)
            s_item.setFlags(s_item.flags() & ~Qt.ItemIsEditable)
            s_item.setTextAlignment(Qt.AlignCenter)
            if matched == 5:
                s_item.setForeground(QColor("#2B8C3C"))
            elif matched > 0:
                s_item.setForeground(QColor("#E89700"))
            else:
                s_item.setForeground(QColor("#CC0000"))
            self.file_table.setItem(i, 3, s_item)

        self.clear_btn.setEnabled(len(self.file_tasks) > 0)

    # ------------------------------------------------------------------
    # 批量处理
    # ------------------------------------------------------------------
    def start_processing(self):
        if not self.file_tasks:
            QMessageBox.warning(self, "提示", "请先添加文件")
            return

        # 检查每个文件的映射
        incomplete = [t for t in self.file_tasks if len(t.get('mapping', {})) != 5]
        if incomplete:
            names = ', '.join(os.path.basename(t['path']) for t in incomplete)
            QMessageBox.warning(self, "提示", f"以下文件列映射不完整（需5个字段）:\n{names}")
            return

        price_strategy = DataCleanerProcessor.PRICE_STRATEGIES.get(
            self.strategy_combo.currentText(), 'first'
        )

        # 构建文件任务列表（每个文件带自己的 mapping）
        file_tasks = []
        for task in self.file_tasks:
            file_tasks.append({
                'path': task['path'],
                'header_row': task['header_row'],
                'mapping': task['mapping'],
                'strategy': price_strategy,
            })

        input_dir = os.path.dirname(self.file_tasks[0]['path']) or os.getcwd()
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(input_dir, f"批量对账清洗结果_{ts}.xlsx")
        self.last_output_path = input_dir

        # 锁定界面
        self.process_btn.setEnabled(False)
        self.add_btn.setEnabled(False)
        self.remove_btn.setEnabled(False)
        self.clear_btn.setEnabled(False)
        self.header_combo.setEnabled(False)
        self.strategy_combo.setEnabled(False)
        self.file_table.setEnabled(False)
        for combo in self.mapping_combos.values():
            combo.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_label.setText("批量对账清洗中...")
        self.status_label.setStyleSheet("color: #2B579A; font-weight: bold;")

        self.log("=" * 40)
        self.log(f"[对账清洗] 开始处理 {len(file_tasks)} 个文件...")
        self.log(f"[对账清洗] 输出: {output_path}")

        self.worker = DataCleanerWorker(file_tasks, output_path)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_progress(self, current, total, message):
        self.progress_bar.setValue(current)
        self.status_label.setText(message)
        if message:
            self.log(f"[对账清洗] {message}")

    def on_finished(self, result):
        self.progress_bar.setValue(100)
        self.status_label.setText("对账清洗完成！")
        self.status_label.setStyleSheet("color: #2B8C3C; font-weight: bold;")
        self._unlock_ui()

        self.log("-" * 40)
        self.log(f"[对账清洗] 完成！成功 {result['success_count']}/{result['total_files']} 个文件")
        self.log(f"  工作表: {', '.join(result['sheet_names'])}")
        for r in result['results']:
            self.log(f"  - {r.get('文件名', '')}: {r.get('状态', '')} "
                     f"({r.get('原始行数', 0)} -> {r.get('清洗后行数', 0)} 行, "
                     f"合并 {r.get('合并/删除行数', 0)} 行)")
        self.log(f"  输出: {os.path.basename(result['output_path'])}")

        msg_lines = [
            f"处理文件: {result['success_count']}/{result['total_files']} 个成功",
            f"工作表: {', '.join(result['sheet_names'])}",
            f"\n各文件明细:",
        ]
        for r in result['results']:
            msg_lines.append(
                f"  {r.get('文件名', '')}: {r.get('状态', '')} "
                f"({r.get('原始行数', 0)} -> {r.get('清洗后行数', 0)} 行)"
            )
        if result.get('warning'):
            msg_lines.append(f"\n警告: {result['warning']}")
        msg_lines.append(f"\n输出文件:\n{result['output_path']}")
        QMessageBox.information(self, "对账清洗完成", "\n".join(msg_lines))

    def on_error(self, msg):
        self.progress_bar.setValue(0)
        self.status_label.setText("出错！")
        self.status_label.setStyleSheet("color: #CC0000; font-weight: bold;")
        self._unlock_ui()
        self.log(f"[对账清洗] 错误: {msg}")
        QMessageBox.critical(self, "错误", f"对账清洗失败:\n{msg}")

    def _unlock_ui(self):
        self.process_btn.setEnabled(True)
        self.add_btn.setEnabled(True)
        self.remove_btn.setEnabled(len(self.file_tasks) > 0)
        self.clear_btn.setEnabled(len(self.file_tasks) > 0)
        self.header_combo.setEnabled(True)
        self.strategy_combo.setEnabled(True)
        self.file_table.setEnabled(True)
        for combo in self.mapping_combos.values():
            combo.setEnabled(True)
        if self.current_edit_row >= 0:
            self._load_mapping_to_combos(self.file_tasks[self.current_edit_row])
        self._update_process_btn()

    def open_output_folder(self):
        if self.last_output_path and os.path.isdir(self.last_output_path):
            os.startfile(self.last_output_path)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("HR 管理工具集 — 考勤 · 宿舍 · 出入库 · 数据清洗")
        self.setMinimumSize(640, 750)
        self.resize(660, 800)

        palette = self.palette()
        palette.setColor(QPalette.Window, QColor("#FFFFFF"))
        self.setPalette(palette)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(0)

        # 标题
        title = QLabel("HR 管理工具集")
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
        self.inventory_tab = InventoryTab(self._log)
        self.cleaner_tab = DataCleanerTab(self._log)
        self.tabs.addTab(self.attendance_tab, "📋 考勤汇总")
        self.tabs.addTab(self.dormitory_tab, "🏠 宿舍分摊")
        self.tabs.addTab(self.inventory_tab, "📦 出入库流水")
        self.tabs.addTab(self.cleaner_tab, "🧹 对账数据清洗")

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
