# HR 管理工具集 — Agent 复刻提示词（v2.0）

> **目标:** 让任何主流 AI Agent（Claude Code、Cursor、GitHub Copilot、Windsurf 等）能够一键复刻本项目的全部源码并打包为 EXE。
> **版本:** v2.0（三模块），v1.0（双模块）见仓库 

---

## 项目简介

一个基于 PyQt5 的桌面工具，包含三个功能模块：

| 模块 | 功能 | 引入版本 |
|------|------|---------|
| 考勤汇总 | 读取考勤打卡 XLSX → 生成每日汇总表 + 迟到判定 + 乐捐金额明细 | v1.0 |
| 宿舍分摊 | 读取宿舍费用 XLSX → 按人按天分摊水电费 + 住宿费（50元/天） | v1.0 |
| 出入库流水 | 读取出入库流水 XLSX → 按部门/存货清洗分类汇总 | v2.0 |

---

## 技术栈

- **Python 3.x**
- **PyQt5** — GUI 界面（标签页切换）
- **openpyxl** — Excel 读写与样式
- **pandas** — 数据处理
- **PyInstaller** — 打包为单文件 EXE

---

## 复刻步骤

请按照以下步骤操作，**严格按文件名和内容创建**：



---

### 步骤 1: 创建 

```text
openpyxl>=3.1.0
PyQt5>=5.15.0
pandas>=1.3.0
pyinstaller>=6.0.0

```

---

### 步骤 2: 创建 

这是考勤处理的核心逻辑模块，负责读取打卡 Excel、匹配班次表、判定迟到、生成汇总表和迟到明细表（含乐捐金额）。

```python
# -*- coding: utf-8 -*-
import math
import os
from collections import defaultdict
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


# ====== Configurable settings ======
LATE_RULES = {
    "白": {"check_time": "08:00", "time_type": "first"},
    "晚": {"check_time": "20:00", "time_type": "last"},
}
LATE_MINUTES_FLOOR = True
SHOW_LATE_MINUTES_IN_CELL = True

# ====== Style constants ======
LATE_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
LATE_FONT = Font(name="Microsoft YaHei", size=9, color="9C0006")


def calc_late_minutes(actual_time, threshold_str):
    th = datetime.strptime(threshold_str, "%H:%M").time()
    if actual_time <= th:
        return 0
    dummy_date = datetime.min.date()
    dt_actual = datetime.combine(dummy_date, actual_time)
    dt_th = datetime.combine(dummy_date, th)
    delta_seconds = (dt_actual - dt_th).total_seconds()
    return int(delta_seconds // 60)


def get_donation(minutes):
    if 1 <= minutes <= 10:
        return (15, f"迟到{minutes}分钟")
    elif 11 <= minutes <= 30:
        return (30, f"迟到{minutes}分钟")
    elif 31 <= minutes <= 59:
        return (100, f"迟到{minutes}分钟")
    elif minutes >= 60:
        return (300, f"迟到{minutes}分钟")
    return (0, "")


class AttendanceProcessor:

    def __init__(self, exclude_names=None, progress_callback=None):
        self.exclude_names = set(exclude_names or [])
        self.progress_callback = progress_callback

    def _report_progress(self, current, total, message=""):
        if self.progress_callback:
            self.progress_callback(current, total, message)

    def parse_shift_schedule(self, shift_file_path):
        wb = openpyxl.load_workbook(shift_file_path, data_only=True)
        ws = wb.active
        date_map = {}
        for col in range(7, ws.max_column + 1):
            val = ws.cell(2, col).value
            if val is None:
                break
            if isinstance(val, (int, float)):
                d = datetime(1899, 12, 30) + timedelta(days=int(val))
                date_map[col] = d.date()
        shift_data = {}
        for row in range(4, ws.max_row + 1, 6):
            emp_name = ws.cell(row, 3).value
            if not emp_name:
                continue
            emp_name = str(emp_name).strip()
            for col, date in date_map.items():
                shift_val = ws.cell(row, col).value
                if shift_val:
                    shift_type = str(shift_val).strip()
                    if shift_type in LATE_RULES:
                        shift_data[(emp_name, date)] = shift_type
        wb.close()
        return shift_data

    @staticmethod
    def _check_late(punch_times, shift_type):
        if shift_type not in LATE_RULES:
            return (False, 0, None)
        rule = LATE_RULES[shift_type]
        time_type = rule["time_type"]
        actual = punch_times[0] if time_type == "first" else punch_times[1]
        minutes = calc_late_minutes(actual, rule["check_time"])
        if minutes >= 1:
            return (True, minutes, time_type)
        return (False, 0, time_type)

    def process(self, input_path, output_path, shift_file_path=None):
        self._report_progress(0, 100, "正在读取打卡文件...")
        wb = openpyxl.load_workbook(input_path)
        source_sheet_name = wb.sheetnames[0]
        for name in wb.sheetnames:
            if "考勤" in name or "打卡" in name:
                source_sheet_name = name
                break
        ws = wb[source_sheet_name]
        max_row = ws.max_row
        self._report_progress(5, 100,
            "找到工作表 {n}, 共 {c} 条记录".format(n=source_sheet_name, c=max_row - 1))

        shift_data = {}
        if shift_file_path and os.path.isfile(shift_file_path):
            self._report_progress(7, 100, "正在读取班次表...")
            shift_data = self.parse_shift_schedule(shift_file_path)
            self._report_progress(10, 100,
                "班次表: {n} 条记录".format(n=len(shift_data)))

        emp_data = defaultdict(lambda: defaultdict(list))
        all_dates = set()
        excluded_count = 0
        skipped_invalid = 0

        for row_idx in range(2, max_row + 1):
            if row_idx % 500 == 0:
                pct = 10 + int((row_idx - 2) / (max_row - 1) * 30)
                self._report_progress(pct, 100,
                    "Row {r}/{t}...".format(r=row_idx, t=max_row))
            date_val = ws.cell(row_idx, 1).value
            time_val = ws.cell(row_idx, 2).value
            name_val = ws.cell(row_idx, 3).value
            if not name_val:
                skipped_invalid += 1
                continue
            emp_name = str(name_val).strip()
            if emp_name in self.exclude_names:
                excluded_count += 1
                continue
            if not date_val:
                skipped_invalid += 1
                continue
            if isinstance(date_val, datetime):
                parsed_date = date_val.date()
            elif isinstance(date_val, str):
                try:
                    parsed_date = datetime.strptime(date_val.strip(), "%Y/%m/%d").date()
                except ValueError:
                    skipped_invalid += 1
                    continue
            else:
                skipped_invalid += 1
                continue
            if isinstance(time_val, datetime):
                parsed_time = time_val.time()
            elif isinstance(time_val, str):
                try:
                    parsed_time = datetime.strptime(time_val.strip(), "%H:%M:%S").time()
                except ValueError:
                    skipped_invalid += 1
                    continue
            else:
                skipped_invalid += 1
                continue
            all_dates.add(parsed_date)
            if not emp_data[emp_name][parsed_date]:
                emp_data[emp_name][parsed_date] = [parsed_time, parsed_time]
            else:
                if parsed_time < emp_data[emp_name][parsed_date][0]:
                    emp_data[emp_name][parsed_date][0] = parsed_time
                if parsed_time > emp_data[emp_name][parsed_date][1]:
                    emp_data[emp_name][parsed_date][1] = parsed_time

        self._report_progress(40, 100, "数据读取完成, 正在生成汇总表...")
        sorted_dates = sorted(all_dates)
        self._report_progress(50, 100,
            "{e} 名员工, {d} 个日期".format(e=len(emp_data), d=len(sorted_dates)))

        for sn in ["汇总考勤表", "迟到明细"]:
            if sn in wb.sheetnames:
                del wb[sn]
        ws_out = wb.create_sheet("汇总考勤表")

        hdr_font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
        hdr_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        count_align = Alignment(horizontal="center", vertical="center")
        name_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
        border = Border(
            left=Side(style="thin", color="B4C6E7"),
            right=Side(style="thin", color="B4C6E7"),
            top=Side(style="thin", color="B4C6E7"),
            bottom=Side(style="thin", color="B4C6E7"),
        )
        even_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
        normal_font = Font(name="Microsoft YaHei", size=9)
        name_font = Font(name="Microsoft YaHei", size=10)

        total_cols = 1 + len(sorted_dates) + 1
        late_count_col = total_cols

        cell = ws_out.cell(1, 1, "姓名")
        cell.font = hdr_font; cell.fill = hdr_fill; cell.alignment = hdr_align; cell.border = border
        for ci, d in enumerate(sorted_dates, start=2):
            cell = ws_out.cell(1, ci, d.day)
            cell.font = hdr_font; cell.fill = hdr_fill; cell.alignment = hdr_align; cell.border = border
        cell = ws_out.cell(1, late_count_col, "迟到总次数")
        cell.font = hdr_font; cell.fill = hdr_fill; cell.alignment = hdr_align; cell.border = border

        sorted_employees = sorted(emp_data.keys())
        total = len(sorted_employees)
        late_count = 0
        late_records = []
        late_days_per_emp = defaultdict(set)

        for ri, emp in enumerate(sorted_employees, start=2):
            if ri % 10 == 0:
                pct = 50 + int((ri - 2) / total * 35)
                self._report_progress(pct, 100, "正在写入 {e}...".format(e=emp))
            nc = ws_out.cell(ri, 1, emp)
            nc.alignment = name_align; nc.border = border; nc.font = name_font
            if ri % 2 == 0:
                nc.fill = even_fill
            for ci, d in enumerate(sorted_dates, start=2):
                cell = ws_out.cell(ri, ci)
                cell.border = border
                cell.alignment = cell_align
                cell.font = normal_font
                is_late = False
                if d in emp_data[emp]:
                    times = emp_data[emp][d]
                    t0_str = times[0].strftime("%H:%M:%S")
                    t1_str = times[1].strftime("%H:%M:%S")
                    shift_key = (emp, d)
                    if shift_key in shift_data:
                        shift_type = shift_data[shift_key]
                        late_result = self._check_late(times, shift_type)
                        is_late, late_mins, time_type = late_result
                        if is_late and SHOW_LATE_MINUTES_IN_CELL:
                            if time_type == "first":
                                cell.value = "{t0}(迟到{m}分钟)\n{t1}".format(
                                    t0=t0_str, t1=t1_str, m=late_mins)
                            else:
                                cell.value = "{t0}\n{t1}(迟到{m}分钟)".format(
                                    t0=t0_str, t1=t1_str, m=late_mins)
                        else:
                            cell.value = "{t0}\n{t1}".format(t0=t0_str, t1=t1_str)
                        if is_late:
                            cell.fill = LATE_FILL
                            cell.font = LATE_FONT
                            late_count += 1
                            late_days_per_emp[emp].add(d)
                            late_records.append({
                                "employee": emp,
                                "date": d,
                                "shift_type": shift_type,
                                "threshold": LATE_RULES[shift_type]["check_time"],
                                "actual_time": times[0] if time_type == "first" else times[1],
                                "late_minutes": late_mins,
                            })
                    else:
                        cell.value = "{t0}\n{t1}".format(t0=t0_str, t1=t1_str)
                else:
                    cell.value = ""
                if not is_late and ri % 2 == 0:
                    cell.fill = even_fill
            tc = ws_out.cell(ri, late_count_col, len(late_days_per_emp.get(emp, set())))
            tc.border = border
            tc.alignment = count_align
            tc.font = normal_font
            if ri % 2 == 0:
                tc.fill = even_fill

        self._report_progress(85, 100, "正在调整格式...")
        ws_out.column_dimensions["A"].width = 10
        for ci in range(2, len(sorted_dates) + 2):
            ws_out.column_dimensions[get_column_letter(ci)].width = 13
        ws_out.column_dimensions[get_column_letter(late_count_col)].width = 11
        ws_out.freeze_panes = "B2"
        ws_out.auto_filter.ref = "A1:{col}{row}".format(
            col=get_column_letter(total_cols),
            row=len(sorted_employees) + 1)

        legend_row = len(sorted_employees) + 3
        if late_count > 0:
            lc = ws_out.cell(legend_row, 1, "图例:")
            lc.font = Font(name="Microsoft YaHei", size=9, bold=True)
            ll = ws_out.cell(legend_row, 2,
                "浅红底 = 迟到 (括号内为分钟数)")
            ll.fill = LATE_FILL
            ll.font = LATE_FONT
            ll.alignment = Alignment(horizontal="left", vertical="center")
            ws_out.merge_cells(start_row=legend_row, start_column=2,
                end_row=legend_row, end_column=min(6, total_cols))

        # Late detail sheet
        self._report_progress(90, 100, "正在生成迟到明细...")
        if "迟到明细" in wb.sheetnames:
            del wb["迟到明细"]
        ws_detail = wb.create_sheet("迟到明细")

        dh_headers = ["员工姓名", "日期", "班次",
                       "应到阈值", "实际打卡时间", "迟到分钟数",
                       "乐捐金额", "乐捐原因", "备注"]
        dh_font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
        dh_fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        dh_align = Alignment(horizontal="center", vertical="center")
        for ci, h in enumerate(dh_headers, start=1):
            cell = ws_detail.cell(1, ci, h)
            cell.font = dh_font; cell.fill = dh_fill; cell.alignment = dh_align; cell.border = border

        late_records.sort(key=lambda x: (x["date"], x["employee"]))
        for ri, rec in enumerate(late_records, start=2):
            donation_amount, donation_reason = get_donation(rec["late_minutes"])
            vals = [
                rec["employee"],
                rec["date"].strftime("%Y-%m-%d"),
                rec["shift_type"],
                rec["threshold"],
                rec["actual_time"].strftime("%H:%M:%S"),
                rec["late_minutes"],
                donation_amount,
                donation_reason,
                ""
            ]
            for ci, v in enumerate(vals, start=1):
                cell = ws_detail.cell(ri, ci, v)
                cell.border = border
                cell.alignment = dh_align
                cell.font = Font(name="Microsoft YaHei", size=9)
                if ri % 2 == 0:
                    cell.fill = even_fill

        detail_widths = [10, 12, 6, 10, 12, 10, 10, 10, 10]
        for ci, w in enumerate(detail_widths, start=1):
            ws_detail.column_dimensions[get_column_letter(ci)].width = w
        ws_detail.freeze_panes = "A2"
        ws_detail.auto_filter.ref = "A1:I{row}".format(row=len(late_records) + 1)

        self._report_progress(98, 100, "正在保存...")
        wb.save(output_path)
        wb.close()
        self._report_progress(100, 100, "处理完成!")

        return {
            "total_records": max_row - 1,
            "valid_records": max_row - 1 - excluded_count - skipped_invalid,
            "excluded_count": excluded_count,
            "skipped_invalid": skipped_invalid,
            "employee_count": len(sorted_employees),
            "date_count": len(sorted_dates),
            "date_range": "{s} ~ {e}".format(s=sorted_dates[0], e=sorted_dates[-1]) if sorted_dates else "N/A",
            "output_path": output_path,
            "late_count": late_count,
            "shift_count": len(shift_data),
            "detail_records": len(late_records),
        }

```

---

### 步骤 3: 创建 

这是宿舍水电住宿费分摊的核心逻辑模块，负责读取宿舍费用表、解析入住人员、按天分摊计算。

```python
# -*- coding: utf-8 -*-
"""宿舍水电住宿费分摊 — 纯逻辑模块（无 UI 依赖）"""

import os
import re
from datetime import datetime, timedelta
from calendar import monthrange

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


def parse_billing_month_from_filename(filepath):
    basename = os.path.splitext(os.path.basename(filepath))[0]
    m = re.search(r'(\d{4})-(\d{1,2})$', basename)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


def _safe_str(val):
    if val is None:
        return ""
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    return str(val).strip()


def _calc_effective_days(checkin, checkout, month_start, month_end):
    if checkin is None or pd.isna(checkin):
        return 0
    if checkout is None or pd.isna(checkout):
        checkout = month_end
    if isinstance(checkin, datetime):
        checkin = checkin.date()
    if isinstance(checkout, datetime):
        checkout = checkout.date()
    if isinstance(month_start, datetime):
        month_start = month_start.date()
    if isinstance(month_end, datetime):
        month_end = month_end.date()
    effective_start = max(checkin, month_start)
    effective_end = min(checkout, month_end)
    days = (effective_end - effective_start).days + 1
    return max(0, days)


def _forward_fill_room_info(df):
    for col in ["房间号码", "房型", "入住人数"]:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: _safe_str(x) if pd.notna(x) else None)
            df[col] = df[col].ffill()
    return df


def _read_utility_sheet(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb["水电总计表"]
    rows = []
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
        rows.append(row)
    if len(rows) < 2:
        wb.close()
        return {}, "room"
    header = str(rows[0][0]) if rows[0][0] else ""
    mode = "room"
    if "月份" in header or (isinstance(rows[1][0], str) and re.match(r"\d{4}-\d{2}", str(rows[1][0]))):
        mode = "month"
    result = {}
    for i, row in enumerate(rows[1:], start=2):
        key = row[0]
        val = row[1] if len(row) > 1 else 0
        if key is None or (isinstance(key, str) and key.strip() in ("", "总计", "合计")):
            continue
        try:
            val = float(val) if val is not None else 0.0
        except (ValueError, TypeError):
            continue
        if mode == "month":
            result[str(key).strip()] = val
        else:
            result[_safe_str(key)] = val
    wb.close()
    return result, mode


def _read_housing_sheet(filepath):
    df = pd.read_excel(filepath, sheet_name="房屋表")
    col_aliases = {
        "房间号码": ["房间号码", "单元房号", "房号"],
        "房型": ["房型"],
        "入住人数": ["入住人数"],
        "入住人员": ["入住人员", "姓名"],
        "住宿计费时间": ["住宿计费时间", "入住日期", "计费开始"],
        "截止日期": ["截止日期", "离开日期", "计费结束"],
    }
    rename_map = {}
    for target, candidates in col_aliases.items():
        for c in candidates:
            if c in df.columns:
                rename_map[c] = target
                break
    df = df.rename(columns=rename_map)
    df = _forward_fill_room_info(df)
    if "入住人员" in df.columns:
        df = df[df["入住人员"].notna()].copy()
        df = df[df["入住人员"].astype(str).str.strip() != ""].copy()
    for col in ["住宿计费时间", "截止日期"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    if "入住人数" in df.columns:
        df["入住人数"] = pd.to_numeric(df["入住人数"], errors="coerce")
    return df


def _write_excel(output_path, df_person, df_room):
    wb = openpyxl.Workbook()
    hf = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    hfill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    halign = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cf = Font(name="微软雅黑", size=10)
    calign = Alignment(horizontal="center", vertical="center")
    tb = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    hb = Border(
        left=Side(style="thin", color="2F5496"),
        right=Side(style="thin", color="2F5496"),
        top=Side(style="thin", color="2F5496"),
        bottom=Side(style="thin", color="2F5496"),
    )

    def write_sheet(ws, df, money_cols):
        for ci, cn in enumerate(df.columns, 1):
            c = ws.cell(row=1, column=ci, value=cn)
            c.font = hf; c.fill = hfill; c.alignment = halign; c.border = hb
        for ri, (_, row) in enumerate(df.iterrows(), 2):
            for ci, cn in enumerate(df.columns, 1):
                v = row[cn]
                if pd.isna(v):
                    v = ""
                c = ws.cell(row=ri, column=ci, value=v)
                c.font = cf; c.alignment = calign; c.border = tb
                if cn in money_cols and isinstance(v, (int, float)):
                    c.number_format = "#,##0.00"
        for ci in range(1, len(df.columns) + 1):
            ml = 0
            for r in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=ci, max_col=ci):
                for c in r:
                    if c.value:
                        s = str(c.value)
                        ml = max(ml, sum(2 if ord(ch) > 127 else 1 for ch in s))
            ws.column_dimensions[get_column_letter(ci)].width = min(ml + 4, 36)
        ws.freeze_panes = "A2"

    ws1 = wb.active; ws1.title = "按个人明细"
    write_sheet(ws1, df_person, ["水电分摊金额(元)", "住宿费(元)", "总应付金额(元)"])
    ws2 = wb.create_sheet("按房间汇总")
    write_sheet(ws2, df_room, ["房间总应付金额(元)"])
    wb.save(output_path)


def process_allocation(filepath, billing_year, billing_month,
                       treat_end_of_prev_month_as_living=False,
                       progress_callback=None):
    utility_data, utility_mode = _read_utility_sheet(filepath)
    housing_df = _read_housing_sheet(filepath)

    month_start = datetime(billing_year, billing_month, 1)
    _, last_day = monthrange(billing_year, billing_month)
    month_end = datetime(billing_year, billing_month, last_day)
    total_days_in_month = last_day
    prev_month_last_day = month_start - timedelta(days=1)

    if treat_end_of_prev_month_as_living:
        count_fixed = 0
        for idx in housing_df.index:
            d = housing_df.at[idx, "截止日期"]
            if pd.notna(d):
                d_date = d.date() if isinstance(d, datetime) else d
                if d_date == prev_month_last_day.date():
                    housing_df.at[idx, "截止日期"] = pd.NaT
                    count_fixed += 1

    results_person = []
    results_room = {}
    all_rooms = set()
    if utility_mode == "room":
        all_rooms.update(utility_data.keys())
    all_rooms.update(housing_df["房间号码"].dropna().apply(_safe_str).unique())
    errors = []

    for room in sorted(all_rooms):
        room_residents = housing_df[housing_df["房间号码"].apply(_safe_str) == room]
        if len(room_residents) == 0:
            if utility_mode == "room" and room in utility_data:
                errors.append(f"房间 {room}: 有账单但无入住人员")
            continue
        room_utility = utility_data.get(room, 0.0) if utility_mode == "room" else 0.0
        resident_days = []
        for idx, row in room_residents.iterrows():
            try:
                days = _calc_effective_days(
                    row["住宿计费时间"], row["截止日期"], month_start, month_end)
            except Exception:
                continue
            resident_days.append((row, days))
        total_person_days = sum(d for _, d in resident_days)
        for row_data, days in resident_days:
            person_name = str(row_data["入住人员"]) if pd.notna(row_data["入住人员"]) else ""
            if total_person_days > 0 and room_utility > 0:
                utility_share = room_utility * (days / total_person_days)
            else:
                utility_share = 0.0
            accommodation = 50.0 * (days / total_days_in_month) if days > 0 else 0.0
            results_person.append({
                "月份": f"{billing_year}-{billing_month:02d}",
                "房间号码": room,
                "入住人员": person_name,
                "有效居住天数": days,
                "水电分摊金额(元)": round(utility_share, 2),
                "住宿费(元)": round(accommodation, 2),
                "总应付金额(元)": round(utility_share + accommodation, 2),
            })
            mk = f"{billing_year}-{billing_month:02d}"
            if room not in results_room:
                results_room[room] = {}
            if mk not in results_room[room]:
                results_room[room][mk] = 0.0
            results_room[room][mk] += round(utility_share + accommodation, 2)

    df_person = pd.DataFrame(results_person)
    if len(df_person) == 0:
        df_person = pd.DataFrame(columns=[
            "月份", "房间号码", "入住人员", "有效居住天数",
            "水电分摊金额(元)", "住宿费(元)", "总应付金额(元)"
        ])

    room_summary = []
    for rn, months in sorted(results_room.items()):
        for m, total in months.items():
            room_summary.append({
                "月份": m, "房间号码": rn,
                "房间总应付金额(元)": round(total, 2)
            })
    df_room = pd.DataFrame(room_summary)
    if len(df_room) == 0:
        df_room = pd.DataFrame(columns=["月份", "房间号码", "房间总应付金额(元)"])

    if progress_callback:
        progress_callback(98, 100, "正在保存...")

    base = os.path.splitext(os.path.basename(filepath))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = os.path.dirname(filepath) or os.getcwd()
    output_path = os.path.join(output_dir, f"分摊结果_{base}_{timestamp}.xlsx")
    _write_excel(output_path, df_person, df_room)

    if progress_callback:
        progress_callback(100, 100, "处理完成!")

    return {
        "output_path": output_path,
        "person_count": len(df_person),
        "room_count": len(df_room),
        "total_amount": df_room["房间总应付金额(元)"].sum()
        if len(df_room) > 0 and "房间总应付金额(元)" in df_room.columns else 0,
        "errors": errors,
        "billing_period": f"{billing_year}-{billing_month:02d}",
    }

```

---

### 步骤 4: 创建   ⭐ v2.0 新增

这是出入库流水清洗模块，负责筛选材料出库单、按经手人映射部门、按存货分组汇总、生成多工作表 Excel 报告。

**核心功能：**
- 自动跳过 Excel 前 8 行元数据（标题 + 双行表头）
- 筛选 （排除进货单、销货单）
- 经手人 → 部门映射：彭建森/罗德浮→VCP, 曹鹏→蚀刻, 邹柳珍→曝光, 刘志军→品质, 其他→其他
- 按存货 + 部门排序，分组计算总数量 / 平均单价 / 总出库金额
- 输出 6 工作表 Excel（汇总 + VCP + 曝光 + 蚀刻 + 品质 + 其他）
- 汇总表列：存货 | 计量单位 | 总数量 | 平均单价 | 出库金额合计
- 分部门表右侧 N 列起附加存货汇总

```python
# -*- coding: utf-8 -*-
"""出入库流水清洗处理器 — 纯逻辑模块，不含 GUI"""

import pandas as pd
import os
import re
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# ============== 配置常量 ==============
DATA_START_ROW = 8  # 0索引：第9行开始是数据

DEPARTMENT_MAP = {
    '彭建森': 'VCP',
    '罗德浮': 'VCP',
    '曹鹏': '蚀刻',
    '邹柳珍': '曝光',
    '刘志军': '品质',
}

KEEP_ORDER_TYPE = '材料出库单'

OUTPUT_COLUMNS = [
    '单据日期', '单据类型', '经手人', '部门', '仓库',
    '存货', '规格型号', '计量单位',
    '出库数量', '出库单价', '出库金额',
]

NORMALIZED_COLUMNS = [
    '单据日期', '单据类型', '部门', '经手人', '往来单位',
    '仓库', '存货', '规格型号', '计量单位',
    '入库数量', '入库单价', '入库金额',
    '出库数量', '出库单价', '出库金额',
]

DEPT_SHEET_ORDER = ['VCP', '曝光', '蚀刻', '品质', '其他']


def _safe_to_numeric(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        return 0
    s = re.sub(r'[¥$￥,\s]', '', s)
    try:
        return float(s)
    except ValueError:
        return 0


def _is_data_row(row_values):
    if row_values is None or len(row_values) < 2:
        return False
    col1 = row_values[1] if len(row_values) > 1 else None
    col2 = row_values[2] if len(row_values) > 2 else None
    if col1 is None or (isinstance(col1, float) and pd.isna(col1)):
        return False
    s = str(col1).strip()
    if not s or '合计' in s or '制表' in s or '打印' in s or '第' in s:
        return False
    if re.match(r'\d{4}[-/.]\d{1,2}[-/.]\d{1,2}', s):
        return True
    if col2 is not None:
        col2_s = str(col2).strip()
        if col2_s in [KEEP_ORDER_TYPE, '进货单', '销货单']:
            return True
    return False


def process_inventory(input_path, output_path=None, progress_callback=None):
    """主处理函数

    Args:
        input_path: 输入 Excel 文件路径
        output_path: 输出 Excel 文件路径（可选，默认自动生成）
        progress_callback: 进度回调 (current_step, total_steps, message)

    Returns:
        dict: 处理结果摘要
    """
    def prog(current, total, message):
        if progress_callback:
            progress_callback(current, total, message)

    if not output_path:
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(
            os.path.dirname(input_path) or '.',
            f'出入库流水_清洗后_{ts}.xlsx'
        )

    # ---- Step 1: 读取 ----
    prog(1, 6, "读取文件...")
    df_raw = pd.read_excel(input_path, header=None, dtype=object)
    if df_raw.empty:
        raise ValueError("文件为空，无法处理")

    data_rows = []
    for r in range(DATA_START_ROW, len(df_raw)):
        row_vals = list(df_raw.iloc[r])
        if _is_data_row(row_vals):
            data_rows.append(row_vals)

    data_rows_trimmed = [row[1:] for row in data_rows]
    df = pd.DataFrame(data_rows_trimmed, columns=NORMALIZED_COLUMNS)

    # ---- Step 2: 筛选 ----
    prog(2, 6, "筛选材料出库单...")
    if '单据类型' not in df.columns:
        raise KeyError(f"找不到'单据类型'列，可用列: {list(df.columns)}")
    before = len(df)
    df = df[df['单据类型'] == KEEP_ORDER_TYPE].copy()
    after = len(df)

    # ---- Step 3: 清洗数值 ----
    prog(3, 6, "清洗数值列...")
    for col in ['出库数量', '出库单价', '出库金额']:
        if col in df.columns:
            df[col] = df[col].apply(_safe_to_numeric)

    # ---- Step 4: 部门映射 ----
    prog(4, 6, "部门映射...")
    if '经手人' not in df.columns:
        raise KeyError(f"找不到'经手人'列，可用列: {list(df.columns)}")

    new_depts = []
    for _, row in df.iterrows():
        person = row['经手人']
        if pd.isna(person):
            person = ''
        person = str(person).strip()
        new_depts.append(DEPARTMENT_MAP.get(person, '其他'))
    df['部门'] = new_depts
    mapped = len([d for d in new_depts if d != '其他'])

    # ---- Step 5: 排序汇总 ----
    prog(5, 6, "排序与汇总...")
    df = df.sort_values(by=['存货', '部门'], ascending=[True, True]).reset_index(drop=True)
    summary = df.groupby('存货').agg(
        出库数量合计=('出库数量', 'sum'),
        出库金额合计=('出库金额', 'sum'),
        记录条数=('出库数量', 'count'),
    ).reset_index()
    summary = summary.sort_values(by='存货', ascending=True).reset_index(drop=True)
    summary['出库数量合计'] = summary['出库数量合计'].round(2)
    summary['出库金额合计'] = summary['出库金额合计'].round(2)

    # ---- Step 6: 生成 Excel ----
    prog(6, 6, "生成 Excel 报告...")
    _generate_excel(df, summary, output_path)

    # 部门统计
    dept_stats = df.groupby('部门').agg(
        记录数=('出库金额', 'count'),
        出库金额合计=('出库金额', 'sum'),
        出库数量合计=('出库数量', 'sum'),
    ).round(2)

    dept_info = {}
    for dn in DEPT_SHEET_ORDER:
        if dn in dept_stats.index:
            r = dept_stats.loc[dn]
            dept_info[dn] = {
                'count': int(r['记录数']),
                'qty': round(r['出库数量合计'], 2),
                'amount': round(r['出库金额合计'], 2),
            }
        else:
            dept_info[dn] = {'count': 0, 'qty': 0, 'amount': 0}

    return {
        'input_path': input_path,
        'output_path': output_path,
        'total_records': len(df),
        'inventory_count': len(summary),
        'total_qty': round(summary['出库数量合计'].sum(), 2),
        'total_amount': round(summary['出库金额合计'].sum(), 2),
        'mapped_count': mapped,
        'other_count': len(df) - mapped,
        'raw_count': before,
        'filtered_count': after,
        'dept_info': dept_info,
    }


def _generate_excel(df, summary, output_path):
    """生成多工作表 Excel 输出文件"""
    wb = Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font_white = Font(name='微软雅黑', bold=True, size=11, color='FFFFFF')
    data_font = Font(name='微软雅黑', size=10)
    total_font = Font(name='微软雅黑', bold=True, size=10)
    total_fill = PatternFill(start_color='D9E2F3', end_color='D9E2F3', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    center_align = Alignment(horizontal='center', vertical='center')
    left_align = Alignment(horizontal='left', vertical='center')
    right_align = Alignment(horizontal='right', vertical='center')

    def style_header(ws, num_cols):
        for col in range(1, num_cols + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = thin_border

    def style_data_cell(ws, row, col, align='center'):
        cell = ws.cell(row=row, column=col)
        cell.font = data_font
        cell.border = thin_border
        if align == 'right':
            cell.alignment = right_align
        elif align == 'left':
            cell.alignment = left_align
        else:
            cell.alignment = center_align

    def auto_width(ws, min_width=8, max_width=40):
        for col_cells in ws.columns:
            col_letter = get_column_letter(col_cells[0].column)
            max_len = 0
            for cell in col_cells:
                if cell.value:
                    val = str(cell.value)
                    char_len = sum(2 if '一' <= ch <= '鿿' or '　' <= ch <= '〿' or '＀' <= ch <= '￯' else 1 for ch in val)
                    max_len = max(max_len, char_len)
            ws.column_dimensions[col_letter].width = max(min_width, min(max_len + 4, max_width))

    # ---- 汇总表 ----
    ws_summary = wb.create_sheet('汇总')
    summary_cols = ['存货', '计量单位', '总数量', '平均单价', '出库金额合计']
    ws_summary.append(summary_cols)
    style_header(ws_summary, len(summary_cols))

    # 从原始数据提取每个存货的计量单位（取第一个非空值）
    unit_map = df.groupby('存货')['计量单位'].first().to_dict() if '计量单位' in df.columns else {}

    for _, row_data in summary.iterrows():
        qty = round(row_data['出库数量合计'], 2)
        amt = round(row_data['出库金额合计'], 2)
        avg_price = round(amt / qty, 2) if qty != 0 else 0
        unit = unit_map.get(row_data['存货'], '')
        ws_summary.append([row_data['存货'], unit, qty, avg_price, amt])

    total_row = len(summary) + 2
    total_qty = round(summary['出库数量合计'].sum(), 2)
    total_amount = round(summary['出库金额合计'].sum(), 2)
    total_avg = round(total_amount / total_qty, 2) if total_qty != 0 else 0
    ws_summary.append(['合计', '', total_qty, total_avg, total_amount])
    for col in range(1, len(summary_cols) + 1):
        cell = ws_summary.cell(row=total_row, column=col)
        cell.font = total_font
        cell.fill = total_fill
        cell.border = thin_border
        cell.alignment = center_align if col == 1 else right_align

    for r in range(2, total_row):
        style_data_cell(ws_summary, r, 1, 'left')
        style_data_cell(ws_summary, r, 2, 'center')
        style_data_cell(ws_summary, r, 3, 'right')
        style_data_cell(ws_summary, r, 4, 'right')
        style_data_cell(ws_summary, r, 5, 'right')
    auto_width(ws_summary)
    ws_summary.freeze_panes = 'A2'

    # ---- 分部门明细 ----
    available_cols = [c for c in OUTPUT_COLUMNS if c in df.columns]
    summary_header_col = 14  # N列
    dept_summary_cols = ['存货', '计量单位', '总数量', '平均单价', '总出库金额']

    for dept_name in DEPT_SHEET_ORDER:
        ws_dept = wb.create_sheet(dept_name)
        dept_df = df[df['部门'] == dept_name].copy()
        dept_df = dept_df.sort_values(by='存货', ascending=True)

        ws_dept.append(available_cols)
        style_header(ws_dept, len(available_cols))

        if len(dept_df) == 0:
            ws_dept.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(available_cols))
            ws_dept.cell(row=2, column=1, value='无数据')
            ws_dept.cell(row=2, column=1).font = Font(name='微软雅黑', size=10, color='999999')
            ws_dept.cell(row=2, column=1).alignment = center_align
            for ci, cn in enumerate(dept_summary_cols):
                c = summary_header_col + ci
                ws_dept.cell(row=1, column=c, value=cn)
                cell = ws_dept.cell(row=1, column=c)
                cell.font = header_font_white
                cell.fill = header_fill
                cell.alignment = center_align
                cell.border = thin_border
            ws_dept.merge_cells(start_row=2, start_column=summary_header_col,
                                end_row=2, end_column=summary_header_col + len(dept_summary_cols) - 1)
            ws_dept.cell(row=2, column=summary_header_col, value='无数据')
            ws_dept.cell(row=2, column=summary_header_col).font = Font(name='微软雅黑', size=10, color='999999')
            ws_dept.cell(row=2, column=summary_header_col).alignment = center_align
            auto_width(ws_dept)
            ws_dept.freeze_panes = 'A2'
            continue

        for r_idx, (_, row_data) in enumerate(dept_df.iterrows()):
            row_num = r_idx + 2
            for c_idx, col_name in enumerate(available_cols):
                val = row_data[col_name]
                if isinstance(val, float) and pd.isna(val):
                    val = ''
                ws_dept.cell(row=row_num, column=c_idx + 1, value=val)

        for r in range(2, len(dept_df) + 2):
            for c in range(1, len(available_cols) + 1):
                col_name = available_cols[c - 1]
                style_data_cell(ws_dept, r, c, 'right' if col_name in ['出库数量', '出库单价', '出库金额'] else 'center')

        # 底部汇总行
        summary_row = len(dept_df) + 2
        qty_sum = round(dept_df['出库数量'].sum(), 2)
        amt_sum = round(dept_df['出库金额'].sum(), 2)

        if '存货' in available_cols:
            ws_dept.cell(row=summary_row, column=available_cols.index('存货') + 1, value='汇总')
        else:
            ws_dept.cell(row=summary_row, column=1, value='汇总')
        if '出库数量' in available_cols:
            ws_dept.cell(row=summary_row, column=available_cols.index('出库数量') + 1, value=qty_sum)
        if '出库金额' in available_cols:
            ws_dept.cell(row=summary_row, column=available_cols.index('出库金额') + 1, value=amt_sum)

        for c in range(1, len(available_cols) + 1):
            cell = ws_dept.cell(row=summary_row, column=c)
            cell.font = total_font
            cell.fill = total_fill
            cell.border = thin_border
            cell.alignment = center_align

        # 右侧按存货汇总 (N列开始)
        for ci, cn in enumerate(dept_summary_cols):
            c = summary_header_col + ci
            ws_dept.cell(row=1, column=c, value=cn)
            cell = ws_dept.cell(row=1, column=c)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = thin_border

        dept_inv = dept_df.groupby('存货').agg(
            总数量=('出库数量', 'sum'),
            总出库金额=('出库金额', 'sum'),
            计量单位=('计量单位', 'first'),
        ).reset_index()
        dept_inv['平均单价'] = (dept_inv['总出库金额'] / dept_inv['总数量'].replace(0, None)).round(2)
        dept_inv['总数量'] = dept_inv['总数量'].round(2)
        dept_inv['总出库金额'] = dept_inv['总出库金额'].round(2)
        dept_inv = dept_inv.sort_values(by='存货', ascending=True)

        for ri, (_, inv_row) in enumerate(dept_inv.iterrows()):
            row_num = ri + 2
            vals = [inv_row['存货'], inv_row['计量单位'], inv_row['总数量'], inv_row['平均单价'], inv_row['总出库金额']]
            for ci, val in enumerate(vals):
                c = summary_header_col + ci
                ws_dept.cell(row=row_num, column=c, value=val)
                cell = ws_dept.cell(row=row_num, column=c)
                cell.font = data_font
                cell.border = thin_border
                if ci == 1:
                    cell.alignment = center_align
                elif ci > 0:
                    cell.alignment = right_align
                else:
                    cell.alignment = left_align

        inv_total_row = len(dept_inv) + 2
        inv_total_qty = round(dept_inv['总数量'].sum(), 2)
        inv_total_amt = round(dept_inv['总出库金额'].sum(), 2)
        inv_total_avg = round(inv_total_amt / inv_total_qty, 2) if inv_total_qty != 0 else 0
        inv_total_vals = ['合计', '', inv_total_qty, inv_total_avg, inv_total_amt]
        for ci, val in enumerate(inv_total_vals):
            c = summary_header_col + ci
            ws_dept.cell(row=inv_total_row, column=c, value=val)
            cell = ws_dept.cell(row=inv_total_row, column=c)
            cell.font = total_font
            cell.fill = total_fill
            cell.border = thin_border
            if ci == 0:
                cell.alignment = center_align
            elif ci == 1:
                cell.alignment = center_align
            else:
                cell.alignment = right_align

        auto_width(ws_dept)
        ws_dept.freeze_panes = 'A2'

    wb.save(output_path)


# 命令行入口
if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='出入库流水自动化清洗')
    parser.add_argument('--input', '-i', default='出入库流水.xlsx', help='输入Excel文件')
    parser.add_argument('--output', '-o', default=None, help='输出Excel文件')
    args = parser.parse_args()

    result = process_inventory(args.input, args.output)
    print("处理完成！")
    print(f"  总记录数: {result['total_records']}")
    print(f"  存货种类: {result['inventory_count']}")
    print(f"  出库金额合计: {result['total_amount']:,.2f}")
    print(f"  输出: {result['output_path']}")

```

---

### 步骤 5: 创建  — 三标签页主程序

包含：考勤汇总 + 宿舍分摊 + 出入库流水，新增  线程和  标签页。

```python
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
from inventory_processor import process_inventory

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

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("HR 管理工具集 — 考勤 · 宿舍 · 出入库")
        self.setMinimumSize(640, 700)
        self.resize(660, 740)

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
        self.tabs.addTab(self.attendance_tab, "📋 考勤汇总")
        self.tabs.addTab(self.dormitory_tab, "🏠 宿舍分摊")
        self.tabs.addTab(self.inventory_tab, "📦 出入库流水")

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

```

---

### 步骤 6: 创建辅助文件

**run.bat:**
```bat
@echo off
chcp 65001 >nul
echo ============================================
echo   考勤与宿舍管理工具集 - 启动
echo ============================================
echo.
python main.py
pause

```

**build.bat:**
```bat
@echo off
chcp 65001 >nul
echo ============================================
echo   考勤与宿舍管理工具集 - 打包为 EXE
echo ============================================
echo.

REM --- 1. Install dependencies ---
echo [1/3] 安装 Python 依赖...
pip install -r requirements.txt -q
if %errorlevel% neq 0 (
    echo 错误: 依赖安装失败。
    pause
    exit /b 1
)
echo 完成.
echo.

REM --- 2. Build with PyInstaller ---
echo [2/3] 使用 PyInstaller 打包...
python -m PyInstaller --onefile --windowed --name "AttendanceSuite" --add-data "attendance_processor.py;." --add-data "dormitory_processor.py;." main.py
if %errorlevel% neq 0 (
    echo 错误: PyInstaller 打包失败。
    pause
    exit /b 1
)
echo 完成.
echo.

REM --- 3. Show output ---
echo [3/3] 打包完成!
echo.
echo EXE 位置: %CD%\dist\AttendanceSuite.exe
echo.

start "" "%CD%\dist"

echo.
pause

```

**.gitignore:**
```gitignore
# Python
__pycache__/
*.py[cod]
*.egg-info/
dist/
build/
*.spec

# IDE
.vscode/
.idea/

# OS
.DS_Store
Thumbs.db

# Output
output/
*.xlsx
!sample*.xlsx

```

---

## 输入文件格式要求

### 考勤打卡文件 (XLSX)

| 日期 | 时间 | 姓名 |
|---|---|---|
| 2024/01/15 | 07:55:30 | 张三 |
| 2024/01/15 | 17:02:00 | 张三 |

### 班次表文件 (XLSX，可选)

- 第 2 行第 7 列起为日期
- 第 3 列(从第 4 行起)为员工姓名
- 单元格值为 白(白班 08:00) 或 晚(晚班 20:00)

### 宿舍费用文件 (XLSX)

- 工作表 水电总计表 — 房间号/月份 → 水电金额
- 工作表 房屋表 — 入住人员信息(房间号、入住人员、入住/截止日期、入住人数)

### 出入库流水文件 (XLSX) — v2.0 新增

- 第 1-2 行：标题"出入库流水"（自动跳过）
- 第 3-6 行：空行/备注/期间信息（自动跳过）
- 第 7-8 行：双行表头（自动合并为规范列名）
- 第 9 行起：数据行（材料出库单 / 进货单 / 销货单）
- 末尾：合计行（含"合计："字样，自动跳过）

| 部门映射规则 | |
|--------|-----|
| 彭建森 | VCP |
| 罗德浮 | VCP |
| 曹鹏 | 蚀刻 |
| 邹柳珍 | 曝光 |
| 刘志军 | 品质 |
| 其他 | 其他 |

---

## 乐捐规则（考勤模块）

| 迟到分钟 | 乐捐金额 |
|---------|---------|
| 1-10 分钟 | 15 元 |
| 11-30 分钟 | 30 元 |
| 31-59 分钟 | 100 元 |
| ≥60 分钟 | 300 元 |

---

## 复刻验证

复刻完成后，请执行以下检查：

1.  能正常启动 GUI
2. 三个标签页"📋 考勤汇总"、"🏠 宿舍分摊"、"📦 出入库流水"显示正常
3. 拖拽 Excel 文件到拖放区能正确识别
4. 出入库模块选文件后点击"开始清洗"，进度条正常更新
5. 输出 Excel 含 6 个工作表，汇总表 5 列，分部门表右侧有存货汇总

---

## Agent 复刻提示

> **直接复制以上所有代码块到对应文件中，按顺序执行步骤 1-6 即可完成复刻。**
>
> 推荐使用 Claude Code、Cursor 或 Windsurf 等 AI 编程助手，直接输入 "请根据 PROMPT.md 复刻这个项目" 即可自动创建所有文件。

---

*项目地址: https://github.com/yourstudendjd/HR-Admin-Toolkit*
*版本: v2.0（三模块）*
*v1.0 源码见 main_v1.py*
