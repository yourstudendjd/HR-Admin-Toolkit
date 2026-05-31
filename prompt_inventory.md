# 出入库流水自动化清洗与分类汇总工具 — Agent 复刻提示词

> **目标:** 让任何主流 AI Agent（Claude Code、Cursor、GitHub Copilot、Windsurf 等）能够一键复刻本项目的全部源码（GUI 版 + 命令行版）。

---

## 项目简介

一个基于 Python + tkinter 的桌面工具，专门用于自动化处理企业出入库流水 Excel 文件。

**核心功能:** 读取原始出入库流水 → 筛选材料出库单 → 按经手人自动映射部门 → 按存货分组汇总 → 生成多工作表 Excel 报告。

---

## 技术栈

- **Python 3.x**  
- **tkinter** — GUI 界面（内置标准库，无需额外安装）  
- **pandas** — 数据处理  
- **openpyxl** — Excel 读写与样式（字体、颜色、边框、合并单元格）  

---

## 复刻步骤

请按照以下步骤创建文件，**严格按文件名和内容创建**：

```
inventory-cleaner/
├── process_inventory.py    # 主程序（GUI + 命令行双模式）
├── README.md               # 项目说明
├── prompt.md               # 本文件（Agent 复刻提示词）
└── .gitignore
```

---

### 步骤 1: 输入文件格式分析

输入 Excel 文件命名为 `出入库流水.xlsx`（或类似名称），结构如下：

| 行号 | 内容 |
|------|------|
| 第 1 行 | 标题（空） |
| 第 2 行 | "出入库流水" |
| 第 3-5 行 | 空行/备注 |
| 第 6 行 | 期间和单据日期信息 |
| 第 7 行 | **第一行表头**（父级分类）：无/单据日期/单据类型/部门/经手人/往来单位/仓库/存货/规格型号/计量单位/本期入库/无/无/本期出库/无/无 |
| 第 8 行 | **第二行表头**（子级明细）：无/无/无/无/无/无/无/无/无/无/入库数量/入库单价/入库金额/出库数量/出库单价/出库金额 |
| 第 9 行起 | 数据行 |
| 末尾 | 合计行（第1列含"合计："字样），之后为制表人、打印日期、页码等元数据 |

**关键点:**
- 第 1 列（Col 0）始终为空占位列，构建 DataFrame 时需要剔除
- 数据行涵盖三种单据类型：`材料出库单`、`进货单`、`销货单`
- 第 3 列（部门）是原始表的部门列，会被 `经手人` 映射结果覆盖
- 第 4 列（经手人）是员工姓名，用于映射到部门

---

### 步骤 2: 创建 `process_inventory.py`

以下为完整源码，分为 **命令行模式** 和 **GUI 模式**，同时支持两种运行方式。

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
出入库流水自动化清洗与分类汇总工具 — GUI版
"""

import pandas as pd
import sys
import os
import re
import threading
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# ============== 配置常量 ==============
DATA_START_ROW = 8

DEPARTMENT_MAP = {
    '彭建森': 'VCP',
    '罗德浮': 'VCP',
    '曹鹏':   '蚀刻',
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
```

#### 2.1 工具函数

```python
def safe_to_numeric(value):
    """尝试将值转为数值，处理货币符号、逗号、空格等"""
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


def is_data_row(row_values):
    """判断一行是否为有效数据行（非合计行、非元数据行）"""
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
```

#### 2.2 数据读取

```python
def read_excel_data(file_path):
    """读取原始Excel文件，跳过元数据行和合计行，返回清洗后的DataFrame"""
    df_raw = pd.read_excel(file_path, header=None, dtype=object)
    if df_raw.empty:
        raise ValueError("文件为空，无法处理")

    data_rows = []
    for r in range(DATA_START_ROW, len(df_raw)):
        row_vals = list(df_raw.iloc[r])
        if is_data_row(row_vals):
            data_rows.append(row_vals)

    # 剔除非数据行的逻辑（合计行、制表人行、页脚行）
    # is_data_row 函数内部完成过滤

    data_rows_trimmed = [row[1:] for row in data_rows]
    df = pd.DataFrame(data_rows_trimmed, columns=NORMALIZED_COLUMNS)
    return df
```

#### 2.3 筛选与映射

```python
def filter_material_orders(df):
    """只保留材料出库单"""
    return df[df['单据类型'] == KEEP_ORDER_TYPE].copy()


def map_department(df):
    """根据经手人映射部门，替换原有部门列"""
    new_depts = []
    for _, row in df.iterrows():
        person = row['经手人']
        if pd.isna(person):
            person = ''
        person = str(person).strip()
        new_depts.append(DEPARTMENT_MAP.get(person, '其他'))
    df['部门'] = new_depts
    return df


def clean_numeric_columns(df):
    """清洗数值列：出库数量、出库单价、出库金额"""
    for col in ['出库数量', '出库单价', '出库金额']:
        if col in df.columns:
            df[col] = df[col].apply(safe_to_numeric)
    return df
```

#### 2.4 排序与汇总

```python
def sort_and_summarize(df):
    """按存货→部门排序，生成存货级分类汇总"""
    df = df.sort_values(by=['存货', '部门'], ascending=[True, True]).reset_index(drop=True)
    summary = df.groupby('存货').agg(
        出库数量合计=('出库数量', 'sum'),
        出库金额合计=('出库金额', 'sum'),
        记录条数=('出库数量', 'count'),
    ).reset_index()
    summary = summary.sort_values(by='存货', ascending=True).reset_index(drop=True)
    summary['出库数量合计'] = summary['出库数量合计'].round(2)
    summary['出库金额合计'] = summary['出库金额合计'].round(2)
    return df, summary
```

#### 2.5 Excel 输出生成（核心）

```python
def generate_excel_report(df, summary, output_path, log_func=None):
    """生成多工作表Excel输出文件"""
    wb = Workbook()
    wb.remove(wb.active)

    # 统一样式定义
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
    
    # ... 样式辅助函数（省略，见源码）
```

**输出工作表结构：**

1. **工作表 "汇总"** — 仅两列：`存货` | `出库金额合计`，底部含总计行，冻结首行
2. **工作表 "VCP"** — 部门=VCP 的明细（彭建森、罗德浮），含左侧明细表 + 底部汇总行 + 右侧存货汇总（N列起）
3. **工作表 "曝光"** — 部门=曝光 的明细（邹柳珍），同上结构
4. **工作表 "蚀刻"** — 部门=蚀刻 的明细（曹鹏），同上结构
5. **工作表 "品质"** — 部门=品质 的明细（刘志军），无数据则显示"无数据"
6. **工作表 "其他"** — 部门=其他 的明细（经手人未匹配或为空），同上结构

每个部门明细表右侧（从 N1 单元格开始）附加存货汇总表：

| N（存货） | O（总数量） | P（平均单价） | Q（总出库金额） |
|-----------|------------|-------------|-----------------|
| AR硫酸98%*越凯 | 3240 | 4.8 | 15552 |
| ... | ... | ... | ... |
| **合计** | **15323** | **37.47** | **574150.23** |

- 平均单价 = 总出库金额 ÷ 总数量（加权平均，数量为 0 时填 None 避免除零）
- 底部有合计行，含蓝色背景 + 粗体

#### 2.6 GUI 界面

使用 tkinter 构建，无需额外依赖：

- **文件选择区域** — 输入文件（浏览/拖放）+ 输出文件（自动生成或手动指定）
- **操作按钮** — "开始清洗" (绿色) + "打开输出文件" (灰色)
- **进度条** — 不确定模式滚条
- **日志区域** — 深色终端风格 Text widget，显示处理步骤和结果摘要
- **状态栏** — 底部状态文字

GUI 布局使用 `ttk.LabelFrame` 分组，`grid` 排列。处理逻辑在后台线程 `threading.Thread` 中执行，避免阻塞界面。

#### 2.7 命令行模式

虽然主流使用 GUI，但保留命令行入口以支持批处理和脚本调用：

```python
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='出入库流水自动化清洗')
    parser.add_argument('--input', '-i', default='出入库流水.xlsx')
    parser.add_argument('--output', '-o', default=None)
    args = parser.parse_args()
```

---

### 步骤 3: 创建 `README.md`

```markdown
# 出入库流水自动化清洗与分类汇总工具

基于 Python + tkinter 的 GUI 工具，用于自动处理企业出入库流水 Excel 文件。

## 功能

- 自动识别并跳过 Excel 元数据行（标题、表头等）
- 筛选**材料出库单**，排除进货单、销货单
- 根据经手人自动映射部门（VCP / 曝光 / 蚀刻 / 品质 / 其他）
- 按存货分组计算分类汇总
- 生成多工作表 Excel 输出（汇总 + 分部门明细 + 各表右侧存货汇总）

## 快速开始

\`\`\`bash
pip install pandas openpyxl
python process_inventory.py
\`\`\`

## 使用方法

1. 点击 **浏览** 选择输入 Excel 文件
2. 点击 **开始清洗**
3. 点击 **打开输出文件** 查看结果

也支持命令行模式：

\`\`\`bash
python process_inventory.py --input 出入库流水.xlsx --output 清洗结果.xlsx
\`\`\`

## 输出结构

| 工作表 | 内容 |
|--------|------|
| 汇总 | 全部存货的出库金额汇总 |
| VCP | 彭建森/罗德浮的出库明细 |
| 曝光 | 邹柳珍的出库明细 |
| 蚀刻 | 曹鹏的出库明细 |
| 品质 | 刘志军的出库明细 |
| 其他 | 未匹配经手人的出库明细 |

每个部门明细表右侧（N列起）附加该表按存货的总数量/平均单价/总出库金额汇总。

## 输入格式

Excel 文件要求：
- 前6行为标题/元信息（自动跳过）
- 第7-8行为列头
- 第9行起为数据
- 末尾含合计行（自动跳过）

## 部门映射

| 经手人 | 映射后的部门 |
|--------|-------------|
| 彭建森 | VCP |
| 罗德浮 | VCP |
| 曹鹏 | 蚀刻 |
| 邹柳珍 | 曝光 |
| 刘志军 | 品质 |
| 其他   | 其他 |

## 技术栈

Python 3.x | tkinter | pandas | openpyxl
```

---

### 步骤 4: 创建 `.gitignore`

```gitignore
*.xlsx
!sample_input.xlsx
__pycache__/
*.pyc
.DS_Store
```

---

## 完整项目文件清单

```
inventory-cleaner/
├── process_inventory.py    # 主程序（约 500 行，含 GUI + 处理逻辑）
├── README.md               # 项目说明（约 80 行）
├── prompt.md               # 本文件（Agent 复刻提示词）
└── .gitignore              # 忽略规则
```

---

## GitHub 仓库

- **仓库名称:** `churuku-cleaning`
- **URL:** `https://github.com/yourstudendjd/churuku-cleaning`
- **可见性:** 私有（Private）
- **主分支:** `main`

完整源码请查看上述仓库的 `process_inventory.py` 文件。本文件（prompt.md）包含了所有关键逻辑的代码片段，可用于复刻或修改。
