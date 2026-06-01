# 考勤与宿舍管理工具集

一个基于 PyQt5 的桌面工具，整合两个模块：

| 模块 | 功能 |
|---|---|
| 考勤汇总 | 读取打卡 XLSX → 每日汇总表 + 迟到判定 + 乐捐金额明细 |
| 宿舍分摊 | 读取宿舍费用 XLSX → 按人按天分摊水电费 + 住宿费（50元/天） |

## 快速开始

```bash
pip install -r requirements.txt
python main.py
```

## 打包为 EXE

```bash
build.bat
```

EXE 生成在 `dist/` 目录下，可直接分发给未安装 Python 的用户。

## 项目结构

```
├── main.py                    # PyQt5 主程序（双标签页）
├── attendance_processor.py    # 考勤处理逻辑
├── dormitory_processor.py     # 宿舍分摊逻辑
├── PROMPT.md                  # Agent 复刻提示词（含完整源码）
├── requirements.txt           # Python 依赖
├── run.bat                    # 一键启动
├── build.bat                  # 一键打包 EXE
└── .gitignore
```

## 输入文件格式

### 考勤打卡文件

| 日期 | 时间 | 姓名 |
|---|---|---|
| 2024/01/15 | 07:55:30 | 张三 |
| 2024/01/15 | 17:02:00 | 张三 |

### 班次表（可选）

- 第 2 行第 7 列起为日期
- 第 3 列（从第 4 行起）为员工姓名
- 单元格值："白"（白班 08:00）或 "晚"（晚班 20:00）

### 宿舍费用文件

- **工作表 "水电总计表"** — 房间/月份 → 水电金额
- **工作表 "房屋表"** — 房间号、入住人员、入住/截止日期、入住人数

## 乐捐规则

| 迟到分钟 | 乐捐金额 |
|---|---|
| 1-10 分钟 | 15 元 |
| 11-30 分钟 | 30 元 |
| 31-59 分钟 | 100 元 |
| ≥60 分钟 | 300 元 |

## 技术栈

Python 3.x | PyQt5 | openpyxl | pandas | PyInstaller

## Agent 复刻

任何 AI Agent 读取 `PROMPT.md` 即可一键复刻整个项目。支持 Claude Code、Cursor、GitHub Copilot、Windsurf 等。
