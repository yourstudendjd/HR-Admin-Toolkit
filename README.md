# HR 管理工具集

一个基于 PyQt5 的桌面工具，持续迭代中。当前整合三个模块：

| 模块 | 功能 | 引入版本 |
|------|------|---------|
| 考勤汇总 | 读取打卡 XLSX → 每日汇总表 + 迟到判定 + 乐捐金额明细 | v1.0 |
| 宿舍分摊 | 读取宿舍费用 XLSX → 按人按天分摊水电费 + 住宿费 | v1.0 |
| 出入库流水 | 读取出入库流水 XLSX → 按部门/存货清洗分类汇总 | v2.0 |

---

## 版本历史

### [v2.0](https://github.com/yourstudendjd/HR-Admin-Toolkit) — 当前版本（三模块）

**新增：出入库流水清洗模块**

- 自动跳过 Excel 元数据行，筛选材料出库单
- 按经手人自动映射部门（VCP / 曝光 / 蚀刻 / 品质 / 其他）
- 按存货分组汇总（总数量 / 计量单位 / 平均单价 / 出库金额合计）
- 生成 6 工作表 Excel 输出，含格式化样式

**运行：**
```bash
python main.py
# 或双击: D:\桌面\AI相关\HR管理工具集.vbs
```

**涉及文件：** `main.py` | `inventory_processor.py` | `prompt_inventory.md`

---

### [v1.0](https://github.com/yourstudendjd/HR-Admin-Toolkit/blob/main/main_v1.py) — 初始版本（双模块）

**包含：考勤汇总 + 宿舍分摊**

- 考勤：打卡记录汇总、班次匹配、迟到判定（白班/晚班）、乐捐金额自动计算
- 宿舍：水电费按人按天分摊、住宿费 50 元/天计费、批量填充支持

**运行：**
```bash
python main_v1.py
# 或双击: D:\桌面\AI相关\HR管理工具集_v1.vbs
```

**涉及文件：** `main_v1.py` | `attendance_processor.py` | `dormitory_processor.py` | `README_v1.md`

---

### 版本对比

| 对比维度 | v1.0 | v2.0（当前） |
|---------|------|-------------|
| 模块数 | 2 | 3 |
| 标签页 | 考勤汇总 / 宿舍分摊 | 考勤汇总 / 宿舍分摊 / 出入库流水 |
| 代码行数 | ~760 行 | ~1210 行 |
| 新增依赖 | PyQt5, openpyxl, pandas | 无新增（同上） |
| 主要改动 | — | 新增 InventoryWorker 线程 + InventoryTab 标签页 |

> **学习建议：** 对比 `main_v1.py` 和 `main.py` 的 diff，可以看到如何在现有 PyQt5 架构中新增一个功能模块。

---

## 快速开始

```bash
# 安装依赖
pip install -r requirements.txt

# 启动当前版本（v2.0，三模块）
python main.py

# 启动旧版本（v1.0，双模块）
python main_v1.py
```

## 打包为 EXE

```bash
build.bat          # 打包当前版本
```

EXE 生成在 `dist/` 目录下。

## 项目结构

```
├── main.py                    # PyQt5 主程序 — v2.0（三标签页）
├── main_v1.py                 # PyQt5 主程序 — v1.0（双标签页，保留学习）
│
├── attendance_processor.py    # 考勤处理逻辑（v1 / v2 共用）
├── dormitory_processor.py     # 宿舍分摊逻辑（v1 / v2 共用）
├── inventory_processor.py     # 出入库流水清洗逻辑（v2 新增）
│
├── PROMPT.md                  # Agent 复刻提示词（含完整源码）
├── prompt_inventory.md        # 出入库模块需求文档
│
├── README.md                  # 本文件（当前版本说明）
├── README_v1.md               # v1.0 项目说明（保留学习）
│
├── requirements.txt           # Python 依赖
├── run.bat                    # 一键启动 v2.0
├── run_v1.bat                 # 一键启动 v1.0
├── build.bat                  # 一键打包 EXE
└── .gitignore
```

## 乐捐规则（考勤模块）

| 迟到分钟 | 乐捐金额 |
|---------|---------|
| 1-10 分钟 | 15 元 |
| 11-30 分钟 | 30 元 |
| 31-59 分钟 | 100 元 |
| ≥60 分钟 | 300 元 |

## 部门映射（出入库模块）

| 经手人 | 映射部门 |
|--------|---------|
| 彭建森 | VCP |
| 罗德浮 | VCP |
| 曹鹏 | 蚀刻 |
| 邹柳珍 | 曝光 |
| 刘志军 | 品质 |
| 其他 | 其他 |

## 技术栈

Python 3.x | PyQt5 | openpyxl | pandas | PyInstaller | tkinter

## Agent 复刻

任何 AI Agent 读取仓库根目录的 `PROMPT.md` 即可一键复刻整个项目。支持 Claude Code、Cursor、GitHub Copilot、Windsurf 等。
