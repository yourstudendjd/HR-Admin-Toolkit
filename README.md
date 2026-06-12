# HR 管理工具集

一个基于 PyQt5 的桌面工具，持续迭代中。当前包含三个功能模块：

| 模块 | 功能 |
|------|------|
| 考勤汇总 | 读取打卡 XLSX → 每日汇总表 + 迟到判定 + 乐捐金额明细 |
| 宿舍分摊 | 读取宿舍费用 XLSX → 按人按天分摊水电费 + 住宿费 |
| 出入库流水 | 读取出入库流水 XLSX → 按部门/存货清洗分类汇总 |

---

## 快速开始

```bash
# 安装依赖
pip install -r requirements.txt

# 启动
python main.py
```

## 打包为 EXE

```
build.bat
```

EXE 生成在 `dist/` 目录下。

## 项目结构

```
├── main.py                      # PyQt5 主程序（三标签页）
├── attendance_processor.py      # 考勤处理逻辑
├── dormitory_processor.py       # 宿舍分摊逻辑
├── inventory_processor.py       # 出入库流水清洗逻辑
├── PROMPT.md                    # Agent 复刻提示词（含完整源码）
├── prompt_inventory.md          # 出入库模块需求文档
├── README.md                    # 本文件
├── requirements.txt             # Python 依赖
├── run.bat                      # 一键启动
├── build.bat                    # 一键打包 EXE
├── .archive/                    # 历史版本归档（不显示在 GitHub 默认视图）
│   ├── dormitory_processor_v1_原始版.py
│   └── v1.0/
│       ├── attendance_processor.py
│       ├── dormitory_processor.py
│       ├── main.py
│       └── ...
└── .gitignore
```

## 乐捐规则（考勤模块）

| 迟到分钟 | 乐捐金额 |
|---------|---------|
| 1-10 分钟 | 15 元 |
| 11-30 分钟 | 30 元 |
| 31-59 分钟 | 100 元 |
| ≥60 分钟 | 300 元 |

## 宿舍分摊规则

- **住宿费**：50 元/天，按当月实际居住天数比例计算
- **水电费**：按房间汇总后，按每人居住天数占房间总人天的比例分摊
- **同名员工**：如员工换宿舍出现在多个房间，住宿费仅在天数最多的房间显示，水电费在各房间均计算

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

Python 3.x | PyQt5 | openpyxl | pandas | PyInstaller

## Agent 复刻

读取仓库根目录的 `PROMPT.md` 即可一键复刻整个项目。支持 Claude Code、Cursor、GitHub Copilot、Windsurf 等。
