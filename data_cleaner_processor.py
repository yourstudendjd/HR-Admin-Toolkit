# -*- coding: utf-8 -*-
"""
对账对账数据清洗处理器 — 客户对账单专用

核心逻辑：按「客户型号 + 加工类型 + 加工要求」组合去重，
相同组合只保留一条记录，单价可按策略处理。
规格（如尺寸、包装方式）不纳入组合条件，即规格不同但前三项相同的行会合并。

支持批量处理：多张表格分别清洗后合并为一个工作簿（多个工作表）。
"""

import os
import re
import pandas as pd
from datetime import datetime


class DataCleanerProcessor:
    """客户对账单对账数据清洗处理器"""

    # 单价冲突处理策略
    PRICE_STRATEGIES = {
        '取首次出现的值': 'first',
        '取平均值': 'mean',
        '取最大值': 'max',
        '提示冲突报错': 'error',
    }

    def __init__(self, progress_callback=None):
        """
        Parameters:
            progress_callback: 进度回调函数，签名 (current, total, message)
        """
        self.progress_callback = progress_callback
        self.df_raw = None
        self.df_clean = None
        self.file_path = None
        self.header_row = 7  # 默认表头行（从0开始）
        self.column_mapping = {}
        self.price_strategy = 'first'
        self.summary = {}

    def set_progress(self, current, total, message=''):
        if self.progress_callback:
            self.progress_callback(current, total, message)

    # ------------------------------------------------------------------
    # 文件加载（单文件）
    # ------------------------------------------------------------------
    def load_file(self, file_path, header_row=7):
        """
        加载Excel文件，自动处理合并单元格和单位说明行。

        Parameters:
            file_path : str  文件路径
            header_row : int 表头所在行（从0开始计数）

        Returns:
            (columns, preview_rows, total_rows, error)
        """
        self.file_path = file_path
        self.header_row = header_row

        try:
            # 用 header=None 读取整个文件，避免合并单元格导致的列名错位
            df_temp = pd.read_excel(file_path, header=None, engine='openpyxl')

            # 提取表头行作为列名
            header = df_temp.iloc[header_row].tolist()
            col_names = []
            for i, h in enumerate(header):
                if pd.isna(h):
                    col_names.append(f'Unnamed_{i}')
                else:
                    col_names.append(str(h).strip())

            # 数据从表头下一行开始
            self.df_raw = df_temp.iloc[header_row + 1:].copy()
            self.df_raw.columns = col_names
            self.df_raw = self.df_raw.reset_index(drop=True)

            # 清理列名
            self.df_raw.columns = [str(c).strip() for c in self.df_raw.columns]

            # 过滤完全空行
            self.df_raw = self.df_raw.dropna(how='all')

            # 过滤单位说明行（如 "mm"、"铜箔" 等短字符串行）
            self._filter_unit_rows()

            columns = self.df_raw.columns.tolist()
            preview = self.df_raw.head(5).fillna('').to_dict(orient='records')
            return columns, preview, len(self.df_raw), None

        except Exception as e:
            return None, None, 0, str(e)

    def auto_detect_header_row(self, file_path, max_scan=25):
        """
        自动检测表头行号。
        扫描前 max_scan 行，找到非空列数最多、文字比例最高的行。
        """
        try:
            df_scan = pd.read_excel(file_path, header=None, nrows=max_scan, engine='openpyxl')
            best_row = 0
            best_score = -1

            for i in range(min(len(df_scan), max_scan)):
                row = df_scan.iloc[i]
                non_null = row.notna().sum()
                text_count = 0
                for v in row:
                    if pd.notna(v):
                        s = str(v).strip()
                        if s and not s.replace('.', '').replace('-', '').isdigit():
                            text_count += 1
                score = non_null + text_count * 2
                if score > best_score:
                    best_score = score
                    best_row = i

            return best_row
        except Exception:
            return 0

    def _filter_unit_rows(self):
        """
        过滤表头后可能存在的单位说明行。
        该行特点：非空值少、全是短字符串、不含数字。
        """
        if self.df_raw is None or len(self.df_raw) == 0:
            return

        def is_unit_row(row):
            non_null = row.dropna()
            if len(non_null) == 0:
                return False
            if len(non_null) > max(3, len(self.df_raw.columns) * 0.5):
                return False
            for v in non_null:
                s = str(v).strip()
                if any(c.isdigit() for c in s):
                    return False
                if len(s) > 5:
                    return False
            return True

        mask = ~self.df_raw.apply(is_unit_row, axis=1)
        dropped = len(self.df_raw) - mask.sum()
        if dropped > 0:
            self.df_raw = self.df_raw[mask].reset_index(drop=True)

    # ------------------------------------------------------------------
    # 核心清洗逻辑（单文件）
    # ------------------------------------------------------------------
    def clean_data(self):
        """
        执行数据清洗与分类（单文件）。

        组合键 = (客户型号, 加工类型, 加工要求)
        - 三个字段都相同 => 同一组 => 合并为1条
        - 规格不同 => 不影响分组 => 可以合并
        - 加工类型不同 => 不同组 => 不合并
        - 加工要求不同 => 不同组 => 不合并

        Returns:
            (df_clean, summary)
        """
        if self.df_raw is None:
            raise ValueError("请先加载文件")

        # 验证映射完整性
        missing = [f for f in ['客户型号', '单价', '加工类型', '规格', '加工要求']
                    if f not in self.column_mapping or self.column_mapping[f] is None]
        if missing:
            raise ValueError(f"请先完成列映射，缺少字段：{', '.join(missing)}")

        df = self.df_raw.copy()

        col_model = self.column_mapping['客户型号']
        col_price = self.column_mapping['单价']
        col_type = self.column_mapping['加工类型']
        col_spec = self.column_mapping['规格']
        col_req = self.column_mapping['加工要求']

        required_cols = [col_model, col_price, col_type, col_spec, col_req]
        for c in required_cols:
            if c not in df.columns:
                raise ValueError(f"列映射错误：找不到列 '{c}'")

        # ==================================================================
        # 关键步骤：构建组合键（分组依据）
        # ==================================================================
        # 组合键由3个字段构成：(客户型号, 加工类型, 加工要求)
        # 规格不纳入组合条件 — 规格不同但前三项相同的行会合并。
        #
        # 合并规则：
        #   - 客户型号 + 加工类型 + 加工要求 三者都相同 => 同一组 => 合并
        #   - 客户型号 + 加工类型 + 加工要求 相同，规格不同 => 合并
        #   - 加工类型不同 => 不同组 => 不合并
        #   - 加工要求不同 => 不同组 => 不合并
        #
        # 将值转为字符串元组，避免 NaN 导致的比较问题。
        def make_key(row):
            return (
                str(row[col_model]) if pd.notna(row[col_model]) else '',
                str(row[col_type]) if pd.notna(row[col_type]) else '',
                str(row[col_req]) if pd.notna(row[col_req]) else '',
            )

        self.set_progress(10, 100, "正在构建分组键...")
        df['__group_key__'] = df.apply(make_key, axis=1)

        total_before = len(df)

        # 按组合键分组
        grouped = df.groupby('__group_key__', sort=False, dropna=False)

        self.set_progress(30, 100, f"正在分组去重，共 {len(grouped)} 个组合...")

        # 构建清洗后的行列表
        clean_rows = []
        price_conflicts = 0
        group_count = len(grouped)
        processed = 0

        for key, group in grouped:
            # 每组取第一条记录作为代表行
            representative = group.iloc[0].copy()

            # 处理该组的单价
            prices = group[col_price].dropna()
            if len(prices) == 0:
                final_price = representative[col_price]
            else:
                unique_prices = prices.unique()
                if len(unique_prices) > 1:
                    price_conflicts += 1

                strategy = self.price_strategy
                if strategy == 'first':
                    final_price = prices.iloc[0]
                elif strategy == 'mean':
                    final_price = round(prices.mean(), 6)
                elif strategy == 'max':
                    final_price = prices.max()
                elif strategy == 'error':
                    if len(unique_prices) > 1:
                        raise ValueError(
                            f"单价冲突：客户型号='{representative[col_model]}', "
                            f"加工类型='{representative[col_type]}', "
                            f"规格='{representative[col_spec]}', "
                            f"加工要求='{representative[col_req]}' 存在多个不同单价："
                            f"{list(unique_prices)}"
                        )
                    final_price = prices.iloc[0]
                else:
                    final_price = prices.iloc[0]

            representative[col_price] = final_price
            clean_rows.append(representative)

            processed += 1
            if processed % 50 == 0 or processed == group_count:
                pct = 30 + int(processed / group_count * 60)
                self.set_progress(pct, 100, f"处理中... {processed}/{group_count} 组")

        # 重建DataFrame（保留所有原始列，去掉辅助列）
        self.df_clean = pd.DataFrame(clean_rows).drop(columns=['__group_key__'], errors='ignore')
        # 恢复原始列顺序
        self.df_clean = self.df_clean[self.df_raw.columns]

        # ==================================================================
        # 输出列处理：格式化日期、隐藏不需要的列
        # ==================================================================
        # 1. 日期列 → 短日期格式（YYYY-MM-DD）
        #    处理两种情况：(a) pandas已识别为datetime (b) Excel序列号整数（存为object因为有NaN）
        for col in self.df_clean.columns:
            if pd.api.types.is_datetime64_any_dtype(self.df_clean[col]):
                # 已是日期类型，直接格式化
                self.df_clean[col] = pd.to_datetime(self.df_clean[col], errors='coerce').dt.strftime('%Y-%m-%d')
            else:
                # 尝试转为数值，检查是否是Excel日期序列号
                numeric = pd.to_numeric(self.df_clean[col], errors='coerce')
                valid = numeric.dropna()
                if len(valid) > 0:
                    int_vals = valid.astype(int)
                    # Excel序列号特征：值在 30000~60000 之间（约对应 1982-2064年）
                    if (int_vals > 30000).all() and (int_vals < 60000).all():
                        self.df_clean[col] = pd.to_datetime(
                            numeric, unit='D', origin='1899-12-30', errors='coerce'
                        ).dt.strftime('%Y-%m-%d')

        # 2. 删除不需要输出的列
        #    - 规格列 及紧随其后的空列（Unnamed_ 开头的合并列）
        #    - 送货单号
        #    - 总价
        drop_keywords = ['规格', '送货单号', '总价']
        cols_to_drop = []
        for col in self.df_clean.columns:
            for kw in drop_keywords:
                if kw in col:
                    cols_to_drop.append(col)
                    break

        # 规格列后面的 Unnamed_ 空列也一并删除
        col_list = list(self.df_clean.columns)
        for col in list(cols_to_drop):
            if col in col_list:
                idx = col_list.index(col)
                if idx + 1 < len(col_list):
                    nxt = col_list[idx + 1]
                    if nxt.startswith('Unnamed_') and nxt not in cols_to_drop:
                        cols_to_drop.append(nxt)

        if cols_to_drop:
            self.df_clean = self.df_clean.drop(columns=cols_to_drop, errors='ignore')
            print(f"[清洗] 已隐藏列: {', '.join(cols_to_drop)}")

        self.set_progress(95, 100, "正在生成输出文件...")

        total_after = len(self.df_clean)
        merged_count = total_before - total_after

        self.summary = {
            '原始行数': total_before,
            '清洗后行数': total_after,
            '合并/删除行数': merged_count,
            '单价冲突组数': price_conflicts,
            '处理时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }

        self.set_progress(100, 100, "处理完成！")
        return self.df_clean, self.summary

    def save_file(self, output_path=None):
        """
        保存清洗后的Excel文件（单文件）。

        Parameters:
            output_path : str 输出文件路径。如果为None，则自动生成。
        """
        if self.df_clean is None:
            raise ValueError("请先执行数据清洗")

        if output_path is None:
            base, ext = os.path.splitext(self.file_path)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_path = f"{base}_清洗后_{timestamp}{ext}"

        self.df_clean.to_excel(output_path, index=False, engine='openpyxl')
        return output_path

    # ------------------------------------------------------------------
    # 批量处理（多文件 → 一个工作簿多工作表）
    # ------------------------------------------------------------------

    @staticmethod
    def _sheet_name_from_path(file_path):
        """从文件路径提取工作表名称（去掉扩展名和路径，限制长度）"""
        name = os.path.splitext(os.path.basename(file_path))[0]
        # 去除不合法的 Excel 工作表名称字符
        name = re.sub(r'[\\/?*\[\]]', '', name)
        # Excel 工作表名最多31字符
        if len(name) > 31:
            name = name[:28] + '...'
        return name if name else 'Sheet'

    # ==================================================================
    # 模糊匹配列名
    # ==================================================================
    # 每个文件的列名可能不同（如"客户产品型号" vs "客户型号"），
    # 这里用关键词包含关系做模糊匹配，不要求完全一致。

    # 每个目标字段对应的匹配关键词列表（按优先级排序）
    FIELD_KEYWORDS = {
        '客户型号': ['客户型号', '客户产品型号', '产品型号', '型号', '品名', '客户品号'],
        '单价':     ['单价', '单价（元/个）', '单价(元/个)', '单价(元)', '单价_元', '单价(元/个)', '价格', 'unit_price'],
        '加工类型': ['加工类型', '加工工艺', '工艺', '加工', '类型'],
        '规格':     ['规格', '客户规格', '产品规格', '尺寸', '大小'],
        '加工要求': ['加工要求', '要求', '加工标准', '工艺要求'],
    }

    @classmethod
    def fuzzy_match_columns(cls, columns):
        """
        对给定的列名列表，用模糊匹配找出每个核心字段对应的列。

        Parameters:
            columns: list of str，文件的列名列表

        Returns:
            mapping: dict，如 {'客户型号': '客户产品型号', '单价': '单价', ...}
                     未能匹配的字段不包含在返回值中。
        """
        mapping = {}
        used_cols = set()

        # 标准化列名用于比较
        normalized = {c: c.strip() for c in columns}

        for field, keywords in cls.FIELD_KEYWORDS.items():
            best_col = None
            best_score = 0

            for col in normalized.values():
                if col in used_cols:
                    continue
                col_lower = col.lower()
                for kw in keywords:
                    kw_lower = kw.lower()
                    # 完全匹配
                    if col_lower == kw_lower:
                        if 2 > best_score:
                            best_score = 2
                            best_col = col
                        break
                    # 包含匹配：列名包含关键词，或关键词包含列名
                    elif kw_lower in col_lower or col_lower in kw_lower:
                        if 1 > best_score:
                            best_score = 1
                            best_col = col
                        break

            if best_col:
                mapping[field] = best_col
                used_cols.add(best_col)

        return mapping

    def clean_and_merge(self, file_tasks, output_path):
        """
        批量清洗多个文件，合并为一个多工作表工作簿。

        Parameters:
            file_tasks : list of dict，每个 dict 包含：
                - 'path'        : str   文件路径
                - 'header_row'  : int   表头行（可选，默认7）
                - 'mapping'     : dict  列映射（可选，使用全局 mapping）
                - 'strategy'    : str   单价策略（可选，使用全局策略）
            output_path : str  输出工作簿路径

        Returns:
            (sheet_names, results_summary, error)
            - sheet_names: 生成的工作表名称列表
            - results_summary: 每个文件的处理摘要列表
            - error: 错误信息（成功则为None）
        """
        if not file_tasks:
            return [], [], "没有可处理的任务"

        import openpyxl

        total_files = len(file_tasks)
        writer = None

        try:
            writer = pd.ExcelWriter(output_path, engine='openpyxl')
        except Exception as e:
            return [], [], f"无法创建输出文件: {e}"

        sheet_names = []
        results_summary = []
        has_error = False

        for idx, task in enumerate(file_tasks):
            file_path = task['path']
            header_row = task.get('header_row', 7)
            mapping = task.get('mapping', self.column_mapping)
            strategy = task.get('strategy', self.price_strategy)

            file_name = os.path.basename(file_path)
            pct_start = int(idx / total_files * 100)
            pct_end = int((idx + 1) / total_files * 100)
            self.set_progress(pct_start, 100, f"正在处理 ({idx+1}/{total_files}): {file_name}")

            # 使用独立的处理器实例处理每个文件
            proc = DataCleanerProcessor(progress_callback=self.progress_callback)
            proc.price_strategy = strategy

            # 加载
            columns, preview, total_rows, error = proc.load_file(file_path, header_row)
            if error:
                results_summary.append({
                    '文件名': file_name,
                    '状态': '加载失败',
                    '错误': error,
                    '原始行数': 0,
                    '清洗后行数': 0,
                    '合并/删除行数': 0,
                    '单价冲突组数': 0,
                })
                has_error = True
                continue

            # 列映射验证
            if not mapping or len(mapping) < 5:
                results_summary.append({
                    '文件名': file_name,
                    '状态': '跳过（未配置列映射）',
                    '错误': '',
                    '原始行数': total_rows,
                    '清洗后行数': 0,
                    '合并/删除行数': 0,
                    '单价冲突组数': 0,
                })
                has_error = True
                continue

            proc.column_mapping = mapping

            # 清洗
            try:
                df_clean, summary = proc.clean_data()
            except Exception as e:
                results_summary.append({
                    '文件名': file_name,
                    '状态': '清洗失败',
                    '错误': str(e),
                    '原始行数': total_rows,
                    '清洗后行数': 0,
                    '合并/删除行数': 0,
                    '单价冲突组数': 0,
                })
                has_error = True
                continue

            # 写入工作表
            sheet_name = self._sheet_name_from_path(file_path)
            # 处理工作表名重复
            original_name = sheet_name
            counter = 2
            while sheet_name in sheet_names:
                sheet_name = original_name[:28] + f'_{counter}' if len(original_name) > 26 else f"{original_name}_{counter}"
                counter += 1

            try:
                proc.df_clean.to_excel(writer, sheet_name=sheet_name, index=False, engine='openpyxl')
            except Exception as e:
                results_summary.append({
                    '文件名': file_name,
                    '状态': '写入失败',
                    '错误': str(e),
                    '原始行数': summary['原始行数'],
                    '清洗后行数': summary['清洗后行数'],
                    '合并/删除行数': summary['合并/删除行数'],
                    '单价冲突组数': summary['单价冲突组数'],
                })
                has_error = True
                continue

            sheet_names.append(sheet_name)
            summary['文件名'] = file_name
            summary['工作表'] = sheet_name
            summary['状态'] = '成功'
            results_summary.append(summary)

        # 生成汇总工作表
        try:
            summary_data = []
            for r in results_summary:
                summary_data.append({
                    '文件名': r.get('文件名', ''),
                    '工作表': r.get('工作表', ''),
                    '状态': r.get('状态', ''),
                    '原始行数': r.get('原始行数', 0),
                    '清洗后行数': r.get('清洗后行数', 0),
                    '合并/删除行数': r.get('合并/删除行数', 0),
                    '单价冲突组数': r.get('单价冲突组数', 0),
                    '处理时间': r.get('处理时间', ''),
                    '错误': r.get('错误', ''),
                })
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='_处理汇总', index=False, engine='openpyxl')
            sheet_names.insert(0, '_处理汇总')
        except Exception:
            pass  # 汇总写入失败不影响主流程

        writer.close()

        self.set_progress(100, 100, "全部完成！")
        return sheet_names, results_summary, None if not has_error else "部分文件处理失败，请查看汇总表"
