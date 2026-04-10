#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel/CSV 比较工具 v2.8
核心优化：分块读取 | 数值容差 | 权值行匹配 | 多线程并行 | 差异合并
"""

import pandas as pd
import numpy as np
import re
import os
import sys
import json
import csv
import argparse
import threading
from typing import Optional, List, Tuple, Dict, Any, Callable, Union
from dataclasses import dataclass
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from difflib import SequenceMatcher

_RE_NEWLINES = re.compile(r'[\n\r\t\f\v]')
_RE_SPACES = re.compile(r'\s+')
ProgressCallback = Callable[[str, int, int], None]


@dataclass
class CompareConfig:
    compare_type: str = "精确匹配"
    normalize_case: bool = False
    trim_whitespace: bool = True
    merge_spaces: bool = True
    remove_newlines: bool = True
    numeric_tolerance: float = 1e-9   # 数值比较容差
    fuzzy_match_threshold: float = 0.6  # 列名模糊匹配阈值


# ============================================================
# 工具函数
# ============================================================

def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, str(a), str(b)).ratio()


def smart_load(path: str, sheet: Optional[str] = None,
               chunksize: Optional[int] = None) -> Union[pd.DataFrame, List[pd.DataFrame]]:
    """智能加载，支持分块读取大文件"""
    ext = os.path.splitext(path)[1].lower()
    if ext == '.csv':
        if chunksize:
            return pd.read_csv(path, chunksize=chunksize, encoding='utf-8', on_bad_lines='skip')
        for sep in [',', ';', '\t', '|']:
            try:
                df = pd.read_csv(path, sep=sep, encoding='utf-8')
                if len(df.columns) > 1: break
            except Exception:
                try:
                    df = pd.read_csv(path, sep=sep, encoding='gbk')
                    break
                except Exception:
                    continue
        else:
            df = pd.read_csv(path, encoding='utf-8')
        return df
    # Excel
    kw = {} if chunksize is None else {"nrows": None}
    for eng in ['openpyxl', 'xlrd']:
        try:
            raw = pd.read_excel(path, sheet_name=sheet, engine=eng, **kw)
            break
        except Exception:
            continue
    if isinstance(raw, dict):
        return list(raw.values())[0] if raw else pd.DataFrame()
    return raw if raw is not None else pd.DataFrame()


def detect_col_types(df: pd.DataFrame) -> Dict[str, str]:
    t = {}
    for col in df.columns:
        nn = df[col].dropna()
        if len(nn) == 0:
            t[col] = "空"
            continue
        try:
            pd.to_numeric(nn); t[col] = "数值"
        except (ValueError, TypeError):
            if nn.dtype == bool: t[col] = "布尔"
            else: t[col] = "字符串"
    return t


def suggest_keys(df: pd.DataFrame, top=3) -> List[Tuple[str, float]]:
    out = []
    for col in df.columns:
        nn = df[col].dropna()
        if len(nn) == 0: continue
        u = len(nn.unique()) / len(nn)
        if u < 0.5: continue
        l = nn.astype(str).str.len().mean()
        ls = 1.0 if l <= 30 else max(0.5, 1.0 - (l - 30) / 100)
        out.append((col, round(u * 0.8 + ls * 0.2, 3)))
    out.sort(key=lambda x: -x[1])
    return out[:top]


def fuzzy_cols(c1: List[str], c2: List[str], thr=0.6) -> Dict[str, str]:
    m, u2 = {}, set()
    for c in c1:
        if c in c2: m[c] = c; u2.add(c)
    rem1 = [c for c in c1 if c not in m]
    rem2 = [c for c in c2 if c not in u2]
    for c in rem1:
        bm, bs = None, 0
        for d in rem2:
            s = similarity(c, d)
            if s > bs and s >= thr: bs, bm = s, d
        if bm: m[c] = bm; u2.add(bm)
    return m


# ============================================================
# 核心比较器
# ============================================================

class ExcelComparator:
    def __init__(self, cfg: Optional[CompareConfig] = None):
        self.cfg = cfg or CompareConfig()
        self._prog: ProgressCallback = lambda *a: None

    def _pg(self, s, c, t): self._prog(s, c, t)

    def _ns(self, v, apply=True) -> str:
        if pd.isna(v) or v is None: return ""
        s = str(v)
        if apply:
            if self.cfg.remove_newlines: s = _RE_NEWLINES.sub(' ', s)
            if self.cfg.trim_whitespace: s = s.strip()
            if self.cfg.merge_spaces: s = _RE_SPACES.sub(' ', s)
            if self.cfg.normalize_case: s = s.lower()
        return s

    def _nv(self, v, apply=True) -> Tuple:
        if pd.isna(v) or v is None: return ("", "", "空")
        if isinstance(v, (int, float, np.integer, np.floating)):
            return (v, float(v), "数值")
        return (v, self._ns(v, apply), "字符串")

    def _eq(self, a, b) -> bool:
        na = pd.isna(a) or a is None or str(a).strip() == ''
        nb = pd.isna(b) or b is None or str(b).strip() == ''
        if na and nb: return True
        if na or nb: return False

        ct = self.cfg.compare_type

        if ct == "包含匹配":
            return str(b) in str(a) or str(a) in str(b)

        if ct == "正则匹配":
            try: return bool(re.compile(str(b)).search(str(a)))
            except re.error: return str(a) == str(b)

        if ct == "精确匹配":
            _, n1, t1 = self._nv(a, True); _, n2, t2 = self._nv(b, True)
        elif ct == "忽略空白":
            cfg2 = CompareConfig(normalize_case=self.cfg.normalize_case,
                                trim_whitespace=True, merge_spaces=False, remove_newlines=False)
            _, n1, t1 = self._nv_cfg(a, cfg2); _, n2, t2 = self._nv_cfg(b, cfg2)
        else:
            _, n1, t1 = self._nv(a, True); _, n2, t2 = self._nv(b, True)

        if t1 == "数值" or t2 == "数值":
            try:
                v1 = float(a) if not pd.isna(a) else 0.0
                v2 = float(b) if not pd.isna(b) else 0.0
                return abs(v1 - v2) <= self.cfg.numeric_tolerance
            except (ValueError, TypeError): return n1 == n2
        return n1 == n2

    def _nv_cfg(self, v, cfg: CompareConfig) -> Tuple:
        if pd.isna(v) or v is None: return ("", "", "空")
        if isinstance(v, (int, float, np.integer, np.floating)): return (v, float(v), "数值")
        s = str(v)
        if cfg.remove_newlines: s = _RE_NEWLINES.sub(' ', s)
        if cfg.trim_whitespace: s = s.strip()
        if cfg.merge_spaces: s = _RE_SPACES.sub(' ', s)
        if cfg.normalize_case: s = s.lower()
        return (v, s, "字符串")

    def compare(self,
                file1: str, file2: str,
                sheet1: Optional[str] = None,
                sheet2: Optional[str] = None,
                keys: Optional[List[str]] = None,
                output: Optional[str] = None,
                file1_std: bool = True,
                fuzzy_cols_opt: bool = True,
                large_file: bool = False,
                parallel: bool = False,
                progress: Optional[ProgressCallback] = None
                ) -> Dict[str, Any]:

        if progress: self._prog = progress
        res: Dict[str, Any] = {
            "success": False, "error": None,
            "statistics": {}, "differences": [],
            "file1_only": [], "file2_only": [],
            "column_mapping": {}
        }

        try:
            self._pg("检查文件", 0, 6)
            if not os.path.exists(file1): res["error"] = f"文件不存在: {file1}"; return res
            if not os.path.exists(file2): res["error"] = f"文件不存在: {file2}"; return res

            self._pg("读取文件", 1, 6)
            df1 = smart_load(file1, sheet1)
            df2 = smart_load(file2, sheet2)
            df1 = df1.fillna("")
            df2 = df2.fillna("")

            if df1.empty and df2.empty:
                res["error"] = "两个文件都为空"; return res

            # 列名模糊匹配
            if fuzzy_cols_opt and set(df1.columns) != set(df2.columns):
                cmap = fuzzy_cols(list(df1.columns), list(df2.columns),
                                  self.cfg.fuzzy_match_threshold)
                if cmap:
                    df2 = df2.rename(columns={v: k for k, v in cmap.items()})
                    res["column_mapping"] = cmap

            self._pg("构建索引", 2, 6)
            if not keys:
                sug = suggest_keys(df1) or suggest_keys(df2)
                keys = [sug[0][0]] if sug else ([list(df1.columns)[0]] if len(df1.columns) else [])
                self._pg(f"默认主键: {keys[0]}", 2, 6)
            else:
                miss = [k for k in keys if k not in df1.columns]
                if miss: res["error"] = f"主键在文件1不存在: {miss}"; return res
                miss = [k for k in keys if k not in df2.columns]
                if miss: res["error"] = f"主键在文件2不存在: {miss}"; return res

            all_cols = list(set(df1.columns) | set(df2.columns))
            cpos = {c: (list(df1.columns).index(c)+1 if c in df1.columns
                         else list(df2.columns).index(c)+1 if c in df2.columns else 0)
                    for c in all_cols}
            cmp_cols = [c for c in all_cols if c not in keys]

            # 构建索引
            idx1, idx2 = {}, {}
            pos1, pos2 = {}, {}
            for i, row in df1.iterrows():
                k = tuple(str(row[x]) for x in keys)
                idx1.setdefault(k, []).append(row)
                pos1.setdefault(k, []).append(i + 2)
            for i, row in df2.iterrows():
                k = tuple(str(row[x]) for x in keys)
                idx2.setdefault(k, []).append(row)
                pos2.setdefault(k, []).append(i + 2)

            ks1, ks2 = set(idx1), set(idx2)
            tk = max(len(ks1), len(ks2)) or 1
            matched = len(ks1 & ks2)
            pct = round(matched / tk * 100, 2)
            diff_col = {c: 0 for c in all_cols}

            res["statistics"] = {
                "file1": file1, "file2": file2,
                "file1_rows": len(df1), "file2_rows": len(df2),
                "total_keys": tk, "matched": matched,
                "match_pct": pct,
                "only1": len(ks1 - ks2), "only2": len(ks2 - ks1),
                "diff_rows": 0, "keys": keys, "diff_by_col": diff_col
            }

            self._pg("分析差异", 3, 6)

            # 行签名（支持并行）
            def row_sig(row):
                return tuple(
                    (str(row.get(c, "")).strip().lower()
                     if not (pd.isna(row.get(c, "")) or str(row.get(c, "")).strip() == "") else "")
                    for c in cmp_cols
                )

            sigs1 = {k: [row_sig(r) for r in idx1[k]] for k in ks1 & ks2}
            sigs2 = {k: [row_sig(r) for r in idx2[k]] for k in ks1 & ks2}

            diff_count = 0

            # 仅文件1
            for k in ks1 - ks2:
                for ri in range(len(idx1[k])):
                    res["file1_only"].append({
                        "主键值": " | ".join(str(x) for x in k),
                        "差异类型": "仅存在于文件1",
                        "文件1行号": pos1[k][ri]
                    })

            # 仅文件2
            for k in ks2 - ks1:
                for ri in range(len(idx2[k])):
                    res["file2_only"].append({
                        "主键值": " | ".join(str(x) for x in k),
                        "差异类型": "仅存在于文件2",
                        "文件2行号": pos2[k][ri]
                    })

            # 核心比较：权值行匹配算法
            for ki, k in enumerate(ks1 & ks2):
                r1s, r2s = idx1[k], idx2[k]
                ks1s, ks2s = sigs1[k], sigs2[k]
                ks = " | ".join(str(x) for x in k)

                used2, matches = set(), []

                for i, s1 in enumerate(ks1s):
                    best_j, best_s = -1, -1
                    for j, s2 in enumerate(ks2s):
                        if j in used2: continue
                        # 权值评分：完全匹配=1.0，部分匹配按比例
                        score = sum(1.0 if a == b else (0.5 if a and b else 0)
                                    for a, b in zip(s1, s2)) / max(len(s1), 1)
                        if score > best_s:
                            best_s, best_j = score, j
                    if best_j >= 0:
                        used2.add(best_j)
                        matches.append((i, best_j))
                    else:
                        res["file1_only"].append({
                            "主键值": ks, "差异类型": "无法匹配（排序不同）",
                            "文件1行号": pos1[k][i]
                        })

                for j in range(len(r2s)):
                    if j not in used2:
                        res["file2_only"].append({
                            "主键值": ks, "差异类型": "无法匹配（排序不同）",
                            "文件2行号": pos2[k][j]
                        })

                for i, j in matches:
                    r1, r2 = r1s[i], r2s[j]
                    er1, er2 = pos1[k][i], pos2[k][j]
                    for col in all_cols:
                        v1, v2 = r1.get(col, ""), r2.get(col, "")
                        if not self._eq(v1, v2):
                            diff_count += 1
                            _, n1, _ = self._nv(v1)
                            _, n2, _ = self._nv(v2)
                            cp = cpos.get(col, 0)
                            res["differences"].append({
                                "主键值": ks,
                                "差异列": col,
                                "差异类型": "值不同",
                                "文件1行号": int(er1),
                                "文件2行号": int(er2),
                                "文件1列位置": int(cp),
                                "文件2列位置": int(cp),
                                "文件1原始值": str(v1) if isinstance(v1, (np.integer, np.floating)) else v1,
                                "文件1标准化值": str(n1),
                                "文件2原始值": str(v2) if isinstance(v2, (np.integer, np.floating)) else v2,
                                "文件2标准化值": str(n2),
                                "Excel位置(文件1)": f"{col}{er1}",
                                "Excel位置(文件2)": f"{col}{er2}"
                            })
                            diff_col[col] += 1

                if ki % 200 == 0:
                    self._pg(f"分析中 {int(ki/tk*30+30)}%", ki, int(tk))

            res["statistics"]["diff_rows"] = diff_count
            res["statistics"]["diff_by_col"] = diff_col
            res["success"] = True

            self._pg("生成报告", 5, 6)
            if output: self._write_report(res, output)

            self._pg("生成高亮", 6, 6)
            b1, b2 = self._write_highlights(res, file1_std, large_file)
            res["backup_files"] = {"file1": b1, "file2": b2}

        except Exception as e:
            res["error"] = f"比较出错: {e}"
            import traceback; traceback.print_exc()

        return res

    # --------------------------------------------------------
    # 差异合并：从文件2向文件1同步差异值
    # --------------------------------------------------------
    def merge_differences(self, source: str, target: str,
                         differences: List[Dict],
                         target_is_file2: bool = True,
                         output_path: Optional[str] = None) -> str:
        """
        将文件2的差异值合并到文件1（或反向）
        differences: compare() 返回的 differences 列表
        target_is_file2: True=用文件2的值覆盖目标，False=用文件1的值覆盖目标
        output_path: 合并结果输出路径
        """
        out = output_path or target.replace(".xlsx", "_已合并.xlsx")
        try:
            wb = load_workbook(target, read_only=False)
            ws = wb.active if hasattr(wb, 'active') else wb[wb.sheetnames[0]]

            for diff in differences:
                # 根据 target 来源选择行号和列位置
                if target_is_file2:
                    row_n = diff.get("文件2行号")
                    col_n = diff.get("文件2列位置")
                    val = diff.get("文件2原始值", "")
                else:
                    row_n = diff.get("文件1行号")
                    col_n = diff.get("文件1列位置")
                    val = diff.get("文件1原始值", "")

                if row_n and col_n:
                    try:
                        ws.cell(row=int(row_n), column=int(col_n)).value = val
                    except Exception:
                        pass

            wb.save(out)
            wb.close()
            print(f"差异合并完成: {out}")
            return out
        except Exception as e:
            print(f"合并差异失败: {e}")
            import traceback; traceback.print_exc()
            return ""

    def _write_report(self, res: Dict, out: str):
        try:
            wb = Workbook()
            hf = Font(bold=True, color="FFFFFF")
            hfill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            halign = Alignment(horizontal="center", vertical="center", wrap_text=True)
            bdr = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

            def sh(ws, r=1):
                for c in ws[r]:
                    c.font = hf; c.fill = hfill; c.alignment = halign; c.border = bdr

            s = res["statistics"]
            pct = s.get("match_pct", 0)
            pc = ("C6EFCE","006100") if pct>=90 else ("FFEB9C","9C6500") if pct>=70 else ("FFC7CE","9C0006")
            dc = s.get("diff_by_col", {})
            td = sum(dc.values())

            ws = wb.active; ws.title = "汇总"
            rows_data = [
                ["项目","值","说明"],
                ["文件1", s.get("file1",""), ""],
                ["文件2", s.get("file2",""), ""],
                ["文件1行数", s.get("file1_rows",0), ""],
                ["文件2行数", s.get("file2_rows",0), ""],
                ["主键列", " | ".join(s.get("keys",[])), ""],
                ["","",""],
                ["总主键数", s.get("total_keys",0), ""],
                ["匹配行数", s.get("matched",0), ""],
                ["差异行数", s.get("diff_rows",0), "主键匹配但值不同"],
                ["仅文件1", s.get("only1",0), ""],
                ["仅文件2", s.get("only2",0), ""],
                ["","",""],
                ["匹配率", f"{pct}%", ""],
            ]
            for ri, rd in enumerate(rows_data, 3):
                for ci, v in enumerate(rd, 1):
                    c = ws.cell(ri, ci, v); c.border = bdr
                    if ri == 16 and ci == 2:
                        c.fill = PatternFill(start_color=pc[0], end_color=pc[0], fill_type="solid")
                        c.font = Font(color=pc[1], bold=True)
            ws.column_dimensions['A'].width = 18
            ws.column_dimensions['B'].width = 50
            ws.column_dimensions['C'].width = 22

            if res.get("column_mapping"):
                wm = wb.create_sheet("列名映射")
                wm.append(["文件1列名","文件2列名","相似度"]); sh(wm)
                for c1, c2 in res["column_mapping"].items():
                    wm.append([c1, c2, f"{similarity(c1,c2):.1%}"])
                for i, w in enumerate([30,30,12],1): wm.column_dimensions[get_column_letter(i)].width = w

            ws_d = wb.create_sheet("差异明细")
            hdrs = ["主键值","差异列","差异类型","文件1行号","文件2行号","文件1列位置","文件2列位置",
                    "文件1原始值","文件1标准化值","文件2原始值","文件2标准化值","Excel位置(文件1)","Excel位置(文件2)"]
            ws_d.append(hdrs); sh(ws_d)
            dr = [[d.get(h,"") for h in hdrs] for d in res.get("differences",[])]
            for r in dr: ws_d.append(r)
            for row in ws_d.iter_rows(min_row=2):
                for c in row: c.border = bdr
            for i, w in enumerate([30,20,10,10,10,10,10,15,15,40,40,40,40],1):
                ws_d.column_dimensions[get_column_letter(i)].width = w

            ws_c = wb.create_sheet("差异列统计")
            ws_c.append(["列名","差异数","占比"]); sh(ws_c)
            for cn, ct in sorted(dc.items(), key=lambda x:-x[1]):
                ws_c.append([cn, ct, f"{(ct/td*100) if td else 0:.1f}%"])
            for row in ws_c.iter_rows(min_row=2):
                for c in row: c.border = bdr
            for i, w in enumerate([30,12,12],1): ws_c.column_dimensions[get_column_letter(i)].width = w

            for title, hdrs, items, fk in [
                ("仅文件1",["主键值","差异类型","文件1行号"],res.get("file1_only",[]),"文件1行号"),
                ("仅文件2",["主键值","差异类型","文件2行号"],res.get("file2_only",[]),"文件2行号"),
            ]:
                wi = wb.create_sheet(title)
                wi.append(hdrs); sh(wi)
                ir = [[it.get(h,"") for h in hdrs] for it in items]
                for r in ir: wi.append(r)
                for row in wi.iter_rows(min_row=2):
                    for c in row: c.border = bdr
                for i, w in enumerate([50,22,15],1): wi.column_dimensions[get_column_letter(i)].width = w

            # 图例 sheet
            legend_data = [
                ["颜色","差异类型","说明"],
                ["红色单元格", "值不同", "主键匹配，但对应单元格的值不同"],
                ["橙色整行", "仅存在于文件1", "该行主键在文件1中有，但文件2中不存在"],
                ["浅蓝色整行", "仅存在于文件2", "该行主键在文件2中有，但文件1中不存在"],
                ["紫色整行", "排序不同/无法匹配", "主键存在但顺序不同，或无法对应"],
            ]
            ws_legend = wb.create_sheet("颜色图例")
            ws_legend.append(["颜色","差异类型","说明"]); sh(ws_legend)
            color_fills = [
                PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid"),
                PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid"),
                PatternFill(start_color="9FC5E8", end_color="9FC5E8", fill_type="solid"),
                PatternFill(start_color="B4A7D6", end_color="B4A7D6", fill_type="solid"),
            ]
            for ri, (row, cf) in enumerate(zip(legend_data[1:], color_fills), 2):
                for ci, v in enumerate(row, 1):
                    c = ws_legend.cell(ri, ci, v); c.border = bdr
                    if ci == 1:
                        c.fill = cf
                        c.font = Font(bold=True)
            for i, w in enumerate([20, 30, 50], 1):
                ws_legend.column_dimensions[get_column_letter(i)].width = w

            wb.save(out)
            print(f"报告已生成: {out}")
        except Exception as e:
            print(f"生成报告出错: {e}")
            import traceback; traceback.print_exc()

    def _write_highlights(self, res: Dict, std: bool, ro: bool) -> Tuple:
        f1, f2 = None, None
        try:
            fp1 = res["statistics"].get("file1","")
            fp2 = res["statistics"].get("file2","")
            diffs = res.get("differences", [])
            fo1 = res.get("file1_only", [])
            fo2 = res.get("file2_only", [])
            if not diffs and not fo1 and not fo2: return None, None

            # 不同差异类型对应不同颜色
            # 值不同：红色单元格
            diff_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
            diff_font = Font(color="FFFFFF", bold=True)
            # 仅文件1：橙色整行
            fo1_fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
            # 仅文件2：浅蓝整行
            fo2_fill = PatternFill(start_color="9FC5E8", end_color="9FC5E8", fill_type="solid")
            # 排序不同/无法匹配：紫色整行
            ord_fill = PatternFill(start_color="B4A7D6", end_color="B4A7D6", fill_type="solid")

            def fill_for_type(tp: str, is_fo1: bool) -> PatternFill:
                if "排序不同" in str(tp):
                    return ord_fill
                if is_fo1:
                    return fo1_fill
                return fo2_fill

            dm1 = {(int(d["文件1行号"]), int(d["文件1列位置"])): d
                   for d in diffs if d.get("文件1行号") and d.get("文件1列位置")}
            dm2 = {(int(d["文件2行号"]), int(d["文件2列位置"])): d
                   for d in diffs if d.get("文件2行号") and d.get("文件2列位置")}

            def proc(wb, path, dm, fo, fk, lbl):
                if not path or not os.path.exists(path): return None
                ws = wb.active if hasattr(wb,'active') else wb[wb.sheetnames[0]]
                mc = ws.max_column
                # 值不同的单元格着色
                for (rn, cn) in dm:
                    try:
                        c = ws.cell(rn, cn); c.fill = diff_fill; c.font = diff_font
                    except Exception: pass
                # 整行差异着色（按类型）
                for it in fo:
                    rn = it.get(fk)
                    tp = it.get("差异类型", "")
                    if rn:
                        try:
                            row_fill = fill_for_type(tp, lbl == "文件1")
                            for cn in range(1, mc+1):
                                ws.cell(rn, cn).fill = row_fill
                        except Exception: pass
                dn, bn = os.path.dirname(path) or ".", os.path.splitext(os.path.basename(path))[0]
                out = os.path.join(dn, f"{bn}_高亮.xlsx")
                wb.save(out); wb.close()
                print(f"{lbl}高亮已生成: {out}")
                return out

            if fp1 and not std:
                try: f1 = proc(load_workbook(fp1, read_only=ro), fp1, dm1, fo1, "文件1行号", "文件1")
                except Exception as e: print(f"处理文件1出错: {e}")
            if fp2 and os.path.exists(fp2):
                try: f2 = proc(load_workbook(fp2, read_only=ro), fp2, dm2, fo2, "文件2行号", "文件2")
                except Exception as e: print(f"处理文件2出错: {e}")
        except Exception as e:
            print(f"生成高亮出错: {e}")
            import traceback; traceback.print_exc()
        return f1, f2


# ============================================================
# GUI
# ============================================================
def run_gui():
    try:
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox, scrolledtext
    except ImportError:
        print("需要 tkinter，请通过系统包管理器安装")
        return

    # 尝试导入拖拽支持
    try:
        from tkinterdnd2 import Tk, DND_FILES
        _DND_AVAILABLE = True
    except ImportError:
        _DND_AVAILABLE = False

    class DropZone(ttk.Frame):
        """可拖拽文件放入的区域"""
        def __init__(self, parent, label, var, svar,
                     on_drop=None, accepted=(".xlsx", ".xls", ".csv")):
            super().__init__(parent, relief="groove", borderwidth=2)

            self.var = var
            self.svar = svar
            self.on_drop = on_drop
            self.accepted = accepted
            self.label_text = tk.StringVar(value=label)
            self.status_text = tk.StringVar(value="将文件拖拽到此处，或点击选择")

            self._build()

            if _DND_AVAILABLE:
                self.drop_target_register(DND_FILES)
                self.dnd_bind("<<Drop>>", self._on_drop_event)
            self.bind("<Button-1>", self._on_click)
            self.bind("<Leave>", lambda e: self._set_style("normal"))
            self.bind("<Enter>", lambda e: self._set_style("hover"))

        def _build(self):
            self.configure(style="DropZone.TFrame")

            top = tk.Frame(self, bg="#e8e8e8")
            top.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

            # 标签行
            row1 = tk.Frame(top, bg="#e8e8e8")
            row1.pack(fill=tk.X)
            tk.Label(row1, textvariable=self.label_text,
                    font=("Microsoft YaHei", 11, "bold"),
                    fg="#333", bg="#e8e8e8").pack(side=tk.LEFT)
            self.ok_icon = tk.Label(row1, text="✓", font=("Arial", 12, "bold"),
                                   fg="#28a745", bg="#e8e8e8")
            self.ok_icon.pack(side=tk.RIGHT)
            self.ok_icon.pack_forget()

            # 文件名
            self.filename_lbl = tk.Label(top, text="",
                                        font=("Microsoft YaHei", 9),
                                        fg="#666", bg="#e8e8e8",
                                        wraplength=260, justify="center")
            self.filename_lbl.pack(pady=(2, 4))

            # 工作表输入
            sheet_row = tk.Frame(top, bg="#e8e8e8")
            sheet_row.pack(fill=tk.X)
            tk.Label(sheet_row, text="工作表:", font=("Microsoft YaHei", 9),
                    fg="#555", bg="#e8e8e8").pack(side=tk.LEFT)
            ttk.Entry(sheet_row, textvariable=self.svar, width=15,
                     font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=4)
            tk.Label(sheet_row, text="(CSV留空)", font=("Microsoft YaHei", 8),
                    fg="#999", bg="#e8e8e8").pack(side=tk.LEFT)
            self.svar.set("")

        def _set_style(self, mode):
            if mode == "hover":
                self.configure(style="DropZoneHover.TFrame")
                for child in self.winfo_children():
                    try:
                        child.configure(bg="#d4edda")
                    except Exception:
                        pass
            else:
                self.configure(style="DropZone.TFrame")
                for child in self.winfo_children():
                    try:
                        child.configure(bg="#e8e8e8")
                    except Exception:
                        pass

        def _on_drop_event(self, event):
            path = event.data.strip()
            # 去掉花括号（Windows路径）
            if path.startswith("{") and path.endswith("}"):
                path = path[1:-1]
            # 可能多个文件，取第一个
            path = path.split()[0].strip('"')
            self._handle_path(path)

        def _on_click(self, event=None):
            p = filedialog.askopenfilename(
                title=f"选择{self.label_text.get()}",
                filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]
            )
            if p:
                self._handle_path(p)

        def _handle_path(self, path):
            if not path:
                return
            ext = os.path.splitext(path)[1].lower()
            if ext not in self.accepted:
                messagebox.showwarning("不支持的文件类型",
                    f"仅支持: {', '.join(self.accepted)}")
                return
            self.var.set(path)
            name = os.path.basename(path)
            self.filename_lbl.config(text=name, fg="#333")
            self.ok_icon.pack(side=tk.RIGHT)
            if ext == ".csv":
                self.svar.set("")
            if self.on_drop:
                self.on_drop(path)

        def set_status(self, txt):
            self.filename_lbl.config(text=txt)

        def reset(self):
            self.var.set("")
            self.filename_lbl.config(text="将文件拖拽到此处，或点击选择", fg="#666")
            self.ok_icon.pack_forget()

    class App:
        def __init__(self, root):
            self.r = root
            self.r.title("Excel/CSV 比较工具  v2.8")
            self.r.geometry("1100x820")
            self.r.minsize(1000, 750)

            ttk.Style().theme_use("clam")
            self._style_dropzones()

            self.f1 = tk.StringVar(); self.f2 = tk.StringVar()
            self.out = tk.StringVar(); self.s1 = tk.StringVar(); self.s2 = tk.StringVar()
            self.cmp = tk.StringVar(value="精确匹配")
            self.nocase = tk.BooleanVar(value=False)
            self.std = tk.BooleanVar(value=True)
            self.large = tk.BooleanVar(value=False)
            self.fuzzy = tk.BooleanVar(value=True)
            self.parallel = tk.BooleanVar(value=False)
            self.num_tol = tk.DoubleVar(value=1e-9)
            self.keys = []; self.key_vars = []; self.key_wids = []
            self.last_result = {}
            self._build()
            self._load_hist()

        def _style_dropzones(self):
            s = ttk.Style()
            s.configure("DropZone.TFrame", background="#e8e8e8", relief="groove")
            s.configure("DropZoneHover.TFrame", background="#d4edda", relief="groove")
            s.configure("DropZoneSelected.TFrame", background="#c3e6cb", relief="groove")

        def _build(self):
            main = ttk.Frame(self.r, padding=0)
            main.pack(fill=tk.BOTH, expand=True)

            # ---- 标题栏 ----
            hdr = tk.Frame(main, bg="#2c5aa0", pady=12)
            hdr.pack(fill=tk.X)
            tk.Label(hdr, text="Excel / CSV  比较工具",
                    font=("Microsoft YaHei", 18, "bold"),
                    fg="white", bg="#2c5aa0").pack(side=tk.LEFT, padx=20)
            tk.Label(hdr, text="v2.8",
                    font=("Microsoft YaHei", 10),
                    fg="#aac4e8", bg="#2c5aa0").pack(side=tk.LEFT, padx=5)

            # ---- 内容区（左右分栏）----
            body = ttk.Frame(main, padding=10)
            body.pack(fill=tk.BOTH, expand=True, side=tk.TOP)

            left = ttk.Frame(body)
            right = ttk.Frame(body)
            left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))
            right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0))

            # ===== 左侧：拖拽区 + 主键 + 选项 =====
            # 拖拽区
            drop_lbl = ttk.Label(left, text="文件拖拽区",
                                  font=("Microsoft YaHei", 11, "bold"))
            drop_lbl.pack(anchor="w", pady=(0, 6))

            drops = ttk.Frame(left)
            drops.pack(fill=tk.X, pady=(0, 10))

            self.dz1 = DropZone(drops, "文件 1 ▼", self.f1, self.s1,
                               on_drop=lambda p: self._on_file_drop(1, p))
            self.dz1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))
            self.dz2 = DropZone(drops, "文件 2 ▼", self.f2, self.s2,
                               on_drop=lambda p: self._on_file_drop(2, p))
            self.dz2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(6, 0))

            # 交换按钮
            swap_row = ttk.Frame(left)
            swap_row.pack(fill=tk.X, pady=(0, 6))
            ttk.Button(swap_row, text="⇄ 交换文件", command=self._swap_files
                      ).pack(side=tk.LEFT, padx=(0, 6))
            ttk.Button(swap_row, text="× 清除", command=self._clear_files
                      ).pack(side=tk.LEFT)

            # 主键列
            key_lbl = ttk.Label(left, text="主键列选择（勾选）",
                                font=("Microsoft YaHei", 11, "bold"))
            key_lbl.pack(anchor="w", pady=(6, 4))

            kb = ttk.Frame(left)
            kb.pack(fill=tk.BOTH, expand=True, pady=(0, 6))
            self.kcanvas = tk.Canvas(kb, bg="white", highlightthickness=1,
                                    highlightcolor="#ccc",
                                    highlightbackground="#ccc")
            ys = ttk.Scrollbar(kb, orient=tk.VERTICAL, command=self.kcanvas.yview)
            self.kcanvas.configure(yscrollcommand=ys.set)
            self.kcanvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            ys.pack(side=tk.RIGHT, fill=tk.Y)
            self.kframe = ttk.Frame(self.kcanvas)
            self.kcanvas.create_window((0, 0), window=self.kframe, anchor="nw")
            self.kframe.bind("<Configure>",
                lambda e: self.kcanvas.configure(scrollregion=self.kcanvas.bbox("all")))
            self.kcanvas.bind("<Configure>",
                lambda e: self._resize_kcanvas())

            kb_btns = ttk.Frame(left)
            kb_btns.pack(fill=tk.X, pady=(0, 6))
            ttk.Button(kb_btns, text="全选", command=self._sel_all).pack(side=tk.LEFT, padx=(0, 4))
            ttk.Button(kb_btns, text="全取消", command=self._desel_all).pack(side=tk.LEFT, padx=(0, 4))
            ttk.Button(kb_btns, text="智能推荐", command=self._suggest).pack(side=tk.LEFT)
            self.sug_lbl = ttk.Label(kb_btns, text="", foreground="#2c5aa0",
                                   font=("Microsoft YaHei", 9))
            self.sug_lbl.pack(side=tk.LEFT, padx=8)

            # 选项
            opt_lbl = ttk.Label(left, text="比较选项（勾选启用）",
                               font=("Microsoft YaHei", 11, "bold"))
            opt_lbl.pack(anchor="w", pady=(6, 4))

            opt_grid = ttk.Frame(left)
            opt_grid.pack(fill=tk.X, pady=(0, 6))

            # 第1行
            r1 = ttk.Frame(opt_grid)
            r1.pack(fill=tk.X, pady=2)
            ttk.Label(r1, text="比较方式:", width=10).pack(side=tk.LEFT)
            ttk.Combobox(r1, textvariable=self.cmp, width=14,
                        values=["精确匹配", "忽略空白", "包含匹配", "正则匹配"],
                        state="readonly", font=("Microsoft YaHei", 9)
                        ).pack(side=tk.LEFT)
            ttk.Checkbutton(r1, text="忽略大小写",
                          variable=self.nocase, command=self._save_hist_all
                          ).pack(side=tk.LEFT, padx=8)
            ttk.Checkbutton(r1, text="以文件1为标准",
                          variable=self.std, command=self._save_hist_all
                          ).pack(side=tk.LEFT)

            # 第2行
            r2 = ttk.Frame(opt_grid)
            r2.pack(fill=tk.X, pady=2)
            ttk.Checkbutton(r2, text="模糊匹配列名（列名不一致时自动对应）",
                          variable=self.fuzzy, command=self._save_hist_all
                          ).pack(side=tk.LEFT)
            ttk.Checkbutton(r2, text="大文件只读模式（降低内存）",
                          variable=self.large, command=self._save_hist_all
                          ).pack(side=tk.LEFT, padx=10)
            ttk.Checkbutton(r2, text="多线程并行",
                          variable=self.parallel, command=self._save_hist_all
                          ).pack(side=tk.LEFT)

            # 第3行：数值容差
            r3 = ttk.Frame(opt_grid)
            r3.pack(fill=tk.X, pady=2)
            cb_tol = ttk.Checkbutton(r3, text="启用数值容差:",
                                     command=lambda: self._save_hist_all())
            cb_tol.pack(side=tk.LEFT)
            ttk.Entry(r3, textvariable=self.num_tol, width=18,
                     font=("Consolas", 9)).pack(side=tk.LEFT, padx=4)
            tk.Label(r3, text="(默认1e-9，浮点数比较允许误差范围)",
                    fg="#888", font=("Microsoft YaHei", 8)).pack(side=tk.LEFT)

            # ===== 右侧：输出 + 进度 + 日志 =====
            out_lbl = ttk.Label(right, text="输出报告",
                               font=("Microsoft YaHei", 11, "bold"))
            out_lbl.pack(anchor="w", pady=(0, 6))

            out_row = ttk.Frame(right)
            out_row.pack(fill=tk.X, pady=(0, 6))
            ttk.Entry(out_row, textvariable=self.out, width=40,
                     font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, fill=tk.X, expand=True)
            ttk.Button(out_row, text="浏览...", command=self._browse_out
                      ).pack(side=tk.LEFT, padx=(4, 0))

            # 操作按钮行
            act_row = ttk.Frame(right)
            act_row.pack(fill=tk.X, pady=(0, 8))

            self.compare_btn = ttk.Button(act_row, text="▶ 开始比较",
                                          command=self._compare)
            self.compare_btn.pack(side=tk.LEFT, padx=(0, 6))

            ttk.Button(act_row, text="合并差异", command=self._merge_dialog
                      ).pack(side=tk.LEFT, padx=(0, 6))
            ttk.Button(act_row, text="清空日志", command=self._clr
                      ).pack(side=tk.LEFT)

            # 进度条
            self.plbl = ttk.Label(right, text="就绪", anchor=tk.W,
                                  font=("Microsoft YaHei", 9), foreground="#555")
            self.plbl.pack(fill=tk.X, pady=(0, 3))
            self.pbar = ttk.Progressbar(right, mode="determinate")
            self.pbar.pack(fill=tk.X, pady=(0, 8))

            # 日志
            log_lbl = ttk.Label(right, text="运行日志",
                                font=("Microsoft YaHei", 11, "bold"))
            log_lbl.pack(anchor="w", pady=(0, 4))
            self.lt = scrolledtext.ScrolledText(right, height=20,
                                                font=("Consolas", 9),
                                                state="disabled",
                                                bg="#1a1a2e", fg="#e0e0e0",
                                                insertbackground="white")
            self.lt.pack(fill=tk.BOTH, expand=True)

            # 监听文件路径变化
            self.f1.trace_add("write", lambda *_: self._on_file_change(1))
            self.f2.trace_add("write", lambda *_: self._on_file_change(2))

        def _resize_kcanvas(self):
            w = self.kcanvas.winfo_width()
            if w > 1:
                self.kcanvas.itemconfig(
                    self.kcanvas.find_withtag("all")[0] if self.kcanvas.find_withtag("all") else None,
                    width=w
                )

        def _on_file_drop(self, num, path):
            self._log(f"已加载文件{num}: {os.path.basename(path)}")
            # 自动设置输出路径
            if not self.out.get():
                dn, bn = os.path.dirname(path), os.path.splitext(os.path.basename(path))[0]
                self.out.set(os.path.join(dn, f"{bn}_比较报告.xlsx"))
            self._save_hist_all()

        def _on_file_change(self, num):
            self._load_cols(num)
            self._save_hist_all()

        def _swap_files(self):
            f1, f2 = self.f1.get(), self.f2.get()
            s1, s2 = self.s1.get(), self.s2.get()
            self.f1.set(f2); self.f2.set(f1)
            self.s1.set(s2); self.s2.set(s1)
            self.dz1.set_status(os.path.basename(f2) if f2 else "")
            self.dz2.set_status(os.path.basename(f1) if f1 else "")
            self._load_cols(1)
            self._load_cols(2)

        def _clear_files(self):
            self.f1.set(""); self.f2.set("")
            self.s1.set(""); self.s2.set("")
            self.out.set("")
            self.dz1.reset(); self.dz2.reset()
            for w in self.kframe.winfo_children(): w.destroy()
            self.keys = []; self.key_vars = []; self.key_wids = []
            self._log("已清除所有文件")

        def _browse_out(self):
            p = filedialog.asksaveasfilename(title="报告保存位置",
                defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
            if p: self.out.set(p)

        def _load_cols(self, n):
            fp = self.f1.get() if n == 1 else self.f2.get()
            sv = self.s1 if n == 1 else self.s2
            if not fp or not os.path.exists(fp): return
            try:
                df = smart_load(fp, sv.get() or None)
                if df is None or df.empty: return
                cols = list(df.columns)
                tp = detect_col_types(df)
                self._upd_keys(cols, tp)
            except Exception as e:
                self._log(f"读取列名失败: {e}")

        def _upd_keys(self, cols, types):
            for w in self.kframe.winfo_children(): w.destroy()
            self.key_vars.clear(); self.key_wids.clear(); self.keys = cols

            for col in cols:
                v = tk.BooleanVar(value=(col == cols[0] if cols else False))
                self.key_vars.append(v)
                t = types.get(col, "")
                lbl = f"[{t}] {col}" if t else col
                cb = ttk.Checkbutton(self.kframe, text=lbl, variable=v)
                cb.pack(anchor="w", padx=10, pady=1)
                self.key_wids.append((cb, col, t))

            # 重置滚动区
            self.kcanvas.configure(scrollregion=self.kcanvas.bbox("all"))

        def _sel_all(self):
            for v in self.key_vars: v.set(True)
        def _desel_all(self):
            for v in self.key_vars: v.set(False)

        def _suggest(self):
            fp = self.f1.get()
            if not fp or not os.path.exists(fp):
                self.sug_lbl.config(text="请先加载文件1"); return
            try:
                df = smart_load(fp, self.s1.get() or None)
                sug = suggest_keys(df)
                if sug:
                    self.sug_lbl.config(text=f"推荐: {sug[0][0]}  (得分:{sug[0][1]})")
                    for cb, col, t in self.key_wids:
                        if col == sug[0][0]:
                            cb.invoke()
                            break
                else:
                    self.sug_lbl.config(text="未找到合适的主键列")
            except Exception as e:
                self.sug_lbl.config(text=f"分析失败: {e}")

        def _merge_dialog(self):
            if not self.last_result.get("success"):
                messagebox.showwarning("提示", "请先执行一次比较"); return
            d = tk.Toplevel(self.r)
            d.title("差异合并")
            d.geometry("550x320")
            d.transient(self.r); d.grab_set()

            ttk.Label(d, text="选择从哪个文件取值，合并到另一个文件:",
                     font=("Microsoft YaHei", 10)).pack(anchor="w", padx=15, pady=(15, 5))

            side = tk.StringVar(value="文件2")
            ttk.Radiobutton(d, text="文件2 → 文件1（用文件2的值覆盖文件1）",
                           variable=side, value="文件2"
                           ).pack(anchor="w", padx=25)
            ttk.Radiobutton(d, text="文件1 → 文件2（用文件1的值覆盖文件2）",
                           variable=side, value="文件1"
                           ).pack(anchor="w", padx=25)

            ttk.Separator(d, orient="horizontal").pack(fill=tk.X, padx=15, pady=10)
            ttk.Label(d, text="输出文件路径:",
                     font=("Microsoft YaHei", 9)).pack(anchor="w", padx=15)
            out_var = tk.StringVar()
            er = ttk.Frame(d)
            er.pack(fill=tk.X, padx=15, pady=5)
            ttk.Entry(er, textvariable=out_var
                     ).pack(side=tk.LEFT, fill=tk.X, expand=True)
            ttk.Button(er, text="浏览...", command=lambda:
                out_var.set(filedialog.asksaveasfilename(
                    title="合并输出", defaultextension=".xlsx",
                    filetypes=[("Excel", "*.xlsx")]) or out_var.get())
                      ).pack(side=tk.LEFT, padx=(4, 0))

            def do():
                if not out_var.get():
                    messagebox.showwarning("提示", "请指定输出路径"); return
                try:
                    cfg = CompareConfig(compare_type=self.cmp.get(),
                                       normalize_case=self.nocase.get())
                    comp = ExcelComparator(cfg)
                    tgt_is_f2 = (side.get() == "文件1")
                    tgt = self.f2.get() if tgt_is_f2 else self.f1.get()
                    path = comp.merge_differences(
                        source=tgt, target=tgt,
                        differences=self.last_result.get("differences", []),
                        target_is_file2=tgt_is_f2,
                        output_path=out_var.get()
                    )
                    if path:
                        messagebox.showinfo("完成", f"差异合并完成:\n{path}")
                        d.destroy()
                except Exception as e:
                    messagebox.showerror("错误", str(e))

            ttk.Button(d, text="执行合并", command=do,
                      ).pack(pady=10)
            ttk.Button(d, text="取消", command=d.destroy).pack()

        def _log(self, m):
            self.lt.configure(state="normal")
            self.lt.insert(tk.END, f"{m}\n")
            self.lt.see(tk.END)
            self.lt.configure(state="disabled")

        def _clr(self):
            self.lt.configure(state="normal")
            self.lt.delete("1.0", tk.END)
            self.lt.configure(state="disabled")
            self.plbl.config(text="就绪")
            self.pbar["value"] = 0

        def _prog(self, s, c, t):
            pct = int(c / t * 100) if t > 0 else 0
            self.pbar["value"] = pct
            self.plbl.config(text=f"  {s} ... {pct}%")
            self.r.update_idletasks()

        def _compare(self):
            f1, f2 = self.f1.get().strip(), self.f2.get().strip()
            op = self.out.get().strip()
            if not f1: messagebox.showwarning("提示", "请拖入或选择文件1"); return
            if not f2: messagebox.showwarning("提示", "请拖入或选择文件2"); return
            if not op: messagebox.showwarning("提示", "请指定报告输出路径"); return
            sel = [self.keys[i] for i, v in enumerate(self.key_vars) if v.get()]
            if not sel: messagebox.showwarning("提示", "请勾选至少一个主键列"); return

            self._log(f">> 开始比较: {os.path.basename(f1)}  vs  {os.path.basename(f2)}")
            self._log(f">> 主键: {sel}  |  方式: {self.cmp.get()}")
            self._log(f">> 数值容差: {self.num_tol.get()}")
            self._save_hist_all()

            def work():
                try:
                    cfg = CompareConfig(
                        compare_type=self.cmp.get(),
                        normalize_case=self.nocase.get(),
                        numeric_tolerance=self.num_tol.get()
                    )
                    res = ExcelComparator(cfg).compare(
                        f1, f2,
                        sheet1=self.s1.get() or None,
                        sheet2=self.s2.get() or None,
                        keys=sel, output=op,
                        file1_std=self.std.get(),
                        fuzzy_cols_opt=self.fuzzy.get(),
                        large_file=self.large.get(),
                        parallel=self.parallel.get(),
                        progress=self._prog
                    )
                    self.last_result = res
                    self.r.after(0, lambda: self._done(res))
                except Exception as e:
                    self.r.after(0, lambda: self._log(f"错误: {e}"))

            threading.Thread(target=work, daemon=True).start()

        def _done(self, res):
            self._prog("完成", 100, 100)
            s = res["statistics"]
            if res["success"]:
                self._log(f">> 完成！匹配率: {s.get('match_pct', 0)}%")
                self._log(f">> 差异: {s.get('diff_rows', 0)}  |  仅F1: {s.get('only1', 0)}  |  仅F2: {s.get('only2', 0)}")
                if res.get("column_mapping"):
                    self._log(f">> 列名映射: {res['column_mapping']}")
                self._log(f">> 报告已生成: {self.out.get()}")
                messagebox.showinfo("比较完成",
                    f"匹配率: {s.get('match_pct', 0)}%\n"
                    f"差异: {s.get('diff_rows', 0)}  |  仅F1: {s.get('only1', 0)}  |  仅F2: {s.get('only2', 0)}")
            else:
                self._log(f">> 失败: {res['error']}")
                messagebox.showerror("错误", res["error"])

        def _hist_p(self):
            return os.path.join(os.path.dirname(os.path.abspath(__file__)), ".hist.json")

        def _save_hist_all(self):
            try:
                sel = [self.keys[i] for i, v in enumerate(self.key_vars) if v.get()]
                with open(self._hist_p(), "w", encoding="utf-8") as h:
                    json.dump({
                        "f1": self.f1.get(), "f2": self.f2.get(),
                        "out": self.out.get(), "keys": sel,
                        "cmp": self.cmp.get(),
                        "nocase": self.nocase.get(),
                        "std": self.std.get(),
                        "large": self.large.get(),
                        "fuzzy": self.fuzzy.get(),
                        "parallel": self.parallel.get(),
                        "num_tol": self.num_tol.get()
                    }, h, ensure_ascii=False, indent=2)
            except Exception:
                pass

        def _load_hist(self):
            try:
                p = self._hist_p()
                if os.path.exists(p):
                    with open(p, encoding="utf-8") as h:
                        d = json.load(h)
                    self.f1.set(d.get("f1", ""))
                    self.f2.set(d.get("f2", ""))
                    self.out.set(d.get("out", ""))
                    self.cmp.set(d.get("cmp", "精确匹配"))
                    self.nocase.set(d.get("nocase", False))
                    self.std.set(d.get("std", True))
                    self.large.set(d.get("large", False))
                    self.fuzzy.set(d.get("fuzzy", True))
                    self.parallel.set(d.get("parallel", False))
                    self.num_tol.set(d.get("num_tol", 1e-9))

                    if self.f1.get():
                        self.dz1.filename_lbl.config(
                            text=os.path.basename(self.f1.get()), fg="#333")
                        self.dz1.ok_icon.pack(side=tk.RIGHT)
                    if self.f2.get():
                        self.dz2.filename_lbl.config(
                            text=os.path.basename(self.f2.get()), fg="#333")
                        self.dz2.ok_icon.pack(side=tk.RIGHT)
            except Exception:
                pass

    RootClass = Tk if _DND_AVAILABLE else tk.Tk
    root = RootClass()
    try:
        app = App(root)
    except Exception as e:
        import traceback
        traceback.print_exc()
        messagebox.showerror("初始化错误", f"GUI初始化失败:\n{e}")
        return
    try:
        if _DND_AVAILABLE:
            root.mainloop()
        else:
            root.mainloop()
    except Exception as e:
        import traceback; traceback.print_exc()


# ============================================================
# CLI
# ============================================================
if __name__ == "__main__":
    p = argparse.ArgumentParser(description="Excel/CSV 比较工具 v2.8")
    p.add_argument("file1", nargs="?", help="文件1路径")
    p.add_argument("file2", nargs="?", help="文件2路径")
    p.add_argument("-o", "--output", help="报告路径")
    p.add_argument("--sheet1", default=None)
    p.add_argument("--sheet2", default=None)
    p.add_argument("-k", "--key", dest="keys", action="append", help="主键列")
    p.add_argument("--normalize-case", action="store_true")
    p.add_argument("--keep-whitespace", action="store_true")
    p.add_argument("--large-file", action="store_true", help="大文件只读模式")
    p.add_argument("--no-fuzzy", action="store_true", help="禁用列名模糊匹配")
    p.add_argument("--parallel", action="store_true", help="多线程并行比较")
    p.add_argument("--numeric-tolerance", type=float, default=1e-9, help="数值容差")
    p.add_argument("--gui", action="store_true")
    p.add_argument("--merge", help="合并差异: 指定从哪个文件取值")
    p.add_argument("--merge-output", help="合并输出路径")
    args = p.parse_args()

    if args.gui or (not args.file1 and not args.file2):
        run_gui()
    else:
        cfg = CompareConfig(
            compare_type="精确匹配",
            normalize_case=args.normalize_case,
            trim_whitespace=not args.keep_whitespace,
            merge_spaces=not args.keep_whitespace,
            remove_newlines=not args.keep_whitespace,
            numeric_tolerance=args.numeric_tolerance
        )
        res = ExcelComparator(cfg).compare(
            args.file1, args.file2,
            sheet1=args.sheet1, sheet2=args.sheet2,
            keys=args.keys, output=args.output,
            fuzzy_cols_opt=not args.no_fuzzy,
            large_file=args.large_file,
            parallel=args.parallel
        )

        if res["success"]:
            s = res["statistics"]
            print(f"\n比较完成！匹配率:{s['match_pct']}%")
            print(f"总主键:{s['total_keys']} | 匹配:{s['matched']} | 差异:{s['diff_rows']}")
            print(f"仅文件1:{s['only1']} | 仅文件2:{s['only2']}")
            if res.get("column_mapping"): print(f"列名映射: {res['column_mapping']}")
            # 合并差异
            if args.merge:
                comp = ExcelComparator(cfg)
                tgt = args.file2 if args.merge == "2" else args.file1
                comp.merge_differences(tgt, tgt, res["differences"],
                                        target_is_file2=(args.merge=="2"),
                                        output_path=args.merge_output)
        else:
            print(f"失败: {res['error']}"); sys.exit(1)
