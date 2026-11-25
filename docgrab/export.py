"""导出：把提取结果落成文件。

- 全文：每份文档一个 .txt，或合并成一个大 .txt
- 字段汇总：一行一份文档的对照表（xlsx / csv）
- 表格列合并：把多个表格文件的同名列按列名对齐，拼成一张大表
"""
from __future__ import annotations

import csv
import os
from dataclasses import dataclass

from .extract import ExtractResult


def _safe_name(name: str) -> str:
    bad = '<>:"/\\|?*'
    for ch in bad:
        name = name.replace(ch, "_")
    return name.strip() or "output"


# ---------- 全文导出 ----------

def export_text_per_file(results: list[ExtractResult], out_dir: str) -> list[str]:
    """每份文档导出一个 txt；重名的自动加序号。返回写出的文件列表。"""
    os.makedirs(out_dir, exist_ok=True)
    written: list[str] = []
    used: set[str] = set()
    for res in results:
        if res.error and not res.text:
            continue
        stem = os.path.splitext(res.doc.name)[0]
        base = _safe_name(stem)
        candidate = base
        i = 1
        while candidate.lower() in used:
            candidate = "%s_%d" % (base, i)
            i += 1
        used.add(candidate.lower())
        path = os.path.join(out_dir, candidate + ".txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write(res.text)
        written.append(path)
    return written


def export_text_merged(results: list[ExtractResult], out_path: str) -> str:
    """把所有文档合并进一个大 txt，用分隔线标出每份来源。"""
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        for idx, res in enumerate(results, 1):
            f.write("=" * 70 + "\n")
            f.write("[%d] %s\n" % (idx, res.doc.path))
            if res.error:
                f.write("!! 错误：%s\n" % res.error)
            f.write("=" * 70 + "\n")
            f.write(res.text)
            f.write("\n\n")
    return out_path


# ---------- 字段汇总表 ----------

def export_rows(rows: list[dict[str, str]], out_path: str, columns: list[str] | None = None) -> str:
    """把"一行一份文档"的字段汇总写成 xlsx 或 csv（按扩展名判断）。"""
    if not rows:
        raise ValueError("没有可导出的数据")
    cols = columns or _ordered_columns(rows)
    ext = os.path.splitext(out_path)[1].lower()
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    if ext == ".csv":
        # Excel 打开 CSV 需要 BOM 才不乱码
        with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
            w.writeheader()
            for r in rows:
                w.writerow(r)
        return out_path

    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "汇总"
    ws.append(cols)
    for r in rows:
        ws.append([r.get(c, "") for c in cols])
    # 简单美化：首行加粗、冻结、按内容设列宽
    from openpyxl.styles import Font, PatternFill
    head_fill = PatternFill("solid", fgColor="4F81BD")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
    ws.freeze_panes = "A2"
    for i, col in enumerate(cols, 1):
        width = max([len(str(col))] + [len(str(r.get(col, ""))) for r in rows[:200]])
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(max(width + 2, 8), 60)
    wb.save(out_path)
    return out_path


def _ordered_columns(rows: list[dict[str, str]]) -> list[str]:
    cols: list[str] = []
    for r in rows:
        for k in r:
            if k not in cols:
                cols.append(k)
    return cols


# ---------- 表格列合并 ----------

@dataclass
class TableSource:
    """一张被纳入合并的表：来自哪个文件、第几个表、表头、数据行。"""
    file_name: str
    file_path: str
    sheet_index: int
    header: list[str]
    rows: list[list[str]]


def collect_tables(results: list[ExtractResult]) -> list[TableSource]:
    """从提取结果里把每张表展开为 TableSource，首行当表头。"""
    out: list[TableSource] = []
    for res in results:
        for idx, grid in enumerate(res.tables):
            if not grid:
                continue
            header = [str(c).strip() for c in grid[0]]
            rows = [[str(c) for c in row] for row in grid[1:]]
            out.append(TableSource(res.doc.name, res.doc.path, idx, header, rows))
    return out


def available_columns(sources: list[TableSource]) -> list[str]:
    """所有表头里出现过的列名（去重保序），供 GUI 让用户勾选。"""
    cols: list[str] = []
    for s in sources:
        for h in s.header:
            if h and h not in cols:
                cols.append(h)
    return cols


def merge_tables(sources: list[TableSource], chosen_cols: list[str],
                 add_source_col: bool = True) -> list[dict[str, str]]:
    """按列名把多张表对齐合并：只取 chosen_cols 中的列，缺失填 ''。"""
    rows_out: list[dict[str, str]] = []
    for s in sources:
        # 建立"列名 -> 该表列索引"映射（同名取第一个）
        index = {}
        for i, h in enumerate(s.header):
            if h and h not in index:
                index[h] = i
        # 该表压根没有用户选的任何列，就跳过，避免产生空行噪音
        if not any(c in index for c in chosen_cols):
            continue
        for row in s.rows:
            item: dict[str, str] = {}
            if add_source_col:
                item["来源文件"] = s.file_name
            for c in chosen_cols:
                idx = index.get(c)
                item[c] = row[idx] if idx is not None and idx < len(row) else ""
            rows_out.append(item)
    return rows_out


def merge_tables_xlsx(sources: list[TableSource], chosen_cols: list[str], out_path: str,
                      add_source_col: bool = True) -> tuple[str, int]:
    rows = merge_tables(sources, chosen_cols, add_source_col)
    cols = (["来源文件"] if add_source_col else []) + list(chosen_cols)
    export_rows(rows, out_path, cols)
    return out_path, len(rows)