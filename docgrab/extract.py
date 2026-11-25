"""扫描文件夹、识别文档类型、提取文本与表格数据。

只读原文件；旧格式转换走临时文件，用完即删。
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field

# 每种扩展名对应到"解析家族"
_EXT_FAMILY = {
    ".docx": "word", ".doc": "word",
    ".xlsx": "excel", ".xls": "excel", ".xlsm": "excel",
    ".pptx": "ppt", ".ppt": "ppt",
    ".txt": "text", ".csv": "text", ".md": "text",
}
_LEGACY = {".doc", ".xls", ".ppt"}  # 需要 Office 转换的旧二进制格式


@dataclass
class DocFile:
    path: str
    ext: str
    family: str          # word / excel / ppt / text
    legacy: bool = False # 是否旧格式（需转换）

    @property
    def name(self) -> str:
        return os.path.basename(self.path)


@dataclass
class ExtractResult:
    doc: DocFile
    text: str = ""
    tables: list[list[list[str]]] = field(default_factory=list)  # 每张表 = 二维字符串
    error: str = ""


def scan(folder: str, recursive: bool = True) -> list[DocFile]:
    """扫描文件夹下的受支持文档。"""
    found: list[DocFile] = []
    if not os.path.isdir(folder):
        return found
    walker = os.walk(folder) if recursive else [(folder, [], os.listdir(folder))]
    for root, _dirs, files in walker:
        for fn in files:
            ext = os.path.splitext(fn)[1].lower()
            family = _EXT_FAMILY.get(ext)
            if not family:
                continue
            full = os.path.join(root, fn)
            found.append(DocFile(path=full, ext=ext, family=family,
                                 legacy=ext in _LEGACY))
    found.sort(key=lambda d: d.path.lower())
    return found


# ---------- 各格式文本提取 ----------

def _extract_docx(path: str) -> tuple[str, list[list[list[str]]]]:
    import docx  # python-docx

    d = docx.Document(path)
    lines: list[str] = []
    for p in d.paragraphs:
        lines.append(p.text)
    # 表格：既收集纯文本，也保留成结构化二维数组
    tables: list[list[list[str]]] = []
    for t in d.tables:
        grid = [[cell.text.strip() for cell in row.cells] for row in t.rows]
        tables.append(grid)
    return "\n".join(lines), tables


def _extract_xlsx(path: str) -> tuple[str, list[list[list[str]]]]:
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    tables: list[list[list[str]]] = []
    text_lines: list[str] = []
    try:
        for ws in wb.worksheets:
            grid: list[list[str]] = []
            for row in ws.iter_rows(values_only=True):
                vals = ["" if v is None else str(v) for v in row]
                if any(cell.strip() for cell in vals):
                    grid.append(vals)
            if grid:
                tables.append(grid)
                for r in grid:
                    text_lines.append("\t".join(r))
    finally:
        wb.close()
    return "\n".join(text_lines), tables


def _extract_pptx(path: str) -> tuple[str, list[list[list[str]]]]:
    from pptx import Presentation

    prs = Presentation(path)
    lines: list[str] = []
    tables: list[list[list[str]]] = []
    for slide_no, slide in enumerate(prs.slides, 1):
        lines.append("【第 %d 页】" % slide_no)
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    t = "".join(run.text for run in para.runs)
                    if t.strip():
                        lines.append(t)
            if getattr(shape, "has_table", False):
                tbl = shape.table
                grid = [[cell.text.strip() for cell in row.cells] for row in tbl.rows]
                tables.append(grid)
    return "\n".join(lines), tables


def _extract_textfile(path: str) -> tuple[str, list[list[list[str]]]]:
    for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read(), []
        except (UnicodeDecodeError, LookupError):
            continue
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read(), []


def extract_file(doc: DocFile) -> ExtractResult:
    """提取单个文件的文本和表格。旧格式会先尝试转换。"""
    res = ExtractResult(doc=doc)
    target = doc.path
    tmp = None
    try:
        if doc.legacy:
            from . import office_convert
            tmp = office_convert.convert(doc.path, doc.family)
            if tmp is None:
                res.error = "旧格式 .%s 需要本机安装 Office/WPS 才能转换，已跳过" % doc.ext.lstrip(".")
                return res
            target = tmp

        if doc.family == "word":
            res.text, res.tables = _extract_docx(target)
        elif doc.family == "excel":
            res.text, res.tables = _extract_xlsx(target)
        elif doc.family == "ppt":
            res.text, res.tables = _extract_pptx(target)
        else:
            res.text, res.tables = _extract_textfile(target)
    except Exception as exc:  # 单个文件失败不应中断批处理
        res.error = "%s: %s" % (type(exc).__name__, exc)
    finally:
        if tmp and os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    return res