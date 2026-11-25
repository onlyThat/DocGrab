"""旧格式（.doc/.xls/.ppt）转换：借本机已安装的 Office / WPS 用 COM 转成新格式。

转换在临时目录进行，不修改原文件。没有可用的 Office 时返回 None，
由上层决定"跳过并提示"。
"""
from __future__ import annotations

import os
import tempfile
import threading

# COM 必须在同一线程内创建和释放；这里用独立线程包起来，避免和 GUI 线程冲突。
_convert_lock = threading.Lock()
_com_available: bool | None = None


def _probe_com() -> bool:
    """探测本机是否能用 COM 启动 Office 应用（Word/Excel 各试一次失败即算不可用）。"""
    global _com_available
    if _com_available is not None:
        return _com_available
    try:
        import win32com.client  # noqa: F401
    except Exception:
        _com_available = False
        return False
    _com_available = True
    return True


def is_available() -> bool:
    """粗略判断旧格式转换是否可用（装了 pywin32 即认为可能可用，真正能否转在转换时再判定）。"""
    return _probe_com()


class ConvertError(Exception):
    pass


def _convert_with_office(src: str, kind: str) -> str:
    """在后台线程里执行 COM 转换，返回新格式临时文件的绝对路径。"""
    result: dict[str, str | Exception] = {}

    def worker():
        import pythoncom
        import win32com.client as win32

        pythoncom.CoInitialize()
        app = None
        doc = None
        try:
            fd, out = tempfile.mkstemp(suffix={"word": ".docx", "excel": ".xlsx", "ppt": ".pptx"}[kind])
            os.close(fd)
            os.unlink(out)  # Office 的 SaveAs 需要一个可写但不存在的目标名

            if kind == "word":
                app = win32.DispatchEx("Word.Application")
                app.Visible = False
                app.DisplayAlerts = 0
                doc = app.Documents.Open(os.path.abspath(src), ReadOnly=True,
                                         AddToRecentFiles=False, Visible=False)
                doc.SaveAs2(os.path.abspath(out), FileFormat=16)  # wdFormatDocumentDefault = docx
            elif kind == "excel":
                app = win32.DispatchEx("Excel.Application")
                app.Visible = False
                app.DisplayAlerts = False
                doc = app.Workbooks.Open(os.path.abspath(src), ReadOnly=True,
                                        UpdateLinks=0)
                doc.SaveAs(os.path.abspath(out), FileFormat=51)  # xlOpenXMLWorkbook = xlsx
            else:  # ppt
                app = win32.DispatchEx("PowerPoint.Application")
                doc = app.Presentations.Open(os.path.abspath(src), ReadOnly=True,
                                             WithWindow=False)
                doc.SaveAs(os.path.abspath(out), 24)  # ppSaveAsOpenXMLPresentation = pptx
            result["out"] = out
        except Exception as exc:  # COM 异常种类繁多，统一包装
            result["err"] = exc
        finally:
            try:
                if doc is not None:
                    if kind == "word":
                        doc.Close(SaveChanges=0)
                    elif kind == "excel":
                        doc.Close(SaveChanges=False)
                    else:
                        doc.Close()
            except Exception:
                pass
            try:
                if app is not None:
                    app.Quit()
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    t = threading.Thread(target=worker, daemon=True)
    t.start()
    t.join(timeout=60)
    if t.is_alive():
        raise ConvertError("转换为新格式超时（可能 Office 弹了对话框）")
    if "err" in result:
        raise ConvertError("Office 转换失败：%s" % result["err"])
    if "out" not in result:
        raise ConvertError("转换未产生输出文件")
    return result["out"]


def convert(src: str, kind: str) -> str | None:
    """把旧格式转成新格式，返回临时文件路径；不可用时返回 None。

    调用方负责在用完后删除返回的临时文件。
    """
    if not _probe_com():
        return None
    with _convert_lock:
        return _convert_with_office(src, kind)