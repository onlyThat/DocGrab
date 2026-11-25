"""旧格式转换测试：用本机 Office 把 .docx/.xlsx 另存为 .doc/.xls，
再用 extract 解析，验证 COM 转换 + 提取链路。无 Office 则跳过。
"""
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docgrab import extract, office_convert

HERE = os.path.dirname(os.path.abspath(__file__))
LEGACY = os.path.join(HERE, "_legacy")
_fail = 0


def check(cond, name):
    global _fail
    print(("[PASS] " if cond else "[FAIL] ") + name)
    if not cond:
        _fail += 1


def _make_legacy():
    """用 Office 生成一份 .doc 和 .xls。"""
    import pythoncom
    import win32com.client as win32

    os.makedirs(LEGACY, exist_ok=True)
    pythoncom.CoInitialize()
    try:
        # Word -> .doc
        w = win32.DispatchEx("Word.Application")
        w.Visible = False
        w.DisplayAlerts = 0
        d = w.Documents.Add()
        d.Content.Text = "甲方：测试公司\n合同编号：LEGACY-001\n金额 12345 元\n日期 2023年1月1日"
        doc_path = os.path.join(LEGACY, "旧合同.doc")
        d.SaveAs2(os.path.abspath(doc_path), FileFormat=0)  # wdFormatDocument = .doc
        d.Close(SaveChanges=0)
        w.Quit()
        # Excel -> .xls
        e = win32.DispatchEx("Excel.Application")
        e.Visible = False
        e.DisplayAlerts = False
        wb = e.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Cells(1, 1).Value = "姓名"; ws.Cells(1, 2).Value = "金额"
        ws.Cells(2, 1).Value = "旧表人"; ws.Cells(2, 2).Value = "777"
        xls_path = os.path.join(LEGACY, "旧表格.xls")
        wb.SaveAs(os.path.abspath(xls_path), FileFormat=56)  # xlExcel8 = .xls
        wb.Close(SaveChanges=False)
        e.Quit()
    finally:
        pythoncom.CoUninitialize()


def main():
    if not office_convert.is_available():
        print("[SKIP] 未安装 pywin32/Office，跳过旧格式测试")
        return 0
    if os.path.isdir(LEGACY):
        shutil.rmtree(LEGACY)
    try:
        _make_legacy()
    except Exception as exc:
        print("[SKIP] 无法用 Office 生成旧格式样例：%s" % exc)
        return 0

    docs = extract.scan(LEGACY)
    check(len(docs) == 2, "扫描到 2 个旧格式文件 (%d)" % len(docs))
    check(all(d.legacy for d in docs), "都被标记为旧格式")

    for d in docs:
        r = extract.extract_file(d)
        print("   %s -> 文本 %d 字，表 %d 张，错误=%s" % (d.name, len(r.text), len(r.tables), r.error or "无"))
        check(not r.error, "%s 转换并提取成功" % d.name)

    doc_res = next((r for r in (extract.extract_file(d) for d in docs) if d.family == "word" or True), None)

    print()
    print("全部通过" if _fail == 0 else "%d 项失败" % _fail)
    return 0 if _fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())