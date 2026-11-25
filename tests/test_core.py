"""核心引擎断言测试。运行：python tests/test_core.py"""
from __future__ import annotations

import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docgrab import extract, fields, export  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLES = os.path.join(HERE, "samples")
_fail = 0


def check(cond, name):
    global _fail
    print(("[PASS] " if cond else "[FAIL] ") + name)
    if not cond:
        _fail += 1


def main():
    docs = extract.scan(SAMPLES)
    check(len(docs) >= 7, "扫描到至少 7 个文件 (实际 %d)" % len(docs))

    fams = {d.family for d in docs}
    check({"word", "excel", "ppt", "text"} <= fams, "四种格式都被识别: %s" % fams)

    results = [extract.extract_file(d) for d in docs]
    errs = [r for r in results if r.error]
    check(not errs, "全部文件提取无错误 (%d 个错误)" % len(errs))
    if errs:
        for e in errs:
            print("   错误：", e.doc.name, e.error)

    # Word 文本
    word = next(r for r in results if r.doc.name == "合同A.docx")
    check("128,500.00" in word.text, "Word 正文金额被提取")
    check(len(word.tables) >= 1, "Word 表格被提取 (%d 张)" % len(word.tables))

    # 字段抓取
    f = fields.extract_fields(word.text)
    check(any("128500" in a for a in f["amount"]), "金额字段命中 128500: %s" % f["amount"])
    check(any("HT-2024-001" in c for c in f["contract_no"]), "合同编号命中")
    check(any("2024" in d for d in f["date"]), "日期命中")
    check(any("杭州" in p for p in f["party"]), "甲方命中")

    # 万元换算
    wc = next(r for r in results if r.doc.name == "合同C.docx")
    fc = fields.extract_fields(wc.text)
    check(any("万元" in a and "188000" in a for a in fc["amount"]), "万元换算正确: %s" % fc["amount"])

    # Excel 列合并
    sources = export.collect_tables([r for r in results if r.doc.family == "excel"])
    check(len(sources) == 3, "收集到 3 张 Excel 表 (%d)" % len(sources))
    cols = export.available_columns(sources)
    for need in ("姓名", "部门", "金额"):
        check(need in cols, "可选中列包含 %s" % need)

    merged = export.merge_tables(sources, ["姓名", "金额"])
    check(len(merged) == 6, "合并后 6 行数据 (%d)" % len(merged))
    check(all("姓名" in m and "金额" in m for m in merged), "合并行含所选列")
    check(all("部门" not in m for m in merged), "未选列未混入")

    # 导出到临时目录
    with tempfile.TemporaryDirectory() as tmp:
        txt_dir = os.path.join(tmp, "txt")
        files = export.export_text_per_file(results, txt_dir)
        check(len(files) == len([r for r in results if r.text]), "每份文档导出 txt (%d)" % len(files))

        merged_txt = export.export_text_merged(results, os.path.join(tmp, "all.txt"))
        check(os.path.getsize(merged_txt) > 0, "合并 txt 非空")

        sum_path = export.export_rows([fields.summarize_row(r) for r in results],
                                      os.path.join(tmp, "汇总.xlsx"))
        check(os.path.exists(sum_path), "字段汇总 xlsx 已生成")

        mp, n = export.merge_tables_xlsx(sources, ["姓名", "部门", "金额"],
                                         os.path.join(tmp, "合并.xlsx"))
        check(n == 6 and os.path.exists(mp), "列合并 xlsx 生成 (%d 行)" % n)

        # 验证 xlsx 能读回
        import openpyxl
        wb = openpyxl.load_workbook(mp)
        ws = wb.active
        check(ws.max_row == 1 + 6, "合并表 1 表头 + 6 行数据")
        check(ws.cell(row=1, column=1).value == "来源文件", "首列为来源文件")
        wb.close()

    print()
    print("全部通过" if _fail == 0 else "%d 项失败" % _fail)
    return 0 if _fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())