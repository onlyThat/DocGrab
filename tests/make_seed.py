"""生成一批用于端到端测试的示例文档（放在 tests/samples/ 下）。

包含：3 份 Word 合同（含金额/编号/日期/甲乙方）、3 份 Excel（同名列但列顺序不同）、
1 份 PPT、1 份 txt。可重复运行，每次覆盖。
"""
from __future__ import annotations

import os
import shutil

import docx
import openpyxl
from pptx import Presentation
from pptx.util import Inches

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLES = os.path.join(HERE, "samples")


def _word_contracts():
    d = os.path.join(SAMPLES, "合同")
    os.makedirs(d, exist_ok=True)
    contracts = [
        ("甲方：杭州某某科技有限公司\n乙方：上海测试贸易有限公司\n"
         "合同编号：HT-2024-001\n签订日期：2024年3月15日\n"
         "经双方友好协商，本合同总金额为人民币 128,500.00 元。", "合同A.docx"),
        ("甲方：北京示例集团\n乙方：深圳方案提供商\n"
         "合同编号：HT-2024-002\n签订日期：2024-06-01\n"
         "服务费用总计 56,000 元，另有押金 5,000 元。", "合同B.docx"),
        ("甲方：广州甲方有限公司\n乙方：成都乙方工作室\n"
         "合同号：HT2024-003\n签订日期：2024/09/20\n"
         "本项目报价为 18.8 万元（含税）。", "合同C.docx"),
    ]
    for text, name in contracts:
        doc = docx.Document()
        doc.add_heading("采购合同", level=1)
        for line in text.split("\n"):
            doc.add_paragraph(line)
        # 加一张小表格，测试表格提取
        t = doc.add_table(rows=2, cols=3)
        t.style = "Table Grid"
        t.rows[0].cells[0].text = "项目"
        t.rows[0].cells[1].text = "数量"
        t.rows[0].cells[2].text = "金额"
        t.rows[1].cells[0].text = "服务费"
        t.rows[1].cells[1].text = "1"
        t.rows[1].cells[2].text = "详见正文"
        doc.save(os.path.join(d, name))


def _excel_tables():
    d = os.path.join(SAMPLES, "表格")
    os.makedirs(d, exist_ok=True)
    # 三份表：列顺序不同，但都有 姓名/部门/金额 三个同名列
    data_orders = [
        (["姓名", "部门", "金额"], [["张三", "财务", "5000"], ["李四", "销售", "8000"]]),
        (["部门", "姓名", "金额"], [["技术", "王五", "12000"], ["技术", "赵六", "9500"]]),
        (["金额", "部门", "姓名", "备注"], [["3000", "行政", "孙七", "报销"], ["2000", "行政", "周八", ""]]),
    ]
    for i, (header, rows) in enumerate(data_orders, 1):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "明细"
        ws.append(header)
        for r in rows:
            ws.append(r)
        wb.save(os.path.join(d, "表格%d.xlsx" % i))


def _ppt():
    d = os.path.join(SAMPLES, "演示")
    os.makedirs(d, exist_ok=True)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "季度汇报"
    slide.placeholders[1].text = "本季度销售额 320000 元\n合同编号：QR-2024-Q3"
    prs.save(os.path.join(d, "汇报.pptx"))


def _text():
    d = os.path.join(SAMPLES, "杂项")
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, "说明.txt"), "w", encoding="utf-8") as f:
        f.write("这是纯文本，金额 66.6 元，日期 2024-12-31。\n")


def main():
    if os.path.isdir(SAMPLES):
        shutil.rmtree(SAMPLES)
    os.makedirs(SAMPLES, exist_ok=True)
    _word_contracts()
    _excel_tables()
    _ppt()
    _text()
    print("示例文档已生成：", SAMPLES)


if __name__ == "__main__":
    main()