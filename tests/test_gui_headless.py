"""GUI 无头测试：构造窗口和三个页签，喂数据，验证不报错、数据正确。

在 Windows 上 tkinter 需要真实显示器；这里用 withdrawn 隐藏窗口，
不进入 mainloop，只验证构造与数据流。运行：python tests/test_gui_headless.py
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import tkinter as tk  # noqa: E402

from docgrab import extract, gui  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLES = os.path.join(HERE, "samples")
_fail = 0


def check(cond, name):
    global _fail
    print(("[PASS] " if cond else "[FAIL] ") + name)
    if not cond:
        _fail += 1


def main():
    root = gui.TkinterDnD.Tk() if gui._HAS_DND else tk.Tk()  # noqa: SLF001
    root.withdraw()
    try:
        app = gui.App(root)
        check(app is not None, "App 构造成功")
        check(app.tab_text is not None and app.tab_fields is not None and app.tab_merge is not None,
              "三个页签都存在")

        # 直接喂提取结果（跳过后台线程）
        docs = extract.scan(SAMPLES)
        results = [extract.extract_file(d) for d in docs]
        app.results = results
        app.tab_text.set_results(results)
        app.tab_fields.set_results(results)
        app.tab_merge.set_results(results)
        root.update_idletasks()

        check(len(app.tab_text.tree.get_children()) == len(results), "全文页签列表行数正确")
        check(len(app.tab_fields.rows) == len(results), "字段页签生成 %d 行" % len(app.tab_fields.rows))
        # 3 张 Excel + 3 份 Word 合同里的表格，都会被收集（Word 表格也可合并）
        check(len(app.tab_merge.sources) == 6,
              "表格页签收集到 6 张表（3 Excel + 3 Word，实际 %d）" % len(app.tab_merge.sources))
        check(len(app.tab_merge.col_vars) >= 4, "可用列被列出 (%d 个)" % len(app.tab_merge.col_vars))

        # 触发合并预览
        app.tab_merge.preview_merge()
        root.update_idletasks()
        check(app.tab_merge.preview_tree is not None, "合并预览表已渲染")

        # 默认输出目录合理
        d = app.default_out_dir()
        check(d.endswith("DocGrab_输出"), "默认输出目录名正确: %s" % d)
    except Exception as exc:
        import traceback
        traceback.print_exc()
        check(False, "GUI 无头测试异常: %s" % exc)
    finally:
        try:
            root.destroy()
        except Exception:
            pass

    print()
    print("全部通过" if _fail == 0 else "%d 项失败" % _fail)
    return 0 if _fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())