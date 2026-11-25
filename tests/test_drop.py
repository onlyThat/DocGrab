"""验证拖拽路径解析 + DropZone 加载逻辑（用真实路径串，不依赖真实鼠标拖拽）。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import tkinter as tk
from docgrab import gui, extract

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLES = os.path.join(HERE, "samples")
_fail = 0


def check(cond, name):
    global _fail
    print(("[PASS] " if cond else "[FAIL] ") + name)
    if not cond:
        _fail += 1


def main():
    # 1) 路径解析：带花括号、带空格、多个
    check(gui._parse_drop("{C:/a b/c}") == ["C:/a b/c"], "带空格花括号路径解析")
    check(gui._parse_drop("C:/x C:/y") == ["C:/x", "C:/y"], "多路径空格分隔解析")
    check(gui._parse_drop("{C:/a b} C:/c") == ["C:/a b", "C:/c"], "混合解析")

    root = gui.TkinterDnD.Tk() if gui._HAS_DND else tk.Tk()
    root.withdraw()
    try:
        app = gui.App(root)

        # 2) 模拟拖入文件夹：直接放到 on_drop 等价路径
        loaded = {}
        orig = app.load_folder
        app.load_folder = lambda folder: (loaded.setdefault("folder", folder), orig(folder))
        app.drop.on_folder = app.load_folder
        app.drop._on_drop(type("E", (), {"data": "{%s}" % SAMPLES})())
        check(loaded.get("folder") == SAMPLES, "拖入文件夹被正确接收: %s" % loaded.get("folder"))

        # 等待后台提取完成
        import time
        for _ in range(100):
            root.update()
            if not app.busy:
                break
            time.sleep(0.05)
        check(not app.busy, "后台提取已完成")
        check(len(app.results) >= 7, "提取到结果 (%d)" % len(app.results))
        check(len(app.tab_merge.sources) >= 3, "合并页签有数据 (%d 表)" % len(app.tab_merge.sources))
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