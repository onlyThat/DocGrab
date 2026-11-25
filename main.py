"""DocGrab 启动入口。

双击 exe 或运行 `python main.py` 即可打开图形界面。
会先把任何未捕获异常写进同目录的 docgrab-error.log，避免 windowed
模式下崩溃时"窗口一闪而过、什么提示都没有"。
"""
import os
import sys
import traceback


def _log_path():
    # 优先写到 exe 所在目录；不可写时退回用户目录
    base = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
    p = os.path.join(base, "docgrab-error.log")
    try:
        with open(p, "a", encoding="utf-8"):
            pass
        return p
    except OSError:
        return os.path.join(os.path.expanduser("~"), "docgrab-error.log")


def _install_crash_logger():
    def hook(exc_type, exc, tb):
        text = "".join(traceback.format_exception(exc_type, exc, tb))
        try:
            with open(_log_path(), "a", encoding="utf-8") as f:
                f.write(text + "\n" + "-" * 60 + "\n")
        except Exception:
            pass
        # 尽力弹个框，windowed 模式下用户至少知道出事了
        try:
            import tkinter as tk
            import tkinter.messagebox as mb
            r = tk.Tk()
            r.withdraw()
            mb.showerror("DocGrab 启动失败",
                         "程序遇到错误：\n%s\n\n详细信息已写入：\n%s" % (exc, _log_path()))
            r.destroy()
        except Exception:
            pass

    sys.excepthook = hook


# 让 `python main.py` 与打包后的 exe 都能找到 docgrab 包
if getattr(sys, "frozen", False):
    sys.path.insert(0, os.path.dirname(sys.executable))
else:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if __name__ == "__main__":
    _install_crash_logger()
    try:
        from docgrab.gui import main
        main()
    except Exception:
        exc_type, exc, tb = sys.exc_info()
        sys.excepthook(exc_type, exc, tb)