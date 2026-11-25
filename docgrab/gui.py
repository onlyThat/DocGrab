"""DocGrab 图形界面：拖入文件夹 → 自动提取 → 预览 → 导出。

三个页签：
  1. 全文提取   —— 每份文档一个 txt / 合并成一个大 txt
  2. 字段汇总   —— 抓金额、合同号、日期、甲乙方，一行一份文档
  3. 表格合并   —— 多个 Excel/Word 表格按列名对齐，合并成一张大表
"""
from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import export, extract, fields

APP_TITLE = "DocGrab · 文档批量提取"

# 拖拽支持是可选增强：装了 tkinterdnd2 就能拖文件夹，否则退回"选择文件夹"按钮
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _HAS_DND = True
except Exception:  # pragma: no cover
    _HAS_DND = False


def _parse_drop(data: str) -> list[str]:
    """解析 tkdnd 拖入路径串（可能带 {} 包裹和空格）。"""
    paths: list[str] = []
    i = 0
    n = len(data)
    while i < n:
        if data[i] == "{":
            j = data.find("}", i + 1)
            if j == -1:
                break
            paths.append(data[i + 1:j])
            i = j + 1
        elif data[i] == " ":
            i += 1
        else:
            j = i
            while j < n and data[j] != " ":
                j += 1
            paths.append(data[i:j])
            i = j
    return paths


class DropZone(ttk.Frame):
    """一个视觉上的拖拽区 + 选择按钮。"""

    def __init__(self, master, on_folder, on_files):
        super().__init__(master, padding=12)
        self.on_folder = on_folder
        self.on_files = on_files

        self.box = tk.Label(
            self, text="把文件夹拖到这里，或点击右侧按钮选择\n（支持 .docx .doc .xlsx .xls .pptx .ppt .txt .csv .md）",
            relief="groove", bd=2, height=3, bg="#f4f7fb", fg="#33506e",
            font=("Microsoft YaHei UI", 10),
        )
        self.box.pack(fill="x", expand=False, ipady=10)

        btns = ttk.Frame(self)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="选择文件夹…", command=self._pick_folder).pack(side="left")
        ttk.Button(btns, text="选择文件…", command=self._pick_files).pack(side="left", padx=6)
        self.hint = ttk.Label(btns, text="会递归扫描子文件夹", foreground="#888")
        self.hint.pack(side="left", padx=8)

        if _HAS_DND:
            for w in (self.box, self):
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:
                    pass
        else:
            self.box.configure(text="点击下方按钮选择文件夹\n（安装 tkinterdnd2 后可直接拖拽）")

    def _on_drop(self, event):
        paths = _parse_drop(event.data)
        folders = [p for p in paths if os.path.isdir(p)]
        files = [p for p in paths if os.path.isfile(p)]
        if folders:
            self.on_folder(folders[0])
        elif files:
            self.on_files(files)

    def _pick_folder(self):
        d = filedialog.askdirectory(title="选择包含文档的文件夹")
        if d:
            self.on_folder(d)

    def _pick_files(self):
        fs = filedialog.askopenfilenames(
            title="选择文档",
            filetypes=[("Office 文档", "*.docx *.doc *.xlsx *.xls *.pptx *.ppt *.txt *.csv *.md"),
                       ("所有文件", "*.*")])
        if fs:
            self.on_files(list(fs))


class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("980x720")
        root.minsize(820, 600)
        try:
            root.iconbitmap(default="")  # 打包时可换成 .ico
        except Exception:
            pass

        self.drop = DropZone(root, self.load_folder, self.load_files)
        self.drop.pack(fill="x", padx=10, pady=(10, 4))

        self.summary_var = tk.StringVar(value="尚未加载文档。")
        ttk.Label(root, textvariable=self.summary_var, foreground="#333").pack(fill="x", padx=16)

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=8)

        self.tab_text = TextTab(self.notebook)
        self.tab_fields = FieldsTab(self.notebook)
        self.tab_merge = MergeTab(self.notebook)
        self.notebook.add(self.tab_text, text="  全文提取  ")
        self.notebook.add(self.tab_fields, text="  字段汇总（Word/合同）  ")
        self.notebook.add(self.tab_merge, text="  表格合并（Excel）  ")

        # 后台状态
        self.results: list[extract.ExtractResult] = []
        self.docs: list[extract.DocFile] = []
        self.busy = False
        self.progress = ttk.Progressbar(root, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(0, 4))
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(root, textvariable=self.status_var, foreground="#666",
                  anchor="w").pack(fill="x", padx=14, pady=(0, 8))

    # ---------- 加载 ----------

    def load_folder(self, folder):
        if self.busy:
            return
        self.docs = extract.scan(folder, recursive=True)
        self._start_extract()

    def load_files(self, files):
        if self.busy:
            return
        docs = []
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            family = extract._EXT_FAMILY.get(ext)  # noqa: SLF001
            if family:
                docs.append(extract.DocFile(path=f, ext=ext, family=family,
                                            legacy=ext in extract._LEGACY))
        self.docs = sorted(docs, key=lambda d: d.path.lower())
        self._start_extract()

    def _start_extract(self):
        if not self.docs:
            self.summary_var.set("该位置没有找到受支持的文档。")
            self._reset_tabs()
            return
        self.busy = True
        self.results = []
        self.progress.configure(maximum=len(self.docs), value=0)
        self.status_var.set("开始提取 %d 个文档…" % len(self.docs))
        self.summary_var.set("正在提取，请稍候…")
        self._reset_tabs()

        q: queue.Queue = queue.Queue()
        docs = list(self.docs)

        def worker():
            out = []
            for i, d in enumerate(docs, 1):
                res = extract.extract_file(d)
                out.append(res)
                q.put(("progress", i, len(docs), res))
            q.put(("done", out))

        threading.Thread(target=worker, daemon=True).start()
        self._poll(q)

    def _poll(self, q):
        try:
            while True:
                msg = q.get_nowait()
                if msg[0] == "progress":
                    _, i, total, res = msg
                    self.progress.configure(value=i)
                    self.status_var.set("已完成 %d/%d：%s" % (i, total, res.doc.name))
                elif msg[0] == "done":
                    self._finish(msg[1])
                    return
        except queue.Empty:
            pass
        self.root.after(80, lambda: self._poll(q))

    def _finish(self, results):
        self.results = results
        self.busy = False
        errs = [r for r in results if r.error]
        fams = {}
        for r in results:
            fams[r.doc.family] = fams.get(r.doc.family, 0) + 1
        name_map = {"word": "Word", "excel": "Excel", "ppt": "PPT", "text": "文本"}
        parts = ["%s %d" % (name_map.get(k, k), v) for k, v in sorted(fams.items())]
        msg = "共 %d 个文档（%s）" % (len(results), "、".join(parts) if parts else "无")
        if errs:
            msg += "，其中 %d 个提取失败" % len(errs)
        self.summary_var.set(msg)
        self.status_var.set("提取完成。" + ("有 %d 个文件出错，详见各页签。" % len(errs) if errs else ""))

        self.tab_text.set_results(results)
        self.tab_fields.set_results(results)
        self.tab_merge.set_results(results)

    def _reset_tabs(self):
        self.tab_text.set_results([])
        self.tab_fields.set_results([])
        self.tab_merge.set_results([])

    def default_out_dir(self):
        if self.docs:
            base = os.path.dirname(self.docs[0].path)
            return os.path.join(base, "DocGrab_输出")
        return os.path.join(os.path.expanduser("~"), "DocGrab_输出")


def _make_tree(master, columns, width_map=None):
    frame = ttk.Frame(master)
    tree = ttk.Treeview(frame, columns=columns, show="headings", height=12)
    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)
    for c in columns:
        tree.heading(c, text=c)
        w = (width_map or {}).get(c, 120)
        tree.column(c, width=w, anchor="w", stretch=True)
    return frame, tree


# ====================== 页签 1：全文提取 ======================

class TextTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=10)
        self.results: list[extract.ExtractResult] = []

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True)

        left = ttk.Frame(body)
        left.pack(side="left", fill="both", expand=True)
        ttk.Label(left, text="文档列表（点击查看提取的纯文本）").pack(anchor="w")
        f, self.tree = _make_tree(left, ["文件", "类型", "字数", "状态"],
                                  {"文件": 320, "类型": 70, "字数": 80, "状态": 200})
        f.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        right = ttk.Frame(body)
        right.pack(side="left", fill="both", expand=True, padx=(10, 0))
        ttk.Label(right, text="预览").pack(anchor="w")
        self.preview = tk.Text(right, wrap="word", height=18, font=("Consolas", 10))
        pv = ttk.Scrollbar(right, command=self.preview.yview)
        self.preview.configure(yscrollcommand=pv.set)
        self.preview.pack(side="left", fill="both", expand=True)
        pv.pack(side="right", fill="y")

        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(8, 0))
        ttk.Button(bar, text="每篇导出为 txt…", command=self.export_each).pack(side="left")
        ttk.Button(bar, text="合并成一个 txt…", command=self.export_merged).pack(side="left", padx=6)

    def set_results(self, results):
        self.results = results
        self.tree.delete(*self.tree.get_children())
        for r in results:
            status = ("失败：" + r.error) if r.error else "成功"
            self.tree.insert("", "end", values=(r.doc.name, r.doc.ext, len(r.text), status))
        self.preview.delete("1.0", "end")

    def _on_select(self, _e):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if 0 <= idx < len(self.results):
            r = self.results[idx]
            self.preview.delete("1.0", "end")
            self.preview.insert("1.0", r.text if not r.error else "【提取失败】\n" + r.error)

    def export_each(self):
        if not self._guard():
            return
        out = filedialog.askdirectory(title="选择导出 txt 的目标文件夹")
        if not out:
            return
        files = export.export_text_per_file(self.results, out)
        messagebox.showinfo("完成", "已导出 %d 个 txt 到：\n%s" % (len(files), out))

    def export_merged(self):
        if not self._guard():
            return
        path = filedialog.asksaveasfilename(title="保存合并文本", defaultextension=".txt",
                                            filetypes=[("文本文件", "*.txt")],
                                            initialfile="合并全文.txt")
        if not path:
            return
        export.export_text_merged(self.results, path)
        messagebox.showinfo("完成", "已合并导出到：\n%s" % path)

    def _guard(self):
        if not self.results:
            messagebox.showwarning("提示", "请先加载并提取文档。")
            return False
        return True


# ====================== 页签 2：字段汇总 ======================

class FieldsTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=10)
        self.results: list[extract.ExtractResult] = []
        self.rows: list[dict[str, str]] = []

        opts = ttk.Frame(self)
        opts.pack(fill="x")
        ttk.Label(opts, text="汇总一份文档一行，自动抓取：金额、合同/编号、日期、甲乙方。").pack(side="left")
        ttk.Button(opts, text="重新抓取", command=self.rebuild).pack(side="right")

        cols = ["文件"] + [r.label for r in fields.DEFAULT_RULES] + ["错误"]
        self.columns = cols
        f, self.tree = _make_tree(self, cols,
                                  {c: 150 for c in cols} | {"文件": 240, "金额": 220, "甲乙方": 240, "错误": 240})
        f.pack(fill="both", expand=True, pady=(8, 0))

        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(8, 0))
        ttk.Button(bar, text="导出汇总表（xlsx）…", command=lambda: self.export(".xlsx")).pack(side="left")
        ttk.Button(bar, text="导出汇总表（csv）…", command=lambda: self.export(".csv")).pack(side="left", padx=6)
        self.count_var = tk.StringVar()
        ttk.Label(bar, textvariable=self.count_var, foreground="#666").pack(side="right")

    def set_results(self, results):
        self.results = results
        self.rebuild()

    def rebuild(self):
        self.tree.delete(*self.tree.get_children())
        self.rows = [fields.summarize_row(r) for r in self.results]
        for row in self.rows:
            vals = [row.get(c, "") for c in self.columns]
            self.tree.insert("", "end", values=vals)
        with_amount = sum(1 for r in self.rows if r.get("金额"))
        self.count_var.set("共 %d 份，%d 份抓到金额" % (len(self.rows), with_amount))

    def export(self, ext):
        if not self.rows:
            messagebox.showwarning("提示", "请先加载并提取文档。")
            return
        path = filedialog.asksaveasfilename(
            title="保存汇总表", defaultextension=ext,
            filetypes=[("Excel", "*.xlsx")] if ext == ".xlsx" else [("CSV", "*.csv")],
            initialfile="字段汇总" + ext)
        if not path:
            return
        try:
            export.export_rows(self.rows, path, self.columns)
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))
            return
        messagebox.showinfo("完成", "已导出到：\n%s" % path)


# ====================== 页签 3：表格合并 ======================

class MergeTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=10)
        self.results: list[extract.ExtractResult] = []
        self.sources: list[export.TableSource] = []
        self.col_vars: dict[str, tk.BooleanVar] = {}

        top = ttk.Frame(self)
        top.pack(fill="x")
        ttk.Label(top, text="扫描到的列（勾选要合并的列，按列名跨文件对齐）：").pack(side="left")

        mid = ttk.Frame(self)
        mid.pack(fill="both", expand=True, pady=(6, 0))
        # 左：列勾选
        leftf = ttk.LabelFrame(mid, text="可用列", padding=6)
        leftf.pack(side="left", fill="y")
        self.col_frame = ttk.Frame(leftf)
        self.col_frame.pack(fill="both", expand=True)
        cbar = ttk.Frame(leftf)
        cbar.pack(fill="x", pady=(6, 0))
        ttk.Button(cbar, text="全选", width=6, command=lambda: self._set_all(True)).pack(side="left")
        ttk.Button(cbar, text="全不选", width=8, command=lambda: self._set_all(False)).pack(side="left", padx=4)

        # 右：预览 + 来源表
        right = ttk.Frame(mid)
        right.pack(side="left", fill="both", expand=True, padx=(10, 0))
        ttk.Label(right, text="数据源表格").pack(anchor="w")
        sf, self.src_tree = _make_tree(right, ["文件", "表序号", "行数", "列数"],
                                       {"文件": 260, "表序号": 70, "行数": 70, "列数": 70})
        sf.pack(fill="both", expand=False)
        self.add_src_col = tk.BooleanVar(value=True)
        ttk.Checkbutton(right, text="合并结果首列加入「来源文件」", variable=self.add_src_col).pack(anchor="w", pady=6)

        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(8, 0))
        ttk.Button(bar, text="预览合并结果", command=self.preview_merge).pack(side="left")
        ttk.Button(bar, text="导出合并表（xlsx）…", command=self.export_merge).pack(side="left", padx=6)
        self.merge_info = tk.StringVar()
        ttk.Label(bar, textvariable=self.merge_info, foreground="#666").pack(side="right")

        pf = ttk.LabelFrame(self, text="合并预览", padding=4)
        pf.pack(fill="both", expand=True, pady=(8, 0))
        self.preview_frame_holder = pf
        self.preview_tree = None

    def set_results(self, results):
        self.results = results
        self.sources = export.collect_tables(results)
        # 只把"有表格"的结果纳入
        tables = [r for r in results if r.tables]
        self.sources = export.collect_tables(tables)

        for w in self.col_frame.winfo_children():
            w.destroy()
        self.col_vars.clear()
        for i, col in enumerate(export.available_columns(self.sources)):
            var = tk.BooleanVar(value=True)
            self.col_vars[col] = var
            ttk.Checkbutton(self.col_frame, text=col, variable=var).grid(
                row=i, column=0, sticky="w")

        self.src_tree.delete(*self.src_tree.get_children())
        for s in self.sources:
            self.src_tree.insert("", "end", values=(s.file_name, s.sheet_index + 1,
                                                    len(s.rows), len(s.header)))
        self.merge_info.set("共 %d 张表" % len(self.sources))

    def _set_all(self, val):
        for v in self.col_vars.values():
            v.set(val)

    def _chosen(self):
        return [c for c, v in self.col_vars.items() if v.get()]

    def preview_merge(self):
        chosen = self._chosen()
        if not self.sources:
            messagebox.showwarning("提示", "没有可合并的表格。请加载含表格的 Excel/Word 文档。")
            return
        if not chosen:
            messagebox.showwarning("提示", "请至少勾选一列。")
            return
        rows = export.merge_tables(self.sources, chosen, self.add_src_col.get())
        self._render_preview(rows, (["来源文件"] if self.add_src_col.get() else []) + chosen)
        self.merge_info.set("预览：%d 行" % len(rows))

    def _render_preview(self, rows, cols):
        if self.preview_tree is not None:
            self.preview_tree.master.destroy()
        f, self.preview_tree = _make_tree(self.preview_frame_holder, cols, {c: 140 for c in cols})
        f.pack(fill="both", expand=True)
        for row in rows[:500]:
            self.preview_tree.insert("", "end", values=[row.get(c, "") for c in cols])

    def export_merge(self):
        chosen = self._chosen()
        if not self.sources:
            messagebox.showwarning("提示", "没有可合并的表格。")
            return
        if not chosen:
            messagebox.showwarning("提示", "请至少勾选一列。")
            return
        path = filedialog.asksaveasfilename(title="保存合并表", defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")],
                                            initialfile="表格合并.xlsx")
        if not path:
            return
        try:
            _, n = export.merge_tables_xlsx(self.sources, chosen, path, self.add_src_col.get())
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))
            return
        messagebox.showinfo("完成", "已合并 %d 行到：\n%s" % (n, path))


def main():
    root = TkinterDnD.Tk() if _HAS_DND else tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()