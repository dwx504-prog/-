# -*- coding: utf-8 -*-
"""
Word公式图片转换助手
将 Word 原生公式（OMath）转换为图片并尽量原位替换。
运行环境：Windows + 已安装 Microsoft Word。
"""
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

try:
    import win32com.client
except Exception:
    win32com = None

WD_PASTE_ENHANCED_METAFILE = 9
WD_FORMAT_XML_DOCUMENT = 12

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Word公式图片转换助手")
        self.root.geometry("700x480")
        self.root.minsize(650, 440)
        self.files = []
        self.output_dir = tk.StringVar(value=str(Path.home() / "Documents" / "公式图片转换结果"))
        self.status = tk.StringVar(value="请选择需要处理的 Word 文档。")
        self._build()

    def _build(self):
        frm = ttk.Frame(self.root, padding=18)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Word公式图片转换助手", font=("Microsoft YaHei UI", 17, "bold")).pack(anchor="w")
        ttk.Label(frm, text="将 Word 原生公式转换为图片，生成新的 DOCX 文件，不覆盖原文件。",
                  font=("Microsoft YaHei UI", 9)).pack(anchor="w", pady=(4, 14))

        row = ttk.Frame(frm)
        row.pack(fill="x")
        ttk.Button(row, text="添加 Word 文件", command=self.add_files).pack(side="left")
        ttk.Button(row, text="添加整个文件夹", command=self.add_folder).pack(side="left", padx=8)
        ttk.Button(row, text="清空列表", command=self.clear).pack(side="left")

        self.listbox = tk.Listbox(frm, height=10, font=("Microsoft YaHei UI", 9))
        self.listbox.pack(fill="both", expand=True, pady=10)

        outrow = ttk.Frame(frm)
        outrow.pack(fill="x", pady=(2, 8))
        ttk.Label(outrow, text="输出目录：").pack(side="left")
        ttk.Entry(outrow, textvariable=self.output_dir).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(outrow, text="浏览", command=self.choose_output).pack(side="left")

        self.progress = ttk.Progressbar(frm, mode="determinate")
        self.progress.pack(fill="x", pady=(4, 8))
        ttk.Label(frm, textvariable=self.status, wraplength=650).pack(anchor="w")

        self.start_btn = ttk.Button(frm, text="开始转换", command=self.start)
        self.start_btn.pack(anchor="e", pady=(12, 0))

        note = ("注意：需要 Windows 电脑并安装桌面版 Microsoft Word。"
                "本工具主要转换 Word 原生公式对象；普通文本形式的公式不会自动识别。")
        ttk.Label(frm, text=note, foreground="#555", wraplength=650).pack(anchor="w", pady=(12, 0))

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择 Word 文档",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")]
        )
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert("end", p)

    def add_folder(self):
        folder = filedialog.askdirectory(title="选择包含 Word 文档的文件夹")
        if not folder:
            return
        for p in sorted(Path(folder).glob("*.docx")):
            s = str(p)
            if not p.name.startswith("~$") and s not in self.files:
                self.files.append(s)
                self.listbox.insert("end", s)

    def clear(self):
        self.files.clear()
        self.listbox.delete(0, "end")
        self.status.set("文件列表已清空。")

    def choose_output(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.output_dir.set(d)

    def start(self):
        if not self.files:
            messagebox.showwarning("提示", "请先添加至少一个 DOCX 文件。")
            return
        if not self.output_dir.get().strip():
            messagebox.showwarning("提示", "请先选择输出目录。")
            return
        if win32com is None:
            messagebox.showerror("缺少组件", "未检测到 pywin32。请重新安装本工具，或运行安装脚本。")
            return
        self.start_btn.config(state="disabled")
        self.progress["maximum"] = len(self.files)
        self.progress["value"] = 0
        threading.Thread(target=self._worker, daemon=True).start()

    def _set_status(self, text):
        self.root.after(0, lambda: self.status.set(text))

    def _worker(self):
        outdir = Path(self.output_dir.get()).expanduser()
        outdir.mkdir(parents=True, exist_ok=True)
        word = None
        successes, failures = [], []
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            for idx, src in enumerate(self.files, 1):
                try:
                    self._set_status(f"正在处理（{idx}/{len(self.files)}）：{Path(src).name}")
                    result = self.convert_one(word, Path(src), outdir)
                    successes.append(result)
                except Exception as e:
                    failures.append((src, str(e)))
                self.root.after(0, lambda v=idx: self.progress.configure(value=v))
        except Exception as e:
            failures.append(("Word 启动失败", str(e)))
        finally:
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            self.root.after(0, lambda: self._finish(successes, failures))

    def convert_one(self, word, src, outdir):
        if src.name.startswith("~$"):
            raise RuntimeError("这是 Word 临时锁定文件，不能转换。")
        doc = word.Documents.Open(str(src.resolve()), ReadOnly=False, AddToRecentFiles=False)
        converted = 0
        try:
            # 从后往前处理，减少前面公式替换对后续位置的影响。
            count = doc.OMaths.Count
            for i in range(count, 0, -1):
                try:
                    equation_range = doc.OMaths.Item(i).Range
                    equation_range.Select()
                    word.Selection.CopyAsPicture()
                    word.Selection.PasteSpecial(DataType=WD_PASTE_ENHANCED_METAFILE)
                    converted += 1
                except Exception:
                    # 单个公式失败时继续处理其余公式。
                    continue

            target = outdir / f"{src.stem}_公式转图片.docx"
            # 避免覆盖已有结果。
            n = 2
            while target.exists():
                target = outdir / f"{src.stem}_公式转图片_{n}.docx"
                n += 1
            doc.SaveAs2(str(target), FileFormat=WD_FORMAT_XML_DOCUMENT)
            return (str(src), str(target), converted, count)
        finally:
            doc.Close(SaveChanges=False)

    def _finish(self, successes, failures):
        self.start_btn.config(state="normal")
        lines = []
        for src, target, done, total in successes:
            lines.append(f"成功：{Path(src).name}，转换 {done}/{total} 个公式\n输出：{target}")
        for src, err in failures:
            lines.append(f"失败：{src}\n原因：{err}")
        report = "\n\n".join(lines) if lines else "没有生成文件。"
        self.status.set(f"处理结束：成功 {len(successes)} 个，失败 {len(failures)} 个。")
        messagebox.showinfo("转换完成", report[:6000])

def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
