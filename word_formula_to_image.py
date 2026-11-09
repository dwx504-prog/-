# -*- coding: utf-8 -*-
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import win32com.client

PASTE_EMF = 9
DOCX_FORMAT = 12

class App:
    def __init__(self, root):
        self.root=root
        root.title("Word公式图片转换助手 V2")
        root.geometry("700x470")
        self.files=[]
        self.out=tk.StringVar(value=str(Path.home()/"Documents"/"公式图片转换结果"))
        self.status=tk.StringVar(value="请选择 DOCX 文件。")
        f=ttk.Frame(root,padding=16); f.pack(fill="both",expand=True)
        ttk.Label(f,text="Word公式图片转换助手 V2",font=("Microsoft YaHei UI",16,"bold")).pack(anchor="w")
        ttk.Label(f,text="修正：复制公式为图片 → 删除原公式 → 在原位置粘贴 EMF 图片。").pack(anchor="w",pady=6)
        r=ttk.Frame(f); r.pack(fill="x")
        ttk.Button(r,text="添加 Word 文件",command=self.add).pack(side="left")
        ttk.Button(r,text="添加文件夹",command=self.folder).pack(side="left",padx=6)
        self.lb=tk.Listbox(f,height=10); self.lb.pack(fill="both",expand=True,pady=8)
        r=ttk.Frame(f); r.pack(fill="x")
        ttk.Label(r,text="输出目录").pack(side="left")
        ttk.Entry(r,textvariable=self.out).pack(side="left",fill="x",expand=True,padx=6)
        ttk.Button(r,text="浏览",command=self.browse).pack(side="left")
        self.pb=ttk.Progressbar(f); self.pb.pack(fill="x",pady=8)
        ttk.Label(f,textvariable=self.status).pack(anchor="w")
        self.btn=ttk.Button(f,text="开始转换",command=self.start); self.btn.pack(anchor="e",pady=8)
        ttk.Label(f,text="需要 Windows 和 Microsoft Word 桌面版。支持 Word 原生公式对象（OMath），不自动识别普通文本公式。",wraplength=650).pack(anchor="w")
    def add(self):
        for p in filedialog.askopenfilenames(filetypes=[("Word文档","*.docx")]):
            if p not in self.files: self.files.append(p); self.lb.insert("end",p)
    def folder(self):
        d=filedialog.askdirectory()
        if d:
            for p in Path(d).glob("*.docx"):
                if not p.name.startswith("~$") and str(p) not in self.files:
                    self.files.append(str(p)); self.lb.insert("end",str(p))
    def browse(self):
        d=filedialog.askdirectory()
        if d:self.out.set(d)
    def start(self):
        if not self.files:
            messagebox.showwarning("提示","请先添加 DOCX 文件"); return
        self.btn.config(state="disabled"); self.pb["maximum"]=len(self.files); self.pb["value"]=0
        threading.Thread(target=self.work,daemon=True).start()
    def work(self):
        output=Path(self.out.get()); output.mkdir(parents=True,exist_ok=True)
        app=None; results=[]; errors=[]
        try:
            app=win32com.client.DispatchEx("Word.Application"); app.Visible=False; app.DisplayAlerts=0
            for n,path in enumerate(self.files,1):
                try:
                    doc=app.Documents.Open(str(Path(path).resolve()),ReadOnly=False,AddToRecentFiles=False)
                    count=doc.OMaths.Count; converted=0; bad=[]
                    try:
                        for i in range(count,0,-1):
                            try:
                                rg=doc.OMaths.Item(i).Range
                                pos1,pos2=rg.Start,rg.End
                                rg.CopyAsPicture()
                                doc.Range(pos1,pos2).Delete()
                                doc.Range(pos1,pos1).PasteSpecial(DataType=PASTE_EMF)
                                converted+=1
                            except Exception as e: bad.append(f"公式{i}: {e}")
                        target=output/(Path(path).stem+"_公式转图片_V2.docx")
                        k=2
                        while target.exists():
                            target=output/(Path(path).stem+f"_公式转图片_V2_{k}.docx"); k+=1
                        doc.SaveAs2(str(target),FileFormat=DOCX_FORMAT)
                        results.append(f"{Path(path).name}: {converted}/{count}，输出 {target}")
                        if bad: results.append("异常："+ "；".join(bad[:5]))
                    finally: doc.Close(SaveChanges=False)
                except Exception as e: errors.append(f"{path}: {e}")
                self.root.after(0,lambda x=n:self.pb.configure(value=x))
        except Exception as e: errors.append("启动 Word 失败："+str(e))
        finally:
            if app:
                try: app.Quit()
                except: pass
            self.root.after(0,lambda:self.finish(results,errors))
    def finish(self,results,errors):
        self.btn.config(state="normal")
        self.status.set(f"完成：{len(results)} 个处理结果，{len(errors)} 个文件错误")
        messagebox.showinfo("转换完成","\n".join(results+errors)[:6500] or "没有输出")

if __name__=="__main__":
    w=tk.Tk()
    try:
        if "vista" in ttk.Style().theme_names(): ttk.Style().theme_use("vista")
    except: pass
    App(w); w.mainloop()
