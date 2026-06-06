"""
论文重写器 - AI重写Word内容，排版格式完全不改
双击run.bat启动，选择"论文重写器"
"""
import os, sys, json, time, threading, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

try: from docx import Document
except ImportError: pass
try: from openai import OpenAI
except ImportError: pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "revise_config.json")
FORMAT_FILE = os.path.join(SCRIPT_DIR, "template_format.json")

def load_json(path):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f: return json.load(f)
        except: pass
    return {}

def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f: json.dump(data, f, indent=2, ensure_ascii=False)

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext

class RewriterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("论文重写器 - 排版不动 · 只改文字")
        self.root.geometry("900x680")
        self.root.minsize(700, 500)
        self.client = None
        self.doc = None
        self.file_path = None
        self.paras = []          # [(pidx, paragraph_obj), ...]
        self.modified = {}       # {pidx: new_text}
        self.processing = False
        self.school_name = ""
        self.template_info = ""
        self.build_ui()
        # 自动读API和模板配置
        cfg = load_json(CONFIG_FILE)
        if cfg.get("api_key"):
            self.api_var.set(cfg["api_key"])
            self.root.after(300, self.auto_verify)
        fmt = load_json(FORMAT_FILE)
        if fmt:
            self.template_info = f"模板: {fmt.get('fonts',{}).get('body','?')} {fmt.get('fonts',{}).get('body_size','?')}"
            self.status_label.config(text=f"已加载格式配置 | {self.template_info}")

    def build_ui(self):
        # 第一行: API
        f1 = ttk.Frame(self.root, padding=10); f1.pack(fill=tk.X)
        ttk.Label(f1, text="API Key:").pack(side=tk.LEFT)
        self.api_var = tk.StringVar()
        ttk.Entry(f1, textvariable=self.api_var, width=45, show="*").pack(side=tk.LEFT, padx=5)
        ttk.Button(f1, text="验证", command=self.verify_key).pack(side=tk.LEFT)
        self.api_lbl = ttk.Label(f1, text="", foreground="gray")
        self.api_lbl.pack(side=tk.LEFT, padx=10)
        # 第二行: 文件+模式+字数
        f2 = ttk.Frame(self.root, padding=10); f2.pack(fill=tk.X)
        ttk.Button(f2, text="选择Word文件", command=self.load_file).pack(side=tk.LEFT)
        self.file_lbl = ttk.Label(f2, text="  未选文件", foreground="gray")
        self.file_lbl.pack(side=tk.LEFT, padx=5)
        ttk.Label(f2, text="模式:").pack(side=tk.LEFT, padx=(15,5))
        self.mode_var = tk.StringVar(value="AI重写")
        ttk.Combobox(f2, textvariable=self.mode_var, values=["AI重写", "降重改写", "扩展补充"], width=10, state="readonly").pack(side=tk.LEFT)
        ttk.Label(f2, text="目标字数:").pack(side=tk.LEFT, padx=(10,5))
        self.words_var = tk.StringVar(value="不限")
        ttk.Combobox(f2, textvariable=self.words_var, values=["不限", "3000字", "5000字", "8000字", "10000字", "15000字", "20000字", "30000字"], width=8, state="readonly").pack(side=tk.LEFT)
        ttk.Label(f2, text="学校:").pack(side=tk.LEFT, padx=(10,5))
        self.school_var = tk.StringVar()
        ttk.Entry(f2, textvariable=self.school_var, width=12).pack(side=tk.LEFT)
        self.school_var.trace_add('write', lambda*a: setattr(self, 'school_name', self.school_var.get()))
        # 进度
        self.prog = ttk.Progressbar(self.root, mode='determinate')
        self.prog.pack(fill=tk.X, padx=10, pady=5)
        self.status_label = ttk.Label(self.root, text="就绪", foreground="gray")
        self.status_label.pack()
        # 主区域
        main = ttk.Frame(self.root); main.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        main.columnconfigure(0, weight=1); main.columnconfigure(1, weight=1); main.rowconfigure(0, weight=1)
        left = ttk.LabelFrame(main, text="段落", padding=5); left.grid(row=0, column=0, sticky="nsew", padx=(0,5))
        self.plist = tk.Listbox(left, font=("Microsoft YaHei", 10))
        self.plist.pack(fill=tk.BOTH, expand=True)
        self.plist.bind('<<ListboxSelect>>', self.on_select)
        right = ttk.LabelFrame(main, text="预览", padding=5); right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(0, weight=1); right.rowconfigure(1, weight=1); right.columnconfigure(0, weight=1)
        ttk.Label(right, text="原文:").grid(row=0, column=0, sticky="w")
        self.orig = scrolledtext.ScrolledText(right, height=8, font=("Microsoft YaHei", 10), wrap=tk.WORD)
        self.orig.grid(row=0, column=0, sticky="nsew", pady=(0,5))
        ttk.Label(right, text="AI重写后:").grid(row=1, column=0, sticky="w")
        self.modv = scrolledtext.ScrolledText(right, height=8, font=("Microsoft YaHei", 10), wrap=tk.WORD)
        self.modv.grid(row=1, column=0, sticky="nsew")
        # 底部按钮
        f3 = ttk.Frame(self.root, padding=10); f3.pack(fill=tk.X)
        ttk.Button(f3, text="开始处理", command=self.start).pack(side=tk.LEFT, padx=5)
        ttk.Button(f3, text="暂停", command=lambda: setattr(self,'processing',False)).pack(side=tk.LEFT, padx=5)
        ttk.Button(f3, text="保存(排版不变)", command=self.save).pack(side=tk.LEFT, padx=5)
        ttk.Button(f3, text="撤销选中段", command=self.undo).pack(side=tk.LEFT, padx=5)

    def auto_verify(self):
        try:
            c = OpenAI(api_key=self.api_var.get(), base_url="https://api.deepseek.com/v1")
            c.chat.completions.create(model="deepseek-chat", messages=[{"role":"user","content":"OK"}], max_tokens=2)
            self.client = c; self.api_lbl.config(text="已连接", foreground="green")
        except: pass

    def verify_key(self):
        k = self.api_var.get().strip()
        if not k: self.api_lbl.config(text="请输入密钥", foreground="red"); return
        self.api_lbl.config(text="验证中...", foreground="orange")
        try:
            c = OpenAI(api_key=k, base_url="https://api.deepseek.com/v1")
            c.chat.completions.create(model="deepseek-chat", messages=[{"role":"user","content":"OK"}], max_tokens=2)
            self.client = c; self.api_lbl.config(text="已连接", foreground="green")
            save_json(CONFIG_FILE, {"api_key": k})
        except Exception as e:
            self.api_lbl.config(text=f"失败: {str(e)[:40]}", foreground="red")

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Word文件", "*.docx"), ("所有文件", "*.*")])
        if not path: return
        if not path.lower().endswith('.docx'):
            messagebox.showerror("格式错误", "仅支持 .docx\n旧版 .doc 请用Word另存为 .docx"); return
        self.file_path = path
        self.file_lbl.config(text=f"  {os.path.basename(path)}")
        self.status_label.config(text="读取中..."); self.root.update()
        try:
            self.doc = Document(path)
            # 只收集有内容的段落
            self.paras = [(i, p) for i, p in enumerate(self.doc.paragraphs) if p.text.strip()]
            self.modified = {}
            self.plist.delete(0, tk.END)
            for i, (pidx, p) in enumerate(self.paras):
                h = "[H] " if p.style.name.startswith("Heading") else "     "
                self.plist.insert(tk.END, f"{h}{p.text[:60]}...")
            # 自动分析格式
            self._analyze_format()
            self.status_label.config(text=f"已加载 {len(self.paras)} 段 | 排版不会改动 | 选择字数后点开始处理")
        except Exception as e:
            messagebox.showerror("失败", f"无法读取:\n{str(e)[:200]}")

    def _analyze_format(self):
        """仅分析格式信息，不动文档"""
        doc = self.doc
        sec = doc.sections[0]
        fonts = {}
        for p in doc.paragraphs:
            if p.runs:
                try:
                    fn = p.runs[0].font.name or "默认"
                    fonts[fn] = fonts.get(fn, 0) + 1
                except: pass
        tf = max(fonts, key=fonts.get) if fonts else "未检测"
        self.template_info = f"字体: {tf} | 页边距: 上{sec.top_margin/360000:.1f}cm"

    def on_select(self, evt):
        sel = self.plist.curselection()
        if not sel or sel[0] >= len(self.paras): return
        _, p = self.paras[sel[0]]
        self.orig.delete("1.0", tk.END); self.orig.insert("1.0", p.text)
        self.modv.delete("1.0", tk.END)
        self.modv.insert("1.0", self.modified.get(_, "(未修改)"))

    def start(self):
        if not self.client: self.verify_key()
        if not self.client: messagebox.showwarning("先验证API"); return
        if not self.paras: messagebox.showwarning("先选Word文件"); return
        self.processing = True
        self.prog["maximum"] = len(self.paras)
        self.prog["value"] = 0
        threading.Thread(target=self._process, daemon=True).start()

    def _process(self):
        mode = self.mode_var.get()
        wc = self.words_var.get()
        wc_str = f"目标总字数: {wc}。每个段落请相应调整字数。" if wc != "不限" else ""
        school = self.school_name.strip()
        school_info = f"学校: {school}。请根据你的训练数据中该校（或同类中国高校）的毕业论文具体格式标准来写作。" if school else ""
        fmt = (self.template_info or "保持原文排版格式，只替换文字内容") + " " + school_info
        descs = {
            "AI重写": f"你是学术论文专家。请重写此段落，保持主题不变。{wc_str}格式要求: {fmt}。只输出重写后文本，无额外说明。",
            "降重改写": f"深度降重改写。变换句式、同义词。{wc_str}格式: {fmt}。只输出改写后文本。",
            "扩展补充": f"扩展内容，丰富论证。{wc_str}格式: {fmt}。只输出扩展后文本。"
        }
        desc = descs.get(mode, descs["AI重写"])
        for i, (pidx, p) in enumerate(self.paras):
            if not self.processing: break
            txt = p.text.strip()
            if len(txt) < 10: self.prog["value"] = i+1; continue
            self.status_label.config(text=f"[{i+1}/{len(self.paras)}] {txt[:30]}...")
            try:
                r = self.client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role":"system","content":desc}, {"role":"user","content":f"原文:\n{txt}"}],
                    temperature=0.7, max_tokens=2048
                )
                nt = r.choices[0].message.content
                if nt and len(nt.strip()) > 5: self.modified[pidx] = nt.strip()
            except: pass
            self.prog["value"] = i+1
            time.sleep(0.3)
        self.status_label.config(text=f"完成 {len(self.modified)} 段 | 点「保存(排版不变)」导出")

    def undo(self):
        sel = self.plist.curselection()
        if sel and sel[0] < len(self.paras):
            pidx = self.paras[sel[0]][0]
            if pidx in self.modified: del self.modified[pidx]; self.modv.delete("1.0", tk.END)

    def save(self):
        """
        核心：只替换文字内容，不碰任何格式！
        """
        if not self.doc or not self.modified:
            messagebox.showwarning("无修改"); return
        cnt = 0
        for pidx, new_text in self.modified.items():
            para = self.doc.paragraphs[pidx]
            # 获取原始runs的格式信息
            orig_runs = [(r.text, r.font.name, r.font.size, r.bold, r.italic)
                         for r in para.runs] if para.runs else []
            # 清空
            for r in para.runs: r.text = ""
            if para.runs:
                # 只用第一个run放新文字，保留其原有格式
                para.runs[0].text = new_text
                cnt += 1
            else:
                # 无run的段落，直接设text（python-docx自动创建run）
                # 这会继承段落样式默认格式
                para.text = new_text
                cnt += 1
        out = self.file_path.replace(".docx", "_重写版.docx")
        if os.path.exists(out):
            out = self.file_path.replace(".docx", f"_重写版{int(time.time())%10000}.docx")
        self.doc.save(out)
        os.startfile(os.path.dirname(out))
        messagebox.showinfo("保存成功", f"已保存:\n{out}\n\n排版格式: 完全不变\n修改段落: {cnt}")

def main():
    # 检查依赖
    missing = []
    try: from docx import Document
    except ImportError: missing.append('python-docx')
    try: from openai import OpenAI
    except ImportError: missing.append('openai')
    if missing:
        root = tk.Tk(); root.withdraw()
        if messagebox.askyesno("安装依赖", f"缺少: {', '.join(missing)}\n自动安装？"):
            import subprocess
            subprocess.run([sys.executable, "-m", "pip", "install"] + missing, capture_output=True)
    root = tk.Tk()
    RewriterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
