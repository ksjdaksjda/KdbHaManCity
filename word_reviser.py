"""
Word论文修改工具 - AI重写内容，自动匹配学校格式
双击 run.bat 启动
"""
import os, sys, json, time, threading, re, io, traceback

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# ---- 依赖检查 ----
missing = []
try: from docx import Document; from docx.shared import Pt, Cm; from docx.oxml.ns import qn
except ImportError: missing.append('python-docx')
try: from openai import OpenAI
except ImportError: missing.append('openai')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "revise_config.json")

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except: pass
    return {}
def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(cfg, f, indent=2)

# ---- GUI ----
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext

class WordReviserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word论文修改工具 - AI重写 · 自动匹配格式")
        self.root.geometry("900x680")
        self.root.minsize(700, 500)
        self.config = load_config()
        self.client = None
        self.doc = None
        self.file_path = None
        self.paras = []
        self.modified_paras = {}
        self.processing = False
        self.template_analysis = None
        self.build_ui()
        if self.config.get("api_key"):
            self.api_var.set(self.config["api_key"])
            self.root.after(500, self.verify_api_silent)

    def build_ui(self):
        # API密钥
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill=tk.X)
        ttk.Label(top, text="DeepSeek API Key:").pack(side=tk.LEFT)
        self.api_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.api_var, width=45, show="*").pack(side=tk.LEFT, padx=5)
        ttk.Button(top, text="验证", command=self.verify_api).pack(side=tk.LEFT, padx=2)
        self.api_status = ttk.Label(top, text="", foreground="gray")
        self.api_status.pack(side=tk.LEFT, padx=10)
        # 文件选择
        mid = ttk.Frame(self.root, padding=10)
        mid.pack(fill=tk.X)
        ttk.Button(mid, text="选择Word文件", command=self.select_file).pack(side=tk.LEFT)
        self.file_label = ttk.Label(mid, text="  未选择文件", foreground="gray")
        self.file_label.pack(side=tk.LEFT, padx=5)
        ttk.Label(mid, text="模式:").pack(side=tk.LEFT, padx=(20, 5))
        self.mode_var = tk.StringVar(value="AI重写")
        ttk.Combobox(mid, textvariable=self.mode_var, values=["AI重写", "降重改写", "扩展补充"], width=10, state="readonly").pack(side=tk.LEFT)
        ttk.Button(mid, text="分析模板", command=self.analyze_template).pack(side=tk.LEFT, padx=5)
        # 进度条
        self.progress = ttk.Progressbar(self.root, mode='determinate')
        self.progress.pack(fill=tk.X, padx=10, pady=5)
        self.status_label = ttk.Label(self.root, text="就绪 - 1.选Word 2.分析模板 3.开始处理", foreground="gray")
        self.status_label.pack()
        # 主区域
        main = ttk.Frame(self.root)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        main.columnconfigure(0, weight=1); main.columnconfigure(1, weight=1); main.rowconfigure(0, weight=1)
        # 左侧
        left = ttk.LabelFrame(main, text="段落列表", padding=5)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        self.para_list = tk.Listbox(left, font=("Microsoft YaHei", 10))
        self.para_list.pack(fill=tk.BOTH, expand=True)
        self.para_list.bind('<<ListboxSelect>>', self.on_para_select)
        # 右侧
        right = ttk.LabelFrame(main, text="预览 - 原文 vs AI修改后", padding=5)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(0, weight=1); right.rowconfigure(1, weight=1); right.columnconfigure(0, weight=1)
        ttk.Label(right, text="原文:").grid(row=0, column=0, sticky="w")
        self.orig_text = scrolledtext.ScrolledText(right, height=8, font=("Microsoft YaHei", 10), wrap=tk.WORD)
        self.orig_text.grid(row=0, column=0, sticky="nsew", pady=(0, 5))
        ttk.Label(right, text="AI修改后:").grid(row=1, column=0, sticky="w")
        self.mod_text = scrolledtext.ScrolledText(right, height=8, font=("Microsoft YaHei", 10), wrap=tk.WORD)
        self.mod_text.grid(row=1, column=0, sticky="nsew")
        # 底部
        bottom = ttk.Frame(self.root, padding=10)
        bottom.pack(fill=tk.X)
        ttk.Button(bottom, text="开始处理", command=self.start_process).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom, text="暂停", command=lambda: setattr(self, 'processing', False)).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom, text="保存修改", command=self.save_doc).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom, text="撤销当前段", command=self.undo_para).pack(side=tk.LEFT, padx=5)

    def verify_api_silent(self):
        try:
            c = OpenAI(api_key=self.api_var.get(), base_url="https://api.deepseek.com/v1")
            c.chat.completions.create(model="deepseek-chat", messages=[{"role":"user","content":"OK"}], max_tokens=2)
            self.client = c; self.api_status.config(text="已连接", foreground="green")
            self.config["api_key"] = self.api_var.get(); save_config(self.config)
        except: pass

    def verify_api(self):
        key = self.api_var.get().strip()
        if not key: self.api_status.config(text="请输入密钥", foreground="red"); return
        self.api_status.config(text="验证中...", foreground="orange")
        try:
            c = OpenAI(api_key=key, base_url="https://api.deepseek.com/v1")
            c.chat.completions.create(model="deepseek-chat", messages=[{"role":"user","content":"OK"}], max_tokens=2)
            self.client = c; self.api_status.config(text="连接成功", foreground="green")
            self.config["api_key"] = key; save_config(self.config)
        except Exception as e:
            self.api_status.config(text=f"失败: {str(e)[:50]}", foreground="red")

    def analyze_template(self):
        if not self.doc: messagebox.showwarning("提示", "请先选择Word文件"); return
        try:
            doc = self.doc; sec = doc.sections[0]
            m = f"上{sec.top_margin/360000:.1f}cm 下{sec.bottom_margin/360000:.1f}cm 左{sec.left_margin/360000:.1f}cm 右{sec.right_margin/360000:.1f}cm"
            psize = f"{sec.page_width/360000:.0f}x{sec.page_height/360000:.0f}cm"
            fonts={}; sizes={}; h_fonts={}
            for p in doc.paragraphs:
                try:
                    if not p.runs or not p.text.strip(): continue
                    r=p.runs[0]; fn=r.font.name or "默认"; fs=str(r.font.size) if r.font.size else "默认"
                    if p.style.name.startswith("Heading"): h_fonts[fn]=h_fonts.get(fn,0)+1
                    else: fonts[fn]=fonts.get(fn,0)+1; sizes[fs]=sizes.get(fs,0)+1
                except: pass
            bf=max(fonts,key=fonts.get) if fonts else "未检测"
            hf=max(h_fonts,key=h_fonts.get) if h_fonts else "未检测"
            bs=max(sizes,key=sizes.get) if sizes else "未检测"
            all_p=[p for p in doc.paragraphs if p.text.strip()]
            h1=[p for p in all_p if p.style.name=='Heading 1']
            h2=[p for p in all_p if p.style.name=='Heading 2']
            body=[p for p in all_p if not p.style.name.startswith("Heading")]
            a=[f"页边距: {m}", f"纸张: {psize}", f"标题字体: {hf}", f"正文字体: {bf} {bs}", f"一级标题: {len(h1)}个", f"二级标题: {len(h2)}个", f"正文段: {len(body)}个", f"表格: {len(doc.tables)}个"]
            if h1: a.append(f"标题示例: {h1[0].text[:50]}")
            self.template_analysis = "\n".join(a)
            messagebox.showinfo("模板分析报告", "【文档完整分析】\n\n" + "\n".join(a))
            self.status_label.config(text=f"模板已分析: {hf}标题 + {bf}正文 | 点「开始处理」开始AI重写")
        except Exception as e:
            messagebox.showerror("分析失败", f"无法分析文档结构\n\n请确认:\n1. 文件是 .docx 格式(非 .doc)\n2. 文件未被其他程序占用\n\n错误: {str(e)[:150]}")

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Word文件", "*.docx"), ("所有文件", "*.*")])
        if not path: return
        if not path.lower().endswith('.docx'):
            messagebox.showerror("格式错误", "仅支持 .docx 格式\n旧版 .doc 请用Word另存为 .docx")
            return
        self.file_path = path
        self.file_label.config(text=f"  {os.path.basename(path)}")
        self.status_label.config(text="读取中...")
        self.root.update()
        try:
            self.doc = Document(path)
            self.paras = [(i, p) for i, p in enumerate(self.doc.paragraphs) if p.text.strip()]
            self.modified_paras = {}
            self.template_analysis = None
            self.para_list.delete(0, tk.END)
            for idx, para in self.paras:
                h = "[H] " if para.style.name.startswith("Heading") else "     "
                self.para_list.insert(tk.END, f"{h}{para.text[:60]}...")
            self.status_label.config(text=f"已加载 {len(self.paras)} 段 | 点「分析模板」查看结构 | 点「开始处理」AI重写")
        except Exception as e:
            messagebox.showerror("读取失败", f"无法读取该Word文件\n\n可能原因:\n1. 旧版 .doc 格式(需另存为 .docx)\n2. 文件被其他程序占用\n3. 文件已损坏\n\n错误: {str(e)[:200]}")

    def on_para_select(self, event):
        sel = self.para_list.curselection()
        if not sel or sel[0] >= len(self.paras): return
        pidx, para = self.paras[sel[0]]
        self.orig_text.delete("1.0", tk.END); self.orig_text.insert("1.0", para.text)
        self.mod_text.delete("1.0", tk.END)
        mod = self.modified_paras.get(pidx, "")
        self.mod_text.insert("1.0", mod if mod else "(未修改)")

    def start_process(self):
        if not self.client: self.verify_api()
        if not self.client: messagebox.showwarning("提示", "请先验证API密钥"); return
        if not self.paras: messagebox.showwarning("提示", "请先选择Word文件"); return
        self.processing = True; self.progress["maximum"] = len(self.paras); self.progress["value"] = 0
        threading.Thread(target=self.process_loop, daemon=True).start()

    def process_loop(self):
        mode = self.mode_var.get()
        fmt = self.template_analysis or """高校论文标准格式 GB/T 7713:
标题: 黑体, 章标题三号(16pt)加粗, 节标题四号(14pt)加粗
正文: 宋体小四(12pt), 首行缩进2字符, 1.5倍行距
页边距: 上2.54cm 下2.54cm 左3.17cm 右3.17cm"""
        modes = {
            "AI重写": f"你是中文学术论文写作专家。根据原文主题完全重写论文内容。{fmt}\n学术严谨、逻辑清晰、善用长短句、避免AI套话。",
            "降重改写": f"深度改写降重。变换句式、同义词替换。{fmt}",
            "扩展补充": f"扩展内容，补充细节和论证。{fmt}"
        }
        desc = modes.get(mode, modes["AI重写"])
        for idx, (pidx, para) in enumerate(self.paras):
            if not self.processing: break
            text = para.text.strip()
            if len(text) < 10: self.progress["value"] = idx+1; continue
            self.status_label.config(text=f"[{idx+1}/{len(self.paras)}] 修改: {text[:30]}...")
            try:
                r = self.client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role":"system","content":desc}, {"role":"user","content":f"原文:\n{text}"}],
                    temperature=0.7, max_tokens=2048
                )
                nt = r.choices[0].message.content
                if nt and len(nt.strip()) > 5: self.modified_paras[pidx] = nt.strip()
            except Exception as e: self.modified_paras[pidx] = f"[错误: {e}]"
            self.progress["value"] = idx+1
            time.sleep(0.3)
        self.status_label.config(text=f"完成！修改 {len(self.modified_paras)} 段，点「保存修改」导出")

    def undo_para(self):
        sel = self.para_list.curselection()
        if not sel or sel[0] >= len(self.paras): return
        pidx = self.paras[sel[0]][0]
        if pidx in self.modified_paras:
            del self.modified_paras[pidx]
            self.mod_text.delete("1.0", tk.END); self.mod_text.insert("1.0", "(已撤销)")

    def save_doc(self):
        if not self.doc or not self.modified_paras: messagebox.showwarning("提示", "没有修改内容"); return
        self.status_label.config(text="保存中...应用GB/T 7713标准格式...")
        self.root.update()
        try:
            for sec in self.doc.sections:
                sec.top_margin = Cm(2.54); sec.bottom_margin = Cm(2.54)
                sec.left_margin = Cm(3.17); sec.right_margin = Cm(3.17)
            for pidx, new_text in self.modified_paras.items():
                para = self.doc.paragraphs[pidx]
                is_heading = para.style.name.startswith("Heading")
                for run in para.runs: run.text = ""
                if para.runs:
                    r = para.runs[0]; r.text = new_text
                    if is_heading:
                        r.font.name = "黑体"; r._element.rPr.rFonts.set(qn('w:eastAsia'), "黑体")
                        lvl = int(para.style.name.replace("Heading ","") or "1")
                        r.font.size = Pt({1:16,2:14}.get(lvl,12)); r.bold = True
                    else:
                        r.font.name = "宋体"; r._element.rPr.rFonts.set(qn('w:eastAsia'), "宋体")
                        r.font.size = Pt(12); r.bold = False
                    para.paragraph_format.line_spacing = 1.5
                else: para.text = new_text
            out = self.file_path.replace(".docx", "_修改版.docx")
            if os.path.exists(out): out = self.file_path.replace(".docx", f"_修改版{int(time.time())%10000}.docx")
            self.doc.save(out)
            self.status_label.config(text=f"已保存: {os.path.basename(out)}")
            os.startfile(os.path.dirname(out))
            messagebox.showinfo("保存成功", f"文件已保存:\n{out}")
        except Exception as e:
            messagebox.showerror("保存失败", f"错误: {str(e)[:200]}")

def main():
    if missing:
        root = tk.Tk(); root.withdraw()
        if messagebox.askyesno("安装依赖", f"缺少: {', '.join(missing)}\n\n自动安装？"):
            import subprocess
            subprocess.run([sys.executable, "-m", "pip", "install"] + missing, capture_output=True)
    root = tk.Tk()
    WordReviserApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
