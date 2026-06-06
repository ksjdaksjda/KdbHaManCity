"""
模板填充器 - 将AI写作内容精准灌入学校Word模板
排版100%不变，只替换文字。操作XML级别的文本节点。
"""
import os, sys, json, io, re, copy
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

try: from docx import Document
except ImportError:
    print("需要 python-docx: pip install python-docx"); input(); sys.exit(1)

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

class FillTemplateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("模板填充器 - 内容灌入模板 · 排版不动")
        self.root.geometry("900x700")
        self.root.minsize(600, 400)
        self.template_doc = None
        self.template_path = None
        self.content_text = ""
        self.build_ui()

    def build_ui(self):
        f1 = ttk.Frame(self.root, padding=10); f1.pack(fill=tk.X)
        ttk.Button(f1, text="1. 选择学校Word模板", command=self.load_template).pack(side=tk.LEFT)
        self.tpl_lbl = ttk.Label(f1, text="  未选择", foreground="gray")
        self.tpl_lbl.pack(side=tk.LEFT, padx=10)

        f2 = ttk.Frame(self.root, padding=10); f2.pack(fill=tk.X)
        ttk.Label(f2, text="2. 粘贴AI写作内容 (从网页复制):").pack(anchor="w")
        self.content_input = scrolledtext.ScrolledText(self.root, height=14, font=("Microsoft YaHei", 11), wrap=tk.WORD)
        self.content_input.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 5))

        f3 = ttk.Frame(self.root, padding=10); f3.pack(fill=tk.X)
        ttk.Label(f3, text="3. 点击填充 →").pack(side=tk.LEFT)
        ttk.Button(f3, text="填充到模板", command=self.fill).pack(side=tk.LEFT, padx=5)
        ttk.Label(f3, text="  (排版100%不变，只替换文字)", foreground="gray").pack(side=tk.LEFT)

        self.log = scrolledtext.ScrolledText(self.root, height=6, font=("Microsoft YaHei", 10))
        self.log.pack(fill=tk.X, padx=10, pady=5)

    def load_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word文件", "*.docx")])
        if not path: return
        if not path.lower().endswith('.docx'):
            messagebox.showerror("格式错误", "仅支持 .docx"); return
        self.template_path = path
        self.template_doc = Document(path)
        self.tpl_lbl.config(text=f"  {os.path.basename(path)} ({len(self.template_doc.paragraphs)}段)")
        # 显示结构预览
        struct = []
        for p in self.template_doc.paragraphs:
            if p.text.strip():
                h = "[H] " if p.style.name.startswith("Heading") else "  - "
                struct.append(f"{h}{p.text[:80]}")
        self.log.delete("1.0", tk.END)
        self.log.insert("1.0", f"模板结构 ({len(self.template_doc.paragraphs)}段):\n" + "\n".join(struct[:30]))

    def fill(self):
        if not self.template_doc:
            messagebox.showwarning("提示", "请先选择学校Word模板"); return
        content = self.content_input.get("1.0", "end-1c").strip()
        if not content:
            messagebox.showwarning("提示", "请粘贴AI写作内容"); return

        # ---- 智能匹配：AI内容段落 → 模板段落 ----
        doc = self.template_doc
        # 收集模板中有内容的段落
        tpl_paras = [(i, p) for i, p in enumerate(doc.paragraphs) if p.text.strip()]
        # 将AI内容按换行拆成段落
        content_paras = [p.strip() for p in content.split('\n') if p.strip()]
        # 去掉明显不是正文的行（如标题标记、分隔线）
        content_paras = [p for p in content_paras if not p.startswith('===') and len(p) > 4]

        self.log.delete("1.0", tk.END)
        log_lines = [f"模板段落: {len(tpl_paras)} 个", f"AI内容段落: {len(content_paras)} 个", ""]

        filled = 0
        # 策略1: 按顺序匹配（1对1映射）
        for i, (pidx, tpl_para) in enumerate(tpl_paras):
            if i >= len(content_paras): break
            old_text = tpl_para.text.strip()
            new_text = content_paras[i]

            # 跳过太短的内容（可能是标题编号等）
            if len(new_text) < 8: continue
            # 不要太长（超过3000字的可能是一整章被当成一段）
            if len(new_text) > 3000:
                # 尝试按句号拆分
                sub_paras = [s.strip() for s in new_text.split('。') if len(s.strip()) > 8]
                if sub_paras:
                    new_text = sub_paras[0] + '。'
                    # 把剩余内容插入到后续位置
                    for j, sp in enumerate(sub_paras[1:]):
                        if i + j + 1 < len(content_paras):
                            continue
                        content_paras.insert(i + j + 1, sp + '。')

            # === 核心：用XML操作替换文字，不动格式 ===
            self._replace_text_only(doc.paragraphs[pidx], new_text)
            log_lines.append(f"  [{i+1}] {old_text[:30]}... → {new_text[:30]}...")
            filled += 1

        # 保存
        out_path = self.template_path.replace(".docx", "_填充版.docx")
        if os.path.exists(out_path):
            out_path = self.template_path.replace(".docx", f"_填充版{int(__import__('time').time())%10000}.docx")
        doc.save(out_path)
        log_lines.append("")
        log_lines.append(f"✅ 完成！填充 {filled} 段 | 排版完全不变")
        log_lines.append(f"输出: {out_path}")
        self.log.delete("1.0", tk.END)
        self.log.insert("1.0", "\n".join(log_lines))
        os.startfile(os.path.dirname(out_path))
        messagebox.showinfo("填充成功", f"已填充 {filled} 个段落\n排版100%不变\n\n文件:\n{out_path}")

    def _replace_text_only(self, para, new_text):
        """
        只替换文字，不动任何格式。
        直接操作XML: 修改所有 <w:t> 节点的文本。
        """
        nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        # 清空所有runs的文本
        runs = para.runs
        if runs:
            for r in runs: r.text = ""
            runs[0].text = new_text
        else:
            # 没有run的段落，创建新run
            para.text = new_text
        # 保证段落里面的所有w:t都被替换
        # 使用lxml直接操作XML（如果可用）
        try:
            from lxml import etree
            t_elements = para._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
            if t_elements:
                # 清空所有
                for t_el in t_elements[1:]:
                    t_el.text = ""
                    t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                t_elements[0].text = new_text
                t_elements[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        except ImportError:
            pass  # lxml not available, runs method already used above

def main():
    try: from docx import Document
    except:
        root = tk.Tk(); root.withdraw()
        if messagebox.askyesno("缺少依赖", "需要 python-docx\n自动安装？"):
            import subprocess, sys
            subprocess.run([sys.executable, "-m", "pip", "install", "python-docx"], capture_output=True)
    root = tk.Tk()
    FillTemplateApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
