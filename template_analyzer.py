"""
学校论文模板分析器 - 完整提取Word排版信息
双击运行，输出格式配置文件供论文重写工具使用
"""
import os, sys, json, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

try: from docx import Document; from docx.shared import Pt, Cm, Inches
except ImportError: print("请先安装: pip install python-docx"); input(); sys.exit(1)

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_CONFIG = os.path.join(SCRIPT_DIR, "template_format.json")

class TemplateAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("学校论文模板分析器 - 提取Word完整排版")
        self.root.geometry("850x650")
        self.root.minsize(600, 400)
        self.doc = None
        self.build_ui()

    def build_ui(self):
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill=tk.X)
        ttk.Button(top, text="选择学校模板Word文件", command=self.load_file).pack(side=tk.LEFT)
        self.file_label = ttk.Label(top, text="  未选择文件", foreground="gray")
        self.file_label.pack(side=tk.LEFT, padx=10)
        ttk.Button(top, text="导出格式配置", command=self.export_config).pack(side=tk.RIGHT)

        self.report = scrolledtext.ScrolledText(self.root, font=("Microsoft YaHei", 11), wrap=tk.WORD)
        self.report.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        bottom = ttk.Frame(self.root, padding=10)
        bottom.pack(fill=tk.X)
        ttk.Label(bottom, text="此工具只分析模板格式，不修改任何内容", foreground="gray").pack()

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Word文件", "*.docx"), ("所有文件", "*.*")])
        if not path: return
        if not path.lower().endswith('.docx'):
            messagebox.showerror("格式错误", "仅支持 .docx\n旧版 .doc 请用Word另存为 .docx"); return
        try:
            self.doc = Document(path)
            self.file_label.config(text=f"  {os.path.basename(path)}")
            self.analyze()
        except Exception as e:
            self.report.insert("1.0", f"[错误] 无法读取文件\n{str(e)[:300]}")

    def analyze(self):
        doc = self.doc
        report = []
        report.append("=" * 60)
        report.append("  学校论文模板 - 完整格式分析报告")
        report.append("=" * 60)
        report.append("")

        # ---- 页面设置 ----
        sec = doc.sections[0]
        tw = sec.page_width / 360000
        th = sec.page_height / 360000
        report.append("【页面设置】")
        report.append(f"  纸张大小: {tw:.0f}cm × {th:.0f}cm ({'A4' if abs(tw-21)<1 else '自定义'})")
        report.append(f"  上边距: {sec.top_margin/360000:.1f}cm")
        report.append(f"  下边距: {sec.bottom_margin/360000:.1f}cm")
        report.append(f"  左边距: {sec.left_margin/360000:.1f}cm")
        report.append(f"  右边距: {sec.right_margin/360000:.1f}cm")
        report.append("")

        # ---- 字体分析 ----
        h1_fonts={}; h2_fonts={}; h3_fonts={}; body_fonts={}; body_sizes={}
        for p in doc.paragraphs:
            if not p.runs or not p.text.strip(): continue
            try:
                r = p.runs[0]
                fn = r.font.name or "继承"
                fs = r.font.size
                fs_str = f"{fs/12700:.0f}pt" if fs else "继承"
                style = p.style.name
                if style == "Heading 1": h1_fonts[fn] = h1_fonts.get(fn, 0) + 1
                elif style == "Heading 2": h2_fonts[fn] = h2_fonts.get(fn, 0) + 1
                elif style == "Heading 3": h3_fonts[fn] = h3_fonts.get(fn, 0) + 1
                elif not style.startswith("Heading"):
                    body_fonts[fn] = body_fonts.get(fn, 0) + 1
                    body_sizes[fs_str] = body_sizes.get(fs_str, 0) + 1
            except: pass

        report.append("【字体统计】")
        def topf(d): return max(d, key=d.get) if d else "未检测"
        report.append(f"  一级标题(Heading 1): {topf(h1_fonts)}")
        report.append(f"  二级标题(Heading 2): {topf(h2_fonts)}")
        report.append(f"  三级标题(Heading 3): {topf(h3_fonts)}")
        report.append(f"  正文: {topf(body_fonts)} {topf(body_sizes)}")
        report.append("")

        # ---- 段落统计 ----
        all_p = [p for p in doc.paragraphs if p.text.strip()]
        headings = {}
        for p in all_p:
            s = p.style.name
            if s.startswith("Heading"):
                lvl = int(s.replace("Heading ", ""))
                headings[lvl] = headings.get(lvl, 0) + 1
        body = [p for p in all_p if not p.style.name.startswith("Heading")]
        total_chars = sum(len(p.text) for p in all_p)

        report.append("【内容统计】")
        report.append(f"  总段落数: {len(all_p)}")
        for lvl in sorted(headings):
            report.append(f"  标题级别{lvl}: {headings[lvl]}个")
        report.append(f"  正文段落: {len(body)}个")
        report.append(f"  总字符数: {total_chars} (约{total_chars//2}字)")
        report.append(f"  表格数量: {len(doc.tables)}")
        report.append("")

        # ---- 标题结构 ----
        report.append("【标题结构】")
        for p in all_p:
            if p.style.name.startswith("Heading"):
                lvl = int(p.style.name.replace("Heading ", ""))
                indent = "  " * (lvl - 1)
                report.append(f"  {indent}[H{lvl}] {p.text[:80]}")
        report.append("")

        # ---- GB/T 7713 对照检查 ----
        report.append("【GB/T 7713标准对照检查】")
        checks = []
        m = sec.top_margin / 360000
        checks.append(f"  {'[PASS]' if 2.0 <= m <= 3.0 else '[注意]'} 上边距 {m:.1f}cm (标准2.54cm)")
        bf = topf(body_fonts)
        checks.append(f"  {'[PASS]' if '宋体' in str(bf) or 'SimSun' in str(bf) else '[注意]'} 正文字体: {bf} (建议宋体)")
        hf = topf(h1_fonts) if h1_fonts else topf(h2_fonts)
        checks.append(f"  {'[PASS]' if '黑体' in str(hf) or 'SimHei' in str(hf) else '[注意]'} 标题字体: {hf} (建议黑体)")
        report.extend(checks)
        report.append("")
        report.append("=" * 60)
        report.append("  分析完成。点「导出格式配置」保存为JSON文件")
        report.append("  论文重写工具可导入此配置自动应用格式")

        self.report.delete("1.0", tk.END)
        self.report.insert("1.0", "\n".join(report))
        self.analysis_data = {
            "page": {"width_cm": tw, "height_cm": th, "top_margin_cm": sec.top_margin/360000,
                     "bottom_margin_cm": sec.bottom_margin/360000, "left_margin_cm": sec.left_margin/360000,
                     "right_margin_cm": sec.right_margin/360000},
            "fonts": {"heading1": topf(h1_fonts), "heading2": topf(h2_fonts),
                      "heading3": topf(h3_fonts), "body": topf(body_fonts), "body_size": topf(body_sizes)},
            "stats": {"total_paras": len(all_p), "headings": headings, "body_paras": len(body),
                      "total_chars": total_chars, "tables": len(doc.tables)}
        }

    def export_config(self):
        if not hasattr(self, 'analysis_data'):
            messagebox.showwarning("提示", "请先分析模板文件"); return
        with open(OUTPUT_CONFIG, "w", encoding="utf-8") as f:
            json.dump(self.analysis_data, f, indent=2, ensure_ascii=False)
        messagebox.showinfo("导出成功", f"格式配置已保存:\n{OUTPUT_CONFIG}\n\n论文重写工具会自动读取此配置")

def main():
    root = tk.Tk()
    TemplateAnalyzer(root)
    root.mainloop()

if __name__ == "__main__":
    main()
