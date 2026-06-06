"""
成品论文Word修改工具 - 保留原排版，只改内容
双击运行，需要 Python 和 DeepSeek API Key
"""
import os, sys, json, copy, time, re
from datetime import datetime

try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    print("请先安装 python-docx: pip install python-docx openai")
    sys.exit(1)

try:
    from openai import OpenAI
except ImportError:
    print("请先安装 openai: pip install openai")
    sys.exit(1)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "revise_config.json")

# ========== 配置 ==========
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {}

def save_config(cfg):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)


def main():
    print("=" * 60)
    print("  成品论文Word修改 - 保留原排版，AI修改内容")
    print("=" * 60)
    print()

    config = load_config()

    # API Key
    api_key = config.get("api_key", "")
    if not api_key:
        api_key = input("请输入 DeepSeek API Key: ").strip()
        if api_key:
            config["api_key"] = api_key
            save_config(config)

    if not api_key:
        print("未提供API Key，退出。")
        return

    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com/v1")

    # 输入文件
    input_file = input("请输入成品论文Word文件路径（或拖拽文件到此处）: ").strip().strip('"').strip("'")
    if not os.path.exists(input_file):
        print(f"文件不存在: {input_file}")
        return

    print(f"\n正在读取: {input_file}")

    # 读取文档
    doc = Document(input_file)

    # 统计
    total_paras = len(doc.paragraphs)
    non_empty = [p for p in doc.paragraphs if p.text.strip()]
    headings = [p for p in doc.paragraphs if p.style.name.startswith("Heading")]
    tables = doc.tables

    print(f"  段落总数: {total_paras}")
    print(f"  有内容段落: {len(non_empty)}")
    print(f"  标题: {len(headings)}")
    print(f"  表格: {len(tables)}")
    print()

    # 选择模式
    print("修改模式:")
    print("  1. 优化语言和学术表达（保留结构）")
    print("  2. 降重改写（保留结构）")
    print("  3. 扩展补充内容（保留结构）")
    print("  4. 仅修改正文，不改标题")
    choice = input("选择 (1-4, 默认1): ").strip() or "1"
    print()

    modes = {
        "1": "优化语言和学术表达，使其更学术化、更专业。保持段落结构和字数基本不变。",
        "2": "深度改写降重。变化句式结构、同义词替换。保持核心信息不变。",
        "3": "适当扩展内容，补充细节和论证。在原段落基础上丰富。",
        "4": "优化语言和学术表达。**注意：是一级标题(Heading 1)、二级标题(Heading 2)的段落不要修改，保持原标题。只改正文段落。**"
    }
    mode_desc = modes.get(choice, modes["1"])
    skip_headings = (choice == "4")

    # 确认
    confirm = input(f"将修改约 {len(non_empty)} 个段落，是否继续？(y/n): ").strip().lower()
    if confirm != "y":
        print("已取消。")
        return

    # ========== 逐段处理 ==========
    modified_count = 0
    para_texts = []  # 收集所有要修改的文本
    para_map = []    # 记录哪些段落要处理

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        is_heading = para.style.name.startswith("Heading")
        if skip_headings and is_heading:
            print(f"  [{i+1}/{total_paras}] ⏭ 跳过标题: {text[:50]}...")
            continue

        para_texts.append(text)
        para_map.append(i)

    print(f"\n实际修改 {len(para_texts)} 个段落\n")

    # 分批处理（每批最多20段，避免token超限）
    batch_size = 15
    for batch_start in range(0, len(para_texts), batch_size):
        batch_end = min(batch_start + batch_size, len(para_texts))
        batch_texts = para_texts[batch_start:batch_end]

        print(f"[批次 {batch_start//batch_size + 1}] 处理 {len(batch_texts)} 段...")

        # 构建提示
        segments = "\n\n---SEGMENT---\n\n".join(
            f"[段落{i+1}]\n{t}" for i, t in enumerate(batch_texts)
        )

        system_prompt = f"""你是学术论文编辑专家。
请对以下{len(batch_texts)}个论文段落进行修改。
修改要求: {mode_desc}

规则:
1. 保持每个段落的编号 [段落N] 不变
2. 用 ---SEGMENT--- 分隔各段落
3. 严格保持段落顺序和数量一致
4. 不要添加额外的说明文字
5. 保留原文中的引用标记 [N] 和公式标记"""

        try:
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": segments}
                ],
                temperature=0.7,
                max_tokens=8192,
            )
            result = response.choices[0].message.content or ""

            # 解析结果
            modified_segments = result.split("---SEGMENT---")
            # 提取 [段落N] 后面的内容
            for i, seg in enumerate(modified_segments):
                # 去掉 [段落N] 标记
                cleaned = re.sub(r'^\[段落\d+\]\s*', '', seg.strip()).strip()
                if cleaned and batch_start + i < len(para_map):
                    orig_idx = para_map[batch_start + i]
                    doc.paragraphs[orig_idx].text = cleaned
                    modified_count += 1

            print(f"  ✅ 完成")
            time.sleep(1)  # 避免API限流

        except Exception as e:
            print(f"  ❌ 失败: {e}")
            continue

    # ========== 保存 ==========
    output_file = input_file.replace(".docx", "_修改版.docx")
    if os.path.exists(output_file):
        name, ext = os.path.splitext(output_file)
        output_file = f"{name}_{datetime.now().strftime('%H%M%S')}{ext}"

    doc.save(output_file)
    print(f"\n{'=' * 60}")
    print(f"  ✅ 修改完成！")
    print(f"  修改段落: {modified_count}")
    print(f"  输出文件: {output_file}")
    print(f"{'=' * 60}")

    # 打开文件夹
    os.startfile(os.path.dirname(output_file))


if __name__ == "__main__":
    main()
