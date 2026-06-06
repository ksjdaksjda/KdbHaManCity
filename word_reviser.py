"""
成品论文Word修改工具 - 保留原排版，只改内容
双击运行，需要 Python + DeepSeek API Key
"""
import os, sys, json, time, re

def check_deps():
    try: from docx import Document; from openai import OpenAI
    except ImportError:
        print("缺少依赖，正在安装...")
        os.system("pip install python-docx openai -q")
        try: from docx import Document; from openai import OpenAI
        except: print("安装失败，请手动: pip install python-docx openai"); sys.exit(1)

check_deps()
from docx import Document
from openai import OpenAI

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "revise_config.json")

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f: return json.load(f)
    return {}
def save_config(cfg):
    with open(CONFIG_FILE, "w") as f: json.dump(cfg, f, indent=2)

def main():
    print("=" * 60)
    print("  成品论文Word修改 - 保留原排版，AI修改内容")
    print("=" * 60)
    print()

    config = load_config()
    api_key = config.get("api_key", "")
    if not api_key:
        api_key = input("请输入 DeepSeek API Key: ").strip()
        if api_key: config["api_key"] = api_key; save_config(config)
    if not api_key: print("未提供API Key，退出。"); return

    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com/v1")

    # 验证API
    print("验证API...")
    try:
        r = client.chat.completions.create(model="deepseek-chat", messages=[{"role":"user","content":"OK"}], max_tokens=5)
        print(f"  API连接成功 (模型: {r.model})")
    except Exception as e:
        print(f"  API连接失败: {e}")
        return

    input_file = input("\n请输入Word文件路径: ").strip().strip('"').strip("'")
    if not os.path.exists(input_file): print(f"文件不存在: {input_file}"); return

    print(f"\n读取: {input_file}")
    doc = Document(input_file)

    # 统计
    all_paras = [p for p in doc.paragraphs if p.text.strip()]
    headings = [p for p in all_paras if p.style.name.startswith("Heading")]
    body = [p for p in all_paras if not p.style.name.startswith("Heading")]
    print(f"  总段落:{len(all_paras)}  标题:{len(headings)}  正文:{len(body)}")

    print("\n修改模式:")
    print("  1. 优化语言和学术表达")
    print("  2. 降重改写")
    print("  3. 扩展补充内容")
    print("  4. 不改标题，只改正文")
    choice = input("选择(默认1): ").strip() or "1"

    modes = {
        "1":"优化语言表达，使其更学术化、专业。保持原意和字数基本不变。只输出修改后的段落文本，不要加任何说明。",
        "2":"深度改写降重。变换句式、同义词替换。保持核心信息不变。只输出修改后的段落文本，不要加任何说明。",
        "3":"适当扩展内容，补充细节和论证。只输出修改后的段落文本，不要加任何说明。",
        "4":"优化语言表达。只输出修改后的段落文本，不要加任何说明。"
    }
    mode_desc = modes.get(choice, modes["1"])
    skip_headings = (choice == "4")

    # 确定要修改的段落
    targets = body if skip_headings else all_paras
    print(f"\n将修改 {len(targets)} 个段落")
    confirm = input("继续?(y/n): ").strip().lower()
    if confirm != "y": print("已取消"); return

    # ========== 逐段修改 ==========
    modified = 0
    for idx, para in enumerate(targets):
        text = para.text.strip()
        if len(text) < 10: continue

        print(f"  [{idx+1}/{len(targets)}] {text[:40]}...", end=" ", flush=True)

        system_prompt = f"你是学术论文编辑专家。{mode_desc}"
        user_prompt = f"请修改以下论文段落：\n\n{text}"

        try:
            resp = client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role":"system","content":system_prompt},
                    {"role":"user","content":user_prompt}
                ],
                temperature=0.7, max_tokens=2048,
            )
            new_text = resp.choices[0].message.content
            if new_text and len(new_text.strip()) > 5:
                # 清理AI可能加的乱七八糟前缀
                new_text = new_text.strip()
                # 保留原始段落中的runs和格式
                _replace_para_text(para, new_text)
                modified += 1
                print(f"✅")
            else:
                print("⏭ (无内容)")
        except Exception as e:
            print(f"❌ {e}")

        time.sleep(0.5)  # 避免限流

    # 保存
    output_file = input_file.replace(".docx", "_修改版.docx")
    if os.path.exists(output_file):
        base = input_file.replace(".docx", "")
        output_file = f"{base}_修改版{int(time.time())%10000}.docx"
    doc.save(output_file)

    print(f"\n{'='*60}")
    print(f"  修改完成！共 {modified}/{len(targets)} 段")
    print(f"  输出: {output_file}")
    print(f"{'='*60}")
    os.startfile(os.path.dirname(output_file))

def _replace_para_text(para, new_text):
    """替换段落文字，保留原有格式(runs)"""
    runs = para.runs
    if runs:
        # 保留第一个run，清空后续runs
        for run in runs[1:]:
            run.text = ""
        runs[0].text = new_text
    else:
        para.text = new_text

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n发生错误: {e}")
        import traceback
        traceback.print_exc()
    print("\n按回车键退出...")
    input()
