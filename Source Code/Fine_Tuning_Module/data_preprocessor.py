import json
import os
from docx import Document
import re

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(os.path.dirname(SCRIPT_DIR))
DOCX_PATH = os.path.join(PROJECT_ROOT, "docs", "paper_with_feedback", "kp_for_fine_tuning", "84 compound verbs from NINJAL (Updated)_17Nov2025.docx")
OUTPUT_JSONL = os.path.join(PROJECT_ROOT, "docs", "paper_with_feedback", "kp_for_fine_tuning", "japanese_vocab_unsloth.jsonl")

def parse_docx(docx_path):
    doc = Document(docx_path)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return paragraphs

def chunk_entries(paragraphs):
    """
    假设你的 Expert 文档里，每个词条是以 “数字．【词汇 (假名)】” 开头，
    然后下面有解释 + 例句若干段。这个函数把它们分组成条目。
    """
    entries = []
    curr = None

    for para in paragraphs:
        # 更新检测逻辑：只匹配 "数字．【" 或 "数字.【" 格式的行作为新条目开头
        if re.match(r"^\d+[．.]\s*【", para):
            # 新条目开始
            if curr:
                entries.append(curr)
            curr = {"header": para, "body": []}
        else:
            # 属于当前条目内容
            if curr is None:
                # 如果还没遇到 header，就忽略
                continue
            curr["body"].append(para)
    # 最后一条
    if curr:
        entries.append(curr)
    return entries

def build_samples(entries):
    samples = []
    for e in entries:
        header = e["header"]
        # 将所有内容合并为一个完整的 output
        body = "\n".join(e["body"])
        
        # 从 header 提取词和读音 (假名)，简单处理
        # header 示例: “２．【思い出す (おもいだす)】”
        try:
            inside = header.split("【")[1].split("】")[0]
        except IndexError:
            continue # 如果 header 格式不正确，跳过此条目

        if "(" in inside and ")" in inside:
            word = inside.split("(")[0].strip()
            reading = inside.split("(")[1].strip(" )")
        else:
            word = inside.strip()
            reading = ""
        
        # 如果 word 为空，则跳过
        if not word:
            continue

        instruction = (
            "【日】次の日本語の複合動詞を、日本語学習者（JLPT N4〜N5レベル）にも分かるように、"
            "できるだけやさしい日本語で詳しく説明してください。語構造、動詞の種類（自動詞／他動詞など）、"
            "意味、そして2〜3個の用例文（それぞれ日文→中文→英文の順）を含めてください。\n"
            "【中】请用【日语】，面向日语学习者（JLPT N4〜N5），详细解释下面的日语复合动词，"
            "包括：1）语构造；2）动词种类（自动词/他动词等）；3）含义；4）2〜3个例句，"
            "并给出【日文例句 + 中文译文 + 英文译文】。"
        )
        input_text = f"{word} ({reading})"
        output_text = body

        sample = {
            "instruction": instruction,
            "input": input_text,
            "output": output_text
        }
        samples.append(sample)
    return samples

def write_jsonl(samples, out_path):
    with open(out_path, "w", encoding="utf-8") as f:
        for s in samples:
            f.write(json.dumps(s, ensure_ascii=False) + "\n")

def main():
    paras = parse_docx(DOCX_PATH)
    entries = chunk_entries(paras)
    samples = build_samples(entries)
    write_jsonl(samples, OUTPUT_JSONL)
    print(f"已写入 {len(samples)} 条样本到 {OUTPUT_JSONL}")

if __name__ == "__main__":
    main()
