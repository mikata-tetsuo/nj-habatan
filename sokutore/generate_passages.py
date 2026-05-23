#!/usr/bin/env python3
"""
Claude API を使って英検速読トレーナー用パッセージを自動生成するスクリプト。

使い方:
  python generate_passages.py --grade 5 --count 10
  python generate_passages.py --grade all --count 5   # 各級5問ずつ
  python generate_passages.py --grade 3 --count 3 --theme 健康・医療

生成後は embed_passages.py を実行して HTML に反映してください。
"""
import argparse
import itertools
import json
import os
import re
import sys

import anthropic
import openpyxl

try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"), override=True)
except ImportError:
    pass

# ------------------------------------------------------------------ 級仕様
GRADE_SPECS = {
    "5": {
        "name": "英検5級",
        "word_count": "40〜60語",
        "target_time": 28,
        "eiken_grade": "5",
        "sub_level": 1,
        "description": "小学生レベル。日常的な話題（家族・食事・学校・ペット等）。",
        "vocab": "基本単語のみ。be動詞・一般動詞・形容詞の単純な文。",
    },
    "4": {
        "name": "英検4級",
        "word_count": "70〜100語",
        "target_time": 42,
        "eiken_grade": "4",
        "sub_level": 1,
        "description": "中学生レベル。学校行事・旅行・趣味・地域の話題等。",
        "vocab": "中学基本語彙。現在完了・受動態が自然に使われる程度。",
    },
    "3": {
        "name": "英検3級",
        "word_count": "100〜130語",
        "target_time": 55,
        "eiken_grade": "3",
        "sub_level": 1,
        "description": "中学卒業レベル。社会問題（環境・テクノロジー・文化等）の入門的内容。",
        "vocab": "中学修了程度の語彙。接続詞や関係代名詞を使った文が多い。",
    },
    "pre2": {
        "name": "英検準2級",
        "word_count": "130〜160語",
        "target_time": 72,
        "eiken_grade": "pre2",
        "sub_level": 1,
        "description": "高校中級レベル。科学・社会・文化などの一般的なトピック。",
        "vocab": "高校中級語彙。譲歩・対比・因果の構造が頻出。",
    },
    "pre2p": {
        "name": "英検準2級+（準2〜2級の橋渡し）",
        "word_count": "160〜190語",
        "target_time": 85,
        "eiken_grade": "pre2",
        "sub_level": 2,
        "description": "準2〜2級の橋渡しレベル。社会問題・ビジネス・科学等のやや高度な内容。",
        "vocab": "高校上級語彙。複雑な文構造と抽象的な概念が登場。",
    },
    "2": {
        "name": "英検2級",
        "word_count": "180〜220語",
        "target_time": 93,
        "eiken_grade": "2",
        "sub_level": 1,
        "description": "高校卒業レベル。学術的な話題（心理・経済・環境・医療等）。",
        "vocab": "大学入試レベルの語彙。学術的な論述スタイル。",
    },
    "pre1": {
        "name": "英検準1級",
        "word_count": "220〜260語",
        "target_time": 115,
        "eiken_grade": "pre1",
        "sub_level": 1,
        "description": "大学中級レベル。専門性の高いトピック（認知科学・政策・環境経済等）。",
        "vocab": "上級語彙・学術専門用語。複雑な論理構造。",
    },
}

THEMES = [
    "環境・気候変動",
    "テクノロジー・AI",
    "健康・医療・科学",
    "教育・学習",
    "社会問題・格差",
    "経済・ビジネス",
    "文化・歴史",
    "食・農業",
    "スポーツ・芸術",
    "宇宙・自然",
    "心理・行動科学",
    "エネルギー・インフラ",
    "グローバル化・国際関係",
    "言語・コミュニケーション",
    "動物・生態系",
]

HEADERS = [
    "id", "type", "eiken_grade", "sub_level", "passage", "japanese_translation",
    "question", "choice_a", "choice_b", "choice_c", "choice_d", "answer",
    "target_time_sec", "reading_point", "evidence_text", "reward_point",
]

PROMPT = """\
英検速読トレーナー用の読解問題を1問作成してください。

## 仕様
- レベル: {name}
- 語数: {word_count}
- テーマ: {theme}
- 内容の難易度: {description}
- 語彙レベル: {vocab}

## 出力形式（JSONのみ。説明文・コードブロック不要）
{{
  "passage": "英文パッセージ",
  "japanese_translation": "日本語訳（原文に忠実に）",
  "question": "英語の問題文（What / Why / How 等で始まる）",
  "choice_a": "選択肢A",
  "choice_b": "選択肢B",
  "choice_c": "選択肢C",
  "choice_d": "選択肢D",
  "answer": "正解（a / b / c / d のいずれか1文字・小文字）",
  "reading_point": "日本語での読解ヒント（どこに注目すればよいか）",
  "evidence_text": "答えの根拠となる本文の一節（英語）"
}}

## 制約
- 不正解の選択肢は「本文にない情報」または「本文と逆の内容」にする
- reading_point は日本語で書く（英語不可）
- answer は小文字1文字（a / b / c / d）
- テーマはあくまでもヒント。自然な英文になるなら多少ずれてもよい
"""


# ------------------------------------------------------------------ ユーティリティ
def load_excel(path: str) -> list[dict]:
    if not os.path.exists(path):
        return []
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if len(rows) < 2:
        return []
    headers = [str(h).strip() if h else f"col{i}" for i, h in enumerate(rows[0])]
    data = []
    for row in rows[1:]:
        if not row[0]:
            continue
        d = {h: (row[i] if i < len(row) else "") for i, h in enumerate(headers)}
        data.append(d)
    return data


def save_excel(path: str, all_data: list[dict]):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "passages"
    ws.append(HEADERS)
    for d in all_data:
        ws.append([d.get(h, "") for h in HEADERS])
    wb.save(path)


def parse_json(text: str) -> dict | None:
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r'\{.*\}', text, re.DOTALL)
        if m:
            try:
                return json.loads(m.group())
            except json.JSONDecodeError:
                pass
    return None


# ------------------------------------------------------------------ 生成
def generate_one(client: anthropic.Anthropic, grade_key: str, theme: str,
                 next_id: int) -> dict | None:
    spec = GRADE_SPECS[grade_key]
    prompt = PROMPT.format(theme=theme, **spec)

    try:
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=2048,
            messages=[{"role": "user", "content": prompt}],
        )
    except Exception as e:
        print(f"[API ERROR] {e}")
        return None

    data = parse_json(response.content[0].text.strip())
    if data is None:
        print("[ERROR] JSON解析失敗")
        return None

    required = ["passage", "japanese_translation", "question",
                "choice_a", "choice_b", "choice_c", "choice_d",
                "answer", "reading_point", "evidence_text"]
    missing = [f for f in required if f not in data]
    if missing:
        print(f"[ERROR] フィールド不足: {missing}")
        return None

    data.update({
        "id": next_id,
        "type": "reading",
        "eiken_grade": spec["eiken_grade"],
        "sub_level": spec["sub_level"],
        "target_time_sec": spec["target_time"],
        "reward_point": 10,
    })
    return data


# ------------------------------------------------------------------ メイン
def main():
    parser = argparse.ArgumentParser(description="英検速読トレーナー 問題自動生成")
    parser.add_argument("--grade", required=True,
                        help="生成する級: 5 / 4 / 3 / pre2 / pre2p / 2 / pre1 / all")
    parser.add_argument("--count", type=int, default=5,
                        help="1つの級に生成する問題数 (default: 5)")
    parser.add_argument("--theme", default=None,
                        help="テーマを固定したい場合に指定（省略時はローテーション）")
    parser.add_argument("--output", default="英検速読トレーナー_DB_v1.xlsx",
                        help="出力先 Excel ファイル")
    args = parser.parse_args()

    grades = list(GRADE_SPECS.keys()) if args.grade == "all" else [args.grade]
    for g in grades:
        if g not in GRADE_SPECS:
            print(f"[ERROR] 不明な級: {g}  有効値: {list(GRADE_SPECS.keys())}")
            sys.exit(1)

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("[ERROR] ANTHROPIC_API_KEY が設定されていません")
        print("  sokutore/.env に ANTHROPIC_API_KEY=xxx を記入してください")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)
    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), args.output)
    existing = load_excel(output_path)

    new_passages: list[dict] = []
    theme_iter = itertools.cycle([args.theme] if args.theme else THEMES)
    def to_int(v):
        try: return int(v)
        except: return 0
    next_id = max((to_int(d.get("id", 0)) for d in existing), default=-1) + 1

    for grade_key in grades:
        print(f"\n=== {GRADE_SPECS[grade_key]['name']} ({args.count}問) ===")
        for i in range(args.count):
            theme = next(theme_iter)
            print(f"  [{i+1:>3}/{args.count}] {theme} ... ", end="", flush=True)
            result = generate_one(client, grade_key, theme, next_id)
            if result:
                new_passages.append(result)
                next_id += 1
                print(f"OK  (id={result['id']})")
            else:
                print("スキップ")

    if not new_passages:
        print("\n生成された問題がありませんでした。")
        return

    all_data = existing + new_passages
    save_excel(output_path, all_data)
    print(f"\n[完了] {len(new_passages)}問を追記 -> {output_path}")
    print("次のコマンドで HTML に反映してください:")
    print("  python embed_passages.py")


if __name__ == "__main__":
    main()
