"""
英検速読トレーナー_DB_v1.xlsx から問題データを読み込み、
sokutore.html の PASSAGES 配列を差し替えるスクリプト。
列順: id, type, eiken_grade, sub_level, passage, japanese_translation,
      question, choice_a, choice_b, choice_c, choice_d, answer,
      target_time_sec, reading_point, evidence_text, reward_point
1行目がヘッダー
"""
import openpyxl, json, re

EXCEL = "英検速読トレーナー_DB_v1.xlsx"
HTML  = "sokutore.html"

print("Excel読み込み中...")
wb = openpyxl.load_workbook(EXCEL, read_only=True)
ws = wb.active

rows = list(ws.iter_rows(values_only=True))
headers = [str(h).strip() if h else f"col{i}" for i, h in enumerate(rows[0])]

passages = []
for row in rows[1:]:
    if not row[0]:
        continue
    d = {}
    for j, h in enumerate(headers):
        v = row[j] if j < len(row) else None
        d[h] = v if v is not None else ""
    # 型変換
    for int_key in ('id', 'sub_level', 'target_time_sec', 'reward_point'):
        try: d[int_key] = int(d.get(int_key, 0))
        except: d[int_key] = 0
    d['eiken_grade'] = str(d.get('eiken_grade', '')).strip()
    for str_key in ('type','passage','japanese_translation','question',
                    'choice_a','choice_b','choice_c','choice_d',
                    'answer','reading_point','evidence_text'):
        d[str_key] = str(d.get(str_key, '')).strip()
    passages.append(d)

wb.close()
print(f"  {len(passages)}問 読み込み完了")

print("HTML更新中...")
with open(HTML, encoding='utf-8') as f:
    content = f.read()

pattern = r'(const PASSAGES\s*=\s*)(\[.*?\])(;)'
m = re.search(pattern, content, re.DOTALL)
if not m:
    print("ERROR: PASSAGES配列が見つかりません")
    exit(1)

new_json = json.dumps(passages, ensure_ascii=False, separators=(',',':'))
new_content = content[:m.start()] + m.group(1) + new_json + m.group(3) + content[m.end():]

with open(HTML, 'w', encoding='utf-8') as f:
    f.write(new_content)

print(f"完了！ sokutore.html を更新しました（{len(passages)}問）")
