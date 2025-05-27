import json
from openpyxl import Workbook, load_workbook
from collections import defaultdict
import sys
import os

# コマンドライン引数から入力ファイルを取得
if len(sys.argv) < 2:
    print("Usage: python sript.py <input_json_file>")
    sys.exit(1)
input_path = sys.argv[1]

# 出力ファイル名を決定
base, _ = os.path.splitext(input_path)
output_path = base + '.xlsx'

# JSONファイルの読み込み
with open(input_path, 'r', encoding='utf-8') as f:
    data = json.load(f)

# profileの全項目・階層を収集
profile_keys = set()
profile_values = defaultdict(set)

def make_hashable(val):
    if isinstance(val, (list, tuple)):
        return tuple(make_hashable(v) for v in val)
    elif isinstance(val, dict):
        return tuple(sorted((k, make_hashable(v)) for k, v in val.items()))
    else:
        return val

def collect_keys(d, path=()):
    for k, v in d.items():
        new_path = path + (k,)
        is_dict = isinstance(v, dict)
        profile_keys.add((new_path, len(new_path), k, is_dict))
        if is_dict:
            collect_keys(v, new_path)
        else:
            try:
                hashable_v = make_hashable(v)
            except Exception:
                hashable_v = str(v)
            profile_values[new_path].add(hashable_v)

for user in data:
    collect_keys(user['profile'])

# profile_keysを階層順・値→親dict→子dictの順で並べる（値が先、dict型は後）
profile_keys = sorted(profile_keys, key=lambda x: (x[1], x[3], x[0]))

# Excelワークブックの作成（テンプレート利用）
wb = load_workbook('template/template.xlsx')
ws = wb.active  # 1枚目のシートを利用

# 書き込み開始位置
start_row = 6
start_col = 1  # A列

# ユーザーIDリスト
ids = [user['id'] for user in data]

# ヘッダー
header = ["項目名", "設定値"] + ids
for col, val in enumerate(header, start_col):
    ws.cell(row=start_row, column=col, value=val)

# 各profile項目ごとに行を作成
prev_disp_name = None
row_idx = start_row + 1
for key_path, depth, key, is_dict in profile_keys:
    # 表示用のインデント付き項目名
    disp_name = '  ' * (depth-1) + key
    values = sorted(profile_values[key_path])
    if not values:
        row_disp_name = disp_name if prev_disp_name != disp_name else ''
        row = [row_disp_name, ''] + ['' for _ in ids]
        for col, val in enumerate(row, start_col):
            ws.cell(row=row_idx, column=col, value=val)
        prev_disp_name = disp_name
        row_idx += 1
    else:
        for value in values:
            # 値の表示用（リストやdictはstrで）
            if isinstance(value, tuple) and any(isinstance(v, tuple) for v in value):
                disp_value = str(value)
            else:
                disp_value = value if not isinstance(value, tuple) else str(value)
            row_disp_name = disp_name if prev_disp_name != disp_name else ''
            row = [row_disp_name, disp_value]
            for user in data:
                v = user['profile']
                try:
                    for part in key_path:
                        v = v[part]
                    user_hashable = make_hashable(v)
                except (KeyError, TypeError):
                    user_hashable = None
                if user_hashable == value:
                    row.append('◯')
                else:
                    row.append('')
            for col, val in enumerate(row, start_col):
                ws.cell(row=row_idx, column=col, value=val)
            prev_disp_name = disp_name
            row_idx += 1

# xlsxファイルとして保存
wb.save(output_path)
