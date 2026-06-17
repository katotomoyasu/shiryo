# make-testdata スキル

支払PGのパンチデータ（固定長ファイル・CSV）に BUMON（8文字）を追記したテストデータを生成する。

## トリガー条件
- 「テストデータを作って」「_testファイルを作成」「BUMONを追記したテスト用」などの依頼
- パンチデータファイル（k8xxxx 等）や検収反映データCSVに BUMON 列を追加したテスト版を作りたい場合

## ファイル種別の判定

| 種別 | 判定 | 処理方法 |
|------|------|---------|
| 固定長ファイル | 拡張子なし | `make_testdata.py` を使用 |
| CSVファイル | 拡張子 `.csv` | 各行末尾にカンマ＋8文字を追加（Pythonインライン処理） |

---

## 固定長ファイルの場合

```
python "C:\Users\KATOTO\Downloads\その他\支払ＰＧ本番_mcframe\bumon_agent\make_testdata.py" <入力ファイル> [出力ファイル]
```

- 出力ファイル省略時は `<元ファイル名>_test` として同フォルダに出力
- レコード長は CRLF 位置から自動検出
- BUMON 値は入力ファイル名（拡張子なし）の先頭8文字を CP932 でエンコードし右パディング

---

## CSVファイルの場合

Pythonインラインで処理する。BUMON値は適当なアルファベット＋数字8文字（例: `BU001234`）。

```python
src = r'<入力CSVパス>'
dst = r'<出力CSVパス>'

with open(src, 'r', encoding='cp932') as f:
    lines = f.readlines()

out_lines = []
for line in lines:
    stripped = line.rstrip('\r\n')
    if stripped:
        out_lines.append(stripped + ',BU001234\r\n')
    else:
        out_lines.append(line)

with open(dst, 'w', encoding='cp932', newline='') as f:
    f.writelines(out_lines)
```

---

## 注意

- 本番ファイルは変更しない（入力ファイルは読み取り専用で扱う）
- 出力ファイル名は katochan-file-rules に従い `YYYYMMDD_案件名_文書種類_vN.拡張子` 形式にする
