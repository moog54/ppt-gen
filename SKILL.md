# /slides — 汎用PPTXスライド生成スキル

## 概要

自然言語の指示から、ブランド制約付きPPTXプレゼンテーションを生成するスキル。

**トリガー**: `/slides` または「〇〇のスライドを作って」「PPTを生成して」などの意図で起動

---

## 環境

```
プロジェクト: /mnt/c/Users/moogs/work/ppt-gen
Python: /home/moog/hr_venv/bin/python
出力先: /mnt/c/Users/moogs/work/ppt-gen/output/
```

---

## ワークフロー（6ステップ）

### Step 1: 要件確認

ユーザーの指示から以下を把握する:
- **目的**: 何のためのプレゼンか（事業報告 / 提案 / 定例報告 / ピッチ等）
- **スライド数**: 何枚か（未指定の場合は目的から推定）
- **テーマカラー**: 指定があれば反映（未指定は `0070C0` デフォルト）
- **内容**: 各スライドに何を載せるか

### Step 2: パターン選定

`PATTERN_CATALOG.md` を参照して適切なパターンを選択する。

よくある組み合わせ:
- 事業報告: `cover → agenda → data_kpi → data_table → flow_timeline → closing`
- 提案書: `cover → quote → body_3col → data_comparison → flow_process → closing`
- 定例報告: `cover → data_kpi → data_table → closing`

### Step 3: コード作成

`CLAUDE.md` を参照してコードを作成する。

```python
# 雛形
import sys
sys.path.insert(0, "/mnt/c/Users/moogs/work/ppt-gen")
from lib import SlideBuilder, add_text, add_kpi_row, CONTENT_TOP, ...

sb = SlideBuilder(theme={"accent": "RRGGBB"})
sb.set_footer("ページ下部テキスト")

# --- 1. 表紙 ---
sb.add_cover("タイトル", "サブタイトル", "年月日")

# --- N. 各スライド ---
slide = sb.add_body("スライドタイトル", section="セクション名")
# ... コンテンツ追加 ...

# --- 最終. 保存 ---
result = sb.save_and_validate("output/FILENAME.pptx")
```

### Step 4: 実行

```bash
cd /mnt/c/Users/moogs/work/ppt-gen
/home/moog/hr_venv/bin/python output_script_or_inline.py
```

### Step 5: バリデーション確認

実行後の出力を確認:
- `ERROR` が出た場合: コードを修正して再実行
- `WARN` のみ: 警告内容をユーザーに報告して完了
- エラーなし: 完了

### Step 6: 結果報告

生成されたファイルパスと、スライド構成のサマリーをユーザーに報告する。

---

## 主要制約（厳守）

1. **pptx native table 禁止** → `add_table_shapes()` を使う
2. **游ゴシック禁止** → `Meiryo UI`（自動設定）
3. **コンテンツははみ出さない** → `x + w ≤ 12.9` / `y + h ≤ 7.0`
4. **コンテンツは CONTENT_TOP (1.4") 以下** → `y = CONTENT_TOP + 0.2 = 1.6` から開始

---

## クイックリファレンス

```python
from lib import (
    SlideBuilder,
    # テキスト
    add_text, add_rich_text,
    # 図形
    add_rect, add_rounded_rect, add_line, add_arrow,
    add_card, add_pill, add_badge, add_quote,
    # データ
    add_kpi_row, add_table_shapes, add_comparison, add_bar_chart_img,
    # フロー
    add_flow_row, add_timeline, add_agenda,
    # ヘッダー・フッター
    add_slide_header, add_slide_footer,
    # 定数
    C, CONTENT_TOP, SLIDE_W, SLIDE_H, TEXT_STYLES,
    # バリデーション
    validate, ValidationResult,
)
```

---

## 使用例

### 「Q1事業報告スライドを5枚作って」

→ Step 1: 事業報告、5枚、デフォルトカラー
→ Step 2: `cover → data_kpi → data_table → flow_timeline → closing`
→ Step 3-4: コード作成・実行
→ Step 6: `output/q1_business_report.pptx` を報告

### 「ピッチデッキを青紫テーマで10枚」

→ `theme={"accent": "5C2D91"}` を設定
→ `cover → quote → body_3col → data_kpi → data_comparison → flow_process → flow_timeline → data_table → body_1col → closing`

### 「examples/business_report.py を参考に〇〇向けに修正して」

→ `examples/business_report.py` を読んで内容を把握してから修正

---

## トラブルシューティング

| エラー | 原因 | 対処 |
|---|---|---|
| `right_edge > 12.9"` | テキストボックス幅が大きすぎ | `w` を小さくして `x + w ≤ 12.8` に |
| `bottom > 7.0"` | 下端はみ出し | `h` を減らすか `y` を上げる |
| `native table` | `prs.slides.add_table()` 使用 | `add_table_shapes()` に変換 |
| `Permission denied` | Windows で PPTX が開かれている | ファイルを閉じて再実行 |
