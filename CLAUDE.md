# ppt-gen — AI向け API リファレンス

## プロジェクト概要

`lib.py` の API を使って `SlideBuilder` でスライドを構築し PPTX を出力するツール。
**ブランド制約**と**バリデーション**が組み込まれており、`save_and_validate()` で自動チェックが走る。

### 実行環境

```bash
# Python venv
/home/moog/hr_venv/bin/python

# 作業ディレクトリ
cd /mnt/c/Users/moogs/work/ppt-gen
```

---

## 基本ワークフロー

```python
import sys
sys.path.insert(0, "/mnt/c/Users/moogs/work/ppt-gen")
from lib import *

# 1. SlideBuilder 初期化
sb = SlideBuilder(theme={"accent": "0070C0"})  # テーマカラー上書き可
sb.set_footer("報告書タイトル")

# 2. スライド追加
sb.add_cover("タイトル", "サブタイトル", "2026年4月")
slide = sb.add_body("スライドタイトル", section="セクション名")
add_text(slide, x=0.5, y=1.5, w=12.0, h=1.0, text="本文テキスト")

# 3. 保存 + バリデーション
result = sb.save_and_validate("output/my_presentation.pptx")
```

---

## カラー定数 `C`

| キー | 値 | 用途 |
|---|---|---|
| `accent` | `0070C0` | メインアクセント（上書き可） |
| `accent2` | `00B050` | サブアクセント（緑） |
| `accent3` | `FF6B00` | 警告・強調（オレンジ） |
| `text` | `1A1A1A` | 本文テキスト |
| `textLight` | `666666` | 補足テキスト |
| `textMuted` | `999999` | キャプション |
| `bg` | `FFFFFF` | 白背景 |
| `bgLight` | `F5F5F5` | 薄グレー背景 |
| `border` | `D9D9D9` | ボーダー |
| `headerBg` | `0070C0` | ヘッダー背景 |
| `headerText` | `FFFFFF` | ヘッダーテキスト |
| `rowAlt` | `EFF4FB` | テーブル交互行 |

カラー引数には **キー名** または **`RRGGBB` / `#RRGGBB` の16進数** が使える。

---

## テキストスタイル `TEXT_STYLES`

| スタイル名 | フォントサイズ | 太字 | 用途 |
|---|---|---|---|
| `heading` | 24pt | ✓ | スライドタイトル |
| `subheading` | 18pt | ✓ | セクション見出し |
| `body` | 14pt | — | 本文 |
| `small` | 11pt | — | テーブル・補足 |
| `caption` | 10pt | — | キャプション・フッター |
| `label` | 11pt | ✓ | ラベル |
| `kpi` | 40pt | ✓ | 数値KPI |
| `title_cover` | 36pt | ✓ | 表紙タイトル |
| `subtitle_cover` | 18pt | — | 表紙サブタイトル |

---

## レイアウト定数

```python
SLIDE_W = 13.33    # スライド幅（インチ）16:9
SLIDE_H = 7.5      # スライド高さ（インチ）
CONTENT_TOP = 1.4  # ヘッダーバンド下端 — コンテンツはここ以下に配置
SAFE_XMAX = 12.9   # コンテンツ右端上限
SAFE_YMAX = 7.0    # コンテンツ下端上限（フッター除く）
```

---

## SlideBuilder メソッド

### `__init__(theme=None, master_path=None)`

```python
sb = SlideBuilder()
sb = SlideBuilder(theme={"accent": "A100FF"})  # テーマカラー上書き
```

### `add_cover(title, subtitle="", date="", bg_color=None) → Slide`

表紙スライドを追加。フルカラー背景。

### `add_section(title, subtitle="", number=None, color=None) → Slide`

セクション区切りスライド（左帯＋番号）。

### `add_body(title, section="", color=None) → Slide`

**コンテンツスライドのベース**。ヘッダーバンド＋フッターを追加して Slide を返す。
返った `slide` に `add_*` 関数でコンテンツを追加する。

```python
slide = sb.add_body("スライドタイトル")
add_text(slide, 0.5, CONTENT_TOP + 0.2, 12.0, 1.0, "テキスト")
```

### `save(path)` / `save_and_validate(path) → ValidationResult`

`save_and_validate` はバリデーション結果を返す。エラー時も保存は完了する。

---

## ヘルパー関数 — テキスト

### `add_text(slide, x, y, w, h, text, style="body", align="left", color=None, bold=None, font_size=None)`

```python
add_text(slide, 0.5, 1.5, 12.0, 0.5, "本文テキスト", style="body")
add_text(slide, 0.5, 2.0, 6.0, 0.4, "見出し", style="heading", color="accent")
```

### `add_rich_text(slide, x, y, w, h, paragraphs)`

```python
add_rich_text(slide, 0.5, 1.5, 12.0, 2.0, [
    {"text": "見出し段落", "style": "heading", "align": "left"},
    {"text": "本文段落", "style": "body", "space_before": 8},
])
```

---

## ヘルパー関数 — 図形

### `add_rect(slide, x, y, w, h, fill=None, border=None, border_width=1.0)`

### `add_rounded_rect(slide, x, y, w, h, fill=None, border=None)`

### `add_line(slide, x1, y1, x2, y2, color="border", width=1.0)`

### `add_arrow(slide, x1, y1, x2, y2, color="accent", width=2.0)`

### `add_card(slide, x, y, w, h, title="", body="", accent=None, bg="bg")`

左アクセントバー付きカード。

### `add_pill(slide, x, y, text, color="accent", text_color="white")`

### `add_badge(slide, x, y, number, color="accent")`

丸い番号バッジ。

### `add_quote(slide, x, y, w, text, source="")`

左ライン付き引用ブロック。

---

## ヘルパー関数 — データ可視化

### `add_kpi_row(slide, y, items, x_start=0.5, total_w=12.33)`

```python
add_kpi_row(slide, y=1.6, items=[
    {"label": "売上高", "value": "12.4", "unit": "億円", "delta": "+8%", "color": "accent"},
    {"label": "顧客数", "value": "4,820", "unit": "社"},
])
```

### `add_table_shapes(slide, x, y, w, headers, rows, col_widths=None, row_h=0.38)`

**pptx native table 禁止** のため、図形で構築する。

```python
add_table_shapes(
    slide, x=0.5, y=1.6, w=12.33,
    headers=["項目", "Q1", "Q2"],
    rows=[["売上", "3.0億", "3.2億"]],
    col_widths=[4.0, 4.0, 4.33],
)
```

### `add_comparison(slide, y, left_title, right_title, left_items, right_items, ...)`

### `add_bar_chart_img(slide, x, y, w, h, data, title="")`

```python
add_bar_chart_img(slide, 0.5, 1.6, 8.0, 3.0, {
    "labels": ["A社", "B社", "当社"],
    "values": [80, 120, 145],
    "colors": ["D9D9D9", "D9D9D9", "0070C0"],
})
```

---

## ヘルパー関数 — フロー・構造

### `add_flow_row(slide, y, steps, color="accent", box_h=0.9)`

### `add_timeline(slide, y, events, color="accent")`

```python
add_timeline(slide, y=2.0, events=[
    {"date": "2026年4月", "title": "開始", "desc": "詳細説明"},
])
```

### `add_agenda(slide, items, current=None, x_start=2.0, y_start=1.8, item_h=0.65, w=9.0)`

### `add_slide_header(slide, title, section="", color="accent")`

### `add_slide_footer(slide, page_num=None, footer_text="")`

---

## ヘッダー・フッター配置ルール

- **ヘッダー**: `y = 0` 〜 `CONTENT_TOP (1.4")` — `add_slide_header()` が担当
- **コンテンツ**: `y >= CONTENT_TOP + 0.1` 以降に配置
- **フッター**: `y = SLIDE_H - 0.5` 付近 — `add_slide_footer()` が担当
- コンテンツ開始推奨座標: `y = CONTENT_TOP + 0.2 = 1.6`

---

## 禁止事項

| 禁止 | 理由 | 代替 |
|---|---|---|
| pptx native table | スタイル制御不能 | `add_table_shapes()` を使う |
| 游ゴシック系フォント | ブランド制約 | `Meiryo UI`（自動設定） |
| `add_rounded_rect()` | 角丸はデザイン方針として使用禁止 | `add_rect()` を使う |
| `shape.text_frame.text = "..."` | スタイル失われる | `add_text()` を使う |
| コンテンツ x > 12.9" | はみ出し | x + w ≤ 12.9 に収める |
| コンテンツ y > 7.0" | はみ出し | y + h ≤ 7.0 に収める |

---

## よくある間違い

### ❌ コンテンツがヘッダーと重なる
```python
add_text(slide, 0.5, 0.5, ...)  # NG: CONTENT_TOP より上
add_text(slide, 0.5, 1.6, ...)  # OK
```

### ❌ 右端はみ出し
```python
add_text(slide, 7.0, 1.6, 6.5, ...)  # NG: 7.0 + 6.5 = 13.5 > SAFE_XMAX
add_text(slide, 7.0, 1.6, 5.8, ...)  # OK: 7.0 + 5.8 = 12.8 < SAFE_XMAX
```

### ❌ 空の `add_body()` slide に直接 save
```python
slide = sb.add_body("タイトル")
# ← ここで add_text() などでコンテンツを追加する
sb.save_and_validate("output/xxx.pptx")
```

---

## バリデーション

```python
from lib import validate
result = validate("output/xxx.pptx")
print(result.ok)       # True/False
print(result.errors)   # 修正必須
print(result.warnings) # 確認推奨（警告のみの場合はOK）
```

### エラー（修正必須）
- 游ゴシック系フォント使用
- コンテンツはみ出し（right > 12.9" or bottom > 7.0"）
- pptx native table 使用

### 警告（確認推奨）
- フォントサイズ < 9pt or > 60pt
- 1スライド300文字超
- 1スライド25シェイプ超（複数カラムでは正常）
