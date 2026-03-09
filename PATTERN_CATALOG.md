# パターンカタログ

汎用PPTXスライド生成ツール — 全13パターン一覧。

## カテゴリ別一覧

### 表紙（Cover）

| パターン | ファイル | 説明 | 主な用途 |
|---|---|---|---|
| `cover_basic` | `patterns/cover_basic.py` | シンプルなカラー背景表紙 | プレゼン開始、報告書表紙 |
| `cover_section` | `patterns/cover_section.py` | セクション区切り（番号付き左帯） | 章・セクション開始 |
| `closing` | `patterns/closing.py` | 最終スライド（Thank you + 次のステップ） | プレゼン締め、連絡先 |

### 本文（Body）

| パターン | ファイル | 説明 | 主な用途 |
|---|---|---|---|
| `body_1col` | `patterns/body_1col.py` | 1カラム箇条書き | 説明スライド、方針発表 |
| `body_2col` | `patterns/body_2col.py` | 2カラム並列 | 比較説明、Before/After |
| `body_3col` | `patterns/body_3col.py` | 3カラム並列 | 3施策・3特徴の並列説明 |

### データ（Data）

| パターン | ファイル | 説明 | 主な用途 |
|---|---|---|---|
| `data_kpi` | `patterns/data_kpi.py` | KPI数値カード横並び | 業績ハイライト、KPIダッシュボード |
| `data_table` | `patterns/data_table.py` | 図形ベーステーブル | 一覧表、データ比較 |
| `data_comparison` | `patterns/data_comparison.py` | 左右2列比較 | 課題と施策、現状と目標 |

### フロー（Flow）

| パターン | ファイル | 説明 | 主な用途 |
|---|---|---|---|
| `flow_process` | `patterns/flow_process.py` | 横並びプロセスフロー | 業務フロー、ステップ解説 |
| `flow_timeline` | `patterns/flow_timeline.py` | 横タイムライン | ロードマップ、スケジュール |

### 構造・強調（Structure / Emphasis）

| パターン | ファイル | 説明 | 主な用途 |
|---|---|---|---|
| `agenda` | `patterns/agenda.py` | アジェンダ（目次） | プレゼン冒頭目次 |
| `quote` | `patterns/quote.py` | 引用・インサイト強調 | キーメッセージ、インタビュー引用 |

---

## パターンの使い方

### 単体実行

```bash
cd /mnt/c/Users/moogs/work/ppt-gen
/home/moog/hr_venv/bin/python patterns/cover_basic.py
# → output/cover_basic.pptx
```

### プログラムから呼び出す

```python
import sys
sys.path.insert(0, "/mnt/c/Users/moogs/work/ppt-gen")
from patterns.cover_basic import run as cover_basic

cover_basic(
    title="Q1事業報告",
    subtitle="XX株式会社",
    date="2026年4月",
    output_path="output/my_cover.pptx",
)
```

### SlideBuilder で複数パターンを組み合わせる（推奨）

```python
from lib import SlideBuilder, add_text, add_kpi_row, CONTENT_TOP

sb = SlideBuilder()
sb.add_cover("プレゼンタイトル", "組織名", "2026年4月")

slide = sb.add_body("業績ハイライト", section="サマリー")
add_kpi_row(slide, y=CONTENT_TOP + 0.3, items=[
    {"label": "売上高", "value": "12.4", "unit": "億円", "delta": "+8%"},
])

sb.save_and_validate("output/my_pptx.pptx")
```

---

## よくある組み合わせパターン

### 事業報告（5〜8スライド）

```
cover_basic → agenda → data_kpi → data_table → body_2col → flow_timeline → closing
```

### ピッチデッキ（8〜12スライド）

```
cover_basic → quote → body_3col → data_kpi → data_comparison → flow_process → flow_timeline → closing
```

### 社内提案（3〜6スライド）

```
cover_basic → body_2col → data_kpi → flow_process → closing
```

### 定例報告（4〜6スライド）

```
cover_basic → data_kpi → data_table → flow_timeline → closing
```

---

## パターン追加方法

`patterns/new_pattern.py` を作成して以下の docstring と `run()` 関数を実装する:

```python
"""
パターン名: new_pattern
カテゴリ: 本文
説明: 新しいパターンの説明
用途: どんな場面で使うか
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, ...

def run(output_path: str = "output/new_pattern.pptx", **kwargs):
    sb = SlideBuilder()
    slide = sb.add_body("タイトル")
    # ... コンテンツ追加 ...
    sb.save_and_validate(output_path)

if __name__ == "__main__":
    run()
```

カタログは `tools/gen_catalog.py` を再実行すると自動更新される。
