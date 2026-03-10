# /slides — 汎用PPTXスライド生成スキル

## 概要

自然言語の指示から、ブランド制約付きPPTXプレゼンテーションを生成するスキル。

**トリガー**: `/ppt-gen` または「〇〇のスライドを作って」「PPTを生成して」などの意図で起動

---

## 環境

```
プロジェクト: /home/moog/work/ppt-gen
Python: /home/moog/.local/bin/uv run python（作業ディレクトリは /home/moog/work/ppt-gen）
出力先: /home/moog/work/ppt-gen/output/
Windowsからのパス: \\wsl$\Ubuntu\home\moog\work\ppt-gen\output\
```

---

## ワークフロー（6ステップ）

### Step 1: 要件確認

ユーザーの指示から以下を把握する:
- **目的**: 何のためのプレゼンか（事業報告 / 提案 / 定例報告 / ピッチ等）
- **スライド数**: 何枚か（未指定の場合は目的から推定）
- **テーマカラー**: 指定があれば反映（未指定は `0070C0` デフォルト）
- **内容**: 各スライドに何を載せるか

### Step 2: ドキュメント種別・パターン・テーマ選定

`catalog/index.html` を参照して適切なパターンを選択する。

**コンサルティングドキュメント種別と推奨構成:**

| ドキュメント種別 | テーマ | 推奨スライド構成 | 参照例 |
|---|---|---|---|
| ドアノッカー | mckinsey | cover → consulting_scr → consulting_findings → data_comparison → closing | `examples/door_knocker.py` |
| 提案書 | default/accenture | cover → consulting_scr → consulting_scope → consulting_pricing → flow_timeline → closing | `examples/proposal.py` |
| SOW | navy | cover → body_1col → consulting_scope → consulting_pricing → flow_timeline → closing | `examples/sow.py` |
| ビジネスケース | default | cover → consulting_scr → data_kpi → data_comparison → flow_timeline → closing | `examples/business_case.py` |
| 定例報告 | default | cover → data_kpi → data_table → consulting_risk → flow_timeline → closing | `examples/status_report.py` |
| ステアリングコミッティ | navy | cover → data_kpi → consulting_risk → flow_timeline → closing | `examples/steering_committee.py` |
| ワークショップ | green | cover → agenda → cover_section → workshop_exercise → closing | `examples/workshop.py` |
| フィンディングス | default | cover → consulting_scr → consulting_findings × 2 → data_comparison → closing | `examples/findings.py` |
| 戦略提言 | mckinsey | cover → consulting_scr → data_comparison → body_3col → flow_timeline → closing | `examples/strategic_recommendation.py` |
| エグゼクティブブリーフィング | navy | cover → data_kpi → consulting_risk → closing | `examples/executive_briefing.py` |

**パターン指定例（ユーザーがこう言ったら従う）:**
- 「KPIスライドは data_kpi パターンで」→ `add_kpi_row` を使う
- 「比較スライドは body_2col で」→ 2カラムレイアウト
- 「フローは flow_process で」→ `add_flow_row` を使う
- 「フィンディングス形式で」→ `consulting_findings` パターン（番号付きカード＋重大度バッジ）
- 「SCR形式で」→ `consulting_scr` パターン（状況・課題・提言の3ブロック）

**テーマプリセット（`lib.THEMES` に定義済み）:**

| テーマ名 | カラー | 用途 |
|---|---|---|
| `default` | 企業ブルー #0070C0 | 汎用・デフォルト |
| `accenture` | パープル #A100FF | コンサルティング・提案書 |
| `navy` | 濃紺 #1A3C6E | 重厚感・金融・官公庁 |
| `green` | グリーン #00875A | ESG・サステナビリティ |
| `warm` | オレンジ #C55A11 | エネルギー・製造 |
| `mckinsey` | 濃紺 #002F6C | 戦略コンサル・経営報告 |

指定方法: `SlideBuilder(theme="mckinsey")`
カタログ確認: `catalog/index.html`（テーマ一覧＋パターン一覧）

### Step 3: コード作成

`CLAUDE.md` を参照してコードを作成する。

```python
# 雛形
import sys
sys.path.insert(0, "/home/moog/work/ppt-gen")
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
cd /home/moog/work/ppt-gen
/home/moog/.local/bin/uv run python output_script_or_inline.py
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

### 「ステアリングコミッティ資料をnavyテーマで」

→ `examples/steering_committee.py` を参照。`SlideBuilder(theme="navy")` を使用

### 「フィンディングスレポートを作って」

→ `examples/findings.py` を参照。`consulting_scr` + `consulting_findings` × 2 の構成

---

## 階層型比較マトリクス表（特別対応）

「比較表」「比較マトリクス」「コンサル表」などのキーワードがある場合、このセクションに従う。

### 表の構造

```
┌──────────┬────────────┬──────────────────┬──────────────────┐
│ グループ  │ サブグループ│   Column A       │   Column B ★    │
│（スパン） │            │   (箇条書き)      │   (箇条書き)     │
├──────────┼────────────┼──────────────────┼──────────────────┤
│          │ サブ1      │ • ...            │ • ...            │
│  画面    ├────────────┼──────────────────┼──────────────────┤
│          │ サブ2      │ • ...            │ • ...            │
├──────────┼────────────┼──────────────────┼──────────────────┤
│  機能    │ サブ3      │ • ...            │ • ...            │
│  (など)  │ サブ4      │ • ...            │ • ...            │
└──────────┴────────────┴──────────────────┴──────────────────┘
```

- グループ列: 同じ名前が連続する行を**縦スパン結合**。`None`なら結合なし（サブグループ列と横結合）
- Column Bはアクセントカラーでハイライト

### Step 1: AskUserQuestion で情報収集

```
Q1: スライドタイトル・Column A/B の列見出し・ハイライトする列（A or B）
Q2: テーマカラー
Q3: 表データを以下の形式で入力してください（1行1行）:
      グループ名（なしは空欄）| サブグループ名 | ColAの箇条書き（改行区切り） | ColBの箇条書き（改行区切り）
    例:
      （空欄）    | 構築プロセス | スクラッチ開発... | 設定シート入力で...
      画面        | 画面項目     | テキスト入力...\nカレンダー... | 設定シートで追加...
      画面        | 画面レイアウト | 柔軟なレイアウト... | 1列/2列選択...
      機能        | エラー制御   | 必須・桁数チェック... | 単項目チェックのみ...
```

### Step 2: コード生成・実行

参照テンプレート: `output/gen_comparison_table.py`（実績あり）

**レイアウト定数:**
```
TX=0.35 / TW=12.45 / GW=0.82 / SW=1.72 / CW=(TW-GW-SW)/2 ≈ 4.955"
HDR_H=0.38 / LINE_H=0.195 / PAD_H=0.22 / MIN_H=0.50
```

**行高さの動的計算（バレット数に応じる）:**
```python
def row_h(ba, bb):
    return max(0.50, max(len(ba), len(bb), 1) * 0.195 + 0.22)
```

**グループスパン計算:**
```python
def calc_groups(rows):
    result, i = [], 0
    while i < len(rows):
        g = rows[i][0]
        if g is None:
            result.append((None, i, i)); i += 1
        else:
            j = i
            while j < len(rows) and rows[j][0] == g: j += 1
            result.append((g, i, j-1)); i = j
    return result
```

**グループなし行（スパンなし）はサブグループ列を左に拡張:**
```python
sub_x = TX + (GW if grp_name else 0)
sub_w = SW + (0  if grp_name else GW)
```

**テーマ別パレット（6色定義）:**

| キー | accenture | mckinsey | default（緑系） |
|---|---|---|---|
| `grp_bg` | 5B0099 | 002F6C | 1E6B3C |
| `sub_bg` | A100FF | 1455A0 | 2E8B57 |
| `hdr_b_bg` | EDD9FF | D6E4F0 | D4EDDA |
| `row_b_even` | FAF5FF | F0F5FA | F0FBF4 |

**コンテンツセルの箇条書き:**
```python
def bullets(items): return "\n".join(f"• {s}" for s in items)
add_text(slide, cx+0.10, ry+0.06, CW-0.16, rh-0.10, bullets(ba),
         style="caption", align="left", color="1A1A1A", font_size=10)
```

**外枠（最後に描画）:**
```python
total_h = sum(row_hs) + HDR_H
add_rect(slide, TX, TY, TW, total_h, fill=None, border="888888", border_width=1.5)
```

---

## ガントチャート生成（特別対応）

「ガントチャート」「工程表」「スケジュール」などのキーワードがある場合、通常フローではなく以下に従う。

### Step 1: AskUserQuestion で情報収集

**1回目（選択式）**: テーマ・粒度を並べて質問する。

```
Q1: テーマ（accenture / mckinsey / navy / default）
Q2: 粒度（週次=各月W1〜W4 / 月次）
```

**2回目（自由記述）**: プロジェクト情報とタスクをまとめて質問する。

```
Q3: プロジェクト名・開始年月（例: 2026-04）・期間（ヶ月数）
Q4: タスク一覧を以下の形式で（1行1タスク）
      タスク名, P1〜P5, 開始[週/月]番号, 終了[週/月]番号
    マイルストーン（任意）:
      マイルストーン名, [週/月]番号
```

ユーザーが途中で指定している情報は再度聞かない。

### Step 2: コード生成・実行

参照テンプレート: `output/gen_gantt.py`（実績あり）

**テーマ別パレット（コード中に埋め込む）:**

```python
THEME_PALETTE = {
    "accenture": {"phases":["5B0099","A100FF","BE4DFF","D98AFF","ECC6FF"],
                  "hdr_bg":"5B0099","hdr_text":"FFFFFF",
                  "cell_a":"EDD9FF","cell_b":"F5EEFF","cell_border":"D4A8FF",
                  "row_alt":"FAF5FF","stripe_sep":"B366FF",
                  "ms_line":"A100FF","ms_badge":"A100FF","ms_text":"5B0099"},
    "mckinsey":  {"phases":["002F6C","1455A0","2980B9","5DADE2","A9CCE3"],
                  "hdr_bg":"002F6C","hdr_text":"FFFFFF",
                  "cell_a":"D6E4F0","cell_b":"EBF5FB","cell_border":"C0D3E8",
                  "row_alt":"F2F5FA","stripe_sep":"1455A0",
                  "ms_line":"002F6C","ms_badge":"002F6C","ms_text":"002F6C"},
    "navy":      {"phases":["1A3C6E","2B5DA7","4C80C4","7CAAD9","AECCE8"],
                  "hdr_bg":"1A3C6E","hdr_text":"FFFFFF",
                  "cell_a":"D4E2F0","cell_b":"EBF2F8","cell_border":"B5CCE0",
                  "row_alt":"F0F4F8","stripe_sep":"2B5DA7",
                  "ms_line":"1A3C6E","ms_badge":"1A3C6E","ms_text":"1A3C6E"},
    "default":   {"phases":["004B8D","0070C0","3498DB","7EC8F5","C5E5F7"],
                  "hdr_bg":"0070C0","hdr_text":"FFFFFF",
                  "cell_a":"D4EBFA","cell_b":"EBF5FC","cell_border":"B3D7F2",
                  "row_alt":"F0F7FC","stripe_sep":"0070C0",
                  "ms_line":"0070C0","ms_badge":"0070C0","ms_text":"0070C0"},
}
```

**レイアウト定数（共通）:**

```
LABEL_X=0.35 / LABEL_W=2.65 / CHART_X=3.05 / CHART_W=9.75（右端12.80"）
BADGE_D=0.28 / SAFE_R=12.9
```

**週次モード（2段ヘッダー）:**
```
MO_HDR_H=0.30 / WK_HDR_H=0.26 / TASKS_Y=2.20
N_WEEKS = n_months × 4  /  WW = CHART_W / N_WEEKS
```

**月次モード（1段ヘッダー）:**
```
HDR_H=0.36 / TASKS_Y=2.00
MW = CHART_W / n_months
```

**行高さの動的計算（タスク数に応じて自動調整）:**
```python
MS_H  = (0.06 + BADGE_D + 0.04 + 0.22) if milestones else 0.08
LEG_H = 0.08 + 0.20
ROW_H = max(0.35, min(0.55, (7.0 - TASKS_Y - MS_H - LEG_H) / len(tasks)))
```

**バーテキスト色の自動判定:**
```python
def bar_text_color(hex):
    r,g,b = int(hex[0:2],16),int(hex[2:4],16),int(hex[4:6],16)
    return "1A1A1A" if 0.299*r+0.587*g+0.114*b > 150 else "FFFFFF"
```

**出力ファイル名:**
```python
from datetime import datetime
fname = f"output/gantt_{proj_title}_{datetime.now().strftime('%m%d_%H%M')}.pptx"
```

### 注意事項
- マイルストーン縦線はバーより先（背面）に描画する
- バッジ右端クランプ: `bx = min(mx - BADGE_D/2, 12.9 - BADGE_D)`
- マイルストークラベル右端クランプ: `lb_x = max(0.35, min(mx - lb_w/2, 12.9 - lb_w))`
- 月境界の縦線は太め（width=1.0）、週境界は細め（width=0.5）

---

## トラブルシューティング

| エラー | 原因 | 対処 |
|---|---|---|
| `right_edge > 12.9"` | テキストボックス幅が大きすぎ | `w` を小さくして `x + w ≤ 12.8` に |
| `bottom > 7.0"` | 下端はみ出し | `h` を減らすか `y` を上げる |
| `native table` | `prs.slides.add_table()` 使用 | `add_table_shapes()` に変換 |
| `Permission denied` | Windows で PPTX が開かれている | ファイルを閉じて再実行 |
