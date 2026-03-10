"""
ビジネスケース — 投資対効果・意思決定支援（5枚）

構成: cover → consulting_scr → data_kpi（現状）→ data_comparison → flow_timeline → closing
テーマ: default
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import (
    SlideBuilder, add_text, add_rect, add_rounded_rect, add_kpi_row,
    add_comparison, add_timeline, add_table_shapes, add_badge, add_line,
    C, CONTENT_TOP, SLIDE_W, SLIDE_H,
)

OUTPUT = "output/business_case.pptx"


def main():
    sb = SlideBuilder(theme="default")
    sb.set_footer("ビジネスケース — 経営企画部 Confidential")

    # --- 1. 表紙 ---
    sb.add_cover(
        title="DX推進投資\nビジネスケース",
        subtitle="株式会社サンプル 経営企画部",
        date="2026年4月",
    )

    # --- 2. エグゼクティブサマリー（SCR） ---
    slide = sb.add_body("エグゼクティブサマリー")
    labels = [
        ("S", "Situation",    "現状",   "データ分断・デジタル人材不足により、意思決定速度が競合比1.8倍遅い。現行ペースでは2028年に市場シェア3pt喪失リスク。", "accent2"),
        ("C", "Complication", "課題",   "現状のシステム維持費は年間8億円超。DX投資を先送りするほど、競合との差が拡大し、将来の改善コストも増大する。", "accent3"),
        ("R", "Resolution",   "推奨策", "総投資額42億円のDX変革プログラムを実施。3年でROI 180%・投資回収18ヶ月を見込む。今期承認を推奨する。", "accent"),
    ]
    bh, gap = 1.55, 0.18
    for i, (letter, eng, jpn, body, color) in enumerate(labels):
        by = CONTENT_TOP + 0.2 + i * (bh + gap)
        add_rounded_rect(slide, 0.5, by, 12.33, bh, fill="bgLight", border="border")
        add_rect(slide, 0.5, by, 0.07, bh, fill=color)
        add_badge(slide, 0.68, by + 0.55, letter, color=color)
        add_text(slide, 1.3, by + 0.12, 2.2, 0.38, eng, style="subheading", color=color)
        add_text(slide, 1.3, by + 0.52, 2.2, 0.3, jpn, style="caption", color="textMuted")
        add_line(slide, 3.6, by + 0.18, 3.6, by + bh - 0.18, color="border", width=0.75)
        add_text(slide, 3.8, by + 0.22, 8.8, bh - 0.44, body, style="body", word_wrap=True)

    # --- 3. 財務インパクト試算 ---
    slide = sb.add_body("財務インパクト試算", section="投資対効果")
    add_kpi_row(slide, y=CONTENT_TOP + 0.3, items=[
        {"label": "総投資額",       "value": "42",   "unit": "億円",   "color": "accent3"},
        {"label": "3年累計リターン", "value": "75.6", "unit": "億円",   "color": "accent2"},
        {"label": "ROI（3年）",     "value": "180",  "unit": "%",      "color": "accent2"},
        {"label": "投資回収期間",   "value": "18",   "unit": "ヶ月",   "color": "accent"},
    ])
    headers = ["施策", "投資額", "Year1効果", "Year2効果", "Year3効果", "累計効果"]
    rows = [
        ["データ統合基盤",    "¥18億", "¥3億",  "¥12億", "¥18億", "¥33億"],
        ["デジタル人材育成",  "¥12億", "¥1億",  "¥5億",  "¥10億", "¥16億"],
        ["ガバナンス整備",    "¥6億",  "¥0.5億", "¥3億",  "¥8億",  "¥11.5億"],
        ["その他（予備費）",  "¥6億",  "—",      "—",     "¥15億", "¥15億"],
        ["合計",              "¥42億", "¥4.5億", "¥20億", "¥51億", "¥75.5億"],
    ]
    col_widths = [3.0, 1.5, 1.7, 1.7, 1.7, 1.73]
    add_table_shapes(slide, x=0.5, y=CONTENT_TOP + 2.1, w=11.4,
                     headers=headers, rows=rows, col_widths=col_widths)

    # --- 4. Do Nothing シナリオ比較 ---
    slide = sb.add_body("投資 vs. 現状維持 比較", section="意思決定")
    add_comparison(
        slide, y=CONTENT_TOP + 0.2,
        left_title="DX推進（推奨シナリオ）",
        right_title="現状維持（Do Nothing）",
        left_items=[
            "ROI 180%（3年） / 回収18ヶ月",
            "意思決定速度 2倍向上",
            "デジタル人材比率 3% → 12%",
            "システム維持費 ▲30%（標準化効果）",
            "2028年市場シェア現状維持〜拡大",
        ],
        right_items=[
            "追加投資ゼロ（見かけ上）",
            "意思決定ラグ継続・拡大",
            "人材採用コスト増（市場競争激化）",
            "老朽システム維持費 年8億円継続",
            "2028年市場シェア ▲3〜5pt リスク",
        ],
        left_color="accent2",
        right_color="accent3",
    )

    # --- 5. ロードマップ ---
    slide = sb.add_body("実現ロードマップ", section="実行計画")
    add_timeline(slide, y=CONTENT_TOP + 1.0, events=[
        {"date": "2026 Q2",  "title": "承認・着手",    "desc": "予算確保\n体制構築"},
        {"date": "2026 Q3",  "title": "基盤構築開始",  "desc": "データ統合\nPJ着手"},
        {"date": "2026 Q4",  "title": "パイロット",    "desc": "2部門で\n先行実施"},
        {"date": "2027 Q1",  "title": "全社展開",      "desc": "ロールアウト\n人材育成"},
        {"date": "2027 Q3",  "title": "効果測定",      "desc": "ROI確認\n追加施策"},
    ])
    add_kpi_row(slide, y=CONTENT_TOP + 3.3, items=[
        {"label": "承認判断期限",   "value": "2026/5", "unit": "末日",  "color": "accent3"},
        {"label": "全社展開完了",   "value": "2027",   "unit": "Q1目標", "color": "accent"},
        {"label": "効果確認時期",   "value": "2027",   "unit": "Q3",    "color": "accent2"},
    ])

    # --- 6. クロージング ---
    slide = sb._new_blank_slide()
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill="accent")
    add_rect(slide, 0, SLIDE_H - 0.25, SLIDE_W, 0.25, fill="white")
    add_text(slide, 1.0, 1.8, SLIDE_W - 2.0, 1.2,
             "2026年5月末までの承認を推奨します", style="title_cover", align="center", font_size=26)
    add_line(slide, 3.0, 3.3, SLIDE_W - 3.0, 3.3, color="white", width=1.0)
    add_text(slide, 1.5, 3.6, SLIDE_W - 3.0, 0.5,
             "投資の先送りは機会損失と回収期間の長期化をもたらします。\n"
             "本ビジネスケースの詳細・補足データは担当PMまでお問い合わせください。",
             style="body", color="white", align="center")

    result = sb.save_and_validate(OUTPUT)
    print(f"\n生成完了: {OUTPUT}  ({len(sb.prs.slides)}枚)")


if __name__ == "__main__":
    main()
