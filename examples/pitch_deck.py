"""
ピッチデッキサンプル — スタートアップ向け投資家提案

使用パターン:
  cover_basic → agenda → quote → body_3col → data_kpi → data_comparison
  → flow_process → flow_timeline → closing

使用方法:
    /home/moog/hr_venv/bin/python examples/pitch_deck.py
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from lib import (
    SlideBuilder,
    add_text, add_rect, add_rounded_rect, add_kpi_row,
    add_table_shapes, add_timeline, add_comparison, add_flow_row,
    add_agenda, add_line, add_badge, add_quote, add_pill, add_card,
    C, CONTENT_TOP, SLIDE_W, SLIDE_H,
)

OUTPUT = "output/pitch_deck.pptx"


def main():
    sb = SlideBuilder(theme={"accent": "5C2D91"})  # パープルテーマ
    sb.set_footer("Confidential — 投資家向け資料")

    # --- 1. 表紙 ---
    sb.add_cover(
        title="〇〇テクノロジー株式会社\nシリーズA 調達説明資料",
        subtitle="AIを活用した次世代HRプラットフォーム",
        date="2026年4月",
    )

    # --- 2. エグゼクティブサマリー (引用スタイル) ---
    slide = sb.add_body("エグゼクティブサマリー")
    add_quote(
        slide, x=0.8, y=CONTENT_TOP + 0.2, w=11.73,
        text="「人材が最大の競争優位」と言われて久しいが、大半の企業は依然として\nExcelと勘で採用・育成を管理している。私たちはAIでこれを変える。",
        source="CEO 〇〇 〇〇",
    )
    add_kpi_row(slide, y=CONTENT_TOP + 2.4, items=[
        {"label": "調達目標", "value": "5", "unit": "億円", "color": "accent"},
        {"label": "ARR（現在）", "value": "1.2", "unit": "億円", "color": "accent2"},
        {"label": "顧客数", "value": "120", "unit": "社", "delta": "+40 QoQ", "color": "accent"},
        {"label": "NRR", "value": "142", "unit": "%", "color": "accent2"},
    ], x_start=0.8, total_w=11.73)

    # --- 3. 課題 ---
    slide = sb.add_body("課題：HRの意思決定はいまだ非科学的", section="Problem")
    add_comparison(
        slide,
        y=CONTENT_TOP + 0.2,
        left_title="現在の痛み",
        right_title="その結果",
        left_items=[
            "採用根拠がデータではなく「感覚」",
            "育成投資と離職の相関が不明",
            "HRシステムが乱立・連携なし",
            "人事担当1人あたり150人管理",
        ],
        right_items=[
            "採用ミスコストは年収の30〜50%",
            "スキルギャップが埋まらない",
            "戦略人事への変革が滞る",
            "エンゲージメント低下→退職率増",
        ],
        left_color="accent3",
        right_color="accent",
    )

    # --- 4. ソリューション（3カラム） ---
    slide = sb.add_body("ソリューション：HR Intelligenceプラットフォーム", section="Solution")
    # 3つの価値柱を手動構築
    cols = [
        {"title": "Talent Analytics", "desc": "採用・評価データをAI分析\n最適人材を科学的に特定", "color": "accent"},
        {"title": "Skill Graph", "desc": "社内スキルマップを自動構築\n成長パスを可視化", "color": "accent2"},
        {"title": "Predict & Act", "desc": "離職リスクを事前検知\n介入施策を自動提案", "color": "5C2D91"},
    ]
    gap = 0.2
    col_w = (12.33 - gap * 2) / 3
    y_base = CONTENT_TOP + 0.2
    for i, col in enumerate(cols):
        cx = 0.5 + i * (col_w + gap)
        add_rounded_rect(slide, cx, y_base, col_w, 5.3, fill="bgLight", border="border")
        add_rect(slide, cx, y_base, col_w, 0.6, fill=col["color"])
        add_text(slide, cx + 0.1, y_base + 0.1, col_w - 0.2, 0.42,
                 col["title"], style="subheading", align="center", color="white", font_size=15)
        add_text(slide, cx + 0.15, y_base + 0.85, col_w - 0.3, 1.6,
                 col["desc"], style="body", align="center", color="text")
        add_badge(slide, cx + col_w / 2 - 0.225, y_base + 0.7, i + 1, color=col["color"])

    # --- 5. 市場規模 ---
    slide = sb.add_body("市場規模：TAM $220B の HRTech 市場", section="Market")
    add_kpi_row(slide, y=CONTENT_TOP + 0.3, items=[
        {"label": "グローバル HRTech TAM", "value": "$220", "unit": "B", "color": "accent"},
        {"label": "日本 HRTech SAM", "value": "¥3,200", "unit": "億円", "color": "accent2"},
        {"label": "ターゲット SOM（3年）", "value": "¥150", "unit": "億円", "color": "accent"},
        {"label": "市場成長率", "value": "18", "unit": "% CAGR", "delta": "加速中", "color": "accent2"},
    ])

    # --- 6. ビジネスモデル ---
    slide = sb.add_body("ビジネスモデル：サブスクリプション×プロサービス", section="Business Model")
    add_table_shapes(
        slide,
        x=0.5, y=CONTENT_TOP + 0.2, w=12.33,
        headers=["収益区分", "単価", "モデル", "粗利率", "成長見込"],
        rows=[
            ["Platform SaaS", "月額30-150万円/社", "年間サブスク", "85%+", "主力 ↑↑"],
            ["Data Insights", "月額10-50万円/社", "従量+サブスク", "75%+", "急成長 ↑↑↑"],
            ["Consulting", "500-2,000万円/案件", "スポット", "40-50%", "安定 →"],
        ],
        col_widths=[2.3, 2.5, 2.3, 1.7, 1.63],
    )
    add_text(slide, 0.5, CONTENT_TOP + 2.4, 12.33, 0.4,
             "Platform SaaSをコアに、Data Insightsでの拡販による LTV 最大化を狙う。",
             style="body", color="accent")

    # --- 7. トラクション ---
    slide = sb.add_body("トラクション：18ヶ月で ARR 1.2億円", section="Traction")
    add_timeline(slide, y=CONTENT_TOP + 1.0, events=[
        {"date": "2024年10月", "title": "創業・β版", "desc": "5社でパイロット開始"},
        {"date": "2025年4月", "title": "正式リリース", "desc": "ARR 2,000万円達成"},
        {"date": "2025年10月", "title": "シード調達", "desc": "1.5億円 / 顧客50社"},
        {"date": "2026年3月", "title": "現在", "desc": "ARR 1.2億円\n顧客120社"},
        {"date": "2026年後半", "title": "シリーズA活用", "desc": "海外展開\nAI機能拡充"},
    ])

    # --- 8. 資金使途 ---
    slide = sb.add_body("資金使途：5億円の活用計画", section="Use of Funds")
    add_flow_row(slide, y=CONTENT_TOP + 0.8, steps=[
        "プロダクト\n開発強化",
        "セールス\nマーケ",
        "カスタマー\nサクセス",
        "グローバル\n展開",
    ], color="accent", box_h=1.1)
    add_table_shapes(
        slide,
        x=1.5, y=CONTENT_TOP + 2.4, w=10.33,
        headers=["用途", "金額", "比率", "期待効果"],
        rows=[
            ["プロダクト開発（エンジニア採用）", "2.5億円", "50%", "AI機能 3倍、開発速度向上"],
            ["セールス・マーケティング", "1.5億円", "30%", "顧客300社・ARR 3億円"],
            ["カスタマーサクセス体制整備", "0.5億円", "10%", "NRR 150%+維持"],
            ["ASEAN展開（SG・TH拠点）", "0.5億円", "10%", "海外ARR 5,000万円"],
        ],
        col_widths=[3.5, 1.7, 1.2, 3.13],
    )

    # --- 9. クロージング ---
    slide = sb._new_blank_slide()
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill="accent")
    add_rect(slide, 0, SLIDE_H - 0.25, SLIDE_W, 0.25, fill="white")
    add_text(slide, 1.0, 1.8, SLIDE_W - 2.0, 1.2,
             "人事を変え、会社を変える", style="title_cover", align="center", font_size=36)
    add_text(slide, 1.0, 3.2, SLIDE_W - 2.0, 0.6,
             "〇〇テクノロジーと共に、データドリブンHRの未来を創りましょう",
             style="subtitle_cover", align="center")
    add_line(slide, 3.0, 4.1, SLIDE_W - 3.0, 4.1, color="white", width=1.5)
    add_text(slide, 1.0, 4.4, SLIDE_W - 2.0, 0.35,
             "〇〇 〇〇  CEO｜oo@example.co.jp｜090-XXXX-XXXX",
             style="caption", align="center", color="white")

    result = sb.save_and_validate(OUTPUT)
    print(f"\n生成完了: {OUTPUT}")
    print(f"スライド数: {len(sb.prs.slides)}")
    if result.ok:
        print("バリデーション: OK")
    else:
        print(f"バリデーション: {len(result.errors)} エラー / {len(result.warnings)} 警告")


if __name__ == "__main__":
    main()
