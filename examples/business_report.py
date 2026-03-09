"""
事業報告サンプル — 四半期業績報告プレゼンテーション

使用パターン:
  cover_basic → agenda → data_kpi → body_2col → data_table → flow_timeline → closing

使用方法:
    /home/moog/hr_venv/bin/python examples/business_report.py
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from lib import (
    SlideBuilder,
    add_text, add_rect, add_rounded_rect, add_kpi_row,
    add_table_shapes, add_timeline, add_comparison, add_flow_row,
    add_agenda, add_line, add_badge, add_quote,
    C, CONTENT_TOP, SLIDE_W, SLIDE_H,
)

OUTPUT = "output/business_report.pptx"


def main():
    sb = SlideBuilder(theme={"accent": "0070C0"})
    sb.set_footer("2026年Q1 事業報告 — Confidential")

    # --- 1. 表紙 ---
    sb.add_cover(
        title="2026年 第1四半期\n事業報告",
        subtitle="XX株式会社 / 経営企画部",
        date="2026年4月25日",
    )

    # --- 2. アジェンダ ---
    slide = sb.add_body("アジェンダ")
    add_agenda(slide, [
        "Q1 業績ハイライト",
        "事業別売上分析",
        "コスト・投資状況",
        "Q2 見通しと施策",
        "まとめ・次のステップ",
    ], x_start=2.5, y_start=CONTENT_TOP + 0.3, w=8.5)

    # --- 3. KPI ハイライト ---
    slide = sb.add_body("Q1 業績ハイライト", section="業績サマリー")
    add_kpi_row(slide, y=CONTENT_TOP + 0.3, items=[
        {"label": "売上高", "value": "12.4", "unit": "億円", "delta": "+8.2% YoY", "color": "accent"},
        {"label": "営業利益", "value": "2.1", "unit": "億円", "delta": "+12.5% YoY", "color": "accent2"},
        {"label": "営業利益率", "value": "16.9", "unit": "%", "delta": "+0.6pt", "color": "accent2"},
        {"label": "顧客数", "value": "4,820", "unit": "社", "delta": "+8.3% YoY", "color": "accent"},
    ])
    add_text(slide, 0.5, CONTENT_TOP + 2.2, 12.0, 0.4,
             "売上・利益ともに計画を上回る進捗。特に新規顧客獲得が好調。",
             style="body", color="accent")

    # --- 4. 事業別売上分析 ---
    slide = sb.add_body("事業別売上分析", section="業績詳細")
    add_comparison(
        slide,
        y=CONTENT_TOP + 0.2,
        left_title="前四半期比（QoQ）",
        right_title="前年同期比（YoY）",
        left_items=[
            "クラウド事業: +5.2%",
            "コンサルティング: +3.1%",
            "ライセンス: -1.2%（想定内）",
        ],
        right_items=[
            "クラウド事業: +22.5% ← 主要成長源",
            "コンサルティング: +8.3%",
            "ライセンス: -5.0%（移行期）",
        ],
        left_color="accent",
        right_color="accent2",
    )

    # --- 5. 四半期実績テーブル ---
    slide = sb.add_body("四半期売上実績", section="業績詳細")
    add_table_shapes(
        slide,
        x=0.5, y=CONTENT_TOP + 0.2, w=12.33,
        headers=["事業セグメント", "Q1実績", "Q1計画", "達成率", "Q1前年", "YoY"],
        rows=[
            ["クラウド事業", "6.2億円", "5.9億円", "105.1%", "5.1億円", "+22.5%"],
            ["コンサルティング", "4.1億円", "4.0億円", "102.5%", "3.8億円", "+8.3%"],
            ["ライセンス", "2.1億円", "2.3億円", "91.3%", "2.2億円", "-4.5%"],
            ["合計", "12.4億円", "12.2億円", "101.6%", "11.1億円", "+8.2%"],
        ],
        col_widths=[2.8, 1.8, 1.8, 1.5, 1.8, 1.55],
    )

    # --- 6. ロードマップ ---
    slide = sb.add_body("Q2-Q4 事業計画", section="見通し")
    add_timeline(slide, y=CONTENT_TOP + 1.2, events=[
        {"date": "2026年4月", "title": "新機能リリース", "desc": "クラウドv3.0\nβ提供開始"},
        {"date": "2026年6月", "title": "パートナー拡大", "desc": "販売代理店\n10社追加"},
        {"date": "2026年9月", "title": "海外展開", "desc": "ASEANパイロット\n開始"},
        {"date": "2026年12月", "title": "年間目標達成", "desc": "売上50億円\n目標"},
    ])
    add_text(slide, 0.5, CONTENT_TOP + 3.5, 12.33, 0.4,
             "Q2以降も2桁成長を維持。クラウド事業のARR拡大とASEAN展開が主要ドライバー。",
             style="body")

    # --- 7. クロージング ---
    slide = sb._new_blank_slide()
    from lib import add_rect
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill="accent")
    add_rect(slide, 0, SLIDE_H - 0.25, SLIDE_W, 0.25, fill="white")
    add_text(slide, 1.0, 2.0, SLIDE_W - 2.0, 1.4,
             "ご清聴ありがとうございました", style="title_cover", align="center", font_size=32)
    add_text(slide, 1.0, 3.6, SLIDE_W - 2.0, 0.6,
             "Q2も引き続きよろしくお願いいたします", style="subtitle_cover", align="center")
    add_line(slide, 3.0, 4.4, SLIDE_W - 3.0, 4.4, color="white", width=1.5)
    add_text(slide, 1.0, 4.7, SLIDE_W - 2.0, 0.35,
             "XX株式会社 経営企画部 | ir@example.co.jp",
             style="caption", align="center", color="white")

    result = sb.save_and_validate(OUTPUT)
    print(f"\n生成完了: {OUTPUT}")
    print(f"スライド数: {len(sb.prs.slides)}")
    if result.ok:
        print("バリデーション: OK")
    else:
        print(f"バリデーション: {len(result.errors)} エラー")


if __name__ == "__main__":
    main()
