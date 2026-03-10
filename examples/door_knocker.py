"""
ドアノッカー — 初回訪問・問題提起型（3〜4枚）

構成: cover → quote → data_kpi → body_2col → closing
テーマ: mckinsey
目的: 初回面談で「うちのことを分かっている」と思わせ、次の会話を引き出す
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import (
    SlideBuilder, add_text, add_rect, add_rounded_rect, add_kpi_row,
    add_comparison, add_quote, add_line,
    C, CONTENT_TOP, SLIDE_W, SLIDE_H,
)

OUTPUT = "output/door_knocker.pptx"


def main():
    sb = SlideBuilder(theme="mckinsey")
    sb.set_footer("Confidential — 貴社限り")

    # --- 1. 表紙 ---
    sb.add_cover(
        title="製造業DXの潮流と\n貴社への示唆",
        subtitle="XX株式会社 / 戦略コンサルティング部門",
        date="2026年4月",
    )

    # --- 2. インサイト（問題提起） ---
    slide = sb.add_body("業界が直面する構造変化", section="エグゼクティブサマリー")
    add_quote(
        slide, x=0.8, y=CONTENT_TOP + 0.3, w=11.5,
        text="製造業の利益率上位25%と下位25%の差は過去10年で2倍に拡大した。"
             "その最大の分岐点は「データ活用による意思決定速度」にある。",
        source="McKinsey Global Institute, 2025",
    )
    add_text(slide, 0.8, CONTENT_TOP + 2.5, 11.5, 0.5,
             "貴社の現状: 意思決定サイクルが競合比1.8倍の長さ。この差を放置すると、"
             "2028年までに市場シェア3〜5pt喪失のリスク。",
             style="body", color="accent3")

    # --- 3. 業界KPI比較 ---
    slide = sb.add_body("業界ベンチマーク対比", section="現状診断")
    add_kpi_row(slide, y=CONTENT_TOP + 0.3, items=[
        {"label": "意思決定サイクル（業界平均）", "value": "8.2",  "unit": "週", "color": "accent2"},
        {"label": "意思決定サイクル（貴社推定）", "value": "14.7", "unit": "週", "color": "accent3"},
        {"label": "デジタル投資比率（業界上位）",  "value": "12",   "unit": "%",  "color": "accent2"},
        {"label": "デジタル投資比率（貴社推定）",  "value": "4",    "unit": "%",  "color": "accent3"},
    ])
    add_text(slide, 0.5, CONTENT_TOP + 2.2, 12.33, 0.4,
             "出所: 業界団体調査・有価証券報告書・求人分析による推計（2026年3月時点）",
             style="caption", color="textMuted")

    # --- 4. ギャップと仮説 ---
    slide = sb.add_body("仮説：3つの構造的ギャップ", section="示唆")
    add_comparison(
        slide, y=CONTENT_TOP + 0.2,
        left_title="業界リーダーの共通特徴",
        right_title="貴社の現状（推定）",
        left_items=[
            "データ基盤を全社統合（単一プラットフォーム）",
            "週次でKPIをリアルタイム可視化",
            "デジタル人材比率 10〜15%",
            "現場判断権限の委譲が進んでいる",
        ],
        right_items=[
            "部門ごとにシステムが分断（7系統以上）",
            "月次レポートに依存、2〜3週のラグ",
            "デジタル人材比率 推定3〜4%",
            "決裁は本部集中、現場判断が困難",
        ],
        left_color="accent2",
        right_color="accent3",
    )

    # --- 5. クロージング ---
    slide = sb._new_blank_slide()
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill="accent")
    add_rect(slide, 0, SLIDE_H - 0.25, SLIDE_W, 0.25, fill="white")
    add_text(slide, 1.0, 1.8, SLIDE_W - 2.0, 1.2,
             "次のステップ", style="title_cover", align="center", font_size=30)
    add_line(slide, 3.0, 3.2, SLIDE_W - 3.0, 3.2, color="white", width=1.0)
    for i, msg in enumerate([
        "30分：貴社の優先課題と経営アジェンダを伺う",
        "仮説の精度を上げ、診断スコープを提案する",
        "必要に応じ、類似事例（製造業3社）をご紹介",
    ]):
        add_text(slide, 2.5, 3.5 + i * 0.7, SLIDE_W - 5.0, 0.55,
                 f"{i+1}.  {msg}", style="body", color="white")
    add_text(slide, 1.0, SLIDE_H - 1.0, SLIDE_W - 2.0, 0.35,
             "XX株式会社 担当: 山田太郎  |  taro.yamada@example.com  |  090-XXXX-XXXX",
             style="caption", color="white", align="center")

    result = sb.save_and_validate(OUTPUT)
    print(f"\n生成完了: {OUTPUT}  ({len(sb.prs.slides)}枚)")


if __name__ == "__main__":
    main()
