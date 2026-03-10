"""
パターン名: quote
カテゴリ: 強調
説明: 引用・インサイト強調スライド
用途: キーメッセージ強調、インタビュー引用、重要な洞察
ドキュメント: ドアノッカー, 提案書, 戦略提言, エグゼクティブブリーフィング
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_quote, add_rect, add_kpi_row, C, CONTENT_TOP


def run(
    title: str = "キーインサイト",
    section: str = "",
    quote_text: str = "デジタル変革の本質は技術ではなく、人と組織の変革にある。最先端のツールを導入しても、使う人材が育たなければ競争優位は生まれない。",
    source: str = "〇〇社 代表取締役CEO",
    supporting_kpis: list[dict] | None = None,
    output_path: str = "output/quote.pptx",
):
    if supporting_kpis is None:
        supporting_kpis = [
            {"label": "DX推進担当者数", "value": "38%", "unit": "不足", "color": "accent3"},
            {"label": "デジタルスキル研修", "value": "2.1x", "unit": "業界平均比", "color": "accent"},
            {"label": "変革プロジェクト完遂率", "value": "67%", "unit": "", "color": "accent2"},
        ]

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    # 引用ブロック
    add_quote(slide, x=1.0, y=CONTENT_TOP + 0.2, w=11.33,
              text=quote_text, source=source)

    # サポートKPI
    if supporting_kpis:
        add_kpi_row(slide, y=CONTENT_TOP + 2.8, items=supporting_kpis,
                    x_start=1.5, total_w=10.33)

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
