"""
パターン名: data_kpi
カテゴリ: データ
説明: KPI数値カードを横並びに表示するスライド
用途: 業績サマリー、KPIダッシュボード、数値報告
ドキュメント: ドアノッカー, 定例報告, ステアリングコミッティ, ビジネスケース, エグゼクティブブリーフィング, 戦略提言
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_kpi_row, add_line, C, CONTENT_TOP


def run(
    title: str = "KPI ダッシュボード",
    section: str = "",
    kpis: list[dict] | None = None,
    note: str = "",
    output_path: str = "output/data_kpi.pptx",
):
    """
    kpis: [{"label": str, "value": str, "unit": str, "delta": str, "color": str}]
    """
    if kpis is None:
        kpis = [
            {"label": "売上高", "value": "12.4", "unit": "億円", "delta": "+8.2%", "color": "accent"},
            {"label": "営業利益", "value": "2.1", "unit": "億円", "delta": "+12.5%", "color": "accent2"},
            {"label": "顧客数", "value": "4,820", "unit": "社", "delta": "+350", "color": "accent"},
            {"label": "NPS", "value": "62", "unit": "pt", "delta": "+7pt", "color": "accent2"},
        ]

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    y = CONTENT_TOP + 0.3
    add_kpi_row(slide, y, kpis)

    # 補足テキスト
    if note:
        add_text(slide, 0.5, y + 2.0, 12.33, 0.35,
                 f"※ {note}", style="caption", color="textMuted")

    # 下部サマリーバー
    add_line(slide, 0.5, y + 2.4, 12.83, y + 2.4, color="border")

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
