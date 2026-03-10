"""
パターン名: consulting_pricing
カテゴリ: コンサルティング
説明: 料金・工数テーブル（フェーズ別・役割別）
用途: SOWの費用明細、投資額の内訳提示
ドキュメント: SOW, 提案書, ビジネスケース
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_rect, add_line, add_table_shapes, C, CONTENT_TOP, SLIDE_W


def run(
    title: str = "投資額・工数内訳",
    section: str = "",
    rows: list[list] | None = None,
    total: str = "¥42,000,000",
    note: str = "※ 費用は税別。交通費・宿泊費等の実費は別途請求。",
    output_path: str = "output/consulting_pricing.pptx",
):
    if rows is None:
        rows = [
            ["Phase 1 現状診断",    "4週間",  "PM×0.5 + Sr.×1 + Jr.×2", "80人日",  "¥8,000,000"],
            ["Phase 2 戦略策定",    "6週間",  "PM×0.5 + Sr.×2 + Jr.×2", "130人日", "¥16,000,000"],
            ["Phase 3 実行計画",    "4週間",  "PM×0.5 + Sr.×1 + Jr.×2", "80人日",  "¥8,000,000"],
            ["プロジェクト管理",    "全期間", "PM×0.3",                   "30人日",  "¥6,000,000"],
            ["ナレッジトランスファー", "最終週", "Sr.×1",                  "5人日",   "¥2,000,000"],
            ["予備費（10%）",       "—",      "—",                        "—",       "¥4,000,000"],
        ]

    headers = ["フェーズ / 項目", "期間", "体制", "工数", "費用（税別）"]
    col_widths = [3.2, 1.2, 3.5, 1.4, 2.03]

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    table_y = CONTENT_TOP + 0.2
    add_table_shapes(slide, x=0.5, y=table_y, w=11.33,
                     headers=headers, rows=rows, col_widths=col_widths)

    # 合計行
    total_y = table_y + 0.38 * (len(rows) + 1) + 0.12
    add_rect(slide, 0.5, total_y, 11.33, 0.48, fill="accent")
    add_text(slide, 0.6, total_y + 0.06, 6.0, 0.36,
             "合計（税別）", style="small", color="white", bold=True)
    add_text(slide, 7.5, total_y + 0.06, 4.3, 0.36,
             total, style="subheading", color="white", align="right")

    # 注記
    add_text(slide, 0.5, total_y + 0.58, 12.33, 0.3,
             note, style="caption", color="textMuted")

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
