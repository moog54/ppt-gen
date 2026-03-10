"""
パターン名: data_content_box
カテゴリ: データ
説明: コンテンツボックス型比較表（各列が独立したカード状ボックス）
用途: 施策比較、オプション評価、選択肢の整理
ドキュメント: 提案書, ビジネスケース, フィンディングス, 戦略提言, エグゼクティブブリーフィング
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_content_box_table, C, CONTENT_TOP


def run(
    title: str = "施策比較",
    section: str = "",
    headers: list[str] | None = None,
    rows: list[dict] | None = None,
    highlight_col: int | None = 0,
    highlight_label: str = "推奨",
    note: str = "",
    output_path: str = "output/data_content_box.pptx",
):
    """
    headers: 列ヘッダー（選択肢・施策名など）
    rows:    [{"label": str, "values": [str, ...]}]
    highlight_col: ハイライトする列番号（0始まり、None でなし）
    """
    if headers is None:
        headers = ["施策A（内製化）", "施策B（外注併用）", "施策C（SaaS導入）"]
    if rows is None:
        rows = [
            {"label": "初期コスト",   "values": ["低",          "中",           "高"]},
            {"label": "運用コスト",   "values": ["高",          "中",           "低"]},
            {"label": "実施期間",     "values": ["6〜12ヶ月",   "3〜6ヶ月",     "1〜3ヶ月"]},
            {"label": "効果・品質",   "values": ["高（長期）",  "中〜高",       "中（標準）"]},
            {"label": "リスク",       "values": ["人材依存",    "品質管理",     "ベンダー依存"]},
            {"label": "スケーラビリティ", "values": ["低",      "中",           "高"]},
            {"label": "推奨度",       "values": ["△ 中長期向け", "◎ バランス型", "○ 即効性重視"]},
        ]

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    add_content_box_table(
        slide,
        x=0.5,
        y=CONTENT_TOP + 0.45,
        w=12.33,
        headers=headers,
        rows=rows,
        highlight_col=highlight_col,
        highlight_label=highlight_label,
        label_w=1.8,
        header_h=0.55,
        row_h=0.5,
        gap=0.15,
    )

    if note:
        note_y = CONTENT_TOP + 0.45 + 0.55 + 0.5 * len(rows) + 0.15
        add_text(slide, 0.5, note_y, 12.33, 0.32,
                 f"※ {note}", style="caption", color="textMuted")

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
