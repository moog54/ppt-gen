"""
パターン名: data_comparison
カテゴリ: データ
説明: 左右2列比較表（メリット/デメリット、現状/目標など）
用途: 比較分析、課題と施策、現状と目標
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_comparison, add_text, C, CONTENT_TOP


def run(
    title: str = "比較分析",
    section: str = "",
    left_title: str = "現状",
    right_title: str = "目標",
    left_items: list[str] | None = None,
    right_items: list[str] | None = None,
    left_color: str = "accent3",
    right_color: str = "accent2",
    note: str = "",
    output_path: str = "output/data_comparison.pptx",
):
    if left_items is None:
        left_items = [
            "人材採用に時間がかかっている",
            "スキルのミスマッチが発生",
            "退職率が業界平均を上回る",
        ]
    if right_items is None:
        right_items = [
            "採用リードタイムを30日以内に短縮",
            "スキル要件の明確化と教育強化",
            "エンゲージメント向上で退職率改善",
        ]

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    add_comparison(
        slide,
        y=CONTENT_TOP + 0.2,
        left_title=left_title,
        right_title=right_title,
        left_items=left_items,
        right_items=right_items,
        left_color=left_color,
        right_color=right_color,
    )

    if note:
        add_text(slide, 0.5, 6.5, 12.33, 0.35, f"※ {note}", style="caption", color="textMuted")

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
