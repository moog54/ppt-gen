"""
パターン名: flow_process
カテゴリ: フロー
説明: 横並びプロセスフロー（矢印付き箱）
用途: 業務フロー、プロセス説明、ステップ解説
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_flow_row, add_rect, C, CONTENT_TOP


def run(
    title: str = "プロセスフロー",
    section: str = "",
    steps: list[str] | None = None,
    descriptions: list[str] | None = None,
    color: str = "accent",
    output_path: str = "output/flow_process.pptx",
):
    if steps is None:
        steps = ["現状分析", "課題特定", "施策立案", "実行", "評価"]
    if descriptions is None:
        descriptions = [
            "データ収集・\nAS-IS整理",
            "ボトルネック\n特定",
            "TO-BE設計\nKPI設定",
            "推進体制\n整備",
            "効果測定\n改善",
        ]

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    y_flow = CONTENT_TOP + 0.8
    add_flow_row(slide, y_flow, steps, color=color, box_h=1.0)

    # 各ステップの説明
    if descriptions:
        n = len(steps)
        arrow_w = 0.3
        total_w = 12.33
        box_w = (total_w - arrow_w * (n - 1)) / n
        for i, desc in enumerate(descriptions[:n]):
            cx = 0.5 + i * (box_w + arrow_w)
            add_text(slide, cx + 0.05, y_flow + 1.1, box_w - 0.1, 0.8,
                     desc, style="caption", align="center", color="textLight")

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
