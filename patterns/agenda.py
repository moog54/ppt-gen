"""
パターン名: agenda
カテゴリ: 構造
説明: アジェンダ（目次）スライド（現在地ハイライト付き）
用途: プレゼン冒頭の目次、セクション開始時の案内
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_agenda, C, CONTENT_TOP, SLIDE_W


def run(
    title: str = "アジェンダ",
    section: str = "",
    items: list[str] | None = None,
    current: int | None = None,
    output_path: str = "output/agenda.pptx",
):
    if items is None:
        items = [
            "現状分析・課題整理",
            "市場環境と競合動向",
            "施策提案と優先順位",
            "実行計画とKPI",
            "まとめと次のステップ",
        ]

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    # 左側デコレーションライン
    from lib import add_rect
    add_rect(slide, 0.5, CONTENT_TOP + 0.1, 0.06, 5.0, fill="accent")

    add_agenda(
        slide,
        items=items,
        current=current,
        x_start=0.75,
        y_start=CONTENT_TOP + 0.2,
        item_h=0.75,
        w=11.8,
    )

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
