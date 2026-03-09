"""
パターン名: body_2col
カテゴリ: 本文
説明: 2カラム本文レイアウト（左右に内容を並列配置）
用途: 比較説明、Before/After、2トピック並列
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_rect, add_rounded_rect, C, CONTENT_TOP


def run(
    title: str = "2カラムレイアウト",
    section: str = "",
    left_title: str = "左カラム",
    right_title: str = "右カラム",
    left_bullets: list[str] | None = None,
    right_bullets: list[str] | None = None,
    output_path: str = "output/body_2col.pptx",
):
    if left_bullets is None:
        left_bullets = ["ポイントA-1", "ポイントA-2", "ポイントA-3"]
    if right_bullets is None:
        right_bullets = ["ポイントB-1", "ポイントB-2", "ポイントB-3"]

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    col_w = 6.0
    gap = 0.33
    left_x = 0.5
    right_x = left_x + col_w + gap
    y_base = CONTENT_TOP + 0.15
    card_h = 5.45  # 1.55 + 5.45 = 7.0 ≤ SAFE_YMAX

    for col_x, col_title, bullets in [(left_x, left_title, left_bullets),
                                       (right_x, right_title, right_bullets)]:
        # カラム背景
        rect = add_rounded_rect(slide, col_x, y_base, col_w, card_h, fill="bgLight", border="border")
        # タイトルバー
        add_rect(slide, col_x, y_base, col_w, 0.55, fill="accent")
        add_text(slide, col_x + 0.15, y_base + 0.08, col_w - 0.3, 0.4,
                 col_title, style="subheading", color="white")
        # 箇条書き
        by = y_base + 0.7
        for bullet in bullets:
            dot = slide.shapes.add_shape(9, *[v * 914400 for v in [col_x + 0.2, by + 0.12, 0.1, 0.1]])
            from lib import _add_shape_fill
            _add_shape_fill(dot, "accent")
            add_text(slide, col_x + 0.4, by, col_w - 0.55, 0.45, bullet, style="body")
            by += 0.5

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
