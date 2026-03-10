"""
パターン名: body_1col
カテゴリ: 本文
説明: 1カラム本文レイアウト（タイトル＋箇条書き）
用途: 説明スライド、方針発表、シングルメッセージ
ドキュメント: SOW, ワークショップ, フィンディングス, 定例報告
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_rect, add_line, C, CONTENT_TOP


def run(
    title: str = "スライドタイトル",
    section: str = "セクション名",
    bullets: list[str] | None = None,
    body_text: str = "",
    output_path: str = "output/body_1col.pptx",
):
    if bullets is None:
        bullets = [
            "ポイント1：主要なメッセージをここに記述します",
            "ポイント2：具体的なデータや根拠を補足します",
            "ポイント3：次のアクションや示唆を明示します",
        ]

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    y = CONTENT_TOP + 0.2
    content_x = 0.5
    content_w = 12.33

    if body_text:
        add_text(slide, content_x, y, content_w, 1.0, body_text,
                 style="body", word_wrap=True)
        y += 1.1

    for bullet in bullets:
        # アクセントドット
        dot = slide.shapes.add_shape(9, *[v * 914400 for v in [content_x, y + 0.12, 0.12, 0.12]])
        from lib import _add_shape_fill
        _add_shape_fill(dot, "accent")
        add_text(slide, content_x + 0.25, y, content_w - 0.3, 0.5, bullet, style="body")
        y += 0.55

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
