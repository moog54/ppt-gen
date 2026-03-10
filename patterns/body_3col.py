"""
パターン名: body_3col
カテゴリ: 本文
説明: 3カラム本文レイアウト（3つのポイントを並列配置）
用途: 3つの施策・戦略軸・特徴の並列説明
ドキュメント: 提案書, 戦略提言, エグゼクティブブリーフィング
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_rect, add_rounded_rect, add_pill, C, CONTENT_TOP


def run(
    title: str = "3カラムレイアウト",
    section: str = "",
    columns: list[dict] | None = None,
    output_path: str = "output/body_3col.pptx",
):
    """
    columns: [{"title": str, "icon": str, "bullets": [str], "color": str}]
    """
    if columns is None:
        columns = [
            {"title": "施策1", "bullets": ["詳細A-1", "詳細A-2", "詳細A-3"], "color": "accent"},
            {"title": "施策2", "bullets": ["詳細B-1", "詳細B-2", "詳細B-3"], "color": "accent2"},
            {"title": "施策3", "bullets": ["詳細C-1", "詳細C-2", "詳細C-3"], "color": "accent3"},
        ]

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    n = len(columns)
    gap = 0.2
    total_w = 12.33
    col_w = (total_w - gap * (n - 1)) / n
    y_base = CONTENT_TOP + 0.15
    card_h = 5.45  # 1.55 + 5.45 = 7.0 ≤ SAFE_YMAX

    for i, col in enumerate(columns):
        cx = 0.5 + i * (col_w + gap)
        color = col.get("color", "accent")

        # カード背景
        add_rounded_rect(slide, cx, y_base, col_w, card_h, fill="bg", border="border")
        # ヘッダーカラーバー
        add_rect(slide, cx, y_base, col_w, 0.6, fill=color)
        add_text(slide, cx + 0.1, y_base + 0.1, col_w - 0.2, 0.42,
                 col.get("title", f"施策{i+1}"), style="subheading",
                 align="center", color="white", font_size=16)

        # 番号バッジ
        from lib import add_badge
        add_badge(slide, cx + col_w / 2 - 0.225, y_base + 0.7, i + 1, color=color)

        # 箇条書き
        by = y_base + 1.3
        for bullet in col.get("bullets", []):
            add_text(slide, cx + 0.15, by, col_w - 0.3, 0.45, f"• {bullet}", style="body")
            by += 0.48

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
