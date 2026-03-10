"""
パターン名: workshop_exercise
カテゴリ: コンサルティング
説明: ワーク指示スライド（作業説明・グループ構成・制限時間）
用途: ワークショップのグループワーク指示、演習説明
ドキュメント: ワークショップ
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_rect, add_badge, add_line, C, CONTENT_TOP, SLIDE_W


def run(
    title: str = "グループワーク",
    section: str = "",
    theme: str = "自社のDX推進における最大の障壁は何か？",
    objective: str = "現場の視点から課題を特定し、優先度の高い施策を3つ導出する。",
    steps: list[str] | None = None,
    duration_min: int = 20,
    group_size: int = 4,
    output_path: str = "output/workshop_exercise.pptx",
):
    if steps is None:
        steps = [
            "個人で課題を付箋に書き出す（5分）",
            "グループ内で共有・グルーピングする（8分）",
            "優先度トップ3を選定して発表準備（7分）",
        ]

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    y = CONTENT_TOP + 0.15

    # テーマブロック
    add_rect(slide, 0.5, y, 12.33, 0.9, fill="accent")
    add_text(slide, 0.7, y + 0.08, 11.93, 0.3,
             "ワークテーマ", style="caption", color="white")
    add_text(slide, 0.7, y + 0.38, 11.93, 0.46,
             theme, style="subheading", color="white", align="center")

    y += 1.05

    # 左カラム：目的 + ステップ
    col_w = 7.5
    add_rect(slide, 0.5, y, col_w, 4.35, fill="bgLight", border="border")
    add_rect(slide, 0.5, y, col_w, 0.38, fill="accent")
    add_text(slide, 0.65, y + 0.06, col_w - 0.3, 0.28,
             "目的", style="caption", color="white", bold=True)
    add_text(slide, 0.65, y + 0.46, col_w - 0.3, 0.5,
             objective, style="body", word_wrap=True)

    add_line(slide, 0.65, y + 1.08, 0.5 + col_w - 0.2, y + 1.08, color="border")
    add_text(slide, 0.65, y + 1.15, col_w - 0.3, 0.3,
             "進め方", style="small", bold=True, color="accent")
    for j, step in enumerate(steps):
        sy = y + 1.5 + j * 0.85
        add_badge(slide, 0.65, sy, j + 1, color="accent")
        add_text(slide, 1.25, sy, col_w - 0.9, 0.75, step, style="body", word_wrap=True)

    # 右カラム：時間・グループ
    rx = 0.5 + col_w + 0.2
    rw = 12.33 - col_w - 0.2

    add_rect(slide, rx, y, rw, 2.1, fill="accent", border=None)
    add_text(slide, rx + 0.15, y + 0.15, rw - 0.3, 0.35,
             "制限時間", style="caption", color="white")
    add_text(slide, rx + 0.15, y + 0.5, rw - 0.3, 1.1,
             str(duration_min), style="kpi", color="white", align="center", font_size=64)
    add_text(slide, rx + 0.15, y + 1.65, rw - 0.3, 0.35,
             "分", style="subheading", color="white", align="center")

    add_rect(slide, rx, y + 2.25, rw, 2.1, fill="bgLight", border="border")
    add_rect(slide, rx, y + 2.25, rw, 0.38, fill="accent2")
    add_text(slide, rx + 0.15, y + 2.31, rw - 0.3, 0.28,
             "グループ構成", style="caption", color="white", bold=True)
    add_text(slide, rx + 0.15, y + 2.72, rw - 0.3, 0.8,
             str(group_size), style="kpi", color="accent2", align="center", font_size=48)
    add_text(slide, rx + 0.15, y + 3.55, rw - 0.3, 0.3,
             "名 / グループ", style="small", color="textLight", align="center")

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
