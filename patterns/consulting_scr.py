"""
パターン名: consulting_scr
カテゴリ: コンサルティング
説明: Situation / Complication / Resolution の3段構造スライド
用途: 課題提起と提言、戦略的メッセージの論理構造
ドキュメント: 戦略提言, 提案書, フィンディングス, エグゼクティブブリーフィング
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_rect, add_rounded_rect, add_line, add_badge, C, CONTENT_TOP, SLIDE_W


def run(
    title: str = "戦略的提言",
    section: str = "",
    situation: str = "現状、市場環境は急速に変化しており、従来の事業モデルでは競争優位を維持することが困難になっている。",
    complication: str = "特にデジタル化の遅れにより、顧客獲得コストが上昇し、既存顧客の離反リスクが高まっている。",
    resolution: str = "DX推進による業務効率化とデジタルチャネル強化を通じ、3年以内にコスト構造を抜本的に改善する。",
    output_path: str = "output/consulting_scr.pptx",
):
    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    block_w = 12.33
    block_x = 0.5
    y = CONTENT_TOP + 0.2

    labels = [
        ("S", "Situation", "現状認識", situation, "accent2"),
        ("C", "Complication", "課題・論点", complication, "accent3"),
        ("R", "Resolution", "提言・解決策", resolution, "accent"),
    ]

    block_h = 1.55
    gap = 0.18

    for i, (letter, eng, jpn, body, color) in enumerate(labels):
        bx = block_x
        by = y + i * (block_h + gap)

        # 背景カード
        add_rounded_rect(slide, bx, by, block_w, block_h, fill="bgLight", border="border")

        # 左アクセントバー
        add_rect(slide, bx, by, 0.07, block_h, fill=color)

        # ラベルバッジ
        circle, t = add_badge(slide, bx + 0.18, by + 0.2, letter, color=color)

        # 英語ラベルと日本語ラベル
        add_text(slide, bx + 0.75, by + 0.1, 2.5, 0.38,
                 eng, style="subheading", color=color)
        add_text(slide, bx + 0.75, by + 0.5, 2.5, 0.3,
                 jpn, style="caption", color="textMuted")

        # 縦区切り線
        add_line(slide, bx + 3.3, by + 0.15, bx + 3.3, by + block_h - 0.15,
                 color="border", width=0.75)

        # 本文
        add_text(slide, bx + 3.5, by + 0.2, block_w - 3.7, block_h - 0.4,
                 body, style="body", word_wrap=True)

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
