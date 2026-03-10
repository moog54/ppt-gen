"""
パターン名: closing
カテゴリ: 表紙
説明: 最終スライド（Thank you / 次のステップ）
用途: プレゼン締め、連絡先、次のアクション提示
ドキュメント: ドアノッカー, 提案書, SOW, ビジネスケース, 定例報告, ステアリングコミッティ, ワークショップ, フィンディングス, 戦略提言, エグゼクティブブリーフィング
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_rect, add_rounded_rect, add_line, C, SLIDE_W, SLIDE_H


def run(
    main_message: str = "ご清聴ありがとうございました",
    subtitle: str = "ご質問・ご相談はお気軽にどうぞ",
    contact_name: str = "担当者名",
    contact_email: str = "contact@example.com",
    contact_phone: str = "",
    next_steps: list[str] | None = None,
    output_path: str = "output/closing.pptx",
):
    if next_steps is None:
        next_steps = [
            "追加資料のご提供",
            "詳細ヒアリングのご調整",
            "提案書の作成・ご提出",
        ]

    sb = SlideBuilder()
    slide = sb._new_blank_slide()

    # 背景
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill="accent")
    add_rect(slide, 0, SLIDE_H * 0.55, SLIDE_W, SLIDE_H * 0.45, fill="bg")

    # メインメッセージ
    add_text(slide, 1.0, 0.8, SLIDE_W - 2.0, 1.4,
             main_message, style="title_cover", align="center",
             font_size=min(36, max(24, 48 - len(main_message) // 4)))
    add_text(slide, 1.0, 2.3, SLIDE_W - 2.0, 0.6,
             subtitle, style="subtitle_cover", align="center")

    # デコライン
    add_line(slide, 3.0, 3.1, SLIDE_W - 3.0, 3.1, color="white", width=1.5)

    # 次のステップ（下部白背景）
    if next_steps:
        add_text(slide, 0.7, SLIDE_H * 0.57, 6.0, 0.4,
                 "次のステップ", style="label", color="accent", bold=True)
        for i, step in enumerate(next_steps):
            sy = SLIDE_H * 0.57 + 0.5 + i * 0.45
            from lib import add_badge
            add_badge(slide, 0.7, sy, i + 1, color="accent")
            add_text(slide, 1.3, sy + 0.05, 5.5, 0.38, step, style="body")

    # 連絡先
    contact_x = 7.5
    contact_w = 5.2  # 7.5 + 5.2 = 12.7 < SAFE_XMAX
    contact_y = SLIDE_H * 0.57
    add_text(slide, contact_x, contact_y, contact_w, 0.4,
             "お問い合わせ", style="label", color="accent", bold=True)
    add_text(slide, contact_x, contact_y + 0.5, contact_w, 0.4,
             contact_name, style="subheading", color="text")
    if contact_email:
        add_text(slide, contact_x, contact_y + 0.95, contact_w, 0.35,
                 contact_email, style="body", color="accent")
    if contact_phone:
        add_text(slide, contact_x, contact_y + 1.3, contact_w, 0.35,
                 contact_phone, style="body", color="textLight")

    # 下部ロゴエリア
    add_line(slide, 0.5, SLIDE_H - 0.5, SLIDE_W - 0.5, SLIDE_H - 0.5, color="border")
    add_text(slide, 0.5, SLIDE_H - 0.45, SLIDE_W - 1.0, 0.3,
             "Confidential", style="caption", color="textMuted", align="right")

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
