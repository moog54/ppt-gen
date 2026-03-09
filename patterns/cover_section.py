"""
パターン名: cover_section
カテゴリ: 表紙
説明: セクション区切りスライド（番号付き左帯）
用途: 章・セクションの開始、アジェンダ区切り
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder


def run(
    title: str = "Section 1",
    subtitle: str = "このセクションの概要説明",
    number: int = 1,
    output_path: str = "output/cover_section.pptx",
):
    sb = SlideBuilder()
    sb.add_section(title, subtitle, number)
    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
