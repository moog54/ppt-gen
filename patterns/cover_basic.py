"""
パターン名: cover_basic
カテゴリ: 表紙
説明: シンプルなカラー背景の表紙スライド
用途: プレゼン開始、報告書表紙、提案書表紙
ドキュメント: ドアノッカー, 提案書, SOW, ビジネスケース, 定例報告, ステアリングコミッティ, ワークショップ, フィンディングス, 戦略提言, エグゼクティブブリーフィング
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder


def run(
    title: str = "プレゼンテーションタイトル",
    subtitle: str = "サブタイトル・組織名",
    date: str = "2026年4月",
    output_path: str = "output/cover_basic.pptx",
    accent: str | None = None,
):
    theme = {"accent": accent} if accent else {}
    sb = SlideBuilder(theme=theme or None)
    sb.add_cover(title, subtitle, date)
    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
