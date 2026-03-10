"""
パターン名: flow_timeline
カテゴリ: フロー
説明: 横タイムライン（日時付きマイルストーン）
用途: プロジェクト計画、ロードマップ、スケジュール
ドキュメント: SOW, ビジネスケース, 定例報告, 戦略提言, ステアリングコミッティ
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_timeline, C, CONTENT_TOP


def run(
    title: str = "プロジェクトロードマップ",
    section: str = "",
    events: list[dict] | None = None,
    output_path: str = "output/flow_timeline.pptx",
):
    """
    events: [{"date": str, "title": str, "desc": str}]
    """
    if events is None:
        events = [
            {"date": "2026年4月", "title": "プロジェクト開始", "desc": "チーム編成・計画策定"},
            {"date": "2026年6月", "title": "現状分析完了", "desc": "AS-IS整理・課題特定"},
            {"date": "2026年9月", "title": "施策実行開始", "desc": "パイロット展開"},
            {"date": "2026年12月", "title": "全社展開", "desc": "定着化・効果測定"},
            {"date": "2027年3月", "title": "評価・改善", "desc": "フェーズ2計画"},
        ]

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    add_timeline(slide, y=CONTENT_TOP + 1.5, events=events)

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
