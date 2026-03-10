"""
パターン名: flow_gantt
カテゴリ: フロー
説明: コンサル式ガントチャート（月/週ヘッダー・グループ行・マイルストン・凡例付き）
用途: プロジェクトスケジュール、実施計画、フェーズ管理
ドキュメント: 提案書, SOW, ビジネスケース, 定例報告, ステアリングコミッティ
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_gantt, CONTENT_TOP


def run(
    title: str = "実施スケジュール",
    section: str = "プロジェクトの要旨",
    description: str = "各作業の実施スケジュールは以下の通りである。",
    periods: list[dict] | None = None,
    rows: list[dict] | None = None,
    legend: dict | None = None,
    output_path: str = "output/flow_gantt.pptx",
):
    """
    periods 例:
        [
            {"label": "10月", "subperiods": ["3週", "10週", "17週", "24週", "31週"]},
            {"label": "11月", "subperiods": ["7週", "14週", "21週", "28週"]},
        ]

    rows 例:
        [
            {"label": "マイルストン", "milestones": [
                {"col": 0,  "label": "契約開始"},
                {"col": 6,  "label": "キックオフ実施"},
            ]},
            {"label": "プロジェクト\n管理", "tasks": [
                {"start": 0, "end": 3,  "label": "事業実施計画書作成", "color": "self"},
                {"start": 4, "end": 16, "label": "ヒアリング結果報告書/PoC機能要件定義書納品", "color": "self"},
                {"start": 0, "end": 19, "label": "週次報告", "color": "self"},
            ]},
            {"label": "機能調査\n拡張", "group": "調査研究", "tasks": [
                {"start": 2, "end": 5,  "label": "調査対象機能候補の抽出", "color": "self"},
                {"start": 4, "end": 7,  "label": "調査対象機能の選定", "color": "joint"},
            ]},
        ]

    legend 例:
        {"弊社作業": "self", "貴社作業": "client", "合同": "joint"}
    """
    # デフォルト: スクリーンショット相当のサンプルデータ
    if periods is None:
        periods = [
            {"label": "10月", "subperiods": ["3週", "10週", "17週", "24週", "31週"]},
            {"label": "11月", "subperiods": ["7週", "14週", "21週", "28週"]},
            {"label": "12月", "subperiods": ["5週", "12週", "19週", "26週"]},
            {"label": "1月",  "subperiods": ["2週", "9週", "16週", "23週", "30週"]},
            {"label": "2月",  "subperiods": ["6週", "13週", "20週", "27週"]},
            {"label": "3月",  "subperiods": ["6週", "13週", "20週", "27週"]},
        ]

    if rows is None:
        rows = [
            {"label": "マイルストン", "milestones": [
                {"col": 0,  "label": "契約開始"},
                {"col": 6,  "label": "キックオフ実施"},
                {"col": 10, "label": "PoC機能決定"},
                {"col": 17, "label": "実機確認開始"},
                {"col": 26, "label": "契約終了"},
            ]},
            {"label": "プロジェクト\n管理", "tasks": [
                {"start": 0,  "end": 3,  "label": "事業実施計画書作成", "color": "self"},
                {"start": 4,  "end": 16, "label": "ヒアリング結果報告書/PoC機能要件定義書納品", "color": "self"},
                {"start": 0,  "end": 25, "label": "週次報告", "color": "self"},
                {"start": 23, "end": 26, "label": "最終成果物納品", "color": "self"},
            ]},
            {"label": "機能調査\n拡張", "group": "調査研究", "tasks": [
                {"start": 2,  "end": 6,  "label": "調査対象機能候補の抽出", "color": "self"},
                {"start": 4,  "end": 8,  "label": "調査対象機能の選定", "color": "joint"},
                {"start": 6,  "end": 10, "label": "調査対象機能の要件詳細確認", "color": "self"},
                {"start": 9,  "end": 14, "label": "調査対象機能の技術検証", "color": "self"},
            ]},
            {"label": "PoC\n実施", "group": "調査研究", "tasks": [
                {"start": 16, "end": 18, "label": "実機確認（初回）", "color": "joint"},
                {"start": 18, "end": 20, "label": "結果とりまとめ、改修", "color": "self"},
                {"start": 19, "end": 20, "label": "課題確認", "color": "client"},
                {"start": 21, "end": 23, "label": "実機確認（2回目）", "color": "joint"},
                {"start": 22, "end": 24, "label": "報告書作成", "color": "self"},
            ]},
            {"label": "資料作成\n引継ぎ", "tasks": [
                {"start": 23, "end": 26, "label": "資料作成・引継ぎ", "color": "self"},
            ]},
        ]

    if legend is None:
        legend = {"弊社作業": "self", "貴社作業": "client", "合同": "joint"}

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    y = CONTENT_TOP + 0.1
    if description:
        add_text(slide, 0.5, y, 12.33, 0.3, description, style="body")
        y += 0.38

    add_gantt(
        slide,
        x=0.4,
        y=y,
        w=12.4,
        periods=periods,
        rows=rows,
        legend=legend,
        header_color="accent",
        row_h=0.42,
        label_w=1.05,
        group_w=0.42,
    )

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
