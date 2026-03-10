"""
パターン名: consulting_scope
カテゴリ: コンサルティング
説明: スコープ・デリバラブル定義表（フェーズ×項目）
用途: SOWのスコープ定義、作業範囲と成果物の明示
ドキュメント: SOW, 提案書
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_rect, add_line, C, CONTENT_TOP, SLIDE_W


def run(
    title: str = "スコープ・デリバラブル定義",
    section: str = "",
    phases: list[dict] | None = None,
    output_path: str = "output/consulting_scope.pptx",
):
    if phases is None:
        phases = [
            {
                "phase": "Phase 1\n現状診断",
                "duration": "4週間",
                "deliverables": ["現状分析レポート", "課題一覧（優先度付き）", "ベンチマーク比較"],
                "in_scope": ["ヒアリング（経営層・現場）", "データ収集・分析", "業界ベンチマーク調査"],
                "out_scope": ["システム実装", "組織変更"],
            },
            {
                "phase": "Phase 2\n戦略策定",
                "duration": "6週間",
                "deliverables": ["変革ロードマップ", "施策優先度マトリクス", "投資対効果試算"],
                "in_scope": ["戦略オプション立案", "ステークホルダー合意形成", "KPI設計"],
                "out_scope": ["実行支援", "ベンダー選定"],
            },
            {
                "phase": "Phase 3\n実行計画",
                "duration": "4週間",
                "deliverables": ["詳細実行計画書", "体制・ガバナンス設計", "クイックウィン特定"],
                "in_scope": ["実行計画詳細化", "推進体制設計", "リスク対応計画"],
                "out_scope": ["実装・開発作業", "運用保守"],
            },
        ]

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    n = len(phases)
    col_w = 12.33 / n
    x = 0.5
    y = CONTENT_TOP + 0.2

    header_h = 0.7
    row_h = 1.4
    label_w = 1.1

    for i, ph in enumerate(phases):
        cx = x + i * col_w
        color = ["accent", "accent2", "accent3"][i % 3]

        # フェーズヘッダー
        add_rect(slide, cx, y, col_w - 0.08, header_h, fill=color)
        add_text(slide, cx + 0.1, y + 0.05, col_w - 0.28, 0.42,
                 ph["phase"], style="small", color="white", bold=True)
        add_text(slide, cx + 0.1, y + 0.45, col_w - 0.28, 0.22,
                 f"期間: {ph['duration']}", style="caption", color="white")

        row_y = y + header_h + 0.1

        # デリバラブル
        add_rect(slide, cx, row_y, col_w - 0.08, row_h, fill="bgLight", border="border")
        add_rect(slide, cx, row_y, col_w - 0.08, 0.28, fill=color)
        add_text(slide, cx + 0.08, row_y + 0.04, col_w - 0.2, 0.22,
                 "成果物", style="caption", color="white", bold=True)
        for j, d in enumerate(ph["deliverables"]):
            add_text(slide, cx + 0.12, row_y + 0.32 + j * 0.34, col_w - 0.24, 0.32,
                     f"• {d}", style="small", word_wrap=True)

        row_y2 = row_y + row_h + 0.1

        # スコープ内
        add_rect(slide, cx, row_y2, col_w - 0.08, row_h, fill="bg", border="accent2")
        add_rect(slide, cx, row_y2, col_w - 0.08, 0.28, fill="accent2")
        add_text(slide, cx + 0.08, row_y2 + 0.04, col_w - 0.2, 0.22,
                 "対象範囲（In Scope）", style="caption", color="white", bold=True)
        for j, s in enumerate(ph["in_scope"]):
            add_text(slide, cx + 0.12, row_y2 + 0.32 + j * 0.32, col_w - 0.24, 0.30,
                     f"✓ {s}", style="small", color="accent2", word_wrap=True)

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
