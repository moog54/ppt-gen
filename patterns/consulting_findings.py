"""
パターン名: consulting_findings
カテゴリ: コンサルティング
説明: 発見事項リスト（番号・重要度バッジ付き）
用途: 診断結果、課題一覧、インタビュー発見事項の提示
ドキュメント: フィンディングス, 提案書, 戦略提言, ドアノッカー
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_rect, add_rounded_rect, add_badge, C, CONTENT_TOP, SLIDE_W

SEVERITY_COLORS = {"高": "accent3", "中": "accent", "低": "accent2"}


def run(
    title: str = "主要発見事項",
    section: str = "",
    findings: list[dict] | None = None,
    output_path: str = "output/consulting_findings.pptx",
):
    if findings is None:
        findings = [
            {"severity": "高", "title": "デジタル基盤の老朽化", "body": "基幹システムが15年以上前の設計のため、API連携・データ活用が困難。モダナイズなしでは競合との差が拡大する。"},
            {"severity": "高", "title": "人材ポートフォリオの偏在", "body": "デジタル人材が全社員の3%未満。業界平均（12%）を大きく下回り、DX推進の実行力が不足している。"},
            {"severity": "中", "title": "顧客データの分散・非統合", "body": "CRM・基幹・ECの顧客データが未統合。パーソナライゼーション施策の精度に限界があり、LTV向上の機会を逸失。"},
            {"severity": "低", "title": "社内承認プロセスの長期化", "body": "新規施策の意思決定に平均6.2週間。スタートアップとの競争において機動力の差が生じている。"},
        ]

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    y = CONTENT_TOP + 0.2
    row_h = 1.2
    gap = 0.1

    for i, f in enumerate(findings):
        by = y + i * (row_h + gap)
        severity = f.get("severity", "中")
        color = SEVERITY_COLORS.get(severity, "accent")

        # カード背景
        add_rounded_rect(slide, 0.5, by, 12.33, row_h, fill="bgLight", border="border")
        add_rect(slide, 0.5, by, 0.07, row_h, fill=color)

        # 番号バッジ
        add_badge(slide, 0.68, by + 0.42, i + 1, color=color)

        # 重要度バッジ
        sev_bg = color
        add_rounded_rect(slide, 1.3, by + 0.38, 0.7, 0.32, fill=sev_bg)
        add_text(slide, 1.32, by + 0.4, 0.66, 0.28,
                 f"重要度:{severity}", style="caption", color="white", bold=True, font_size=9)

        # タイトル
        add_text(slide, 2.15, by + 0.1, 10.5, 0.45,
                 f.get("title", ""), style="subheading", color=color)

        # 本文
        add_text(slide, 2.15, by + 0.58, 10.5, 0.65,
                 f.get("body", ""), style="body", word_wrap=True)

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
