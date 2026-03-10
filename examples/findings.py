"""
フィンディングス — 調査・診断結果報告（6枚）

構成: cover → consulting_scr → consulting_findings × 2 → data_comparison → closing
テーマ: default
目的: ヒアリング・分析結果を構造化して経営層に提示する
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import (
    SlideBuilder, add_text, add_rect, add_rounded_rect,
    add_comparison, add_badge, add_line, add_kpi_row,
    C, CONTENT_TOP, SLIDE_W, SLIDE_H,
)

OUTPUT = "output/findings.pptx"


def main():
    sb = SlideBuilder(theme="default")
    sb.set_footer("現状診断レポート — 株式会社サンプル — Confidential")

    # --- 1. 表紙 ---
    sb.add_cover(
        title="現状診断レポート\nPhase 1 フィンディングス",
        subtitle="株式会社サンプル × XX株式会社\n調査期間: 2026年2月〜4月 / ヒアリング対象: 30名",
        date="2026年4月25日",
    )

    # --- 2. エグゼクティブサマリー（SCR） ---
    slide = sb.add_body("エグゼクティブサマリー")
    labels = [
        ("S", "Situation",    "現状",   "全社売上の85%を占める主力事業において、意思決定に必要なデータが分断されており、週次レポートの作成に平均32時間を要している。", "accent2"),
        ("C", "Complication", "課題",   "データ分断により現場の感覚値経営が常態化。市場変化への対応が競合比1.8倍遅く、過去2期連続で計画未達。このまま放置すると2028年に市場シェア3pt喪失のリスク。", "accent3"),
        ("R", "Resolution",   "推奨策", "3フェーズ・18ヶ月のDX変革プログラムを実施。データ基盤統合・意思決定プロセス改革・デジタル人材育成を並走させ、ROI 180%（3年）を実現する。", "accent"),
    ]
    bh, gap = 1.55, 0.18
    for i, (letter, eng, jpn, body, color) in enumerate(labels):
        by = CONTENT_TOP + 0.2 + i * (bh + gap)
        add_rounded_rect(slide, 0.5, by, 12.33, bh, fill="bgLight", border="border")
        add_rect(slide, 0.5, by, 0.07, bh, fill=color)
        add_badge(slide, 0.68, by + 0.55, letter, color=color)
        add_text(slide, 1.3, by + 0.12, 2.2, 0.38, eng, style="subheading", color=color)
        add_text(slide, 1.3, by + 0.52, 2.2, 0.3, jpn, style="caption", color="textMuted")
        add_line(slide, 3.6, by + 0.18, 3.6, by + bh - 0.18, color="border", width=0.75)
        add_text(slide, 3.8, by + 0.22, 8.8, bh - 0.44, body, style="body", word_wrap=True)

    # --- 3. フィンディングス（戦略・組織） ---
    slide = sb.add_body("主要フィンディングス（戦略・組織）", section="診断結果")
    findings_1 = [
        {
            "no": 1, "severity": "Critical", "color": "accent3",
            "title": "中長期戦略と現場KPIの乖離",
            "body": "経営層が掲げる「2028年シェア20%」目標に対し、現場部門の月次KPIが売上・件数中心で非整合。戦略実行の推進力が機能していない。",
            "evidence": "ヒアリング対象30名中22名が「自部門KPIと中期目標の関係が不明」と回答（73%）",
        },
        {
            "no": 2, "severity": "High", "color": "accent",
            "title": "デジタル人材の絶対的不足",
            "body": "データ分析・DX推進スキルを持つ人材が全社でわずか12名（全体比0.8%）。業界平均の1/4水準。採用・育成ともに具体計画なし。",
            "evidence": "競合A社: 3.2%、競合B社: 2.9%に対し、当社は0.8%（出典: 各社有価証券報告書）",
        },
        {
            "no": 3, "severity": "High", "color": "accent",
            "title": "部門縦割りによるサイロ化",
            "body": "営業・製造・物流の各部門がそれぞれ独自システムを運用。顧客データが3システムに分断され、統合ビューが存在しない。",
            "evidence": "データ統合に費やす工数: 月間延べ480時間（アンケート集計）",
        },
    ]
    y = CONTENT_TOP + 0.15
    for f in findings_1:
        card_h = 1.2
        add_rounded_rect(slide, 0.5, y, 12.33, card_h, fill="bgLight", border="border")
        add_rect(slide, 0.5, y, 0.07, card_h, fill=f["color"])
        add_rect(slide, 0.5, y, 12.33, 0.36, fill=f["color"])
        add_badge(slide, 0.65, y + 0.0, f["no"], color="white")
        add_rounded_rect(slide, 1.3, y + 0.04, 1.3, 0.28, fill="white")
        add_text(slide, 1.32, y + 0.06, 1.26, 0.24,
                 f["severity"], style="caption", color=f["color"], bold=True, align="center")
        add_text(slide, 2.75, y + 0.05, 9.8, 0.28,
                 f["title"], style="small", color="white", bold=True)
        add_text(slide, 0.65, y + 0.42, 8.5, 0.65, f["body"], style="body", word_wrap=True)
        add_line(slide, 9.3, y + 0.4, 9.3, y + card_h - 0.08, color="border", width=0.5)
        add_text(slide, 9.4, y + 0.42, 3.2, 0.22, "根拠・エビデンス", style="caption", color="textMuted")
        add_text(slide, 9.4, y + 0.64, 3.2, 0.46, f["evidence"], style="caption", word_wrap=True)
        y += card_h + 0.1

    # --- 4. フィンディングス（IT・データ） ---
    slide = sb.add_body("主要フィンディングス（IT・データ）", section="診断結果")
    findings_2 = [
        {
            "no": 4, "severity": "Critical", "color": "accent3",
            "title": "レガシーシステムによる技術的負債",
            "body": "基幹システムが15年以上前の設計で、API連携不可。クラウド移行も未着手のため、モダンなデータ活用基盤の構築が困難。維持費は年間8億円超。",
            "evidence": "システム更改コスト試算: 現状維持 8.2億/年 vs 移行後 3.1億/年（3年後）",
        },
        {
            "no": 5, "severity": "Medium", "color": "accent2",
            "title": "データガバナンスの不在",
            "body": "データ定義・品質管理ルールが存在せず、同一指標でも部門により数値が異なる。経営会議での「数字の信頼性」議論に毎回1〜2時間を費やしている。",
            "evidence": "直近12回の経営会議のうち9回でデータ不整合に起因する議論が発生",
        },
        {
            "no": 6, "severity": "Medium", "color": "accent2",
            "title": "セキュリティリスクの蓄積",
            "body": "個人情報を含む顧客データが部門ごとに分散管理され、アクセスログ取得も不完全。ISMS認証取得済みながら、実態との乖離が拡大している。",
            "evidence": "内部監査指摘: 2025年度に17件のセキュリティ関連指摘（前年比+42%）",
        },
    ]
    y = CONTENT_TOP + 0.15
    for f in findings_2:
        card_h = 1.2
        add_rounded_rect(slide, 0.5, y, 12.33, card_h, fill="bgLight", border="border")
        add_rect(slide, 0.5, y, 0.07, card_h, fill=f["color"])
        add_rect(slide, 0.5, y, 12.33, 0.36, fill=f["color"])
        add_badge(slide, 0.65, y + 0.0, f["no"], color="white")
        add_rounded_rect(slide, 1.3, y + 0.04, 1.3, 0.28, fill="white")
        add_text(slide, 1.32, y + 0.06, 1.26, 0.24,
                 f["severity"], style="caption", color=f["color"], bold=True, align="center")
        add_text(slide, 2.75, y + 0.05, 9.8, 0.28,
                 f["title"], style="small", color="white", bold=True)
        add_text(slide, 0.65, y + 0.42, 8.5, 0.65, f["body"], style="body", word_wrap=True)
        add_line(slide, 9.3, y + 0.4, 9.3, y + card_h - 0.08, color="border", width=0.5)
        add_text(slide, 9.4, y + 0.42, 3.2, 0.22, "根拠・エビデンス", style="caption", color="textMuted")
        add_text(slide, 9.4, y + 0.64, 3.2, 0.46, f["evidence"], style="caption", word_wrap=True)
        y += card_h + 0.1

    # --- 5. 競合比較 ---
    slide = sb.add_body("競合比較 — DX成熟度", section="競合分析")
    add_comparison(
        slide, y=CONTENT_TOP + 0.2,
        left_title="当社（現状）",
        right_title="業界リーダー（競合A社）",
        left_items=[
            "DX人材比率: 0.8%（業界平均の1/4）",
            "意思決定リードタイム: 平均11日",
            "データ活用部門: 一部（営業のみ）",
            "クラウド利用率: 12%（非基幹のみ）",
            "週次レポート作成工数: 32時間/部門",
        ],
        right_items=[
            "DX人材比率: 3.2%（積極採用継続中）",
            "意思決定リードタイム: 平均6日",
            "データ活用部門: 全社（製造・物流含む）",
            "クラウド利用率: 68%（基幹含む移行完了）",
            "週次レポート作成工数: 4時間/部門（自動化）",
        ],
        left_color="accent3",
        right_color="accent2",
    )
    add_text(slide, 0.5, CONTENT_TOP + 4.45, 12.33, 0.35,
             "出典: 各社有価証券報告書・統合報告書（2025年度）/ DX成熟度調査（XX株式会社実施、2026年3月）",
             style="caption", color="textMuted")

    # --- 6. クロージング ---
    slide = sb._new_blank_slide()
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill="accent")
    add_rect(slide, 0, SLIDE_H - 0.25, SLIDE_W, 0.25, fill="white")
    add_text(slide, 1.0, 1.8, SLIDE_W - 2.0, 1.0,
             "Phase 2 へ", style="title_cover", align="center", font_size=32)
    add_line(slide, 3.0, 3.2, SLIDE_W - 3.0, 3.2, color="white", width=1.0)
    add_text(slide, 1.0, 3.5, SLIDE_W - 2.0, 1.5,
             "6つのフィンディングスを踏まえ、Phase 2 では\n"
             "変革ロードマップと施策優先度マトリクスを策定します。\n"
             "次回ステコミ: 2026年5月15日（予定）",
             style="body", color="white", align="center")

    result = sb.save_and_validate(OUTPUT)
    print(f"\n生成完了: {OUTPUT}  ({len(sb.prs.slides)}枚)")


if __name__ == "__main__":
    main()
