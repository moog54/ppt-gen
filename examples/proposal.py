"""
提案書 — DX変革コンサルティング提案（7枚）

構成: cover → agenda → quote → consulting_findings → consulting_scr → body_3col → flow_process → closing
テーマ: default
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import (
    SlideBuilder, add_text, add_rect, add_kpi_row, add_agenda,
    add_flow_row, add_quote,
    C, CONTENT_TOP, SLIDE_W, SLIDE_H,
)
from patterns.consulting_findings import run as findings_slide
from patterns.consulting_scr import run as scr_slide

OUTPUT = "output/proposal.pptx"


def main():
    sb = SlideBuilder(theme="default")
    sb.set_footer("XX株式会社 提案書 — Confidential")

    # --- 1. 表紙 ---
    sb.add_cover(
        title="デジタル変革推進\nコンサルティング提案",
        subtitle="株式会社サンプル御中 / XX株式会社",
        date="2026年4月25日",
    )

    # --- 2. アジェンダ ---
    slide = sb.add_body("本日のご提案内容")
    add_agenda(slide, [
        "貴社の現状認識と課題仮説",
        "変革の方向性（SCRアプローチ）",
        "ご提案するソリューション",
        "推進アプローチ・体制",
        "投資対効果・期待成果",
    ], x_start=2.5, y_start=CONTENT_TOP + 0.3, w=8.5)

    # --- 3. 課題インサイト ---
    slide = sb.add_body("貴社の現状認識", section="1. 課題仮説")
    add_quote(
        slide, x=0.8, y=CONTENT_TOP + 0.2, w=11.5,
        text="ヒアリングと公開情報の分析から、貴社の競争力低下は「意思決定速度」と"
             "「データ活用基盤の未整備」に起因すると仮説を立てています。",
        source="XX株式会社 診断仮説（2026年3月）",
    )
    add_kpi_row(slide, y=CONTENT_TOP + 2.1, items=[
        {"label": "意思決定ラグ（競合比）", "value": "1.8×", "unit": "遅い", "color": "accent3"},
        {"label": "データ活用投資不足額",   "value": "▲8",   "unit": "億円/年", "color": "accent3"},
        {"label": "改善による利益ポテンシャル", "value": "+15", "unit": "%", "color": "accent2"},
    ])

    # --- 4. 発見事項 ---
    slide = sb.add_body("主要課題：3つの構造的問題", section="1. 課題仮説")
    from patterns.consulting_findings import run as _f
    # インラインで発見事項を描画
    from lib import add_rounded_rect, add_badge, add_line
    findings = [
        {"severity": "高", "title": "データサイロ化（7部門・システム非統合）",
         "body": "各部門がExcel管理。全社横断の分析に2〜3週間かかり、市場変化への対応が後手に回る。"},
        {"severity": "高", "title": "デジタル人材の絶対的不足（推定3%）",
         "body": "業界平均12%に対し、推定3%程度。プロジェクトを内製化できず、外部依存コストが膨張。"},
        {"severity": "中", "title": "変革推進のガバナンス不在",
         "body": "DX専任組織・KPIが未設定。各部門が個別最適を追い、全社変革が進まない構造になっている。"},
    ]
    SEVERITY_COLORS = {"高": "accent3", "中": "accent", "低": "accent2"}
    row_h, gap = 1.4, 0.18
    for i, f in enumerate(findings):
        by = CONTENT_TOP + 0.2 + i * (row_h + gap)
        color = SEVERITY_COLORS[f["severity"]]
        add_rounded_rect(slide, 0.5, by, 12.33, row_h, fill="bgLight", border="border")
        add_rect(slide, 0.5, by, 0.07, row_h, fill=color)
        add_badge(slide, 0.68, by + 0.48, i + 1, color=color)
        add_rounded_rect(slide, 1.3, by + 0.45, 0.8, 0.32, fill=color)
        add_text(slide, 1.32, by + 0.47, 0.76, 0.28,
                 f"重要度:{f['severity']}", style="caption", color="white", bold=True, font_size=9)
        add_text(slide, 2.25, by + 0.1, 10.4, 0.45, f["title"], style="subheading", color=color)
        add_text(slide, 2.25, by + 0.62, 10.4, 0.65, f["body"], style="body", word_wrap=True)

    # --- 5. SCR構造 ---
    slide = sb.add_body("変革の方向性", section="2. ソリューション")
    from lib import add_line as _line
    labels = [
        ("S", "Situation", "現状認識", "データ分断・人材不足・ガバナンス欠如により、競争力が構造的に低下している。", "accent2"),
        ("C", "Complication", "課題・論点", "このまま放置すると2028年までに市場シェア3〜5pt喪失。競合のDX加速で差は拡大する一方。", "accent3"),
        ("R", "Resolution", "提言", "「データ統合基盤構築」「デジタル人材育成」「CDO設置によるガバナンス整備」の3本柱で変革を推進する。", "accent"),
    ]
    block_h, gap2 = 1.55, 0.18
    for i, (letter, eng, jpn, body, color) in enumerate(labels):
        by = CONTENT_TOP + 0.2 + i * (block_h + gap2)
        add_rounded_rect(slide, 0.5, by, 12.33, block_h, fill="bgLight", border="border")
        add_rect(slide, 0.5, by, 0.07, block_h, fill=color)
        add_badge(slide, 0.68, by + 0.55, letter, color=color)
        add_text(slide, 1.3, by + 0.12, 2.2, 0.38, eng, style="subheading", color=color)
        add_text(slide, 1.3, by + 0.52, 2.2, 0.3, jpn, style="caption", color="textMuted")
        add_line(slide, 3.6, by + 0.18, 3.6, by + block_h - 0.18, color="border", width=0.75)
        add_text(slide, 3.8, by + 0.22, 8.8, block_h - 0.44, body, style="body", word_wrap=True)

    # --- 6. ソリューション3本柱 ---
    slide = sb.add_body("ご提案ソリューション", section="2. ソリューション")
    cols = [
        ("01\nデータ統合基盤", "accent",  ["全社データ統合プラットフォーム構築", "BI/ダッシュボード整備", "リアルタイムKPI可視化", "データガバナンスポリシー策定"]),
        ("02\nデジタル人材育成", "accent2", ["デジタル人材ロードマップ策定", "社内研修プログラム設計", "採用・育成・配置計画", "外部人材との協業モデル"]),
        ("03\nガバナンス整備", "accent3", ["CDO設置・DX推進室立上げ", "変革KPI・管理体制設計", "部門横断プロジェクト管理", "取締役会へのDX報告体制"]),
    ]
    col_w = 3.9
    gap3 = 0.27
    for i, (title, color, items) in enumerate(cols):
        cx = 0.5 + i * (col_w + gap3)
        add_rounded_rect(slide, cx, CONTENT_TOP + 0.2, col_w, 5.2, fill="bgLight", border="border")
        add_rect(slide, cx, CONTENT_TOP + 0.2, col_w, 0.7, fill=color)
        add_text(slide, cx + 0.15, CONTENT_TOP + 0.28, col_w - 0.3, 0.58,
                 title, style="small", color="white", bold=True, align="center")
        for j, item in enumerate(items):
            add_text(slide, cx + 0.2, CONTENT_TOP + 1.05 + j * 0.9, col_w - 0.4, 0.8,
                     f"• {item}", style="body", word_wrap=True)

    # --- 7. 推進フロー ---
    slide = sb.add_body("推進アプローチ", section="3. 実行計画")
    add_flow_row(slide, y=CONTENT_TOP + 0.5, steps=[
        "Phase 1\n現状診断\n（4週間）",
        "Phase 2\n戦略策定\n（6週間）",
        "Phase 3\n実行計画\n（4週間）",
        "Phase 4\n実行支援\n（12週間〜）",
    ])
    add_kpi_row(slide, y=CONTENT_TOP + 2.2, items=[
        {"label": "総期間（目安）", "value": "6",   "unit": "ヶ月〜", "color": "accent"},
        {"label": "期待コスト削減",  "value": "▲20", "unit": "%",     "color": "accent2"},
        {"label": "意思決定速度向上", "value": "2×",  "unit": "以上",  "color": "accent2"},
        {"label": "投資回収期間",    "value": "18",  "unit": "ヶ月",  "color": "accent"},
    ])

    # --- 8. クロージング ---
    slide = sb._new_blank_slide()
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill="accent")
    add_rect(slide, 0, SLIDE_H - 0.25, SLIDE_W, 0.25, fill="white")
    from lib import add_line
    add_text(slide, 1.0, 2.0, SLIDE_W - 2.0, 1.2,
             "ご清聴ありがとうございました", style="title_cover", align="center", font_size=30)
    add_line(slide, 3.0, 3.4, SLIDE_W - 3.0, 3.4, color="white", width=1.0)
    add_text(slide, 1.0, 3.7, SLIDE_W - 2.0, 0.5,
             "次のステップ: 診断フェーズ開始に向けたキックオフミーティング（2週間以内）",
             style="body", color="white", align="center")
    add_text(slide, 1.0, SLIDE_H - 1.0, SLIDE_W - 2.0, 0.35,
             "XX株式会社 担当PM: 山田太郎  |  taro.yamada@example.com",
             style="caption", color="white", align="center")

    result = sb.save_and_validate(OUTPUT)
    print(f"\n生成完了: {OUTPUT}  ({len(sb.prs.slides)}枚)")


if __name__ == "__main__":
    main()
