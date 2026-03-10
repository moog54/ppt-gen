"""
戦略提言 — 経営判断支援・方向性提示（6枚）

構成: cover → consulting_scr → body_2col（論点整理）→ data_comparison → flow_timeline → closing
テーマ: mckinsey
目的: 経営層に対して戦略的方向性と実行優先度を提言する
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import (
    SlideBuilder, add_text, add_rect, add_rounded_rect,
    add_comparison, add_timeline, add_badge, add_line, add_kpi_row,
    C, CONTENT_TOP, SLIDE_W, SLIDE_H,
)

OUTPUT = "output/strategic_recommendation.pptx"


def main():
    sb = SlideBuilder(theme="mckinsey")
    sb.set_footer("DX変革戦略 提言書 — 株式会社サンプル — Strictly Confidential")

    # --- 1. 表紙 ---
    sb.add_cover(
        title="DX変革戦略\n提言書",
        subtitle="株式会社サンプル 経営会議 御中\nXX株式会社 戦略コンサルティング部",
        date="2026年4月25日",
    )

    # --- 2. 提言の核心（SCR） ---
    slide = sb.add_body("提言の核心")
    labels = [
        ("S", "Situation",    "現状",   "Phase 1診断の結果、データ分断・DX人材不足・レガシーシステムの3重苦が確認された。現状維持では2028年に市場シェア3pt喪失（売上換算▲24億円）が確実視される。", "accent2"),
        ("C", "Complication", "論点",   "「いつ・何から・どこまで」が最大の論点。投資を絞るとリターンが限定的、広げると実行リスクが高まる。スピードと確実性のトレードオフをどう解くか。", "accent3"),
        ("R", "Resolution",   "提言",   "「コアデータ基盤先行・人材育成並走」の集中戦略を採用。42億円を3フェーズで執行し、18ヶ月でROI回収・競合水準のDX成熟度を実現する。今期承認が最短での成果実現を保証する。", "accent"),
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

    # --- 3. 戦略オプション比較 ---
    slide = sb.add_body("戦略オプション比較", section="オプション分析")

    # ヘッダー
    headers = [("評価軸", 2.8), ("A案: 段階的・低リスク", 3.0), ("B案: 集中・高リターン（推奨）", 3.3), ("C案: 現状維持", 2.83)]
    hx = 0.5
    hy = CONTENT_TOP + 0.15
    for label, w in headers:
        color = "accent" if "推奨" in label else ("textMuted" if "現状" in label else "accent2")
        add_rect(slide, hx, hy, w - 0.06, 0.38, fill=color)
        add_text(slide, hx + 0.08, hy + 0.06, w - 0.2, 0.28,
                 label, style="small", color="white", bold=True)
        hx += w

    criteria = [
        ("ROI（3年）",       "120%",       "180%",       "▲3%"),
        ("投資回収期間",     "24ヶ月",     "18ヶ月",     "—"),
        ("実行リスク",       "低",         "中",         "高（機会損失）"),
        ("シェア影響",       "現状維持",   "拡大余地あり", "▲3〜5pt"),
        ("経営負担",         "小",         "中",         "小（見かけ上）"),
        ("推奨度",           "—",          "◎ 推奨",     "✕ 非推奨"),
    ]
    HIGHLIGHT = {"◎ 推奨": "accent2", "✕ 非推奨": "accent3"}
    for i, row in enumerate(criteria):
        ry = hy + 0.38 + 0.06 + i * 0.44
        bg = "rowAlt" if i % 2 == 1 else "bg"
        rx = 0.5
        for j, (cell, w) in enumerate(zip(row, [2.8, 3.0, 3.3, 2.83])):
            add_rounded_rect(slide, rx, ry, w - 0.06, 0.4, fill=bg, border="border")
            color = HIGHLIGHT.get(cell, "text") if j > 0 else "text"
            bold = cell in HIGHLIGHT
            add_text(slide, rx + 0.1, ry + 0.08, w - 0.2, 0.28,
                     cell, style="small", color=color, bold=bold, align="center" if j > 0 else "left")
            rx += w

    add_text(slide, 0.5, CONTENT_TOP + 3.25, 12.33, 0.3,
             "B案を推奨: コスト集中によるスケールメリットと、スピード優先による競合優位確保が決め手。分散投資（A案）は効果が拡散し競合との差を縮められない。",
             style="small", color="accent", bold=True)

    # --- 4. 推奨戦略の詳細 ---
    slide = sb.add_body("推奨戦略（B案）の骨子", section="提言詳細")
    pillars = [
        {
            "no": 1, "title": "コアデータ基盤の構築",
            "invest": "¥18億", "period": "〜2026年末",
            "body": "営業・製造・物流データを統合するクラウドデータ基盤を構築。週次レポートの自動化により工数を90%削減。全社一元的な意思決定基盤を確立する。",
            "kpi": "レポート工数 32h→3h / データ信頼性 100%",
            "color": "accent",
        },
        {
            "no": 2, "title": "デジタル人材の育成・採用",
            "invest": "¥12億", "period": "〜2027年中",
            "body": "データ人材を現在12名から150名体制へ。社内育成プログラム（年200名）と戦略採用（年20名）を並走。全管理職のDXリテラシー研修も実施。",
            "kpi": "DX人材比率 0.8%→3.5% / 全管理職研修完了",
            "color": "accent2",
        },
        {
            "no": 3, "title": "ガバナンス・推進体制整備",
            "invest": "¥12億", "period": "〜2027年末",
            "body": "CDO直下にDX推進室を設置。データガバナンス方針・変更管理プロセスを制定。四半期ごとのステコミで進捗・効果をモニタリングし機動的に軌道修正する。",
            "kpi": "施策実行率 >90% / 四半期レビュー体制確立",
            "color": "accent3",
        },
    ]
    col_w = 12.33 / 3
    y = CONTENT_TOP + 0.15
    for p in pillars:
        cx = 0.5 + (p["no"] - 1) * col_w
        add_rect(slide, cx, y, col_w - 0.1, 0.55, fill=p["color"])
        add_badge(slide, cx + 0.12, y + 0.08, p["no"], color="white")
        add_text(slide, cx + 0.6, y + 0.1, col_w - 0.75, 0.38,
                 p["title"], style="small", color="white", bold=True)
        ry = y + 0.55 + 0.06
        card_h = 4.55
        add_rounded_rect(slide, cx, ry, col_w - 0.1, card_h, fill="bgLight", border="border")
        add_text(slide, cx + 0.12, ry + 0.12, col_w - 0.3, 0.32,
                 f"投資額: {p['invest']}  期間: {p['period']}", style="caption", color=p["color"], bold=True)
        add_line(slide, cx + 0.1, ry + 0.5, cx + col_w - 0.2, ry + 0.5, color="border")
        add_text(slide, cx + 0.12, ry + 0.58, col_w - 0.3, 2.0,
                 p["body"], style="small", word_wrap=True)
        add_line(slide, cx + 0.1, ry + 2.7, cx + col_w - 0.2, ry + 2.7, color="border")
        add_text(slide, cx + 0.12, ry + 2.78, col_w - 0.3, 0.26, "目標KPI", style="caption", color=p["color"], bold=True)
        add_text(slide, cx + 0.12, ry + 3.05, col_w - 0.3, 1.3, p["kpi"], style="caption", word_wrap=True)

    # --- 5. 実行ロードマップ ---
    slide = sb.add_body("実行ロードマップ", section="実行計画")
    add_timeline(slide, y=CONTENT_TOP + 0.8, events=[
        {"date": "2026 Q2",  "title": "承認・着手",      "desc": "CDO体制\n確立"},
        {"date": "2026 Q3",  "title": "データ基盤\n構築開始", "desc": "クラウド\n移行着手"},
        {"date": "2026 Q4",  "title": "パイロット\n実施",   "desc": "営業部門\n先行展開"},
        {"date": "2027 Q1",  "title": "全社展開",          "desc": "全部門\n統合完了"},
        {"date": "2027 Q3",  "title": "効果確認",          "desc": "ROI測定\n追加施策"},
    ])
    add_kpi_row(slide, y=CONTENT_TOP + 3.2, items=[
        {"label": "承認判断期限",   "value": "2026/5", "unit": "末",     "color": "accent3"},
        {"label": "投資回収",       "value": "18",     "unit": "ヶ月",   "color": "accent"},
        {"label": "3年ROI",         "value": "180",    "unit": "%",      "color": "accent2"},
    ])

    # --- 6. クロージング ---
    slide = sb._new_blank_slide()
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill="accent")
    add_rect(slide, 0, SLIDE_H - 0.25, SLIDE_W, 0.25, fill="white")
    add_text(slide, 1.0, 1.6, SLIDE_W - 2.0, 1.0,
             "今期の承認が最速の成果を生む", style="title_cover", align="center", font_size=28)
    add_line(slide, 3.0, 2.9, SLIDE_W - 3.0, 2.9, color="white", width=1.0)
    add_text(slide, 1.5, 3.2, SLIDE_W - 3.0, 1.5,
             "投資の先送りは機会損失の拡大を意味します。\n"
             "2026年5月末までにご承認いただくことで、\n"
             "最短ルートでの競合キャッチアップが実現します。",
             style="body", color="white", align="center")
    add_text(slide, 1.0, SLIDE_H - 1.0, SLIDE_W - 2.0, 0.35,
             "本提言書の詳細・補足資料は担当PMまでお問い合わせください。",
             style="caption", color="white", align="center")

    result = sb.save_and_validate(OUTPUT)
    print(f"\n生成完了: {OUTPUT}  ({len(sb.prs.slides)}枚)")


if __name__ == "__main__":
    main()
