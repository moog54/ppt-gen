"""
エグゼクティブブリーフィング — 経営層向け短時間報告（4枚）

構成: cover → data_kpi（状況一覧）→ consulting_risk（判断事項）→ closing
テーマ: navy
目的: 多忙な経営層が5〜10分で状況把握・意思決定できる資料
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import (
    SlideBuilder, add_text, add_rect, add_rounded_rect,
    add_kpi_row, add_badge, add_line,
    C, CONTENT_TOP, SLIDE_W, SLIDE_H,
)

OUTPUT = "output/executive_briefing.pptx"


def main():
    sb = SlideBuilder(theme="navy")
    sb.set_footer("エグゼクティブブリーフィング — 2026年4月25日 — Confidential")

    # --- 1. 表紙 ---
    sb.add_cover(
        title="エグゼクティブ\nブリーフィング",
        subtitle="DXコンサルティングプロジェクト 緊急報告\n2026年4月25日 / 所要時間: 約10分",
        date="2026年4月25日",
    )

    # --- 2. 現状サマリー ---
    slide = sb.add_body("現状サマリー — 3点で把握")

    add_kpi_row(slide, y=CONTENT_TOP + 0.3, items=[
        {"label": "全体進捗",    "value": "68", "unit": "%",   "delta": "計画比 +3%",  "color": "accent2"},
        {"label": "予算消化",    "value": "61", "unit": "%",   "delta": "計画範囲内",  "color": "accent2"},
        {"label": "残期間",      "value": "27", "unit": "週",  "delta": "〜10月末",    "color": "accent"},
        {"label": "要判断事項",  "value": "2",  "unit": "件",  "delta": "本日決定希望", "color": "accent3"},
    ])

    # 3ポイントサマリー
    points = [
        {
            "icon": "✓", "color": "accent2", "title": "良いニュース",
            "body": "Phase 1（現状診断）が3日前倒しで完了。Phase 2は計画比3%超過と堅調に進行中。"
                    "戦略オプションの方向性についても主要ステークホルダーの合意を得た。",
        },
        {
            "icon": "!", "color": "accent3", "title": "懸念事項",
            "body": "IT部門からのデータ提供が2週間超過。このまま放置するとPhase 2完了が"
                    "7月末→8月中旬にずれ込み、最終納品も遅延する見込み。",
        },
        {
            "icon": "→", "color": "accent", "title": "本日のお願い",
            "body": "2件の判断事項（次ページ）について、本ブリーフィング中にご決定いただけますと、"
                    "プロジェクトを計画通りに進行できます。",
        },
    ]
    y = CONTENT_TOP + 2.15
    for p in points:
        add_rounded_rect(slide, 0.5, y, 12.33, 1.05, fill="bgLight", border="border")
        add_rect(slide, 0.5, y, 0.07, 1.05, fill=p["color"])
        add_rounded_rect(slide, 0.65, y + 0.28, 0.45, 0.45, fill=p["color"], border=None)
        add_text(slide, 0.65, y + 0.28, 0.45, 0.45, p["icon"], style="small", color="white",
                 bold=True, align="center")
        add_text(slide, 1.25, y + 0.1, 2.5, 0.32, p["title"], style="small", color=p["color"], bold=True)
        add_text(slide, 1.25, y + 0.42, 11.35, 0.55, p["body"], style="body", word_wrap=True)
        y += 1.15

    # --- 3. 判断事項 ---
    slide = sb.add_body("本日の判断事項 — 2件", section="意思決定")

    decisions = [
        {
            "no": 1, "priority": "要判断", "color": "accent3",
            "title": "ITデータ提供期限の設定（緊急）",
            "situation": "IT部門からのデータ提供が2週間超過。指示系統が曖昧なため担当者が動けない状況。",
            "impact": "放置の場合: Phase 2完了が7/11→8/18にずれ込み、最終納品も6週間遅延。",
            "options": [
                "【A案】本日中にCIOからIT部門へ期限（5/9）を直接指示 → 計画通り進行",
                "【B案】スコープを縮小しデータなしで戦略策定 → 完成度・精度が低下",
            ],
            "rec": "A案。CIOから直接指示いただければ即日対応可能。5分で解決できます。",
        },
        {
            "no": 2, "priority": "要確認", "color": "accent",
            "title": "第4回ステコミの開催日程",
            "situation": "Phase 2完了（7/11予定）後、戦略オプション最終承認のための経営層レビューが必要。",
            "impact": "日程未確定のまま進むと、承認プロセスが遅延し実行開始が後ろ倒しになる。",
            "options": [
                "【A案】7月14日（月）15:00〜16:30（第4回ステコミ）",
                "【B案】7月18日（金）10:00〜11:30",
            ],
            "rec": "いずれも対応可能。本日中にご指定いただけると日程調整を即開始できます。",
        },
    ]

    y = CONTENT_TOP + 0.15
    for dec in decisions:
        card_h = 2.55
        add_rounded_rect(slide, 0.5, y, 12.33, card_h, fill="bgLight", border="border")
        add_rect(slide, 0.5, y, 0.07, card_h, fill=dec["color"])

        # ヘッダー
        add_rect(slide, 0.5, y, 12.33, 0.42, fill=dec["color"])
        add_badge(slide, 0.65, y + 0.0, dec["no"], color="white")
        add_rounded_rect(slide, 1.3, y + 0.06, 1.1, 0.3, fill="white")
        add_text(slide, 1.32, y + 0.08, 1.06, 0.26,
                 dec["priority"], style="caption", color=dec["color"], bold=True, align="center")
        add_text(slide, 2.55, y + 0.06, 9.8, 0.3,
                 dec["title"], style="small", color="white", bold=True)

        # 状況・影響
        add_text(slide, 0.7, y + 0.5, 5.8, 0.38, dec["situation"], style="body", word_wrap=True)
        add_text(slide, 0.7, y + 0.9, 5.8, 0.38,
                 dec["impact"], style="small", color="accent3", word_wrap=True)

        # 選択肢
        add_line(slide, 6.65, y + 0.48, 6.65, y + card_h - 0.12, color="border", width=0.5)
        add_text(slide, 6.75, y + 0.5, 5.8, 0.22, "選択肢", style="caption", color="textMuted")
        for j, opt in enumerate(dec["options"]):
            add_text(slide, 6.75, y + 0.76 + j * 0.45, 5.8, 0.4, opt, style="small", word_wrap=True)

        # 推奨
        add_line(slide, 0.7, y + 1.95, 12.7, y + 1.95, color="border", width=0.5)
        add_text(slide, 0.75, y + 2.02, 11.8, 0.42,
                 f"推奨: {dec['rec']}", style="small", color=dec["color"], bold=True, word_wrap=True)

        y += card_h + 0.2

    # --- 4. クロージング ---
    slide = sb._new_blank_slide()
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill="accent")
    add_rect(slide, 0, SLIDE_H - 0.25, SLIDE_W, 0.25, fill="white")
    add_text(slide, 1.0, 1.8, SLIDE_W - 2.0, 0.8,
             "ご確認ありがとうございました", style="title_cover", align="center", font_size=28)
    add_line(slide, 3.0, 3.0, SLIDE_W - 3.0, 3.0, color="white", width=1.0)
    add_text(slide, 1.5, 3.3, SLIDE_W - 3.0, 1.2,
             "本日ご決定いただいた内容は、翌営業日中に議事録として共有いたします。\n"
             "次回定例報告: 2026年5月末（月次ステータスレポート）\n"
             "緊急連絡先: 担当PM（XX株式会社）xxx-xxxx-xxxx",
             style="body", color="white", align="center")

    result = sb.save_and_validate(OUTPUT)
    print(f"\n生成完了: {OUTPUT}  ({len(sb.prs.slides)}枚)")


if __name__ == "__main__":
    main()
