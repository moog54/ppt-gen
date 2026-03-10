"""
ステアリングコミッティ — 意思決定者向け（4枚）

構成: cover → data_kpi → consulting_risk（判断事項付き）→ flow_timeline → closing
テーマ: navy
目的: 経営層が10分で状況把握・意思決定できる資料
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import (
    SlideBuilder, add_text, add_rect, add_rounded_rect, add_kpi_row,
    add_timeline, add_badge, add_line,
    C, CONTENT_TOP, SLIDE_W, SLIDE_H,
)

OUTPUT = "output/steering_committee.pptx"


def main():
    sb = SlideBuilder(theme="navy")
    sb.set_footer("ステアリングコミッティ 第3回 — 2026年4月25日 — Confidential")

    # --- 1. 表紙 ---
    sb.add_cover(
        title="ステアリングコミッティ\n第3回",
        subtitle="DXコンサルティングプロジェクト\n2026年4月25日 / 参加者: CEO・CFO・CDO・PM",
        date="2026年4月25日",
    )

    # --- 2. プロジェクト状況サマリー ---
    slide = sb.add_body("プロジェクト状況サマリー", section="エグゼクティブサマリー")
    add_kpi_row(slide, y=CONTENT_TOP + 0.3, items=[
        {"label": "全体進捗",    "value": "68", "unit": "%",    "delta": "計画比 +3%", "color": "accent2"},
        {"label": "予算消化",    "value": "61", "unit": "%",    "delta": "計画範囲内", "color": "accent2"},
        {"label": "残期間",      "value": "27", "unit": "週",   "delta": "6ヶ月",      "color": "accent"},
        {"label": "オープン課題", "value": "2",  "unit": "件",  "delta": "要ご判断",   "color": "accent3"},
    ])
    add_text(slide, 0.5, CONTENT_TOP + 2.2, 12.33, 0.5,
             "Phase 1（現状診断）は3日前倒しで完了。Phase 2（戦略策定）は68%進捗・計画内。"
             "ただし下記2点について、本ステコミでのご判断をお願いします。",
             style="body")

    # --- 3. 判断事項 ---
    slide = sb.add_body("本日のご判断事項", section="意思決定")

    decisions = [
        {
            "no": 1, "priority": "要判断", "color": "accent3",
            "title": "ITデータ提供期限の設定",
            "background": "IT部門からのデータ提供が2週間超過。このまま放置するとPhase 2完了が7月末→8月中旬にずれ込む。",
            "options": [
                "【A案】CIOからIT部門へ期限（5/9）を指示 → 計画通り進行",
                "【B案】スコープを縮小しデータなしで戦略策定 → 精度リスクあり",
            ],
            "recommendation": "A案を推奨。CIOから直接指示いただけると即日対応可能。",
        },
        {
            "no": 2, "priority": "要確認", "color": "accent",
            "title": "Phase 2 完了後のステコミ開催日程",
            "background": "Phase 2完了（7/11予定）後、戦略オプションの最終承認のため経営層レビューが必要。",
            "options": [
                "【A案】7月14日（月）15:00〜16:30（第4回ステコミ）",
                "【B案】7月18日（金）10:00〜11:30",
            ],
            "recommendation": "いずれの日程でも対応可能。ご都合の良い日をご指定ください。",
        },
    ]

    y = CONTENT_TOP + 0.2
    for dec in decisions:
        card_h = 2.4
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

        # 背景
        add_text(slide, 0.7, y + 0.5, 11.9, 0.38, dec["background"], style="body", word_wrap=True)

        # 選択肢
        add_line(slide, 0.7, y + 0.98, 12.7, y + 0.98, color="border", width=0.5)
        for j, opt in enumerate(dec["options"]):
            add_text(slide, 0.75, y + 1.05 + j * 0.42, 11.8, 0.38, opt, style="small", word_wrap=True)

        # 推奨
        add_line(slide, 0.7, y + 1.92, 12.7, y + 1.92, color="border", width=0.5)
        add_text(slide, 0.75, y + 1.97, 11.8, 0.35,
                 f"推奨: {dec['recommendation']}", style="small", color=dec["color"], bold=True)

        y += card_h + 0.2

    # --- 4. 次回まで ---
    slide = sb.add_body("ネクストステップ・スケジュール", section="今後の予定")
    add_timeline(slide, y=CONTENT_TOP + 0.8, events=[
        {"date": "4/25",  "title": "本ステコミ",    "desc": "判断事項\n決定"},
        {"date": "5/9",   "title": "IT部門\nデータ受領", "desc": "（A案採用時）"},
        {"date": "5/15",  "title": "施策優先度\n確定",   "desc": "経営層\nレビュー"},
        {"date": "7/11",  "title": "Phase 2\n完了",      "desc": "戦略確定"},
        {"date": "7/14",  "title": "第4回\nステコミ",    "desc": "最終承認"},
    ])
    add_text(slide, 0.5, CONTENT_TOP + 3.3, 12.33, 0.35,
             "次回第4回ステコミ: 2026年7月14日（月）15:00〜16:30（予定）/ 議題: 戦略オプション最終承認",
             style="small", color="accent", bold=True)

    result = sb.save_and_validate(OUTPUT)
    print(f"\n生成完了: {OUTPUT}  ({len(sb.prs.slides)}枚)")


if __name__ == "__main__":
    main()
