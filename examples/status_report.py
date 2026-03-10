"""
定例報告 — プロジェクト月次ステータスレポート（5枚）

構成: cover → data_kpi → data_table → consulting_risk → flow_timeline → closing
テーマ: default
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import (
    SlideBuilder, add_text, add_rect, add_rounded_rect, add_kpi_row,
    add_table_shapes, add_timeline, add_line,
    C, CONTENT_TOP, SLIDE_W, SLIDE_H,
)

OUTPUT = "output/status_report.pptx"
REPORT_DATE = "2026年4月25日（第3回）"


def main():
    sb = SlideBuilder(theme="default")
    sb.set_footer(f"DXコンサルティング 月次定例報告 — {REPORT_DATE}")

    # --- 1. 表紙 ---
    sb.add_cover(
        title="月次プロジェクト\nステータスレポート",
        subtitle=f"株式会社サンプル × XX株式会社\n{REPORT_DATE}",
        date="2026年4月",
    )

    # --- 2. KPIサマリー ---
    slide = sb.add_body("プロジェクトKPIサマリー", section="エグゼクティブサマリー")
    add_kpi_row(slide, y=CONTENT_TOP + 0.3, items=[
        {"label": "全体進捗",         "value": "68",  "unit": "%",    "delta": "計画比 +3%", "color": "accent2"},
        {"label": "予算消化率",        "value": "61",  "unit": "%",    "delta": "計画比 ▲2%", "color": "accent2"},
        {"label": "完了タスク",        "value": "34",  "unit": "/ 50", "delta": "+8（今月）", "color": "accent"},
        {"label": "オープンリスク数",  "value": "3",   "unit": "件",   "delta": "前月比 ▲1",  "color": "accent2"},
    ])
    add_text(slide, 0.5, CONTENT_TOP + 2.2, 12.33, 0.5,
             "全体は計画比3%前倒しで進行中。Phase 2（戦略策定）は来月完了見込み。"
             "主要リスク：データ提供遅延（IT部門）はステコミでの対処が必要。",
             style="body", color="accent")

    # --- 3. フェーズ別進捗 ---
    slide = sb.add_body("フェーズ別進捗", section="詳細報告")
    headers = ["フェーズ", "計画完了", "実績", "進捗率", "ステータス", "備考"]
    rows = [
        ["Phase 1 現状診断",  "2026/5/30",  "2026/5/27 ✓", "100%", "完了",  "3日前倒しで完了"],
        ["Phase 2 戦略策定",  "2026/7/11",  "進行中",      "68%",  "順調",  "主要施策の方向性合意済"],
        ["Phase 3 実行計画",  "2026/8/8",   "未着手",      "—",    "未開始", "Phase 2完了後に着手"],
        ["全体",              "2026/10/31", "—",           "68%",  "計画内", "前倒しトレンド継続"],
    ]
    col_widths = [2.5, 1.7, 1.8, 1.3, 1.3, 3.0]
    add_table_shapes(slide, x=0.5, y=CONTENT_TOP + 0.2, w=11.6,
                     headers=headers, rows=rows, col_widths=col_widths)

    add_text(slide, 0.5, CONTENT_TOP + 2.6, 12.33, 0.3,
             "今月の主要成果物: ① 戦略オプション比較レポート（v0.8）  ② KPI設計案（ドラフト）  ③ ITアーキテクチャ現状図",
             style="small", color="accent")

    # --- 4. リスク・課題 ---
    slide = sb.add_body("リスク・課題ログ", section="リスク管理")
    from lib import add_badge
    items = [
        {"id": "R-02", "type": "リスク", "title": "データ提供の遅れ（IT部門）",
         "impact": "高", "status": "エスカレーション", "owner": "クライアントIT",
         "action": "ステコミでの決定依頼済。今週中に回答予定。"},
        {"id": "I-01", "type": "課題", "title": "ヒアリング対象者のスケジュール調整難航",
         "impact": "中", "status": "対応中", "owner": "PM",
         "action": "一部リモートに変更。日程を2週間後ろ倒し。"},
        {"id": "R-03", "type": "リスク", "title": "Phase 2 範囲拡大の可能性",
         "impact": "低", "status": "監視中", "owner": "PM",
         "action": "スコープ変更管理プロセスを明確化済。"},
    ]
    STATUS_COLORS = {"対応中": "accent", "監視中": "accent2", "解決済": "textMuted", "エスカレーション": "accent3"}
    IMPACT_COLORS = {"高": "accent3", "中": "accent", "低": "accent2"}
    headers2 = [("ID", 0.7), ("種別", 0.8), ("タイトル / 対応アクション", 5.5), ("影響", 0.7), ("状況", 1.5), ("担当", 1.4)]
    hx = 0.5
    hy = CONTENT_TOP + 0.2
    for label, w in headers2:
        add_rect(slide, hx, hy, w - 0.04, 0.35, fill="accent")
        add_text(slide, hx + 0.06, hy + 0.05, w - 0.14, 0.25,
                 label, style="caption", color="white", bold=True, align="center")
        hx += w
    row_h = 0.95
    for i, item in enumerate(items):
        ry = hy + 0.35 + 0.06 + i * (row_h + 0.06)
        bg = "rowAlt" if i % 2 == 1 else "bg"
        rx = 0.5
        for label2, w in [(item["id"], 0.7), (item["type"], 0.8), (None, 5.5),
                          (item["impact"], 0.7), (item["status"], 1.5), (item["owner"], 1.4)]:
            add_rounded_rect(slide, rx, ry, w - 0.04, row_h, fill=bg, border="border")
            if label2 is not None:
                color = "text"
                if label2 == item["impact"]:
                    color = IMPACT_COLORS.get(item["impact"], "accent")
                elif label2 == item["status"]:
                    color = STATUS_COLORS.get(item["status"], "accent")
                add_text(slide, rx + 0.06, ry + 0.12, w - 0.14, row_h - 0.24,
                         label2, style="small", align="center", color=color,
                         bold=(label2 in (item["impact"], item["status"])))
            else:
                add_text(slide, rx + 0.1, ry + 0.06, 5.25, 0.38, item["title"], style="small", bold=True)
                add_text(slide, rx + 0.1, ry + 0.48, 5.25, 0.4,
                         f"→ {item['action']}", style="caption", color="textLight", word_wrap=True)
            rx += w

    # --- 5. 来月の計画 ---
    slide = sb.add_body("来月の予定・マイルストーン", section="次ステップ")
    add_timeline(slide, y=CONTENT_TOP + 0.8, events=[
        {"date": "5/1",  "title": "IT部門\nデータ受領", "desc": "ステコミ\n決定次第"},
        {"date": "5/15", "title": "施策優先度\n確定",   "desc": "経営層\nレビュー"},
        {"date": "5/22", "title": "KPI設計\n最終化",    "desc": "クライアント\n承認"},
        {"date": "5/30", "title": "Phase 2\n中間報告",  "desc": "定例報告会"},
    ])
    add_text(slide, 0.5, CONTENT_TOP + 3.0, 12.33, 0.35,
             "【判断依頼事項】 IT部門のデータ提供期限について、今週中のご回答をお願いします。"
             "遅延の場合、Phase 2完了が最大2週間後ろ倒しになる見込みです。",
             style="body", color="accent3")

    result = sb.save_and_validate(OUTPUT)
    print(f"\n生成完了: {OUTPUT}  ({len(sb.prs.slides)}枚)")


if __name__ == "__main__":
    main()
