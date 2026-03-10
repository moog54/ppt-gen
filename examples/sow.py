"""
SOW — Statement of Work（5枚）

構成: cover → body_1col（概要） → consulting_scope → consulting_pricing → flow_timeline → closing
テーマ: navy
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import (
    SlideBuilder, add_text, add_rect, add_rounded_rect, add_timeline,
    add_flow_row, add_line,
    C, CONTENT_TOP, SLIDE_W, SLIDE_H,
)

OUTPUT = "output/sow.pptx"


def main():
    sb = SlideBuilder(theme="navy")
    sb.set_footer("業務委託契約書 別紙 — Statement of Work v1.0")

    # --- 1. 表紙 ---
    sb.add_cover(
        title="業務委託契約書\nStatement of Work",
        subtitle="株式会社サンプル × XX株式会社\nDXコンサルティング業務",
        date="2026年4月1日",
    )

    # --- 2. プロジェクト概要 ---
    slide = sb.add_body("プロジェクト概要")
    items = [
        ("委託者",     "株式会社サンプル（以下「クライアント」）"),
        ("受託者",     "XX株式会社（以下「コンサルタント」）"),
        ("業務名称",   "デジタル変革推進コンサルティング業務"),
        ("契約期間",   "2026年5月1日 〜 2026年10月31日（6ヶ月間）"),
        ("契約金額",   "金42,000,000円（税別）/ 別途実費精算"),
        ("主管部門",   "クライアント: 経営企画部 / コンサルタント: 戦略コンサルティング部"),
        ("レポートライン", "月次ステアリングコミッティ（クライアントCEO・CFO・CDO出席）"),
    ]
    y = CONTENT_TOP + 0.2
    for label, value in items:
        add_rounded_rect(slide, 0.5, y, 12.33, 0.48, fill="bgLight", border="border")
        add_rect(slide, 0.5, y, 2.5, 0.48, fill="accent")
        add_text(slide, 0.65, y + 0.08, 2.2, 0.32, label, style="small", color="white", bold=True)
        add_text(slide, 3.15, y + 0.08, 9.5, 0.32, value, style="small")
        y += 0.55

    # --- 3. スコープ・デリバラブル ---
    slide = sb.add_body("スコープ・デリバラブル定義", section="業務範囲")
    from patterns.consulting_scope import run as _s
    # インラインで描画
    phases = [
        {"phase": "Phase 1\n現状診断", "duration": "4週間",
         "deliverables": ["現状分析レポート", "課題一覧（優先度付き）", "ベンチマーク調査結果"],
         "in_scope": ["経営層・現場ヒアリング（30名）", "データ収集・統計分析", "競合ベンチマーク"],
         "out_scope": []},
        {"phase": "Phase 2\n戦略策定", "duration": "6週間",
         "deliverables": ["変革ロードマップ", "施策優先度マトリクス", "ROI試算モデル"],
         "in_scope": ["戦略オプション立案", "ステークホルダー合意", "KPI設計"],
         "out_scope": []},
        {"phase": "Phase 3\n実行計画", "duration": "4週間",
         "deliverables": ["詳細実行計画書", "ガバナンス設計書", "クイックウィン特定"],
         "in_scope": ["実行計画詳細化", "推進体制設計", "リスク対応計画"],
         "out_scope": []},
    ]
    n = len(phases)
    col_w = 12.33 / n
    y = CONTENT_TOP + 0.2
    header_h, row_h = 0.65, 1.35
    for i, ph in enumerate(phases):
        cx = 0.5 + i * col_w
        color = ["accent", "accent2", "accent3"][i % 3]
        add_rect(slide, cx, y, col_w - 0.08, header_h, fill=color)
        add_text(slide, cx + 0.1, y + 0.06, col_w - 0.28, 0.38,
                 ph["phase"], style="small", color="white", bold=True)
        add_text(slide, cx + 0.1, y + 0.44, col_w - 0.28, 0.2,
                 f"期間: {ph['duration']}", style="caption", color="white")
        ry = y + header_h + 0.08
        add_rounded_rect(slide, cx, ry, col_w - 0.08, row_h, fill="bgLight", border="border")
        add_rect(slide, cx, ry, col_w - 0.08, 0.26, fill=color)
        add_text(slide, cx + 0.08, ry + 0.04, col_w - 0.2, 0.2,
                 "成果物", style="caption", color="white", bold=True)
        for j, d in enumerate(ph["deliverables"]):
            add_text(slide, cx + 0.12, ry + 0.3 + j * 0.33, col_w - 0.24, 0.31,
                     f"• {d}", style="small", word_wrap=True)
        ry2 = ry + row_h + 0.08
        add_rounded_rect(slide, cx, ry2, col_w - 0.08, row_h, fill="bg", border="accent2")
        add_rect(slide, cx, ry2, col_w - 0.08, 0.26, fill="accent2")
        add_text(slide, cx + 0.08, ry2 + 0.04, col_w - 0.2, 0.2,
                 "対象範囲", style="caption", color="white", bold=True)
        for j, s in enumerate(ph["in_scope"]):
            add_text(slide, cx + 0.12, ry2 + 0.3 + j * 0.33, col_w - 0.24, 0.31,
                     f"✓ {s}", style="small", color="accent2", word_wrap=True)

    # --- 4. 費用内訳 ---
    slide = sb.add_body("投資額・工数内訳", section="費用")
    from lib import add_table_shapes
    headers = ["フェーズ / 項目", "期間", "体制", "工数", "費用（税別）"]
    rows = [
        ["Phase 1 現状診断",    "4週間",  "PM×0.5 + Sr.×1 + Jr.×2", "80人日",  "¥8,000,000"],
        ["Phase 2 戦略策定",    "6週間",  "PM×0.5 + Sr.×2 + Jr.×2", "130人日", "¥16,000,000"],
        ["Phase 3 実行計画",    "4週間",  "PM×0.5 + Sr.×1 + Jr.×2", "80人日",  "¥8,000,000"],
        ["プロジェクト管理",    "全期間", "PM×0.3",                   "30人日",  "¥6,000,000"],
        ["予備費（10%）",       "—",      "—",                        "—",       "¥4,000,000"],
    ]
    col_widths = [3.2, 1.2, 3.7, 1.4, 1.83]
    add_table_shapes(slide, x=0.5, y=CONTENT_TOP + 0.2, w=11.33,
                     headers=headers, rows=rows, col_widths=col_widths)
    total_y = CONTENT_TOP + 0.2 + 0.38 * (len(rows) + 1) + 0.1
    add_rect(slide, 0.5, total_y, 11.33, 0.46, fill="accent")
    add_text(slide, 0.6, total_y + 0.06, 6.0, 0.34,
             "合計（税別）", style="small", color="white", bold=True)
    add_text(slide, 7.5, total_y + 0.06, 4.3, 0.34,
             "¥42,000,000", style="subheading", color="white", align="right")
    add_text(slide, 0.5, total_y + 0.58, 12.33, 0.28,
             "※ 費用は税別。交通費・宿泊費等の実費は別途精算。支払条件: 月末締め翌月末払い。",
             style="caption", color="textMuted")

    # --- 5. スケジュール ---
    slide = sb.add_body("プロジェクトスケジュール", section="実行計画")
    add_timeline(slide, y=CONTENT_TOP + 1.0, events=[
        {"date": "2026年5月",  "title": "キックオフ",    "desc": "体制確立\nヒアリング開始"},
        {"date": "2026年6月",  "title": "Phase 1 完了",  "desc": "現状診断報告\n中間ステコミ"},
        {"date": "2026年7月",  "title": "戦略策定",      "desc": "オプション検討\n経営層レビュー"},
        {"date": "2026年8月",  "title": "Phase 2 完了",  "desc": "ロードマップ確定\n最終ステコミ"},
        {"date": "2026年10月", "title": "Phase 3 完了",  "desc": "実行計画書納品\n成果報告会"},
    ])
    add_text(slide, 0.5, CONTENT_TOP + 3.4, 12.33, 0.35,
             "変更管理: スコープ変更は書面合意を必須とし、追加費用が発生する場合は事前承認を得る。",
             style="caption", color="textMuted")

    # --- 6. クロージング ---
    slide = sb._new_blank_slide()
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill="accent")
    add_rect(slide, 0, SLIDE_H - 0.25, SLIDE_W, 0.25, fill="white")
    add_text(slide, 1.0, 2.2, SLIDE_W - 2.0, 0.9,
             "本SOWは両社の合意に基づき効力を生ずる", style="title_cover", align="center", font_size=26)
    add_line(slide, 2.0, 3.3, SLIDE_W - 2.0, 3.3, color="white", width=0.8)
    for i, line in enumerate(["委託者署名欄: ___________________", "受託者署名欄: ___________________", "合意日: 2026年　　月　　日"]):
        add_text(slide, 2.5, 3.6 + i * 0.7, SLIDE_W - 5.0, 0.55,
                 line, style="body", color="white", align="center")

    result = sb.save_and_validate(OUTPUT)
    print(f"\n生成完了: {OUTPUT}  ({len(sb.prs.slides)}枚)")


if __name__ == "__main__":
    main()
