"""
パターン名: consulting_risk
カテゴリ: コンサルティング
説明: リスク・課題ログ表（影響度・対応状況付き）
用途: プロジェクト課題管理、リスク一覧の定例報告
ドキュメント: 定例報告, ステアリングコミッティ, SOW
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_rect, add_rounded_rect, C, CONTENT_TOP, SLIDE_W

STATUS_COLORS = {"対応中": "accent", "監視中": "accent2", "解決済": "textMuted", "エスカレーション": "accent3"}
IMPACT_COLORS = {"高": "accent3", "中": "accent", "低": "accent2"}


def run(
    title: str = "リスク・課題ログ",
    section: str = "",
    items: list[dict] | None = None,
    output_path: str = "output/consulting_risk.pptx",
):
    if items is None:
        items = [
            {"id": "R-01", "type": "リスク", "title": "主要メンバーのアサイン遅延",
             "impact": "高", "status": "対応中", "owner": "PM", "action": "代替メンバーのアサイン検討中。2週間以内に決定予定。"},
            {"id": "R-02", "type": "リスク", "title": "データ提供の遅れ（IT部門）",
             "impact": "高", "status": "エスカレーション", "owner": "クライアントIT", "action": "ステコミでの決定依頼済。今週中に回答予定。"},
            {"id": "I-01", "type": "課題", "title": "ヒアリング対象者のスケジュール調整難航",
             "impact": "中", "status": "対応中", "owner": "PM", "action": "一部リモート実施に変更。日程を2週間後ろ倒し。"},
            {"id": "I-02", "type": "課題", "title": "競合ベンチマークデータの入手困難",
             "impact": "中", "status": "監視中", "owner": "Sr.コンサル", "action": "公開情報・業界レポートで代替。精度への影響を注記する方針。"},
            {"id": "R-03", "type": "リスク", "title": "Phase 2 範囲拡大の可能性",
             "impact": "低", "status": "監視中", "owner": "PM", "action": "スコープ変更管理プロセスを明確化済。変更は書面合意を必須とする。"},
        ]

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    # ヘッダー
    headers = [("ID", 0.7), ("種別", 0.8), ("タイトル / 対応アクション", 5.5), ("影響", 0.7), ("状況", 1.5), ("担当", 1.4)]
    hx = 0.5
    hy = CONTENT_TOP + 0.2
    hh = 0.35
    for label, w in headers:
        add_rect(slide, hx, hy, w - 0.04, hh, fill="accent")
        add_text(slide, hx + 0.06, hy + 0.05, w - 0.14, hh - 0.1,
                 label, style="caption", color="white", bold=True, align="center")
        hx += w

    row_h = 0.9
    for i, item in enumerate(items):
        ry = hy + hh + 0.06 + i * (row_h + 0.06)
        bg = "rowAlt" if i % 2 == 1 else "bg"
        rx = 0.5

        impact_c = IMPACT_COLORS.get(item["impact"], "accent")
        status_c = STATUS_COLORS.get(item["status"], "accent")

        for label, w in [(item["id"], 0.7), (item["type"], 0.8), (None, 5.5),
                         (item["impact"], 0.7), (item["status"], 1.5), (item["owner"], 1.4)]:
            add_rounded_rect(slide, rx, ry, w - 0.04, row_h, fill=bg, border="border")
            if label is not None:
                color = "text"
                if label == item["impact"]:
                    color = impact_c
                elif label == item["status"]:
                    color = status_c
                add_text(slide, rx + 0.06, ry + 0.1, w - 0.14, row_h - 0.2,
                         label, style="small", align="center", color=color,
                         bold=(label in (item["impact"], item["status"])))
            else:
                # タイトル + アクション
                add_text(slide, rx + 0.1, ry + 0.06, 5.25, 0.35,
                         item["title"], style="small", bold=True)
                add_text(slide, rx + 0.1, ry + 0.44, 5.25, 0.4,
                         f"→ {item['action']}", style="caption", color="textLight", word_wrap=True)
            rx += w

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
