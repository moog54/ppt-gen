"""
ワークショップ — DX戦略ワークショップ（5枚）

構成: cover → agenda → cover_section → workshop_exercise × 2 → closing
テーマ: green
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import (
    SlideBuilder, add_text, add_rect, add_rounded_rect, add_agenda,
    add_flow_row, add_line, add_badge,
    C, CONTENT_TOP, SLIDE_W, SLIDE_H,
)

OUTPUT = "output/workshop.pptx"


def main():
    sb = SlideBuilder(theme="green")
    sb.set_footer("DX戦略ワークショップ — 2026年4月28日")

    # --- 1. 表紙 ---
    sb.add_cover(
        title="DX戦略ワークショップ",
        subtitle="株式会社サンプル × XX株式会社\n2026年4月28日 / 参加者: 各部門リーダー14名",
        date="2026年4月28日",
    )

    # --- 2. 本日のアジェンダ ---
    slide = sb.add_body("本日のアジェンダ")
    add_agenda(slide, [
        "イントロダクション・目的共有（10分）",
        "【ワーク1】 DX阻害要因の特定（25分）",
        "グループ発表・共有（15分）",
        "【ワーク2】 優先施策の合意形成（25分）",
        "まとめ・ネクストステップ（15分）",
    ], x_start=2.5, y_start=CONTENT_TOP + 0.3, w=8.5)

    # --- 3. セクション区切り ---
    sb.add_section("ワーク 1", subtitle="DX阻害要因の特定", number=1)

    # --- 4. ワーク1指示 ---
    slide = sb.add_body("ワーク 1：グループワーク", section="ワーク1")
    y = CONTENT_TOP + 0.15

    add_rect(slide, 0.5, y, 12.33, 0.9, fill="accent")
    add_text(slide, 0.7, y + 0.08, 11.93, 0.3, "ワークテーマ", style="caption", color="white")
    add_text(slide, 0.7, y + 0.38, 11.93, 0.46,
             "自社のDX推進における最大の障壁は何か？",
             style="subheading", color="white", align="center")
    y += 1.05

    col_w = 7.5
    add_rounded_rect(slide, 0.5, y, col_w, 4.35, fill="bgLight", border="border")
    add_rect(slide, 0.5, y, col_w, 0.38, fill="accent")
    add_text(slide, 0.65, y + 0.06, col_w - 0.3, 0.28, "目的", style="caption", color="white", bold=True)
    add_text(slide, 0.65, y + 0.46, col_w - 0.3, 0.5,
             "現場の視点から課題を特定し、優先度の高い障壁を3つ導出する。",
             style="body", word_wrap=True)
    add_line(slide, 0.65, y + 1.08, 0.5 + col_w - 0.2, y + 1.08, color="border")
    add_text(slide, 0.65, y + 1.15, col_w - 0.3, 0.3, "進め方", style="small", bold=True, color="accent")
    for j, step in enumerate([
        "個人で課題を付箋に書き出す（5分）",
        "グループ内で共有・グルーピングする（10分）",
        "優先度トップ3を選定して発表準備（10分）",
    ]):
        sy = y + 1.5 + j * 0.85
        add_badge(slide, 0.65, sy, j + 1, color="accent")
        add_text(slide, 1.25, sy, col_w - 0.9, 0.75, step, style="body", word_wrap=True)

    rx = 0.5 + col_w + 0.2
    rw = 12.33 - col_w - 0.2
    add_rounded_rect(slide, rx, y, rw, 2.1, fill="accent", border=None)
    add_text(slide, rx + 0.15, y + 0.15, rw - 0.3, 0.35, "制限時間", style="caption", color="white")
    add_text(slide, rx + 0.15, y + 0.5, rw - 0.3, 1.1, "25", style="kpi", color="white", align="center", font_size=64)
    add_text(slide, rx + 0.15, y + 1.65, rw - 0.3, 0.35, "分", style="subheading", color="white", align="center")
    add_rounded_rect(slide, rx, y + 2.25, rw, 2.1, fill="bgLight", border="border")
    add_rect(slide, rx, y + 2.25, rw, 0.38, fill="accent2")
    add_text(slide, rx + 0.15, y + 2.31, rw - 0.3, 0.28, "グループ構成", style="caption", color="white", bold=True)
    add_text(slide, rx + 0.15, y + 2.72, rw - 0.3, 0.8, "4", style="kpi", color="accent2", align="center", font_size=48)
    add_text(slide, rx + 0.15, y + 3.55, rw - 0.3, 0.3, "名 / グループ", style="small", color="textLight", align="center")

    # --- 5. セクション区切り ---
    sb.add_section("ワーク 2", subtitle="優先施策の合意形成", number=2)

    # --- 6. ワーク2指示 ---
    slide = sb.add_body("ワーク 2：グループワーク", section="ワーク2")
    y = CONTENT_TOP + 0.15

    add_rect(slide, 0.5, y, 12.33, 0.9, fill="accent2")
    add_text(slide, 0.7, y + 0.08, 11.93, 0.3, "ワークテーマ", style="caption", color="white")
    add_text(slide, 0.7, y + 0.38, 11.93, 0.46,
             "ワーク1の阻害要因を踏まえ、来年度に取り組むべき施策TOP3を選べ",
             style="subheading", color="white", align="center")
    y += 1.05

    add_rounded_rect(slide, 0.5, y, col_w, 4.35, fill="bgLight", border="border")
    add_rect(slide, 0.5, y, col_w, 0.38, fill="accent2")
    add_text(slide, 0.65, y + 0.06, col_w - 0.3, 0.28, "目的", style="caption", color="white", bold=True)
    add_text(slide, 0.65, y + 0.46, col_w - 0.3, 0.5,
             "全社で合意できる優先施策を絞り込み、次のアクションを具体化する。",
             style="body", word_wrap=True)
    add_line(slide, 0.65, y + 1.08, 0.5 + col_w - 0.2, y + 1.08, color="border")
    add_text(slide, 0.65, y + 1.15, col_w - 0.3, 0.3, "進め方", style="small", bold=True, color="accent2")
    for j, step in enumerate([
        "ワーク1の結果を見ながら施策を発散（5分）",
        "評価軸（効果 × 実現性）で施策をマッピング（10分）",
        "TOP3に絞り込み、担当・期限を設定（10分）",
    ]):
        sy = y + 1.5 + j * 0.85
        add_badge(slide, 0.65, sy, j + 1, color="accent2")
        add_text(slide, 1.25, sy, col_w - 0.9, 0.75, step, style="body", word_wrap=True)

    add_rounded_rect(slide, rx, y, rw, 2.1, fill="accent2", border=None)
    add_text(slide, rx + 0.15, y + 0.15, rw - 0.3, 0.35, "制限時間", style="caption", color="white")
    add_text(slide, rx + 0.15, y + 0.5, rw - 0.3, 1.1, "25", style="kpi", color="white", align="center", font_size=64)
    add_text(slide, rx + 0.15, y + 1.65, rw - 0.3, 0.35, "分", style="subheading", color="white", align="center")
    add_rounded_rect(slide, rx, y + 2.25, rw, 2.1, fill="bgLight", border="border")
    add_rect(slide, rx, y + 2.25, rw, 0.38, fill="accent")
    add_text(slide, rx + 0.15, y + 2.31, rw - 0.3, 0.28, "グループ構成", style="caption", color="white", bold=True)
    add_text(slide, rx + 0.15, y + 2.72, rw - 0.3, 0.8, "全体", style="kpi", color="accent", align="center", font_size=36)
    add_text(slide, rx + 0.15, y + 3.55, rw - 0.3, 0.3, "（合同ワーク）", style="small", color="textLight", align="center")

    # --- 7. クロージング ---
    slide = sb._new_blank_slide()
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill="accent")
    add_rect(slide, 0, SLIDE_H - 0.25, SLIDE_W, 0.25, fill="white")
    add_text(slide, 1.0, 2.0, SLIDE_W - 2.0, 1.0,
             "お疲れ様でした", style="title_cover", align="center", font_size=32)
    add_line(slide, 3.0, 3.2, SLIDE_W - 3.0, 3.2, color="white", width=1.0)
    add_text(slide, 1.0, 3.5, SLIDE_W - 2.0, 1.0,
             "今日のアウトプットをもとに、\n2週間以内に優先施策の実行計画を策定します。",
             style="body", color="white", align="center")
    add_text(slide, 1.0, SLIDE_H - 1.0, SLIDE_W - 2.0, 0.35,
             "議事録・アウトプット共有: 5営業日以内にXX株式会社より送付",
             style="caption", color="white", align="center")

    result = sb.save_and_validate(OUTPUT)
    print(f"\n生成完了: {OUTPUT}  ({len(sb.prs.slides)}枚)")


if __name__ == "__main__":
    main()
