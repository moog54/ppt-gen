"""
パターン名: data_comparison_matrix
カテゴリ: データ
説明: 階層型比較マトリクス表（左側グループ列スパン結合＋サブグループ＋2列箇条書き比較）
用途: 機能比較、As-Is/To-Be、Before/After、オプション評価
ドキュメント: 提案書, SOW, フィンディングス, 戦略提言, エグゼクティブブリーフィング
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_rect, add_rounded_rect, add_line, CONTENT_TOP


# ── デフォルトパレット（テーマ別に差し替え可） ─────────────────────
DEFAULT_PAL = {
    "grp_bg":     "1E6B3C",
    "grp_text":   "FFFFFF",
    "sub_bg":     "2E8B57",
    "sub_text":   "FFFFFF",
    "hdr_a_bg":   "ECECEC",
    "hdr_b_bg":   "D4EDDA",
    "hdr_text":   "1A1A1A",
    "row_a_even": "FFFFFF",
    "row_a_odd":  "F9F9F9",
    "row_b_even": "F0FBF4",
    "row_b_odd":  "E8F5EE",
    "border":     "CCCCCC",
    "outer":      "888888",
}

THEME_PAL = {
    "accenture": {**DEFAULT_PAL,
                  "grp_bg":"5B0099","sub_bg":"A100FF",
                  "hdr_b_bg":"EDD9FF","row_b_even":"FAF5FF","row_b_odd":"F5EEFF"},
    "mckinsey":  {**DEFAULT_PAL,
                  "grp_bg":"002F6C","sub_bg":"1455A0",
                  "hdr_b_bg":"D6E4F0","row_b_even":"F0F5FA","row_b_odd":"E6EFF7"},
    "navy":      {**DEFAULT_PAL,
                  "grp_bg":"1A3C6E","sub_bg":"2B5DA7",
                  "hdr_b_bg":"D4E2F0","row_b_even":"F0F4F8","row_b_odd":"E5EDF5"},
    "default":   {**DEFAULT_PAL,
                  "grp_bg":"004B8D","sub_bg":"0070C0",
                  "hdr_b_bg":"D4EBFA","row_b_even":"F0F7FC","row_b_odd":"E5F2FA"},
}


def _calc_groups(rows):
    """連続する同名グループをスパン結合: [(grp_name, start_idx, end_idx), ...]"""
    result, i = [], 0
    while i < len(rows):
        g = rows[i][0]
        if g is None:
            result.append((None, i, i))
            i += 1
        else:
            j = i
            while j < len(rows) and rows[j][0] == g:
                j += 1
            result.append((g, i, j - 1))
            i = j
    return result


def _bullets(items):
    return "\n".join(f"• {s}" for s in items) if items else ""


def render_comparison_matrix(
    slide,
    x: float,
    y: float,
    w: float,
    col_a_label: str,
    col_b_label: str,
    rows: list,
    pal: dict,
    group_w: float = 0.82,
    sub_w: float = 1.72,
    hdr_h: float = 0.38,
    line_h: float = 0.195,
    pad_h: float = 0.22,
    min_row_h: float = 0.50,
    font_size: int = 10,
):
    """
    rows: [(group_or_None, subgroup, col_a_bullets: list[str], col_b_bullets: list[str]), ...]
    """
    cw = (w - group_w - sub_w) / 2   # content column width

    def row_h(ba, bb):
        return max(min_row_h, max(len(ba), len(bb), 1) * line_h + pad_h)

    row_hs = [row_h(r[2], r[3]) for r in rows]
    groups = _calc_groups(rows)

    # ── ヘッダー ──
    add_rect(slide, x, y, group_w + sub_w, hdr_h,
             fill=pal["sub_bg"], border=pal["border"])
    add_text(slide, x, y, group_w + sub_w, hdr_h, "分類",
             style="small", align="center", bold=True, color=pal["sub_text"])
    add_rect(slide, x + group_w + sub_w, y, cw, hdr_h,
             fill=pal["hdr_a_bg"], border=pal["border"])
    add_text(slide, x + group_w + sub_w, y, cw, hdr_h, col_a_label,
             style="small", align="center", bold=True, color=pal["hdr_text"])
    add_rect(slide, x + group_w + sub_w + cw, y, cw, hdr_h,
             fill=pal["hdr_b_bg"], border=pal["border"])
    add_text(slide, x + group_w + sub_w + cw, y, cw, hdr_h, col_b_label,
             style="small", align="center", bold=True, color=pal["grp_bg"])

    # ── データ行 ──
    cy = y + hdr_h
    for (grp_name, g_start, g_end) in groups:
        grp_total_h = sum(row_hs[g_start:g_end + 1])

        if grp_name is not None:
            add_rect(slide, x, cy, group_w, grp_total_h,
                     fill=pal["grp_bg"], border=pal["border"])
            add_text(slide, x, cy, group_w, grp_total_h, grp_name,
                     style="small", align="center", bold=True,
                     color=pal["grp_text"], font_size=font_size)

        ry = cy
        for ri in range(g_start, g_end + 1):
            _, sub, ba, bb = rows[ri]
            rh = row_hs[ri]
            is_even = ri % 2 == 0

            sub_x = x + (group_w if grp_name else 0)
            sw    = sub_w + (0 if grp_name else group_w)
            add_rect(slide, sub_x, ry, sw, rh, fill=pal["sub_bg"], border=pal["border"])
            add_text(slide, sub_x, ry, sw, rh, sub,
                     style="small", align="center", bold=True,
                     color=pal["sub_text"], font_size=font_size)

            bg_a = pal["row_a_even"] if is_even else pal["row_a_odd"]
            cx_a = x + group_w + sub_w
            add_rect(slide, cx_a, ry, cw, rh, fill=bg_a, border=pal["border"])
            if ba:
                add_text(slide, cx_a + 0.10, ry + 0.06, cw - 0.16, rh - 0.10,
                         _bullets(ba), style="caption", align="left",
                         color="1A1A1A", font_size=font_size)

            bg_b = pal["row_b_even"] if is_even else pal["row_b_odd"]
            cx_b = x + group_w + sub_w + cw
            add_rect(slide, cx_b, ry, cw, rh, fill=bg_b, border=pal["border"])
            if bb:
                add_text(slide, cx_b + 0.10, ry + 0.06, cw - 0.16, rh - 0.10,
                         _bullets(bb), style="caption", align="left",
                         color="1A1A1A", font_size=font_size)

            ry += rh
        cy += grp_total_h

    # 外枠
    total_h = sum(row_hs) + hdr_h
    add_rect(slide, x, y, w, total_h, fill=None, border=pal["outer"], border_width=1.5)


def run(
    title: str = "機能比較：As-Is / To-Be",
    section: str = "現状分析",
    col_a_label: str = "現行（As-Is）",
    col_b_label: str = "将来（To-Be）",
    theme: str = "default",
    rows: list | None = None,
    output_path: str = "output/data_comparison_matrix.pptx",
):
    """
    rows: [(group_or_None, subgroup, col_a_bullets: list[str], col_b_bullets: list[str]), ...]
    """
    if rows is None:
        rows = [
            (
                None, "プロセス",
                ["紙ベースの申請書で窓口受付。", "担当者が手作業でデータ入力。"],
                ["Web申請フォームで完結。", "データは自動連携・即時反映。"],
            ),
            (
                "管理機能", "承認フロー",
                ["上長への紙回覧・押印が必要。", "承認状況の把握が困難。"],
                ["システム上で並行承認・多段階承認が可能。", "承認状況をリアルタイム確認。"],
            ),
            (
                "管理機能", "データ連携",
                ["基幹システムとの連携なし。", "CSVによる手動取込みが必要。"],
                ["API連携で基幹システムとリアルタイム同期。", "CSV/PDF出力に対応。"],
            ),
            (
                "品質・統制", "エラーチェック",
                ["目視確認のみ。", "入力ミスの発見が遅れる。"],
                ["必須・桁数・相関チェックを自動実施。", "エラー即時通知で品質向上。"],
            ),
            (
                "品質・統制", "監査証跡",
                ["紙ファイル保管のみ。", "検索・集計が困難。"],
                ["全操作ログを自動記録。", "クエリで即時抽出・分析が可能。"],
            ),
        ]

    pal = THEME_PAL.get(theme, THEME_PAL["default"])
    sb = SlideBuilder(theme=theme)
    slide = sb.add_body(title, section=section)

    render_comparison_matrix(
        slide,
        x=0.35,
        y=CONTENT_TOP + 0.25,
        w=12.45,
        col_a_label=col_a_label,
        col_b_label=col_b_label,
        rows=rows,
        pal=pal,
    )

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
