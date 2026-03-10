"""
パターン名: data_feature_list
カテゴリ: データ
説明: 番号付き機能・要件一覧表（ラベル列カラー付き、テキスト列動的行高）
用途: 機能一覧、要件一覧、課題一覧、リスク一覧、施策一覧
ドキュメント: 提案書, SOW, ビジネスケース, フィンディングス, 戦略提言
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_rect, add_line, CONTENT_TOP


def render_feature_list(
    slide,
    x: float,
    y: float,
    w: float,
    col_headers: list[str],       # 例: ["#", "機能名称", "機能概要", "背景・目的"]
    rows: list[dict],             # {"label": str, "cols": [str, str, ...]}
    col_widths: list[float] | None = None,
    accent_color: str = "1E6B3C", # ラベル列の背景色
    accent_text: str = "FFFFFF",
    hdr_h: float = 0.38,
    line_h: float = 0.185,        # 1行あたりの高さ
    pad_h: float = 0.22,          # セル上下パディング合計
    min_row_h: float = 0.52,
    font_size: int = 10,
):
    """
    col_headers: [番号列ヘッダー, ラベル列ヘッダー, テキスト列1ヘッダー, ...]
    rows: [{"label": "ラベル（色付き）", "cols": ["テキスト1", "テキスト2", ...]}, ...]
          cols の長さは col_headers - 2 に合わせる
    """
    n_cols = len(col_headers)   # 番号列 + ラベル列 + テキスト列×n

    # ── 列幅決定 ──────────────────────────────────────────────────
    if col_widths is None:
        num_w   = 0.42
        label_w = 1.80
        text_w  = (w - num_w - label_w) / max(n_cols - 2, 1)
        col_widths = [num_w, label_w] + [text_w] * (n_cols - 2)

    assert len(col_widths) == n_cols

    # ── 行高さ計算（テキスト量から推定）─────────────────────────────
    def estimate_lines(text: str, cell_w: float, fsize: int = font_size) -> int:
        """テキストの折り返し行数を推定"""
        if not text:
            return 1
        chars_per_line = max(1, int(cell_w / (fsize * 0.013)))  # 簡易推定
        lines = 0
        for para in text.split("\n"):
            lines += max(1, (len(para) + chars_per_line - 1) // chars_per_line)
        return lines

    row_hs = []
    for row in rows:
        label_lines = estimate_lines(row.get("label", ""), col_widths[1])
        col_lines   = [estimate_lines(c, col_widths[2 + i]) for i, c in enumerate(row.get("cols", []))]
        max_lines   = max(label_lines, *col_lines) if col_lines else label_lines
        row_hs.append(max(min_row_h, max_lines * line_h + pad_h))

    # ── ヘッダー行（線＋テキスト、塗りつぶしなし） ──────────────────
    # 上部ライン
    add_line(slide, x, y, x + w, y, color="888888", width=1.0)
    cx = x
    for i, (hdr, cw) in enumerate(zip(col_headers, col_widths)):
        align = "center" if i <= 1 else "left"
        tx = cx + (0.08 if i > 1 else 0)
        tw = cw - (0.12 if i > 1 else 0)
        add_text(slide, tx, y + 0.04, tw, hdr_h - 0.06,
                 hdr, style="small", align=align, bold=True, color="1A1A1A")
        cx += cw
    # 下部ライン（ヘッダーとデータの区切り）
    add_line(slide, x, y + hdr_h, x + w, y + hdr_h, color="888888", width=1.0)

    # ── データ行 ──────────────────────────────────────────────────
    ry = y + hdr_h
    for idx, (row, rh) in enumerate(zip(rows, row_hs)):
        is_even = idx % 2 == 0
        row_bg  = "FFFFFF" if is_even else "F7F7F7"
        cx = x

        # 番号セル
        add_rect(slide, cx, ry, col_widths[0], rh, fill=row_bg, border="DDDDDD")
        add_text(slide, cx, ry, col_widths[0], rh, str(idx + 1),
                 style="small", align="center", bold=True, color="888888")
        cx += col_widths[0]

        # ラベルセル（アクセントカラー）
        add_rect(slide, cx, ry, col_widths[1], rh, fill=accent_color, border="BBBBBB")
        add_text(slide, cx + 0.08, ry + 0.06, col_widths[1] - 0.14, rh - 0.10,
                 row.get("label", ""), style="small", align="center", bold=True,
                 color=accent_text, font_size=font_size)
        cx += col_widths[1]

        # テキスト列
        for i, col_text in enumerate(row.get("cols", [])):
            cw = col_widths[2 + i]
            add_rect(slide, cx, ry, cw, rh, fill=row_bg, border="DDDDDD")
            add_text(slide, cx + 0.10, ry + 0.07, cw - 0.16, rh - 0.12,
                     col_text, style="caption", align="left", color="1A1A1A",
                     font_size=font_size)
            cx += cw

        ry += rh

    # 外枠
    total_h = sum(row_hs) + hdr_h
    add_rect(slide, x, y, w, total_h, fill=None, border="888888", border_width=1.5)


def run(
    title: str = "検証対象機能一覧",
    section: str = "検証対象機能の概要",
    col_headers: list[str] | None = None,
    rows: list[dict] | None = None,
    theme: str = "default",
    accent_color: str = "1E6B3C",
    output_path: str = "output/data_feature_list.pptx",
):
    if col_headers is None:
        col_headers = ["#", "機能名称", "機能概要", "背景・目的"]

    if rows is None:
        rows = [
            {
                "label": "マスタ管理機能",
                "cols": [
                    "各事業で管理している事業者情報等のマスタデータ上の項目と汎化された画面項目を項目単位でマッピングすることで、データの更新及び入力時の参照処理を実施できる機能。",
                    "電子化されていない手続の中にはデータのマスタ管理が必要な手続が多く、汎化申請を利用する手続の対象範囲を拡大するために、データを管理する仕組みが必要。",
                ],
            },
            {
                "label": "汎化申請CSV\nデータ出力",
                "cols": [
                    "手続情報をCSV形式で出力する機能。",
                    "手続情報を利活用しての集計業務や他システムへの連携等が存在することから、手続情報を一括で取得する機能が求められた。",
                ],
            },
            {
                "label": "表形式",
                "cols": [
                    "複数のレコードの入力が必要なリスト形式の画面項目に関して、Excelのように表形式で入力可能とする機能。",
                    "報告系の手続等は入力項目が多いことに加えて、現状の様式で表形式を利用しているため、事業者目線でより入力の操作性、工数を考えたレイアウトが必要。",
                ],
            },
            {
                "label": "詳細画面の\nページ分け",
                "cols": [
                    "手続毎に設定可能な画面項目は、これまで1つの詳細画面のみに配置可能であったが、詳細画面を複数画面に分割できる機能。",
                    "項目数が多いと縦スクロールが発生することに加えて、報告系の手続等は既存の様式で複数のシートに分割されているという点から、事業者による入力をより直感的に実施できるようにするための機能が必要。",
                ],
            },
            {
                "label": "過去手続からの\n複写機能",
                "cols": [
                    "新たに手続を開始する際に、過去既に申請した手続内容と同様の内容を手続入力画面上に初期表示する機能。",
                    "定期的に提出が必要な手続等、類似した入力内容を複数回提出するような手続が存在しており、そのような場合に入力工数を削減したいと考える。",
                ],
            },
        ]

    sb = SlideBuilder(theme=theme)
    slide = sb.add_body(title, section=section)

    # 説明文（省略可）
    add_text(slide, 0.35, CONTENT_TOP + 0.12, 12.45, 0.36,
             "本事業において調査・技術検証対象とした機能の一覧に関しては以下の通り。",
             style="body", align="left", color="1A1A1A")

    render_feature_list(
        slide,
        x=0.35,
        y=CONTENT_TOP + 0.58,
        w=12.45,
        col_headers=col_headers,
        rows=rows,
        accent_color=accent_color,
    )

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
