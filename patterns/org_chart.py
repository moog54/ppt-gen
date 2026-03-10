"""
パターン名: org_chart
カテゴリ: 構造
説明: 組織図・体制図（ツリー構造、レベル別カラーリング、自動レイアウト）
用途: 組織体制、プロジェクト体制、役割分担、責任マトリクス
ドキュメント: 提案書, SOW, ビジネスケース, ステアリングコミッティ, ワークショップ
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_rect, add_line, CONTENT_TOP

# ──────────────────────────────────────────────────────────────
# データ構造:
#   tree = {
#       "label": "役職名\nサブタイトル",  # \n で2行
#       "name":  "氏名（省略可）",
#       "color": "RRGGBB（省略可、省略時はレベルで自動決定）",
#       "dotted": False,  # True にすると点線接続（兼務・アドバイザー等）
#       "children": [ {...}, ... ]
#   }
# ──────────────────────────────────────────────────────────────

LEVEL_COLORS_DEFAULT = [
    "1A3C6E",   # Level 0: root（濃紺）
    "2B6CB0",   # Level 1
    "4A90D9",   # Level 2
    "7AB3E0",   # Level 3
    "B3D4F0",   # Level 4+
]
LEVEL_TEXT_COLORS = [
    "FFFFFF", "FFFFFF", "FFFFFF", "FFFFFF", "1A1A1A",
]

THEME_LEVEL_COLORS = {
    "accenture": ["5B0099", "A100FF", "BE4DFF", "D98AFF", "ECC6FF"],
    "mckinsey":  ["002F6C", "1455A0", "2980B9", "5DADE2", "A9CCE3"],
    "navy":      ["1A3C6E", "2B5DA7", "4C80C4", "7CAAD9", "AECCE8"],
    "default":   ["1A3C6E", "2B6CB0", "4A90D9", "7AB3E0", "B3D4F0"],
}


def _calc_subtree_w(node, box_w, h_gap):
    """このノードの部分ツリーが必要とする横幅を返す"""
    children = node.get("children", [])
    if not children:
        return box_w
    children_w = sum(_calc_subtree_w(c, box_w, h_gap) for c in children)
    children_w += h_gap * (len(children) - 1)
    return max(box_w, children_w)


def _layout(node, cx, y, box_w, box_h, h_gap, v_gap, level, result):
    """再帰的にノードの (center_x, top_y, level) を result に追加"""
    result.append({
        "node":  node,
        "cx":    cx,
        "y":     y,
        "level": level,
    })
    children = node.get("children", [])
    if not children:
        return
    # 子ノード群の合計幅を計算して左端を決定
    total_w = sum(_calc_subtree_w(c, box_w, h_gap) for c in children)
    total_w += h_gap * (len(children) - 1)
    child_y = y + box_h + v_gap
    lx = cx - total_w / 2   # 左端 center x
    for child in children:
        sw = _calc_subtree_w(child, box_w, h_gap)
        child_cx = lx + sw / 2
        _layout(child, child_cx, child_y, box_w, box_h, h_gap, v_gap, level + 1, result)
        lx += sw + h_gap


def render_org_chart(
    slide,
    tree: dict,
    x: float,
    y: float,
    w: float,
    max_h: float,
    level_colors: list[str] | None = None,
    box_h: float = 0.58,
    h_gap: float = 0.22,
    v_gap: float = 0.50,
    connector_color: str = "888888",
    connector_width: float = 1.0,
):
    """
    tree: ノードの辞書ツリー（上記データ構造参照）
    x, y: 描画エリア左上
    w: 描画エリア幅
    max_h: 描画エリア最大高さ
    """
    if level_colors is None:
        level_colors = LEVEL_COLORS_DEFAULT

    # ── Step 1: デフォルト box_w で全ツリー幅を計算 ──────────────
    box_w_default = 1.80
    total_w = _calc_subtree_w(tree, box_w_default, h_gap)

    # 収まらない場合はスケール調整
    if total_w > w:
        scale = w / total_w
        box_w = box_w_default * scale
        h_gap = h_gap * scale
    else:
        box_w = box_w_default

    # ── Step 2: レイアウト計算 ───────────────────────────────────
    positioned = []
    root_cx = x + w / 2
    _layout(tree, root_cx, y, box_w, box_h, h_gap, v_gap, 0, positioned)

    # ノードIDで引けるよう辞書化（cx検索用）
    pos_map = {id(p["node"]): p for p in positioned}

    # ── Step 3: コネクタ描画（ボックスより先に描いて背面に） ─────
    def draw_connectors(node):
        children = node.get("children", [])
        if not children:
            return
        p = pos_map[id(node)]
        parent_bx = p["cx"]
        parent_by = p["y"] + box_h   # 親ボックス下端

        child_tops = []
        for child in children:
            c = pos_map[id(child)]
            child_tops.append((c["cx"], c["y"], child.get("dotted", False)))

        mid_y = parent_by + v_gap / 2

        # 親から中点まで縦線
        add_line(slide, parent_bx, parent_by, parent_bx, mid_y,
                 color=connector_color, width=connector_width)

        # 子が複数なら中点で横線
        if len(child_tops) > 1:
            lx = min(cx for cx, _, _ in child_tops)
            rx = max(cx for cx, _, _ in child_tops)
            add_line(slide, lx, mid_y, rx, mid_y,
                     color=connector_color, width=connector_width)

        # 各子の中点→子ボックス上端まで縦線
        for child_cx, child_y_top, dotted in child_tops:
            add_line(slide, child_cx, mid_y, child_cx, child_y_top,
                     color=connector_color, width=connector_width)

        for child in children:
            draw_connectors(child)

    draw_connectors(tree)

    # ── Step 4: ボックス描画 ──────────────────────────────────────
    for p in positioned:
        node  = p["node"]
        cx    = p["cx"]
        ny    = p["y"]
        level = p["level"]

        col_idx = min(level, len(level_colors) - 1)
        fill    = node.get("color") or level_colors[col_idx]
        txt_col = LEVEL_TEXT_COLORS[col_idx] if not node.get("color") else (
            "1A1A1A" if int(fill[0:2], 16) * 0.299 +
                        int(fill[2:4], 16) * 0.587 +
                        int(fill[4:6], 16) * 0.114 > 150 else "FFFFFF"
        )

        bx = cx - box_w / 2
        add_rect(slide, bx, ny, box_w, box_h, fill=fill, border="FFFFFF", border_width=1.5)

        label = node.get("label", "")
        name  = node.get("name", "")

        if name:
            # 上部: ラベル / 下部: 氏名
            add_text(slide, bx + 0.06, ny + 0.04, box_w - 0.12, box_h * 0.58,
                     label, style="small", align="center", bold=True,
                     color=txt_col, font_size=9)
            add_text(slide, bx + 0.06, ny + box_h * 0.58, box_w - 0.12, box_h * 0.38,
                     name, style="caption", align="center",
                     color=txt_col, font_size=9)
        else:
            add_text(slide, bx + 0.06, ny + 0.06, box_w - 0.12, box_h - 0.12,
                     label, style="small", align="center", bold=True,
                     color=txt_col, font_size=9)


def run(
    title: str = "プロジェクト体制図",
    section: str = "プロジェクト概要",
    theme: str = "default",
    tree: dict | None = None,
    output_path: str = "output/org_chart.pptx",
):
    if tree is None:
        tree = {
            "label": "プロジェクト\nマネージャー",
            "name": "山田 太郎",
            "children": [
                {
                    "label": "アドバイザー",
                    "name": "佐藤 顧問",
                    "dotted": True,
                },
                {
                    "label": "PMO",
                    "name": "田中 花子",
                    "children": [
                        {"label": "スケジュール\n管理"},
                        {"label": "リスク\n管理"},
                        {"label": "品質\n管理"},
                    ],
                },
                {
                    "label": "WS1 リーダー\n業務改革",
                    "name": "鈴木 一郎",
                    "children": [
                        {"label": "業務分析\n担当"},
                        {"label": "業務設計\n担当"},
                    ],
                },
                {
                    "label": "WS2 リーダー\nシステム",
                    "name": "伊藤 二郎",
                    "children": [
                        {"label": "要件定義\n担当"},
                        {"label": "開発\n担当"},
                        {"label": "テスト\n担当"},
                    ],
                },
            ],
        }

    level_colors = THEME_LEVEL_COLORS.get(theme, THEME_LEVEL_COLORS["default"])

    sb = SlideBuilder(theme=theme)
    slide = sb.add_body(title, section=section)

    render_org_chart(
        slide,
        tree=tree,
        x=0.35,
        y=CONTENT_TOP + 0.15,
        w=12.45,
        max_h=7.0 - CONTENT_TOP - 0.15 - 0.4,
        level_colors=level_colors,
    )

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
