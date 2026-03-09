"""
master.pptx を生成するスクリプト。
python-pptx でプログラム的にスライドマスターを構築する。

使用方法:
    /home/moog/hr_venv/bin/python tools/gen_master.py
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from lib import _rgb, C, SLIDE_W, SLIDE_H, DEFAULT_FONT

OUTPUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "master.pptx")


def set_font_scheme(prs: Presentation):
    """テーマフォントを Meiryo UI に設定する"""
    from lxml import etree
    from pptx.oxml.ns import qn

    theme_elem = prs.slide_master.theme_color_map
    # フォント変更はテーマXMLレベルで行う
    # slide_master の txStyles を更新
    txStyles = prs.slide_master._element.find(qn("p:txStyles"))
    if txStyles is None:
        return

    for defRPr in txStyles.iter(qn("a:defRPr")):
        latin = defRPr.find(qn("a:latin"))
        if latin is None:
            latin = etree.SubElement(defRPr, qn("a:latin"))
        latin.set("typeface", DEFAULT_FONT)


def main():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    # 空白レイアウト（index 6）を確認
    layouts = prs.slide_layouts
    print(f"利用可能なレイアウト数: {len(layouts)}")
    for i, layout in enumerate(layouts):
        print(f"  [{i}] {layout.name}")

    # master.pptx として保存（SlideBuilder が参照する）
    prs.save(OUTPUT)
    print(f"\nmaster.pptx を生成しました: {OUTPUT}")


if __name__ == "__main__":
    main()
