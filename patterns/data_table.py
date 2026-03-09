"""
パターン名: data_table
カテゴリ: データ
説明: 図形ベースのデータテーブル（pptx native table 禁止）
用途: 一覧表、データ比較、スケジュール表
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib import SlideBuilder, add_text, add_table_shapes, C, CONTENT_TOP


def run(
    title: str = "データテーブル",
    section: str = "",
    headers: list[str] | None = None,
    rows: list[list[str]] | None = None,
    col_widths: list[float] | None = None,
    output_path: str = "output/data_table.pptx",
):
    if headers is None:
        headers = ["項目", "Q1", "Q2", "Q3", "Q4", "合計"]
    if rows is None:
        rows = [
            ["売上高（億円）", "2.8", "3.1", "3.2", "3.3", "12.4"],
            ["営業利益（億円）", "0.4", "0.5", "0.6", "0.6", "2.1"],
            ["利益率（%）", "14.3", "16.1", "18.8", "18.2", "16.9"],
            ["顧客数（社）", "4,200", "4,450", "4,680", "4,820", "—"],
        ]
    if col_widths is None:
        total_w = 12.33
        n = len(headers)
        col_widths = [2.5] + [(total_w - 2.5) / (n - 1)] * (n - 1)

    sb = SlideBuilder()
    slide = sb.add_body(title, section)

    add_table_shapes(
        slide,
        x=0.5,
        y=CONTENT_TOP + 0.2,
        w=12.33,
        headers=headers,
        rows=rows,
        col_widths=col_widths,
        row_h=0.45,
    )

    sb.save_and_validate(output_path)
    print(f"生成: {output_path}")


if __name__ == "__main__":
    run()
