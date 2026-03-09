"""
パターンカタログを生成するスクリプト。

処理フロー:
1. patterns/*.py を全て検出・実行 → catalog/tmp/{name}.pptx
2. LibreOffice で PPTX → PNG 変換（未インストールの場合はスキップ）
3. docstring からメタデータ抽出
4. catalog/index.html 生成

使用方法:
    /home/moog/hr_venv/bin/python tools/gen_catalog.py
"""
import sys
import os
import re
import importlib.util
import shutil
import subprocess
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

PATTERNS_DIR = PROJECT_ROOT / "patterns"
CATALOG_DIR = PROJECT_ROOT / "catalog"
IMAGES_DIR = CATALOG_DIR / "images"
TMP_DIR = CATALOG_DIR / "tmp"


def extract_metadata(path: Path) -> dict:
    """パターンファイルの docstring からメタデータを抽出する"""
    content = path.read_text(encoding="utf-8")
    docstring_match = re.search(r'"""(.*?)"""', content, re.DOTALL)
    if not docstring_match:
        return {"name": path.stem, "category": "未分類", "description": "", "usage": ""}

    docstring = docstring_match.group(1)
    meta = {}
    for line in docstring.splitlines():
        line = line.strip()
        for key, attr in [
            ("パターン名:", "name"),
            ("カテゴリ:", "category"),
            ("説明:", "description"),
            ("用途:", "usage"),
        ]:
            if line.startswith(key):
                meta[attr] = line[len(key):].strip()

    return {
        "name": meta.get("name", path.stem),
        "category": meta.get("category", "未分類"),
        "description": meta.get("description", ""),
        "usage": meta.get("usage", ""),
        "filename": path.name,
    }


def run_pattern(path: Path, output_pptx: Path) -> bool:
    """パターンファイルを実行して PPTX を生成する"""
    spec = importlib.util.spec_from_file_location(path.stem, path)
    mod = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(mod)
        if hasattr(mod, "run"):
            mod.run(output_path=str(output_pptx))
            return True
    except Exception as e:
        print(f"  ERROR running {path.name}: {e}")
    return False


def convert_pptx_to_png(pptx_path: Path, images_dir: Path) -> Path | None:
    """LibreOffice で PPTX → PNG 変換する"""
    libreoffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not libreoffice:
        return None

    try:
        result = subprocess.run(
            [libreoffice, "--headless", "--convert-to", "png",
             "--outdir", str(images_dir), str(pptx_path)],
            capture_output=True, text=True, timeout=30
        )
        # LibreOffice は ファイル名の1枚目のみ {name}.png として出力
        # 複数スライドの場合は {name}.png, {name}2.png, ...
        png_path = images_dir / (pptx_path.stem + ".png")
        if png_path.exists():
            return png_path
        # 先頭スライドとして連番が付く場合
        for suffix in ["0001.png", "0000.png"]:
            alt = images_dir / (pptx_path.stem + suffix)
            if alt.exists():
                shutil.copy(alt, png_path)
                return png_path
    except Exception as e:
        print(f"  LibreOffice 変換エラー: {e}")
    return None


def build_html(patterns_meta: list[dict]) -> str:
    """カタログ HTML を生成する"""
    categories = sorted(set(m["category"] for m in patterns_meta))

    filter_buttons = ""
    for cat in ["すべて"] + categories:
        data_cat = "" if cat == "すべて" else cat
        active = ' class="active"' if cat == "すべて" else ""
        filter_buttons += f'<button onclick="filterCat(\'{data_cat}\')" data-cat="{data_cat}"{active}>{cat}</button>\n'

    cards_html = ""
    for m in patterns_meta:
        img_path = f"images/{m['name']}.png"
        has_img = (IMAGES_DIR / f"{m['name']}.png").exists()
        img_html = (f'<img src="{img_path}" alt="{m["name"]}">'
                    if has_img else
                    '<div class="no-img">プレビューなし<br>LibreOfficeをインストールして<br>gen_catalog.py を再実行</div>')
        cards_html += f"""
        <div class="card" data-category="{m['category']}">
          <div class="thumbnail">{img_html}</div>
          <div class="card-body">
            <div class="category-badge">{m['category']}</div>
            <h3>{m['name']}</h3>
            <p class="description">{m['description']}</p>
            <p class="usage"><span>用途:</span> {m['usage']}</p>
            <code>{m['filename']}</code>
          </div>
        </div>"""

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ppt-gen パターンカタログ</title>
<style>
  :root {{
    --accent: #0070C0;
    --text: #1A1A1A;
    --textLight: #666666;
    --bg: #ffffff;
    --bgLight: #F5F5F5;
    --border: #D9D9D9;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: "Meiryo UI", "Segoe UI", sans-serif; background: var(--bgLight); color: var(--text); }}
  header {{
    background: var(--accent); color: white; padding: 24px 32px;
    display: flex; align-items: center; gap: 16px;
  }}
  header h1 {{ font-size: 22px; font-weight: bold; }}
  header p {{ font-size: 13px; opacity: 0.85; margin-top: 4px; }}
  .filter-bar {{
    background: white; border-bottom: 1px solid var(--border);
    padding: 12px 32px; display: flex; gap: 8px; flex-wrap: wrap;
  }}
  .filter-bar button {{
    background: var(--bgLight); border: 1px solid var(--border);
    padding: 6px 16px; border-radius: 20px; cursor: pointer;
    font-size: 13px; color: var(--textLight); transition: all 0.2s;
  }}
  .filter-bar button:hover, .filter-bar button.active {{
    background: var(--accent); color: white; border-color: var(--accent);
  }}
  .grid {{
    display: grid; grid-template-columns: repeat(3, 1fr);
    gap: 20px; padding: 24px 32px;
    max-width: 1400px; margin: 0 auto;
  }}
  @media (max-width: 1100px) {{ .grid {{ grid-template-columns: repeat(2, 1fr); }} }}
  @media (max-width: 700px) {{ .grid {{ grid-template-columns: 1fr; }} }}
  .card {{
    background: white; border: 1px solid var(--border); border-radius: 8px;
    overflow: hidden; transition: transform 0.2s, box-shadow 0.2s;
  }}
  .card:hover {{ transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,112,192,0.12); }}
  .card.hidden {{ display: none; }}
  .thumbnail {{ background: var(--bgLight); height: 200px; overflow: hidden; }}
  .thumbnail img {{ width: 100%; height: 100%; object-fit: contain; }}
  .no-img {{
    height: 100%; display: flex; align-items: center; justify-content: center;
    font-size: 12px; color: var(--textLight); text-align: center; line-height: 1.8;
  }}
  .card-body {{ padding: 16px; }}
  .category-badge {{
    display: inline-block; background: var(--accent); color: white;
    font-size: 11px; padding: 2px 10px; border-radius: 10px; margin-bottom: 8px;
  }}
  .card-body h3 {{ font-size: 15px; font-weight: bold; margin-bottom: 6px; }}
  .description {{ font-size: 13px; color: var(--textLight); margin-bottom: 6px; line-height: 1.5; }}
  .usage {{ font-size: 12px; color: #999; margin-bottom: 10px; }}
  .usage span {{ font-weight: bold; color: var(--textLight); }}
  code {{ font-size: 11px; background: var(--bgLight); padding: 3px 8px; border-radius: 4px; color: var(--accent); }}
  footer {{ text-align: center; padding: 24px; color: var(--textLight); font-size: 12px; }}
</style>
</head>
<body>
<header>
  <div>
    <h1>ppt-gen パターンカタログ</h1>
    <p>汎用PPTXスライド生成ツール — 全 {len(patterns_meta)} パターン</p>
  </div>
</header>
<div class="filter-bar">
  {filter_buttons}
</div>
<div class="grid" id="grid">
  {cards_html}
</div>
<footer>ppt-gen &copy; 2026 — <code>tools/gen_catalog.py</code> で再生成</footer>
<script>
  function filterCat(cat) {{
    document.querySelectorAll('.filter-bar button').forEach(b => {{
      b.classList.toggle('active', b.dataset.cat === cat);
    }});
    document.querySelectorAll('.card').forEach(card => {{
      if (!cat || card.dataset.category === cat) {{
        card.classList.remove('hidden');
      }} else {{
        card.classList.add('hidden');
      }}
    }});
  }}
</script>
</body>
</html>"""


def main():
    CATALOG_DIR.mkdir(exist_ok=True)
    IMAGES_DIR.mkdir(exist_ok=True)
    TMP_DIR.mkdir(exist_ok=True)

    pattern_files = sorted(PATTERNS_DIR.glob("*.py"))
    pattern_files = [f for f in pattern_files if f.name != "__init__.py"]
    print(f"パターンファイル: {len(pattern_files)} 件")

    patterns_meta = []
    for pf in pattern_files:
        print(f"\n処理中: {pf.name}")
        meta = extract_metadata(pf)
        meta["filename"] = pf.name

        # PPTX 生成
        out_pptx = TMP_DIR / f"{pf.stem}.pptx"
        success = run_pattern(pf, out_pptx)
        if not success:
            continue

        # PNG 変換
        if out_pptx.exists():
            png = convert_pptx_to_png(out_pptx, IMAGES_DIR)
            if png:
                print(f"  PNG: {png.name}")
            else:
                print(f"  PNG: スキップ（LibreOffice未検出）")

        patterns_meta.append(meta)

    # HTML 生成
    html = build_html(patterns_meta)
    html_path = CATALOG_DIR / "index.html"
    html_path.write_text(html, encoding="utf-8")
    print(f"\nカタログ生成完了: {html_path}")
    print(f"  パターン数: {len(patterns_meta)}")
    print(f"  ブラウザで開く: file://{html_path.resolve()}")


if __name__ == "__main__":
    main()
