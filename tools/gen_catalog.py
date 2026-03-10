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
            ("ドキュメント:", "documents"),
        ]:
            if line.startswith(key):
                meta[attr] = line[len(key):].strip()

    return {
        "name": meta.get("name", path.stem),
        "category": meta.get("category", "未分類"),
        "description": meta.get("description", ""),
        "usage": meta.get("usage", ""),
        "documents": meta.get("documents", ""),
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


THEME_LABELS = {
    "default":   ("企業ブルー",           "汎用・デフォルト"),
    "accenture": ("アクセンチュアパープル", "コンサルティング・提案書"),
    "navy":      ("ネイビー",              "重厚感・金融・官公庁"),
    "green":     ("グリーン",              "ESG・サステナビリティ"),
    "warm":      ("ウォームオレンジ",      "エネルギー・製造・スタートアップ"),
    "mckinsey":  ("マッキンゼー濃紺",      "戦略コンサル・経営報告"),
}


def build_theme_swatches(themes: dict) -> str:
    """テーマ一覧のスウォッチHTMLを生成する"""
    cards = ""
    for name, colors in themes.items():
        from lib import _C_DEFAULT
        merged = {**_C_DEFAULT, **colors}
        accent = f"#{merged['accent']}"
        accent2 = f"#{merged['accent2']}"
        accent3 = f"#{merged['accent3']}"
        row_alt = f"#{merged['rowAlt']}"
        label, desc = THEME_LABELS.get(name, (name, ""))

        img_path = f"images/theme_{name}.png"
        has_img = (IMAGES_DIR / f"theme_{name}.png").exists()
        img_html = (f'<img src="{img_path}" alt="{name}">'
                    if has_img else
                    '<div class="no-img">プレビューなし</div>')

        cards += f"""
        <div class="theme-card">
          <div class="thumbnail">{img_html}</div>
          <div class="swatches">
            <span class="swatch" style="background:{accent}" title="accent {accent}"></span>
            <span class="swatch" style="background:{accent2}" title="accent2 {accent2}"></span>
            <span class="swatch" style="background:{accent3}" title="accent3 {accent3}"></span>
            <span class="swatch" style="background:{row_alt};border:1px solid #ccc" title="rowAlt {row_alt}"></span>
          </div>
          <div class="card-body">
            <h3>{name}</h3>
            <p class="description">{label}</p>
            <p class="usage">{desc}</p>
            <code>SlideBuilder(theme="{name}")</code>
          </div>
        </div>"""
    return cards


def build_html(patterns_meta: list[dict], themes: dict, styles_html: str = "") -> str:
    """カタログ HTML を生成する"""
    categories = sorted(set(m["category"] for m in patterns_meta))

    # カテゴリフィルター
    cat_buttons = ""
    for cat in ["すべて"] + categories:
        data_cat = "" if cat == "すべて" else cat
        active = ' class="active"' if cat == "すべて" else ""
        cat_buttons += f'<button onclick="filterCat(\'{data_cat}\')" data-cat="{data_cat}"{active}>{cat}</button>\n'

    # ドキュメントタイプフィルター
    all_docs = set()
    for m in patterns_meta:
        for d in m.get("documents", "").split(","):
            d = d.strip()
            if d:
                all_docs.add(d)
    doc_order = ["ドアノッカー", "提案書", "SOW", "ビジネスケース",
                 "定例報告", "ステアリングコミッティ", "ワークショップ",
                 "フィンディングス", "戦略提言", "エグゼクティブブリーフィング"]
    sorted_docs = [d for d in doc_order if d in all_docs] + sorted(all_docs - set(doc_order))

    doc_buttons = ""
    for doc in ["すべて"] + sorted_docs:
        data_doc = "" if doc == "すべて" else doc
        active = ' class="active"' if doc == "すべて" else ""
        doc_buttons += f'<button onclick="filterDoc(\'{data_doc}\')" data-doc="{data_doc}"{active}>{doc}</button>\n'

    cards_html = ""
    for m in patterns_meta:
        img_path = f"images/{m['name']}.png"
        has_img = (IMAGES_DIR / f"{m['name']}.png").exists()
        img_html = (f'<img src="{img_path}" alt="{m["name"]}">'
                    if has_img else
                    '<div class="no-img">プレビューなし<br>LibreOfficeをインストールして<br>gen_catalog.py を再実行</div>')
        docs_attr = m.get("documents", "")
        doc_tags = ""
        for d in docs_attr.split(","):
            d = d.strip()
            if d:
                doc_tags += f'<span class="doc-tag">{d}</span>'
        cards_html += f"""
        <div class="card" data-category="{m['category']}" data-documents="{docs_attr}">
          <div class="thumbnail">{img_html}</div>
          <div class="card-body">
            <div class="category-badge">{m['category']}</div>
            <h3>{m['name']}</h3>
            <p class="description">{m['description']}</p>
            <p class="usage"><span>用途:</span> {m['usage']}</p>
            <div class="doc-tags">{doc_tags}</div>
            <code>{m['filename']}</code>
          </div>
        </div>"""

    theme_swatches = build_theme_swatches(themes)

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
  .section-title {{
    font-size: 18px; font-weight: bold; padding: 24px 32px 8px;
    max-width: 1400px; margin: 0 auto; color: var(--text);
    border-bottom: 2px solid var(--accent); margin-bottom: 0;
  }}
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
  .theme-grid {{
    display: grid; grid-template-columns: repeat(3, 1fr);
    gap: 20px; padding: 16px 32px 32px;
    max-width: 1400px; margin: 0 auto;
  }}
  @media (max-width: 1100px) {{
    .grid, .theme-grid {{ grid-template-columns: repeat(2, 1fr); }}
  }}
  @media (max-width: 700px) {{
    .grid, .theme-grid {{ grid-template-columns: 1fr; }}
  }}
  .card, .theme-card {{
    background: white; border: 1px solid var(--border); border-radius: 8px;
    overflow: hidden; transition: transform 0.2s, box-shadow 0.2s;
  }}
  .card:hover, .theme-card:hover {{ transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,112,192,0.12); }}
  .card.hidden {{ display: none; }}
  .thumbnail {{ background: var(--bgLight); height: 200px; overflow: hidden; }}
  .thumbnail img {{ width: 100%; height: 100%; object-fit: contain; }}
  .no-img {{
    height: 100%; display: flex; align-items: center; justify-content: center;
    font-size: 12px; color: var(--textLight); text-align: center; line-height: 1.8;
  }}
  .swatches {{ display: flex; gap: 0; }}
  .swatch {{ flex: 1; height: 12px; }}
  .card-body, .theme-card .card-body {{ padding: 16px; }}
  .category-badge {{
    display: inline-block; background: var(--accent); color: white;
    font-size: 11px; padding: 2px 10px; border-radius: 10px; margin-bottom: 8px;
  }}
  .card-body h3, .theme-card h3 {{ font-size: 15px; font-weight: bold; margin-bottom: 6px; }}
  .description {{ font-size: 13px; color: var(--textLight); margin-bottom: 4px; line-height: 1.5; }}
  .usage {{ font-size: 12px; color: #999; margin-bottom: 10px; }}
  .usage span {{ font-weight: bold; color: var(--textLight); }}
  code {{ font-size: 11px; background: var(--bgLight); padding: 3px 8px; border-radius: 4px; color: var(--accent); display: inline-block; margin-top: 4px; }}
  .doc-tags {{ display: flex; flex-wrap: wrap; gap: 4px; margin: 6px 0; }}
  .doc-tag {{ font-size: 10px; background: #EFF4FB; color: #1A3C6E; border: 1px solid #C0D0E8; padding: 2px 7px; border-radius: 10px; }}
  .filter-label {{ font-size: 11px; font-weight: bold; color: var(--textLight); margin-right: 4px; white-space: nowrap; align-self: center; }}
  footer {{ text-align: center; padding: 24px; color: var(--textLight); font-size: 12px; }}
</style>
</head>
<body>
<header>
  <div>
    <h1>ppt-gen パターンカタログ</h1>
    <p>汎用PPTXスライド生成ツール — {len(themes)} テーマ / {len(patterns_meta)} パターン</p>
  </div>
</header>

<div class="section-title" style="padding-top:28px">テーマ一覧</div>
<div class="theme-grid">
  {theme_swatches}
</div>

{styles_html}

<div class="section-title">パターン一覧</div>
<div class="filter-bar">
  <span class="filter-label">カテゴリ:</span>
  {cat_buttons}
</div>
<div class="filter-bar" style="border-top:none">
  <span class="filter-label">ドキュメント:</span>
  {doc_buttons}
</div>
<div class="grid" id="grid">
  {cards_html}
</div>
<footer>ppt-gen &copy; 2026 — <code>tools/gen_catalog.py</code> で再生成</footer>
<script>
  let activeCat = "";
  let activeDoc = "";

  function applyFilters() {{
    document.querySelectorAll('.card').forEach(card => {{
      const catOk = !activeCat || card.dataset.category === activeCat;
      const docOk = !activeDoc || (card.dataset.documents || "").split(",").map(d => d.trim()).includes(activeDoc);
      card.classList.toggle('hidden', !(catOk && docOk));
    }});
  }}

  function filterCat(cat) {{
    activeCat = cat;
    document.querySelectorAll('[data-cat]').forEach(b => b.classList.toggle('active', b.dataset.cat === cat));
    applyFilters();
  }}

  function filterDoc(doc) {{
    activeDoc = doc;
    document.querySelectorAll('[data-doc]').forEach(b => b.classList.toggle('active', b.dataset.doc === doc));
    applyFilters();
  }}
</script>
</body>
</html>"""


def generate_styles_thumbnail() -> bool:
    """TEXT_STYLES 一覧スライドを生成してPNG変換する"""
    from lib import SlideBuilder, add_text, add_rect, add_line, TEXT_STYLES, CONTENT_TOP, SLIDE_W, C

    out_pptx = TMP_DIR / "styles_preview.pptx"
    sb = SlideBuilder()
    slide = sb._new_blank_slide()

    # 背景
    add_rect(slide, 0, 0, SLIDE_W, 7.5, fill="bg")

    # ヘッダー
    add_rect(slide, 0, 0, SLIDE_W, 0.6, fill="accent")
    add_text(slide, 0.3, 0.1, SLIDE_W - 0.6, 0.4, "TEXT_STYLES — テキストスタイル一覧",
             style="body", color="white", bold=True)

    style_order = ["heading", "subheading", "body", "small", "caption", "label", "kpi", "title_cover", "subtitle_cover"]
    sample_texts = {
        "heading":       "heading — 見出し 24pt Bold",
        "subheading":    "subheading — サブ見出し 18pt Bold",
        "body":          "body — 本文テキスト 14pt",
        "small":         "small — 補足・表セル 11pt",
        "caption":       "caption — キャプション・フッター 10pt",
        "label":         "label — ラベル 11pt Bold",
        "kpi":           "40",
        "title_cover":   "title_cover 36pt",
        "subtitle_cover":"subtitle_cover 18pt",
    }
    kpi_labels = {"kpi": "kpi — KPI数値 40pt Bold", "title_cover": "", "subtitle_cover": ""}

    y = 0.75
    row_h = 0.62
    for name in style_order:
        s = TEXT_STYLES[name]
        # 区切り線
        add_line(slide, 0.3, y - 0.04, SLIDE_W - 0.3, y - 0.04, color="border", width=0.5)
        # スタイル名バッジ
        add_rect(slide, 0.3, y + 0.05, 1.5, 0.35, fill="bgLight", border="border")
        add_text(slide, 0.32, y + 0.07, 1.46, 0.32, name, style="caption", color="accent", bold=True)
        # サンプルテキスト
        add_text(slide, 1.95, y, SLIDE_W - 2.3, row_h,
                 sample_texts.get(name, name), style=name)
        # サイズ情報
        size_str = f"{s['font_size']}pt {'Bold' if s['bold'] else ''}"
        add_text(slide, SLIDE_W - 1.6, y + 0.1, 1.3, 0.3,
                 size_str, style="caption", align="right", color="textMuted")
        y += row_h

    sb.save(str(out_pptx))
    png = convert_pptx_to_png(out_pptx, IMAGES_DIR)
    if png:
        dest = IMAGES_DIR / "styles_preview.png"
        if png != dest:
            png.rename(dest)
        return True
    return False


def build_styles_html() -> str:
    """テキストスタイル一覧HTMLを生成する"""
    from lib import TEXT_STYLES

    has_img = (IMAGES_DIR / "styles_preview.png").exists()
    img_html = ('<img src="images/styles_preview.png" alt="styles" style="width:100%;border:1px solid #D9D9D9;border-radius:6px">'
                if has_img else '<p style="color:#999;font-size:13px">プレビューなし</p>')

    rows = ""
    for name, s in TEXT_STYLES.items():
        rows += f"""
        <tr>
          <td><code>{name}</code></td>
          <td style="font-size:{s['font_size']}px;font-weight:{'bold' if s['bold'] else 'normal'};line-height:1.2"
              title="{name}">Aa あア</td>
          <td>{s['font_size']}pt</td>
          <td>{"Bold" if s['bold'] else "—"}</td>
          <td><code style="font-size:10px">{s.get('color','text')}</code></td>
        </tr>"""

    return f"""
<div class="section-title">テキストスタイル一覧</div>
<div style="max-width:1400px;margin:0 auto;padding:16px 32px 32px;display:grid;grid-template-columns:1fr 1fr;gap:24px">
  <div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;background:white;border:1px solid #D9D9D9;border-radius:8px;overflow:hidden">
      <thead>
        <tr style="background:#0070C0;color:white">
          <th style="padding:8px 12px;text-align:left">スタイル名</th>
          <th style="padding:8px 12px;text-align:left">プレビュー</th>
          <th style="padding:8px 12px;text-align:left">サイズ</th>
          <th style="padding:8px 12px;text-align:left">太字</th>
          <th style="padding:8px 12px;text-align:left">デフォルト色</th>
        </tr>
      </thead>
      <tbody>{rows}</tbody>
    </table>
  </div>
  <div>{img_html}</div>
</div>"""


def generate_theme_thumbnails():
    """テーマごとに cover_basic のサムネイルを生成する"""
    import importlib.util
    from lib import THEMES

    cover_path = PROJECT_ROOT / "patterns" / "cover_basic.py"
    if not cover_path.exists():
        print("  cover_basic.py が見つかりません。テーマサムネイルをスキップ。")
        return

    spec = importlib.util.spec_from_file_location("cover_basic", cover_path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)

    for name, colors in THEMES.items():
        print(f"\nテーマサムネイル: {name}")
        out_pptx = TMP_DIR / f"theme_{name}.pptx"
        try:
            from lib import SlideBuilder, _C_DEFAULT
            merged = {**_C_DEFAULT, **colors}
            mod.run(
                title=name.upper(),
                subtitle={"default": "企業ブルー", "accenture": "アクセンチュアパープル",
                          "navy": "ネイビー", "green": "グリーン",
                          "warm": "ウォームオレンジ", "mckinsey": "マッキンゼー濃紺"}.get(name, name),
                date="ppt-gen",
                accent=merged["accent"],
                output_path=str(out_pptx),
            )
            png = convert_pptx_to_png(out_pptx, IMAGES_DIR)
            if png:
                # theme_{name}.png にリネーム
                dest = IMAGES_DIR / f"theme_{name}.png"
                if png != dest:
                    png.rename(dest)
                print(f"  PNG: theme_{name}.png")
            else:
                print(f"  PNG: スキップ（LibreOffice未検出）")
        except Exception as e:
            print(f"  ERROR: {e}")


def main():
    CATALOG_DIR.mkdir(exist_ok=True)
    IMAGES_DIR.mkdir(exist_ok=True)
    TMP_DIR.mkdir(exist_ok=True)

    # テキストスタイルサムネイル生成
    print("=== テキストスタイルサムネイル生成 ===")
    ok = generate_styles_thumbnail()
    print(f"  {'styles_preview.png' if ok else 'スキップ（LibreOffice未検出）'}")

    # テーマサムネイル生成
    print("\n=== テーマサムネイル生成 ===")
    generate_theme_thumbnails()

    # パターン処理
    print(f"\n=== パターン処理 ===")
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
    from lib import THEMES
    styles_html = build_styles_html()
    html = build_html(patterns_meta, THEMES, styles_html)
    html_path = CATALOG_DIR / "index.html"
    html_path.write_text(html, encoding="utf-8")
    print(f"\nカタログ生成完了: {html_path}")
    print(f"  テーマ数: {len(THEMES)} / パターン数: {len(patterns_meta)}")
    print(f"  ブラウザで開く: file://{html_path.resolve()}")


if __name__ == "__main__":
    main()
