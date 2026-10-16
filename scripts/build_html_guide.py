"""Build Interactive Trilingual Standalone HTML User Guide for Kyocera SSBOM Manager.

Generates a single, self-contained HTML file (docs/HUONG_DAN_SU_DUNG.html) with:
- Instant client-side tab switching between Tiếng Việt (VI), 日本語 (JA), and 中文 (ZH)
- Embedded Base64 images for 100% portable offline usage across all factory LAN locations
- Kyocera Corporate UI design system (Kyocera Red #E60012, Slate #1E293B, Royal Blue #2563EB)
- Print-ready CSS for 1-click Export to A4 PDF
"""

from __future__ import annotations

import base64
import json
import html as html_lib
import re
from pathlib import Path
import markdown
from markdown.extensions.toc import slugify_unicode


def image_to_base64(img_path: Path) -> str:
    """Encode image to base64 data URI for 100% portable standalone HTML."""
    if not img_path.exists():
        return ""
    suffix = img_path.suffix.lower().replace(".", "")
    if suffix == "jpg":
        suffix = "jpeg"
    data = base64.b64encode(img_path.read_bytes()).decode("utf-8")
    return f"data:image/{suffix};base64,{data}"


def _plain_text(fragment: str) -> str:
    """Strip tags/entities and normalize whitespace for heading/link text matching."""
    text = re.sub(r"<[^>]+>", "", fragment)
    text = html_lib.unescape(text)
    return re.sub(r"\s+", " ", text).strip().lower()


def convert_md_to_html(md_path: Path, img_cache: dict[str, str], lang: str) -> str:
    """Render markdown to HTML with working, language-unique internal anchors."""
    if not md_path.exists():
        return f"<p>Error: {md_path.name} not found.</p>"

    md_content = md_path.read_text(encoding="utf-8")

    def slugify(value: str, separator: str) -> str:
        return f"{lang}-{slugify_unicode(value, separator)}"

    body = markdown.markdown(
        md_content,
        extensions=["tables", "fenced_code", "toc"],
        extension_configs={"toc": {"slugify": slugify}},
    )

    # Images are embedded once (see IMG map in the page script) instead of once per language.
    body = re.sub(r'<img ([^>]*?)src="images/([^"]+)"', r'<img \1data-img="\2"', body)

    # Map heading text -> generated id
    heading_ids: dict[str, str] = {}
    for hid, inner in re.findall(r'<h[1-6] id="([^"]+)">(.*?)</h[1-6]>', body, flags=re.S):
        heading_ids.setdefault(_plain_text(inner), hid)
    valid_ids = set(heading_ids.values())

    def fix_link(match: re.Match) -> str:
        target, inner = match.group(1), match.group(2)
        new_id = heading_ids.get(_plain_text(inner))
        if not new_id:
            candidate = slugify(html_lib.unescape(target), "-")
            new_id = candidate if candidate in valid_ids else target
        return f'<a href="#{new_id}">{inner}</a>'

    body = re.sub(r'<a href="#([^"]+)">(.*?)</a>', fix_link, body, flags=re.S)

    broken = [h for h in re.findall(r'href="#([^"]+)"', body) if h not in valid_ids]
    print(f"[{lang}] headings={len(valid_ids)} broken_links={len(broken)}")
    return body, build_side_nav(body, lang)


def build_side_nav(body: str, lang: str) -> str:
    """Side menu listing every H2 (and its H3 children) with in-page anchors."""
    items: list[str] = []
    open_sub = False
    for level, hid, inner in re.findall(r'<h([23]) id="([^"]+)">(.*?)</h[23]>', body, flags=re.S):
        text = html_lib.escape(html_lib.unescape(re.sub(r"<[^>]+>", "", inner)).strip())
        if level == "2":
            if open_sub:
                items.append("</ul></li>")
                open_sub = False
            items.append(f'<li class="l2"><a href="#{hid}" data-target="{hid}">{text}</a>')
            items.append("<ul>")
            open_sub = True
        else:
            items.append(f'<li class="l3"><a href="#{hid}" data-target="{hid}">{text}</a></li>')
    if open_sub:
        items.append("</ul></li>")
    nav = "\n".join(items).replace("<ul>\n</ul>", "")
    return f'<ul class="side-menu" data-lang="{lang}">\n{nav}\n</ul>'


def main() -> None:
    doc_dir = Path("docs")
    vi_md = doc_dir / "HUONG_DAN_SU_DUNG.md"
    ja_md = doc_dir / "HUONG_DAN_SU_DUNG_JA.md"
    zh_md = doc_dir / "HUONG_DAN_SU_DUNG_ZH.md"
    html_file = doc_dir / "HUONG_DAN_SU_DUNG.html"

    # Pre-cache all images into base64
    img_dir = doc_dir / "images"
    img_cache: dict[str, str] = {}
    if img_dir.exists():
        for p in sorted(img_dir.glob("*.png")):
            img_cache[p.name] = image_to_base64(p)
    print(f"[+] Cached {len(img_cache)} images as Base64 data URIs.")

    html_vi, nav_vi = convert_md_to_html(vi_md, img_cache, "vi")
    html_ja, nav_ja = convert_md_to_html(ja_md, img_cache, "ja")
    html_zh, nav_zh = convert_md_to_html(zh_md, img_cache, "zh")

    img_js = "{" + ",".join(f'{json.dumps(k)}:{json.dumps(v)}' for k, v in img_cache.items() if v) + "}"

    full_html = f"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sổ Tay Hướng Dẫn Sử Dụng - Kyocera SSBOM Manager (VI / JA / ZH)</title>
    <style>
        :root {{
            --primary: #E60012;
            --primary-dark: #B3000E;
            --secondary: #0284C7;
            --dark: #0F172A;
            --slate: #1E293B;
            --light: #F8FAFC;
            --border: #E2E8F0;
            --card-bg: #FFFFFF;
            --text-main: #334155;
            --text-muted: #64748B;
            --accent-green: #15803D;
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }}

        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Hiragino Kaku Gothic ProN", "Yu Gothic", "PingFang SC", "Microsoft YaHei", sans-serif;
            background-color: var(--light);
            color: var(--text-main);
            line-height: 1.68;
            min-height: 100vh;
        }}

        /* Sticky Top Navigation Bar */
        .top-navbar {{
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            height: 64px;
            background: linear-gradient(90deg, #1E293B 0%, #0F172A 100%);
            border-bottom: 3px solid var(--primary);
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 24px;
            z-index: 1000;
            color: #FFFFFF;
            box-shadow: 0 4px 14px rgba(0,0,0,0.15);
        }}

        .top-navbar .logo-area {{
            display: flex;
            align-items: center;
            gap: 12px;
            font-weight: 700;
            font-size: 17px;
            letter-spacing: 0.3px;
        }}

        .top-navbar .logo-badge {{
            background-color: var(--primary);
            color: white;
            padding: 5px 11px;
            border-radius: 4px;
            font-size: 13px;
            font-weight: 900;
            letter-spacing: 1px;
        }}

        /* Interactive Language Switcher Bar */
        .lang-switch-container {{
            display: flex;
            align-items: center;
            gap: 6px;
            background: rgba(255, 255, 255, 0.1);
            padding: 4px;
            border-radius: 8px;
            border: 1px solid rgba(255, 255, 255, 0.15);
        }}

        .lang-btn {{
            background: transparent;
            color: #CBD5E1;
            border: none;
            padding: 6px 14px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 13px;
            font-weight: 700;
            transition: all 0.2s ease-in-out;
            display: flex;
            align-items: center;
            gap: 6px;
        }}

        .lang-btn:hover {{
            background: rgba(255, 255, 255, 0.15);
            color: #FFFFFF;
        }}

        .lang-btn.active {{
            background: var(--primary);
            color: #FFFFFF;
            box-shadow: 0 2px 8px rgba(230, 0, 18, 0.4);
        }}

        .top-actions {{
            display: flex;
            align-items: center;
            gap: 14px;
        }}

        .btn-print {{
            background-color: #334155;
            color: white;
            border: 1px solid #475569;
            padding: 7px 16px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 13px;
            font-weight: 600;
            transition: all 0.2s;
            display: flex;
            align-items: center;
            gap: 6px;
        }}

        .btn-print:hover {{
            background-color: var(--primary);
            border-color: var(--primary);
        }}

        /* Main Container */
        .wrapper {{
            max-width: 1480px;
            margin: 80px auto 40px auto;
            padding: 0 20px;
        }}

        .content {{
            background: var(--card-bg);
            box-shadow: 0 4px 20px rgba(0,0,0,0.04);
            border-radius: 10px;
            padding: 44px 52px;
            border: 1px solid var(--border);
        }}

        /* Typography */
        h1 {{
            color: var(--slate);
            font-size: 26px;
            font-weight: 800;
            border-bottom: 2px solid var(--border);
            padding-bottom: 12px;
            margin-bottom: 20px;
            line-height: 1.3;
        }}

        h2 {{
            color: var(--dark);
            font-size: 21px;
            font-weight: 700;
            margin-top: 36px;
            margin-bottom: 16px;
            border-left: 4px solid var(--primary);
            padding-left: 12px;
        }}

        h3 {{
            color: var(--slate);
            font-size: 17px;
            font-weight: 600;
            margin-top: 24px;
            margin-bottom: 12px;
        }}

        h4 {{
            font-size: 15px;
            color: var(--slate);
            margin-top: 18px;
            margin-bottom: 8px;
        }}

        p {{
            margin-bottom: 14px;
        }}

        blockquote {{
            background-color: #F8FAFC;
            border-left: 4px solid var(--secondary);
            padding: 12px 18px;
            margin-bottom: 20px;
            border-radius: 0 6px 6px 0;
            font-size: 14px;
            color: var(--slate);
        }}

        ul, ol {{
            margin-left: 24px;
            margin-bottom: 16px;
        }}

        li {{
            margin-bottom: 6px;
        }}

        code {{
            background-color: #F1F5F9;
            color: #BE185D;
            padding: 2px 6px;
            border-radius: 4px;
            font-family: Consolas, Monaco, monospace;
            font-size: 13px;
            border: 1px solid #E2E8F0;
        }}

        pre {{
            background-color: #0F172A;
            color: #E2E8F0;
            padding: 16px;
            border-radius: 8px;
            overflow-x: auto;
            margin-bottom: 20px;
            font-family: Consolas, Monaco, monospace;
            font-size: 13px;
        }}

        pre code {{
            background: transparent;
            color: inherit;
            border: none;
            padding: 0;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            font-size: 13.5px;
            background: white;
            border-radius: 6px;
            overflow: hidden;
            border: 1px solid var(--border);
        }}

        th {{
            background-color: #F8FAFC;
            color: var(--slate);
            font-weight: 700;
            text-align: left;
            padding: 11px 15px;
            border-bottom: 2px solid var(--border);
            font-size: 13px;
        }}

        td {{
            padding: 10px 15px;
            border-bottom: 1px solid var(--border);
            vertical-align: middle;
        }}

        tr:nth-child(even) {{
            background-color: #FAFAFA;
        }}

        tr:hover {{
            background-color: #F1F5F9;
        }}

        img {{
            max-width: 100%;
            height: auto;
            border-radius: 8px;
            border: 1px solid #CBD5E1;
            box-shadow: 0 4px 16px rgba(0,0,0,0.08);
            margin: 16px 0 24px 0;
            display: block;
        }}

        hr {{
            border: none;
            border-top: 1px solid var(--border);
            margin: 32px 0;
        }}

        .lang-section {{
            display: none;
        }}

        html {{
            scroll-behavior: smooth;
        }}

        h1[id], h2[id], h3[id], h4[id] {{
            scroll-margin-top: 84px;
        }}

        .lang-section.active {{
            display: block;
            animation: fadeIn 0.25s ease-in-out;
        }}

        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(4px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        /* Side menu (jump to any section) */
        .layout {{
            display: grid;
            grid-template-columns: 290px minmax(0, 1fr);
            gap: 36px;
            align-items: start;
        }}
        .side-nav {{
            position: sticky;
            top: 84px;
            max-height: calc(100vh - 100px);
            overflow-y: auto;
            border-right: 1px solid var(--border);
            padding-right: 14px;
            font-size: 13.5px;
        }}
        .side-menu, .side-menu ul {{ list-style: none; margin: 0; padding: 0; }}
        .side-menu li {{ margin: 0; }}
        .side-menu a {{
            display: block;
            padding: 5px 10px;
            color: var(--text-main);
            text-decoration: none;
            border-left: 3px solid transparent;
            border-radius: 0 4px 4px 0;
            line-height: 1.4;
        }}
        .side-menu a:hover {{ background: #F1F5F9; color: var(--primary-dark); }}
        .side-menu .l2 > a {{ font-weight: 700; color: var(--dark); margin-top: 8px; }}
        .side-menu .l3 > a {{ padding-left: 24px; font-weight: 400; }}
        .side-menu a.current {{ border-left-color: var(--primary); background: #FEF2F2; color: var(--primary-dark); }}
        @media (max-width: 960px) {{
            .layout {{ grid-template-columns: 1fr; }}
            .side-nav {{ position: static; max-height: none; border-right: none; border-bottom: 1px solid var(--border); padding-bottom: 12px; }}
        }}
        /* Print Media */
        @media print {{
            .top-navbar, .side-nav {{
                display: none !important;
            }}
            .layout {{ display: block !important; }}
            .wrapper {{
                margin-top: 0 !important;
                padding: 0 !important;
                max-width: 100% !important;
            }}
            .content {{
                border: none !important;
                box-shadow: none !important;
                padding: 0 !important;
            }}
            img {{
                page-break-inside: avoid;
                max-width: 95% !important;
            }}
            h2, h3 {{
                page-break-after: avoid;
            }}
        }}
    </style>
</head>
<body>
    <div class="top-navbar">
        <div class="logo-area">
            <span class="logo-badge">KYOCERA</span>
            <span id="nav-title">SSBOM MANAGER - SỔ TAY HƯỚNG DẪN SỬ DỤNG</span>
        </div>

        <div class="lang-switch-container">
            <button class="lang-btn active" id="btn-vi" onclick="switchLanguage('vi')">🇻🇳 Tiếng Việt</button>
            <button class="lang-btn" id="btn-ja" onclick="switchLanguage('ja')">🇯🇵 日本語</button>
            <button class="lang-btn" id="btn-zh" onclick="switchLanguage('zh')">🇨🇳 中文</button>
        </div>

        <div class="top-actions">
            <button class="btn-print" onclick="window.print()">🖨️ <span id="lbl-print">In PDF (Ctrl+P)</span></button>
        </div>
    </div>

    <div class="wrapper">
        <div class="content">
            <!-- Vietnamese Section -->
            <div id="section-vi" class="lang-section active">
                <div class="layout"><nav class="side-nav">{nav_vi}</nav><div class="doc">{html_vi}</div></div>
            </div>

            <!-- Japanese Section -->
            <div id="section-ja" class="lang-section">
                <div class="layout"><nav class="side-nav">{nav_ja}</nav><div class="doc">{html_ja}</div></div>
            </div>

            <!-- Chinese Section -->
            <div id="section-zh" class="lang-section">
                <div class="layout"><nav class="side-nav">{nav_zh}</nav><div class="doc">{html_zh}</div></div>
            </div>
        </div>
    </div>

    <script>
        const titles = {{
            'vi': 'SSBOM MANAGER - SỔ TAY HƯỚNG DẪN SỬ DỤNG',
            'ja': 'SSBOM MANAGER - 自動BOM照合システム 取扱説明書',
            'zh': 'SSBOM MANAGER - 自动BOM比对系统 操作手册'
        }};

        const printLabels = {{
            'vi': 'In PDF (Ctrl+P)',
            'ja': 'PDF印刷 (Ctrl+P)',
            'zh': '打印PDF (Ctrl+P)'
        }};

        function switchLanguage(lang) {{
            document.querySelectorAll('.lang-section').forEach(el => el.classList.remove('active'));
            document.querySelectorAll('.lang-btn').forEach(el => el.classList.remove('active'));

            const targetSection = document.getElementById('section-' + lang);
            const targetBtn = document.getElementById('btn-' + lang);

            if (targetSection) targetSection.classList.add('active');
            if (targetBtn) targetBtn.classList.add('active');

            if (titles[lang]) {{
                document.getElementById('nav-title').textContent = titles[lang];
                document.title = titles[lang] + ' - Kyocera SSBOM Manager';
            }}
            if (printLabels[lang]) {{
                document.getElementById('lbl-print').textContent = printLabels[lang];
            }}

            localStorage.setItem('ssbom_guide_lang', lang);
            window.scrollTo({{ top: 0, behavior: 'smooth' }});
            initScrollSpy();
        }}

        // Each screenshot is stored once and attached to every <img data-img=...>
        const IMG = {img_js};
        document.querySelectorAll('img[data-img]').forEach(el => {{ const d = IMG[el.dataset.img]; if (d) el.src = d; }});

        // Highlight the menu entry of the section currently on screen
        function initScrollSpy() {{
            const links = Array.from(document.querySelectorAll('.lang-section.active .side-menu a'));
            const targets = links.map(a => document.getElementById(a.dataset.target)).filter(Boolean);
            const onScroll = () => {{
                let cur = targets.length ? targets[0].id : null;
                for (const el of targets) {{
                    if (el.getBoundingClientRect().top <= 110) cur = el.id;
                }}
                if (targets.length && (window.innerHeight + window.scrollY) >= document.body.scrollHeight - 4) cur = targets[targets.length - 1].id;
                links.forEach(a => a.classList.toggle('current', a.dataset.target === cur));
            }};
            window.removeEventListener('scroll', window._ssbomSpy || (() => {{}}));
            window._ssbomSpy = onScroll;
            window.addEventListener('scroll', onScroll, {{ passive: true }});
            onScroll();
        }}
        // Restore language preference (URL hash prefix wins, e.g. #ja-...)
        document.addEventListener('DOMContentLoaded', () => {{
            const hashLang = (location.hash.match(/^#(vi|ja|zh)-/) || [])[1];
            const savedLang = hashLang || localStorage.getItem('ssbom_guide_lang');
            if (savedLang && ['vi', 'ja', 'zh'].includes(savedLang)) {{
                switchLanguage(savedLang);
            }}
            if (hashLang) {{
                const el = document.getElementById(decodeURIComponent(location.hash.slice(1)));
                if (el) el.scrollIntoView();
            }}
        }});
    </script>
</body>
</html>
"""

    html_file.write_text(full_html, encoding="utf-8")
    file_size_mb = html_file.stat().st_size / (1024 * 1024)
    print(f"\n[SUCCESS] Generated Standalone Trilingual HTML Guide: {html_file.resolve()}")
    print(f"[+] Total Size: {file_size_mb:.2f} MB (screenshots embedded as Base64)")


if __name__ == "__main__":
    main()
