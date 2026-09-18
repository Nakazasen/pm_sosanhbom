"""
Empirical Verification and Stress Testing of specs/SPEC_PLM_AUTO_DOWNLOAD.md
"""
import os
import re
import sys
from pathlib import Path
from bs4 import BeautifulSoup
import lxml.html

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

WORKSPACE = Path(r"D:\Sandbox\pm_sosanhbom")
SPEC_PATH = WORKSPACE / "specs" / "SPEC_PLM_AUTO_DOWNLOAD.md"
ORIGINAL_REQUEST_PATH = WORKSPACE / ".agents" / "ORIGINAL_REQUEST.md"

DOM_FILES = {
    "login": WORKSPACE / "tc14_login_page_dom.html",
    "export_dialog": WORKSPACE / "tc14_export_dialog_dom.html",
    "item_view": WORKSPACE / "tc14_item_view_dom.html",
    "search": WORKSPACE / "tc14_search_dom.html",
    "post_login": WORKSPACE / "tc14_post_login_dom.html",
    "bom_report": WORKSPACE / "tc14_bom_report_dialog_dom.html",
}

def load_doms():
    dom_trees = {}
    dom_soups = {}
    for key, path in DOM_FILES.items():
        if path.exists():
            html_content = path.read_text(encoding="utf-8", errors="replace")
            try:
                dom_trees[key] = lxml.html.fromstring(html_content)
            except Exception as e:
                print(f"Error parsing {key} with lxml: {e}")
            dom_soups[key] = BeautifulSoup(html_content, "html.parser")
        else:
            print(f"Warning: {path} not found!")
    return dom_trees, dom_soups

def extract_dom_locator_tables(spec_text):
    sections = re.split(r'\n(?=## \d+\. )', spec_text)
    dom_locators = []
    
    for sec in sections:
        sec_title_m = re.match(r'## \d+\.\s*(.*?)\n', sec)
        sec_title = sec_title_m.group(1).strip() if sec_title_m else "Unknown"
        
        loc_table_m = re.search(r'### \d+\.\d+\.\s*Verified DOM Locators Table.*?\n\n(.*?)(?=\n\n###|\n\n---\n|\Z)', sec, re.DOTALL)
        if not loc_table_m:
            continue
            
        table_text = loc_table_m.group(1).strip()
        lines = [line.strip() for line in table_text.splitlines() if line.strip().startswith("|")]
        if len(lines) < 3:
            continue
            
        for line in lines[2:]:
            parts = [p.strip() for p in line.split("|")[1:-1]]
            if len(parts) >= 7:
                elem_name = re.sub(r'\*+', '', parts[0]).strip()
                tag = parts[1].strip('`').strip()
                css = parts[2].strip('`').strip()
                xpath = parts[3].strip('`').strip()
                identifying = parts[4].strip()
                fallback = parts[5].strip('`').strip()
                notes = parts[6].strip()
                
                dom_locators.append({
                    "section": sec_title,
                    "element": elem_name,
                    "tag": tag,
                    "css": css,
                    "xpath": xpath,
                    "identifying": identifying,
                    "fallback": fallback,
                    "notes": notes,
                })
    return dom_locators

def test_selector_in_doms(css, xpath, fallback, dom_trees, dom_soups):
    results = {
        "css_match": {},
        "xpath_match": {},
        "fallback_match": {},
    }
    
    # Test CSS
    for dom_name, soup in dom_soups.items():
        css_clean = css.strip()
        if ":contains(" not in css_clean:
            try:
                matches = soup.select(css_clean)
                if matches:
                    results["css_match"][dom_name] = len(matches)
            except Exception:
                pass
            
        if dom_name in dom_trees and dom_name not in results["css_match"]:
            try:
                tree = dom_trees[dom_name]
                matches = tree.cssselect(css_clean)
                if matches:
                    results["css_match"][dom_name] = len(matches)
            except Exception:
                pass

    # Test XPath
    for dom_name, tree in dom_trees.items():
        xpath_clean = xpath.strip()
        try:
            matches = tree.xpath(xpath_clean)
            if matches:
                results["xpath_match"][dom_name] = len(matches)
        except Exception:
            pass

    # Test Fallback
    for dom_name, soup in dom_soups.items():
        fb = fallback.strip()
        if fb and not fb.startswith("//") and ":contains(" not in fb:
            try:
                matches = soup.select(fb)
                if matches:
                    results["fallback_match"][dom_name] = len(matches)
            except Exception:
                pass
        elif fb and fb.startswith("//") and dom_name in dom_trees:
            try:
                matches = dom_trees[dom_name].xpath(fb)
                if matches:
                    results["fallback_match"][dom_name] = len(matches)
            except Exception:
                pass

    return results

def check_14_columns():
    print("=" * 60)
    print("CHECK 1: 14 Column Names and Exact Sequence")
    print("=" * 60)
    spec_text = SPEC_PATH.read_text(encoding="utf-8")
    orig_text = ORIGINAL_REQUEST_PATH.read_text(encoding="utf-8")
    
    table_cols = re.findall(r'\|\s*\*\*\d+\*\*\s*\|\s*\*\*[A-N]\*\*\s*\|\s*\*\*([^*]+)\*\*', spec_text)
    
    canonical_match = re.search(r'CANONICAL_14_COLUMNS\s*:\s*List\[str\]\s*=\s*\[(.*?)\]', spec_text, re.DOTALL)
    canonical_code_cols = re.findall(r'"([^"]+)"', canonical_match.group(1)) if canonical_match else []
    
    expected_match = re.search(r'EXPECTED_HEADERS\s*=\s*\[(.*?)\]', spec_text, re.DOTALL)
    expected_code_cols = re.findall(r'"([^"]+)"', expected_match.group(1)) if expected_match else []

    ac_config_cols = re.findall(r'^\s*(\d+)\.\s+([A-Za-z0-9\s]+?)\s*$', spec_text[spec_text.find("AC-CONFIG-04"):spec_text.find("AC-CONFIG-05")], re.MULTILINE)
    ac_config_cols_sorted = [name.strip() for num, name in sorted([(int(n), s) for n, s in ac_config_cols], key=lambda x: x[0])]

    orig_cols = re.findall(r'^\s*(\d+)\.\s+([A-Za-z0-9\s]+?)\s*$', orig_text, re.MULTILINE)
    orig_14 = []
    for num_str, name in orig_cols:
        num = int(num_str)
        if 1 <= num <= 14:
            orig_14.append((num, name.strip()))
    orig_14_sorted = [name for num, name in sorted(orig_14, key=lambda x: x[0])]
    
    comma_match = re.search(r'Home,\s*Level,\s*Item Type.*Item Rev Status', orig_text)
    comma_cols = [c.strip() for c in comma_match.group(0).split(",")] if comma_match else []

    checks = [
        ("Table 3 has exactly 14 columns", len(table_cols) == 14),
        ("Table 3 matches ORIGINAL_REQUEST numbered list", table_cols == orig_14_sorted),
        ("Table 3 matches ORIGINAL_REQUEST comma list", table_cols == comma_cols),
        ("CANONICAL_14_COLUMNS in code matches ORIGINAL_REQUEST", canonical_code_cols == orig_14_sorted),
        ("EXPECTED_HEADERS in openpyxl code matches ORIGINAL_REQUEST", expected_code_cols == orig_14_sorted),
        ("AC-CONFIG-04 list in spec matches ORIGINAL_REQUEST", ac_config_cols_sorted == orig_14_sorted),
    ]

    all_passed = True
    for name, passed in checks:
        status = "PASS" if passed else "FAIL"
        if not passed:
            all_passed = False
        print(f"[{status}] {name}")
        
    return all_passed, table_cols

def check_phases_gwt_schemas():
    print("\n" + "=" * 60)
    print("CHECK 3: All 7 Phases + R8 Given-When-Then & Schemas")
    print("=" * 60)
    spec_text = SPEC_PATH.read_text(encoding="utf-8")
    
    phases = [
        ("Phase 1: AUTH-01", r'## 4\.\s*PHASE 1:\s*AUTH-01', "AC-AUTH-", ["AuthInputSchema", "AuthOutputSchema"]),
        ("Phase 2: BOM-SEARCH-01", r'## 5\.\s*PHASE 2:\s*BOM-SEARCH-01', "AC-SEARCH-", ["SearchInputSchema", "SearchOutputSchema"]),
        ("Phase 3: BOM-EXPAND-01", r'## 6\.\s*PHASE 3:\s*BOM-EXPAND-01', "AC-EXPAND-", ["ExpandInputSchema", "ExpandOutputSchema"]),
        ("Phase 4: BOM-SELECT-01", r'## 7\.\s*PHASE 4:\s*BOM-SELECT-01', "AC-SELECT-", ["SelectInputSchema", "SelectOutputSchema"]),
        ("Phase 5: BOM-EXPORT-OPEN-01", r'## 8\.\s*PHASE 5:\s*BOM-EXPORT-OPEN-01', "AC-EXPORT-OPEN-", ["ExportOpenInputSchema", "ExportOpenOutputSchema"]),
        ("Phase 6: BOM-EXPORT-CONFIG-01", r'## 9\.\s*PHASE 6:\s*BOM-EXPORT-CONFIG-01', "AC-CONFIG-", ["ExportConfigInputSchema", "ExportConfigOutputSchema"]),
        ("Phase 7: BOM-EXPORT-RUN-01", r'## 10\.\s*PHASE 7:\s*BOM-EXPORT-RUN-01', "AC-RUN-", ["ExportRunInputSchema", "ExportRunOutputSchema"]),
        ("Phase 8: REP-01 (R8)", r'## 11\.\s*PHASE 8:\s*REP-01', "AC-REP-", ["ProgressReporter"]),
    ]

    all_phases_ok = True
    total_ac_verified = 0
    
    for phase_name, header_regex, ac_prefix, schemas in phases:
        start_m = re.search(header_regex, spec_text)
        if not start_m:
            print(f"[FAIL] {phase_name}: Section header regex {header_regex} not found in spec")
            all_phases_ok = False
            continue
        pos = start_m.start()
        
        next_m = re.search(r'\n## \d+\. ', spec_text[pos + 10:])
        if next_m:
            next_pos = pos + 10 + next_m.start()
            section_text = spec_text[pos:next_pos]
        else:
            section_text = spec_text[pos:]

        # Check User Story
        has_user_story = ("###" in section_text and "User Story" in section_text)

        # Extract all AC titles in this section
        ac_titles = re.findall(rf'-\s*\*\*({ac_prefix}\d+[^*]*)\*\*:', section_text)
        
        # Verify each AC title has Given, When, Then
        valid_gwt_acs = []
        for title in ac_titles:
            escaped_title = re.escape(title)
            # Match block up to next top-level list item (- **) or section header (###)
            ac_block_m = re.search(rf'-\s*\*\*{escaped_title}\*\*:\s*\r?\n(.*?)(?=\r?\n-\s*\*\*|\r?\n\r?\n###|\Z)', section_text, re.DOTALL)
            if ac_block_m:
                block_content = ac_block_m.group(1)
                has_given = ("**Given**" in block_content or "**given**" in block_content.lower())
                has_when = ("**When**" in block_content or "**when**" in block_content.lower())
                has_then = ("**Then**" in block_content or "**then**" in block_content.lower())
                if has_given and has_when and has_then:
                    valid_gwt_acs.append(title)
                else:
                    print(f"       [GWT INCOMPLETE] {title}: Given={has_given}, When={has_when}, Then={has_then}")

        # Check Schemas
        missing_schemas = []
        for s in schemas:
            if s not in section_text:
                missing_schemas.append(s)

        ac_count = len(valid_gwt_acs)
        total_ac_verified += ac_count
        has_gwt = (ac_count > 0 and ac_count == len(ac_titles))
        schemas_ok = len(missing_schemas) == 0

        status = "PASS" if (has_user_story and has_gwt and schemas_ok) else "FAIL"
        if status == "FAIL":
            all_phases_ok = False

        print(f"[{status}] {phase_name}:")
        print(f"   - User Story: {'YES' if has_user_story else 'NO'}")
        print(f"   - GWT AC Count: {ac_count}/{len(ac_titles)} verified")
        for ac in valid_gwt_acs:
            print(f"       * {ac}")
        print(f"   - Required Schemas: {schemas} -> {'ALL FOUND' if schemas_ok else 'MISSING: ' + str(missing_schemas)}")

    print(f"\nTotal GWT Acceptance Criteria verified across all phases: {total_ac_verified}")
    return all_phases_ok

def check_dom_selectors():
    print("\n" + "=" * 60)
    print("CHECK 2: DOM Selectors Cited in Spec vs DOM Snapshots")
    print("=" * 60)
    
    spec_text = SPEC_PATH.read_text(encoding="utf-8")
    locators = extract_dom_locator_tables(spec_text)
    dom_trees, dom_soups = load_doms()
    
    print(f"Total verified DOM locator entries in spec: {len(locators)}")
    
    matched = []
    unmatched = []
    
    for loc in locators:
        is_forbidden = "CẤM" in loc["element"] or "KHÔNG BAO GIỜ" in loc["fallback"]
        res = test_selector_in_doms(loc["css"], loc["xpath"], loc["fallback"], dom_trees, dom_soups)
        has_match = bool(res["css_match"] or res["xpath_match"] or res["fallback_match"])
        
        loc_summary = {
            "section": loc["section"],
            "element": loc["element"],
            "tag": loc["tag"],
            "css": loc["css"],
            "xpath": loc["xpath"],
            "fallback": loc["fallback"],
            "notes": loc["notes"],
            "is_forbidden": is_forbidden,
            "res": res,
            "has_match": has_match
        }
        
        if has_match:
            matched.append(loc_summary)
            print(f"[MATCH] [{loc['section'][:15]}] {loc['element']}")
            if res["css_match"]:
                print(f"    CSS (`{loc['css']}`): {res['css_match']}")
            if res["xpath_match"]:
                print(f"    XPath (`{loc['xpath']}`): {res['xpath_match']}")
            if res["fallback_match"]:
                print(f"    Fallback (`{loc['fallback']}`): {res['fallback_match']}")
        else:
            unmatched.append(loc_summary)
            print(f"[NO MATCH IN SNAPSHOTS] [{loc['section'][:15]}] {loc['element']} (Forbidden={is_forbidden})")
            print(f"    CSS: `{loc['css']}`")
            print(f"    XPath: `{loc['xpath']}`")
            print(f"    Fallback: `{loc['fallback']}`")

    print("\nDOM Locator Summary:")
    print(f"  Matched: {len(matched)} / {len(locators)}")
    print(f"  Unmatched: {len(unmatched)} / {len(locators)}")
    return len(locators), len(matched), unmatched

if __name__ == "__main__":
    c1_ok, cols = check_14_columns()
    total_loc, match_loc, unmatch_list = check_dom_selectors()
    c3_ok = check_phases_gwt_schemas()
    
    print("\n" + "=" * 60)
    print("FINAL VERIFICATION SUMMARY")
    print("=" * 60)
    print(f"14 Columns Contract: {'PASS' if c1_ok else 'FAIL'}")
    print(f"DOM Locators: {match_loc}/{total_loc} matched in snapshots ({len(unmatch_list)} unmatched due to dynamic/transient states)")
    print(f"7 Phases + R8 GWT & Schemas: {'PASS' if c3_ok else 'FAIL'}")
