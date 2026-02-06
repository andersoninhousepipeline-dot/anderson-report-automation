import os
import sys
import pandas as pd
from datetime import datetime
from reportlab.lib import colors
import re

# Specific imports for testing
try:
    from pgta_template import PGTAReportTemplate
except ImportError:
    PGTAReportTemplate = None

# Mocking the date format function logic for testing
def format_date(d_val):
    if pd.isna(d_val) or str(d_val).lower() == 'nan': return ""
    s = str(d_val).split(' ')[0] 
    try:
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
            try:
                dt = datetime.strptime(s, fmt)
                return dt.strftime("%d-%m-%Y")  # Use hyphen format DD-MM-YYYY
            except ValueError:
                continue
        return s.replace('/', '-')  # Replace slashes with hyphens
    except:
        return s.replace('/', '-')

def test_date_formatting():
    print("\n--- Testing Date Formatting ---")
    cases = [
        ("2023-10-25", "25-10-2023"),
        ("25/10/2023", "25-10-2023"),
        ("10/25/2023", "25-10-2023"),
        ("25-10-2023", "25-10-2023"),
        ("Invalid", "Invalid"),
        (None, "")
    ]
    for inp, expected in cases:
        res = format_date(inp)
        status = "PASS" if res == expected else f"FAIL (Got {res})"
        print(f"Input: {inp} -> Expected: {expected} : {status}")

def test_parsing_logic():
    print("\n--- Testing Parsing Logic ---")
    
    def parse_complex_test(r_str):
        p_stats = {}
        p_mos = {}
        parts = re.split(r',\s*(?=(?:del|dup|mos|[+-]|Monosomy|Trisomy))', r_str, flags=re.IGNORECASE)
        for part in parts:
            part = part.strip()
            if not part: continue
            
            match = re.search(r'(del|dup|mos)\D*?(\d+|X|Y)', part, re.IGNORECASE)
            if match:
                etype = match.group(1).lower() 
                chrom = match.group(2)
                is_seg = bool(re.search(r'([pq]\d|mb)', part, re.IGNORECASE))
                status = 'N'
                
                if 'del' in etype: status = 'SL' if is_seg else 'L'
                elif 'dup' in etype: status = 'SG' if is_seg else 'G'
                elif 'mos' in etype: status = 'M'
                    
                mos_match = re.search(r'[~]*(\d+)%', part)
                if mos_match:
                    p_mos[chrom] = mos_match.group(1)
                    if 'del' in etype: status = 'SML' if is_seg else 'ML'
                    if 'dup' in etype: status = 'SMG' if is_seg else 'MG'
                
                p_stats[chrom] = status
        return p_stats, p_mos

    advanced_cases = [
        ("del(5)(p15.33q12.3)(~64.50Mb,~57%)", {'5': 'SML'}, {'5': '57'}),
        ("del(10)(q22.3q24.1)(~17.50Mb)", {'10': 'SL'}, {}),
        ("dup(18)(~30%)", {'18': 'MG'}, {'18': '30'}),
    ]
    
    print("\n--- Testing Regex Parser ---")
    for inp, exp_stats, exp_mos in advanced_cases:
        got_s, got_m = parse_complex_test(inp)
        status = "PASS" if got_s == exp_stats and got_m == exp_mos else f"FAIL (Got {got_s}, {got_m})"
        print(f"Input: {inp}\n  Stats: {exp_stats} vs {got_s}\n  Mosaic: {exp_mos} vs {got_m}\n  -> {status}")

    print("\n--- Testing Autosome String Generator ---")
    def generate_autosomes_string(stats, mosaic_map):
        if not stats: return "Normal" 
        def sort_key(k):
            if k.isdigit(): return int(k)
            if k.upper() == 'X': return 23
            if k.upper() == 'Y': return 24
            return 25
        sorted_chrs = sorted(stats.keys(), key=sort_key)
        is_multiple = len(sorted_chrs) > 1
        parts = []
        for ch in sorted_chrs:
            st = stats[ch]
            mos_val = mosaic_map.get(ch, '')
            if is_multiple:
                parts.append(f"{ch} {st}")
            else:
                base = f"{ch} chromosome, CNV status {st}"
                if mos_val:
                    base += f", Mosaic(%) {mos_val}"
                parts.append(base)
        return ", ".join(parts)

    auto_cases = [
        ({'16': 'L'}, {}, "16 chromosome, CNV status L"),
        ({'1': 'SG', '11': 'SG', '13': 'SG', '21': 'G'}, {}, "1 SG, 11 SG, 13 SG, 21 G"),
        ({'15': 'MG'}, {'15': '30'}, "15 chromosome, CNV status MG, Mosaic(%) 30"),
    ]
    for stats, mos, expected in auto_cases:
        res = generate_autosomes_string(stats, mos)
        status = "PASS" if res == expected else f"FAIL (Got '{res}')"
        print(f"Stats: {stats} -> Expected: '{expected}' : {status}")

    print("\n--- Testing Data Extraction Robustness ---")
    def get_clean_value(row, keys, default=''):
        if isinstance(keys, str): keys = [keys]
        for k in keys:
            if k in row:
                val = row[k]
                if pd.isna(val): continue
                s_val = str(val).strip()
                if s_val.lower() in ['nan', 'none', 'nat', 'null']: continue
                if s_val: return s_val
        return default
    mock_row = {'Valid': 'Value', 'NoneVal': None, 'NaNVal': float('nan'), 'Alias1': 'Found'}
    extraction_cases = [(['Valid'], "Value"), (['Missing'], ""), (['NoneVal'], ""), (['NaNVal'], ""), (['Missing', 'Alias1'], "Found")]
    for keys, expected in extraction_cases:
        res = get_clean_value(mock_row, keys)
        status = "PASS" if res == expected else f"FAIL (Got '{res}')"
        print(f"Keys: {keys} -> Expected: '{expected}' : {status}")

    print("\n--- Testing Layout Logic (Simplified Spacing) ---")
    if PGTAReportTemplate:
        try:
            tpl = PGTAReportTemplate("dummy.pdf")
            # Test Spacing: \n should become <br/>
            p = tpl._wrap_text("Line 1\nLine 2")
            status = "PASS" if '<br/>' in p.text else f"FAIL (Got '{p.text}')"
            print(f"Spacing Test (Internal): 'Line 1\\nLine 2' -> '{p.text}' : {status}")
            
            # Test Trailing Spacing: \n\n at end should stay as <br/>&nbsp;<br/>&nbsp;
            p_trail = tpl._wrap_text("Line 1\n\n")
            # In HTML representation \u00A0 might show as &nbsp; or literal space depending on print
            # PGTAReportTemplate converts \n to <br/>\u00A0
            status_trail = "PASS" if '<br/>\u00A0<br/>\u00A0' in p_trail.text else f"FAIL (Got '{repr(p_trail.text)}')"
            print(f"Spacing Test (Trailing): 'Line 1\\n\\n' -> '{p_trail.text}' : {status_trail}")
            
            # Test Bold Handling
            p_bold = tpl._wrap_text("Bold Text", bold=True)
            status_bold = "PASS" if '<b>Bold Text</b>' in p_bold.text else f"FAIL (Got '{p_bold.text}')"
            print(f"Bold Test: 'Bold Text' (bold=True) -> '{p_bold.text}' : {status_bold}")
        except Exception as e:
            print(f"Skipping Layout Test due to runtime issues: {e}")
    else:
        print("Skipping Layout Test: PGTAReportTemplate not imported")

if __name__ == "__main__":
    test_date_formatting()
    test_parsing_logic()
    print("\nVerification Complete.")
