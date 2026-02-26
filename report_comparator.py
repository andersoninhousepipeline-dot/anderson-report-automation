import PyPDF2
import os
import re
import pandas as pd
from datetime import datetime

class PGTAReportComparator:
    def __init__(self, manual_dir=None, automated_dir=None):
        self.manual_dir = manual_dir
        self.automated_dir = automated_dir
        self.manual_files = sorted([f for f in os.listdir(manual_dir) if f.endswith('.pdf')]) if manual_dir and os.path.isdir(manual_dir) else []
        self.automated_files = sorted([f for f in os.listdir(automated_dir) if f.endswith('.pdf')]) if automated_dir and os.path.isdir(automated_dir) else []
        
    def normalize_name(self, name):
        """Normalize patient name for matching."""
        if not name: return ""
        # Remove titles and labels
        name = re.sub(r'^(MRS\.|MR\.|SMT\.|DR\.|MS\.|MISS\.|PATIENT NAME|PATIENT NAME\s*:)\s*', '', name, flags=re.IGNORECASE)
        # Remove hospital codes/extra info in parenthesis
        name = re.sub(r'\(.*?\)', '', name)
        # Remove underscores, dots, and common suffixes
        name = name.replace('_', ' ').replace('.', ' ')
        # Special case for "Mrs._" prefix in automated files
        name = re.sub(r'^MRS\s+', '', name, flags=re.IGNORECASE)
        # Remove "PGT-A report" and "PGTA REPORT" suffix
        name = re.sub(r'PGT-A report.*', '', name, flags=re.IGNORECASE)
        name = re.sub(r'PGTA REPORT.*', '', name, flags=re.IGNORECASE)
        # Remove "withlogo" etc
        name = re.sub(r'withlogo|withoutlogo', '', name, flags=re.IGNORECASE)
        # Cleanup whitespace and non-alphanumeric
        name = re.sub(r'[^A-Z0-9\s]', '', name.upper())
        return ' '.join(name.split())

    def extract_manual_data(self, file_path):
        """Extract data from Manual report format."""
        with open(file_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            text = reader.pages[0].extract_text()
            
        data = {'patient_name': '', 'pin': '', 'sample_number': '', 'indication': '', 'embryos': []}
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        
        for i, line in enumerate(lines):
            if "Patient name" in line:
                name_match = re.search(r'Patient name\s*:\s*(.*?)(?:\s*PIN|$)', line, re.IGNORECASE)
                if name_match: data['patient_name'] = name_match.group(1).strip()
            if "PIN" in line and not data['pin']:
                pin_match = re.search(r'PIN\s*:\s*([A-Z0-9]+)', line)
                if pin_match: data['pin'] = pin_match.group(1)
            if "Sample Number" in line and not data['sample_number']:
                sn_match = re.search(r'Sample Number\s*:\s*([0-9]+)', line)
                if sn_match: data['sample_number'] = sn_match.group(1)
        
        results_start = False
        for line in lines:
            if "Results summary" in line: results_start = True; continue
            if "Methodology" in line or "This test does not reveal" in line: results_start = False; continue
            if results_start:
                match = re.match(r'^(\d+)\s+([A-Z0-9]+(?:\s*\(D\d\))?)\s+(.*?)\s+([\d\.]+|NA)\s+(.*)$', line)
                if match:
                    res, mt, intp = match.group(3).strip(), match.group(4).strip(), match.group(5).strip()
                    if mt.isdigit() and "NA" in intp:
                        res = f"{res} {mt}"; parts = intp.split(); mt = parts[0]; intp = ' '.join(parts[1:])
                    data['embryos'].append({'id': match.group(2).strip(), 'result': res, 'mtcopy': mt, 'interpretation': intp})
        return data

    def extract_automated_data(self, file_path):
        """Extract data from Automated report format."""
        with open(file_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            text = reader.pages[0].extract_text()
            
        data = {'patient_name': '', 'pin': '', 'sample_number': '', 'indication': '', 'embryos': []}
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        
        for i, line in enumerate(lines):
            if line.strip() == "PIN": data['pin'] = lines[i+2].strip()
            if line.strip() == "Patient name": data['patient_name'] = lines[i+2].strip()
            if line.strip() == "Sample Number": data['sample_number'] = lines[i+2].strip()

        results_start = False
        embryo_blocks, current_block = [], []
        for line in lines:
            if "Results summary" in line: results_start = True; continue
            if "Methodology" in line: results_start = False; break
            if results_start:
                if re.match(r'^\d+$', line):
                    if current_block: embryo_blocks.append(current_block)
                    current_block = [line]
                elif current_block:
                    if not line.startswith('[') and not "PNDT act" in line:
                        current_block.append(line)
        if current_block: embryo_blocks.append(current_block)

        for block in embryo_blocks:
            if len(block) >= 5:
                s_id = block[1]
                if len(block) == 5:
                    res, mt, intp = block[2], block[3], block[4]
                else:
                    remaining = ' '.join(block[2:])
                    mt_match = re.search(r'\s+([\d\.]+|NA)\s+([A-Z].*)$', remaining)
                    if mt_match:
                        res, mt, intp = remaining[:mt_match.start()].strip(), mt_match.group(1).strip(), mt_match.group(2).strip()
                    else:
                        res, mt, intp = ' '.join(block[2:-2]), block[-2], block[-1]
                data['embryos'].append({'id': s_id, 'result': res, 'mtcopy': mt, 'interpretation': intp})
        return data

    def compare_embryos(self, me, ae):
        discrepancies = []
        norm_mid = re.sub(r'\(D\d\)', '', me['id']).strip()
        norm_aid = re.sub(r'\(D\d\)', '', ae['id']).strip()
        if norm_mid != norm_aid: discrepancies.append(f"ID Mismatch: Manual({me['id']}) vs Auto({ae['id']})")
        
        norm_mres = ' '.join(re.sub(r'\(D\d\)', '', me['result']).split()).strip().upper()
        norm_ares = ' '.join(re.sub(r'\(D\d\)', '', ae['result']).split()).strip().upper()
        if norm_mres != norm_ares: discrepancies.append(f"Result Mismatch: Manual({me['result']}) vs Auto({ae['result']})")
            
        if me['mtcopy'] != ae['mtcopy']: discrepancies.append(f"MTcopy Mismatch: Manual({me['mtcopy']}) vs Auto({ae['mtcopy']})")
             
        norm_mintp = ' '.join(me['interpretation'].split()).strip().upper()
        norm_aintp = ' '.join(ae['interpretation'].split()).strip().upper()
        if norm_mintp != norm_aintp:
             if not (norm_mintp in norm_aintp or norm_aintp in norm_mintp):
                 discrepancies.append(f"Intp Mismatch: Manual({me['interpretation']}) vs Auto({ae['interpretation']})")
        return discrepancies

    def check_name_match(self, manual_path, auto_path):
        """Validate if two report files belong to the same patient."""
        try:
            m_data = self.extract_manual_data(manual_path)
            a_data = self.extract_automated_data(auto_path)
            
            m_norm = self.normalize_name(m_data['patient_name'])
            a_norm = self.normalize_name(a_data['patient_name'])
            
            # Also consider filename if data extraction fails to find a name properly
            m_file_norm = self.normalize_name(os.path.basename(manual_path))
            a_file_norm = self.normalize_name(os.path.basename(auto_path))
            
            # Check for name in either extracted data or filename
            # CRITICAL: Ensure we don't match against empty strings
            is_match = False
            primary_m = m_norm or m_file_norm
            primary_a = a_norm or a_file_norm
            
            if primary_m and primary_a:
                is_match = (primary_m == primary_a) or (primary_m in primary_a) or (primary_a in primary_m)
            
            return {
                'match': is_match,
                'manual_name': m_data['patient_name'] or os.path.basename(manual_path),
                'auto_name': a_data['patient_name'] or os.path.basename(auto_path),
                'm_norm': m_norm,
                'a_norm': a_norm,
                'is_swapped': "Automated" in manual_path and "Manual" in auto_path
            }
        except Exception as e:
            return {'match': False, 'error': str(e)}

    def compare_single_pair(self, manual_path, auto_path):
        """Compare a specific pair of manual and automated reports."""
        try:
            m_data = self.extract_manual_data(manual_path)
            a_data = self.extract_automated_data(auto_path)
            
            discrepancies = []
            
            # Robustness check: if one file is in the wrong directory, it will likely have no PIN
            if not m_data['pin'] and not a_data['pin']:
                discrepancies.append("CRITICAL: Failed to extract PIN from both reports. Check if files are swapped or invalid.")
            elif not m_data['pin']:
                discrepancies.append("CRITICAL: Manual source file does not appear to be a valid PGT-A manual report.")
            elif not a_data['pin']:
                discrepancies.append("CRITICAL: Automated source file does not appear to be a valid automated report.")

            if m_data['pin'] != a_data['pin']: 
                discrepancies.append(f"PIN Mismatch: Manual({m_data['pin']}) vs Auto({a_data['pin']})")
                
            if m_data['sample_number'] and a_data['sample_number'] and m_data['sample_number'] != a_data['sample_number']:
                discrepancies.append(f"Sample # Mismatch: Manual({m_data['sample_number']}) vs Auto({a_data['sample_number']})")
            
            if len(m_data['embryos']) != len(a_data['embryos']):
                discrepancies.append(f"Embryo Count: Manual({len(m_data['embryos'])}) vs Auto({len(a_data['embryos'])})")
            elif not m_data['embryos']:
                discrepancies.append("Warning: No embryo data extracted from either report.")
            else:
                for i in range(len(m_data['embryos'])):
                    emb_dis = self.compare_embryos(m_data['embryos'][i], a_data['embryos'][i])
                    if emb_dis:
                        discrepancies.append(f"Embryo {i+1} differences:")
                        for d in emb_dis: discrepancies.append(f"  {d}")
                        
            return {
                'patient': m_data['patient_name'] or os.path.basename(manual_path),
                'manual_file': os.path.basename(manual_path),
                'auto_file': os.path.basename(auto_path),
                'discrepancies': discrepancies
            }
        except Exception as e:
            return {
                'patient': os.path.basename(manual_path),
                'manual_file': os.path.basename(manual_path),
                'auto_file': os.path.basename(auto_path),
                'discrepancies': [f"Processing Error: {str(e)}"]
            }

    def compare(self):
        """Compare all files found in manual_dir with matching files in automated_dir."""
        comparison_results = []
        for m_file in self.manual_files:
            if "Investigation" in m_file: continue
            m_path = os.path.join(self.manual_dir, m_file)
            m_data = self.extract_manual_data(m_path)
            norm_m_name = self.normalize_name(m_data['patient_name']) or self.normalize_name(m_file)
            
            matched_auto = next((f for f in self.automated_files if norm_m_name in self.normalize_name(f)), None)
            
            if matched_auto:
                a_path = os.path.join(self.automated_dir, matched_auto)
                comparison_results.append(self.compare_single_pair(m_path, a_path))
            else:
                comparison_results.append({
                    'patient': m_data['patient_name'] or m_file,
                    'manual_file': m_file,
                    'auto_file': 'MISSING',
                    'discrepancies': ['Automated report not found']
                })
        return comparison_results

    def generate_report(self, results):
        report = f"# PGT-A Report Comparison Summary\nGenerated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        perfect = len([r for r in results if not r['discrepancies']])
        report += f"**Total Samples**: {len(results)} | **Matches**: {perfect} | **Discrepancies**: {len(results)-perfect}\n\n---\n\n"
        for res in results:
            report += f"## Patient: {res['patient']}\n- **Manual File**: `{res['manual_file']}`\n- **Automated File**: `{res['auto_file']}`\n"
            if not res['discrepancies']: report += "- ✅ **Status**: Perfect Match\n"
            else:
                report += "- ⚠️ **Discrepancies Found**:\n"
                for d in res['discrepancies']: report += f"  {d}\n"
            report += "\n---\n"
        return report

    def generate_html_report(self, results, output_path):
        """Generate a premium HTML comparison report with detailed stats and modern design."""
        total_samples = len(results)
        if total_samples == 0: return
        
        perfect_matches = len([r for r in results if not r['discrepancies']])
        match_rate = (perfect_matches / total_samples) * 100
        
        # Calculate granular stats
        stats = {
            'pin': {'match': 0, 'total': 0},
            'sample_num': {'match': 0, 'total': 0},
            'embryo_count': {'match': 0, 'total': 0}
        }
        
        for res in results:
            dis = "; ".join(res['discrepancies'])
            
            if "PIN Mismatch" not in dis: stats['pin']['match'] += 1
            stats['pin']['total'] += 1
            
            if "Sample # Mismatch" not in dis: stats['sample_num']['match'] += 1
            stats['sample_num']['total'] += 1
            
            if "Embryo Count" not in dis: stats['embryo_count']['match'] += 1
            stats['embryo_count']['total'] += 1
            
        pin_acc = (stats['pin']['match'] / stats['pin']['total']) * 100 if stats['pin']['total'] > 0 else 0
        sn_acc = (stats['sample_num']['match'] / stats['sample_num']['total']) * 100 if stats['sample_num']['total'] > 0 else 0
        ec_acc = (stats['embryo_count']['match'] / stats['embryo_count']['total']) * 100 if stats['embryo_count']['total'] > 0 else 0
        
        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PGT-A Comparison Performance Dashboard</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {{
            --primary: #2563eb;
            --primary-light: #eff6ff;
            --success: #10b981;
            --success-light: #ecfdf5;
            --warning: #f59e0b;
            --warning-light: #fffbeb;
            --danger: #ef4444;
            --danger-light: #fef2f2;
            --bg: #f3f4f6;
            --card: #ffffff;
            --text-main: #1f2937;
            --text-muted: #6b7280;
            --border: #e5e7eb;
        }}
        * {{ box-sizing: border-box; }}
        body {{
            font-family: 'Inter', system-ui, -apple-system, sans-serif;
            background-color: var(--bg);
            color: var(--text-main);
            margin: 0;
            padding: 40px 20px;
            line-height: 1.5;
        }}
        .container {{
            max-width: 1100px;
            margin: 0 auto;
        }}
        .header {{
            background: white;
            padding: 30px;
            border-radius: 16px;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
            margin-bottom: 30px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-left: 6px solid var(--primary);
        }}
        .header-title h1 {{ margin: 0; font-size: 26px; color: var(--primary); font-weight: 700; }}
        .header-title p {{ margin: 5px 0 0; color: var(--text-muted); font-size: 14px; }}
        
        .concordance-ring {{
            position: relative;
            width: 100px;
            height: 100px;
            display: flex;
            align-items: center;
            justify-content: center;
            background: conic-gradient(var(--success) {match_rate}%, #e5e7eb 0);
            border-radius: 50%;
        }}
        .concordance-ring::after {{
            content: '{match_rate:.0f}%';
            position: absolute;
            width: 80px;
            height: 80px;
            background: white;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 700;
            font-size: 20px;
            color: var(--text-main);
        }}

        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 20px;
            margin-bottom: 40px;
        }}
        .stat-card {{
            background: var(--card);
            padding: 24px;
            border-radius: 16px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            transition: transform 0.2s;
        }}
        .stat-card:hover {{ transform: translateY(-4px); }}
        .stat-label {{ color: var(--text-muted); font-size: 13px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; }}
        .stat-value {{ font-size: 28px; font-weight: 700; margin: 8px 0; }}
        .progress-bar {{
            height: 8px;
            background: #e5e7eb;
            border-radius: 4px;
            overflow: hidden;
            margin-top: 12px;
        }}
        .progress-fill {{ height: 100%; transition: width 1s ease-in-out; }}

        .section-title {{
            font-size: 20px;
            font-weight: 700;
            margin: 40px 0 20px;
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        .section-title::after {{
            content: "";
            flex: 1;
            height: 2px;
            background: var(--border);
        }}

        .card {{
            background: var(--card);
            border-radius: 16px;
            border: 1px solid var(--border);
            margin-bottom: 24px;
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }}
        .card-header {{
            padding: 20px 24px;
            background: #f9fafb;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid var(--border);
        }}
        .card-title {{ font-size: 18px; font-weight: 600; color: var(--text-main); }}
        .card-body {{ padding: 24px; }}
        
        .badge {{
            padding: 6px 14px;
            border-radius: 99px;
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
        }}
        .badge-success {{ background: var(--success-light); color: var(--success); }}
        .badge-danger {{ background: var(--danger-light); color: var(--danger); }}
        
        .file-box {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 15px;
            margin-bottom: 20px;
        }}
        .file-item {{
            background: #f8fafc;
            padding: 12px 16px;
            border-radius: 10px;
            border: 1px dashed #cbd5e1;
            font-size: 13px;
        }}
        .file-item strong {{ display: block; color: var(--text-muted); font-size: 11px; margin-bottom: 4px; text-transform: uppercase; }}
        .file-item span {{ font-family: 'Cascadia Code', monospace; word-break: break-all; }}

        .discrepancy-container {{
            background: var(--danger-light);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid #fee2e2;
        }}
        .discrepancy-title {{
            color: var(--danger);
            font-weight: 700;
            font-size: 14px;
            margin-bottom: 12px;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .discrepancy-list {{
            list-style: none;
            padding: 0;
            margin: 0;
        }}
        .discrepancy-item {{
            padding: 10px 14px;
            background: white;
            border-radius: 8px;
            margin-bottom: 8px;
            font-size: 14px;
            color: #b91c1c;
            border-left: 4px solid var(--danger);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }}
        .diff-marker {{
            font-weight: 700;
            background: #fca5a5;
            padding: 2px 6px;
            border-radius: 4px;
            color: #7f1d1d;
            font-size: 11px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-title">
                <h1>Performance Dashboard</h1>
                <p>PGT-A Automated vs Manual Report Validation</p>
                <p style="font-size: 12px; margin-top: 10px;">Generated on: {datetime.now().strftime('%B %d, %Y at %H:%M:%S')}</p>
            </div>
            <div class="concordance-ring"></div>
        </div>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label">Total Samples</div>
                <div class="stat-value">{total_samples}</div>
                <div style="font-size: 12px; color: var(--text-muted)">Matched pairs verified</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">PIN Accuracy</div>
                <div class="stat-value" style="color: var(--primary)">{pin_acc:.1f}%</div>
                <div class="progress-bar"><div class="progress-fill" style="width: {pin_acc}%; background: var(--primary)"></div></div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Sample # Accuracy</div>
                <div class="stat-value" style="color: var(--warning)">{sn_acc:.1f}%</div>
                <div class="progress-bar"><div class="progress-fill" style="width: {sn_acc}%; background: var(--warning)"></div></div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Embryo Count Consistency</div>
                <div class="stat-value" style="color: var(--success)">{ec_acc:.1f}%</div>
                <div class="progress-bar"><div class="progress-fill" style="width: {ec_acc}%; background: var(--success)"></div></div>
            </div>
        </div>

        <div class="section-title">Validation Details</div>
"""
        
        for res in results:
            is_perfect = not res['discrepancies']
            status_badge = '<span class="badge badge-success">Match verified</span>' if is_perfect else '<span class="badge badge-danger">Discrepancy</span>'
            
            html += f"""
            <div class="card">
                <div class="card-header">
                    <div class="card-title">{res['patient']}</div>
                    {status_badge}
                </div>
                <div class="card-body">
                    <div class="file-box">
                        <div class="file-item">
                            <strong>Manual Report Source</strong>
                            <span>{res['manual_file']}</span>
                        </div>
                        <div class="file-item">
                            <strong>Automated System Output</strong>
                            <span>{res['auto_file']}</span>
                        </div>
                    </div>"""
            
            if not is_perfect:
                html += """
                    <div class="discrepancy-container">
                        <div class="discrepancy-title">⚠️ Discrepancies identified in extraction:</div>
                        <ul class="discrepancy-list">"""
                for d in res['discrepancies']:
                    html += f'<li class="discrepancy-item">{d} <span class="diff-marker">ISSUE</span></li>'
                html += """
                        </ul>
                    </div>"""
            else:
                html += f"""
                    <div style="display: flex; align-items: center; gap: 10px; color: var(--success); font-size: 14px; background: var(--success-light); padding: 15px; border-radius: 10px;">
                        <svg width="20" height="20" viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clip-rule="evenodd"></path></svg>
                        All extraction fields (Patient Info & embryo results) are identical between reports.
                    </div>"""
            
            html += """
                </div>
            </div>"""
            
        html += """
        </div>
    </div>
</body>
</html>"""
        with open(output_path, "w", encoding='utf-8') as f:
            f.write(html)

if __name__ == "__main__":
    comparator = PGTAReportComparator("/data/Sethu/PGTA-Report/Comparison/Manual", "/data/Sethu/PGTA-Report/Comparison/Automated")
    results = comparator.compare()
    comparator.generate_html_report(results, "comparison_report.html")
    open("comparison_results.md", "w").write(comparator.generate_report(results))
    print("Comparison complete. Results saved to comparison_results.md and comparison_report.html")
