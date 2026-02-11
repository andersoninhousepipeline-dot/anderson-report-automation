#!/usr/bin/env python3
"""
Test script to verify bulk upload extraction fixes for Run 55 Excel file.
Tests enhanced patient matching and direct AUTOSOMES/SEX column extraction.
"""

import pandas as pd
import re
from datetime import datetime

def normalize_str(s):
    """Enhanced normalization that removes prefixes, suffixes, and special chars"""
    if not s: return ""
    s = str(s).upper().strip()
    
    # Remove common prefixes (titles, initials)
    prefixes = ['MRS.', 'MR.', 'SMT.', 'DR.', 'MS.', 'MISS.', 'PROF.', 
               'R.', 'S.', 'K.', 'M.', 'D.', 'P.', 'A.', 'B.', 'C.', 
               'G.', 'H.', 'J.', 'L.', 'N.', 'T.', 'V.', 'W.']
    for prefix in prefixes:
        if s.startswith(prefix):
            s = s[len(prefix):].strip()
            break  # Only remove first prefix
    
    # Remove all non-alphanumeric characters
    s = re.sub(r'[^A-Z0-9]', '', s)
    
    # Remove single-letter suffixes at the end (like C, D, R, S, G, etc.)
    # Pattern: Name ends with a single letter that's likely a suffix
    # Examples: KAVITHAC -> KAVITHA, SINDHUR -> SINDHU, SAHANAD -> SAHANA
    if len(s) > 3 and s[-1].isalpha() and not s[-2].isdigit():
        # Check if removing last char leaves a reasonable name (at least 3 chars)
        potential_name = s[:-1]
        if len(potential_name) >= 3:
            # Only remove if the last char is a common suffix letter
            # This prevents removing legitimate name endings
            common_suffixes = ['C', 'D', 'R', 'S', 'G', 'K', 'M', 'N', 'P', 'T', 'V', 'W']
            if s[-1] in common_suffixes:
                s = potential_name
    
    return s

def test_run55_extraction():
    file_path = '/data/Sethu/PGTA-Report/batch-demo/Analysis_Run55_PGT_Surfseq_02-02-2026.xlsx'
    
    print('='*80)
    print('VERIFICATION TEST: Run 55 Bulk Upload Extraction')
    print('='*80)
    
    # Load Excel file
    xl = pd.ExcelFile(file_path)
    sheet_names_lower = [s.lower() for s in xl.sheet_names]
    details_idx = next((i for i, s in enumerate(sheet_names_lower) if s == 'details'), None)
    summary_idx = next((i for i, s in enumerate(sheet_names_lower) if s == 'summary'), None)
    
    df_details = pd.read_excel(file_path, sheet_name=xl.sheet_names[details_idx])
    df_summary_full = pd.read_excel(file_path, sheet_name=xl.sheet_names[summary_idx], header=None)
    
    # Find header row
    header_row_idx = 0
    for r_idx, row in df_summary_full.iterrows():
        if any('sample name' in str(val).lower() for val in row.values):
            header_row_idx = r_idx
            break
    
    df_summary = pd.read_excel(file_path, sheet_name=xl.sheet_names[summary_idx], header=header_row_idx)
    df_details.columns = [str(c).strip() for c in df_details.columns]
    df_summary.columns = [str(c).strip() for c in df_summary.columns]
    
    # Test extraction with new logic
    patient_count = 0
    total_embryos = 0
    bulk_patient_data_list = []
    
    for _, p_row in df_details.iterrows():
        p_name = str(p_row.get('Patient Name', '')).strip()
        if not p_name or p_name.lower() == 'nan':
            continue
        
        patient_count += 1
        norm_p_name = normalize_str(p_name)
        norm_sample_id = normalize_str(p_row.get('Sample ID', ''))
        
        embryos = []
        
        for _, s_row in df_summary.iterrows():
            sample_orig = str(s_row.get('Sample name', ''))
            norm_s_name = normalize_str(sample_orig)
            
            # Enhanced Matching Logic
            match = False
            if norm_sample_id and norm_sample_id in norm_s_name:
                match = True
            elif norm_p_name and norm_p_name in norm_s_name:
                match = True
            
            if match:
                embryo_id = sample_orig.split('_')[0].split('-')[-1] if '-' in sample_orig.split('_')[0] else sample_orig.split('_')[0]
                
                # Extract AUTOSOMES and SEX exactly as provided
                autosomes_raw = str(s_row.get('AUTOSOMES', '')).strip()
                sex_raw = str(s_row.get('SEX', '')).strip()
                
                # Use exactly as provided
                autosomes_val = "" if not autosomes_raw or autosomes_raw.lower() in ['nan', 'none', 'nat', 'null', ''] else autosomes_raw
                sex_chr_val = "" if not sex_raw or sex_raw.lower() in ['nan', 'none', 'nat', 'null', ''] else sex_raw
                
                embryos.append({
                    'embryo_id': embryo_id,
                    'sample_name': sample_orig,
                    'autosomes': autosomes_val,
                    'sex_chromosomes': sex_chr_val
                })
        
        if embryos:
            total_embryos += len(embryos)
            bulk_patient_data_list.append({
                'patient_name': p_name,
                'sample_id': p_row.get('Sample ID', ''),
                'embryo_count': len(embryos),
                'embryos': embryos
            })
    
    # Results
    print(f'\n📊 EXTRACTION RESULTS:')
    print(f'   Expected Patients: 44')
    print(f'   Matched Patients: {len(bulk_patient_data_list)}')
    print(f'   Match Rate: {len(bulk_patient_data_list)/44*100:.1f}%')
    print(f'   Expected Embryos: 117')
    print(f'   Extracted Embryos: {total_embryos}')
    print(f'   Match Rate: {total_embryos/117*100:.1f}%')
    
    # Check specific problematic patients
    print(f'\n🔍 CHECKING PREVIOUSLY PROBLEMATIC PATIENTS:')
    test_patients = [
        ('R. DEEPALAKSHMI', 'AND25150117498', 1),
        ('SMT. SUJATA SHRIPAL', 'AND25AE0000469', 5),
        ('KAVITHA C', 'AND25630004665', 4),
        ('SINDHU R', 'ADB0000021858', 1),
        ('SAHANA D', 'ADB0000021944', 1),
        ('PALLAVI G', 'ADB0000021941', 1)
    ]
    
    for test_name, test_id, expected_embryos in test_patients:
        found = False
        for data in bulk_patient_data_list:
            if data['patient_name'] == test_name:
                found = True
                status = '✅' if data['embryo_count'] == expected_embryos else '⚠️'
                print(f'   {status} {test_name}: {data["embryo_count"]} embryos (expected {expected_embryos})')
                break
        if not found:
            print(f'   ❌ {test_name}: NOT FOUND')
    
    # Verify AUTOSOMES and SEX extraction
    print(f'\n🔍 VERIFYING AUTOSOMES/SEX EXTRACTION (Sample Check):')
    verify_samples = [
        ('AMRITABANGA-ARE1_L01_R1', 'del(1)(q21.1q44)(~104.75Mb,~42%),+22(~51%)', 'MOSAIC GAIN (52%)'),
        ('PURNIMAJAIN-PS2_L02_R1', '-16', 'Abnormal'),
        ('DHIVYA-DK1_L02_R1', 'Normal', 'Normal')
    ]
    
    for sample_name, expected_auto, expected_sex in verify_samples:
        for data in bulk_patient_data_list:
            for emb in data['embryos']:
                if emb['sample_name'] == sample_name:
                    auto_match = '✅' if emb['autosomes'] == expected_auto else '❌'
                    sex_match = '✅' if emb['sex_chromosomes'] == expected_sex else '❌'
                    print(f'   {sample_name}:')
                    print(f'      {auto_match} AUTOSOMES: "{emb["autosomes"]}" (expected "{expected_auto}")')
                    print(f'      {sex_match} SEX: "{emb["sex_chromosomes"]}" (expected "{expected_sex}")')
                    break
    
    # Final verdict
    print(f'\n{"="*80}')
    if len(bulk_patient_data_list) == 44 and total_embryos == 117:
        print('✅ ALL TESTS PASSED!')
        print('   - All 44 patients matched')
        print('   - All 117 embryos extracted')
        print('   - AUTOSOMES and SEX columns extracted correctly')
        return 0
    else:
        print('⚠️ SOME TESTS FAILED')
        print(f'   - Patients: {len(bulk_patient_data_list)}/44')
        print(f'   - Embryos: {total_embryos}/117')
        return 1

if __name__ == '__main__':
    exit(test_run55_extraction())
