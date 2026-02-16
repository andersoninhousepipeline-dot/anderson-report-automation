#!/usr/bin/env python3
"""
Script to remove all TRF-related methods from pgta_report_generator.py
This is a one-time cleanup script.
"""

import re

# Read the file
with open('pgta_report_generator.py', 'r', encoding='utf-8') as f:
    content = f.read()

# List of TRF methods to remove (found via grep)
trf_methods = [
    'extract_text_from_trf',
    'upload_trf_manual',
    'verify_trf_manual',
    'upload_trf_batch',
    'verify_trf_batch',
    'upload_bulk_trf',
    'upload_bulk_trf_single_pdf',
    'upload_bulk_trf_multiple_files',
    'parse_extracted_trf_text',
    'verify_all_bulk_trf',
    'show_bulk_trf_verification_dialog',
    'upload_bulk_trf_pdf_for_dialog',
    'upload_trf_for_selected_patient',
    'show_bulk_trf_comparison_dialog',
    'upload_bulk_trf_dialog',
    'verify_bulk_trf_single_pdf',
    'verify_bulk_trf_multiple_files',
    'get_trf_patients_dict',
    'populate_trf_patient_list',
    'on_trf_patient_selected',
    'verify_selected_patient_trf',
    'compare_trf_to_patient',
    'get_trf_preview_image',
    'apply_trf_value_with_preview',
    'apply_all_trf_values_with_preview',
    'auto_match_bulk_trfs',
]

# Create backup
with open('dev_tools/trf_verification_backup/pgta_report_generator_before_method_removal.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("Backup created: dev_tools/trf_verification_backup/pgta_report_generator_before_method_removal.py")

# For each method, find and remove it
methods_removed = []
for method_name in trf_methods:
    # Pattern to match the method definition and its entire body
    # This regex finds "def method_name" and continues until the next "def " at the same indentation level
    pattern = rf'(\n    def {method_name}\(.*?\):.*?)(?=\n    def |\n\nclass |\Z)'
    
    matches = list(re.finditer(pattern, content, re.DOTALL))
    if matches:
        for match in matches:
            methods_removed.append(method_name)
            # Replace with a stub
            stub = f'\n    def {method_name}(self, *args, **kwargs):\n        """TRF method removed 2026-02-16 - See dev_tools/trf_verification_backup/"""\n        QMessageBox.information(self, "Feature Removed", "TRF Verification has been removed.\\nSee dev_tools/trf_verification_backup/ for restoration.")\n        pass\n'
            content = content[:match.start()] + stub + content[match.end():]

print(f"\nRemoved {len(set(methods_removed))} TRF methods")
print("Methods:", set(methods_removed))

# Write the modified content
with open('pgta_report_generator.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("\nDone! File updated.")
print("Test with: python -m py_compile pgta_report_generator.py")
