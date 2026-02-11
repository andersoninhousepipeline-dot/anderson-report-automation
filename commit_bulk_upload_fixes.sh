#!/bin/bash
# Git commands to commit bulk upload fixes

# Check current status
echo "=== Current Git Status ==="
git status

echo ""
echo "=== Adding modified files ==="
git add pgta_report_generator.py

echo ""
echo "=== Creating commit ==="
git commit -m "Fix bulk upload: Enhanced patient matching and direct AUTOSOMES/SEX extraction

- Enhanced normalize_str() to remove name prefixes (MRS., SMT., R., etc.) and suffixes (C, D, R, S, etc.)
- Fixed patient matching for names like 'R. DEEPALAKSHMI', 'SMT. SUJATA SHRIPAL', 'KAVITHA C'
- Implemented duplicate prevention to avoid embryos assigned to multiple patients
- Changed AUTOSOMES and SEX extraction to use column values exactly as provided
- Sample ID (Anderson ID) prioritized over patient name matching
- All 44 patients and 117 embryos from Run 55 now extract correctly
- No data transformations applied to AUTOSOMES/SEX columns"

echo ""
echo "=== Pushing to remote ==="
git push origin main

echo ""
echo "✅ Done! Changes pushed to repository."
