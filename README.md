# PGT-A Report Generator

**Automated Desktop Application for Generating PGT-A Reports**

A professional desktop application that generates PDF and DOCX reports for Preimplantation Genetic Testing for Aneuploidies (PGT-A) based on a standardized template.

---

## ✨ Features

- **📝 Manual Data Entry**: User-friendly form interface for entering patient and embryo data
- **📊 Bulk Upload**: Import multiple patient records from Excel/TSV files
- **🖼️ Image Management**: Upload and manage CNV chart images for embryos
- **📄 Dual Format Output**: Generate both PDF and DOCX reports simultaneously
- **🔄 Batch Processing**: Generate reports for multiple patients in one go
- **💾 Settings Persistence**: Remembers your last used directories and settings
- **🎨 Professional Templates**: Pixel-perfect recreation of the original PGT-A report design
- **✅ Data Validation**: Built-in validation to ensure data completeness
- **📏 Vertical Spacing Control**: Press **Enter** in any field to add manual line breaks/gaps to the PDF.
- **🖱️ Safe Dropdowns**: Menus ignore mouse wheel to prevent accidental changes.
- **✏️ Editable Interpretations**: Type custom results directly into the "Interpretation" menu.

---

## 📋 Requirements

- **Python 3.10+ (64-bit Recommended)**
- **Operating System**: Windows (10/11), macOS, or Linux
- **Architecture**: 64-bit is highly recommended to avoid PDF rendering issues.
- **Dependencies**: See `requirements.txt` (approx. 150MB total)

---

## 🚀 Quick Start

### For Linux Users
1. **Grant execution permission**:
   ```bash
   chmod +x launch.sh
   ```
2. **Run the application**:
   ```bash
   ./launch.sh
   ```

### For Windows Users
1. **Install Python**: Download from [python.org](https://www.python.org/) (Check "Add to PATH").
2. **Launch**: Double-click **`launch.bat`**.

---

## 🛠️ Detailed Installation (Manual)

### Step 1: Install Python Dependencies
```bash
pip install -r requirements.txt
```

### Step 2: Run the App
```bash
python pgta_report_generator.py
```

---

## 📖 User Guide

### Getting Started

1. **Launch the Application**
   ```bash
   python pgta_report_generator.py
   ```

2. **Choose Your Input Method**
   - **Manual Entry**: For single patient reports
   - **Bulk Upload**: For multiple patients from Excel/TSV

### Tab 1: Manual Entry 📝

**Patient Information Section:**
1. Fill in all required patient fields:
   - Patient Name
   - Spouse Name
   - PIN
   - Age
   - Sample Number
   - Referring Clinician
   - Biopsy Date (DD-MM-YYYY)
   - Hospital/Clinic
   - Sample Collection Date
   - Specimen
   - Sample Receipt Date
   - Biopsy Performed By
   - Report Date
   - Indication

**Embryo Data Section:**
1. Select the number of embryos using the spinner
2. For each embryo, enter:
   - Embryo ID (e.g., PS1, PS2)
   - Result Summary
   - Result Description
   - Autosomes findings
   - Sex Chromosomes status
   - Interpretation (Euploid/Aneuploid/Mosaic)
   - MTcopy value
   - Chromosome statuses (1-22): Select N/G/L/SG/SL/M/MG/ML/SMG/SML for each

3. Click **💾 Save Data** to store your entries

### Tab 2: Bulk Upload 📊

1. Click **📥 Download Template** to get the Excel template
2. Fill in the template with your patient data
3. Click **📁 Browse** to select your filled Excel/TSV file
4. Review the data preview
5. Click **✅ Load Data** to import

**Excel Template Columns:**
- Patient_Name
- Spouse_Name
- PIN
- Age
- Sample_Number
- Referring_Clinician
- Biopsy_Date
- Hospital_Clinic
- Sample_Collection_Date
- Specimen
- Sample_Receipt_Date
- Biopsy_Performed_By
- Report_Date
- Indication
- Embryo_ID
- Result_Summary
- Result_Description
- Autosomes
- Sex_Chromosomes
- Interpretation
- MTcopy

### Tab 3: Image Management 🖼️

1. Click **➕ Add Image(s)** to upload CNV chart images (PNG/JPG/JPEG format)
2. A dialog will appear for each image asking for the **Embryo ID** (e.g., PS1, PS2)
   - The app attempts to auto-suggest the ID if it's in the filename
3. Images are displayed in a table showing the assigned Embryo ID and filename
4. Use the **🗑️** button in the Actions column to remove specific images
5. Use **🗑️ Clear All** to remove all uploaded images

### Tab 4: Generate Reports ⚙️

1. **Select Output Directory**: Click **📁 Browse** to choose where reports will be saved
2. **Choose Formats**: 
   - ✅ Generate PDF
   - ✅ Generate DOCX
3. **Select Branding**:
   - **With Logo**: Includes full header/footer branding.
   - **Without Logo**: Removes main branding but **preserves the GenQA logo**.
4. **Review Data Summary**: Verify patient and embryo count
5. Click **🚀 Generate Reports**
6. Monitor progress bar
7. Open output folder when complete

---

## 📁 Output Files

Reports are automatically named using the format:
```
{Sample_Number}_{Patient_Name}_{Timestamp}.pdf
{Sample_Number}_{Patient_Name}_{Timestamp}.docx
```

Example:
```
632504349_Mrs_Priya_20260122_103045.pdf
632504349_Mrs_Priya_20260122_103045.docx
```

---

## 🎯 Report Structure

### Page 1: Cover Page
- Header logo and branding
- Patient information table
- PNDT disclaimer
- Indication
- Results summary table (all embryos)

### Page 2: Methodology & Static Content
- Methodology description
- Conditions for reporting mosaicism
- Limitations
- References

### Page 3+: Individual Embryo Results
- One page per embryo
- Patient info and PIN
- Embryo details (ID, result, autosomes, sex chromosomes)
- Copy Number Chart (CNV status for chromosomes 1-22)
- Legend
- Signature section

---

## 🔧 Troubleshooting

### Application Won't Start
- Ensure all dependencies are installed: `pip install -r requirements_app.txt`
- Check Python version: `python --version` (should be 3.8+)

### Images Not Appearing in PDF
- Ensure images are in PNG format
- Check that `assets/pgta/` folder exists with logo files
- Verify image paths are correct

### Excel Template Issues
- Use the downloaded template as a starting point
- Ensure date format is DD-MM-YYYY
- Don't modify column headers

### Reports Not Generating
- Check that output directory is writable
- Ensure all required patient fields are filled
- Verify at least one embryo is defined

---

## 🏗️ Project Structure

```
PGTA-Report/
├── pgta_report_generator.py      # Main GUI application
├── pgta_template.py               # PDF template engine
├── pgta_docx_generator.py         # DOCX generator
├── requirements.txt               # Python dependencies
├── template_specification.md      # Template documentation
├── assets/pgta/                   # Logo and branding images
│   ├── image_page1_0.png          # Header logo
│   ├── image_page1_1.png          # Footer banner
│   ├── genqa_logo.png             # GenQA logo
│   └── fonts/                     # Professional fonts (Segoe UI, etc.)
└── README.md                      # This file
```

---

## 🔐 Data Privacy

- All data processing happens **locally** on your computer
- No internet connection required after installation
- No data is sent to external servers
- Reports are saved only to your selected output directory

---

## 📝 Static Content

The following content is **automatically included** in all reports and does not need to be entered:

- **Methodology**: ChromInst® PGT-A description
- **Mosaicism Conditions**: NGS sensitivity and reporting criteria
- **Limitations**: Technical limitations of the assay
- **References**: 7 scientific references
- **Signatures**: Three authorized signatories
  - Anand Babu. K, Ph.D - Molecular Biologist
  - Sachin D Honguntikar, Ph.D - Molecular Geneticist
  - Dr Suriyakumar G - Director

---

## 🎨 Customization

### Updating Logos/Branding
Replace files in `assets/pgta/`:
- `image_page1_0.png` - Header logo (1280x193px recommended)
- `image_page1_1.png` - Footer banner (1299x182px recommended)
- `genqa_logo.png`    - GenQA logo

### Modifying Static Content
Edit the static text in:
- `pgta_template.py` - For PDF reports
- `pgta_docx_generator.py` - For DOCX reports

---

## 🐛 Known Issues

- Large batch uploads (>100 patients) may take several minutes
- Very long patient names may affect table formatting
- Image quality depends on source PNG resolution

---

## 📞 Support

For issues or questions:
1. Check the Troubleshooting section
2. Review the template specification document
3. Verify all dependencies are correctly installed

---

## 📄 License

This software is provided for internal use in PGT-A report generation.

---

## 🙏 Acknowledgments

- **ReportLab** for PDF generation
- **python-docx** for DOCX generation
- **PyQt6** for the GUI framework
- **pandas** for data processing

---

## 🔄 Version History

### Version 1.1.0 (2026-01-31)
- **NEW**: Full multi-line support for all patient details.
- **NEW**: Automatic newline-to-PDF-break conversion (**Enter** key spacing).
- **NEW**: Safe dropdowns (Wheel event ignored to prevent accidental changes).
- **NEW**: Editable interpretation dropdowns with "NA" and "Manual Entry".
- **FIX**: Removed experimental Auto-Shrink in favor of manual spacing control.

### Version 1.0.0 (2026-01-22)
- Initial release
- Manual entry and bulk upload support
- PDF and DOCX generation
- Image management
- Settings persistence
- Batch processing

---

**Built with ❤️ for accurate and professional PGT-A reporting**
