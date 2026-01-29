# PGT-A Report Generator - Quick Start Guide

## 🚀 Installation (One-Time Setup)

```bash
cd /data/Sethu/PGTA-Report
pip install -r requirements.txt
```

## ▶️ Launch Application

```bash
./launch.sh
```
**OR**
```bash
python3 pgta_report_generator.py
```

## 📝 Generate Your First Report

### Method 1: Manual Entry (Single Patient)

1. **Open "Manual Entry" tab**
2. **Fill patient information:**
   - Patient Name, Spouse Name, PIN, Age
   - Sample Number, Dates, Hospital, etc.
3. **Set embryo count** (use spinner)
4. **For each embryo:**
   - Enter Embryo ID (e.g., PS1, PS2)
   - Fill result details
   - Select chromosome statuses
5. **Click "Save Data"**
6. **Go to "Generate Reports" tab**
7. **Select output folder**
8. **Click "Generate Reports"** 🚀

### Method 2: Bulk Upload (Multiple Patients)

1. **Open "Bulk Upload" tab**
2. **Click "Download Template"** → Save Excel file
3. **Open template in Excel**
4. **Fill in your data** (one row per patient)
5. **Save Excel file**
6. **Click "Browse"** → Select your filled file
7. **Review preview**
8. **Click "Load Data"**
9. **Go to "Generate Reports" tab**
10. **Click "Generate Reports"** 🚀

### Adding CNV Chart Images (Optional)

1. **Open "Image Management" tab**
2. **Click "Add Image(s)"**
3. **Select image files**
4. **Enter Embryo ID** for each image when prompted (e.g., PS1)
5. **Verify** assignments in the table

## 📊 Excel Template Columns

| Column | Example | Required |
|--------|---------|----------|
| Patient_Name | Mrs. Priya (PNM00791) | ✅ |
| Spouse_Name | Mr. Saranraj | ✅ |
| PIN | AND25630004206 | ✅ |
| Age | 34 Years | ✅ |
| Sample_Number | 632504349 | ✅ |
| Referring_Clinician | Dr. Ajantha. B | ✅ |
| Biopsy_Date | 03-01-2026 | ✅ |
| Hospital_Clinic | Rhea Healthcare | ✅ |
| Sample_Collection_Date | 03-01-2026 | ✅ |
| Specimen | Day 6 Trophectoderm Biopsy | ✅ |
| Sample_Receipt_Date | 03-01-2026 | ✅ |
| Biopsy_Performed_By | Dr. Example | ✅ |
| Report_Date | 14-01-2026 | ✅ |
| Indication | History of implantation failure | ✅ |
| Embryo_ID | PS1 | ✅ |
| Result_Summary | Trisomy of chromosome 16 | ✅ |
| Result_Description | Abnormal chromosome complement | ✅ |
| Autosomes | Trisomy of chromosome 16 | ✅ |
| Sex_Chromosomes | Normal | ✅ |
| Interpretation | Aneuploid | ✅ |
| MTcopy | NA | ✅ |

## 🔤 Chromosome Status Codes

| Code | Meaning |
|------|---------|
| **N** | Normal |
| **G** | Gain |
| **L** | Loss |
| **SG** | Segmental Gain |
| **SL** | Segmental Loss |
| **M** | Mosaic |
| **MG** | Mosaic Gain |
| **ML** | Mosaic Loss |
| **SMG** | Segmental Mosaic Gain |
| **SML** | Segmental Mosaic Loss |

## 📋 Interpretation Values

- **Euploid** - Normal chromosome complement
- **Aneuploid** - Abnormal chromosome complement
- **Low level mosaic** - 30-50% mosaicism
- **High level mosaic** - 51-80% mosaicism
- **Complex mosaic** - 3+ chromosomes with mosaic changes

## 📁 Output Files

Reports are saved as:
```
{SampleNumber}_{PatientName}_{Timestamp}.pdf
{SampleNumber}_{PatientName}_{Timestamp}.docx
```

Example:
```
632504349_Mrs_Priya_20260122_104530.pdf
632504349_Mrs_Priya_20260122_104530.docx
```

## ❓ Common Issues

### "No module named 'reportlab'"
```bash
pip install -r requirements_app.txt
```

### "No data loaded"
- Make sure you clicked "Save Data" in Manual Entry tab
- OR loaded data from Bulk Upload tab

### "Please select an output directory"
- Click "Browse" in Generate Reports tab
- Choose where to save reports

### Reports look wrong
- Check that `assets/pgta/` folder exists
- Verify logo images are present

## 📞 Need Help?

1. Check [README.md](file:///data/Sethu/PGTA-Report/README.md) for detailed documentation
2. Review [walkthrough.md](file:///home/nextflowserver/.gemini/antigravity/brain/a29068cc-68fc-487e-8321-31bebc6388b6/walkthrough.md) for complete feature list
3. See [template_specification.md](file:///data/Sethu/PGTA-Report/template_specification.md) for technical details

---

**That's it! You're ready to generate professional PGT-A reports! 🎉**
