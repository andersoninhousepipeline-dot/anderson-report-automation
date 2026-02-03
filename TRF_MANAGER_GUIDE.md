# TRF Manager - Complete Setup & User Guide

> **Test Request Form (TRF) Verification System**  
> Automatically extract and verify patient information from TRF documents against entered data.

---

## Table of Contents

1. [Overview](#overview)
2. [Installation & Setup](#installation--setup)
   - [Basic Requirements](#basic-requirements)
   - [EasyOCR Setup (Recommended)](#easyocr-setup-recommended)
   - [Ollama LLaVA Setup (Best Accuracy)](#ollama-llava-setup-best-accuracy)
   - [Tesseract Setup (Fallback)](#tesseract-setup-fallback)
3. [Using TRF Manager](#using-trf-manager)
   - [Opening TRF Manager](#opening-trf-manager)
   - [Single TRF Verification](#single-trf-verification)
   - [Bulk TRF Verification](#bulk-trf-verification)
4. [Workflow Examples](#workflow-examples)
5. [Troubleshooting](#troubleshooting)

---

## Overview

The TRF Manager allows you to:

- ✅ **Upload TRF documents** (PDF or images) for each patient
- ✅ **Extract text automatically** using OCR/AI
- ✅ **Compare extracted data** with entered patient information
- ✅ **Apply corrections** from TRF to patient data with one click
- ✅ **Bulk process** a single multi-page PDF containing all TRFs
- ✅ **Auto-match** TRF pages to patients by name/PIN

### Supported File Formats

| Type | Extensions |
|------|------------|
| PDF | `.pdf` |
| Images | `.png`, `.jpg`, `.jpeg`, `.tiff`, `.bmp` |

### Extraction Methods Comparison

| Method | Accuracy | Speed | Offline | Setup Difficulty |
|--------|----------|-------|---------|------------------|
| **EasyOCR** | Good | Medium | ✅ Yes | Easy |
| **Ollama LLaVA** | Excellent | Slow | ✅ Yes | Medium |
| **Tesseract** | Basic | Fast | ✅ Yes | Easy |

---

## Installation & Setup

### Basic Requirements

First, ensure you have the core dependencies installed:

```bash
# Navigate to the project directory
cd /data/Sethu/PGTA-Report

# Install base requirements
pip install -r requirements.txt
```

The `requirements.txt` includes:
- `pdfplumber>=0.10.0` - PDF text extraction
- `Pillow>=10.0.0` - Image processing
- `requests>=2.31.0` - For Ollama API calls

---

### EasyOCR Setup (Recommended)

EasyOCR is the **recommended** method - it works offline and provides good accuracy.

#### Step 1: Install EasyOCR

```bash
pip install easyocr>=1.7.0
```

#### Step 2: First Run (Model Download)

On first use, EasyOCR will download language models (~100MB for English):

```bash
# Test EasyOCR installation
python -c "import easyocr; reader = easyocr.Reader(['en']); print('EasyOCR ready!')"
```

> ⏳ **Note**: The first run takes a few minutes to download models. Subsequent runs are fast.

#### Step 3: Verify Installation

```bash
python -c "
import easyocr
reader = easyocr.Reader(['en'], gpu=False)
print('✅ EasyOCR installed and working!')
"
```

#### GPU Acceleration (Optional)

For faster processing with NVIDIA GPU:

```bash
# Install PyTorch with CUDA support
pip install torch torchvision --index-url https://download.pytorch.org/whl/cu118
```

Then EasyOCR will automatically use GPU.

---

### Ollama LLaVA Setup (Best Accuracy)

Ollama with LLaVA provides the **best accuracy** using a local AI vision model. No API keys needed!

#### Step 1: Install Ollama

**Linux:**
```bash
curl -fsSL https://ollama.com/install.sh | sh
```

**Windows:**
Download from: https://ollama.com/download/windows

**macOS:**
```bash
brew install ollama
```

#### Step 2: Start Ollama Server

```bash
# Start Ollama service
ollama serve
```

Keep this terminal open, or run as a background service.

#### Step 3: Download LLaVA Model

```bash
# Download the LLaVA vision model (~4GB)
ollama pull llava
```

#### Step 4: Verify Installation

```bash
# Test Ollama is running
curl http://localhost:11434/api/tags

# Or test with Python
python -c "
import requests
r = requests.get('http://localhost:11434/api/tags')
print('✅ Ollama running!' if r.status_code == 200 else '❌ Ollama not running')
"
```

#### Step 5: Test LLaVA

```bash
# Quick test
ollama run llava "What can you see in this image?" --verbose
```

> 💡 **Tip**: LLaVA requires ~8GB RAM minimum. For better performance, use a machine with 16GB+ RAM.

---

### Tesseract Setup (Fallback)

Tesseract is a basic OCR option that's fast but less accurate.

#### Linux (Ubuntu/Debian):

```bash
sudo apt update
sudo apt install tesseract-ocr tesseract-ocr-eng
pip install pytesseract
```

#### Windows:

1. Download installer from: https://github.com/UB-Mannheim/tesseract/wiki
2. Install to default location: `C:\Program Files\Tesseract-OCR`
3. Add to PATH or set in Python:

```python
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

4. Install Python wrapper:
```bash
pip install pytesseract
```

#### macOS:

```bash
brew install tesseract
pip install pytesseract
```

#### Verify Installation:

```bash
tesseract --version
python -c "import pytesseract; print('✅ Tesseract ready!')"
```

---

## Using TRF Manager

### Opening TRF Manager

1. **Load patient data** first (either via Manual Entry or Bulk Excel Import)
2. Go to **Bulk Editor** tab
3. Click **"📋 TRF Manager"** button in the toolbar

![TRF Manager Button Location](assets/pgta/trf_manager_btn.png)

### TRF Manager Interface

```
┌─────────────────────────────────────────────────────────────────┐
│  TRF Verification Manager                                    ✕  │
├─────────────────────────────────────────────────────────────────┤
│  ⚙️ Settings                                                    │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │ Extraction Method: [EasyOCR (Recommended) ▼]              │  │
│  │ ☑ Use AI Enhancement (Ollama LLaVA)                       │  │
│  │ [🔗 Test Ollama Connection]                               │  │
│  └───────────────────────────────────────────────────────────┘  │
│                                                                 │
│  Patient List              │  TRF Details                       │
│  ┌───────────────────────┐ │ ┌─────────────────────────────────┐│
│  │ Name    │ PIN  │ TRF  │ │ │                                 ││
│  ├─────────┼──────┼──────┤ │ │  Patient: John Doe              ││
│  │ John Doe│ P001 │  ✅  │ │ │  PIN: P001                      ││
│  │ Jane Doe│ P002 │  ❌  │ │ │  Hospital: ABC Clinic           ││
│  │ Bob Smith│P003 │  ✅  │ │ │                                 ││
│  └───────────────────────┘ │ │  ✅ TRF Linked: trf_001.pdf     ││
│                            │ │  Match Score: 95%               ││
│  [📄 Upload TRF]           │ │                                 ││
│  [📑 Upload Bulk TRF]      │ └─────────────────────────────────┘│
│  [✓ Verify Selected]       │                                    │
│  [✓ Verify All]            │                                    │
├─────────────────────────────────────────────────────────────────┤
│                                                    [Close]      │
└─────────────────────────────────────────────────────────────────┘
```

---

### Single TRF Verification

For verifying individual patient TRFs one at a time:

#### Step 1: Select Patient
Click on a patient in the list to select them.

#### Step 2: Upload TRF
Click **"📄 Upload TRF"** and select the TRF file (PDF or image).

#### Step 3: Verify
Click **"✓ Verify Selected"** to extract and compare data.

#### Step 4: Review & Apply
A comparison dialog appears:

```
┌─────────────────────────────────────────────────────────────────┐
│  TRF Verification - John Doe (P001)                          ✕  │
├─────────────────────────────────────────────────────────────────┤
│  ✅ All fields verified!                                        │
├─────────────────────────────────────────────────────────────────┤
│  Field          │ Current Value │ TRF Value    │ Status │Action │
│  ───────────────┼───────────────┼──────────────┼────────┼───────│
│  Patient Name   │ John Doe      │ John Doe     │ ✓ Match│       │
│  Hospital       │ ABC Clinic    │ ABC Hospital │ ⚠ Diff │[Apply]│
│  PIN            │ P001          │ P001         │ ✓ Match│       │
│  Biopsy Date    │ 2026-01-15    │ 15/01/2026   │ ✓ Match│       │
├─────────────────────────────────────────────────────────────────┤
│  [✓ Apply All (1)]                                    [Close]   │
└─────────────────────────────────────────────────────────────────┘
```

- **Green rows**: Data matches
- **Red rows**: Data differs - click **[Apply]** to use TRF value
- **Blue rows**: Data found in TRF but empty in current record

---

### Bulk TRF Verification

For processing a **single multi-page PDF** containing all TRFs:

#### Step 1: Upload Bulk TRF

Click **"📑 Upload Bulk TRF"** and choose:

```
┌─────────────────────────────────────────┐
│  Select Bulk TRF Upload Type            │
│                                         │
│  [📑 Single Multi-Page PDF]             │
│  One PDF containing all TRFs            │
│                                         │
│  [📁 Multiple Individual Files]         │
│  Separate files for each patient        │
│                                         │
│                           [Cancel]      │
└─────────────────────────────────────────┘
```

#### Step 2: Select PDF

Choose your bulk TRF PDF file. The system will:
1. Count all pages
2. Show: "Loaded bulk TRF with X pages"

#### Step 3: Auto-Match (Verify All)

Click **"✓ Verify All"** to automatically:
1. Extract text from each page
2. Match pages to patients by name/PIN
3. Link matched pages to patients

Progress is shown:
```
Processing page 1 of 50...
Processing page 2 of 50...
...
✅ Matched 45 of 50 pages to patients
```

#### Step 4: Manual Assignment (If Needed)

For unmatched pages, select the patient and click **"📄 Upload TRF"**:

```
┌─────────────────────────────────────────┐
│  TRF Source                             │
│                                         │
│  [📑 Select Page from Bulk PDF]         │
│  [📄 Upload Individual File]            │
│                           [Cancel]      │
└─────────────────────────────────────────┘
```

Choose **"Select Page from Bulk PDF"** to see a page selector:

```
┌─────────────────────────────────────────────────┐
│  Assign Page to Jane Doe (P002)              ✕  │
├─────────────────────────────────────────────────┤
│  Select page from: bulk_trf.pdf                 │
│  Total pages: 50                                │
│                                                 │
│  Page: [  3  ] [▲] [▼]                          │
│                                                 │
│  Preview:                                       │
│  ┌─────────────────────────────────────────┐    │
│  │ Patient Name: Jane Doe                  │    │
│  │ PIN: P002                               │    │
│  │ Hospital: XYZ Medical Center            │    │
│  │ Biopsy Date: 20/01/2026                 │    │
│  │ ...                                     │    │
│  └─────────────────────────────────────────┘    │
│                                                 │
│                    [Cancel]  [✓ Assign Page]    │
└─────────────────────────────────────────────────┘
```

#### Step 5: Verify Individual Patients

After matching, select each patient and click **"✓ Verify Selected"** to see detailed comparison.

---

## Workflow Examples

### Workflow A: Small Batch (5-10 patients)

```
1. Load patients via Bulk Excel Import
2. Open TRF Manager
3. For each patient:
   a. Select patient
   b. Upload individual TRF file
   c. Click "Verify Selected"
   d. Apply any corrections
4. Close TRF Manager
5. Generate reports
```

### Workflow B: Large Batch (50+ patients) with Bulk PDF

```
1. Load patients via Bulk Excel Import
2. Open TRF Manager
3. Click "Upload Bulk TRF" → "Single Multi-Page PDF"
4. Select the bulk TRF PDF
5. Click "Verify All" (auto-matching)
6. Review matches in patient list (✅ = matched)
7. For unmatched patients:
   a. Select patient
   b. Upload TRF → Select Page from Bulk PDF
   c. Preview and assign correct page
8. Click "Verify All" again to validate all data
9. Close TRF Manager
10. Generate reports
```

### Workflow C: Mixed TRFs (Some individual, some bulk)

```
1. Load patients via Bulk Excel Import
2. Open TRF Manager
3. Upload Bulk PDF for most patients
4. Click "Verify All"
5. For patients with separate TRF files:
   a. Select patient
   b. Upload TRF → Upload Individual File
   c. Select the individual TRF file
6. Verify all patients
7. Generate reports
```

---

## Troubleshooting

### EasyOCR Issues

**Problem**: `ModuleNotFoundError: No module named 'easyocr'`
```bash
pip install easyocr>=1.7.0
```

**Problem**: First run is very slow
> This is normal - EasyOCR downloads models on first use (~100MB)

**Problem**: Out of memory errors
```python
# Use CPU mode (add to code or set environment)
import easyocr
reader = easyocr.Reader(['en'], gpu=False)
```

---

### Ollama Issues

**Problem**: "Ollama not running"
```bash
# Start Ollama server
ollama serve

# Or check if running
curl http://localhost:11434/api/tags
```

**Problem**: "Model not found"
```bash
# Download LLaVA model
ollama pull llava

# Verify it's downloaded
ollama list
```

**Problem**: Very slow responses
> LLaVA is compute-intensive. Ensure:
> - At least 8GB RAM available
> - Close other heavy applications
> - Consider using EasyOCR instead for speed

---

### Tesseract Issues

**Problem**: `TesseractNotFoundError`
```bash
# Linux
sudo apt install tesseract-ocr

# macOS
brew install tesseract

# Windows: Add to PATH or set manually
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

---

### PDF Issues

**Problem**: "Cannot extract text from PDF"
```bash
pip install pdfplumber>=0.10.0
```

**Problem**: Scanned PDF shows no text
> Use EasyOCR or Ollama LLaVA - they can read images from PDFs

---

### General Issues

**Problem**: Poor extraction accuracy
1. Try a different extraction method (Ollama LLaVA is most accurate)
2. Ensure TRF image quality is good
3. Check "Use AI Enhancement" option

**Problem**: Auto-matching fails
1. Ensure patient names in Excel match TRF names closely
2. Use consistent PIN formats
3. Manually assign pages for edge cases

**Problem**: Slow bulk processing
1. Use EasyOCR (faster than Ollama)
2. Process in smaller batches
3. Ensure sufficient system resources

---

## Quick Reference Card

| Task | Action |
|------|--------|
| Open TRF Manager | Bulk Editor → 📋 TRF Manager |
| Upload single TRF | Select patient → 📄 Upload TRF |
| Upload bulk PDF | 📑 Upload Bulk TRF → Single Multi-Page PDF |
| Auto-match all | ✓ Verify All |
| Verify one patient | Select → ✓ Verify Selected |
| Apply TRF value | Click [Apply] in comparison dialog |
| Apply all values | Click [✓ Apply All] in comparison dialog |
| Assign page manually | Select patient → 📄 Upload TRF → Select Page from Bulk PDF |
| Change extraction method | Settings → Extraction Method dropdown |
| Test Ollama | Settings → 🔗 Test Ollama Connection |

---

## Support

For issues or feature requests, contact the development team or raise an issue in the repository.

---

*Last updated: February 2026*
