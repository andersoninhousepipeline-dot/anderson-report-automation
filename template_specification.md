# PGT-A Report Template Specification

## Document Properties
- **Page Size**: 612 x 792 points (US Letter / A4)
- **Page Size (mm)**: 215.9 x 279.4 mm
- **Margins**: 
  - Left/Right: 72 points (1 inch / 25.4mm)
  - Top: ~72 points
  - Bottom: ~66 points

## Extracted Assets

### Header Images
1. **image_page1_0.png** (1280x193) - Main header logo/banner
   - Position: Top of page (y: 720-790)
   - Width: 468 points (full content width)
   
2. **image_page1_1.png** (1299x182) - Footer banner
   - Position: Bottom of page (y: 0-66)
   - Width: 468 points

3. **image_page1_2.png** (186x99) - Small logo (bottom right)
   - Position: Bottom right (x: 454-521, y: 35-71)
   - Width: 67 points

### Content Images
4. **page4_image_3.jpg** - CNV Chart/Graph (signature section)
5. **page4_image_4.jpg** - Copy Number Chart table

## Color Scheme
- **Primary Text**: Black (#000000)
- **Headers**: Dark Blue/Black
- **Table Borders**: Light gray
- **Background**: White

## Typography
- **Main Font**: Helvetica / Arial
- **Title**: ~16pt, Bold
- **Section Headers**: ~12pt, Bold
- **Body Text**: ~10pt, Regular
- **Table Text**: ~9pt, Regular
- **Small Text**: ~8pt (disclaimers)

## Page Structure

### Page 1: Cover Page with Patient Info & Results Summary

#### Header Section (Top)
- Logo/Banner image (full width)
- Title: "Preimplantation Genetic Testing for Aneuploidies (PGT-A)"
- Centered, bold, ~14-16pt

#### Patient Information Table
Two-column layout:

**Left Column:**
- Patient name
- Date of Birth/Age
- Referring Clinician
- Hospital/Clinic
- Specimen
- Report date

**Right Column:**
- PIN
- Sample Number
- Biopsy date
- Sample collection date
- Sample receipt date
- Biopsy performed by

#### PNDT Disclaimer
- Italic text: "This test does not reveal sex of the fetus & confers to PNDT act, 1994"
- Centered, below patient info

#### Indication Section
- **Label**: "Indication"
- **Content**: Free text (e.g., "History of implantation failure")

#### Results Summary Table
Columns:
1. S. No.
2. Sample (Embryo ID)
3. Result
4. MTcopy
5. Interpretation

### Page 2: Methodology & Mosaicism Conditions (STATIC)

#### Methodology Section
- **Title**: "Methodology"
- **Content**: Static text about ChromInst® PGT-A kit
- Bullet points about probe coverage and CNV detection

#### Conditions for Reporting Mosaicism
- **Title**: "Conditions for reporting mosaicism"
- **Content**: Static text with bullet points
  - <30%: Euploid
  - 30-50%: Low level mosaic
  - 51-80%: High level mosaic
  - >80%: Aneuploid
  - 3+ chromosomes: Complex mosaic

#### Limitations Section
- **Title**: "Limitations"
- **Content**: Static bullet points about technique limitations

#### References Section
- **Title**: "References"
- **Content**: Static numbered list of 7 references

### Page 3+: Individual Embryo Results (REPEATABLE)

**For each embryo, create a page with:**

#### Header
- Same logo/banner as page 1
- Patient name and PIN
- PNDT disclaimer

#### Embryo Details Box
- **EMBRYO**: [ID] (e.g., PS4)
- **Result**: [Description]
- **Autosomes**: [Findings]
- **Sex chromosomes**: [Findings]
- **Interpretation**: [Aneuploid/Euploid/Mosaic]
- **MTcopy**: [Value or NA]

#### Copy Number Chart Table
Table with columns:
- Chromosome (1-22)
- CNV status (N/G/L/SG/SL/M/MG/ML/SMG/SML)
- Mosaic (%) - if applicable

**Legend below table:**
N – Normal, G-Gain, L-Loss, SG-Segmental Gain, SL-Segmental Loss, M-Mosaic, MG- Mosaic Gain, ML-Mosaic Loss, SMG-Segmental Mosaic Gain, SML-Segmental Mosaic Loss

#### Signature Section
Three columns:
1. **Anand Babu. K, Ph.D** - Molecular Biologist
2. **Sachin D Honguntikar, Ph.D** - Molecular Geneticist
3. **Dr Suriyakumar G** - Director

#### Footer
- Same footer banner as page 1
- Small logo (bottom right)

## Dynamic Fields (from requirements.txt)

### Patient Information
- `Patient_Name`
- `PIN`
- `Date_of_Birth_Age`
- `Sample_Number`
- `Referring_Clinician`
- `Biopsy_Date`
- `Hospital_Clinic`
- `Sample_Collection_Date`
- `Specimen`
- `Sample_Receipt_Date`
- `Biopsy_Performed_By`
- `Report_Date`
- `Indication` (new field)

### Results Summary (Per Embryo)
- `S_No`
- `Embryo_ID` (Sample)
- `Result`
- `MTcopy`
- `Interpretation`

### Embryo Details (Per Embryo)
- `Embryo_ID`
- `Result_Description`
- `Autosomes`
- `Sex_Chromosomes`
- `Interpretation`
- `MTcopy`
- `CNV_Chart_Image` (user uploaded PNG)
- `Chromosome_1_Status` through `Chromosome_22_Status`
- `Mosaic_Percentages` (if applicable)

## Static Content

### Methodology Text
```
Chromosomal aneuploidy analysis was performed using ChromInst® PGT-A from Yikon Genomics (Suzhou)
Co., Ltd - China. The Yikon - ChromInst® PGT-A kit with the Genemind - SURFSeq 5000* High-throughput
Sequencing Platform allows detection of aneuploidies in all 23 sets of Chromosomes. Probes are not
covering the p arm of acrocentric chromosomes as they are rich in repeat regions and RNA markers and
devoid of genes. Changes in this region will not be detected. However, these regions have less clinical
significance due to the absence of genes. Chromosomal aneuploidy can be detected by copy number
variations (CNVs), which represent a class of variation in which segments of the genome have been
duplicated (gains) or deleted (losses). Large, genomic copy number imbalances can range from sub-
chromosomal regions to entire chromosomes. Inherited and de-novo CNVs (up to 10 Mb) have been
associated with many disease conditions. This assay was performed on DNA extracted from embryo
biopsy samples.
```

### Mosaicism Conditions Text
```
Mosaicism arises in the embryo due to mitotic errors which lead to the production of karyotypically
distinct cell lineages within a single embryo [1]. NGS has the sensitivity to detect mosaicism when 30% or
the above cells are abnormal [2]. Mosaicism is reported in our laboratory as follows [3].

• Embryos with less than 30% mosaicism are considered as euploid.
• Embryos with 30% to 50% mosaicism will be reported as low level mosaic, 51% to 80% mosaicism
  will be reported as high level mosaic.
• When three chromosomes or more than three chromosomes showing mosaic change, it will be
  denoted as complex mosaic.
• If greater than 80% mosaicism detected in an embryo it will be considered aneuploid.

Clinical significance of transferring mosaic embryos is still under evaluation. Based on Preimplantation
Genetic Diagnosis International Society (PGDIS) Position Statement – 2019 transfer of these embryos
should be considered only after appropriate counselling of the patient and alternatives have been
discussed. Invasive prenatal testing with karyotyping in the amniotic fluid needs to be advised in such
cases [4]. As shown in published literature evidence, such transfers can result in normal pregnancy or
miscarriage or an offspring with chromosomal mosaicism [5,6,7].
```

### Limitations Text
```
• This technique cannot detect point mutations, balanced translocations, inversions, triploidy,
  uniparental disomy and epigenetic modifications.
• Probes used do not cover the p arm of acrocentric chromosomes as they are rich in repeat regions
  and RNA markers and devoid of genes. Changes in this region will not be detected. However, these
  regions have less clinical significance due to the absence of genes.
• Deletions and duplications with the size of < 10 Mb cannot be detected.
• Risk of misinterpretation of the actual embryo karyotype due to the presence of chromosomal
  mosaicism, either at cleavage-stage or at blastocyst stage may exist.
• This technique cannot detect variants of polyploidy and haploidy
• NGS without genotyping cannot identify the nature (meiotic or mitotic) nor the parental origin of
  aneuploidies
• Due to the intrinsic nature of chromosomal mosaicism, the chromosomal make-up achieved from a
  biopsy only may represent a picture of a small part of the embryo and may not necessarily reflect
  the chromosomal content of the entire embryo. Also, the mosaicism level inferred from a multi-cell
  TE biopsy might not unequivocally represent the exact chromosomal mosaicism percentage of the
  TE cells or the inner cell mass constitution.
```

### References
```
1. McCoy, Rajiv C. "Mosaicism in Preimplantation human embryos: when chromosomal abnormalities
   are the norm." Trends in genetics 33.7 (2017): 448-463.
2. ESHRE PGT-SR/PGT-A Working Group, et al. "ESHRE PGT Consortium good practice recommendations
   for the detection of structural and numerical chromosomal aberrations." Human reproduction open
   2020.3 (2020): hoaa017.
3. ESHRE Working Group on Chromosomal Mosaicism, et al. "ESHRE survey results and good practice
   recommendations on managing chromosomal mosaicism." Hum Reprod Open. 2022 Nov
   7;2022(4):hoac044.
4. Cram, D. S., et al. "PGDIS position statement on the transfer of mosaic embryos 2019." Reproductive
   biomedicine online 39 (2019): e1-e4.
5. Victor, Andrea R., et al. "One hundred mosaic embryos transferred prospectively in a single clinic:
   exploring when and why they result in healthy pregnancies." Fertility and sterility 111.2 (2019): 280-
   293.
6. Lin, Pin-Yao, et al. "Clinical outcomes of single mosaic embryo transfer: high-level or low-level mosaic
   embryo, does it matter?" Journal of clinical medicine 9.6 (2020): 1695.
7. Kahraman, Semra, et al. "The birth of a baby with mosaicism resulting from a known mosaic embryo
   transfer: a case report." Human Reproduction 35.3 (2020): 727-733.
```

## Notes
- All static content (methodology, mosaicism, limitations, references) remains identical across all reports
- Signature section remains identical (same three names and titles)
- Header and footer images remain identical
- Only patient info, embryo data, and CNV charts are dynamic
