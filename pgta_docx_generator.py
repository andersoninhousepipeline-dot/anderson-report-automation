"""
DOCX Report Generator for PGT-A Reports
Generates Word documents matching the PDF template
"""

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import sys
import io
from PIL import Image as PILImage


class PGTADocxGenerator:
    """Generates DOCX reports for PGT-A"""
    
    # Static content (same as PDF template)
    METHODOLOGY_TEXT = """Chromosomal aneuploidy analysis was performed using ChromInst® PGT-A from Yikon Genomics (Suzhou) Co., Ltd - China. The Yikon - ChromInst® PGT-A kit with the Genemind - SURFSeq 5000* High-throughput Sequencing Platform allows detection of aneuploidies in all 23 sets of Chromosomes. Probes are not covering the p arm of acrocentric chromosomes as they are rich in repeat regions and RNA markers and devoid of genes. Changes in this region will not be detected. However, these regions have less clinical significance due to the absence of genes. Chromosomal aneuploidy can be detected by copy number variations (CNVs), which represent a class of variation in which segments of the genome have been duplicated (gains) or deleted (losses). Large, genomic copy number imbalances can range from sub-chromosomal regions to entire chromosomes. Inherited and de-novo CNVs (up to 10 Mb) have been associated with many disease conditions. This assay was performed on DNA extracted from embryo biopsy samples."""
    
    MOSAICISM_TEXT = """Mosaicism arises in the embryo due to mitotic errors which lead to the production of karyotypically distinct cell lineages within a single embryo [1]. NGS has the sensitivity to detect mosaicism when 30% or the above cells are abnormal [2]. Mosaicism is reported in our laboratory as follows [3]."""
    
    MOSAICISM_BULLETS = [
        "Embryos with less than 30% mosaicism are considered as euploid.",
        "Embryos with 30% to 50% mosaicism will be reported as low level mosaic, 51% to 80% mosaicism will be reported as high level mosaic.",
        "When three chromosomes or more than three chromosomes showing mosaic change, it will be denoted as complex mosaic.",
        "If greater than 80% mosaicism detected in an embryo it will be considered aneuploid."
    ]
    
    MOSAICISM_CLINICAL = """Clinical significance of transferring mosaic embryos is still under evaluation. Based on Preimplantation Genetic Diagnosis International Society (PGDIS) Position Statement – 2019 transfer of these embryos should be considered only after appropriate counselling of the patient and alternatives have been discussed. Invasive prenatal testing with karyotyping in the amniotic fluid needs to be advised in such cases [4]. As shown in published literature evidence, such transfers can result in normal pregnancy or miscarriage or an offspring with chromosomal mosaicism [5,6,7]."""
    
    LIMITATIONS = [
        "This technique cannot detect point mutations, balanced translocations, inversions, triploidy, uniparental disomy and epigenetic modifications.",
        "Probes used do not cover the p arm of acrocentric chromosomes as they are rich in repeat regions and RNA markers and devoid of genes. Changes in this region will not be detected. However, these regions have less clinical significance due to the absence of genes.",
        "Deletions and duplications with the size of < 10 Mb cannot be detected.",
        "Risk of misinterpretation of the actual embryo karyotype due to the presence of chromosomal mosaicism, either at cleavage-stage or at blastocyst stage may exist.",
        "This technique cannot detect variants of polyploidy and haploidy",
        "NGS without genotyping cannot identify the nature (meiotic or mitotic) nor the parental origin of aneuploidies",
        "Due to the intrinsic nature of chromosomal mosaicism, the chromosomal make-up achieved from a biopsy only may represent a picture of a small part of the embryo and may not necessarily reflect the chromosomal content of the entire embryo. Also, the mosaicism level inferred from a multi-cell TE biopsy might not unequivocally represent the exact chromosomal mosaicism percentage of the TE cells or the inner cell mass constitution."
    ]
    
    REFERENCES = [
        'McCoy, Rajiv C. "Mosaicism in Preimplantation human embryos: when chromosomal abnormalities are the norm." Trends in genetics 33.7 (2017): 448-463.',
        'ESHRE PGT-SR/PGT-A Working Group, et al. "ESHRE PGT Consortium good practice recommendations for the detection of structural and numerical chromosomal aberrations." Human reproduction open 2020.3 (2020): hoaa017.',
        'ESHRE Working Group on Chromosomal Mosaicism, et al. "ESHRE survey results and good practice recommendations on managing chromosomal mosaicism." Hum Reprod Open. 2022 Nov 7;2022(4):hoac044.',
        'Cram, D. S., et al. "PGDIS position statement on the transfer of mosaic embryos 2019." Reproductive biomedicine online 39 (2019): e1-e4.',
        'Victor, Andrea R., et al. "One hundred mosaic embryos transferred prospectively in a single clinic: exploring when and why they result in healthy pregnancies." Fertility and sterility 111.2 (2019): 280-293.',
        'Lin, Pin-Yao, et al. "Clinical outcomes of single mosaic embryo transfer: high-level or low-level mosaic embryo, does it matter?" Journal of clinical medicine 9.6 (2020): 1695.',
        'Kahraman, Semra, et al. "The birth of a baby with mosaicism resulting from a known mosaic embryo transfer: a case report." Human Reproduction 35.3 (2020): 727-733.'
    ]
    
    SIGNATURES = [
        {"name": "Anand Babu. K, Ph.D", "title": "Molecular Biologist"},
        {"name": "Sachin D Honguntikar, Ph.D", "title": "Molecular Geneticist"},
        {"name": "Dr Suriyakumar G", "title": "Director"}
    ]
    
    @staticmethod
    def get_resource_path(relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        # Try PyInstaller path
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
            
        # Try relative to script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        path1 = os.path.join(script_dir, relative_path)
        if os.path.exists(path1):
            return path1
            
        # Try relative to CWD
        path2 = os.path.join(os.getcwd(), relative_path)
        if os.path.exists(path2):
            return path2
            
        return path1 # Fallback 

    def __init__(self, assets_dir="assets/pgta"):
        """Initialize DOCX generator"""
        # Resolve the assets directory relative to the script location
        self.assets_dir = self.get_resource_path(assets_dir)
        print(f"INFO DOCX: Assets Directory: {self.assets_dir}")
        
        self.header_logo = os.path.join(self.assets_dir, "image_page1_0.png")
        self.footer_banner = os.path.join(self.assets_dir, "image_page1_1.png")
        self.footer_logo = os.path.join(self.assets_dir, "image_page1_2.png")
        self.genqa_logo = os.path.join(self.assets_dir, "genqa_logo.png")
        self.signs_image = os.path.join(self.assets_dir, "signs.png")
        
        for label, path in [("DOCX_Logo", self.header_logo), ("DOCX_Signs", self.signs_image)]:
             if not os.path.exists(path):
                 print(f"WARNING DOCX: {label} not found at {path}")
    
    def _set_cell_background(self, cell, fill):
        """Set background shading for a table cell"""
        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:fill'), fill)
        cell._tc.get_or_add_tcPr().append(shading_elm)
    
    def generate_docx(self, output_path, patient_data, embryos_data):
        """Generate DOCX report"""
        doc = Document()
        
        # Set margins
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(0.9)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)
        
        # Page 1: Cover page
        self._add_cover_page(doc, patient_data, embryos_data)
        doc.add_page_break()
        
        # Setup headers and footers for all sections
        self._setup_page_header_footer(doc)
        
        # Page 2: Methodology
        self._add_methodology_page(doc)
        doc.add_page_break()
        
        # Pages 3+: Embryo results
        for idx, embryo in enumerate(embryos_data):
            self._add_embryo_page(doc, patient_data, embryo)
            # Add signature under EACH embryo detail section as requested
            self._add_signature_section(doc)
            if idx < len(embryos_data) - 1:
                doc.add_page_break()
        
        # Save document
        doc.save(output_path)
        return output_path
    
    def _setup_page_header_footer(self, doc):
        """Setup repeating headers and footers for all sections"""
        for section in doc.sections:
            # Header
            header = section.header
            header_para = header.paragraphs[0]
            header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if os.path.exists(self.header_logo):
                run = header_para.add_run()
                run.add_picture(self.header_logo, width=Inches(6.5))

            # Footer
            footer = section.footer
            footer_para = footer.paragraphs[0]
            footer_para.alignment = WD_ALIGN_PARAGRAPH.LEFT # Banner on left
            
            # Create a table for footer to position banner and GenQA logo
            footer_table = footer.add_table(rows=1, cols=2, width=Inches(6.5))
            # Set table width
            footer_table.autofit = False
            footer_table.columns[0].width = Inches(5.0)
            footer_table.columns[1].width = Inches(1.5)
            
            # Add Footer Banner
            if os.path.exists(self.footer_banner):
                para_banner = footer_table.rows[0].cells[0].paragraphs[0]
                run_banner = para_banner.add_run()
                run_banner.add_picture(self.footer_banner, width=Inches(5.0))
            
            # Add GenQA Logo
            if os.path.exists(self.genqa_logo):
                para_logo = footer_table.rows[0].cells[1].paragraphs[0]
                para_logo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                run_logo = para_logo.add_run()
                run_logo.add_picture(self.genqa_logo, width=Inches(0.8))

    def _add_cover_page(self, doc, patient_data, embryos_data):
        """Add cover page with patient info"""
        # Header is now handled by _setup_page_header_footer
        
        # Title
        title = doc.add_paragraph()
        title_run = title.add_run("Preimplantation Genetic Testing for Aneuploidies (PGT-A)")
        title_run.bold = True
        title_run.font.size = Pt(14)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph()  # Spacer
        
        # Patient info table
        table = doc.add_table(rows=13, cols=6)
        table.style = 'Table Grid'
        
        # Populate patient info
        self._populate_patient_table(table, patient_data)
        
        doc.add_paragraph()  # Spacer
        
        # PNDT Disclaimer
        disclaimer = doc.add_paragraph()
        disclaimer_run = disclaimer.add_run("This test does not reveal sex of the fetus & confers to PNDT act, 1994")
        disclaimer_run.italic = True
        disclaimer_run.font.size = Pt(9)
        disclaimer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph()  # Spacer
        
        # Indication
        if 'indication' in patient_data and patient_data['indication']:
            indication_header = doc.add_paragraph()
            indication_header.add_run("Indication").bold = True
            indication_text = doc.add_paragraph(patient_data['indication'])
            indication_text.style = 'Normal'
            doc.add_paragraph()  # Spacer
        
        # Results summary
        results_header = doc.add_paragraph()
        results_header.add_run("Results summary").bold = True
        results_header.style = 'Heading 2'
        
        # Results table
        results_table = doc.add_table(rows=len(embryos_data) + 1, cols=5)
        results_table.style = 'Table Grid'
        
        # Header row
        headers = ['S. No.', 'Sample', 'Result', 'MTcopy', 'Interpretation']
        for idx, header in enumerate(headers):
            cell = results_table.rows[0].cells[idx]
            cell.text = header
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.size = Pt(9)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._set_cell_background(cell, "FABF8F") # Exact peach from source
        
        # Data rows
        for idx, embryo in enumerate(embryos_data, 1):
            row = results_table.rows[idx]
            row.cells[0].text = str(idx)
            row.cells[1].text = embryo.get('embryo_id', '')
            row.cells[2].text = embryo.get('result_summary', '')
            
            # MTcopy: NA for non-euploid
            interp = embryo.get('interpretation', '')
            mtcopy = embryo.get('mtcopy', 'NA')
            if interp.upper() != "EUPLOID":
                mtcopy = "NA"
            row.cells[3].text = mtcopy
            row.cells[4].text = interp
            
            for idx_cell, cell in enumerate(row.cells):
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                res_sum = embryo.get('result_summary', '')
                res_interp = embryo.get('interpretation', '')
                text_color = self._get_result_color_hex(res_sum, res_interp)
                
                # ONLY first column gets peach
                if idx_cell == 0:
                    self._set_cell_background(cell, "FABF8F")  # Exact peach
                # All other data cells get light blue-grey
                else:
                    self._set_cell_background(cell, "F1F1F7")  # Exact from source
                
                # Apply text color
                if text_color and (idx_cell == 2 or idx_cell == 4): # Result and Interpretation
                    for p in cell.paragraphs:
                        for r in p.runs:
                            r.font.color.rgb = RGBColor.from_string(text_color[1:])
        
        doc.add_paragraph() # Spacer
    
    def _populate_patient_table(self, table, patient_data):
        """Populate patient information table"""
        # Row 0: Patient name and PIN
        table.rows[0].cells[0].text = "Patient name"
        table.rows[0].cells[0].paragraphs[0].runs[0].bold = True
        table.rows[0].cells[1].text = ":"
        table.rows[0].cells[2].text = patient_data.get('patient_name', '')
        table.rows[0].cells[2].paragraphs[0].runs[0].bold = True # Bold name
        table.rows[0].cells[3].text = "PIN"
        table.rows[0].cells[3].paragraphs[0].runs[0].bold = True
        table.rows[0].cells[4].text = ":"
        table.rows[0].cells[5].text = patient_data.get('pin', '')
        table.rows[0].cells[5].paragraphs[0].runs[0].bold = True # Bold PIN
        
        # Row 1: Spouse name
        table.rows[1].cells[2].text = patient_data.get('spouse_name', '')
        table.rows[1].cells[2].paragraphs[0].runs[0].bold = True
        
        # Row 4: Age and Sample Number
        table.rows[4].cells[0].text = "Date of Birth/ Age"
        table.rows[4].cells[0].paragraphs[0].runs[0].bold = True
        table.rows[4].cells[1].text = ":"
        table.rows[4].cells[2].text = patient_data.get('age', '')
        table.rows[4].cells[3].text = "Sample Number"
        table.rows[4].cells[3].paragraphs[0].runs[0].bold = True
        table.rows[4].cells[4].text = ":"
        table.rows[4].cells[5].text = patient_data.get('sample_number', '')
        
        # Row 6: Referring Clinician and Biopsy date
        table.rows[6].cells[0].text = "Referring Clinician"
        table.rows[6].cells[0].paragraphs[0].runs[0].bold = True
        table.rows[6].cells[1].text = ":"
        table.rows[6].cells[2].text = patient_data.get('referring_clinician', '')
        table.rows[6].cells[3].text = "Biopsy date"
        table.rows[6].cells[3].paragraphs[0].runs[0].bold = True
        table.rows[6].cells[4].text = ":"
        table.rows[6].cells[5].text = patient_data.get('biopsy_date', '')
        
        # Row 8: Hospital and Sample collection date
        table.rows[8].cells[0].text = "Hospital/Clinic"
        table.rows[8].cells[0].paragraphs[0].runs[0].bold = True
        table.rows[8].cells[1].text = ":"
        table.rows[8].cells[2].text = patient_data.get('hospital_clinic', '')
        table.rows[8].cells[3].text = "Sample collection date"
        table.rows[8].cells[3].paragraphs[0].runs[0].bold = True
        table.rows[8].cells[4].text = ":"
        table.rows[8].cells[5].text = patient_data.get('sample_collection_date', '')
        
        # Row 10: Specimen and Sample receipt date
        table.rows[10].cells[0].text = "Specimen"
        table.rows[10].cells[0].paragraphs[0].runs[0].bold = True
        table.rows[10].cells[1].text = ":"
        table.rows[10].cells[2].text = patient_data.get('specimen', '')
        table.rows[10].cells[3].text = "Sample receipt date"
        table.rows[10].cells[3].paragraphs[0].runs[0].bold = True
        table.rows[10].cells[4].text = ":"
        table.rows[10].cells[5].text = patient_data.get('sample_receipt_date', '')
        
        # Row 12: Report date and Biopsy performed by
        table.rows[12].cells[0].text = "Report date"
        table.rows[12].cells[0].paragraphs[0].runs[0].bold = True
        table.rows[12].cells[1].text = ":"
        table.rows[12].cells[2].text = patient_data.get('report_date', '')
        table.rows[12].cells[3].text = "Biopsy performed by"
        table.rows[12].cells[3].paragraphs[0].runs[0].bold = True
        table.rows[12].cells[4].text = ":"
        table.rows[12].cells[5].text = patient_data.get('biopsy_performed_by', '')
        
        # Set column widths precisely (Total = ~6.5 Inches)
        widths = [Inches(1.2), Inches(0.2), Inches(1.8), Inches(1.3), Inches(0.2), Inches(1.8)]
        for row in table.rows:
            for i, width in enumerate(widths):
                row.cells[i].width = width

        # Set font size for all cells
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    for run in paragraph.runs:
                        run.font.size = Pt(9)
                # All patient info cells get light blue-grey background
                self._set_cell_background(cell, "F1F1F7")  # Exact from source
    
    def _add_methodology_page(self, doc):
        """Add methodology and static content page"""
        # Methodology
        method_header = doc.add_paragraph()
        method_header.add_run("Methodology").bold = True
        method_header.style = 'Heading 2'
        
        method_text = doc.add_paragraph(self.METHODOLOGY_TEXT)
        method_text.style = 'Normal'
        for run in method_text.runs:
            run.font.size = Pt(9)
        
        doc.add_paragraph()  # Spacer
        
        # Mosaicism
        mosaic_header = doc.add_paragraph()
        mosaic_header.add_run("Conditions for reporting mosaicism").bold = True
        mosaic_header.style = 'Heading 2'
        
        mosaic_text = doc.add_paragraph(self.MOSAICISM_TEXT)
        for run in mosaic_text.runs:
            run.font.size = Pt(9)
        
        # Bullets
        for bullet in self.MOSAICISM_BULLETS:
            p = doc.add_paragraph(bullet, style='List Bullet')
            for run in p.runs:
                run.font.size = Pt(9)
        
        clinical_text = doc.add_paragraph(self.MOSAICISM_CLINICAL)
        for run in clinical_text.runs:
            run.font.size = Pt(9)
        
        doc.add_paragraph()  # Spacer
        
        # Limitations
        limit_header = doc.add_paragraph()
        limit_header.add_run("Limitations").bold = True
        limit_header.style = 'Heading 2'
        
        for limitation in self.LIMITATIONS:
            p = doc.add_paragraph(limitation, style='List Bullet')
            for run in p.runs:
                run.font.size = Pt(9)
        
        doc.add_paragraph()  # Spacer
        
        # References
        ref_header = doc.add_paragraph()
        ref_header.add_run("References").bold = True
        ref_header.style = 'Heading 2'
        
        for idx, ref in enumerate(self.REFERENCES, 1):
            ref_text = doc.add_paragraph(f"{idx}. {ref}")
            for run in ref_text.runs:
                run.font.size = Pt(8)
    
    def _add_embryo_page(self, doc, patient_data, embryo_data):
        """Add individual embryo results page"""
        # Patient info
        patient_p = doc.add_paragraph()
        run1 = patient_p.add_run(f"Patient name : ")
        run1.bold = True
        patient_p.add_run(f"{patient_data.get('patient_name', '')} {patient_data.get('spouse_name', '')}").bold = True
        
        pin_p = doc.add_paragraph()
        run2 = pin_p.add_run(f"PIN : ")
        run2.bold = True
        pin_p.add_run(f"{patient_data.get('pin', '')}").bold = True
        
        doc.add_paragraph()  # Spacer
        
        # PNDT Disclaimer in a grey box
        disclaimer_table = doc.add_table(rows=1, cols=1)
        disclaimer_table.width = Inches(6.0)
        cell = disclaimer_table.rows[0].cells[0]
        self._set_cell_background(cell, "F2F2F2") # Grey bg
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run("This test does not reveal sex of the fetus & confers to PNDT act, 1994")
        run.bold = True
        run.font.size = Pt(9)
        
        doc.add_paragraph()  # Spacer
        
        doc.add_paragraph(f"EMBRYO: {embryo_data.get('embryo_id', '')}").style = 'Heading 2'
        
        # Details in a colored table for background support
        res_text = embryo_data.get('result_description', '')
        interp_text = embryo_data.get('interpretation', '')
        autosomes_text = embryo_data.get('autosomes', '')
        
        # Color logic
        interp_color = self._get_result_color_hex('', interp_text)
        auto_color = self._get_result_color_hex(autosomes_text, '')
        
        # MTcopy logic
        mtcopy = embryo_data.get('mtcopy', 'NA')
        if interp_text.upper() != "EUPLOID":
            mtcopy = "NA"
            
        summary_table = doc.add_table(rows=5, cols=2)
        summary_table.style = 'Table Grid'
        
        rows = [
            ("Result:", res_text, None), # Result remains black
            ("Autosomes:", autosomes_text, auto_color),
            ("Sex chromosomes:", embryo_data.get('sex_chromosomes', ''), None),
            ("Interpretation:", interp_text, interp_color),
            ("MTcopy:", mtcopy, None)
        ]
        
        for i, (label, val, color) in enumerate(rows):
            row = summary_table.rows[i]
            # Bold the label
            p0 = row.cells[0].paragraphs[0]
            r0 = p0.add_run(label)
            r0.bold = True
            
            # Bold the value and color if needed
            p1 = row.cells[1].paragraphs[0]
            r1 = p1.add_run(val)
            r1.bold = True
            
            for cell in row.cells:
                self._set_cell_background(cell, "F1F1F7")
                for p in cell.paragraphs:
                    p.runs[0].font.size = Pt(9)
                    if color: 
                        p.runs[0].font.color.rgb = RGBColor.from_string(color[1:])

        doc.add_paragraph()  # Spacer
        
        # CNV Chart
        cnv_header = doc.add_paragraph()
        cnv_header.add_run("COPY NUMBER CHART").bold = True
        cnv_header.style = 'Heading 2'
        
        # CNV Chart Image
        if 'cnv_image_path' in embryo_data and embryo_data['cnv_image_path'] and os.path.exists(embryo_data['cnv_image_path']):
            try:
                doc.add_picture(embryo_data['cnv_image_path'], width=Inches(6.0))
                doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph()  # Spacer
            except Exception as e:
                print(f"Error loading image into DOCX: {e}")
        
        # CNV table
        chr_statuses = embryo_data.get('chromosome_statuses', {})
        mosaic_percentages = embryo_data.get('mosaic_percentages', {})
        has_mosaic = any(mosaic_percentages.values())
        
        if has_mosaic:
            cnv_table = doc.add_table(rows=3, cols=23)
        else:
            cnv_table = doc.add_table(rows=2, cols=23)
        
        cnv_table.style = 'Table Grid'
        
        # Header row
        cnv_table.rows[0].cells[0].text = "Chromosome"
        for i in range(1, 23):
            cnv_table.rows[0].cells[i].text = str(i)
        
        # CNV status row
        cnv_table.rows[1].cells[0].text = "CNV status"
        for i in range(1, 23):
            cnv_table.rows[1].cells[i].text = chr_statuses.get(str(i), 'N')
        
        # Mosaic row if applicable
        if has_mosaic:
            cnv_table.rows[2].cells[0].text = "Mosaic (%)"
            for i in range(1, 23):
                cnv_table.rows[2].cells[i].text = str(mosaic_percentages.get(str(i), '-'))
        
        # Color headers
        for i in range(23):
             header_cell = cnv_table.rows[0].cells[i]
             self._set_cell_background(header_cell, "FABF8F") # Exact peach from source
             # Bold header text
             if header_cell.paragraphs[0].runs:
                  header_cell.paragraphs[0].runs[0].bold = True
             else:
                  header_cell.paragraphs[0].add_run().bold = True

        # Format table
        # Total width = 6.0 inches. Column 0 = 1.0 inch, others = (5.0 / 22) = ~0.227 inches
        first_col_width = Inches(1.0)
        other_col_width = Inches(5.0 / 22)
        
        for row in cnv_table.rows:
            for idx_cell, cell in enumerate(row.cells):
                if idx_cell == 0:
                    cell.width = first_col_width
                else:
                    cell.width = other_col_width
                    
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        run.font.size = Pt(8)
                
                if idx_cell == 0: # ONLY first column gets peach
                    self._set_cell_background(cell, "FABF8F")  # Exact peach
                else: # All other data cells get light blue-grey
                    self._set_cell_background(cell, "F1F1F7")  # Exact from source
                
                # Apply status color (using current loop index k for statuses)
                # Wait, the previous code had a bug here using 'i' from outer scope
                if idx_cell > 0:
                    status = chr_statuses.get(str(idx_cell), 'N')
                    s_color = self._get_status_color_hex(status)
                    if s_color:
                        for p in cell.paragraphs:
                            for r in p.runs:
                                 r.font.color.rgb = RGBColor.from_string(s_color[1:])
        
        doc.add_paragraph()  # Spacer
        
        # Legend
        legend = doc.add_paragraph()
        legend_run = legend.add_run(
            "N – Normal, G-Gain, L-Loss, SG-Segmental Gain, SL-Segmental Loss, "
            "M-Mosaic, MG- Mosaic Gain, ML-Mosaic Loss, SMG-Segmental Mosaic Gain, "
            "SML-Segmental Mosaic Loss"
        )
        legend_run.italic = True
        legend_run.font.size = Pt(8)
        
        doc.add_paragraph()  # Spacer
        # Signature section removed from intermediate pages

    def _add_signature_section(self, doc):
        """Add signature section (image preferred, fallback to table)"""
        sig_image_path = self.signs_image
        if os.path.exists(sig_image_path):
            try:
                sig_p = doc.add_paragraph()
                sig_p.add_run("This report has been reviewed and approved by:").bold = True
                
                try:
                    # Optimized width to match PDF (approx 400pt ~ 5.5 inches)
                    doc.add_picture(sig_image_path, width=Inches(5.5))
                except Exception:
                    # Conversion fallback for UnrecognizedImageError
                    with PILImage.open(sig_image_path) as img:
                        img = img.convert("RGB")
                        img_byte_arr = io.BytesIO()
                        img.save(img_byte_arr, format='PNG')
                        img_byte_arr.seek(0)
                        doc.add_picture(img_byte_arr, width=Inches(5.5))
                
                doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                return
            except Exception as e:
                print(f"Error adding signature image to DOCX: {repr(e)}")

        # Fallback to table
        sig_header = doc.add_paragraph("This report has been reviewed and approved by:")
        sig_header.runs[0].bold = True
        
        sig_table = doc.add_table(rows=2, cols=3)
        
        for idx, sig in enumerate(self.SIGNATURES):
            sig_table.rows[0].cells[idx].text = sig['name']
            sig_table.rows[0].cells[idx].paragraphs[0].runs[0].bold = True
            sig_table.rows[0].cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            sig_table.rows[1].cells[idx].text = sig['title']
            sig_table.rows[1].cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    def _get_result_color_hex(self, result_text, interpretation_text):
        """Determine if text should be Red (#FF0000), Blue (#0000FF) or Black"""
        res_up = result_text.upper() if result_text else ""
        int_up = interpretation_text.upper() if interpretation_text else ""
        
        red_keywords = ["MONOSOMY", "TRISOMY", "SEGMENTAL GAIN", "SEGMENTAL LOSS", 
                        "MULTIPLE CHROMOSOMAL ABNORMALITIES", "ANEUPLOID", "CHAOTIC EMBRYO"]
        if any(kw in res_up for kw in red_keywords) or any(kw in int_up for kw in red_keywords):
            return "#FF0000"
            
        blue_keywords = ["MOSAIC CHROMOSOME COMPLEMENT", "LOW LEVEL MOSAIC", 
                         "HIGH LEVEL MOSAIC", "COMPLEX MOSAIC", "MULTIPLE MOSAIC"]
        if any(kw in res_up for kw in blue_keywords) or any(kw in int_up for kw in blue_keywords):
            return "#0000FF"
            
        return None

    def _get_autosome_color_hex(self, autosome_text):
        """Special color logic for autosomes field"""
        if not autosome_text: return None
        txt = autosome_text.upper()
        if "MULTIPLE MOSAIC CHROMOSOME COMPLEMENT" in txt:
            return "#0000FF"
        return None

    def _get_status_color_hex(self, status):
        """Color logic for CNV status codes"""
        if not status: return None
        s = status.upper().strip()
        red_codes = ["G", "L", "SG", "SL"]
        blue_codes = ["M", "MG", "ML", "SMG", "SML"]
        
        if s in red_codes: return "#FF0000"
        if s in blue_codes: return "#0000FF"
        
        try:
            val = float(s.replace('%', ''))
            if val > 0: return "#0000FF"
        except:
            pass
            
        return None


if __name__ == "__main__":
    # Test the DOCX generator
    generator = PGTADocxGenerator()
    
    # Sample data
    patient_data = {
        'patient_name': 'Mrs. Priya (PNM00791)',
        'patient_spouse': 'Mrs. Priya (PNM00791)',
        'spouse_name': 'Mr. Saranraj',
        'pin': 'AND25630004206',
        'age': '34 Years',
        'sample_number': '632504349',
        'referring_clinician': 'Dr. Ajantha. B',
        'biopsy_date': '03-01-2026',
        'hospital_clinic': 'Rhea Healthcare Private Limited Annanagar (NOVA IVF)',
        'sample_collection_date': '03-01-2026',
        'specimen': 'Day 6 Trophectoderm Biopsy',
        'sample_receipt_date': '03-01-2026',
        'biopsy_performed_by': 'Raj Priya Pandian',
        'report_date': '14-01-2026',
        'indication': 'History of implantation failure.'
    }
    
    embryos_data = [
        {
            'embryo_id': 'PS4',
            'result_summary': 'Trisomy of chromosome 16',
            'mtcopy': 'NA',
            'interpretation': 'Aneuploid',
            'result_description': 'The embryo contains abnormal chromosome complement',
            'autosomes': 'Trisomy of chromosome 16',
            'sex_chromosomes': 'Normal',
            'chromosome_statuses': {str(i): 'N' for i in range(1, 23)},
            'mosaic_percentages': {},
            'cnv_image_path': 'assets/pgta/page4_image_3.jpg'
        }
    ]
    
    embryos_data[0]['chromosome_statuses']['16'] = 'G'
    
    # Generate DOCX
    output_path = "test_report.docx"
    generator.generate_docx(output_path, patient_data, embryos_data)
    print(f"Test DOCX report generated: {output_path}")
