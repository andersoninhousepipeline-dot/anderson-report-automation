"""
PGT-A Report Template Engine
Generates PDF and DOCX reports from patient and embryo data
"""

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch, mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, 
    Spacer, PageBreak, Image, KeepTogether, CondPageBreak
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.pdfgen import canvas
from PIL import Image as PILImage
import os
import sys
from datetime import datetime


class PGTAReportTemplate:
    """Template engine for PGT-A reports"""
    
    COLORS = {
        'patient_info_bg': '#F1F1F7',
        'results_header_bg': '#F9BE8F',
        'grey_bg': '#F2F2F2',
        'blue_title': '#1F497D',
        'approval_blue': '#4F81BD'
    }
    
    # Page dimensions (US Letter)
    PAGE_WIDTH = 612  # points
    PAGE_HEIGHT = 792  # points
    MARGIN_LEFT = 58  # Reduced from 72 to allow title in single line
    MARGIN_RIGHT = 58
    MARGIN_TOP = 100
    MARGIN_BOTTOM = 60
    CONTENT_WIDTH = PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT  # 496 points
    
    # Asset paths
    ASSETS_DIR = "extracted_assets"
    HEADER_LOGO = os.path.join(ASSETS_DIR, "image_page1_0.png")
    FOOTER_BANNER = os.path.join(ASSETS_DIR, "image_page1_1.png")
    FOOTER_LOGO = os.path.join(ASSETS_DIR, "image_page1_2.png")
    
    # Static content
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
            
        return path1 # Fallback to path1 even if not exists for error reporting

    def __init__(self, assets_dir="assets/pgta"):
        """Initialize template with asset directory"""
        # Resolve the assets directory relative to the script location
        self.ASSETS_DIR = self.get_resource_path(assets_dir)
        print(f"INFO: Assets Directory resolved to: {self.ASSETS_DIR}")
        
        # Hardcode specific asset filenames to avoid manual adaptation
        self.HEADER_LOGO = os.path.join(self.ASSETS_DIR, "image_page1_0.png")
        self.FOOTER_BANNER = os.path.join(self.ASSETS_DIR, "image_page1_1.png")
        self.FOOTER_LOGO = os.path.join(self.ASSETS_DIR, "image_page1_2.png")
        self.GENQA_LOGO = os.path.join(self.ASSETS_DIR, "genqa_logo.png")
        
        # Confirm critical assets exist
        for label, path in [("HeaderLogo", self.HEADER_LOGO), ("FooterBanner", self.FOOTER_BANNER)]:
            if not os.path.exists(path):
                print(f"CRITICAL WARNING: Asset not found: {label} at {path}")
            else:
                print(f"SUCCESS: Found asset: {label}")
        
        # Create custom styles
        self.styles = getSampleStyleSheet()
        self._register_fonts()
        self._create_custom_styles()
    
    def _register_fonts(self):
        """Register custom fonts if they exist in assets/fonts"""
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.pdfbase.pdfmetrics import registerFontFamily
        
        fonts_dir = os.path.join(self.ASSETS_DIR, "fonts")
        if not os.path.exists(fonts_dir):
            return
            
        # Actual filenames from user upload
        font_configs = [
            # Segoe UI
            {'name': 'SegoeUI', 'file': 'SEGOEUI.TTF'},
            {'name': 'SegoeUI-Bold', 'file': 'SEGOEUIB.TTF'},
            {'name': 'SegoeUI-Italic', 'file': 'SEGOEUII.TTF'},
            {'name': 'SegoeUI-BoldItalic', 'file': 'SEGOEUIZ.TTF'},
            # Gill Sans MT
            {'name': 'GillSansMT', 'file': 'GIL_____.TTF'},
            {'name': 'GillSansMT-Bold', 'file': 'GILB____.TTF'},
            {'name': 'GillSansMT-Italic', 'file': 'GILI____.TTF'},
            {'name': 'GillSansMT-BoldItalic', 'file': 'GILBI___.TTF'},
            # Calibri
            {'name': 'Calibri', 'file': 'CALIBRI.TTF'},
            {'name': 'Calibri-Bold', 'file': 'CALIBRIB.TTF'},
            {'name': 'Calibri-Italic', 'file': 'CALIBRII.TTF'},
            {'name': 'Calibri-BoldItalic', 'file': 'CALIBRIZ.TTF'},
        ]
        
        registered = []
        for cfg in font_configs:
            font_path = os.path.join(fonts_dir, cfg['file'])
            if os.path.exists(font_path):
                try:
                    pdfmetrics.registerFont(TTFont(cfg['name'], font_path))
                    registered.append(cfg['name'])
                    print(f"Registered font: {cfg['name']}")
                except Exception as e:
                    print(f"Error registering font {cfg['name']}: {e}")
        
        # Register font families if components exist
        if 'SegoeUI' in registered and 'SegoeUI-Bold' in registered:
            registerFontFamily('SegoeUI', normal='SegoeUI', bold='SegoeUI-Bold', italic='SegoeUI-Italic', boldItalic='SegoeUI-BoldItalic')
        if 'GillSansMT' in registered and 'GillSansMT-Bold' in registered:
            registerFontFamily('GillSansMT', normal='GillSansMT', bold='GillSansMT-Bold', italic='GillSansMT-Italic', boldItalic='GillSansMT-BoldItalic')
        if 'Calibri' in registered and 'Calibri-Bold' in registered:
            registerFontFamily('Calibri', normal='Calibri', bold='Calibri-Bold', italic='Calibri-Italic', boldItalic='Calibri-BoldItalic')
    
    def _get_font(self, name, fallback):
        """Helper to get best available font"""
        try:
            from reportlab.pdfbase import pdfmetrics
            pdfmetrics.getFont(name)
            return name
        except:
            return fallback

    def _create_custom_styles(self):
        """Create custom paragraph styles"""

        # Title style
        self.styles.add(ParagraphStyle(
            name='PGTAReportTitle',
            parent=self.styles['Heading1'],
            fontSize=16, # Adjusted to match source and ensure single line
            leading=18,
            textColor=colors.HexColor(self.COLORS['blue_title']),
            alignment=TA_CENTER,
            spaceAfter=12,
            fontName=self._get_font('GillSansMT-Bold', 'Helvetica-Bold')
        ))
        
        # Section header
        self.styles.add(ParagraphStyle(
            name='PGTASectionHeader',
            parent=self.styles['Heading2'],
            fontSize=11,
            leading=13,
            textColor=colors.HexColor(self.COLORS['blue_title']), # Color as in source
            spaceBefore=12,
            spaceAfter=3, # Reduced to accommodate line
            fontName=self._get_font('SegoeUI-Bold', 'Helvetica-Bold')
        ))
        
        # Body text
        self.styles.add(ParagraphStyle(
            name='PGTABodyText',
            parent=self.styles['Normal'],
            fontSize=11,  # Matches source 11.04pt
            leading=13,
            alignment=TA_LEFT,
            fontName=self._get_font('Calibri', 'Helvetica')
        ))
        
        # Small text
        self.styles.add(ParagraphStyle(
            name='PGTASmallText',
            parent=self.styles['Normal'],
            fontSize=8,
            leading=10,
            fontName=self._get_font('SegoeUI', 'Helvetica')
        ))
        
        # Bold disclaimer (PNDT)
        self.styles.add(ParagraphStyle(
            name='PGTADisclaimer',
            parent=self.styles['Normal'],
            fontSize=11, # Matched to Indication (approx 11pt)
            leading=13,
            alignment=TA_CENTER,
            fontName=self._get_font('SegoeUI-Bold', 'Helvetica-Bold'),
            textColor=colors.black
        ))
        
        # Bullet style
        self.styles.add(ParagraphStyle(
            name='PGTABulletText',
            parent=self.styles['Normal'],
            fontSize=11, # Increased to match Methodology (11pt)
            leading=13, # Increased leading
            leftIndent=20,
            bulletIndent=10,
            fontName=self._get_font('Calibri', 'Helvetica')
        ))
        
        # Signature Approval line style
        self.styles.add(ParagraphStyle(
            name='PGTASigApproval',
            parent=self.styles['Normal'],
            fontSize=12.48, # Exact source size
            leading=14.5,
            textColor=colors.HexColor(self.COLORS['approval_blue']),
            fontName=self._get_font('SegoeUI-Bold', 'Helvetica-Bold')
        ))
        
        # Centered Body style for reliable table alignment
        self.styles.add(ParagraphStyle(
            name='PGTACenteredBodyText',
            parent=self.styles['PGTABodyText'],
            alignment=TA_CENTER
        ))
    
    def generate_pdf(self, output_path, patient_data, embryos_data):
        """
        Generate PDF report
        
        Args:
            output_path: Path to save PDF
            patient_data: Dictionary with patient information
            embryos_data: List of dictionaries with embryo data
        """
        # Create PDF document
        doc = SimpleDocTemplate(
            output_path,
            pagesize=letter,
            leftMargin=self.MARGIN_LEFT,
            rightMargin=self.MARGIN_RIGHT,
            topMargin=self.MARGIN_TOP,
            bottomMargin=self.MARGIN_BOTTOM
        )
        
        # Build story (content)
        story = []
        
        # Page 1: Cover with patient info and results summary
        story.extend(self._build_cover_page(patient_data, embryos_data))
        # No PageBreak here to allow Methodology to flow immediately after Results
        
        # Methodolgy content (Flows after cover, includes CondPageBreak for References)
        story.extend(self._build_methodology_page())
        
        # Strict page break before embryo results as requested
        story.append(PageBreak())
        
        # Pages 3+: Individual embryo results
        for idx, embryo in enumerate(embryos_data):
            story.extend(self._build_embryo_page(patient_data, embryo))
            # Add signature under EACH embryo detail section as requested
            story.append(Spacer(1, 12))
            story.append(self._create_signature_table())
            story.append(PageBreak())
        
        # Build PDF
        doc.build(story, onFirstPage=self._add_header_footer, 
                  onLaterPages=self._add_header_footer)
        
        return output_path
    
    def _add_header_footer(self, canvas, doc):
        """Add header and footer to each page"""
        canvas.saveState()
        
        # Helper to draw using PIL for better compatibility
        def draw_img_safe(path, x, y, w, h):
            if os.path.exists(path):
                try:
                    # Drawing PIL images in reportlab is more robust
                    with PILImage.open(path) as img:
                        canvas.drawInlineImage(img, x, y, width=w, height=h, preserveAspectRatio=True)
                        return True
                except Exception as e:
                    print(f"Error drawing image {path}: {e}")
            return False

        # Add header logo
        draw_img_safe(self.HEADER_LOGO, 72, 720, 468, 72)
        
        # Add footer banner
        draw_img_safe(self.FOOTER_BANNER, 72, 0.4, 468, 66)
        
        # Add small GenQA logo
        draw_img_safe(self.GENQA_LOGO, 454, 35, 67, 36)
        
        canvas.restoreState()
    
    def _build_cover_page(self, patient_data, embryos_data):
        """Build cover page with patient info and results summary"""
        elements = []
        
        # Title - Blue color as in source PDF
        title_style = self.styles['PGTAReportTitle']
        title = Paragraph(
            "Preimplantation Genetic Testing for Aneuploidies (PGT-A)",
            title_style
        )
        elements.append(title)
        elements.append(Spacer(1, 8)) # Reduced spacer
        
        # Patient information table
        patient_table = self._create_patient_info_table(patient_data)
        elements.append(patient_table)
        elements.append(Spacer(1, 12))
        
        # PNDT Disclaimer in a grey box
        disclaimer = Paragraph(
            "<b>This test does not reveal sex of the fetus & confers to PNDT act, 1994</b>",
            self.styles['PGTADisclaimer']
        )
        # Use a single-cell table for the background color (Exact grey from source)
        disclaimer_table = Table([[disclaimer]], colWidths=[self.CONTENT_WIDTH])
        disclaimer_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(self.COLORS['grey_bg'])),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        elements.append(KeepTogether(disclaimer_table))
        elements.append(Spacer(1, 12))
        
        # Indication
        if 'indication' in patient_data and patient_data['indication']:
            elements.append(self._create_section_header("Indication"))
            elements.append(Spacer(1, 8)) # Gap after header line
            indication_text = Paragraph(patient_data['indication'], self.styles['PGTABodyText'])
            elements.append(indication_text)
            elements.append(Spacer(1, 12))
        
        # Results summary header
        elements.append(self._create_section_header("Results summary"))
        elements.append(Spacer(1, 8)) # Gap after header line
        
        # Results summary table
        results_table = self._create_results_summary_table(embryos_data)
        elements.append(results_table)
        elements.append(Spacer(1, 12))
        
        return elements
    
    def _wrap_text(self, text, bold=False, font_size=None, align='LEFT'):
        """Wrap text in a Paragraph for table cells"""
        if not text: return ""
        style = self.styles['PGTABodyText']
        if align == 'CENTER':
            style = self.styles['PGTACenteredBodyText']
        
        # Apply font size override if needed
        content = str(text)
        if font_size:
            content = f'<font size="{font_size}">{content}</font>'
            
        if bold:
            return Paragraph(f"<b>{content}</b>", style)
        return Paragraph(content, style)

    def _create_patient_info_table(self, patient_data):
        """Create patient information table"""
        # Prepare data with Paragraph wrapping to prevent overlap
        data = [
            [self._wrap_text('<b>Patient name</b>', True), self._wrap_text(':'), self._wrap_text(f"<b>{patient_data.get('patient_name', '')}</b>"), self._wrap_text('<b>PIN</b>', True), self._wrap_text(':'), self._wrap_text(f"<b>{patient_data.get('pin', '')}</b>")],
            [self._wrap_text(''), self._wrap_text(''), self._wrap_text(f"<b>{patient_data.get('spouse_name', '')}</b>"), self._wrap_text(''), self._wrap_text(''), self._wrap_text('')],
            [self._wrap_text('<b>Date of Birth/ Age</b>', True), self._wrap_text(':'), self._wrap_text(patient_data.get('age', '')), self._wrap_text('<b>Sample Number</b>', True), self._wrap_text(':'), self._wrap_text(patient_data.get('sample_number', ''))],
            [self._wrap_text('<b>Referring Clinician</b>', True), self._wrap_text(':'), self._wrap_text(patient_data.get('referring_clinician', '')), self._wrap_text('<b>Biopsy date</b>', True), self._wrap_text(':'), self._wrap_text(patient_data.get('biopsy_date', ''))],
            [self._wrap_text('<b>Hospital/Clinic</b>', True), self._wrap_text(':'), self._wrap_text(patient_data.get('hospital_clinic', '')), self._wrap_text('<b>Sample collection date</b>', True), self._wrap_text(':'), self._wrap_text(patient_data.get('sample_collection_date', ''))],
            [self._wrap_text('<b>Specimen</b>', True), self._wrap_text(':'), self._wrap_text(patient_data.get('specimen', '')), self._wrap_text('<b>Sample receipt date</b>', True), self._wrap_text(':'), self._wrap_text(patient_data.get('sample_receipt_date', ''))],
            [self._wrap_text('<b>Report date</b>', True), self._wrap_text(':'), self._wrap_text(patient_data.get('report_date', '')), self._wrap_text('<b>Biopsy performed by</b>', True), self._wrap_text(':'), self._wrap_text(patient_data.get('biopsy_performed_by', ''))]
        ]
        
        # Create table with optimized widths [Total: 496pt]
        table = Table(data, colWidths=[105, 10, 135, 105, 10, 131])
        
        # Style table
        table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), self._get_font('SegoeUI-Bold', 'Helvetica-Bold')),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(self.COLORS['patient_info_bg'])), # Exact light blue-grey from source
        ]))
        
        return table
    
    def _create_results_summary_table(self, embryos_data):
        """Create results summary table"""
        # Header row
        header_labels = ['S. No.', 'Sample', 'Result', 'MTcopy', 'Interpretation']
        data = [[self._wrap_text(label, bold=True, align='CENTER') for label in header_labels]]
        
        # Add embryo rows
        for idx, embryo in enumerate(embryos_data, 1):
            res_sum = embryo.get('result_summary', '')
            interp = embryo.get('interpretation', '')
            
            # Application of Red/Blue color logic
            res_color = self._get_result_color(res_sum, interp)
            
            # MTcopy: NA for non-euploid
            mtcopy = embryo.get('mtcopy', 'NA')
            if interp.upper() != "EUPLOID":
                mtcopy = "NA"
            
            data.append([
                self._wrap_text(str(idx), align='CENTER'),
                self._wrap_text(embryo.get('embryo_id', ''), align='CENTER'),
                # Color only, no bold as per latest request
                self._wrap_text(self._wrap_colored(res_sum, res_color, bold=False), align='CENTER'),
                self._wrap_text(mtcopy, align='CENTER'),
                self._wrap_text(self._wrap_colored(interp, res_color, bold=False), align='CENTER')
            ])
        
        # Create table [Total: 496pt]
        table = Table(data, colWidths=[50, 90, 185, 80, 91])
        
        # Style table
        table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, 0), self._get_font('Calibri-Bold', 'Helvetica-Bold')),
            ('FONTNAME', (0, 1), (-1, -1), self._get_font('Calibri', 'Helvetica')),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            # Header row - peach (exact from source)
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(self.COLORS['results_header_bg'])),
            # First column (S.No) - NO peach as per latest request (only header row)
            # All other data cells - light blue-grey (exact from source)
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor(self.COLORS['patient_info_bg'])),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            # Explicitly ensure center alignment matching request for ALL data cells including S.No
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ]))
        
        return table
    
    def _build_methodology_page(self):
        """Build methodology and static content page"""
        elements = []
        
        # Methodology section
        # Methodology section
        elements.append(self._create_section_header("Methodology"))
        elements.append(Spacer(1, 8)) # Gap after header line
        elements.append(Paragraph(self.METHODOLOGY_TEXT, self.styles['PGTABodyText']))
        elements.append(Spacer(1, 12))
        
        # Mosaicism section
        elements.append(self._create_section_header("Conditions for reporting mosaicism"))
        elements.append(Spacer(1, 8)) # Gap after header line
        elements.append(Paragraph(self.MOSAICISM_TEXT, self.styles['PGTABodyText']))
        elements.append(Spacer(1, 6)) # Added spacing between intro and bullets
        
        # Mosaicism bullets
        for bullet in self.MOSAICISM_BULLETS:
            elements.append(Paragraph(f"• {bullet}", self.styles['PGTABulletText']))
        elements.append(Spacer(1, 6)) # Added spacing before clinical text
        elements.append(Paragraph(self.MOSAICISM_CLINICAL, self.styles['PGTABodyText']))
        elements.append(Spacer(1, 12)) # Increased spacing between sections
        
        # Limitations section
        elements.append(self._create_section_header("Limitations"))
        elements.append(Spacer(1, 8)) # Gap after header line
        for limitation in self.LIMITATIONS:
            elements.append(Paragraph(f"• {limitation}", self.styles['PGTABulletText']))
        
        elements.append(Spacer(1, 12))
        
        elements.append(Spacer(1, 12))
        
        # References section
        elements.append(CondPageBreak(60)) # Only break if < ~0.8 inch of space remains (header protection)
        elements.append(self._create_section_header("References"))
        elements.append(Spacer(1, 8)) # Gap after header line
        for idx, ref in enumerate(self.REFERENCES, 1):
            # Using PGTABodyText logic (11pt) instead of small text
            elements.append(Paragraph(f"{idx}. {ref}", self.styles['PGTABodyText']))
        
        return elements
    
    def _build_embryo_page(self, patient_data, embryo_data):
        """Build individual embryo results page"""
        elements = []
        
        # Title repeated in embryo section as requested
        title = Paragraph(
            "Preimplantation Genetic Testing for Aneuploidies (PGT-A)",
            self.styles['PGTAReportTitle']
        )
        elements.append(title)
        elements.append(Spacer(1, 8))
        
        # Patient info line in a table as in source
        # Using 6-column grid logic from cover page to handle long names
        info_data = [[
            self._wrap_text('<b>Patient name</b>', True),
            self._wrap_text(':'),
            self._wrap_text(f"<b>{patient_data.get('patient_name', '')} {patient_data.get('spouse_name', '')}</b>"),
            self._wrap_text('<b>PIN</b>', True),
            self._wrap_text(':'),
            self._wrap_text(f"<b>{patient_data.get('pin', '')}</b>")
        ]]
        
        # Optimized widths for detailed banner [Total: 496pt] - Tightened to remove gaps
        info_table = Table(info_data, colWidths=[72, 8, 170, 30, 8, 208])
        info_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(self.COLORS['patient_info_bg'])),
            ('FONTNAME', (0, 0), (-1, -1), self._get_font('SegoeUI-Bold', 'Helvetica-Bold')),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        elements.append(info_table)
        elements.append(Spacer(1, 8))
        
        # PNDT Disclaimer in a grey box (Exact grey from source)
        disclaimer = Paragraph(
            "<b>This test does not reveal sex of the fetus & confers to PNDT act, 1994</b>",
            self.styles['PGTADisclaimer']
        )
        disclaimer_table = Table([[disclaimer]], colWidths=[self.CONTENT_WIDTH])
        disclaimer_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(self.COLORS['grey_bg'])),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        elements.append(KeepTogether(disclaimer_table))
        elements.append(Spacer(1, 12))
        
        # Application of Red/Blue color logic
        res_text = embryo_data.get('result_description', '')
        autosomes_text = embryo_data.get('autosomes', '')
        interp_text = embryo_data.get('interpretation', '')
        
        # Color based on interpretation for interpretation field
        # Color based on keywords for autosomes field
        interp_color = self._get_result_color('', interp_text)
        auto_color = self._get_result_color(autosomes_text, '')
        
        # MTcopy: NA for non-euploid
        mtcopy = embryo_data.get('mtcopy', 'NA')
        if interp_text.upper() != "EUPLOID":
            mtcopy = "NA"
            
        # Embryo Identification matching source style
        # Font: Gill Sans MT,Bold, Size: 12.00
        embryo_id_style = ParagraphStyle(
            name='EmbryoIDStyle',
            parent=self.styles['Normal'],
            fontSize=12,
            leading=14,
            fontName=self._get_font('GillSansMT-Bold', 'Helvetica-Bold'),
            textColor=colors.HexColor(self.COLORS['blue_title'])
        )
        elements.append(Paragraph(f"<b>EMBRYO: {embryo_data.get('embryo_id', '')}</b>", embryo_id_style))
        elements.append(Spacer(1, 6))
        
        detail_data = [
            [self._wrap_text(f"<b>Result:</b> {self._wrap_colored(res_text, self._get_result_color(res_text, interp_text), bold=False)}", False)],
            [self._wrap_text(f"<b>Autosomes:</b> {self._wrap_colored(autosomes_text, auto_color, bold=False)}", False)],
            [self._wrap_text(f"<b>Sex chromosomes:</b> {embryo_data.get('sex_chromosomes', '')}", False)],
            [self._wrap_text(f"<b>Interpretation:</b> {self._wrap_colored(interp_text, interp_color, bold=False)}", False)],
            [self._wrap_text(f"<b>MTcopy:</b> {mtcopy}", False)]
        ]
        
        detail_table = Table(detail_data, colWidths=[self.CONTENT_WIDTH])
        detail_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(self.COLORS['patient_info_bg'])),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('RIGHTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(detail_table)
        elements.append(Spacer(1, 12))
        
        # CNV Chart title - No line for this one
        elements.append(self._create_section_header("COPY NUMBER CHART", show_line=False))
        elements.append(Spacer(1, 6))
        
        # CNV Chart Image
        if 'cnv_image_path' in embryo_data and embryo_data['cnv_image_path'] and os.path.exists(embryo_data['cnv_image_path']):
            try:
                # Add image, keeping aspect ratio but fitting within width
                img = Image(embryo_data['cnv_image_path'], width=self.CONTENT_WIDTH)
                
                # Adjust height proportionally
                aspect = img.imageWidth / img.imageHeight
                img.drawHeight = self.CONTENT_WIDTH / aspect
                
                img.hAlign = 'CENTER'
                elements.append(img)
                elements.append(Spacer(1, 12))
            except Exception as e:
                print(f"Error loading image: {e}")
        
        # CNV table
        cnv_table = self._create_cnv_table(embryo_data)
        elements.append(cnv_table)
        elements.append(Spacer(1, 6))
        
        # Legend
        legend = Paragraph(
            "<i>N – Normal, G-Gain, L-Loss, SG-Segmental Gain, SL-Segmental Loss, "
            "M-Mosaic, MG- Mosaic Gain, ML-Mosaic Loss, SMG-Segmental Mosaic Gain, "
            "SML-Segmental Mosaic Loss</i>",
            self.styles['PGTASmallText']
        )
        elements.append(legend)
        elements.append(Spacer(1, 12))
        
        return elements
    
    def _create_cnv_table(self, embryo_data):
        """Create CNV status table"""
        # Get chromosome statuses
        chr_statuses = embryo_data.get('chromosome_statuses', {})
        mosaic_percentages = embryo_data.get('mosaic_percentages', {})
        
        # Build header
        has_mosaic = any(mosaic_percentages.values())
        
        if has_mosaic:
            # Header set to 9pt
            header = [self._wrap_text('Chromosome', bold=True, align='CENTER', font_size=9)] + [self._wrap_text(str(i), bold=True, align='CENTER', font_size=9) for i in range(1, 23)]
            cnv_row = [self._wrap_text('CNV status', bold=True, align='CENTER', font_size=9)]
            mosaic_row = [self._wrap_text('Mosaic (%)', bold=True, align='CENTER', font_size=9)]
            
            for i in range(1, 23):
                status = chr_statuses.get(str(i), 'N')
                perc = mosaic_percentages.get(str(i), '-')
                s_color = self._get_status_color(status)
                
                # Dynamic font sizing for balancing visibility (Normal: 9pt, SMG/SML: 8pt)
                f_size = 8 if len(status) > 2 else 9
                cnv_row.append(self._wrap_text(self._wrap_colored(status, s_color, bold=True), bold=True, font_size=f_size, align='CENTER'))
                mosaic_row.append(self._wrap_text(self._wrap_colored(str(perc), s_color, bold=True), bold=True, font_size=9, align='CENTER'))
            
            data = [header, cnv_row, mosaic_row]
            # Final optimized width: "Chromosome" widened to 75pt to ensure NO wrap. 
            # Remaining width (496 - 75 = 421) / 22 columns = ~19.13pt per data column
            col_widths = [75] + [19.13] * 22
        else:
            header = [self._wrap_text('Chromosome', bold=True, align='CENTER', font_size=9)] + [self._wrap_text(str(i), bold=True, align='CENTER', font_size=9) for i in range(1, 23)]
            cnv_row = [self._wrap_text('CNV status', bold=True, align='CENTER', font_size=9)]
            for i in range(1, 23):
                status = chr_statuses.get(str(i), 'N')
                s_color = self._get_status_color(status)
                
                # Dynamic font sizing for balancing visibility (Normal: 9pt, SMG/SML: 8pt)
                f_size = 8 if len(status) > 2 else 9
                cnv_row.append(self._wrap_text(self._wrap_colored(status, s_color, bold=True), bold=True, font_size=f_size, align='CENTER'))
                
            data = [header, cnv_row]
            col_widths = [75] + [19.13] * 22
        
        # Create table
        table = Table(data, colWidths=col_widths)
        
        # Style table
        table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), self._get_font('SegoeUI-Bold', 'Helvetica-Bold')),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            # All cells - light blue-grey (exact from source) as requested
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(self.COLORS['patient_info_bg'])),
            # White grid lines for "white line gaps"
            # White grid lines for "white line gaps"
            ('GRID', (0, 0), (-1, -1), 1.0, colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            # Zero padding for data columns to maximize horizontal space for 8pt/9pt fonts
            ('LEFTPADDING', (1, 0), (-1, -1), 0),
            ('RIGHTPADDING', (1, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        
        return table
    
    def _create_signature_table(self):
        """Create signature section with precise structural metrics from source PDF"""
        elements = []
        
        # Leading text: exact color #4F81BD, size 12.48pt
        elements.append(Paragraph(
            "<b>This report has been reviewed and approved by: </b>", 
            self.styles['PGTASigApproval']
        ))
        # Exact source vertical gap (12.7pt from text baseline to image top)
        elements.append(Spacer(1, 12.7))
        
        # Signatures image: exact source dimensions (395pt x 42pt)
        sig_image_path = os.path.join(self.ASSETS_DIR, "signs.png")
        if os.path.exists(sig_image_path):
            try:
                img_w = 395
                img_h = 42
                img = Image(sig_image_path, width=img_w, height=img_h)
                img.hAlign = 'CENTER'
                elements.append(img)
            except Exception:
                pass

        # Names and Titles below signatures - Using SegoeUI Normal weight as in source
        # Source Centers: 152.1, 307.7, 463.1 (Page Center approx 306)
        # Using a centered 3-column table with 156pt columns (centers at 78, 234, 390 relative to table)
        # Table width: 468 (total content width)
        data = []
        names_row = []
        titles_row = []
        
        # Use Normal weight for names as identified in source extraction (Bold: False)
        sig_name_style = ParagraphStyle('SigName', parent=self.styles['Normal'], 
                                       fontName=self._get_font('SegoeUI', 'Helvetica'), 
                                       fontSize=11.04, alignment=TA_CENTER)
        sig_title_style = ParagraphStyle('SigTitle', parent=self.styles['Normal'], 
                                        fontName=self._get_font('SegoeUI', 'Helvetica'), 
                                        fontSize=11.04, alignment=TA_CENTER)

        for sig in self.SIGNATURES:
            names_row.append(Paragraph(sig['name'], sig_name_style))
            titles_row.append(Paragraph(sig['title'], sig_title_style))
        
        data = [names_row, titles_row]
        
        # Table width 468, 3 columns of 156
        table = Table(data, colWidths=[156, 156, 156])
        table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('TOPPADDING', (0, 0), (-1, -1), 1),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ]))
        
        elements.append(table)
        
        return KeepTogether(elements)

    def _create_section_header(self, text, show_line=True):
        """Create a section header with navy blue text and a slight lighter line below"""
        header = Paragraph(f"<b>{text}</b>", self.styles["PGTASectionHeader"])
        
        if not show_line:
            return KeepTogether([header])
            
        # Create a single-cell table for the line
        header_table = Table([[header]], colWidths=[self.CONTENT_WIDTH])
        header_table.setStyle(TableStyle([
            # Slight grey line, lighter than full black
            ("LINEBELOW", (0, 0), (-1, -1), 0.5, colors.HexColor("#989998")),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            # Increased bottom padding to 6pt as requested for visual gap between text and line
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ]))
        return KeepTogether(header_table)

    def _get_result_color(self, result_text, interpretation_text):
        """Determine if text should be Red (Aneuploid), Blue (Mosaic) or Black"""
        res_up = result_text.upper() if result_text else ""
        int_up = interpretation_text.upper() if interpretation_text else ""
        
        # Red Logic
        red_keywords = ["MONOSOMY", "TRISOMY", "SEGMENTAL GAIN", "SEGMENTAL LOSS", 
                        "MULTIPLE CHROMOSOMAL ABNORMALITIES", "ANEUPLOID", "CHAOTIC EMBRYO"]
        if any(kw in res_up for kw in red_keywords) or any(kw in int_up for kw in red_keywords):
            return colors.red
            
        # Blue Logic
        blue_keywords = ["MOSAIC CHROMOSOME COMPLEMENT", "LOW LEVEL MOSAIC", 
                         "HIGH LEVEL MOSAIC", "COMPLEX MOSAIC", "MULTIPLE MOSAIC"]
        if any(kw in res_up for kw in blue_keywords) or any(kw in int_up for kw in blue_keywords):
            return colors.blue
            
        return colors.black

    def _get_autosome_color(self, autosome_text):
        """Special color logic for autosomes field"""
        if not autosome_text: return colors.black
        txt = autosome_text.upper()
        if "MULTIPLE MOSAIC CHROMOSOME COMPLEMENT" in txt:
            return colors.blue
        return colors.black

    def _get_status_color(self, status):
        """Color logic for CNV status codes"""
        if not status: return colors.black
        s = status.upper().strip()
        red_codes = ["G", "L", "SG", "SL"]
        blue_codes = ["M", "MG", "ML", "SMG", "SML"]
        
        if s in red_codes: return colors.red
        if s in blue_codes: return colors.blue
        
        # Numeric check for mosaic percentage
        try:
            val = float(s.replace('%', ''))
            if val > 0: return colors.blue
        except:
            pass
            
        return colors.black

    def _wrap_colored(self, text, color, bold=False):
        """Standard wrapper for colored text with optional bolding"""
        if not text: return text
        if color == colors.black:
            return f"<b>{text}</b>" if bold else str(text)
        hex_color = color.hexval()[2:] # Remove 0x
        if len(hex_color) > 6: hex_color = hex_color[2:] # Handle alpha
        
        if bold:
            return f'<b><font color="#{hex_color}">{text}</font></b>'
        return f'<font color="#{hex_color}">{text}</font>'


if __name__ == "__main__":
    # Test the template
    template = PGTAReportTemplate()
    
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
            'cnv_image_path': os.path.join(os.path.dirname(os.path.abspath(__file__)), "PRIYA-PS4_L00_R1_noXY_nomos.png")
        }
    ]
    
    # Update chromosome 16 to Segmental Mosaic Loss (SML) to test font scaling
    embryos_data[0]['chromosome_statuses']['16'] = 'SML'
    
    # Generate PDF
    output_path = "test_report.pdf"
    template.generate_pdf(output_path, patient_data, embryos_data)
    print(f"Test report generated: {output_path}")
