"""
PGT-A Report Generator - Desktop Application
Main GUI application for generating PGT-A reports
"""

import sys
import os
import json
from datetime import datetime
from pathlib import Path
import subprocess
import platform

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QLineEdit, QTextEdit, QPushButton, QFileDialog,
    QTableWidget, QTableWidgetItem, QMessageBox, QProgressBar,
    QGroupBox, QFormLayout, QScrollArea, QCheckBox, QSpinBox,
    QComboBox, QListWidget, QListWidgetItem, QStyle, QGridLayout,
    QSplitter, QTextBrowser, QRadioButton, QDialog, QDialogButtonBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSettings, QTimer
from PyQt6.QtGui import QPixmap, QIcon, QColor, QBrush
try:
    from PyQt6.QtPdf import QPdfDocument
    from PyQt6.QtPdfWidgets import QPdfView
except ImportError:
    print("Warning: QtPdf module not found. Preview may not work.")
    QPdfDocument = None
    QPdfView = None

import pandas as pd
# Add templates directory to path to import modular generators (currently only PGT-A)
# We will keep the structure but use ReportLab classes
from pgta_template import PGTAReportTemplate
from pgta_docx_generator import PGTADocxGenerator

# TRF Verification imports
try:
    import pytesseract
    from PIL import Image
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False
    print("Warning: pytesseract not found. TRF image verification may not work.")

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False
    print("Warning: pdfplumber not found. TRF PDF verification may not work.")

# EasyOCR - Simple, accurate, works offline (RECOMMENDED)
try:
    import easyocr
    EASYOCR_AVAILABLE = True
except ImportError:
    EASYOCR_AVAILABLE = False
    print("Info: easyocr not found. Install for better OCR: pip install easyocr")

# Ollama - Local LLM with vision (LLaVA model)
try:
    import requests
    OLLAMA_AVAILABLE = True
except ImportError:
    OLLAMA_AVAILABLE = False

import base64
import re
from difflib import SequenceMatcher
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

class ClickOnlyComboBox(QComboBox):
    """Subclass of QComboBox that ignores mouse wheel events to prevent accidental changes when scrolling."""
    def wheelEvent(self, event):
        event.ignore()


def add_colored_items_to_combo(combo, items_with_colors):
    """
    Add items with specific colors to a combo box.
    items_with_colors: list of tuples (text, color) where color is 'black', 'red', or 'blue'
    """
    color_map = {
        'black': QColor(0, 0, 0),
        'red': QColor(255, 0, 0),
        'blue': QColor(0, 0, 255)
    }
    
    for text, color_name in items_with_colors:
        combo.addItem(text)
        index = combo.count() - 1
        combo.setItemData(index, QBrush(color_map.get(color_name, QColor(0, 0, 0))), Qt.ItemDataRole.ForegroundRole)


class PreviewWorker(QThread):
    """Worker thread for generating preview PDF"""
    finished = pyqtSignal(str) # Path to generated PDF
    error = pyqtSignal(str)
    
    def __init__(self, patient_data, embryos_data, output_path, show_logo=True):
        super().__init__()
        self.patient_data = patient_data
        self.embryos_data = embryos_data
        self.output_path = output_path
        self.show_logo = show_logo
        
    def run(self):
        try:
            # Generate PDF using native template
            gen = PGTAReportTemplate(assets_dir="assets/pgta")
            gen.generate_pdf(self.output_path, self.patient_data, self.embryos_data, show_logo=self.show_logo)
            self.finished.emit(self.output_path)
        except Exception as e:
            self.error.emit(str(e))


class ReportGeneratorWorker(QThread):
    """Worker thread for generating reports"""
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(list, list)
    error = pyqtSignal(str)
    
    def __init__(self, patient_data_list, output_dir, generate_pdf=True, generate_docx=True, template_type="PGT-A", show_logo=True):
        super().__init__()
        self.patient_data_list = patient_data_list
        self.output_dir = output_dir
        self.generate_pdf = generate_pdf
        self.generate_docx = generate_docx
        self.template_type = template_type
        self.show_logo = show_logo
    
    def run(self):
        """Generate reports"""
        success_reports = []
        failed_reports = []
        
        if self.generate_pdf:
            # Revert to pure ReportLab template
            # In a full PGT-A system, we would select the class based on template_type
            if self.template_type == "PGT-A":
                pdf_generator = PGTAReportTemplate(assets_dir="assets/pgta")
            else:
                pdf_generator = PGTAReportTemplate(assets_dir="assets/pgta")
        else:
            pdf_generator = None
            
        if self.generate_docx:
            docx_generator = PGTADocxGenerator(assets_dir="assets/pgta")
        else:
            docx_generator = None
        
        total = len(self.patient_data_list)
        
        for idx, patient_data in enumerate(self.patient_data_list, 1):
            try:
                sample_num = patient_data['patient_info'].get('sample_number', f'Sample_{idx}')
                patient_name = patient_data['patient_info'].get('patient_name', 'Unknown')
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                logo_suffix = "_withlogo" if self.show_logo else "_withoutlogo"
                base_filename = f"{sample_num}_{patient_name.replace(' ', '_')}_{timestamp}{logo_suffix}"
                
                self.progress.emit(
                    int((idx - 1) / total * 100),
                    f"Generating reports for {patient_name} ({idx}/{total})..."
                )
                
                # Generate using Original Template
                # Generate PDF
                if self.generate_pdf:
                    pdf_path = os.path.join(self.output_dir, f"{base_filename}.pdf")
                    # Use the ReportLab generator directly
                    # assets_dir is handled internally by the generator classes using resource_path
                    pdf_generator.generate_pdf(
                        pdf_path,
                        patient_data['patient_info'],
                        patient_data['embryos'],
                        show_logo=self.show_logo
                    )
                
                # Generate DOCX
                if self.generate_docx:
                    docx_path = os.path.join(self.output_dir, f"{base_filename}.docx")
                    docx_generator.generate_docx(
                        docx_path,
                        patient_data['patient_info'],
                        patient_data['embryos'],
                        show_logo=self.show_logo
                    )
                
                success_reports.append(base_filename)
                
            except Exception as e:
                failed_reports.append((patient_data['patient_info'].get('patient_name', 'Unknown'), str(e)))
                self.error.emit(f"Error generating report for {patient_data['patient_info'].get('patient_name', 'Unknown')}: {str(e)}")
        
        self.progress.emit(100, "Report generation complete!")
        self.finished.emit(success_reports, failed_reports)


class PGTAReportGeneratorApp(QMainWindow):
    """Main application window"""
    
    def __init__(self):
        super().__init__()
        self.settings = QSettings('PGTA', 'ReportGenerator')
        self.current_patient_data = {} # For manual entry
        self.bulk_patient_data_list = [] # For bulk upload
        self.current_embryos = []
        self.uploaded_images = {}
        
        self.init_ui()
        self.load_settings()
    
        # ... (skipping some methods) ...

    def update_data_summary(self):
        """Update data summary display"""
        summary = ""
        
        # Check manual data
        if self.current_patient_data:
            p_name = self.current_patient_data['patient_info'].get('patient_name', 'Unknown')
            e_count = len(self.current_patient_data['embryos'])
            summary += f"Manual Entry: {p_name} ({e_count} embryos)\n"
            
        # Check bulk data
        if self.bulk_patient_data_list:
            p_count = len(self.bulk_patient_data_list)
            e_count = sum(len(p['embryos']) for p in self.bulk_patient_data_list)
            summary += f"Bulk Data: {p_count} patients ({e_count} total embryos)\n"
            
        if not summary:
            summary = "No data loaded"
            
        self.data_summary_label.setText(summary)

    def generate_reports(self):
        """Generate reports"""
        # Determine source
        patient_data_list = []
        
        # If we have bulk data, prioritize that? Or combine?
        # Let's combine if both exist, or use whichever is available
        
        if self.bulk_patient_data_list:
            patient_data_list.extend(self.bulk_patient_data_list)
            
        if self.current_patient_data:
            # Check if this patient is already in bulk list by Sample Number
            current_sn = self.current_patient_data.get('patient_info', {}).get('sample_number')
            is_duplicate = any(p.get('patient_info', {}).get('sample_number') == current_sn for p in patient_data_list)
            
            if not is_duplicate:
                patient_data_list.append(self.current_patient_data)
            else:
                # If duplicate, we prioritize manual entry as it's likely a correction
                patient_data_list = [p for p in patient_data_list if p.get('patient_info', {}).get('sample_number') != current_sn]
                patient_data_list.append(self.current_patient_data)
            
        # Validate
        if not patient_data_list:
            QMessageBox.warning(self, "Warning", "No data loaded. Please enter data or upload a file.")
            return
        
        output_dir = self.output_dir_label.text()
        if output_dir == "No directory selected":
            QMessageBox.warning(self, "Warning", "Please select an output directory.")
            return
        
        if not self.generate_pdf_check.isChecked() and not self.generate_docx_check.isChecked():
            QMessageBox.warning(self, "Warning", "Please select at least one output format.")
            return
        
        # Start generation
        self.progress_bar.setVisible(True)
        self.progress_label.setVisible(True)
        self.generate_btn.setEnabled(False)
        
        # Create worker thread
        self.worker = ReportGeneratorWorker(
            patient_data_list,
            output_dir,
            self.generate_pdf_check.isChecked(),
            self.generate_docx_check.isChecked(),
            template_type=self.template_combo.currentText(),
            show_logo=(self.logo_combo.currentText() == "With Logo")
        )
        
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.generation_finished)
        self.worker.error.connect(self.generation_error)
        
        self.worker.start()
    
    def init_ui(self):
        """Initialize user interface"""
        self.setWindowTitle("PGT-A Report Generator")
        self.setGeometry(100, 100, 1200, 800)
        
        # Set Application Icon
        icon_path = resource_path(os.path.join("assets", "pgta", "app_icon.png"))
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # Header layout with title and template selector
        header_layout = QHBoxLayout()
        
        # Title
        title_label = QLabel("PGT-A Report Generator")
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; padding: 10px;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        header_layout.addWidget(title_label)
        
        header_layout.addStretch()
        
        # Template selector
        header_layout.addWidget(QLabel("Select Template:"))
        self.template_combo = ClickOnlyComboBox()
        self.template_combo.addItems(["PGT-A", "PGT-M (Coming Soon)", "NIPT (Coming Soon)"])
        self.template_combo.setMinimumWidth(200)
        header_layout.addWidget(self.template_combo)
        
        main_layout.addLayout(header_layout)
        
        # Tab widget
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Create tabs
        self.manual_entry_tab = self.create_manual_entry_tab()
        self.bulk_upload_tab = self.create_bulk_upload_tab()
        self.user_guide_tab = self.create_user_guide_tab()
        
        self.tabs.addTab(self.manual_entry_tab, "Manual Entry")
        self.tabs.setTabIcon(0, self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))
        
        self.tabs.addTab(self.bulk_upload_tab, "Bulk Upload")
        self.tabs.setTabIcon(1, self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogListView))
        
        self.tabs.addTab(self.user_guide_tab, "User Guide")
        self.tabs.setTabIcon(2, self.style().standardIcon(QStyle.StandardPixmap.SP_MessageBoxInformation))
        
        # Status bar
        self.statusBar().showMessage("Ready")
    
    def create_manual_entry_tab(self):
        """Create manual entry_tab with split preview"""
        tab = QWidget()
        main_layout = QHBoxLayout()
        tab.setLayout(main_layout)
        
        # Splitter
        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)
        
        # --- Left Panel: Input Form ---
        left_widget = QWidget()
        layout = QVBoxLayout()
        left_widget.setLayout(layout)
        
        # Scroll area for form
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout()
        scroll_widget.setLayout(scroll_layout)
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        
        # Patient Information Group
        patient_group = QGroupBox("Patient Information")
        patient_form = QFormLayout()
        patient_group.setLayout(patient_form)
        scroll_layout.addWidget(patient_group)
        
        # Create input fields
        self.patient_name_input = QTextEdit()
        self.patient_name_input.setMaximumHeight(45)
        self.spouse_name_input = QTextEdit()
        self.spouse_name_input.setMaximumHeight(45)
        self.spouse_name_input.setPlaceholderText("w/o")
        self.spouse_name_input.setText("w/o")  # Default value, editable by user
        self.pin_input = QTextEdit()
        self.pin_input.setMaximumHeight(40)
        self.age_input = QTextEdit()
        self.age_input.setMaximumHeight(40)
        self.sample_number_input = QTextEdit()
        self.sample_number_input.setMaximumHeight(40)
        self.referring_clinician_input = QTextEdit()
        self.referring_clinician_input.setMaximumHeight(40)
        self.biopsy_date_input = QLineEdit()
        
        self.biopsy_date_input.setPlaceholderText("DD-MM-YYYY")
        self.hospital_clinic_input = QTextEdit()
        self.hospital_clinic_input.setMaximumHeight(45)
        self.sample_collection_date_input = QLineEdit()
        self.sample_collection_date_input.setPlaceholderText("DD-MM-YYYY")
        self.specimen_input = QTextEdit()
        self.specimen_input.setMaximumHeight(40)
        self.specimen_input.setText("Day 6 Trophectoderm Biopsy")
        self.sample_receipt_date_input = QLineEdit()
        self.sample_receipt_date_input.setPlaceholderText("DD-MM-YYYY")
        self.biopsy_performed_by_input = QTextEdit()
        self.biopsy_performed_by_input.setMaximumHeight(40)
        self.report_date_input = QLineEdit()
        self.report_date_input.setPlaceholderText("DD-MM-YYYY")
        self.report_date_input.setText(datetime.now().strftime("%d-%m-%Y"))
        self.indication_input = QTextEdit()
        self.indication_input.setMaximumHeight(60)
        
        # Add fields to form
        patient_form.addRow("Patient Name:", self.patient_name_input)
        patient_form.addRow("Spouse Name:", self.spouse_name_input)
        patient_form.addRow("PIN:", self.pin_input)
        patient_form.addRow("Age:", self.age_input)
        patient_form.addRow("Sample Number:", self.sample_number_input)
        patient_form.addRow("Referring Clinician:", self.referring_clinician_input)
        patient_form.addRow("Biopsy Date:", self.biopsy_date_input)
        patient_form.addRow("Hospital/Clinic:", self.hospital_clinic_input)
        patient_form.addRow("Sample Collection Date:", self.sample_collection_date_input)
        patient_form.addRow("Specimen:", self.specimen_input)
        patient_form.addRow("Sample Receipt Date:", self.sample_receipt_date_input)
        patient_form.addRow("Biopsy Performed By:", self.biopsy_performed_by_input)
        patient_form.addRow("Report Date:", self.report_date_input)
        patient_form.addRow("Indication:", self.indication_input)
        
        # --- TRF Verification Section ---
        trf_group = QGroupBox("TRF Verification (Optional)")
        trf_layout = QVBoxLayout()
        trf_group.setLayout(trf_layout)
        scroll_layout.addWidget(trf_group)
        
        trf_upload_layout = QHBoxLayout()
        self.trf_path_label = QLabel("No TRF uploaded")
        self.trf_path_label.setStyleSheet("color: #666; font-style: italic;")
        trf_upload_layout.addWidget(self.trf_path_label, 1)
        
        self.trf_upload_btn = QPushButton("📄 Upload TRF")
        self.trf_upload_btn.clicked.connect(self.upload_trf_manual)
        trf_upload_layout.addWidget(self.trf_upload_btn)
        
        self.trf_verify_btn = QPushButton("✓ Verify")
        self.trf_verify_btn.clicked.connect(self.verify_trf_manual)
        self.trf_verify_btn.setEnabled(False)
        trf_upload_layout.addWidget(self.trf_verify_btn)
        
        trf_layout.addLayout(trf_upload_layout)
        
        # Verification result display
        self.trf_result_text = QTextBrowser()
        self.trf_result_text.setMaximumHeight(120)
        self.trf_result_text.setStyleSheet("background-color: #f8f9fa; border: 1px solid #ddd; border-radius: 4px;")
        self.trf_result_text.setHtml("<i style='color:#888;'>Upload a TRF (image or PDF) to verify patient details</i>")
        trf_layout.addWidget(self.trf_result_text)
        
        # Store TRF path
        self.manual_trf_path = None
        
        # --- Page 1 Summary Section ---
        summary_group = QGroupBox("Page 1: Results Summary")
        summary_layout = QVBoxLayout()
        summary_group.setLayout(summary_layout)
        scroll_layout.addWidget(summary_group)
        
        self.summary_table = QTableWidget()
        self.summary_table.setColumnCount(4)
        self.summary_table.setHorizontalHeaderLabels(["Embryo ID", "Result (Summary)", "Interpretation", "MTcopy"])
        self.summary_table.horizontalHeader().setStretchLastSection(True)
        # Column widths
        self.summary_table.setColumnWidth(0, 80)  # ID
        self.summary_table.setColumnWidth(1, 230) # Summary
        self.summary_table.setColumnWidth(2, 110) # Interp
        self.summary_table.setColumnWidth(3, 70) # MTcopy
        self.summary_table.setAlternatingRowColors(True)
        # We'll update rows in update_embryo_forms
        summary_layout.addWidget(self.summary_table)
        
        # Embryo Management Group (Page 4+)
        embryo_group = QGroupBox("Page 4+: Embryo Details (Detail View)")
        embryo_layout = QVBoxLayout()
        embryo_group.setLayout(embryo_layout)
        scroll_layout.addWidget(embryo_group)
        
        # Embryo count selector
        embryo_count_layout = QHBoxLayout()
        embryo_count_layout.addWidget(QLabel("Number of Embryos:"))
        self.embryo_count_spin = QSpinBox()
        self.embryo_count_spin.setMinimum(1)
        self.embryo_count_spin.setMaximum(20)
        self.embryo_count_spin.setValue(1)
        self.embryo_count_spin.valueChanged.connect(self.update_embryo_forms)
        embryo_count_layout.addWidget(self.embryo_count_spin)
        
        # New: Copy Last Embryo Button
        self.copy_embryo_btn = QPushButton("Copy Last Embryo")
        self.copy_embryo_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))
        self.copy_embryo_btn.clicked.connect(self.copy_last_embryo)
        self.copy_embryo_btn.setToolTip("Create a new embryo entry with data copied from the current last one")
        embryo_count_layout.addWidget(self.copy_embryo_btn)
        
        embryo_count_layout.addStretch()
        embryo_layout.addLayout(embryo_count_layout)
        
        # Embryo forms container
        self.embryo_forms_container = QWidget()
        self.embryo_forms_layout = QVBoxLayout()
        self.embryo_forms_container.setLayout(self.embryo_forms_layout)
        embryo_layout.addWidget(self.embryo_forms_container)
        
        # Initialize with one embryo form (Moved to end)
        self.embryo_forms = []
        
        # Buttons
        button_layout = QHBoxLayout()
        
        save_draft_btn = QPushButton("Save Draft")
        save_draft_btn.clicked.connect(self.save_draft)
        
        load_draft_btn = QPushButton("Load Draft")
        load_draft_btn.clicked.connect(self.load_draft)
        
        save_btn = QPushButton("Save Data")
        save_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton))
        save_btn.clicked.connect(self.save_manual_data)
        
        clear_btn = QPushButton("Clear Form")
        clear_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon))
        clear_btn.clicked.connect(self.clear_manual_form)
        
        button_layout.addWidget(save_draft_btn)
        button_layout.addWidget(load_draft_btn)
        button_layout.addWidget(save_btn)
        button_layout.addWidget(clear_btn)
        button_layout.addStretch()
        layout.addLayout(button_layout)
        
        # --- NEW: Report Generation Tools ---
        gen_group = QGroupBox("Fast Report Generation")
        gen_layout = QVBoxLayout()
        gen_group.setLayout(gen_layout)
        
        # Output directory selection
        out_row = QHBoxLayout()
        self.output_dir_label = QLabel("No directory selected")
        self.output_dir_label.setStyleSheet("padding: 5px; border: 1px solid #ccc; background: white;")
        browse_out_btn = QPushButton("Select Output Folder")
        browse_out_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton))
        browse_out_btn.clicked.connect(self.browse_output_dir)
        out_row.addWidget(self.output_dir_label, 1)
        out_row.addWidget(browse_out_btn)
        gen_layout.addLayout(out_row)
        
        # Formats and Generate Button
        action_row = QHBoxLayout()
        self.generate_pdf_check = QCheckBox("PDF")
        self.generate_pdf_check.setChecked(True)
        self.generate_docx_check = QCheckBox("DOCX")
        self.generate_docx_check.setChecked(True)
        
        self.generate_btn = QPushButton("Generate Report(s)")
        self.generate_btn.setStyleSheet("background-color: #1F497D; color: white; font-weight: bold; padding: 8px;")
        self.generate_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
        self.generate_btn.clicked.connect(self.generate_reports)
        
        action_row.addWidget(QLabel("Output Forms:"))
        action_row.addWidget(self.generate_pdf_check)
        action_row.addWidget(self.generate_docx_check)
        
        # Logo Preference
        action_row.addSpacing(20)
        action_row.addWidget(QLabel("Branding:"))
        self.logo_combo = ClickOnlyComboBox()
        self.logo_combo.addItems(["With Logo", "Without Logo"])
        self.logo_combo.currentIndexChanged.connect(self.update_preview)
        action_row.addWidget(self.logo_combo)
        
        action_row.addStretch()
        action_row.addWidget(self.generate_btn)
        gen_layout.addLayout(action_row)
        
        # Progress Bar inside Manual Tab
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_label = QLabel("")
        self.progress_label.setVisible(False)
        
        # Restore data summary label to fix crash
        self.data_summary_label = QLabel("No data loaded")
        self.data_summary_label.setStyleSheet("padding: 5px; color: #666; font-style: italic;")
        
        gen_layout.addWidget(self.progress_label)
        gen_layout.addWidget(self.progress_bar)
        gen_layout.addWidget(self.data_summary_label)
        
        layout.addWidget(gen_group)
        
        splitter.addWidget(left_widget)
        
        # --- Right Panel: Preview ---
        right_widget = QGroupBox("Report Preview")
        right_layout = QVBoxLayout()
        right_widget.setLayout(right_layout)
        
        preview_label = QLabel("Report Preview (PDF)")
        preview_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        right_layout.addWidget(preview_label)
        
        if QPdfView and QPdfDocument:
            self.pdf_document = QPdfDocument(self)
            self.pdf_view = QPdfView(self)
            self.pdf_view.setDocument(self.pdf_document)
            self.pdf_view.setPageMode(QPdfView.PageMode.MultiPage)
            right_layout.addWidget(self.pdf_view)
        else:
            self.pdf_view = QLabel("PDF Preview not available (QtPdf missing)")
            self.pdf_view.setAlignment(Qt.AlignmentFlag.AlignCenter)
            right_layout.addWidget(self.pdf_view)
        
        refresh_btn = QPushButton("Refresh Preview")
        refresh_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))
        refresh_btn.clicked.connect(self.update_preview)
        right_layout.addWidget(refresh_btn)
        
        splitter.addWidget(right_widget)
        splitter.setSizes([600, 400])
        
        # Connect signals for live preview update
        for field in [self.patient_name_input, self.spouse_name_input, self.pin_input, self.age_input,
                      self.sample_number_input, self.referring_clinician_input, self.biopsy_date_input,
                      self.hospital_clinic_input, self.sample_collection_date_input, self.specimen_input,
                      self.sample_receipt_date_input, self.biopsy_performed_by_input, self.report_date_input]:
            field.textChanged.connect(self.update_preview)
            
        
        # Initialize with one embryo form (After preview browser is created)
        self.indication_input.textChanged.connect(self.update_preview)
        
        self.update_embryo_forms(1)
        
        return tab

    def schedule_preview_update(self):
        """Debounce preview updates"""
        if not hasattr(self, 'preview_timer'):
            self.preview_timer = QTimer()
            self.preview_timer.setSingleShot(True)
            self.preview_timer.setInterval(1000) # 1 second debounce
            self.preview_timer.timeout.connect(self.start_preview_generation)
            
        self.preview_timer.start()

    def update_preview(self):
        """Alias for schedule_preview_update to maintain compatibility with existing signals"""
        self.schedule_preview_update()

    def start_preview_generation(self):
        """Generate temp PDF and show in preview (Background Thread)"""
        if not hasattr(self, 'pdf_view') or isinstance(self.pdf_view, QLabel):
            return

        # Gather data from form
        p_data = {
            'patient_name': self.patient_name_input.toPlainText(),
            'spouse_name': self.spouse_name_input.toPlainText(),
            'pin': self.pin_input.toPlainText(),
            'age': self.age_input.toPlainText(),
            'sample_number': self.sample_number_input.toPlainText(),
            'referring_clinician': self.referring_clinician_input.toPlainText(),
            'biopsy_date': self.biopsy_date_input.text(),
            'hospital_clinic': self.hospital_clinic_input.toPlainText(),
            'sample_collection_date': self.sample_collection_date_input.text(),
            'specimen': self.specimen_input.toPlainText(),
            'sample_receipt_date': self.sample_receipt_date_input.text(),
            'biopsy_performed_by': self.biopsy_performed_by_input.toPlainText(),
            'report_date': self.report_date_input.text(),
            'indication': self.indication_input.toPlainText()
        }
        
        # Gather Embryo Data - Correctly iterating through all forms
        e_data = []
        
        # First check if we have data in detailed forms
        if hasattr(self, 'embryo_forms') and self.embryo_forms:
            for idx, form_dict in enumerate(self.embryo_forms):
                # We need to access widgets from the form_dict if stored, or find them
                # create_embryo_form DOES NOT currently return widget references, just layout structure?
                # Wait, I need to check create_embryo_form implementation. 
                # It currently returns {'group': group}. It needs to return input widgets to harvest data.
                
                # RECOVERY STRATEGY: 
                # Since create_embryo_form logic is inside this class, I should verify what it returns.
                # Assuming I fix create_embryo_form to return references, here is how we harvest:
                
                embryo_id = f"PS{idx+1}"
                
                # Try to get data from summary table first for high-level info
                res_sum = ""
                interp = "NA"
                mtcopy = "NA"
                
                if self.summary_table.rowCount() > idx:
                     # Item 0 is ID
                     item_id = self.summary_table.item(idx, 0)
                     if item_id: embryo_id = item_id.text()
                     
                     # Column 1 is now a dropdown widget (Result Summary)
                     widget_res = self.summary_table.cellWidget(idx, 1)
                     if widget_res: res_sum = widget_res.currentText()
                     
                     widget_interp = self.summary_table.cellWidget(idx, 2)
                     if widget_interp: interp = widget_interp.currentText()
                     
                     item_mt = self.summary_table.item(idx, 3)
                     if item_mt: mtcopy = item_mt.text()

                # Get Detailed Info (Result Desc, Autosomes, Image, Chromosomes)
                # If form_dict has references:
                # result_description and sex_chromosomes are now combo boxes
                result_desc_widget = form_dict.get('result_description')
                if result_desc_widget and hasattr(result_desc_widget, 'currentText'):
                    result_desc = result_desc_widget.currentText()
                else:
                    result_desc = ""
                    
                autosomes = form_dict.get('autosomes', QLineEdit()).text()
                
                sex_widget = form_dict.get('sex_chromosomes')
                if sex_widget and hasattr(sex_widget, 'currentText'):
                    sex = sex_widget.currentText()
                else:
                    sex = "Normal"
                    
                img_path = form_dict.get('chart_path_label', QLabel()).property("filepath")
                
                chr_statuses = {}
                mosaic_percentages = {}
                
                chr_inputs = form_dict.get('chr_inputs', {})
                for k, inputs in chr_inputs.items():
                    chr_statuses[str(k)] = inputs['status'].currentText()
                    mosaic_percentages[str(k)] = inputs['mosaic'].text()

                e_data.append({
                    'embryo_id': embryo_id,
                    'result_summary': res_sum, 
                    'interpretation': interp,
                    'result_description': result_desc,
                    'autosomes': autosomes,
                    'mtcopy': mtcopy,
                    'cnv_image_path': img_path,
                    'chromosome_statuses': chr_statuses,
                    'mosaic_percentages': mosaic_percentages
                })
        else:
            # Fallback if no forms (initial state)
             e_data = [{
                'embryo_id': 'E1',
                'result_summary': 'Euploid', 
                'interpretation': 'Euploid',
                'chromosome_statuses': {str(i): 'N' for i in range(1, 23)}
             }]

        temp_pdf = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp_preview.pdf")
        
        # Run in worker
        if hasattr(self, 'preview_worker') and self.preview_worker.isRunning():
            return # Skip if already running (debounce handles most, but safety check)
            
        show_logo = self.logo_combo.currentText() == "With Logo"
        self.preview_worker = PreviewWorker(p_data, e_data, temp_pdf, show_logo=show_logo)
        self.preview_worker.finished.connect(self.on_preview_generated)
        self.preview_worker.error.connect(lambda e: print(f"PREVIEW ERROR: {e}"))
        self.preview_worker.start()

    def on_preview_generated(self, pdf_path):
        """Load generated PDF into viewer with robust reloading"""
        if QPdfDocument and self.pdf_document and os.path.exists(pdf_path):
            try:
                # Explicitly unload current file to prevent locks
                self.pdf_document.close()
                
                # Reload fresh file
                self.pdf_document.load(pdf_path)
                self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitInView)
            except Exception as e:
                print(f"PREVIEW LOAD ERROR: {e}")
    
    def update_embryo_forms(self, count):
        """Update number of embryo forms and summary table"""
        # Update Table Rows
        self.summary_table.setRowCount(count)
        
        # Ensure connection (safe to multiple connect? No. Use unique connection or just connect once if possible)
        try:
            self.summary_table.itemChanged.disconnect(self.update_preview)
        except:
            pass
        self.summary_table.itemChanged.connect(self.update_preview)
        
        for r in range(count):
            if not self.summary_table.item(r, 0): # ID
                self.summary_table.setItem(r, 0, QTableWidgetItem(f"PS{r+1}"))
            
            if not self.summary_table.cellWidget(r, 1): # Result (Summary) - now a dropdown with colors
                result_combo = ClickOnlyComboBox()
                # Color scheme: Normal=Black, Multiple=Red, Mosaic=Blue, Inconclusive=Black, Low DNA=Black
                add_colored_items_to_combo(result_combo, [
                    ("Normal chromosome complement", "black"),
                    ("Multiple chromosomal abnormalities", "red"), 
                    ("Mosaic chromosome complement", "blue"),
                    ("Inconclusive", "black"),
                    ("Low DNA concentration", "black")
                ])
                result_combo.setEditable(True)
                result_combo.setInsertPolicy(ClickOnlyComboBox.InsertPolicy.NoInsert)
                result_combo.currentTextChanged.connect(self.update_preview)
                self.summary_table.setCellWidget(r, 1, result_combo)
            
            if not self.summary_table.cellWidget(r, 2): # Interpretation - with colors
                combo = ClickOnlyComboBox()
                # Color scheme: NA=Black, Chaotic=Red, Mosaics=Blue
                add_colored_items_to_combo(combo, [
                    ("NA", "black"),
                    ("Chaotic embryo", "red"),
                    ("Low level mosaic", "blue"),
                    ("High level mosaic", "blue"),
                    ("Complex mosaic", "blue")
                ])
                combo.setEditable(True)
                combo.setInsertPolicy(ClickOnlyComboBox.InsertPolicy.NoInsert)
                combo.currentTextChanged.connect(self.update_preview)
                self.summary_table.setCellWidget(r, 2, combo)
                
            if not self.summary_table.item(r, 3): # MTcopy
                self.summary_table.setItem(r, 3, QTableWidgetItem("NA"))
        
        # Clear existing forms
        while self.embryo_forms_layout.count():
            child = self.embryo_forms_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        
        self.embryo_forms = []
        
        # Create new forms
        for i in range(count):
            embryo_form = self.create_embryo_form(i + 1)
            self.embryo_forms.append(embryo_form)
            self.embryo_forms_layout.addWidget(embryo_form['group'])
    
    def create_embryo_form(self, embryo_num):
        """Create a single embryo data entry form for detailed view"""
        group = QGroupBox(f"Embryo {embryo_num} Details")
        form = QFormLayout()
        group.setLayout(form)
        
        # Result Description dropdown (Embryo Result Page) - All black as per spec
        result_description = ClickOnlyComboBox()
        add_colored_items_to_combo(result_description, [
            ("The embryo contains normal chromosome complement", "black"),
            ("The embryo contains abnormal chromosome complement", "black"),
            ("The embryo contains mosaic chromosome complement", "black"),
            ("Inconclusive", "black")
        ])
        result_description.setEditable(True)
        result_description.setInsertPolicy(ClickOnlyComboBox.InsertPolicy.NoInsert)
        
        autosomes = QLineEdit()
        
        # Sex Chromosomes dropdown - Normal=Black, Abnormal=Red
        sex_chromosomes = ClickOnlyComboBox()
        add_colored_items_to_combo(sex_chromosomes, [
            ("Normal", "black"),
            ("Abnormal", "red")
        ])
        sex_chromosomes.setEditable(True)
        sex_chromosomes.setInsertPolicy(ClickOnlyComboBox.InsertPolicy.NoInsert)
        
        # Connect signals
        result_description.currentTextChanged.connect(self.update_preview)
        autosomes.textChanged.connect(self.update_preview)
        sex_chromosomes.currentTextChanged.connect(self.update_preview)
        
        form.addRow("Result Description (Page 4):", result_description)
        form.addRow("Autosomes:", autosomes)
        form.addRow("Sex Chromosomes:", sex_chromosomes)
        
        # Image Upload
        img_layout = QHBoxLayout()
        img_path_label = QLabel("No image selected")
        img_path_label.setStyleSheet("color: #666; font-style: italic;")
        img_btn = QPushButton("Select Chart")
        img_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
        
        def select_image():
            path, _ = QFileDialog.getOpenFileName(self, "Select CNV Chart", "", "Images (*.png *.jpg *.jpeg)")
            if path:
                img_path_label.setText(os.path.basename(path))
                img_path_label.setProperty("filepath", path)
                self.update_preview()
        
        img_btn.clicked.connect(select_image)
        img_layout.addWidget(img_btn)
        img_layout.addWidget(img_path_label)
        img_layout.addStretch()
        form.addRow("CNV Chart:", img_layout)
    
        # Chromosome status section using Grid
        chr_group = QGroupBox("Chromosome Details")
        chr_grid = QGridLayout()
        chr_group.setLayout(chr_grid)
        form.addRow(chr_group)
    
        # Headers
        chr_grid.addWidget(QLabel("<b>Chr</b>"), 0, 0)
        chr_grid.addWidget(QLabel("<b>Status</b>"), 0, 1)
        chr_grid.addWidget(QLabel("<b>Mosaic %</b>"), 0, 2)
        chr_grid.addWidget(QLabel("   "), 0, 3) # Spacer
        chr_grid.addWidget(QLabel("<b>Chr</b>"), 0, 4)
        chr_grid.addWidget(QLabel("<b>Status</b>"), 0, 5)
        chr_grid.addWidget(QLabel("<b>Mosaic %</b>"), 0, 6)
    
        chr_inputs = {}
    
        for i in range(1, 23):
            # Determine column (Left: 1-11, Right: 12-22)
            if i <= 11:
                row = i
                col_base = 0
            else:
                row = i - 11
                col_base = 4
            
            # Label
            chr_grid.addWidget(QLabel(str(i)), row, col_base)
            
            # Status Combo
            chr_combo = ClickOnlyComboBox()
            chr_combo.addItems(["N", "G", "L", "SG", "SL", "M", "MG", "ML", "SMG", "SML"])
            chr_combo.currentTextChanged.connect(self.update_preview)
            chr_grid.addWidget(chr_combo, row, col_base + 1)
            
            # Mosaic Input
            mos_input = QLineEdit()
            mos_input.setPlaceholderText("%")
            mos_input.setMaximumWidth(60)
            mos_input.textChanged.connect(self.update_preview)
            chr_grid.addWidget(mos_input, row, col_base + 2)
            
            chr_inputs[str(i)] = {'status': chr_combo, 'mosaic': mos_input}
    
        return {
            'group': group,
            'result_description': result_description,
            'autosomes': autosomes,
            'sex_chromosomes': sex_chromosomes,
            'chr_inputs': chr_inputs,
            'chart_path_label': img_path_label
        }
    
    def create_bulk_upload_tab(self):
        """Create standalone bulk upload tab with batch editing"""
        tab = QWidget()
        main_layout = QVBoxLayout()
        tab.setLayout(main_layout)
        
        # File selection
        file_group = QGroupBox("1. Select Excel File")
        file_layout = QVBoxLayout()
        file_group.setLayout(file_layout)
        
        file_row = QHBoxLayout()
        self.bulk_file_label = QLabel("No file selected")
        file_row.addWidget(QLabel("File:"))
        file_row.addWidget(self.bulk_file_label, 1)
        
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self.browse_and_parse_bulk_file)
        file_row.addWidget(browse_btn)
        file_layout.addLayout(file_row)
        
        main_layout.addWidget(file_group)
        
        # Output folder selection
        output_group = QGroupBox("2. Select Output Folder")
        output_layout = QHBoxLayout()
        output_group.setLayout(output_layout)
        
        self.bulk_output_label = QLabel("No folder selected")
        output_layout.addWidget(QLabel("Folder:"))
        output_layout.addWidget(self.bulk_output_label, 1)
        
        output_browse_btn = QPushButton("Browse")
        output_browse_btn.clicked.connect(self.browse_bulk_output_folder)
        output_layout.addWidget(output_browse_btn)
        
        main_layout.addWidget(output_group)
        
        # Batch list and editor
        content_group = QGroupBox("3. Review and Edit Patients")
        content_layout = QHBoxLayout()
        content_group.setLayout(content_layout)
        
        # LEFT: Patient list
        left_panel = QWidget()
        left_layout = QVBoxLayout()
        left_panel.setLayout(left_layout)
        
        left_layout.addWidget(QLabel("Patients:"))
        
        # Search box for filtering patients
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("🔍"))
        self.batch_search_input = QLineEdit()
        self.batch_search_input.setPlaceholderText("Search by patient name...")
        self.batch_search_input.textChanged.connect(self.filter_batch_list)
        self.batch_search_input.setClearButtonEnabled(True)
        search_layout.addWidget(self.batch_search_input)
        left_layout.addLayout(search_layout)
        
        self.batch_list_widget = QListWidget()
        self.batch_list_widget.currentItemChanged.connect(self.on_batch_selection_changed)
        left_layout.addWidget(self.batch_list_widget)
        
        # Draft buttons
        draft_layout = QHBoxLayout()
        save_all_draft_btn = QPushButton("Save All Draft")
        save_all_draft_btn.clicked.connect(self.save_bulk_draft)
        load_draft_btn = QPushButton("Load Draft")
        load_draft_btn.clicked.connect(self.load_bulk_draft)
        draft_layout.addWidget(save_all_draft_btn)
        draft_layout.addWidget(load_draft_btn)
        left_layout.addLayout(draft_layout)
        
        # Bulk TRF Verification Section
        trf_bulk_group = QGroupBox("📋 Bulk TRF Verification")
        trf_bulk_layout = QVBoxLayout()
        trf_bulk_group.setLayout(trf_bulk_layout)
        
        # TRF upload row
        trf_upload_row = QHBoxLayout()
        self.bulk_trf_label = QLabel("No TRFs uploaded")
        self.bulk_trf_label.setStyleSheet("color: #666; font-style: italic;")
        trf_upload_row.addWidget(self.bulk_trf_label, 1)
        
        upload_bulk_trf_btn = QPushButton("📁 Upload TRFs")
        upload_bulk_trf_btn.clicked.connect(self.upload_bulk_trf)
        trf_upload_row.addWidget(upload_bulk_trf_btn)
        trf_bulk_layout.addLayout(trf_upload_row)
        
        # TRF action buttons
        trf_action_row = QHBoxLayout()
        self.bulk_trf_verify_all_btn = QPushButton("🔄 Verify All")
        self.bulk_trf_verify_all_btn.setEnabled(False)
        self.bulk_trf_verify_all_btn.clicked.connect(self.verify_all_bulk_trf)
        trf_action_row.addWidget(self.bulk_trf_verify_all_btn)
        
        trf_mgmt_btn = QPushButton("⚙️ TRF Manager")
        trf_mgmt_btn.clicked.connect(self.show_bulk_trf_verification_dialog)
        trf_action_row.addWidget(trf_mgmt_btn)
        trf_bulk_layout.addLayout(trf_action_row)
        
        # TRF status display
        self.bulk_trf_status = QTextBrowser()
        self.bulk_trf_status.setMaximumHeight(60)
        self.bulk_trf_status.setStyleSheet("background-color: #f8f9fa; border: 1px solid #ddd; border-radius: 4px; font-size: 11px;")
        self.bulk_trf_status.setHtml("<i style='color:#888;'>Upload TRF files to verify patient data</i>")
        trf_bulk_layout.addWidget(self.bulk_trf_status)
        
        left_layout.addWidget(trf_bulk_group)
        
        # Initialize bulk TRF storage
        self.bulk_trf_files = []
        self.patient_trf_mapping = {}
        self.pending_bulk_trfs = []
        
        # Logo selection for bulk
        logo_layout = QHBoxLayout()
        logo_layout.addWidget(QLabel("Branding:"))
        self.bulk_logo_combo = ClickOnlyComboBox()
        self.bulk_logo_combo.addItems(["With Logo", "Without Logo"])
        self.bulk_logo_combo.currentIndexChanged.connect(self.update_batch_preview)
        logo_layout.addWidget(self.bulk_logo_combo)
        left_layout.addLayout(logo_layout)
        
        generate_all_btn = QPushButton("Generate All Reports")
        generate_all_btn.clicked.connect(self.generate_all_batch_reports)
        left_layout.addWidget(generate_all_btn)
        
        content_layout.addWidget(left_panel)
        
        # RIGHT Panel: Layout for Editor and Preview
        right_panel = QWidget()
        right_layout = QHBoxLayout() # Horizontal to hold Splitter
        right_panel.setLayout(right_layout)
        
        # Splitter between Editor and Preview
        self.bulk_editor_splitter = QSplitter(Qt.Orientation.Horizontal)
        right_layout.addWidget(self.bulk_editor_splitter)
        
        # --- Editor Side (Left of Right Panel) ---
        editor_widget = QWidget()
        editor_layout = QVBoxLayout()
        editor_widget.setLayout(editor_layout)
        
        editor_layout.addWidget(QLabel("Patient Editor:"))
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        editor_container = QWidget()
        self.batch_editor_layout = QVBoxLayout()
        editor_container.setLayout(self.batch_editor_layout)
        scroll.setWidget(editor_container)
        editor_layout.addWidget(scroll)
        
        self.batch_editor_placeholder = QLabel("Select a patient from the list")
        self.batch_editor_placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.batch_editor_layout.addWidget(self.batch_editor_placeholder)
        
        self.bulk_editor_splitter.addWidget(editor_widget)
        
        # --- Preview Side (Right of Right Panel) ---
        preview_group = QGroupBox("Batch Report Preview")
        preview_layout = QVBoxLayout()
        preview_group.setLayout(preview_layout)
        
        if QPdfView and QPdfDocument:
            self.batch_pdf_document = QPdfDocument(self)
            self.batch_pdf_view = QPdfView(self)
            self.batch_pdf_view.setDocument(self.batch_pdf_document)
            self.batch_pdf_view.setPageMode(QPdfView.PageMode.MultiPage)
            preview_layout.addWidget(self.batch_pdf_view)
        else:
            self.batch_pdf_view = QLabel("PDF Preview not available")
            self.batch_pdf_view.setAlignment(Qt.AlignmentFlag.AlignCenter)
            preview_layout.addWidget(self.batch_pdf_view)
            
        refresh_batch_btn = QPushButton("Refresh Preview")
        refresh_batch_btn.clicked.connect(self.update_batch_preview)
        preview_layout.addWidget(refresh_batch_btn)
        
        self.bulk_editor_splitter.addWidget(preview_group)
        self.bulk_editor_splitter.setSizes([500, 500])
        
        content_layout.addWidget(right_panel)
        content_layout.setStretch(0, 1)
        content_layout.setStretch(1, 4) # Give more space to editor/preview
        
        main_layout.addWidget(content_group)
        
        return tab
    
    def create_user_guide_tab(self):
        """Create a helpful, premium user guide tab"""
        tab = QWidget()
        layout = QVBoxLayout()
        tab.setLayout(layout)
        
        guide = QTextBrowser()
        guide.setOpenExternalLinks(True)
        
        # Premium CSS for the guide
        styles = """
        <style>
            body { 
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
                line-height: 1.6; 
                color: #333; 
                background-color: #f8f9fa;
                padding: 20px;
            }
            .container { max-width: 900px; margin: auto; }
            .header { 
                background-color: #1F497D; 
                color: white; 
                padding: 30px; 
                border-radius: 10px; 
                margin-bottom: 30px;
                text-align: center;
            }
            .header h1 { margin: 0; font-size: 28px; }
            .card {
                background: white;
                border-radius: 8px;
                padding: 20px;
                margin-bottom: 20px;
                border-left: 5px solid #1F497D;
                box-shadow: 0 2px 5px rgba(0,0,0,0.05);
            }
            .card h3 { 
                color: #1F497D; 
                margin-top: 0; 
                border-bottom: 1px solid #eee;
                padding-bottom: 10px;
            }
            .feature-list { list-style-type: none; padding-left: 0; }
            .feature-list li { 
                padding: 8px 0; 
                border-bottom: 1px solid #f1f1f1;
                display: flex;
                align-items: flex-start;
            }
            .feature-list li:last-child { border-bottom: none; }
            .icon { 
                background: #e9ecef;
                color: #1F497D;
                width: 24px;
                height: 24px;
                border-radius: 50%;
                display: inline-block;
                text-align: center;
                line-height: 24px;
                margin-right: 12px;
                font-weight: bold;
                font-size: 14px;
            }
            .tip {
                background-color: #e7f3ff;
                border: 1px solid #b8daff;
                padding: 15px;
                border-radius: 5px;
                color: #004085;
                font-style: italic;
            }
        </style>
        """
        
        content = f"""
        <html>
        <head>{styles}</head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>PGT-A Report Generator Guide</h1>
                    <p>Efficient Clinical Reporting Suite</p>
                </div>

                <div class="card">
                    <h3>1. Manual Report Entry</h3>
                    <ul class="feature-list">
                        <li><span class="icon">P</span> <b>Patient Information:</b> Enter clinician and hospital details. Use DD-MM-YYYY for all dates.</li>
                        <li><span class="icon">E</span> <b>Sample Management:</b> Add multiple samples/embryos using the 'Number of Samples' selector.</li>
                        <li><span class="icon">L</span> <b>Live View:</b> The side-by-side preview updates in real-time as you type for immediate verification.</li>
                        <li><span class="icon">G</span> <b>Generation:</b> Set your output folder at the bottom and click 'Generate Report(s)'.</li>
                    </ul>
                </div>

                <div class="card">
                    <h3>2. Bulk Upload & Batch Processing</h3>
                    <ul class="feature-list">
                        <li><span class="icon">1</span> <b>Import:</b> Click 'Browse' to select your analysis run Excel file (requires standard 'Details' and 'summary' sheets).</li>
                        <li><span class="icon">2</span> <b>Review:</b> Select a patient from the list to view their automatically mapped data in the editor.</li>
                        <li><span class="icon">3</span> <b>Refine:</b> Make individual corrections in the batch editor; the live preview will follow your changes.</li>
                        <li><span class="icon">4</span> <b>Process:</b> Use 'Generate All Reports' to export the entire batch to your chosen directory.</li>
                    </ul>
                </div>

                <div class="card">
                    <h3>3. Productivity Features</h3>
                    <ul class="feature-list">
                        <li><span class="icon">⚡</span> <b>Copy Logic:</b> Use 'Copy Last Sample' to duplicate data across multiple entries instantly.</li>
                        <li><span class="icon">💾</span> <b>Drafts:</b> Save your current work as a JSON draft to reload and finish later.</li>
                        <li><span class="icon">🖼️</span> <b>Image Auto-Match:</b> Place sample charts in the same folder as the Excel file for automatic bulk matching.</li>
                        <li><span class="icon">📄</span> <b>Dual Export:</b> Generate Word (DOCX) files alongside PDFs for further manual customization.</li>
                    </ul>
                </div>

                <div class="card">
                    <h3>4. Color Coding System</h3>
                    <ul class="feature-list">
                        <li><span class="icon" style="background:#FF0000;color:white;">R</span> <b>Red:</b> Non-mosaic abnormalities - del/dup without %, CNV status L/G/SL/SG</li>
                        <li><span class="icon" style="background:#0000FF;color:white;">B</span> <b>Blue:</b> Mosaic abnormalities - any result containing % (e.g., +15(~30%), dup with ~32%)</li>
                        <li><span class="icon" style="background:#000000;color:white;">N</span> <b>Black:</b> Normal/Euploid results</li>
                    </ul>
                </div>

                <div class="card">
                    <h3>5. Excel File Format</h3>
                    <ul class="feature-list">
                        <li><span class="icon">📊</span> <b>Required Sheets:</b> 'Details' (patient info) and 'summary' (embryo results)</li>
                        <li><span class="icon">📝</span> <b>Details Columns:</b> Patient Name, Sample ID, Center name, Date of Biopsy, Date Sample Received, EMBRYOLOGIST NAME</li>
                        <li><span class="icon">🧬</span> <b>Summary Columns:</b> Sample name, QC, Conclusion, Result, MTcopy, AUTOSOMES, SEX</li>
                        <li><span class="icon">⚠️</span> <b>Note:</b> Referring Clinician and Sample Number fields are NOT auto-filled - enter manually</li>
                    </ul>
                </div>

                <div class="card">
                    <h3>6. System Requirements</h3>
                    <ul class="feature-list">
                        <li><span class="icon">🐍</span> <b>Python:</b> Version 3.8 or higher</li>
                        <li><span class="icon">📦</span> <b>Required Packages:</b> PyQt6, ReportLab, python-docx, pandas, openpyxl, Pillow</li>
                        <li><span class="icon">💻</span> <b>OS:</b> Windows, Linux, or macOS</li>
                        <li><span class="icon">📁</span> <b>Storage:</b> 100MB for application + space for generated reports</li>
                    </ul>
                    <p style="margin-top:15px;"><b>Installation:</b> <code style="background:#f1f1f1;padding:5px 10px;border-radius:3px;">pip install -r requirements.txt</code></p>
                </div>

                <div class="card">
                    <h3>7. Quick Start</h3>
                    <ul class="feature-list">
                        <li><span class="icon">1</span> Install requirements: <code>pip install -r requirements.txt</code></li>
                        <li><span class="icon">2</span> Run application: <code>python pgta_report_generator.py</code></li>
                        <li><span class="icon">3</span> For Manual Entry: Fill patient details → Add embryo data → Generate</li>
                        <li><span class="icon">4</span> For Bulk Upload: Browse Excel → Search/Select patient → Review → Generate All</li>
                    </ul>
                </div>

                <div class="tip">
                    <b>Pro Tip:</b> Use the search box in Bulk Upload to quickly find patients by name instead of scrolling through the list.
                </div>
                
                <p style="text-align: center; color: #888; margin-top: 40px; font-size: 12px;">
                    PGT-A Clinical Reporting Suite v2.0 | © 2026 Andrology Center
                </p>
            </div>
        </body>
        </html>
        """
        guide.setHtml(content)
        layout.addWidget(guide)
        return tab


    def copy_last_embryo(self):
        """Duplicate the data from the last embryo to a new one"""
        current_data = self.get_manual_data_dict()
        embryos = current_data.get('embryos', [])
        if not embryos:
            return
            
        last_embryo = embryos[-1]
        
        # Increment spinner
        current_count = self.embryo_count_spin.value()
        if current_count >= 20:
            QMessageBox.warning(self, "Limit Reached", "Maximum of 20 embryos allowed.")
            return
            
        # Prevent the signal from triggering multiple updates while we populate
        self.embryo_count_spin.blockSignals(True)
        self.embryo_count_spin.setValue(current_count + 1)
        self.update_embryo_forms(current_count + 1) # Manual call
        self.embryo_count_spin.blockSignals(False)
        
        # Now fill the new embryo (which is the last one now) at index 'current_count'
        new_index = current_count 
        
        if new_index < len(self.embryo_forms):
            form = self.embryo_forms[new_index]
            # result_description and sex_chromosomes are now combo boxes
            form['result_description'].setCurrentText(last_embryo.get('result_description', ''))
            form['autosomes'].setText(last_embryo.get('autosomes', ''))
            form['sex_chromosomes'].setCurrentText(last_embryo.get('sex_chromosomes', 'Normal'))
            
            # Chromosomes
            chr_statuses = last_embryo.get('chromosome_statuses', {})
            mosaic_data = last_embryo.get('mosaic_percentages', {})
            for k in range(1, 23):
                s_i = str(k)
                if s_i in form['chr_inputs']:
                    inputs = form['chr_inputs'][s_i]
                    inputs['status'].setCurrentText(chr_statuses.get(s_i, 'N'))
                    inputs['mosaic'].setText(mosaic_data.get(s_i, ''))
            
            # Summary Table
            # Note: summary_table ID is handled by update_embryo_forms, but we copy the rest
            # Column 1 (Result Summary) is now a combo box
            val_summary = last_embryo.get('result_summary', '')
            w_res_sum = self.summary_table.cellWidget(new_index, 1)
            if w_res_sum and isinstance(w_res_sum, QComboBox):
                w_res_sum.setCurrentText(val_summary)
            w_interp = self.summary_table.cellWidget(new_index, 2)
            if w_interp and isinstance(w_interp, QComboBox):
                w_interp.setCurrentText(last_embryo.get('interpretation', ''))
            self.summary_table.setItem(new_index, 3, QTableWidgetItem(last_embryo.get('mtcopy', 'NA')))
            
        self.update_preview()
        self.statusBar().showMessage(f"Embryo data copied to {new_index+1}")

    def generate_preview_html(self):
        """Generate high-fidelity HTML for preview"""
        # Collect data
        p_data = self.get_manual_data_dict()
        p_info = p_data.get('patient_info', {})
        embryos = p_data.get('embryos', [])
        
        # Colors matching improved original template
        PATIENT_INFO_BG = "#F1F1F7"
        SUMMARY_HEADER_BG = "#F9BE8F"
        GREY_SECTION_BG = "#F1F1F7" # Matches source extracted color for disclaimers
        BORDER_COLOR = "#D1D1D1"
        BLUE_TITLE = "#1F497D"
        TEXT_RED = "#FF0000"
        TEXT_BLUE = "#0000FF"
        
        # Logo Path
        # Logo Path (Updated to assets/pgta/)
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "pgta", "image_page1_0.png")
        logo_html = f'<img src="file://{logo_path}" height="50">' if os.path.exists(logo_path) else '<div style="font-size: 24px; font-weight: bold; color: {BLUE_TITLE};">CHROMINST</div>'
        
        html = f"""
        <html>
        <head>
            <style>
                body {{ 
                    font-family: Arial, sans-serif; 
                    padding: 40px; 
                    color: #333; 
                    background-color: #FAFAFA;
                    line-height: 1.4;
                }}
                .report-page {{
                    background: white;
                    box-shadow: 0 0 15px rgba(0,0,0,0.1);
                    max-width: 800px;
                    margin: 0 auto 40px;
                    padding: 50px;
                    border: 1px solid #eee;
                }}
                .header {{
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    margin-bottom: 20px;
                }}
                .title {{
                    font-size: 18px;
                    font-weight: bold;
                    color: {BLUE_TITLE};
                    text-align: center;
                    padding-bottom: 8px;
                    margin-bottom: 25px;
                    text-transform: uppercase;
                }}
                
                table {{ 
                    width: 100%; 
                    border-collapse: collapse; 
                    margin-bottom: 15px; 
                    font-size: 12px; 
                }}
                th, td {{ 
                    border: 1px solid {BORDER_COLOR}; 
                    padding: 6px 10px; 
                    text-align: left; 
                }}
                
                .patient-table td {{ 
                    background-color: {PATIENT_INFO_BG}; 
                }}
                
                .summary-table th {{ 
                    background-color: {SUMMARY_HEADER_BG}; 
                    color: black; 
                    font-weight: bold;
                    text-align: center;
                }}
                .summary-table td {{ 
                    text-align: center;
                    background-color: {PATIENT_INFO_BG};
                }}
                
                .section-header {{
                    font-weight: bold;
                    margin-top: 15px;
                    margin-bottom: 8px;
                    font-size: 13px;
                }}
                
                .img-container {{ 
                    text-align: center; 
                    margin: 15px 0; 
                    border: 1px solid #f0f0f0; 
                    padding: 8px;
                    background-color: white;
                }}
                
                .chr-table {{
                    font-size: 10px;
                }}
                .chr-table th {{
                    background-color: {SUMMARY_HEADER_BG};
                    text-align: center;
                }}
                .chr-table td {{
                    text-align: center;
                    background-color: {PATIENT_INFO_BG};
                }}
                .text-red {{ color: {TEXT_RED}; font-weight: bold; }}
                .text-blue {{ color: {TEXT_BLUE}; font-weight: bold; }}
                
                .disclaimer {{
                    background-color: {GREY_SECTION_BG};
                    color: black;
                    font-weight: bold;
                    text-align: center;
                    padding: 8px;
                    margin: 15px 0;
                    font-size: 11px;
                }}
                
                .footer-signatures {{
                    margin-top: 40px;
                    display: flex;
                    justify-content: space-between;
                }}
                .sig-box {{
                    text-align: center;
                    flex: 1;
                    font-size: 11px;
                }}
                .sig-line {{
                    border-top: 1px solid #333;
                    margin: 8px 15px;
                    padding-top: 3px;
                }}
            </style>
        </head>
        <body>
            <div class="report-page">
                <div class="header">
                    <div style="height: 50px;">{logo_html}</div>
                    <div style="text-align: right; color: #666; font-size: 10px;">PREIMPLANTATION GENETIC TESTING<br>FOR ANEUPLOIDIES (PGT-A)</div>
                </div>
                
                <div class="title">Preimplantation Genetic Testing for Aneuploidies (PGT-A)</div>
                
                <table class="patient-table">
                    <tr>
                        <td><b>Patient Name :</b> {p_info.get('patient_name', '')} {p_info.get('spouse_name', '')}</td>
                        <td><b>PIN :</b> {p_info.get('pin', '')}</td>
                    </tr>
                    <tr>
                        <td><b>Age :</b> {p_info.get('age', '')}</td>
                        <td><b>Sample Number :</b> {p_info.get('sample_number', '')}</td>
                    </tr>
                </table>
                <div class="section-header">Results summary</div>
                <table class="summary-table">
                    <tr>
                        <th>S.No</th>
                        <th>Sample</th>
                        <th>Result</th>
                        <th>MTcopy</th>
                        <th>Interpretation</th>
                    </tr>
        """
        
        for i, embryo in enumerate(embryos):
            res_sum = embryo.get('result_summary', '')
            interp = embryo.get('interpretation', '')
            color_class = self._get_preview_color_class(res_sum, interp)
            
            # MTcopy: NA for non-euploid
            mtcopy = embryo.get('mtcopy', 'NA')
            if interp.upper() != "EUPLOID":
                mtcopy = "NA"
                
            html += f"""
                <tr>
                    <td><b>{i+1}</b></td>
                    <td><b>{embryo.get('embryo_id', '')}</b></td>
                    <td class="{color_class}"><b>{res_sum}</b></td>
                    <td><b>{mtcopy}</b></td>
                    <td class="{color_class}"><b>{interp}</b></td>
                </tr>
            """
            
        html += """
                </table>
            </div>
            
            <div class="report-page">
                <div class="title">Preimplantation Genetic Testing for Aneuploidies (PGT-A)</div>
                <div class="section-header">Methodology</div>
                <div style="font-size: 11px;">Chromosomal aneuploidy analysis was performed using ChromInst® PGT-A from Yikon Genomics (Suzhou) Co., Ltd - China. The Yikon - ChromInst® PGT-A kit with the Genemind - SURFSeq 5000* High-throughput Sequencing Platform allows detection of aneuploidies in all 23 sets of Chromosomes. Probes are not covering the p arm of acrocentric chromosomes as they are rich in repeat regions and RNA markers and devoid of genes. Changes in this region will not be detected. However, these regions have less clinical significance due to the absence of genes. Chromosomal aneuploidy can be detected by copy number variations (CNVs), which represent a class of variation in which segments of the genome have been duplicated (gains) or deleted (losses). Large, genomic copy number imbalances can range from sub-chromosomal regions to entire chromosomes. Inherited and de-novo CNVs (up to 10 Mb) have been associated with many disease conditions. This assay was performed on DNA extracted from embryo biopsy samples.</div>
                
                <div class="section-header">Conditions for reporting mosaicism</div>
                <div style="font-size: 11px;">Mosaicism arises in the embryo due to mitotic errors which lead to the production of karyotypically distinct cell lineages within a single embryo [1]. NGS has the sensitivity to detect mosaicism when 30% or the above cells are abnormal [2]. Mosaicism is reported in our laboratory as follows [3].</div>
                <ul style="font-size: 11px; margin-top: 5px;">
                    <li>Embryos with less than 30% mosaicism are considered as euploid.</li>
                    <li>Embryos with 30% to 50% mosaicism will be reported as low level mosaic, 51% to 80% mosaicism will be reported as high level mosaic.</li>
                    <li>When three chromosomes or more than three chromosomes showing mosaic change, it will be denoted as complex mosaic.</li>
                    <li>If greater than 80% mosaicism detected in an embryo it will be considered aneuploid.</li>
                </ul>
                <div style="font-size: 11px; margin-top: 10px;">Clinical significance of transferring mosaic embryos is still under evaluation. Based on Preimplantation Genetic Diagnosis International Society (PGDIS) Position Statement – 2019 transfer of these embryos should be considered only after appropriate counselling of the patient and alternatives have been discussed. Invasive prenatal testing with karyotyping in the amniotic fluid needs to be advised in such cases [4]. As shown in published literature evidence, such transfers can result in normal pregnancy or miscarriage or an offspring with chromosomal mosaicism [5,6,7].</div>
                
                <div class="section-header">Limitations</div>
                <ul style="font-size: 11px; margin-top: 5px;">
                    <li>This technique cannot detect point mutations, balanced translocations, inversions, triploidy, uniparental disomy and epigenetic modifications.</li>
                    <li>Probes used do not cover the p arm of acrocentric chromosomes as they are rich in repeat regions and RNA markers and devoid of genes. Changes in this region will not be detected. However, these regions have less clinical significance due to the absence of genes.</li>
                    <li>Deletions and duplications with the size of < 10 Mb cannot be detected.</li>
                    <li>Risk of misinterpretation of the actual embryo karyotype due to the presence of chromosomal mosaicism, either at cleavage-stage or at blastocyst stage may exist.</li>
                    <li>This technique cannot detect variants of polyploidy and haploidy</li>
                    <li>NGS without genotyping cannot identify the nature (meiotic or mitotic) nor the parental origin of aneuploidies</li>
                    <li>Due to the intrinsic nature of chromosomal mosaicism, the chromosomal make-up achieved from a biopsy only may represent a picture of a small part of the embryo and may not necessarily reflect the chromosomal content of the entire embryo. Also, the mosaicism level inferred from a multi-cell TE biopsy might not unequivocally represent the exact chromosomal mosaicism percentage of the TE cells or the inner cell mass constitution.</li>
                </ul>
                
                <div class="section-header">References</div>
                <ol style="font-size: 10px; margin-top: 5px; color: #555;">
                    <li>McCoy, Rajiv C. "Mosaicism in Preimplantation human embryos: when chromosomal abnormalities are the norm." Trends in genetics 33.7 (2017): 448-463.</li>
                    <li>ESHRE PGT-SR/PGT-A Working Group, et al. "ESHRE PGT Consortium good practice recommendations for the detection of structural and numerical chromosomal aberrations." Human reproduction open 2020.3 (2020): hoaa017.</li>
                    <li>ESHRE Working Group on Chromosomal Mosaicism, et al. "ESHRE survey results and good practice recommendations on managing chromosomal mosaicism." Hum Reprod Open. 2022 Nov 7;2022(4):hoac044.</li>
                    <li>Cram, D. S., et al. "PGDIS position statement on the transfer of mosaic embryos 2019." Reproductive biomedicine online 39 (2019): e1-e4.</li>
                    <li>Victor, Andrea R., et al. "One hundred mosaic embryos transferred prospectively in a single clinic: exploring when and why they result in healthy pregnancies." Fertility and sterility 111.2 (2019): 280-293.</li>
                    <li>Lin, Pin-Yao, et al. "Clinical outcomes of single mosaic embryo transfer: high-level or low-level mosaic embryo, does it matter?" Journal of clinical medicine 9.6 (2020): 1695.</li>
                    <li>Kahraman, Semra, et al. "The birth of a baby with mosaicism resulting from a known mosaic embryo transfer: a case report." Human Reproduction 35.3 (2020): 727-733.</li>
                </ol>
            </div>
        """

        
    def get_manual_data_dict(self):
        """Collect current manual entry data into a dictionary with 'nan' sanitation"""
        def clean(val):
            if val is None: return ""
            s = str(val).strip(' \t\r\f\v') # Preserves newlines (\n)
            return "" if s.lower() == "nan" else s

        patient_info = {
            'patient_name': clean(self.patient_name_input.toPlainText()),
            'spouse_name': clean(self.spouse_name_input.toPlainText()),
            'pin': clean(self.pin_input.toPlainText()),
            'age': clean(self.age_input.toPlainText()),
            'sample_number': clean(self.sample_number_input.toPlainText()),
            'referring_clinician': clean(self.referring_clinician_input.toPlainText()),
            'biopsy_date': clean(self.biopsy_date_input.text()),
            'hospital_clinic': clean(self.hospital_clinic_input.toPlainText()),
            'sample_collection_date': clean(self.sample_collection_date_input.text()),
            'specimen': clean(self.specimen_input.toPlainText()),
            'sample_receipt_date': clean(self.sample_receipt_date_input.text()),
            'biopsy_performed_by': clean(self.biopsy_performed_by_input.toPlainText()),
            'report_date': clean(self.report_date_input.text()),
            'indication': clean(self.indication_input.toPlainText())
        }
        
        embryos = []
        count = self.summary_table.rowCount()
        
        for i in range(count):
            # 1. Summary Table Data
            item_id = self.summary_table.item(i, 0)
            t_id = item_id.text() if item_id else ""
            
            # Column 1 (Result Summary) is now a combo box
            w_sum = self.summary_table.cellWidget(i, 1)
            t_sum = w_sum.currentText() if w_sum else ""
            
            w_interp = self.summary_table.cellWidget(i, 2)
            t_interp = w_interp.currentText() if w_interp else ""
            
            item_mt = self.summary_table.item(i, 3)
            t_mt = item_mt.text() if item_mt else ""
            
            # 2. Detail Data
            result_desc = ""
            autosomes = ""
            sex = ""
            cnv_image_path = None
            chr_statuses = {}
            mosaic_percentages = {}
            
            if i < len(self.embryo_forms):
                form = self.embryo_forms[i]
                # result_description and sex_chromosomes are now combo boxes
                result_desc = form['result_description'].currentText() if hasattr(form['result_description'], 'currentText') else form['result_description'].text()
                autosomes = form['autosomes'].text()
                sex = form['sex_chromosomes'].currentText() if hasattr(form['sex_chromosomes'], 'currentText') else form['sex_chromosomes'].text()
                
                # Image
                if 'chart_path_label' in form:
                     path = form['chart_path_label'].property("filepath")
                     if path and os.path.exists(path):
                         cnv_image_path = path

                # Chromosomes
                for k in range(1, 23):
                    if str(k) in form['chr_inputs']:
                        inputs = form['chr_inputs'][str(k)]
                        chr_statuses[str(k)] = inputs['status'].currentText()
                        val = inputs['mosaic'].text()
                        if val:
                            mosaic_percentages[str(k)] = val
            
            # Fallback Image Lookup
            if not cnv_image_path and t_id:
                target_id = t_id.strip().upper()
                for stored_id, paths in self.uploaded_images.items():
                    if stored_id.strip().upper() == target_id and paths:
                         cnv_image_path = paths[0]
                         break
            
            embryo = {
                'embryo_id': t_id,
                'result_summary': t_sum,
                'interpretation': t_interp,
                'mtcopy': t_mt,
                'result_description': result_desc,
                'autosomes': autosomes,
                'sex_chromosomes': sex,
                'cnv_image_path': cnv_image_path,
                'chromosome_statuses': chr_statuses,
                'mosaic_percentages': mosaic_percentages
            }
            embryos.append(embryo)
            
        return {'patient_info': patient_info, 'embryos': embryos}

    def save_manual_data(self):
        """Save manually entered data"""
        self.current_patient_data = self.get_manual_data_dict()
        self.update_data_summary()
        
        QMessageBox.information(self, "Success", "Data saved successfully! You can now generate reports.")
        self.statusBar().showMessage("Manual data saved")
    
    
    def save_draft(self):
        """Save current form data to JSON"""
        data = self.get_manual_data_dict()
        path, _ = QFileDialog.getSaveFileName(self, "Save Draft", "pgta_draft.json", "JSON Files (*.json)")
        
        if path:
            try:
                # Add metadata
                data['version'] = "1.0"
                data['timestamp'] = datetime.now().isoformat()
                
                with open(path, 'w') as f:
                    json.dump(data, f, indent=4)
                self.statusBar().showMessage(f"Draft saved to {path}")
                QMessageBox.information(self, "Draft Saved", f"Draft saved successfully to:\n{path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def load_draft(self):
        """Load form data from JSON"""
        path, _ = QFileDialog.getOpenFileName(self, "Load Draft", "", "JSON Files (*.json)")
        if not path:
            return
            
        try:
            with open(path, 'r') as f:
                data = json.load(f)
            
            self.populate_manual_form(data)
            self.update_preview()
            self.statusBar().showMessage(f"Draft loaded from {path}")
            QMessageBox.information(self, "Draft Loaded", "Draft loaded successfully.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load draft: {str(e)}")

    def populate_manual_form(self, data):
        """Populate form with data from dictionary"""
        p_info = data.get('patient_info', {})
        
        # Patient fields
        self.patient_name_input.setText(p_info.get('patient_name', ''))
        self.spouse_name_input.setText(p_info.get('spouse_name', ''))
        self.pin_input.setText(p_info.get('pin', ''))
        self.age_input.setText(p_info.get('age', ''))
        self.sample_number_input.setText(p_info.get('sample_number', ''))
        self.referring_clinician_input.setText(p_info.get('referring_clinician', ''))
        self.biopsy_date_input.setText(p_info.get('biopsy_date', ''))
        self.hospital_clinic_input.setText(p_info.get('hospital_clinic', ''))
        self.sample_collection_date_input.setText(p_info.get('sample_collection_date', ''))
        self.specimen_input.setText(p_info.get('specimen', ''))
        self.sample_receipt_date_input.setText(p_info.get('sample_receipt_date', ''))
        self.biopsy_performed_by_input.setText(p_info.get('biopsy_performed_by', ''))
        self.report_date_input.setText(p_info.get('report_date', ''))
        self.indication_input.setText(p_info.get('indication', ''))
        
        # Embryos
        embryos = data.get('embryos', [])
        count = len(embryos)
        
        # Update spinbox (triggers form/table creation)
        self.embryo_count_spin.setValue(count)
        if len(self.embryo_forms) != count:
            # Note: setValue(count) above already triggers update_embryo_forms via signal
            # Only call manually if the count didn't change but we need a refresh
            pass
            
        # 1. Fill Summary Table
        for i, embryo in enumerate(embryos):
            if i >= self.summary_table.rowCount():
                break
                
            # ID
            self.summary_table.setItem(i, 0, QTableWidgetItem(embryo.get('embryo_id', f'PS{i+1}')))
            # Summary - now a combo box
            w_sum = self.summary_table.cellWidget(i, 1)
            if w_sum and isinstance(w_sum, QComboBox):
                 w_sum.setCurrentText(embryo.get('result_summary', ''))
            # Interp
            w_interp = self.summary_table.cellWidget(i, 2)
            if w_interp and isinstance(w_interp, QComboBox):
                 w_interp.setCurrentText(embryo.get('interpretation', ''))
            # MTcopy
            self.summary_table.setItem(i, 3, QTableWidgetItem(embryo.get('mtcopy', 'NA')))
            
        # 2. Fill Detail Forms
        for idx, embryo in enumerate(embryos):
            if idx >= len(self.embryo_forms):
                break
                
            form = self.embryo_forms[idx]
            
            # result_description and sex_chromosomes are now combo boxes
            if hasattr(form['result_description'], 'setCurrentText'):
                form['result_description'].setCurrentText(embryo.get('result_description', ''))
            else:
                form['result_description'].setText(embryo.get('result_description', ''))
            form['autosomes'].setText(embryo.get('autosomes', ''))
            if hasattr(form['sex_chromosomes'], 'setCurrentText'):
                form['sex_chromosomes'].setCurrentText(embryo.get('sex_chromosomes', 'Normal'))
            else:
                form['sex_chromosomes'].setText(embryo.get('sex_chromosomes', ''))
            
            # Image
            path = embryo.get('cnv_image_path')
            if path and os.path.exists(path) and 'chart_path_label' in form:
                 form['chart_path_label'].setText(os.path.basename(path))
                 form['chart_path_label'].setProperty("filepath", path)
            
            # Chromosomes
            chr_statuses = embryo.get('chromosome_statuses', {})
            mosaic_data = embryo.get('mosaic_percentages', {})
            
            for i in range(1, 23):
                s_i = str(i)
                if s_i in form['chr_inputs']:
                    inputs = form['chr_inputs'][s_i]
                    inputs['status'].setCurrentText(chr_statuses.get(s_i, 'N'))
                    inputs['mosaic'].setText(mosaic_data.get(s_i, ''))
    
    # ==================== TRF Verification Methods ====================
    
    def extract_text_from_trf(self, file_path):
        """Extract text from TRF file (image or PDF)"""
        text = ""
        file_ext = os.path.splitext(file_path)[1].lower()
        
        try:
            if file_ext in ['.pdf']:
                # Extract text from PDF
                if PDFPLUMBER_AVAILABLE:
                    with pdfplumber.open(file_path) as pdf:
                        for page in pdf.pages:
                            page_text = page.extract_text()
                            if page_text:
                                text += page_text + "\n"
                else:
                    return None, "PDF extraction requires pdfplumber. Please install it."
                    
            elif file_ext in ['.png', '.jpg', '.jpeg', '.tiff', '.bmp']:
                # Extract text from image using OCR
                if TESSERACT_AVAILABLE:
                    img = Image.open(file_path)
                    text = pytesseract.image_to_string(img)
                else:
                    return None, "Image OCR requires pytesseract. Please install: pip install pytesseract\nAlso install Tesseract OCR: sudo apt install tesseract-ocr"
            else:
                return None, f"Unsupported file format: {file_ext}"
                
        except Exception as e:
            return None, f"Error extracting text: {str(e)}"
        
        return text, None
    
    def fuzzy_match(self, str1, str2, threshold=0.7):
        """Check if two strings are similar using fuzzy matching"""
        if not str1 or not str2:
            return False, 0.0
        str1 = str(str1).lower().strip()
        str2 = str(str2).lower().strip()
        ratio = SequenceMatcher(None, str1, str2).ratio()
        return ratio >= threshold, ratio
    
    def find_in_text(self, text, search_term, context_words=3):
        """Find a term in text and return if found with context"""
        if not text or not search_term:
            return False, ""
        
        text_lower = text.lower()
        search_lower = str(search_term).lower().strip()
        
        # Direct match
        if search_lower in text_lower:
            return True, "exact match"
        
        # Word-by-word fuzzy match for names
        search_words = search_lower.split()
        text_words = text_lower.split()
        
        matched_words = 0
        for sw in search_words:
            for tw in text_words:
                is_match, ratio = self.fuzzy_match(sw, tw, 0.8)
                if is_match:
                    matched_words += 1
                    break
        
        if len(search_words) > 0 and matched_words / len(search_words) >= 0.6:
            return True, f"fuzzy match ({matched_words}/{len(search_words)} words)"
        
        return False, "not found"
    
    def extract_field_from_trf(self, trf_text, field_key):
        """Extract specific field value from TRF text using patterns"""
        if not trf_text:
            return None
        
        lines = trf_text.split('\n')
        
        # Enhanced field-specific patterns
        patterns = {
            'patient_name': [
                r'patient\s*name[:\s]*([A-Za-z\s\.]+?)(?:\n|$|wife|husband|w/o|s/o|d/o)',
                r'name\s*of\s*patient[:\s]*([A-Za-z\s\.]+?)(?:\n|$)',
                r'(?:mrs?\.?|miss|dr\.?)\s+([A-Za-z\s\.]+?)(?:\n|$|wife)',
                r'name[:\s]*([A-Za-z][A-Za-z\s\.]{2,30})(?:\n|$)',
            ],
            'hospital_clinic': [
                r'(?:hospital|clinic|center|centre)[:\s]*([^\n]+)',
                r'(?:ivf|fertility)\s*(?:center|centre|clinic)[:\s]*([^\n]*)',
                r'from[:\s]*([^\n]+(?:hospital|clinic|center|centre|ivf)[^\n]*)',
            ],
            'pin': [
                r'(?:pin|patient\s*id|sample\s*id|id\s*no)[:\s\.]*([A-Z0-9]+)',
                r'(?:AND|PIN)[:\s]*([A-Z0-9]+)',
                r'sample[:\s]*([A-Z0-9]+)',
            ],
            'biopsy_date': [
                r'(?:date\s*of\s*)?biopsy[:\s]*(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})',
                r'biopsy\s*date[:\s]*(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})',
                r'collection\s*date[:\s]*(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})',
            ],
            'sample_receipt_date': [
                r'(?:sample\s*)?receipt\s*date[:\s]*(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})',
                r'received[:\s]*(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})',
                r'date\s*received[:\s]*(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})',
            ],
            'referring_clinician': [
                r'(?:referring\s*)?(?:clinician|doctor|dr)[:\s]*([A-Za-z\s\.]+?)(?:\n|$)',
                r'referred\s*by[:\s]*([A-Za-z\s\.]+?)(?:\n|$)',
                r'(?:dr\.?|doctor)\s+([A-Za-z\s\.]+?)(?:\n|$)',
            ],
            'embryologist': [
                r'embryologist[:\s]*([A-Za-z\s\.]+?)(?:\n|$)',
                r'biologist[:\s]*([A-Za-z\s\.]+?)(?:\n|$)',
            ],
        }
        
        if field_key in patterns:
            for pattern in patterns[field_key]:
                for line in lines:
                    match = re.search(pattern, line, re.IGNORECASE)
                    if match:
                        extracted = match.group(1).strip()
                        # Clean up extracted value
                        extracted = re.sub(r'\s+', ' ', extracted)
                        if len(extracted) > 2:
                            return extracted
                # Also try on full text
                match = re.search(pattern, trf_text, re.IGNORECASE)
                if match:
                    extracted = match.group(1).strip()
                    extracted = re.sub(r'\s+', ' ', extracted)
                    if len(extracted) > 2:
                        return extracted
        
        return None
    
    def verify_patient_data_enhanced(self, trf_text, patient_data):
        """Enhanced verification with extracted values and comparison"""
        results = []
        all_correct = True
        has_suggestions = False
        
        fields_to_check = [
            ('patient_name', 'Patient Name', 'patient_name_input'),
            ('hospital_clinic', 'Hospital/Clinic', 'hospital_clinic_input'),
            ('pin', 'PIN/Sample ID', 'pin_input'),
            ('biopsy_date', 'Biopsy Date', 'biopsy_date_input'),
            ('sample_receipt_date', 'Receipt Date', 'sample_receipt_date_input'),
            ('referring_clinician', 'Referring Clinician', 'referring_clinician_input'),
        ]
        
        for field_key, field_label, widget_name in fields_to_check:
            entered_value = patient_data.get(field_key, '')
            trf_value = self.extract_field_from_trf(trf_text, field_key)
            
            # Skip if no data entered and nothing found in TRF
            if (not entered_value or entered_value.lower() in ['nan', 'none', '', 'w/o']) and not trf_value:
                results.append({
                    'field': field_label,
                    'field_key': field_key,
                    'widget': widget_name,
                    'status': 'skip',
                    'entered': entered_value or '(empty)',
                    'trf_value': '(not found)',
                    'message': 'No data',
                    'icon': '⚪',
                    'can_apply': False
                })
                continue
            
            # If entered value exists, check if it matches TRF
            if entered_value and entered_value.lower() not in ['nan', 'none', '', 'w/o']:
                found, match_type = self.find_in_text(trf_text, entered_value)
                
                if found:
                    results.append({
                        'field': field_label,
                        'field_key': field_key,
                        'widget': widget_name,
                        'status': 'ok',
                        'entered': entered_value,
                        'trf_value': trf_value or entered_value,
                        'message': f'✓ Match ({match_type})',
                        'icon': '✅',
                        'can_apply': False
                    })
                else:
                    all_correct = False
                    if trf_value:
                        has_suggestions = True
                        results.append({
                            'field': field_label,
                            'field_key': field_key,
                            'widget': widget_name,
                            'status': 'mismatch',
                            'entered': entered_value,
                            'trf_value': trf_value,
                            'message': '✗ Mismatch',
                            'icon': '❌',
                            'can_apply': True
                        })
                    else:
                        results.append({
                            'field': field_label,
                            'field_key': field_key,
                            'widget': widget_name,
                            'status': 'warning',
                            'entered': entered_value,
                            'trf_value': '(not found in TRF)',
                            'message': '⚠ Not verified',
                            'icon': '⚠️',
                            'can_apply': False
                        })
            else:
                # No entered value but TRF has value - suggest auto-fill
                if trf_value:
                    has_suggestions = True
                    results.append({
                        'field': field_label,
                        'field_key': field_key,
                        'widget': widget_name,
                        'status': 'suggestion',
                        'entered': '(empty)',
                        'trf_value': trf_value,
                        'message': '💡 Found in TRF',
                        'icon': '💡',
                        'can_apply': True
                    })
        
        return results, all_correct, has_suggestions
    
    def show_trf_comparison_dialog(self, results, is_batch=False):
        """Show a dialog with side-by-side comparison and apply options"""
        dialog = QDialog(self)
        dialog.setWindowTitle("TRF Verification Results")
        dialog.setMinimumWidth(700)
        dialog.setMinimumHeight(450)
        
        layout = QVBoxLayout()
        dialog.setLayout(layout)
        
        # Header
        all_match = all(r['status'] in ['ok', 'skip'] for r in results)
        if all_match:
            header = QLabel("✅ All fields verified successfully!")
            header.setStyleSheet("font-size: 16px; font-weight: bold; color: #28a745; padding: 10px;")
        else:
            header = QLabel("⚠️ Some fields need attention - Review differences below")
            header.setStyleSheet("font-size: 16px; font-weight: bold; color: #dc3545; padding: 10px;")
        layout.addWidget(header)
        
        # Comparison Table
        table = QTableWidget()
        table.setColumnCount(5)
        table.setHorizontalHeaderLabels(["Field", "Current Value", "TRF Value", "Status", "Action"])
        table.horizontalHeader().setStretchLastSection(True)
        table.setColumnWidth(0, 120)
        table.setColumnWidth(1, 180)
        table.setColumnWidth(2, 180)
        table.setColumnWidth(3, 100)
        table.setColumnWidth(4, 80)
        
        table.setRowCount(len(results))
        self.trf_apply_buttons = []
        
        for i, r in enumerate(results):
            # Field name
            table.setItem(i, 0, QTableWidgetItem(r['field']))
            
            # Current value
            current_item = QTableWidgetItem(r['entered'])
            if r['status'] == 'mismatch':
                current_item.setBackground(QColor('#ffcccc'))
            table.setItem(i, 1, current_item)
            
            # TRF value
            trf_item = QTableWidgetItem(r['trf_value'])
            if r['status'] in ['mismatch', 'suggestion']:
                trf_item.setBackground(QColor('#ccffcc'))
            table.setItem(i, 2, trf_item)
            
            # Status
            status_item = QTableWidgetItem(r['message'])
            if r['status'] == 'ok':
                status_item.setForeground(QColor('#28a745'))
            elif r['status'] in ['mismatch', 'warning']:
                status_item.setForeground(QColor('#dc3545'))
            elif r['status'] == 'suggestion':
                status_item.setForeground(QColor('#007bff'))
            table.setItem(i, 3, status_item)
            
            # Apply button
            if r['can_apply'] and r['trf_value'] and r['trf_value'] != '(not found)':
                apply_btn = QPushButton("Apply")
                apply_btn.setStyleSheet("background-color: #28a745; color: white; padding: 3px 8px;")
                apply_btn.clicked.connect(lambda checked, idx=i, res=r, batch=is_batch: self.apply_single_trf_value(res, batch, dialog, table, idx))
                table.setCellWidget(i, 4, apply_btn)
                self.trf_apply_buttons.append((apply_btn, r))
            else:
                table.setItem(i, 4, QTableWidgetItem(""))
        
        layout.addWidget(table)
        
        # Store for apply all
        self.trf_results = results
        self.trf_is_batch = is_batch
        self.trf_dialog = dialog
        self.trf_table = table
        
        # Buttons
        btn_layout = QHBoxLayout()
        
        # Apply All Suggestions button
        suggestions = [r for r in results if r['can_apply']]
        if suggestions:
            apply_all_btn = QPushButton(f"✓ Apply All Suggestions ({len(suggestions)})")
            apply_all_btn.setStyleSheet("background-color: #007bff; color: white; padding: 8px 16px; font-weight: bold;")
            apply_all_btn.clicked.connect(lambda: self.apply_all_trf_values(is_batch, dialog))
            btn_layout.addWidget(apply_all_btn)
        
        btn_layout.addStretch()
        
        # Close button
        close_btn = QPushButton("Close")
        close_btn.setStyleSheet("padding: 8px 16px;")
        close_btn.clicked.connect(dialog.accept)
        btn_layout.addWidget(close_btn)
        
        layout.addLayout(btn_layout)
        
        # Info text
        info = QLabel("💡 Tip: Click 'Apply' to use TRF value, or 'Apply All' to accept all suggestions")
        info.setStyleSheet("color: #666; font-style: italic; padding: 5px;")
        layout.addWidget(info)
        
        dialog.exec()
    
    def apply_single_trf_value(self, result, is_batch, dialog, table, row_idx):
        """Apply a single TRF value to the corresponding field"""
        field_key = result['field_key']
        trf_value = result['trf_value']
        widget_name = result['widget']
        
        if is_batch:
            widget_map = {
                'patient_name_input': self.batch_patient_name,
                'hospital_clinic_input': self.batch_hospital,
                'pin_input': self.batch_pin,
                'biopsy_date_input': self.batch_biopsy_date,
                'sample_receipt_date_input': self.batch_sample_receipt_date,
                'referring_clinician_input': self.batch_referring_clinician,
            }
        else:
            widget_map = {
                'patient_name_input': self.patient_name_input,
                'hospital_clinic_input': self.hospital_clinic_input,
                'pin_input': self.pin_input,
                'biopsy_date_input': self.biopsy_date_input,
                'sample_receipt_date_input': self.sample_receipt_date_input,
                'referring_clinician_input': self.referring_clinician_input,
            }
        
        if widget_name in widget_map:
            widget = widget_map[widget_name]
            if hasattr(widget, 'setText'):
                widget.setText(trf_value)
            elif hasattr(widget, 'setPlainText'):
                widget.setPlainText(trf_value)
            
            # Update table to show applied
            table.item(row_idx, 1).setText(trf_value)
            table.item(row_idx, 1).setBackground(QColor('#ccffcc'))
            table.item(row_idx, 3).setText("✓ Applied")
            table.item(row_idx, 3).setForeground(QColor('#28a745'))
            
            # Remove apply button
            table.removeCellWidget(row_idx, 4)
            table.setItem(row_idx, 4, QTableWidgetItem("Done"))
            
            self.statusBar().showMessage(f"Applied TRF value to {result['field']}")
    
    def apply_all_trf_values(self, is_batch, dialog):
        """Apply all TRF suggestions"""
        applied_count = 0
        
        for i, r in enumerate(self.trf_results):
            if r['can_apply'] and r['trf_value'] and r['trf_value'] != '(not found)':
                self.apply_single_trf_value(r, is_batch, dialog, self.trf_table, i)
                applied_count += 1
        
        QMessageBox.information(self, "Applied", f"Successfully applied {applied_count} values from TRF!")
        self.statusBar().showMessage(f"Applied {applied_count} values from TRF")
    
    def upload_trf_manual(self):
        """Upload TRF for manual entry verification"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select TRF Document",
            "",
            "TRF Files (*.pdf *.png *.jpg *.jpeg *.tiff *.bmp);;PDF Files (*.pdf);;Image Files (*.png *.jpg *.jpeg *.tiff *.bmp)"
        )
        
        if file_path:
            self.manual_trf_path = file_path
            self.trf_path_label.setText(os.path.basename(file_path))
            self.trf_path_label.setStyleSheet("color: #28a745; font-weight: bold;")
            self.trf_verify_btn.setEnabled(True)
            self.trf_result_text.setHtml("<i style='color:#007bff;'>TRF uploaded. Click 'Verify' to check patient details.</i>")
    
    def verify_trf_manual(self):
        """Verify manual entry data against uploaded TRF"""
        if not self.manual_trf_path:
            QMessageBox.warning(self, "No TRF", "Please upload a TRF document first.")
            return
        
        # Show processing
        self.trf_result_text.setHtml("<i style='color:#007bff;'>🔄 Extracting text from TRF...</i>")
        QApplication.processEvents()
        
        # Extract text from TRF
        trf_text, error = self.extract_text_from_trf(self.manual_trf_path)
        
        if error:
            self.trf_result_text.setHtml(f"<span style='color:#dc3545;'>❌ {error}</span>")
            return
        
        if not trf_text or len(trf_text.strip()) < 10:
            self.trf_result_text.setHtml("<span style='color:#dc3545;'>❌ Could not extract readable text from TRF. Try a clearer image or PDF.</span>")
            return
        
        # Get current patient data
        patient_data = {
            'patient_name': self.patient_name_input.toPlainText().strip(),
            'hospital_clinic': self.hospital_clinic_input.toPlainText().strip(),
            'pin': self.pin_input.toPlainText().strip(),
            'biopsy_date': self.biopsy_date_input.text().strip(),
            'sample_receipt_date': self.sample_receipt_date_input.text().strip(),
            'referring_clinician': self.referring_clinician_input.toPlainText().strip(),
        }
        
        # Enhanced verification
        results, all_correct, has_suggestions = self.verify_patient_data_enhanced(trf_text, patient_data)
        
        # Show comparison dialog
        self.show_trf_comparison_dialog(results, is_batch=False)
        
        # Update result text
        if all_correct:
            self.trf_result_text.setHtml("<span style='color:#28a745;'>✅ All fields verified! Click 'Verify' again to re-check.</span>")
        else:
            self.trf_result_text.setHtml("<span style='color:#ffc107;'>⚠️ Review completed. Some fields may need attention.</span>")
    
    def upload_trf_batch(self):
        """Upload TRF for batch entry verification"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select TRF Document",
            "",
            "TRF Files (*.pdf *.png *.jpg *.jpeg *.tiff *.bmp);;PDF Files (*.pdf);;Image Files (*.png *.jpg *.jpeg *.tiff *.bmp)"
        )
        
        if file_path:
            self.batch_trf_path = file_path
            self.batch_trf_path_label.setText(os.path.basename(file_path))
            self.batch_trf_path_label.setStyleSheet("color: #28a745; font-weight: bold;")
            self.batch_trf_result_text.setHtml("<i style='color:#007bff;'>TRF uploaded. Click 'Verify' to check patient details.</i>")
    
    def verify_trf_batch(self):
        """Verify batch entry data against uploaded TRF"""
        if not hasattr(self, 'batch_trf_path') or not self.batch_trf_path:
            QMessageBox.warning(self, "No TRF", "Please upload a TRF document first.")
            return
        
        # Show processing
        self.batch_trf_result_text.setHtml("<i style='color:#007bff;'>🔄 Extracting text from TRF...</i>")
        QApplication.processEvents()
        
        # Extract text from TRF
        trf_text, error = self.extract_text_from_trf(self.batch_trf_path)
        
        if error:
            self.batch_trf_result_text.setHtml(f"<span style='color:#dc3545;'>❌ {error}</span>")
            return
        
        if not trf_text or len(trf_text.strip()) < 10:
            self.batch_trf_result_text.setHtml("<span style='color:#dc3545;'>❌ Could not extract readable text from TRF. Try a clearer image or PDF.</span>")
            return
        
        # Get current batch patient data
        patient_data = {
            'patient_name': self.batch_patient_name.toPlainText().strip(),
            'hospital_clinic': self.batch_hospital.toPlainText().strip(),
            'pin': self.batch_pin.toPlainText().strip(),
            'biopsy_date': self.batch_biopsy_date.text().strip(),
            'sample_receipt_date': self.batch_sample_receipt_date.text().strip(),
            'referring_clinician': self.batch_referring_clinician.toPlainText().strip(),
        }
        
        # Enhanced verification
        results, all_correct, has_suggestions = self.verify_patient_data_enhanced(trf_text, patient_data)
        
        # Show comparison dialog
        self.show_trf_comparison_dialog(results, is_batch=True)
        
        # Update result text
        if all_correct:
            self.batch_trf_result_text.setHtml("<span style='color:#28a745;'>✅ All fields verified!</span>")
        else:
            self.batch_trf_result_text.setHtml("<span style='color:#ffc107;'>⚠️ Review completed. Some fields may need attention.</span>")
    
    # ==================== AI-Enhanced TRF Extraction (Open Source) ====================
    
    def init_easyocr_reader(self):
        """Initialize EasyOCR reader (lazy loading for performance)"""
        if not hasattr(self, '_easyocr_reader'):
            if EASYOCR_AVAILABLE:
                try:
                    # Use GPU if available, fall back to CPU
                    import torch
                    use_gpu = torch.cuda.is_available()
                except:
                    use_gpu = False
                try:
                    self._easyocr_reader = easyocr.Reader(['en'], gpu=use_gpu)
                except Exception as e:
                    print(f"EasyOCR init error: {e}")
                    self._easyocr_reader = None
            else:
                self._easyocr_reader = None
        return self._easyocr_reader
    
    def preprocess_image_for_ocr(self, image):
        """Preprocess image for better OCR accuracy (mobile camera/scanner images)"""
        from PIL import Image, ImageEnhance, ImageFilter, ImageOps
        
        try:
            # Convert to RGB if needed
            if image.mode != 'RGB':
                image = image.convert('RGB')
            
            # 1. Auto-orient based on EXIF (mobile photos often have orientation data)
            try:
                image = ImageOps.exif_transpose(image)
            except:
                pass
            
            # 2. Convert to grayscale for better text detection
            gray = image.convert('L')
            
            # 3. Increase contrast - helps with scanner/camera lighting issues
            enhancer = ImageEnhance.Contrast(gray)
            gray = enhancer.enhance(2.0)
            
            # 4. Increase sharpness - helps with slightly blurry camera photos
            enhancer = ImageEnhance.Sharpness(gray)
            gray = enhancer.enhance(2.0)
            
            # 5. Denoise using median filter (removes camera noise)
            gray = gray.filter(ImageFilter.MedianFilter(size=1))
            
            # 6. Binarize (convert to black and white) using adaptive threshold
            # This helps with uneven lighting from mobile cameras
            import numpy as np
            img_array = np.array(gray)
            
            # Simple adaptive thresholding
            threshold = np.mean(img_array)
            binary = ((img_array > threshold) * 255).astype(np.uint8)
            
            # Convert back to PIL Image
            processed = Image.fromarray(binary)
            
            # 7. Scale up small images (OCR works better on larger text)
            min_dimension = 1500
            if processed.width < min_dimension or processed.height < min_dimension:
                scale = max(min_dimension / processed.width, min_dimension / processed.height)
                new_size = (int(processed.width * scale), int(processed.height * scale))
                processed = processed.resize(new_size, Image.Resampling.LANCZOS)
            
            return processed.convert('RGB')  # EasyOCR needs RGB
            
        except Exception as e:
            print(f"Image preprocessing warning: {e}")
            return image  # Return original if preprocessing fails
    
    def extract_text_with_easyocr(self, file_path):
        """Extract text from image using EasyOCR (open-source, offline)
        
        Optimized for:
        - Mobile camera photos
        - Document scanner apps
        - Scanned PDFs
        """
        if not EASYOCR_AVAILABLE:
            return None, "EasyOCR not available. Install: pip install easyocr"
        
        try:
            from PIL import Image
            import io
            
            reader = self.init_easyocr_reader()
            if not reader:
                return None, "Failed to initialize EasyOCR"
            
            file_ext = os.path.splitext(file_path)[1].lower()
            
            # For PDFs, convert to image first with HIGH resolution for better OCR
            if file_ext == '.pdf':
                if PDFPLUMBER_AVAILABLE:
                    with pdfplumber.open(file_path) as pdf:
                        if pdf.pages:
                            page = pdf.pages[0]
                            # Higher resolution = better OCR accuracy
                            pil_img = page.to_image(resolution=300).original
                            
                            # Preprocess for better OCR
                            processed_img = self.preprocess_image_for_ocr(pil_img)
                            
                            img_buffer = io.BytesIO()
                            processed_img.save(img_buffer, format='PNG')
                            img_buffer.seek(0)
                            
                            # EasyOCR with optimized settings
                            results = reader.readtext(
                                img_buffer.getvalue(), 
                                detail=1, 
                                paragraph=False,
                                contrast_ths=0.1,  # Lower threshold for low contrast images
                                adjust_contrast=0.5,  # Auto-adjust contrast
                                text_threshold=0.6,  # Lower threshold for more text detection
                                low_text=0.3,  # Better detection for small text
                            )
                else:
                    return None, "PDF processing requires pdfplumber"
            else:
                # Load and preprocess image file
                pil_img = Image.open(file_path)
                processed_img = self.preprocess_image_for_ocr(pil_img)
                
                img_buffer = io.BytesIO()
                processed_img.save(img_buffer, format='PNG')
                img_buffer.seek(0)
                
                # EasyOCR with optimized settings for camera photos
                results = reader.readtext(
                    img_buffer.getvalue(),
                    detail=1, 
                    paragraph=False,
                    contrast_ths=0.1,
                    adjust_contrast=0.5,
                    text_threshold=0.6,
                    low_text=0.3,
                )
            
            # Combine all detected text - sort by vertical position for better reading order
            # results format: [[bbox, text, confidence], ...]
            sorted_results = sorted(results, key=lambda x: (x[0][0][1], x[0][0][0]))  # Sort by Y then X
            text_lines = [item[1] for item in sorted_results if item[2] > 0.25]  # Lower threshold for mobile photos
            full_text = '\n'.join(text_lines)
            
            return full_text, None
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            return None, f"EasyOCR error: {str(e)}"
    
    def extract_text_with_ollama(self, file_path):
        """Extract text from TRF using Ollama with LLaVA vision model (local AI)"""
        if not OLLAMA_AVAILABLE:
            return None, "requests library not available"
        
        # Check if Ollama is running
        ollama_url = self.settings.value('ollama_url', 'http://localhost:11434')
        
        try:
            # Test connection
            response = requests.get(f"{ollama_url}/api/tags", timeout=5)
            if response.status_code != 200:
                return None, f"Ollama not running at {ollama_url}"
        except requests.exceptions.ConnectionError:
            return None, f"Cannot connect to Ollama at {ollama_url}. Start Ollama first: ollama serve"
        except Exception as e:
            return None, f"Ollama connection error: {str(e)}"
        
        try:
            file_ext = os.path.splitext(file_path)[1].lower()
            
            # Read image and convert to base64
            if file_ext == '.pdf':
                if PDFPLUMBER_AVAILABLE:
                    with pdfplumber.open(file_path) as pdf:
                        if pdf.pages:
                            page = pdf.pages[0]
                            img = page.to_image(resolution=150)
                            import io
                            img_buffer = io.BytesIO()
                            img.save(img_buffer, format='PNG')
                            img_data = base64.b64encode(img_buffer.getvalue()).decode('utf-8')
                else:
                    return None, "PDF processing requires pdfplumber"
            else:
                with open(file_path, 'rb') as f:
                    img_data = base64.b64encode(f.read()).decode('utf-8')
            
            # Get model name from settings (default: llava)
            model_name = self.settings.value('ollama_vision_model', 'llava')
            
            # Call Ollama API with vision
            prompt = """You are extracting patient information from a medical Test Request Form (TRF).
Extract these fields and return ONLY a JSON object:
{
    "patient_name": "full patient name",
    "hospital_clinic": "hospital or clinic name", 
    "pin": "sample ID or PIN number",
    "biopsy_date": "date of biopsy (DD-MM-YYYY)",
    "sample_receipt_date": "date received (DD-MM-YYYY)",
    "referring_clinician": "doctor name"
}
Use null for fields not found. Return ONLY valid JSON."""

            response = requests.post(
                f"{ollama_url}/api/generate",
                json={
                    "model": model_name,
                    "prompt": prompt,
                    "images": [img_data],
                    "stream": False,
                    "options": {
                        "temperature": 0.1
                    }
                },
                timeout=120  # Vision models can be slow
            )
            
            if response.status_code != 200:
                return None, f"Ollama API error: {response.status_code}"
            
            result = response.json()
            result_text = result.get('response', '')
            
            # Try to parse as JSON
            try:
                import json
                # Find JSON in response
                json_match = re.search(r'\{[^{}]*\}', result_text, re.DOTALL)
                if json_match:
                    extracted_data = json.loads(json_match.group())
                    return extracted_data, None
                else:
                    return result_text, None
            except json.JSONDecodeError:
                return result_text, None
                
        except Exception as e:
            return None, f"Ollama extraction error: {str(e)}"
    
    def extract_text_enhanced(self, file_path, method='auto'):
        """Enhanced text extraction with multiple methods
        
        Methods:
        - 'auto': Try best available method
        - 'easyocr': Use EasyOCR (recommended, offline)
        - 'ollama': Use Ollama LLaVA (local AI)
        - 'tesseract': Use pytesseract (basic)
        """
        errors = []
        
        if method == 'auto':
            # Try methods in order of preference
            methods_to_try = []
            
            # Check settings for preferred method
            preferred = self.settings.value('trf_extraction_method', 'easyocr')
            
            if preferred == 'ollama' and OLLAMA_AVAILABLE:
                methods_to_try = ['ollama', 'easyocr', 'tesseract']
            elif EASYOCR_AVAILABLE:
                methods_to_try = ['easyocr', 'tesseract']
            else:
                methods_to_try = ['tesseract']
            
            for m in methods_to_try:
                result, error = self.extract_text_enhanced(file_path, method=m)
                if result:
                    return result, None
                if error:
                    errors.append(f"{m}: {error}")
            
            return None, "All extraction methods failed: " + "; ".join(errors)
        
        elif method == 'easyocr':
            return self.extract_text_with_easyocr(file_path)
        
        elif method == 'ollama':
            return self.extract_text_with_ollama(file_path)
        
        elif method == 'tesseract':
            return self.extract_text_from_trf(file_path)
        
        else:
            return None, f"Unknown extraction method: {method}"
    
    def verify_with_ai(self, file_path, patient_data, use_ai=True):
        """Enhanced verification using local AI when available"""
        # Try Ollama extraction first if enabled (returns structured data)
        if use_ai:
            preferred_method = self.settings.value('trf_extraction_method', 'easyocr')
            
            if preferred_method == 'ollama' and OLLAMA_AVAILABLE:
                ai_result, ai_error = self.extract_text_with_ollama(file_path)
                if ai_result and isinstance(ai_result, dict):
                    return self.compare_structured_data(ai_result, patient_data)
        
        # Fallback to OCR/text extraction
        trf_text, error = self.extract_text_enhanced(file_path)
        if error:
            return [], False, False, error
        
        if isinstance(trf_text, dict):
            # Ollama returned structured data
            return self.compare_structured_data(trf_text, patient_data)
        
        results, all_correct, has_suggestions = self.verify_patient_data_enhanced(trf_text, patient_data)
        return results, all_correct, has_suggestions, None
    
    def compare_structured_data(self, ai_data, patient_data):
        """Compare AI-extracted structured data with patient data"""
        results = []
        all_correct = True
        has_suggestions = False
        
        field_mapping = [
            ('patient_name', 'Patient Name', 'patient_name_input'),
            ('hospital_clinic', 'Hospital/Clinic', 'hospital_clinic_input'),
            ('pin', 'PIN/Sample ID', 'pin_input'),
            ('biopsy_date', 'Biopsy Date', 'biopsy_date_input'),
            ('sample_receipt_date', 'Receipt Date', 'sample_receipt_date_input'),
            ('referring_clinician', 'Referring Clinician', 'referring_clinician_input'),
        ]
        
        for field_key, field_label, widget_name in field_mapping:
            entered_value = patient_data.get(field_key, '') or ''
            ai_value = ai_data.get(field_key) or ''
            
            # Clean values
            entered_clean = str(entered_value).strip().lower()
            ai_clean = str(ai_value).strip().lower() if ai_value else ''
            
            # Skip if both empty
            if not entered_clean and not ai_clean:
                results.append({
                    'field': field_label,
                    'field_key': field_key,
                    'widget': widget_name,
                    'status': 'skip',
                    'entered': '(empty)',
                    'trf_value': '(not found)',
                    'message': 'No data',
                    'icon': '⚪',
                    'can_apply': False
                })
                continue
            
            # Check match using fuzzy matching
            if entered_clean and ai_clean:
                is_match, ratio = self.fuzzy_match(entered_clean, ai_clean, 0.8)
                
                if is_match:
                    results.append({
                        'field': field_label,
                        'field_key': field_key,
                        'widget': widget_name,
                        'status': 'ok',
                        'entered': entered_value,
                        'trf_value': ai_value,
                        'message': f'✓ Match ({int(ratio*100)}%)',
                        'icon': '✅',
                        'can_apply': False
                    })
                else:
                    all_correct = False
                    has_suggestions = True
                    results.append({
                        'field': field_label,
                        'field_key': field_key,
                        'widget': widget_name,
                        'status': 'mismatch',
                        'entered': entered_value,
                        'trf_value': ai_value,
                        'message': f'✗ Mismatch ({int(ratio*100)}%)',
                        'icon': '❌',
                        'can_apply': True
                    })
            elif ai_clean and not entered_clean:
                # AI found value but nothing entered - suggest
                has_suggestions = True
                results.append({
                    'field': field_label,
                    'field_key': field_key,
                    'widget': widget_name,
                    'status': 'suggestion',
                    'entered': '(empty)',
                    'trf_value': ai_value,
                    'message': '💡 Found in TRF (AI)',
                    'icon': '💡',
                    'can_apply': True
                })
            else:
                # Entered but not in AI result
                results.append({
                    'field': field_label,
                    'field_key': field_key,
                    'widget': widget_name,
                    'status': 'warning',
                    'entered': entered_value,
                    'trf_value': '(not found)',
                    'message': '⚠ Not in TRF',
                    'icon': '⚠️',
                    'can_apply': False
                })
        
        return results, all_correct, has_suggestions, None
    
    # ==================== Bulk TRF Verification ====================
    
    def upload_bulk_trf(self):
        """Upload TRF files - supports single multi-page PDF or multiple individual files"""
        # Ask user what type of upload
        msg = QMessageBox(self)
        msg.setWindowTitle("TRF Upload Type")
        msg.setText("How are your TRF files organized?")
        msg.setInformativeText("Choose the format that matches your TRF documents:")
        
        single_btn = msg.addButton("📄 Single Multi-Page PDF\n(All TRFs in one file)", QMessageBox.ButtonRole.ActionRole)
        multiple_btn = msg.addButton("📁 Multiple Individual Files\n(One file per TRF)", QMessageBox.ButtonRole.ActionRole)
        cancel_btn = msg.addButton(QMessageBox.StandardButton.Cancel)
        
        msg.exec()
        
        if msg.clickedButton() == single_btn:
            self.upload_bulk_trf_single_pdf()
        elif msg.clickedButton() == multiple_btn:
            self.upload_bulk_trf_multiple_files()
    
    def upload_bulk_trf_single_pdf(self):
        """Upload a single multi-page PDF containing all TRFs"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Bulk TRF PDF (Multi-Page)",
            "",
            "PDF Files (*.pdf)"
        )
        
        if not file_path:
            return
        
        if not PDFPLUMBER_AVAILABLE:
            QMessageBox.warning(self, "Error", "PDF processing requires pdfplumber. Install: pip install pdfplumber")
            return
        
        # Count pages in PDF
        try:
            with pdfplumber.open(file_path) as pdf:
                num_pages = len(pdf.pages)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to open PDF: {str(e)}")
            return
        
        if num_pages == 0:
            QMessageBox.warning(self, "Error", "PDF has no pages")
            return
        
        # Store the bulk PDF info
        self.bulk_trf_pdf_path = file_path
        self.bulk_trf_pdf_pages = num_pages
        self.bulk_trf_files = []  # Clear individual files
        self.bulk_trf_is_single_pdf = True
        
        self.bulk_trf_label.setText(f"📄 {os.path.basename(file_path)} ({num_pages} pages)")
        self.bulk_trf_label.setStyleSheet("color: #28a745; font-weight: bold;")
        self.bulk_trf_verify_all_btn.setEnabled(True)
        
        self.bulk_trf_status.setHtml(
            f"<b>Bulk TRF PDF loaded:</b><br>"
            f"📄 {os.path.basename(file_path)}<br>"
            f"📑 {num_pages} pages detected<br><br>"
            f"<i>Click 'Verify All' to auto-match pages to patients,<br>"
            f"or use 'TRF Manager' for manual page assignment.</i>"
        )
    
    def upload_bulk_trf_multiple_files(self):
        """Upload multiple individual TRF files"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Multiple TRF Documents",
            "",
            "TRF Files (*.pdf *.png *.jpg *.jpeg *.tiff *.bmp);;PDF Files (*.pdf);;Image Files (*.png *.jpg *.jpeg *.tiff *.bmp)"
        )
        
        if file_paths:
            self.bulk_trf_files = file_paths
            self.bulk_trf_is_single_pdf = False
            self.bulk_trf_pdf_path = None
            
            self.bulk_trf_label.setText(f"{len(file_paths)} TRF files selected")
            self.bulk_trf_label.setStyleSheet("color: #28a745; font-weight: bold;")
            self.bulk_trf_verify_all_btn.setEnabled(True)
            
            file_names = [os.path.basename(f) for f in file_paths]
            self.bulk_trf_status.setHtml(
                f"<b>Selected TRF files ({len(file_paths)}):</b><br>" +
                "<br>".join([f"📄 {name}" for name in file_names[:10]]) +
                (f"<br><i>... and {len(file_paths) - 10} more</i>" if len(file_paths) > 10 else "")
            )
    
    def extract_page_from_bulk_pdf(self, page_number):
        """Extract a single page from the bulk TRF PDF and return as image bytes"""
        if not hasattr(self, 'bulk_trf_pdf_path') or not self.bulk_trf_pdf_path:
            return None, "No bulk PDF loaded"
        
        try:
            with pdfplumber.open(self.bulk_trf_pdf_path) as pdf:
                if page_number < 0 or page_number >= len(pdf.pages):
                    return None, f"Page {page_number + 1} out of range"
                
                page = pdf.pages[page_number]
                
                # Convert to image for OCR
                img = page.to_image(resolution=200)
                import io
                img_buffer = io.BytesIO()
                img.save(img_buffer, format='PNG')
                img_buffer.seek(0)
                
                return img_buffer.getvalue(), None
        except Exception as e:
            return None, f"Error extracting page: {str(e)}"
    
    def extract_text_from_bulk_pdf_page(self, page_number):
        """Extract text from a specific page of the bulk TRF PDF"""
        if not hasattr(self, 'bulk_trf_pdf_path') or not self.bulk_trf_pdf_path:
            return None, "No bulk PDF loaded"
        
        try:
            with pdfplumber.open(self.bulk_trf_pdf_path) as pdf:
                if page_number < 0 or page_number >= len(pdf.pages):
                    return None, f"Page {page_number + 1} out of range"
                
                page = pdf.pages[page_number]
                
                # Try direct text extraction first
                text = page.extract_text()
                
                if text and len(text.strip()) > 20:
                    return text, None
                
                # Fallback to OCR
                extraction_method = self.settings.value('trf_extraction_method', 'easyocr')
                
                if extraction_method == 'easyocr' and EASYOCR_AVAILABLE:
                    img = page.to_image(resolution=200)
                    import io
                    img_buffer = io.BytesIO()
                    img.save(img_buffer, format='PNG')
                    img_buffer.seek(0)
                    
                    reader = self.init_easyocr_reader()
                    if reader:
                        results = reader.readtext(img_buffer.getvalue())
                        text_lines = [item[1] for item in results]
                        return '\n'.join(text_lines), None
                
                # Try Tesseract
                if TESSERACT_AVAILABLE:
                    img = page.to_image(resolution=200)
                    import io
                    img_buffer = io.BytesIO()
                    img.save(img_buffer, format='PNG')
                    img_buffer.seek(0)
                    
                    pil_img = Image.open(img_buffer)
                    text = pytesseract.image_to_string(pil_img)
                    return text, None
                
                return text or "", None
                
        except Exception as e:
            return None, f"Error extracting text: {str(e)}"
    
    def verify_all_bulk_trf(self):
        """Verify all uploaded TRFs against all patients in the batch"""
        # Check if we have bulk PDF or individual files
        is_single_pdf = getattr(self, 'bulk_trf_is_single_pdf', False)
        
        if is_single_pdf:
            if not hasattr(self, 'bulk_trf_pdf_path') or not self.bulk_trf_pdf_path:
                QMessageBox.warning(self, "No TRF", "Please upload a bulk TRF PDF first.")
                return
        else:
            if not hasattr(self, 'bulk_trf_files') or not self.bulk_trf_files:
                QMessageBox.warning(self, "No TRFs", "Please upload TRF files first.")
                return
        
        if not hasattr(self, 'bulk_patient_data_list') or not self.bulk_patient_data_list:
            QMessageBox.warning(self, "No Patients", "Please load patient data first using an Excel file.")
            return
        
        if is_single_pdf:
            self.verify_bulk_trf_single_pdf()
        else:
            self.verify_bulk_trf_multiple_files()
    
    def get_trf_patients_dict(self):
        """Get patients data as dictionary for TRF matching"""
        patients_dict = {}
        if hasattr(self, 'bulk_patient_data_list') and self.bulk_patient_data_list:
            for i, data in enumerate(self.bulk_patient_data_list):
                p_info = data.get('patient_info', {})
                p_name = p_info.get('patient_name', f'Patient_{i}')
                # Use index as unique key to handle duplicate names
                patient_key = f"{p_name}_{i}"
                patients_dict[patient_key] = {
                    'patient_info': p_info,
                    'embryos': data.get('embryos', []),
                    'list_index': i  # Store original index for updates
                }
        return patients_dict
    
    def verify_bulk_trf_single_pdf(self):
        """Verify all pages in bulk TRF PDF against patients"""
        num_pages = getattr(self, 'bulk_trf_pdf_pages', 0)
        if num_pages == 0:
            return
        
        # Show progress dialog
        progress_dialog = QDialog(self)
        progress_dialog.setWindowTitle("Bulk TRF Verification (Multi-Page PDF)")
        progress_dialog.setMinimumWidth(600)
        progress_layout = QVBoxLayout()
        progress_dialog.setLayout(progress_layout)
        
        status_label = QLabel(f"🔄 Processing {num_pages} pages from bulk TRF...")
        status_label.setStyleSheet("font-size: 14px; padding: 10px;")
        progress_layout.addWidget(status_label)
        
        progress_bar = QProgressBar()
        progress_bar.setMaximum(num_pages)
        progress_layout.addWidget(progress_bar)
        
        results_text = QTextBrowser()
        results_text.setMinimumHeight(350)
        progress_layout.addWidget(results_text)
        
        close_btn = QPushButton("Close")
        close_btn.setEnabled(False)
        close_btn.clicked.connect(progress_dialog.accept)
        progress_layout.addWidget(close_btn)
        
        progress_dialog.show()
        QApplication.processEvents()
        
        # Get patient list for matching using helper
        patients_dict = self.get_trf_patients_dict()
        patients_list = list(patients_dict.items())
        matched_count = 0
        unmatched_pages = []
        page_patient_matches = {}
        
        results_text.append(f"<b>Processing {num_pages} pages, {len(patients_list)} patients...</b><br>")
        
        for page_idx in range(num_pages):
            status_label.setText(f"🔄 Processing page {page_idx + 1}/{num_pages}...")
            progress_bar.setValue(page_idx + 1)
            QApplication.processEvents()
            
            # Extract text from this page
            page_text, error = self.extract_text_from_bulk_pdf_page(page_idx)
            
            if error or not page_text:
                unmatched_pages.append((page_idx + 1, error or "No text extracted"))
                results_text.append(f"⚠️ Page {page_idx + 1}: {error or 'No text extracted'}")
                continue
            
            # Try to match with a patient
            best_match = None
            best_score = 0
            
            for patient_name, patient_data in patients_list:
                p_info = patient_data.get('patient_info', {})
                p_name = p_info.get('patient_name', '')
                p_pin = p_info.get('pin', '')
                
                score = 0
                
                if p_name:
                    found, _ = self.find_in_text(page_text, p_name)
                    if found:
                        score += 60
                
                if p_pin:
                    found, _ = self.find_in_text(page_text, p_pin)
                    if found:
                        score += 40
                
                if score > best_score:
                    best_score = score
                    best_match = (patient_name, patient_data)
            
            if best_match and best_score >= 50:
                matched_count += 1
                patient_name, patient_data = best_match
                
                # Store page-patient mapping
                page_patient_matches[page_idx] = patient_name
                
                # Store TRF association
                if not hasattr(self, 'patient_trf_mapping'):
                    self.patient_trf_mapping = {}
                
                self.patient_trf_mapping[patient_name] = {
                    'trf_path': self.bulk_trf_pdf_path,
                    'trf_page': page_idx,
                    'trf_data': None,
                    'match_score': best_score,
                    'is_bulk_pdf': True
                }
                
                results_text.append(f"✅ Page {page_idx + 1} → <b>{patient_name}</b> ({int(best_score)}%)")
            else:
                unmatched_pages.append((page_idx + 1, f"Best score: {int(best_score)}%"))
                results_text.append(f"⚠️ Page {page_idx + 1} - No match (best: {int(best_score)}%)")
            
            QApplication.processEvents()
        
        # Summary
        status_label.setText("✅ Bulk TRF verification complete!")
        results_text.append(f"\n<hr><b>Summary:</b>")
        results_text.append(f"• Total pages: {num_pages}")
        results_text.append(f"• Matched: {matched_count}")
        results_text.append(f"• Unmatched: {len(unmatched_pages)}")
        
        if unmatched_pages:
            results_text.append(f"\n<b>Unmatched pages:</b>")
            for page_num, reason in unmatched_pages[:10]:
                results_text.append(f"  • Page {page_num}: {reason}")
            if len(unmatched_pages) > 10:
                results_text.append(f"  ... and {len(unmatched_pages) - 10} more")
        
        close_btn.setEnabled(True)
        
        # Update status
        self.bulk_trf_status.setHtml(
            f"<span style='color:#28a745;'>✅ Processed {num_pages} pages: "
            f"{matched_count} matched, {len(unmatched_pages)} unmatched</span>"
        )
    
    def verify_bulk_trf_multiple_files(self):
        """Verify all uploaded TRFs against all patients in the batch"""
        if not hasattr(self, 'bulk_trf_files') or not self.bulk_trf_files:
            QMessageBox.warning(self, "No TRFs", "Please upload TRF files first.")
            return
        
        if not hasattr(self, 'bulk_patient_data_list') or not self.bulk_patient_data_list:
            QMessageBox.warning(self, "No Patients", "Please load patient data first using an Excel file.")
            return
        
        # Show progress dialog
        progress_dialog = QDialog(self)
        progress_dialog.setWindowTitle("Bulk TRF Verification")
        progress_dialog.setMinimumWidth(500)
        progress_layout = QVBoxLayout()
        progress_dialog.setLayout(progress_layout)
        
        status_label = QLabel("🔄 Processing TRF files...")
        status_label.setStyleSheet("font-size: 14px; padding: 10px;")
        progress_layout.addWidget(status_label)
        
        progress_bar = QProgressBar()
        progress_bar.setMaximum(len(self.bulk_trf_files))
        progress_layout.addWidget(progress_bar)
        
        results_text = QTextBrowser()
        results_text.setMinimumHeight(300)
        progress_layout.addWidget(results_text)
        
        close_btn = QPushButton("Close")
        close_btn.setEnabled(False)
        close_btn.clicked.connect(progress_dialog.accept)
        progress_layout.addWidget(close_btn)
        
        progress_dialog.show()
        QApplication.processEvents()
        
        # Process each TRF
        all_results = []
        matched_count = 0
        unmatched_trfs = []
        
        # Get extraction method from settings
        extraction_method = self.settings.value('trf_extraction_method', 'easyocr')
        
        for i, trf_path in enumerate(self.bulk_trf_files):
            trf_name = os.path.basename(trf_path)
            status_label.setText(f"🔄 Processing: {trf_name} ({i+1}/{len(self.bulk_trf_files)})")
            progress_bar.setValue(i + 1)
            QApplication.processEvents()
            
            # Extract data from TRF using selected method
            trf_data = None
            trf_text = None
            
            if extraction_method == 'ollama' and OLLAMA_AVAILABLE:
                # Ollama returns structured data
                trf_data, error = self.extract_text_with_ollama(trf_path)
                if error or not isinstance(trf_data, dict):
                    # Fallback to EasyOCR
                    trf_text, error = self.extract_text_enhanced(trf_path, method='easyocr')
                    trf_data = None
            else:
                # Use EasyOCR or Tesseract
                trf_text, error = self.extract_text_enhanced(trf_path, method=extraction_method)
            
            if error and not trf_text and not trf_data:
                unmatched_trfs.append((trf_name, f"Error: {error}"))
                continue
            
            # Try to match with a patient
            best_match = None
            best_score = 0
            
            patients_dict = self.get_trf_patients_dict()
            for patient_name, patient_data in patients_dict.items():
                p_name = patient_data.get('patient_info', {}).get('patient_name', '')
                p_pin = patient_data.get('patient_info', {}).get('pin', '')
                
                # Calculate match score
                score = 0
                if trf_data and isinstance(trf_data, dict):
                    # AI extracted data
                    trf_name_val = trf_data.get('patient_name', '') or ''
                    trf_pin_val = trf_data.get('pin', '') or ''
                    
                    if p_name and trf_name_val:
                        is_match, ratio = self.fuzzy_match(p_name, trf_name_val, 0.6)
                        if is_match:
                            score += ratio * 60
                    
                    if p_pin and trf_pin_val:
                        if p_pin.lower() in trf_pin_val.lower() or trf_pin_val.lower() in p_pin.lower():
                            score += 40
                else:
                    # Text-based matching
                    if trf_text:
                        if p_name:
                            found, _ = self.find_in_text(trf_text, p_name)
                            if found:
                                score += 60
                        if p_pin:
                            found, _ = self.find_in_text(trf_text, p_pin)
                            if found:
                                score += 40
                
                if score > best_score:
                    best_score = score
                    best_match = (patient_name, patient_data)
            
            if best_match and best_score >= 50:
                matched_count += 1
                patient_name, patient_data = best_match
                
                # Store TRF association
                if not hasattr(self, 'patient_trf_mapping'):
                    self.patient_trf_mapping = {}
                self.patient_trf_mapping[patient_name] = {
                    'trf_path': trf_path,
                    'trf_data': trf_data if isinstance(trf_data, dict) else None,
                    'match_score': best_score
                }
                
                all_results.append({
                    'trf': trf_name,
                    'patient': patient_name,
                    'score': best_score,
                    'status': 'matched'
                })
                
                results_text.append(f"✅ <b>{trf_name}</b> → {patient_name} ({int(best_score)}% confidence)")
            else:
                unmatched_trfs.append((trf_name, "No matching patient found"))
                all_results.append({
                    'trf': trf_name,
                    'patient': None,
                    'score': best_score,
                    'status': 'unmatched'
                })
                results_text.append(f"⚠️ <b>{trf_name}</b> - No match (best: {int(best_score)}%)")
            
            QApplication.processEvents()
        
        # Show summary
        status_label.setText("✅ Bulk TRF verification complete!")
        results_text.append(f"\n<hr><b>Summary:</b>")
        results_text.append(f"• Total TRFs: {len(self.bulk_trf_files)}")
        results_text.append(f"• Matched: {matched_count}")
        results_text.append(f"• Unmatched: {len(unmatched_trfs)}")
        
        if unmatched_trfs:
            results_text.append(f"\n<b>Unmatched TRFs:</b>")
            for name, reason in unmatched_trfs:
                results_text.append(f"  • {name}: {reason}")
        
        close_btn.setEnabled(True)
        
        # Store results
        self.bulk_trf_results = all_results
        
        # Update main status
        self.bulk_trf_status.setHtml(
            f"<span style='color:#28a745;'>✅ Verified {len(self.bulk_trf_files)} TRFs: "
            f"{matched_count} matched, {len(unmatched_trfs)} unmatched</span>"
        )
    
    def show_bulk_trf_verification_dialog(self):
        """Show comprehensive bulk TRF verification dialog"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Bulk TRF Verification")
        dialog.setMinimumWidth(900)
        dialog.setMinimumHeight(600)
        
        layout = QVBoxLayout()
        dialog.setLayout(layout)
        
        # Header
        header = QLabel("📋 Bulk TRF Verification")
        header.setStyleSheet("font-size: 18px; font-weight: bold; padding: 10px;")
        layout.addWidget(header)
        
        # Simple status info - EasyOCR only
        status_group = QGroupBox("🔧 OCR Engine: EasyOCR (Offline)")
        status_layout = QHBoxLayout()
        status_group.setLayout(status_layout)
        
        if EASYOCR_AVAILABLE:
            status_label = QLabel("✅ EasyOCR Ready - High accuracy offline text extraction")
            status_label.setStyleSheet("color: #28a745; font-weight: bold;")
        else:
            status_label = QLabel("❌ EasyOCR not installed. Run: pip install easyocr")
            status_label.setStyleSheet("color: #dc3545; font-weight: bold;")
        status_layout.addWidget(status_label)
        
        layout.addWidget(status_group)
        
        # Splitter for list and details
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Left side - Patient list with TRF status
        left_widget = QWidget()
        left_layout = QVBoxLayout()
        left_widget.setLayout(left_layout)
        
        left_layout.addWidget(QLabel("<b>Patients & TRF Status:</b>"))
        
        self.trf_patient_list = QTableWidget()
        self.trf_patient_list.setColumnCount(4)
        self.trf_patient_list.setHorizontalHeaderLabels(["Patient", "PIN", "TRF Status", "Action"])
        self.trf_patient_list.horizontalHeader().setStretchLastSection(True)
        self.trf_patient_list.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.trf_patient_list.itemSelectionChanged.connect(self.on_trf_patient_selected)
        left_layout.addWidget(self.trf_patient_list)
        
        splitter.addWidget(left_widget)
        
        # Right side - TRF details and verification
        right_widget = QWidget()
        right_layout = QVBoxLayout()
        right_widget.setLayout(right_layout)
        
        right_layout.addWidget(QLabel("<b>Verification Details:</b>"))
        
        self.trf_details_text = QTextBrowser()
        self.trf_details_text.setMinimumWidth(400)
        right_layout.addWidget(self.trf_details_text)
        
        # Action buttons for selected patient
        action_layout = QHBoxLayout()
        
        upload_single_btn = QPushButton("📄 Upload TRF for Selected")
        upload_single_btn.clicked.connect(self.upload_trf_for_selected_patient)
        action_layout.addWidget(upload_single_btn)
        
        verify_single_btn = QPushButton("✓ Verify Selected")
        verify_single_btn.clicked.connect(self.verify_selected_patient_trf)
        action_layout.addWidget(verify_single_btn)
        
        right_layout.addLayout(action_layout)
        
        splitter.addWidget(right_widget)
        splitter.setSizes([450, 450])
        
        layout.addWidget(splitter)
        
        # Bulk TRF status label
        self.bulk_trf_status_label = QLabel("No bulk TRF uploaded")
        self.bulk_trf_status_label.setStyleSheet("color: #666; font-style: italic; padding: 5px;")
        layout.addWidget(self.bulk_trf_status_label)
        
        # Bottom buttons
        bottom_layout = QHBoxLayout()
        
        bulk_pdf_btn = QPushButton("📑 Upload Bulk TRF PDF")
        bulk_pdf_btn.clicked.connect(self.upload_bulk_trf_pdf_for_dialog)
        bulk_pdf_btn.setStyleSheet("padding: 8px 16px; background-color: #17a2b8; color: white;")
        bulk_pdf_btn.setToolTip("Upload a single PDF containing all TRFs (one page per patient)")
        bottom_layout.addWidget(bulk_pdf_btn)
        
        auto_match_btn = QPushButton("🔄 Auto-Match All")
        auto_match_btn.clicked.connect(self.auto_match_bulk_pdf_pages)
        auto_match_btn.setStyleSheet("padding: 8px 16px; background-color: #007bff; color: white;")
        auto_match_btn.setToolTip("Automatically match PDF pages to patients using OCR")
        bottom_layout.addWidget(auto_match_btn)
        
        bottom_layout.addStretch()
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.accept)
        close_btn.setStyleSheet("padding: 8px 16px;")
        bottom_layout.addWidget(close_btn)
        
        layout.addLayout(bottom_layout)
        
        # Populate patient list
        self.populate_trf_patient_list()
        
        dialog.exec()
    
    def upload_bulk_trf_pdf_for_dialog(self):
        """Upload a bulk TRF PDF from the TRF Manager dialog"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Bulk TRF PDF (All TRFs in one file)",
            "",
            "PDF Files (*.pdf)"
        )
        
        if not file_path:
            return
        
        if not PDFPLUMBER_AVAILABLE:
            QMessageBox.warning(self, "Error", "PDF processing requires pdfplumber. Install: pip install pdfplumber")
            return
        
        try:
            with pdfplumber.open(file_path) as pdf:
                num_pages = len(pdf.pages)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to open PDF: {str(e)}")
            return
        
        if num_pages == 0:
            QMessageBox.warning(self, "Error", "PDF has no pages")
            return
        
        # Store the bulk PDF info
        self.dialog_bulk_pdf_path = file_path
        self.dialog_bulk_pdf_pages = num_pages
        
        # Update status label
        self.bulk_trf_status_label.setText(f"📄 {os.path.basename(file_path)} - {num_pages} pages loaded")
        self.bulk_trf_status_label.setStyleSheet("color: #28a745; font-weight: bold; padding: 5px;")
        
        QMessageBox.information(
            self,
            "Bulk TRF Loaded",
            f"Loaded: {os.path.basename(file_path)}\n"
            f"Pages: {num_pages}\n\n"
            f"Click 'Auto-Match All' to automatically match pages to patients."
        )
    
    def auto_match_bulk_pdf_pages(self):
        """Auto-match bulk PDF pages to patients using EasyOCR"""
        if not hasattr(self, 'dialog_bulk_pdf_path') or not self.dialog_bulk_pdf_path:
            QMessageBox.warning(self, "No PDF", "Please upload a bulk TRF PDF first.")
            return
        
        if not hasattr(self, 'bulk_patient_data_list') or not self.bulk_patient_data_list:
            QMessageBox.warning(self, "No Patients", "No patient data loaded.")
            return
        
        if not EASYOCR_AVAILABLE:
            QMessageBox.warning(self, "Error", "EasyOCR is required. Install: pip install easyocr")
            return
        
        num_pages = self.dialog_bulk_pdf_pages
        patients_dict = self.get_trf_patients_dict()
        
        # Progress dialog
        progress = QProgressDialog("Processing TRF pages...", "Cancel", 0, num_pages, self)
        progress.setWindowTitle("Auto-Matching TRFs")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)
        progress.show()
        
        matched_count = 0
        results_log = []
        
        try:
            with pdfplumber.open(self.dialog_bulk_pdf_path) as pdf:
                for page_idx in range(num_pages):
                    if progress.wasCanceled():
                        break
                    
                    progress.setValue(page_idx)
                    progress.setLabelText(f"Processing page {page_idx + 1} of {num_pages}...")
                    QApplication.processEvents()
                    
                    # Extract text from this page
                    page = pdf.pages[page_idx]
                    
                    # First try direct text extraction (faster)
                    page_text = page.extract_text() or ""
                    
                    # If no text or insufficient text, use enhanced OCR with preprocessing
                    if len(page_text.strip()) < 50:
                        try:
                            from PIL import Image
                            import io
                            
                            reader = self.init_easyocr_reader()
                            if reader:
                                # Get page as image
                                pil_img = page.to_image(resolution=300).original
                                
                                # Apply preprocessing for mobile/scanner images
                                processed_img = self.preprocess_image_for_ocr(pil_img)
                                
                                img_buffer = io.BytesIO()
                                processed_img.save(img_buffer, format='PNG')
                                img_buffer.seek(0)
                                
                                # OCR with optimized settings
                                ocr_results = reader.readtext(
                                    img_buffer.getvalue(), 
                                    detail=0,
                                    contrast_ths=0.1,
                                    adjust_contrast=0.5,
                                    text_threshold=0.6,
                                    low_text=0.3,
                                )
                                page_text = ' '.join(ocr_results)
                        except Exception as e:
                            results_log.append(f"Page {page_idx + 1}: OCR error - {str(e)}")
                            continue
                    
                    if not page_text.strip():
                        results_log.append(f"Page {page_idx + 1}: No text found")
                        continue
                    
                    # Find best matching patient
                    best_match = None
                    best_score = 0
                    
                    for patient_key, patient_data in patients_dict.items():
                        p_info = patient_data.get('patient_info', {})
                        p_name = p_info.get('patient_name', '')
                        p_pin = p_info.get('pin', '')
                        
                        score = 0
                        page_text_lower = page_text.lower()
                        
                        # Check name match
                        if p_name:
                            name_parts = p_name.lower().split()
                            for part in name_parts:
                                if len(part) > 2 and part in page_text_lower:
                                    score += 30
                        
                        # Check PIN match (exact or partial)
                        if p_pin:
                            if p_pin.lower() in page_text_lower:
                                score += 50
                            elif any(part in page_text_lower for part in p_pin.lower().split() if len(part) > 3):
                                score += 25
                        
                        if score > best_score:
                            best_score = score
                            best_match = patient_key
                    
                    # Match if score is good enough
                    if best_match and best_score >= 30:
                        if not hasattr(self, 'patient_trf_mapping'):
                            self.patient_trf_mapping = {}
                        
                        self.patient_trf_mapping[best_match] = {
                            'trf_path': self.dialog_bulk_pdf_path,
                            'trf_page': page_idx,
                            'trf_data': None,
                            'match_score': best_score,
                            'is_bulk_pdf': True
                        }
                        matched_count += 1
                        
                        # Get display name
                        p_name = patients_dict[best_match].get('patient_info', {}).get('patient_name', best_match)
                        results_log.append(f"Page {page_idx + 1}: ✅ Matched to {p_name} (score: {best_score})")
                    else:
                        results_log.append(f"Page {page_idx + 1}: ❌ No match found (best score: {best_score})")
        
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Processing error: {str(e)}")
            return
        finally:
            progress.setValue(num_pages)
        
        # Update the patient list
        self.populate_trf_patient_list()
        
        # Show results
        msg = f"✅ Matched {matched_count} of {num_pages} pages to patients.\n\n"
        msg += "Results:\n" + "\n".join(results_log[:15])
        if len(results_log) > 15:
            msg += f"\n... and {len(results_log) - 15} more"
        
        QMessageBox.information(self, "Auto-Match Complete", msg)
    
    def test_ollama_connection(self):
        """Test connection to Ollama server"""
        if not OLLAMA_AVAILABLE:
            QMessageBox.warning(self, "Not Available", "requests library not installed")
            return
        
        ollama_url = self.settings.value('ollama_url', 'http://localhost:11434')
        
        try:
            response = requests.get(f"{ollama_url}/api/tags", timeout=5)
            if response.status_code == 200:
                data = response.json()
                models = [m['name'] for m in data.get('models', [])]
                vision_models = [m for m in models if any(v in m.lower() for v in ['llava', 'bakllava', 'moondream'])]
                
                msg = f"✅ Connected to Ollama at {ollama_url}\n\n"
                msg += f"Total models: {len(models)}\n"
                if vision_models:
                    msg += f"\nVision models available:\n• " + "\n• ".join(vision_models[:5])
                else:
                    msg += "\n⚠️ No vision models found.\nInstall one: ollama pull llava"
                
                QMessageBox.information(self, "Ollama Connection", msg)
            else:
                QMessageBox.warning(self, "Connection Failed", f"Ollama returned status {response.status_code}")
        except requests.exceptions.ConnectionError:
            QMessageBox.warning(
                self, 
                "Connection Failed", 
                f"Cannot connect to Ollama at {ollama_url}\n\n"
                "Make sure Ollama is running:\n"
                "1. Install: curl -fsSL https://ollama.com/install.sh | sh\n"
                "2. Start: ollama serve\n"
                "3. Pull vision model: ollama pull llava"
            )
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Connection error: {str(e)}")
    
    def populate_trf_patient_list(self):
        """Populate the TRF patient list table"""
        if not hasattr(self, 'bulk_patient_data_list') or not self.bulk_patient_data_list:
            self.trf_patient_list.setRowCount(0)
            return
        
        patients_dict = self.get_trf_patients_dict()
        self.trf_patient_list.setRowCount(len(patients_dict))
        
        for i, (patient_key, data) in enumerate(patients_dict.items()):
            p_info = data.get('patient_info', {})
            
            # Patient name
            name_item = QTableWidgetItem(p_info.get('patient_name', patient_key))
            self.trf_patient_list.setItem(i, 0, name_item)
            
            # PIN
            self.trf_patient_list.setItem(i, 1, QTableWidgetItem(p_info.get('pin', '')))
            
            # TRF Status
            trf_mapping = getattr(self, 'patient_trf_mapping', {})
            if patient_key in trf_mapping:
                status_item = QTableWidgetItem("✅ Linked")
                status_item.setForeground(QColor('#28a745'))
            else:
                status_item = QTableWidgetItem("⚪ No TRF")
                status_item.setForeground(QColor('#6c757d'))
            self.trf_patient_list.setItem(i, 2, status_item)
            
            # Store patient key and list index for later reference
            self.trf_patient_list.item(i, 0).setData(Qt.ItemDataRole.UserRole, patient_key)
            self.trf_patient_list.item(i, 0).setData(Qt.ItemDataRole.UserRole + 1, data.get('list_index', i))
    
    def on_trf_patient_selected(self):
        """Handle patient selection in TRF verification dialog"""
        selected = self.trf_patient_list.selectedItems()
        if not selected:
            return
        
        row = selected[0].row()
        patient_key = self.trf_patient_list.item(row, 0).data(Qt.ItemDataRole.UserRole)
        list_index = self.trf_patient_list.item(row, 0).data(Qt.ItemDataRole.UserRole + 1)
        
        patients_dict = self.get_trf_patients_dict()
        if not patient_key or patient_key not in patients_dict:
            return
        
        patient_data = patients_dict[patient_key]
        p_info = patient_data.get('patient_info', {})
        
        # Show patient details
        html = f"<h3>{p_info.get('patient_name', 'Unknown')}</h3>"
        html += f"<p><b>PIN:</b> {p_info.get('pin', 'N/A')}</p>"
        html += f"<p><b>Hospital:</b> {p_info.get('hospital_clinic', 'N/A')}</p>"
        html += f"<p><b>Biopsy Date:</b> {p_info.get('biopsy_date', 'N/A')}</p>"
        
        # Check if TRF is linked
        trf_mapping = getattr(self, 'patient_trf_mapping', {})
        if patient_key in trf_mapping:
            trf_info = trf_mapping[patient_key]
            
            # Check if it's from bulk PDF
            if trf_info.get('is_bulk_pdf'):
                page_num = trf_info.get('trf_page', 0) + 1
                html += f"<hr><p style='color:#28a745;'><b>✅ TRF Linked:</b> {os.path.basename(trf_info['trf_path'])} (Page {page_num})</p>"
            else:
                html += f"<hr><p style='color:#28a745;'><b>✅ TRF Linked:</b> {os.path.basename(trf_info['trf_path'])}</p>"
            
            html += f"<p><b>Match Score:</b> {int(trf_info.get('match_score', 0))}%</p>"
            
            if trf_info.get('trf_data'):
                html += "<p><b>Extracted Data:</b></p><ul>"
                for k, v in trf_info['trf_data'].items():
                    if v:
                        html += f"<li>{k}: {v}</li>"
                html += "</ul>"
        else:
            html += "<hr><p style='color:#6c757d;'><i>No TRF linked to this patient</i></p>"
            
            # Show option to assign page from bulk PDF if available
            if hasattr(self, 'bulk_trf_pdf_path') and self.bulk_trf_pdf_path:
                num_pages = getattr(self, 'bulk_trf_pdf_pages', 0)
                html += f"<p style='color:#007bff;'><i>💡 Bulk TRF loaded ({num_pages} pages). Use 'Assign Page' to link.</i></p>"
        
        self.trf_details_text.setHtml(html)
    
    def upload_trf_for_selected_patient(self):
        """Upload TRF for the selected patient - supports individual file or page from bulk PDF"""
        selected = self.trf_patient_list.selectedItems()
        if not selected:
            QMessageBox.warning(self, "No Selection", "Please select a patient first.")
            return
        
        row = selected[0].row()
        patient_key = self.trf_patient_list.item(row, 0).data(Qt.ItemDataRole.UserRole)
        
        # Check if bulk PDF is loaded - offer choice
        if hasattr(self, 'bulk_trf_pdf_path') and self.bulk_trf_pdf_path:
            msg = QMessageBox(self)
            msg.setWindowTitle("TRF Source")
            msg.setText(f"How would you like to assign TRF for {patient_key}?")
            
            page_btn = msg.addButton(f"📑 Select Page from Bulk PDF", QMessageBox.ButtonRole.ActionRole)
            file_btn = msg.addButton("📄 Upload Individual File", QMessageBox.ButtonRole.ActionRole)
            cancel_btn = msg.addButton(QMessageBox.StandardButton.Cancel)
            
            msg.exec()
            
            if msg.clickedButton() == page_btn:
                self.assign_bulk_pdf_page_to_patient(patient_key, row)
                return
            elif msg.clickedButton() == cancel_btn:
                return
        
        # Upload individual file
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            f"Select TRF for {patient_key}",
            "",
            "TRF Files (*.pdf *.png *.jpg *.jpeg *.tiff *.bmp)"
        )
        
        if file_path:
            if not hasattr(self, 'patient_trf_mapping'):
                self.patient_trf_mapping = {}
            
            self.patient_trf_mapping[patient_key] = {
                'trf_path': file_path,
                'trf_data': None,
                'match_score': 100,  # Manual upload = 100%
                'is_bulk_pdf': False
            }
            
            # Update list
            self.populate_trf_patient_list()
            
            # Re-select the row
            self.trf_patient_list.selectRow(row)
            
            self.statusBar().showMessage(f"TRF linked to {patient_key}")
    
    def assign_bulk_pdf_page_to_patient(self, patient_key, row):
        """Assign a specific page from bulk PDF to a patient"""
        if not hasattr(self, 'bulk_trf_pdf_path') or not self.bulk_trf_pdf_path:
            QMessageBox.warning(self, "Error", "No bulk PDF loaded")
            return
        
        num_pages = getattr(self, 'bulk_trf_pdf_pages', 0)
        
        # Create page selection dialog
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Assign Page to {patient_key}")
        dialog.setMinimumWidth(400)
        layout = QVBoxLayout()
        dialog.setLayout(layout)
        
        layout.addWidget(QLabel(f"<b>Select page from:</b> {os.path.basename(self.bulk_trf_pdf_path)}"))
        layout.addWidget(QLabel(f"Total pages: {num_pages}"))
        
        # Page selection
        page_layout = QHBoxLayout()
        page_layout.addWidget(QLabel("Page:"))
        page_spin = QSpinBox()
        page_spin.setMinimum(1)
        page_spin.setMaximum(num_pages)
        page_spin.setValue(1)
        page_layout.addWidget(page_spin)
        page_layout.addStretch()
        layout.addLayout(page_layout)
        
        # Preview area
        preview_label = QLabel("Page preview will appear here...")
        preview_label.setMinimumHeight(200)
        preview_label.setStyleSheet("border: 1px solid #ccc; padding: 10px;")
        preview_label.setWordWrap(True)
        layout.addWidget(preview_label)
        
        def update_preview():
            page_num = page_spin.value() - 1
            text, error = self.extract_text_from_bulk_pdf_page(page_num)
            if text:
                preview_label.setText(text[:500] + "..." if len(text) > 500 else text)
            else:
                preview_label.setText(f"Error: {error}" if error else "No text extracted")
        
        page_spin.valueChanged.connect(update_preview)
        update_preview()  # Initial preview
        
        # Buttons
        btn_layout = QHBoxLayout()
        assign_btn = QPushButton("✓ Assign Page")
        assign_btn.setStyleSheet("background-color: #28a745; color: white; padding: 8px 16px;")
        cancel_btn = QPushButton("Cancel")
        
        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(assign_btn)
        layout.addLayout(btn_layout)
        
        def do_assign():
            page_num = page_spin.value() - 1
            
            if not hasattr(self, 'patient_trf_mapping'):
                self.patient_trf_mapping = {}
            
            self.patient_trf_mapping[patient_key] = {
                'trf_path': self.bulk_trf_pdf_path,
                'trf_page': page_num,
                'trf_data': None,
                'match_score': 100,  # Manual assignment = 100%
                'is_bulk_pdf': True
            }
            
            dialog.accept()
            
            # Update list
            self.populate_trf_patient_list()
            self.trf_patient_list.selectRow(row)
            self.statusBar().showMessage(f"Page {page_num + 1} assigned to {patient_key}")
        
        assign_btn.clicked.connect(do_assign)
        cancel_btn.clicked.connect(dialog.reject)
        
        dialog.exec()
    
    def verify_selected_patient_trf(self):
        """Verify TRF for the selected patient"""
        selected = self.trf_patient_list.selectedItems()
        if not selected:
            QMessageBox.warning(self, "No Selection", "Please select a patient first.")
            return
        
        row = selected[0].row()
        patient_key = self.trf_patient_list.item(row, 0).data(Qt.ItemDataRole.UserRole)
        
        trf_mapping = getattr(self, 'patient_trf_mapping', {})
        if patient_key not in trf_mapping:
            QMessageBox.warning(self, "No TRF", "No TRF linked to this patient. Upload one first.")
            return
        
        trf_info = trf_mapping[patient_key]
        patients_dict = self.get_trf_patients_dict()
        if patient_key not in patients_dict:
            QMessageBox.warning(self, "Error", "Patient data not found.")
            return
        
        patient_data = patients_dict[patient_key]
        p_info = patient_data.get('patient_info', {})
        
        # Get patient data for comparison
        comparison_data = {
            'patient_name': p_info.get('patient_name', ''),
            'hospital_clinic': p_info.get('hospital_clinic', ''),
            'pin': p_info.get('pin', ''),
            'biopsy_date': p_info.get('biopsy_date', ''),
            'sample_receipt_date': p_info.get('sample_receipt_date', ''),
            'referring_clinician': p_info.get('referring_clinician', ''),
        }
        
        # Get extraction method from settings
        extraction_method = self.settings.value('trf_extraction_method', 'easyocr')
        use_ai = extraction_method == 'ollama'
        
        # Check if this is from bulk PDF
        if trf_info.get('is_bulk_pdf'):
            page_num = trf_info.get('trf_page', 0)
            
            # Extract text from specific page
            raw_text, error = self.extract_text_from_bulk_pdf_page(page_num)
            
            if error:
                QMessageBox.warning(self, "Extraction Error", f"Could not extract text from page {page_num + 1}: {error}")
                return
            
            # Parse extracted text
            trf_data = self.parse_extracted_trf_text(raw_text)
            
            # Compare and generate results
            results = self.compare_trf_to_patient(trf_data, comparison_data)
            all_correct = all(r['status'] in ['ok', 'skip'] for r in results)
            has_suggestions = any(r['can_apply'] for r in results)
            
        else:
            # Standard file verification
            results, all_correct, has_suggestions, error = self.verify_with_ai(
                trf_info['trf_path'], 
                comparison_data,
                use_ai=use_ai
            )
            
            if error:
                QMessageBox.warning(self, "Error", error)
                return
        
        # Show comparison dialog (modified for bulk context)
        self.show_bulk_trf_comparison_dialog(results, patient_key)
    
    def compare_trf_to_patient(self, trf_data, patient_data):
        """Compare TRF extracted data to patient data and return comparison results"""
        from difflib import SequenceMatcher
        
        def fuzzy_match(s1, s2):
            if not s1 or not s2:
                return 0.0
            return SequenceMatcher(None, s1.lower(), s2.lower()).ratio()
        
        results = []
        field_mapping = {
            'patient_name': 'Patient Name',
            'hospital_clinic': 'Hospital/Clinic',
            'pin': 'PIN',
            'biopsy_date': 'Biopsy Date',
            'sample_receipt_date': 'Sample Receipt Date',
            'referring_clinician': 'Referring Clinician',
        }
        
        for field_key, field_name in field_mapping.items():
            entered_value = patient_data.get(field_key, '') or ''
            trf_value = trf_data.get(field_key, '') or ''
            
            if not trf_value:
                trf_value = '(not found)'
            
            if not entered_value:
                status = 'suggestion' if trf_value != '(not found)' else 'skip'
                message = 'Value found in TRF' if trf_value != '(not found)' else 'No data'
                can_apply = trf_value != '(not found)'
            elif entered_value.lower() == trf_value.lower():
                status = 'ok'
                message = '✓ Match'
                can_apply = False
            elif trf_value == '(not found)':
                status = 'warning'
                message = 'Not found in TRF'
                can_apply = False
            else:
                score = fuzzy_match(entered_value, trf_value)
                if score >= 0.85:
                    status = 'ok'
                    message = f'✓ Similar ({int(score*100)}%)'
                    can_apply = False
                else:
                    status = 'mismatch'
                    message = f'⚠ Different ({int(score*100)}% match)'
                    can_apply = True
            
            results.append({
                'field': field_name,
                'field_key': field_key,
                'entered': entered_value,
                'trf_value': trf_value,
                'status': status,
                'message': message,
                'can_apply': can_apply
            })
        
        return results
    
    def get_trf_preview_image(self, patient_key):
        """Get TRF image for preview"""
        from PIL import Image
        import io
        
        trf_mapping = getattr(self, 'patient_trf_mapping', {})
        if patient_key not in trf_mapping:
            return None
        
        trf_info = trf_mapping[patient_key]
        trf_path = trf_info.get('trf_path')
        
        if not trf_path or not os.path.exists(trf_path):
            return None
        
        try:
            file_ext = os.path.splitext(trf_path)[1].lower()
            
            if file_ext == '.pdf':
                if not PDFPLUMBER_AVAILABLE:
                    return None
                
                with pdfplumber.open(trf_path) as pdf:
                    # Check if it's from bulk PDF (specific page)
                    if trf_info.get('is_bulk_pdf'):
                        page_idx = trf_info.get('trf_page', 0)
                        if page_idx < len(pdf.pages):
                            page = pdf.pages[page_idx]
                        else:
                            return None
                    else:
                        page = pdf.pages[0]
                    
                    # Convert to image
                    pil_img = page.to_image(resolution=150).original
                    return pil_img
            else:
                # Image file
                return Image.open(trf_path)
                
        except Exception as e:
            print(f"TRF preview error: {e}")
            return None
    
    def show_bulk_trf_comparison_dialog(self, results, patient_key):
        """Show enhanced TRF comparison dialog with 3-panel view:
        - Left: Entered values
        - Center: TRF extracted values + comparison
        - Right: TRF image preview
        """
        dialog = QDialog(self)
        dialog.setWindowTitle(f"📋 TRF Verification - {patient_key}")
        dialog.setMinimumWidth(1400)
        dialog.setMinimumHeight(700)
        
        main_layout = QVBoxLayout()
        dialog.setLayout(main_layout)
        
        # Get patient info for display
        patients_dict = self.get_trf_patients_dict()
        patient_data = patients_dict.get(patient_key, {})
        p_info = patient_data.get('patient_info', {})
        
        # Header with summary
        header_layout = QHBoxLayout()
        
        all_match = all(r['status'] in ['ok', 'skip'] for r in results)
        mismatch_count = sum(1 for r in results if r['status'] == 'mismatch')
        suggestion_count = sum(1 for r in results if r['can_apply'])
        
        if all_match:
            header = QLabel(f"✅ All fields verified for: {p_info.get('patient_name', patient_key)}")
            header.setStyleSheet("font-size: 16px; font-weight: bold; color: #28a745; padding: 10px;")
        else:
            header = QLabel(f"⚠️ Review differences for: {p_info.get('patient_name', patient_key)}")
            header.setStyleSheet("font-size: 16px; font-weight: bold; color: #dc3545; padding: 10px;")
        header_layout.addWidget(header)
        header_layout.addStretch()
        
        # Summary badges
        if mismatch_count > 0:
            mismatch_badge = QLabel(f"⚠️ {mismatch_count} Differences")
            mismatch_badge.setStyleSheet("background-color: #dc3545; color: white; padding: 5px 10px; border-radius: 10px;")
            header_layout.addWidget(mismatch_badge)
        if suggestion_count > 0:
            suggest_badge = QLabel(f"💡 {suggestion_count} Can Apply")
            suggest_badge.setStyleSheet("background-color: #007bff; color: white; padding: 5px 10px; border-radius: 10px;")
            header_layout.addWidget(suggest_badge)
        
        main_layout.addLayout(header_layout)
        
        # TRF info row
        trf_mapping = getattr(self, 'patient_trf_mapping', {})
        trf_info = trf_mapping.get(patient_key, {})
        trf_path = trf_info.get('trf_path', 'N/A')
        page_info = ""
        if trf_info.get('is_bulk_pdf'):
            page_info = f" (Page {trf_info.get('trf_page', 0) + 1})"
        
        info_label = QLabel(f"PIN: {p_info.get('pin', 'N/A')} | TRF: {os.path.basename(trf_path)}{page_info}")
        info_label.setStyleSheet("color: #666; padding: 0 10px 10px 10px;")
        main_layout.addWidget(info_label)
        
        # === 3-PANEL LAYOUT ===
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # ===== LEFT PANEL: Entered Values =====
        left_panel = QWidget()
        left_layout = QVBoxLayout()
        left_panel.setLayout(left_layout)
        
        left_header = QLabel("📝 Entered Values")
        left_header.setStyleSheet("font-size: 14px; font-weight: bold; background-color: #e3f2fd; padding: 8px; border-radius: 5px;")
        left_layout.addWidget(left_header)
        
        entered_table = QTableWidget()
        entered_table.setColumnCount(2)
        entered_table.setHorizontalHeaderLabels(["Field", "Value"])
        entered_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        entered_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        entered_table.setRowCount(len(results))
        entered_table.setAlternatingRowColors(True)
        entered_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        
        for i, r in enumerate(results):
            field_item = QTableWidgetItem(r['field'])
            field_item.setFont(QFont("", -1, QFont.Weight.Bold))
            entered_table.setItem(i, 0, field_item)
            
            value = r['entered'] or '(empty)'
            value_item = QTableWidgetItem(value)
            if r['status'] == 'mismatch':
                value_item.setBackground(QColor('#ffe6e6'))
            elif r['status'] == 'ok':
                value_item.setBackground(QColor('#d4edda'))
            entered_table.setItem(i, 1, value_item)
        
        entered_table.resizeRowsToContents()
        left_layout.addWidget(entered_table)
        splitter.addWidget(left_panel)
        
        # ===== CENTER PANEL: TRF Values + Actions =====
        center_panel = QWidget()
        center_layout = QVBoxLayout()
        center_panel.setLayout(center_layout)
        
        center_header = QLabel("📄 TRF Extracted Values")
        center_header.setStyleSheet("font-size: 14px; font-weight: bold; background-color: #e8f5e9; padding: 8px; border-radius: 5px;")
        center_layout.addWidget(center_header)
        
        trf_table = QTableWidget()
        trf_table.setColumnCount(3)
        trf_table.setHorizontalHeaderLabels(["Field", "TRF Value", "Action"])
        trf_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        trf_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        trf_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        trf_table.setRowCount(len(results))
        trf_table.setAlternatingRowColors(True)
        trf_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        
        for i, r in enumerate(results):
            field_item = QTableWidgetItem(r['field'])
            field_item.setFont(QFont("", -1, QFont.Weight.Bold))
            trf_table.setItem(i, 0, field_item)
            
            trf_value = r['trf_value'] or '(not found)'
            trf_item = QTableWidgetItem(trf_value)
            if r['status'] in ['mismatch', 'suggestion'] and trf_value != '(not found)':
                trf_item.setBackground(QColor('#d4edda'))
                trf_item.setForeground(QColor('#155724'))
            elif trf_value == '(not found)':
                trf_item.setForeground(QColor('#999'))
            trf_table.setItem(i, 1, trf_item)
            
            # Status + Action
            if r['can_apply'] and r['trf_value'] and r['trf_value'] != '(not found)':
                apply_btn = QPushButton("Apply →")
                apply_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #28a745; 
                        color: white; 
                        padding: 4px 10px;
                        border-radius: 3px;
                    }
                    QPushButton:hover {
                        background-color: #218838;
                    }
                """)
                apply_btn.clicked.connect(
                    lambda checked, res=r, pk=patient_key, et=entered_table, idx=i: 
                    self.apply_trf_value_with_preview(res, pk, et, idx)
                )
                trf_table.setCellWidget(i, 2, apply_btn)
            else:
                status_item = QTableWidgetItem("✓" if r['status'] == 'ok' else "—")
                if r['status'] == 'ok':
                    status_item.setForeground(QColor('#28a745'))
                trf_table.setItem(i, 2, status_item)
        
        trf_table.resizeRowsToContents()
        center_layout.addWidget(trf_table)
        
        # Apply All button in center panel
        suggestions = [r for r in results if r['can_apply'] and r['trf_value'] and r['trf_value'] != '(not found)']
        if suggestions:
            apply_all_btn = QPushButton(f"✓ Apply All ({len(suggestions)})")
            apply_all_btn.setStyleSheet("""
                QPushButton {
                    background-color: #007bff; 
                    color: white; 
                    padding: 10px;
                    font-weight: bold;
                    border-radius: 5px;
                }
                QPushButton:hover {
                    background-color: #0056b3;
                }
            """)
            apply_all_btn.clicked.connect(
                lambda: self.apply_all_trf_values_with_preview(results, patient_key, entered_table)
            )
            center_layout.addWidget(apply_all_btn)
        
        splitter.addWidget(center_panel)
        
        # ===== RIGHT PANEL: TRF Image Preview =====
        right_panel = QWidget()
        right_layout = QVBoxLayout()
        right_panel.setLayout(right_layout)
        
        right_header = QLabel("🖼️ TRF Preview")
        right_header.setStyleSheet("font-size: 14px; font-weight: bold; background-color: #fff3e0; padding: 8px; border-radius: 5px;")
        right_layout.addWidget(right_header)
        
        # Image preview area with scroll
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("background-color: #f5f5f5; border: 1px solid #ddd;")
        
        preview_label = QLabel()
        preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Load TRF image
        trf_image = self.get_trf_preview_image(patient_key)
        if trf_image:
            from PIL import Image
            import io
            
            # Scale image to fit preview (max width 500px)
            max_width = 500
            if trf_image.width > max_width:
                ratio = max_width / trf_image.width
                new_height = int(trf_image.height * ratio)
                trf_image = trf_image.resize((max_width, new_height), Image.Resampling.LANCZOS)
            
            # Convert PIL to QPixmap
            img_buffer = io.BytesIO()
            trf_image.save(img_buffer, format='PNG')
            img_buffer.seek(0)
            
            from PyQt6.QtGui import QPixmap
            pixmap = QPixmap()
            pixmap.loadFromData(img_buffer.getvalue())
            preview_label.setPixmap(pixmap)
        else:
            preview_label.setText("TRF preview not available")
            preview_label.setStyleSheet("color: #999; padding: 50px;")
        
        scroll_area.setWidget(preview_label)
        right_layout.addWidget(scroll_area)
        
        # Zoom controls
        zoom_layout = QHBoxLayout()
        zoom_in_btn = QPushButton("🔍+")
        zoom_out_btn = QPushButton("🔍-")
        zoom_fit_btn = QPushButton("Fit")
        
        for btn in [zoom_in_btn, zoom_out_btn, zoom_fit_btn]:
            btn.setStyleSheet("padding: 5px 15px;")
        
        self._preview_scale = 1.0
        self._preview_original_image = trf_image
        
        def zoom_preview(factor):
            if not self._preview_original_image:
                return
            self._preview_scale *= factor
            self._preview_scale = max(0.25, min(3.0, self._preview_scale))
            
            new_width = int(self._preview_original_image.width * self._preview_scale)
            new_height = int(self._preview_original_image.height * self._preview_scale)
            
            from PIL import Image
            import io
            scaled_img = self._preview_original_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            img_buffer = io.BytesIO()
            scaled_img.save(img_buffer, format='PNG')
            img_buffer.seek(0)
            
            from PyQt6.QtGui import QPixmap
            pixmap = QPixmap()
            pixmap.loadFromData(img_buffer.getvalue())
            preview_label.setPixmap(pixmap)
        
        def fit_preview():
            if not self._preview_original_image:
                return
            self._preview_scale = 500 / self._preview_original_image.width
            zoom_preview(1.0)
        
        zoom_in_btn.clicked.connect(lambda: zoom_preview(1.25))
        zoom_out_btn.clicked.connect(lambda: zoom_preview(0.8))
        zoom_fit_btn.clicked.connect(fit_preview)
        
        zoom_layout.addWidget(zoom_out_btn)
        zoom_layout.addWidget(zoom_fit_btn)
        zoom_layout.addWidget(zoom_in_btn)
        zoom_layout.addStretch()
        right_layout.addLayout(zoom_layout)
        
        splitter.addWidget(right_panel)
        
        # Set splitter sizes (30% / 30% / 40%)
        splitter.setSizes([350, 350, 500])
        
        main_layout.addWidget(splitter)
        
        # Bottom close button
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        close_btn = QPushButton("Close")
        close_btn.setStyleSheet("padding: 10px 30px;")
        close_btn.clicked.connect(dialog.accept)
        btn_layout.addWidget(close_btn)
        main_layout.addLayout(btn_layout)
        
        dialog.exec()
    
    def apply_trf_value_with_preview(self, result, patient_key, entered_table, row_idx):
        """Apply a TRF value and update the entered values table"""
        patients_dict = self.get_trf_patients_dict()
        if patient_key not in patients_dict:
            return
        
        list_index = patients_dict[patient_key].get('list_index')
        if list_index is None or list_index >= len(self.bulk_patient_data_list):
            return
        
        field_key = result['field_key']
        trf_value = result['trf_value']
        
        # Update the stored patient data
        self.bulk_patient_data_list[list_index]['patient_info'][field_key] = trf_value
        
        # Update the entered values table
        entered_table.item(row_idx, 1).setText(trf_value)
        entered_table.item(row_idx, 1).setBackground(QColor('#d4edda'))
        
        # Mark as applied
        result['can_apply'] = False
        
        self.statusBar().showMessage(f"Applied {result['field']}: {trf_value}")
    
    def apply_all_trf_values_with_preview(self, results, patient_key, entered_table):
        """Apply all TRF suggestions"""
        count = 0
        for i, r in enumerate(results):
            if r['can_apply'] and r['trf_value'] and r['trf_value'] != '(not found)':
                self.apply_trf_value_with_preview(r, patient_key, entered_table, i)
                count += 1
        
        QMessageBox.information(self, "Applied", f"Applied {count} values from TRF")
    
    def upload_bulk_trf_dialog(self):
        """Upload multiple TRFs from the dialog"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Multiple TRF Documents",
            "",
            "TRF Files (*.pdf *.png *.jpg *.jpeg *.tiff *.bmp)"
        )
        
        if file_paths:
            if not hasattr(self, 'pending_bulk_trfs'):
                self.pending_bulk_trfs = []
            self.pending_bulk_trfs.extend(file_paths)
            
            QMessageBox.information(
                self, 
                "TRFs Added",
                f"Added {len(file_paths)} TRF files.\nClick 'Auto-Match All TRFs' to match them with patients."
            )
    
    def auto_match_bulk_trfs(self):
        """Auto-match all pending TRFs with patients"""
        if not hasattr(self, 'pending_bulk_trfs') or not self.pending_bulk_trfs:
            QMessageBox.warning(self, "No TRFs", "No TRF files to match. Upload some first.")
            return
        
        if not hasattr(self, 'bulk_patient_data_list') or not self.bulk_patient_data_list:
            QMessageBox.warning(self, "No Patients", "No patient data loaded.")
            return
        
        # Get extraction method from settings
        extraction_method = self.settings.value('trf_extraction_method', 'easyocr')
        
        matched = 0
        unmatched = []
        
        for trf_path in self.pending_bulk_trfs:
            trf_name = os.path.basename(trf_path)
            
            # Extract data using selected method
            trf_data = None
            trf_text = None
            
            if extraction_method == 'ollama' and OLLAMA_AVAILABLE:
                trf_data, error = self.extract_text_with_ollama(trf_path)
                if error or not isinstance(trf_data, dict):
                    trf_text, _ = self.extract_text_enhanced(trf_path, method='easyocr')
                    trf_data = None
            else:
                trf_text, error = self.extract_text_enhanced(trf_path, method=extraction_method)
            
            if error and not trf_text and not trf_data:
                unmatched.append((trf_name, error))
                continue
            
            # Find best match
            best_match = None
            best_score = 0
            
            patients_dict = self.get_trf_patients_dict()
            for patient_key, patient_data in patients_dict.items():
                p_info = patient_data.get('patient_info', {})
                p_name = p_info.get('patient_name', '')
                p_pin = p_info.get('pin', '')
                
                score = 0
                
                if trf_data and isinstance(trf_data, dict):
                    trf_name_val = trf_data.get('patient_name', '') or ''
                    trf_pin_val = trf_data.get('pin', '') or ''
                    
                    if p_name and trf_name_val:
                        is_match, ratio = self.fuzzy_match(p_name, trf_name_val, 0.5)
                        score += ratio * 60
                    
                    if p_pin and trf_pin_val and p_pin.lower() in trf_pin_val.lower():
                        score += 40
                else:
                    if trf_text:
                        if p_name:
                            found, _ = self.find_in_text(trf_text, p_name)
                            if found:
                                score += 60
                        if p_pin:
                            found, _ = self.find_in_text(trf_text, p_pin)
                            if found:
                                score += 40
                
                if score > best_score:
                    best_score = score
                    best_match = patient_key
            
            if best_match and best_score >= 40:
                if not hasattr(self, 'patient_trf_mapping'):
                    self.patient_trf_mapping = {}
                
                self.patient_trf_mapping[best_match] = {
                    'trf_path': trf_path,
                    'trf_data': trf_data if isinstance(trf_data, dict) else None,
                    'match_score': best_score
                }
                matched += 1
            else:
                unmatched.append((trf_name, f"Best score: {int(best_score)}%"))
        
        # Clear pending
        self.pending_bulk_trfs = []
        
        # Update list
        self.populate_trf_patient_list()
        
        # Show results
        msg = f"Matched {matched} TRFs with patients."
        if unmatched:
            msg += f"\n\nUnmatched ({len(unmatched)}):\n"
            msg += "\n".join([f"• {name}: {reason}" for name, reason in unmatched[:5]])
            if len(unmatched) > 5:
                msg += f"\n... and {len(unmatched) - 5} more"
        
        QMessageBox.information(self, "Auto-Match Complete", msg)
    
    # ==================== End TRF Verification Methods ====================
    
    def clear_manual_form(self):
        """Clear all manual entry fields"""
        reply = QMessageBox.question(
            self, 'Confirm Clear',
            'Are you sure you want to clear all fields?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            # Clear patient fields
            for field in [self.patient_name_input, self.spouse_name_input, self.pin_input,
                         self.age_input, self.sample_number_input, self.referring_clinician_input,
                         self.biopsy_date_input, self.hospital_clinic_input, self.sample_collection_date_input,
                         self.sample_receipt_date_input, self.biopsy_performed_by_input]:
                field.clear()
            
            self.indication_input.clear()
            self.report_date_input.setText(datetime.now().strftime("%d-%m-%Y"))
            self.specimen_input.setText("Day 6 Trophectoderm Biopsy")
            
            # Reset embryo count
            self.embryo_count_spin.setValue(1)
    
    def browse_bulk_file(self):
        """Browse for bulk upload file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel or TSV File",
            self.settings.value('last_bulk_file_dir', ''),
            "Data Files (*.xlsx *.xls *.tsv *.csv)"
        )
        
        if file_path:
            self.bulk_file_label.setText(file_path)
            self.settings.setValue('last_bulk_file_dir', os.path.dirname(file_path))
            
            # Preview file
            try:
                if file_path.endswith(('.xlsx', '.xls')):
                    df = pd.read_excel(file_path)
                else:
                    df = pd.read_csv(file_path, sep='\t' if file_path.endswith('.tsv') else ',')
                
                # Display preview
                self.bulk_preview_table.setRowCount(min(10, len(df)))
                self.bulk_preview_table.setColumnCount(len(df.columns))
                self.bulk_preview_table.setHorizontalHeaderLabels(df.columns.tolist())
                
                for i in range(min(10, len(df))):
                    for j, col in enumerate(df.columns):
                        self.bulk_preview_table.setItem(i, j, QTableWidgetItem(str(df.iloc[i, j])))
                
                self.statusBar().showMessage(f"Loaded preview: {len(df)} rows")
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load file: {str(e)}")
    
    def download_template(self):
        """Download Excel template"""
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Template",
            "PGTA_Report_Template.xlsx",
            "Excel Files (*.xlsx)"
        )
        
        if save_path:
            # Create template DataFrame
            template_data = {
                'Patient_Name': ['Mrs. Example'],
                'Spouse_Name': ['Mr. Example'],
                'PIN': ['PIN12345'],
                'Age': ['30 Years'],
                'Sample_Number': ['123456'],
                'Referring_Clinician': ['Dr. Example'],
                'Biopsy_Date': ['01-01-2026'],
                'Hospital_Clinic': ['Example Hospital'],
                'Sample_Collection_Date': ['01-01-2026'],
                'Specimen': ['Day 6 Trophectoderm Biopsy'],
                'Sample_Receipt_Date': ['01-01-2026'],
                'Biopsy_Performed_By': ['Dr. Example'],
                'Report_Date': ['01-01-2026'],
                'Indication': ['Example indication'],
                'Embryo_ID': ['PS1'],
                'Result_Summary': ['Normal'],
                'Result_Description': ['The embryo contains normal chromosome complement'],
                'Autosomes': ['Normal'],
                'Sex_Chromosomes': ['Normal'],
                'Interpretation': ['Euploid'],
                'MTcopy': ['NA']
            }
            
            df = pd.DataFrame(template_data)
            df.to_excel(save_path, index=False)
            
            QMessageBox.information(self, "Success", f"Template saved to:\n{save_path}")
    
    def browse_bulk_output_folder(self):
        """Browse for bulk output directory"""
        dir_path = QFileDialog.getExistingDirectory(
            self,
            "Select Output Directory for Batch Reports",
            self.settings.value('last_bulk_output_dir', '')
        )
    
        if dir_path:
            self.bulk_output_label.setText(dir_path)
            self.settings.setValue('last_bulk_output_dir', dir_path)

    def browse_and_parse_bulk_file(self):
        """Browse for Excel file and automatically parse it"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Analysis RUN Excel File",
            self.settings.value('last_bulk_file_dir', ''),
            "Excel Files (*.xlsx *.xls)"
        )
        
        if not file_path:
            return
            
        self.bulk_file_label.setText(file_path)
        self.settings.setValue('last_bulk_file_dir', os.path.dirname(file_path))
        
        # Automatically parse the file
        self.parse_bulk_excel(file_path)

    def parse_bulk_excel(self, file_path):
        """Parse RUN Excel file and populate batch list"""
        try:
            xl = pd.ExcelFile(file_path)
            sheet_names_lower = [s.lower() for s in xl.sheet_names]
            
            # Find Details and summary sheets
            details_idx = next((i for i, s in enumerate(sheet_names_lower) if s == 'details'), None)
            summary_idx = next((i for i, s in enumerate(sheet_names_lower) if s == 'summary'), None)
            
            if details_idx is None or summary_idx is None:
                QMessageBox.warning(self, "Invalid Format", 
                    "Excel file must contain 'Details' and 'summary' sheets")
                return
            
            # Read sheets
            df_details = pd.read_excel(file_path, sheet_name=xl.sheet_names[details_idx])
            
            # Find the header row in summary sheet by searching for 'Sample name'
            df_summary_full = pd.read_excel(file_path, sheet_name=xl.sheet_names[summary_idx], header=None)
            header_row_idx = 0
            for r_idx, row in df_summary_full.iterrows():
                if any('sample name' in str(val).lower() for val in row.values):
                    header_row_idx = r_idx
                    break
            
            df_summary = pd.read_excel(file_path, sheet_name=xl.sheet_names[summary_idx], header=header_row_idx)
            
            # Clean columns
            df_details.columns = [str(c).strip() for c in df_details.columns]
            df_summary.columns = [str(c).strip() for c in df_summary.columns]
            
            # Parse and group data
            self.bulk_patient_data_list = []
            file_dir = os.path.dirname(file_path)
            
            for _, p_row in df_details.iterrows():
                p_name = str(p_row.get('Patient Name', '')).strip()
                if not p_name or p_name.lower() == 'nan':
                    continue
                
                # Extract patient info
                def get_clean_value(row, keys, default=''):
                    if isinstance(keys, str): keys = [keys]
                    for k in keys:
                        if k in row:
                            val = row[k]
                            if pd.isna(val): continue
                            s_val = str(val).strip(' \t\r\f\v') # Preserves newlines (\n)
                            if s_val.lower() in ['nan', 'none', 'nat', 'null']: continue
                            if s_val: return s_val
                    return default

                def format_date(d_val):
                    if not d_val: return ""
                    s = str(d_val).split(' ')[0] # Remove time if present
                    try:
                        # Try parsing common formats
                        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
                            try:
                                dt = datetime.strptime(s, fmt)
                                return dt.strftime("%d/%m/%Y")
                            except ValueError:
                                continue
                        return s # Return original if parse fails
                    except:
                        return s

                b_date_raw = get_clean_value(p_row, ['Date of Biopsy', 'Biopsy Date'])
                r_date_raw = get_clean_value(p_row, ['Date Sample Received', 'Receipt Date'])
                
                b_date = format_date(b_date_raw)
                r_date = format_date(r_date_raw)
                rep_date = datetime.now().strftime("%d/%m/%Y")
                
                patient_info = {
                    'patient_name': p_name,
                    'spouse_name': get_clean_value(p_row, ['Spouse Name', 'Husband Name', 'Partner Name']) or 'w/o',
                    'pin': get_clean_value(p_row, ['Sample ID', 'PIN', 'Patient ID']),
                    'age': get_clean_value(p_row, ['Age', 'Patient Age']),
                    'sample_number': '',  # Not extracted from Excel - user must fill manually
                    'referring_clinician': '',  # Not extracted from Excel - user must fill manually
                    'biopsy_date': b_date,
                    'hospital_clinic': get_clean_value(p_row, ['Center name', 'Hospital', 'Clinic', 'Center']),
                    'sample_collection_date': b_date,
                    'specimen': get_clean_value(p_row, ['Specimen Type', 'Sample Type'], 'Day 6 Trophectoderm Biopsy'),
                    'sample_receipt_date': r_date,
                    'biopsy_performed_by': get_clean_value(p_row, ['EMBRYOLOGIST NAME', 'Biologist']),
                    'report_date': rep_date,
                    'indication': get_clean_value(p_row, ['Indication', 'Clinical Indication']) # Removed 'Remarks' to avoid random data
                }
                
                # Find matching embryos
                # Robust Normalization: Remove all non-alphanumeric characters
                import re
                def normalize_str(s):
                    if not s: return ""
                    return re.sub(r'[^A-Z0-9]', '', str(s).upper())

                norm_p_name = normalize_str(p_name)
                norm_sample_id = normalize_str(p_row.get('Sample ID', ''))
                
                embryos = []
                
                for _, s_row in df_summary.iterrows():
                    sample_orig = str(s_row.get('Sample name', ''))
                    norm_s_name = normalize_str(sample_orig)
                    
                    # Enhanced Matching Logic: Check Sample ID first (more specific), then Patient Name
                    match = False
                    if norm_sample_id and norm_sample_id in norm_s_name:
                        match = True
                    elif norm_p_name and norm_p_name in norm_s_name:
                        match = True
                    
                    if match:
                        # Extract embryo ID
                        base_id = sample_orig.split('_')[0]
                        embryo_id = base_id.split('-')[-1] if '-' in base_id else base_id
                        
                        # Get values from summary sheet
                        conclusion = str(s_row.get('Conclusion', ''))
                        result_col = str(s_row.get('Result', ''))
                        qc_status = str(s_row.get('QC', '')).upper()
                        
                        # Determine Result Summary (for summary table) based on Conclusion
                        conclusion_upper = conclusion.upper()
                        result_summary_val = "Normal chromosome complement"  # Default
                        interp = "NA"  # Default interpretation
                        
                        # Handle QC failures first
                        if qc_status == 'FAIL' or 'INCONCLUSIVE' in result_col.upper() or 'RESEQUENCING' in result_col.upper():
                            result_summary_val = "Inconclusive"
                            interp = "NA"
                        elif 'LOW' in result_col.upper() and ('DNA' in result_col.upper() or 'READS' in result_col.upper()):
                            result_summary_val = "Low DNA concentration"
                            interp = "NA"
                        elif "CHAOTIC" in conclusion_upper or "CHAOTIC" in result_col.upper():
                            result_summary_val = "Multiple chromosomal abnormalities"
                            interp = "Chaotic embryo"
                        elif "NO COPY NUMBER ABNORMALITY" in conclusion_upper or conclusion_upper == 'EUPLOID' or result_col.upper() == 'EUPLOID':
                            result_summary_val = "Normal chromosome complement"
                            interp = "NA"
                        elif "MOSAIC" in conclusion_upper or "MOSAIC" in result_col.upper():
                            result_summary_val = "Mosaic chromosome complement"
                            # Determine mosaic level from percentage if available
                            import re as re_local
                            mos_match = re_local.search(r'(\d+)%', result_col)
                            if mos_match:
                                mos_pct = int(mos_match.group(1))
                                if mos_pct >= 50:
                                    interp = "High level mosaic"
                                else:
                                    interp = "Low level mosaic"
                            else:
                                interp = "Low level mosaic"
                        elif "ABNORMAL" in conclusion_upper or conclusion_upper == 'ABNORMAL':
                            # Check if multiple abnormalities
                            abnormal_count = result_col.count(',') + 1 if ',' in result_col else 1
                            if abnormal_count > 2:
                                result_summary_val = "Multiple chromosomal abnormalities"
                            else:
                                result_summary_val = "Multiple chromosomal abnormalities"  # Could also be single, but safer default
                            interp = "NA"
                        
                        # Advanced Parsing for Chromosome Statuses
                        # Example: del(5)(p15.33q12.3)(~64.50Mb,~57%)
                        res_sum = str(s_row.get('Result', ''))
                        chr_statuses = {str(i): 'N' for i in range(1, 23)}
                        mosaic_percentages = {}

                        def parse_complex_result(r_str):
                            c_map = {}
                            m_map = {}
                            if not r_str or r_str.lower() == 'nan': return c_map, m_map
                            
                            # Intelligent Split: Split by comma ONLY if followed by a keyword (del, dup, mos, +, -)
                            # This keeps "del(5)(... , ~57%)" together but splits "del(5)... , del(10)..."
                            parts = re.split(r',\s*(?=(?:del|dup|mos|[+-]|Monosomy|Trisomy))', r_str, flags=re.IGNORECASE)
                            
                            for part in parts:
                                part = part.strip()
                                if not part: continue
                                
                                # Regex for del/dup format: del(5)(...)
                                # Capture: Type(freq), Chr
                                match = re.search(r'(del|dup|mos)\D*?(\d+|X|Y)', part, re.IGNORECASE)
                                if match:
                                    etype = match.group(1).lower() # del/dup
                                    chrom = match.group(2)
                                    
                                    # Determine Status - Seg check looks for p/q band or 'Mb'
                                    is_seg = bool(re.search(r'([pq]\d|mb)', part, re.IGNORECASE))
                                    status = 'N'
                                    
                                    if 'del' in etype:
                                        status = 'SL' if is_seg else 'L'
                                    elif 'dup' in etype:
                                        status = 'SG' if is_seg else 'G'
                                    elif 'mos' in etype:
                                        status = 'M' # Generic Mosaic
                                        
                                    # Update Mosaic Percentage
                                    # Look for ~XX% or just XX%
                                    mos_match = re.search(r'[~]*(\d+)%', part)
                                    if mos_match:
                                        m_map[chrom] = mos_match.group(1)
                                        # Use mosaic specific codes if mosaic detected
                                        if 'del' in etype: status = 'SML' if is_seg else 'ML'
                                        if 'dup' in etype: status = 'SMG' if is_seg else 'MG'
                                    
                                    c_map[chrom] = status
                                    continue
                                
                                # Fallback/Alternative for simple format: Monosomy-5 or -5
                                # Check for Leading + or -
                                if part.startswith('+'):
                                     # Gain/Trisomy
                                     ch = part.replace('+', '').strip()
                                     if ch in chr_statuses: c_map[ch] = 'G'
                                elif part.startswith('-'):
                                     # Loss/Monosomy
                                     ch = part.replace('-', '').strip()
                                     if ch in chr_statuses: c_map[ch] = 'L'
                                     
                            
                            return c_map, m_map

                        p_stats, p_mos = parse_complex_result(res_sum)
                        chr_statuses.update(p_stats)
                        mosaic_percentages.update(p_mos)
                        
                        # Legacy Dash fallback if complex parse found nothing and dash exists
                        if not p_stats and '-' in res_sum and not 'del' in res_sum.lower():
                            parts = res_sum.split('-')
                            if len(parts) >= 2:
                                s_type = parts[0]
                                s_chrs = parts[1].split(',')
                                for ch in s_chrs:
                                    ch = ch.strip()
                                    if ch in chr_statuses:
                                        chr_statuses[ch] = s_type
                                        p_stats[ch] = s_type # Update local stats for generator

                        # --- Phase 4: Advanced Autosomes String Generator ---
                        def generate_autosomes_string(stats, mosaic_map):
                            if not stats: return "Normal" # or handle as Euploid logic below
                            
                            # Sort keys: 1-22, X, Y
                            def sort_key(k):
                                if k.isdigit(): return int(k)
                                if k.upper() == 'X': return 23
                                if k.upper() == 'Y': return 24
                                return 25
                            
                            sorted_chrs = sorted(stats.keys(), key=sort_key)
                            
                            # Logic: If SINGLE event -> Verbose. If MULTIPLE -> Concise.
                            is_multiple = len(sorted_chrs) > 1
                            parts = []
                            
                            for ch in sorted_chrs:
                                st = stats[ch]
                                mos_val = mosaic_map.get(ch, '')
                                
                                if is_multiple:
                                    # Format: "1 SG, 11 SG, 21 G"
                                    # If mosaic, maybe add %? User example "dup(1)..., +21: 1 SG, ... 21 G" (Red color)
                                    # User example also showed "dup(9)...: 9 chormosome..." when it was arguably single line in example?
                                    # Wait, the user example:
                                    # "dup(1)...,dup(11)...,dup(13)...,+21: 1 SG, 11 SG, 13 SG, 21 G" -> Concise for multiple
                                    parts.append(f"{ch} {st}")
                                else:
                                    # Format: "16 chromosome CNV status L"
                                    # Or "1 chromosome, CNV status SG"
                                    # Or "15 chromosome, CNV status MG, Mosaic(%) 30"
                                    base = f"{ch} chromosome, CNV status {st}"
                                    if mos_val:
                                        base += f", Mosaic(%) {mos_val}"
                                    parts.append(base)
                            
                            return ", ".join(parts)

                        # --- Phase 4: Autosomes - STRICTLY from AUTOSOMES column ---
                        # Get autosomes directly from the AUTOSOMES column in Excel
                        autosomes_raw = str(s_row.get('AUTOSOMES', '')).strip()
                        
                        # Handle nan/empty values
                        if not autosomes_raw or autosomes_raw.lower() in ['nan', 'none', 'nat', 'null', '']:
                            # Only if AUTOSOMES column is empty, derive from result summary
                            if result_summary_val == "Normal chromosome complement":
                                autosomes_val = "Euploid ( Normal )"
                            else:
                                autosomes_val = ""  # Leave empty if no data
                        elif autosomes_raw.lower() == 'normal':
                            autosomes_val = "Euploid ( Normal )"
                        else:
                            # Use the AUTOSOMES column value directly
                            autosomes_val = autosomes_raw

                        # --- Phase 4: Result Description Mapping ---
                        # Map result_summary to long description for the PDF body (Embryo Result Page)
                        long_desc = "The embryo contains normal chromosome complement" # Default
                        if result_summary_val == "Multiple chromosomal abnormalities":
                            long_desc = "The embryo contains abnormal chromosome complement"
                        elif result_summary_val == "Mosaic chromosome complement":
                            long_desc = "The embryo contains mosaic chromosome complement"
                        elif result_summary_val == "Inconclusive" or result_summary_val == "Low DNA concentration":
                            long_desc = "Inconclusive"
                        elif result_summary_val == "Normal chromosome complement":
                            long_desc = "The embryo contains normal chromosome complement"
                        
                        # --- Phase 4: Sex Chromosomes - STRICTLY from SEX column ---
                        sex_raw = str(s_row.get('SEX', '')).strip()
                        
                        # Handle nan/empty values
                        if not sex_raw or sex_raw.lower() in ['nan', 'none', 'nat', 'null', '']:
                            sex_chr_val = "Normal"  # Default to Normal if no data
                        elif sex_raw.lower() == 'normal':
                            sex_chr_val = "Normal"
                        else:
                            # Any other value (e.g., "MOSAIC GAIN (52%)", abnormality descriptions)
                            sex_chr_val = "Abnormal"

                        # Auto-match CNV image (Restored)
                        cnv_image_path = None
                        sample_base = sample_orig.split('_')[0]
                        if os.path.exists(file_dir):
                            for f in os.listdir(file_dir):
                                if f.upper().startswith(sample_base.upper()) and f.lower().endswith(('.png', '.jpg')):
                                    cnv_image_path = os.path.join(file_dir, f)
                                    break
                        
                        embryos.append({
                            'embryo_id': embryo_id,
                            'result_summary': result_summary_val,  # Use the mapped result summary
                            'interpretation': interp,
                            'result_description': long_desc, # Mapped long text
                            'autosomes': autosomes_val,
                            'sex_chromosomes': sex_chr_val,
                            'mtcopy': str(s_row.get('MTcopy', 'NA')),
                            'cnv_image_path': cnv_image_path,
                            'chromosome_statuses': chr_statuses,
                            'mosaic_percentages': mosaic_percentages # Pass properly
                        })
                
                if embryos:
                    self.bulk_patient_data_list.append({
                        'patient_info': patient_info,
                        'embryos': embryos
                    })
            
            # Populate batch list
            self.batch_list_widget.clear()
            for i, data in enumerate(self.bulk_patient_data_list):
                p_name = data['patient_info']['patient_name']
                e_count = len(data['embryos'])
                item = QListWidgetItem(f"{p_name} ({e_count} embryos)")
                item.setData(Qt.ItemDataRole.UserRole, i)
                self.batch_list_widget.addItem(item)
            
            self.statusBar().showMessage(f"Loaded {len(self.bulk_patient_data_list)} patients")
            self.update_data_summary()
            QMessageBox.information(self, "Success", 
                f"Successfully parsed {len(self.bulk_patient_data_list)} patients with "
                f"{sum(len(d['embryos']) for d in self.bulk_patient_data_list)} total embryos")
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Parse Error", f"Failed to parse Excel file:\n{str(e)}")

    def filter_batch_list(self, search_text):
        """Filter the patient list based on search text"""
        search_text = search_text.lower().strip()
        
        for i in range(self.batch_list_widget.count()):
            item = self.batch_list_widget.item(i)
            if item:
                # Show all if search is empty, otherwise filter by name
                if not search_text:
                    item.setHidden(False)
                else:
                    item_text = item.text().lower()
                    item.setHidden(search_text not in item_text)

    def on_batch_selection_changed(self, current, previous):
        """Handle batch list selection change - populate comprehensive batch editor"""
        if not current:
            return
    
        idx = current.data(Qt.ItemDataRole.UserRole)
        if idx >= len(self.bulk_patient_data_list):
            return
    
        # Clear existing editor
        while self.batch_editor_layout.count():
            child = self.batch_editor_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
    
        # Get patient data
        data = self.bulk_patient_data_list[idx]
        self.current_batch_index = idx
    
        # Patient Info Section - ALL FIELDS like manual entry
        patient_group = QGroupBox("Patient Information")
        patient_form = QFormLayout()
        patient_group.setLayout(patient_form)
    
        self.batch_patient_name = QTextEdit(data['patient_info']['patient_name'])
        self.batch_patient_name.setMaximumHeight(40)
        spouse_val = data['patient_info'].get('spouse_name', '') or 'w/o'
        self.batch_spouse_name = QTextEdit(spouse_val)
        self.batch_spouse_name.setMaximumHeight(40)
        self.batch_spouse_name.setPlaceholderText("w/o")
        self.batch_pin = QTextEdit(data['patient_info']['pin'])
        self.batch_pin.setMaximumHeight(40)
        self.batch_age = QTextEdit(data['patient_info'].get('age', ''))
        self.batch_age.setMaximumHeight(40)
        self.batch_sample_number = QTextEdit(data['patient_info']['sample_number'])
        self.batch_sample_number.setMaximumHeight(40)
        self.batch_referring_clinician = QTextEdit(data['patient_info']['referring_clinician'])
        self.batch_referring_clinician.setMaximumHeight(40)
        self.batch_biopsy_date = QLineEdit(data['patient_info']['biopsy_date'])
        self.batch_hospital = QTextEdit(data['patient_info']['hospital_clinic'])
        self.batch_hospital.setMaximumHeight(40)
        self.batch_sample_collection_date = QLineEdit(data['patient_info'].get('sample_collection_date', ''))
        self.batch_specimen = QTextEdit(data['patient_info'].get('specimen', 'Day 6 Trophectoderm Biopsy'))
        self.batch_specimen.setMaximumHeight(40)
        self.batch_sample_receipt_date = QLineEdit(data['patient_info'].get('sample_receipt_date', ''))
        self.batch_biopsy_performed_by = QTextEdit(data['patient_info'].get('biopsy_performed_by', ''))
        self.batch_biopsy_performed_by.setMaximumHeight(40)
        self.batch_report_date = QLineEdit(data['patient_info'].get('report_date', datetime.now().strftime("%d-%m-%Y")))
        self.batch_indication = QTextEdit(data['patient_info'].get('indication', ''))
        self.batch_indication.setMaximumHeight(80)
    
        # Connect Patient Fields to Live Preview
        for field in [self.batch_patient_name, self.batch_spouse_name, self.batch_pin, self.batch_age,
                      self.batch_sample_number, self.batch_referring_clinician, self.batch_biopsy_date,
                      self.batch_hospital, self.batch_sample_collection_date, self.batch_specimen,
                      self.batch_sample_receipt_date, self.batch_biopsy_performed_by, self.batch_report_date]:
            field.textChanged.connect(self.update_batch_preview)
        self.batch_indication.textChanged.connect(self.update_batch_preview)
    
        patient_form.addRow("Patient Name:", self.batch_patient_name)
        patient_form.addRow("Spouse Name:", self.batch_spouse_name)
        patient_form.addRow("PIN:", self.batch_pin)
        patient_form.addRow("Age:", self.batch_age)
        patient_form.addRow("Sample Number:", self.batch_sample_number)
        patient_form.addRow("Referring Clinician:", self.batch_referring_clinician)
        patient_form.addRow("Biopsy Date:", self.batch_biopsy_date)
        patient_form.addRow("Hospital/Clinic:", self.batch_hospital)
        patient_form.addRow("Sample Collection Date:", self.batch_sample_collection_date)
        patient_form.addRow("Specimen:", self.batch_specimen)
        patient_form.addRow("Sample Receipt Date:", self.batch_sample_receipt_date)
        patient_form.addRow("Biopsy Performed By:", self.batch_biopsy_performed_by)
        patient_form.addRow("Report Date:", self.batch_report_date)
        patient_form.addRow("Indication:", self.batch_indication)
        
        # --- TRF Verification Section for Batch ---
        trf_group = QGroupBox("TRF Verification")
        trf_layout = QVBoxLayout()
        trf_group.setLayout(trf_layout)
        
        trf_upload_layout = QHBoxLayout()
        self.batch_trf_path_label = QLabel("No TRF uploaded")
        self.batch_trf_path_label.setStyleSheet("color: #666; font-style: italic;")
        trf_upload_layout.addWidget(self.batch_trf_path_label, 1)
        
        batch_trf_upload_btn = QPushButton("📄 Upload TRF")
        batch_trf_upload_btn.clicked.connect(self.upload_trf_batch)
        trf_upload_layout.addWidget(batch_trf_upload_btn)
        
        batch_trf_verify_btn = QPushButton("✓ Verify")
        batch_trf_verify_btn.clicked.connect(self.verify_trf_batch)
        trf_upload_layout.addWidget(batch_trf_verify_btn)
        
        trf_layout.addLayout(trf_upload_layout)
        
        # Verification result display for batch
        self.batch_trf_result_text = QTextBrowser()
        self.batch_trf_result_text.setMaximumHeight(120)
        self.batch_trf_result_text.setStyleSheet("background-color: #f8f9fa; border: 1px solid #ddd; border-radius: 4px;")
        self.batch_trf_result_text.setHtml("<i style='color:#888;'>Upload a TRF to verify patient details</i>")
        trf_layout.addWidget(self.batch_trf_result_text)
        
        patient_form.addRow("", trf_group)
        
        # Store batch TRF path
        self.batch_trf_path = None
    
        self.batch_editor_layout.addWidget(patient_group)
    
        # Embryos Section with ALL fields
        embryos_group = QGroupBox(f"Embryos ({len(data['embryos'])})")
        embryos_layout = QVBoxLayout()
        embryos_group.setLayout(embryos_layout)
    
        self.batch_embryo_editors = []
        for i, embryo in enumerate(data['embryos']):
            embryo_frame = QGroupBox(f"Embryo: {embryo['embryo_id']}")
            embryo_form = QFormLayout()
            embryo_frame.setLayout(embryo_form)
        
            e_id = QLineEdit(embryo['embryo_id'])
            
            # Result Summary dropdown with colors
            e_result_summary = ClickOnlyComboBox()
            add_colored_items_to_combo(e_result_summary, [
                ("Normal chromosome complement", "black"),
                ("Multiple chromosomal abnormalities", "red"), 
                ("Mosaic chromosome complement", "blue"),
                ("Inconclusive", "black"),
                ("Low DNA concentration", "black")
            ])
            e_result_summary.setEditable(True)
            e_result_summary.setInsertPolicy(ClickOnlyComboBox.InsertPolicy.NoInsert)
            e_result_summary.setCurrentText(embryo['result_summary'])
            
            # Result Description dropdown - All black as per spec
            e_result_desc = ClickOnlyComboBox()
            add_colored_items_to_combo(e_result_desc, [
                ("The embryo contains normal chromosome complement", "black"),
                ("The embryo contains abnormal chromosome complement", "black"),
                ("The embryo contains mosaic chromosome complement", "black"),
                ("Inconclusive", "black")
            ])
            e_result_desc.setEditable(True)
            e_result_desc.setInsertPolicy(ClickOnlyComboBox.InsertPolicy.NoInsert)
            e_result_desc.setCurrentText(embryo.get('result_description', 'The embryo contains normal chromosome complement'))
            
            e_autosomes = QLineEdit(embryo.get('autosomes', ''))
            
            # Sex Chromosomes dropdown - Normal=Black, Abnormal=Red
            e_sex_chr = ClickOnlyComboBox()
            add_colored_items_to_combo(e_sex_chr, [
                ("Normal", "black"),
                ("Abnormal", "red")
            ])
            e_sex_chr.setEditable(True)
            e_sex_chr.setInsertPolicy(ClickOnlyComboBox.InsertPolicy.NoInsert)
            e_sex_chr.setCurrentText(embryo.get('sex_chromosomes', 'Normal'))
            
            # Interpretation dropdown with colors
            e_interp = ClickOnlyComboBox()
            add_colored_items_to_combo(e_interp, [
                ("NA", "black"),
                ("Chaotic embryo", "red"),
                ("Low level mosaic", "blue"),
                ("High level mosaic", "blue"),
                ("Complex mosaic", "blue")
            ])
            e_interp.setEditable(True)
            e_interp.setInsertPolicy(ClickOnlyComboBox.InsertPolicy.NoInsert)
            e_interp.setCurrentText(embryo['interpretation'])
            e_mtcopy = QLineEdit(embryo['mtcopy'])
        
            # Connect Embryo Fields to Live Preview
            e_id.textChanged.connect(self.update_batch_preview)
            e_result_summary.currentTextChanged.connect(self.update_batch_preview)
            e_result_desc.currentTextChanged.connect(self.update_batch_preview)
            e_autosomes.textChanged.connect(self.update_batch_preview)
            e_sex_chr.currentTextChanged.connect(self.update_batch_preview)
            e_interp.currentTextChanged.connect(self.update_batch_preview)
            e_mtcopy.textChanged.connect(self.update_batch_preview)
        
            # CNV Image with upload button
            image_layout = QHBoxLayout()
            e_image_label = QLabel(os.path.basename(embryo['cnv_image_path']) if embryo['cnv_image_path'] else "No image")
            e_image_path = embryo['cnv_image_path']
            image_layout.addWidget(e_image_label, 1)
        
            upload_img_btn = QPushButton("Upload Image")
            upload_img_btn.clicked.connect(lambda checked, idx=i: self.upload_embryo_image_batch(idx))
            image_layout.addWidget(upload_img_btn)
        
            embryo_form.addRow("Embryo ID:", e_id)
            embryo_form.addRow("Result Summary:", e_result_summary)
            embryo_form.addRow("Result Description:", e_result_desc)
            embryo_form.addRow("Autosomes:", e_autosomes)
            embryo_form.addRow("Sex Chromosomes:", e_sex_chr)
            embryo_form.addRow("Interpretation:", e_interp)
            embryo_form.addRow("MTcopy:", e_mtcopy)
        
            image_widget = QWidget()
            image_widget.setLayout(image_layout)
            embryo_form.addRow("CNV Image:", image_widget)

            # Chromosome status section using Grid (Same as manual form)
            chr_group = QGroupBox("Chromosome Details")
            chr_grid = QGridLayout()
            chr_group.setLayout(chr_grid)
            embryo_form.addRow(chr_group)
        
            # Headers
            chr_grid.addWidget(QLabel("<b>Chr</b>"), 0, 0)
            chr_grid.addWidget(QLabel("<b>Status</b>"), 0, 1)
            chr_grid.addWidget(QLabel("<b>Mosaic %</b>"), 0, 2)
            chr_grid.addWidget(QLabel("   "), 0, 3) # Spacer
            chr_grid.addWidget(QLabel("<b>Chr</b>"), 0, 4)
            chr_grid.addWidget(QLabel("<b>Status</b>"), 0, 5)
            chr_grid.addWidget(QLabel("<b>Mosaic %</b>"), 0, 6)
        
            chr_inputs = {}
            chr_statuses = embryo.get('chromosome_statuses', {})
            mosaic_percentages = embryo.get('mosaic_percentages', {})
        
            for j in range(1, 23):
                # Determine column (Left: 1-11, Right: 12-22)
                if j <= 11:
                    row = j
                    col_base = 0
                else:
                    row = j - 11
                    col_base = 4
                
                s_j = str(j)
                
                # Label
                chr_grid.addWidget(QLabel(s_j), row, col_base)
                
                # Status Combo
                chr_combo = ClickOnlyComboBox()
                chr_combo.setEditable(True) # Manual Entry Enabled
                chr_combo.addItems(["N", "G", "L", "SG", "SL", "M", "MG", "ML", "SMG", "SML", "NA"])
                chr_combo.setCurrentText(chr_statuses.get(s_j, 'N'))
                chr_combo.currentTextChanged.connect(self.update_batch_preview)
                chr_grid.addWidget(chr_combo, row, col_base + 1)
                
                # Mosaic Input
                mos_input = QLineEdit()
                mos_input.setPlaceholderText("%")
                mos_input.setText(str(mosaic_percentages.get(s_j, '')))
                mos_input.setMaximumWidth(60)
                mos_input.textChanged.connect(self.update_batch_preview)
                chr_grid.addWidget(mos_input, row, col_base + 2)
                
                chr_inputs[s_j] = {'status': chr_combo, 'mosaic': mos_input}
        
            embryos_layout.addWidget(embryo_frame)
        
            self.batch_embryo_editors.append({
                'embryo_id': e_id,
                'result_summary': e_result_summary,
                'result_description': e_result_desc,
                'autosomes': e_autosomes,
                'sex_chromosomes': e_sex_chr,
                'interpretation': e_interp,
                'mtcopy': e_mtcopy,
                'image_label': e_image_label,
                'image_path': e_image_path,
                'chr_inputs': chr_inputs
            })
    
        self.batch_editor_layout.addWidget(embryos_group)
    
        # Action buttons
        button_layout = QHBoxLayout()
    
        save_btn = QPushButton("Save Changes")
        save_btn.clicked.connect(self.save_batch_edits)
        button_layout.addWidget(save_btn)
    
        save_individual_draft_btn = QPushButton("Save This Patient as Draft")
        save_individual_draft_btn.clicked.connect(self.save_individual_patient_draft)
        button_layout.addWidget(save_individual_draft_btn)
    
        preview_btn = QPushButton("Preview PDF")
        preview_btn.clicked.connect(self.preview_batch_patient_pdf)
        button_layout.addWidget(preview_btn)
    
        generate_one_btn = QPushButton("Generate This Report")
        generate_one_btn.clicked.connect(self.generate_single_batch_report)
        button_layout.addWidget(generate_one_btn)
    
        self.batch_editor_layout.addWidget(QWidget())  # Spacer widget
        button_widget = QWidget()
        button_widget.setLayout(button_layout)
        self.batch_editor_layout.addWidget(button_widget)
    
        self.batch_editor_layout.addStretch()
        
        # Trigger initial preview
        self.update_batch_preview()

    def save_batch_edits(self):
        """Save edits from batch editor back to data list with 'nan' sanitation"""
        if not hasattr(self, 'current_batch_index'):
            return
            
        def clean(val):
            if val is None: return ""
            s = str(val).strip(' \t\r\f\v') # Preserves newlines (\n)
            return "" if s.lower() == "nan" else s

        idx = self.current_batch_index
        data = self.bulk_patient_data_list[idx]
    
        # Update ALL patient info fields
        data['patient_info']['patient_name'] = clean(self.batch_patient_name.toPlainText())
        data['patient_info']['spouse_name'] = clean(self.batch_spouse_name.toPlainText())
        data['patient_info']['pin'] = clean(self.batch_pin.toPlainText())
        data['patient_info']['age'] = clean(self.batch_age.toPlainText())
        data['patient_info']['sample_number'] = clean(self.batch_sample_number.toPlainText())
        data['patient_info']['referring_clinician'] = clean(self.batch_referring_clinician.toPlainText())
        data['patient_info']['biopsy_date'] = clean(self.batch_biopsy_date.text())
        data['patient_info']['hospital_clinic'] = clean(self.batch_hospital.toPlainText())
        data['patient_info']['sample_collection_date'] = clean(self.batch_sample_collection_date.text())
        data['patient_info']['specimen'] = clean(self.batch_specimen.toPlainText())
        data['patient_info']['sample_receipt_date'] = clean(self.batch_sample_receipt_date.text())
        data['patient_info']['biopsy_performed_by'] = clean(self.batch_biopsy_performed_by.toPlainText())
        data['patient_info']['report_date'] = clean(self.batch_report_date.text())
        data['patient_info']['indication'] = clean(self.batch_indication.toPlainText())
    
        # Update ALL embryo fields
        for i, editor in enumerate(self.batch_embryo_editors):
            if i < len(data['embryos']):
                data['embryos'][i]['embryo_id'] = editor['embryo_id'].text()
                # result_summary, result_description, sex_chromosomes are now combo boxes
                data['embryos'][i]['result_summary'] = editor['result_summary'].currentText()
                data['embryos'][i]['result_description'] = editor['result_description'].currentText()
                data['embryos'][i]['autosomes'] = editor['autosomes'].text()
                data['embryos'][i]['sex_chromosomes'] = editor['sex_chromosomes'].currentText()
                data['embryos'][i]['interpretation'] = editor['interpretation'].currentText()
                data['embryos'][i]['mtcopy'] = editor['mtcopy'].text()
                data['embryos'][i]['cnv_image_path'] = editor['image_path']
                
                # Update chromosome data
                if 'chr_inputs' in editor:
                    for ch, inputs in editor['chr_inputs'].items():
                        data['embryos'][i]['chromosome_statuses'][ch] = inputs['status'].currentText()
                        data['embryos'][i]['mosaic_percentages'][ch] = inputs['mosaic'].text()
    
        self.statusBar().showMessage("Changes saved to batch")
        QMessageBox.information(self, "Saved", "Changes saved successfully!")

    def upload_embryo_image_batch(self, embryo_idx):
        """Upload CNV image for specific embryo in batch"""
        if not hasattr(self, 'current_batch_index'):
            return
    
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select CNV Chart Image",
            "",
            "Image Files (*.png *.jpg *.jpeg)"
        )
    
        if file_path:
            # Update the data
            batch_idx = self.current_batch_index
            self.bulk_patient_data_list[batch_idx]['embryos'][embryo_idx]['cnv_image_path'] = file_path
        
            # Update the UI
            if embryo_idx < len(self.batch_embryo_editors):
                self.batch_embryo_editors[embryo_idx]['image_label'].setText(os.path.basename(file_path))
                self.batch_embryo_editors[embryo_idx]['image_path'] = file_path
        
            self.statusBar().showMessage(f"Image uploaded for embryo {embryo_idx + 1}")

    def save_individual_patient_draft(self):
        """Save individual patient as JSON draft"""
        if not hasattr(self, 'current_batch_index'):
            return
            
        idx = self.current_batch_index
        # Fix for KeyError: check if index is valid before access
        if idx < 0 or idx >= len(self.bulk_patient_data_list):
            return

        data = self.bulk_patient_data_list[idx]
    
        p_name = data['patient_info']['patient_name'].replace(' ', '_')
        default_name = f"{p_name}_draft.json"
    
        path, _ = QFileDialog.getSaveFileName(self, "Save Patient Draft", default_name, "JSON Files (*.json)")
        if path:
            try:
                with open(path, 'w') as f:
                    json.dump(data, f, indent=4)
                QMessageBox.information(self, "Success", f"Patient draft saved to {os.path.basename(path)}")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def update_batch_preview(self):
        """Update live preview for batch editor"""
        self.schedule_batch_preview_update()

    def schedule_batch_preview_update(self):
        """Debounce batch preview updates"""
        if not hasattr(self, 'batch_preview_timer'):
            self.batch_preview_timer = QTimer()
            self.batch_preview_timer.setSingleShot(True)
            self.batch_preview_timer.setInterval(1000) # 1 second debounce
            self.batch_preview_timer.timeout.connect(self.start_batch_preview_generation)
            
        self.batch_preview_timer.start()

    def start_batch_preview_generation(self):
        """Generate temp PDF for batch patient and show in preview"""
        if not hasattr(self, 'batch_pdf_view') or isinstance(self.batch_pdf_view, QLabel):
            return
            
        if not hasattr(self, 'current_batch_index'):
            return

        # Gather data from batch editor fields (Live data, not stored data)
        p_data = {
            'patient_name': self.batch_patient_name.toPlainText(),
            'spouse_name': self.batch_spouse_name.toPlainText(),
            'pin': self.batch_pin.toPlainText(),
            'age': self.batch_age.toPlainText(),
            'sample_number': self.batch_sample_number.toPlainText(),
            'referring_clinician': self.batch_referring_clinician.toPlainText(),
            'biopsy_date': self.batch_biopsy_date.text(),
            'hospital_clinic': self.batch_hospital.toPlainText(),
            'sample_collection_date': self.batch_sample_collection_date.text(),
            'specimen': self.batch_specimen.toPlainText(),
            'sample_receipt_date': self.batch_sample_receipt_date.text(),
            'biopsy_performed_by': self.batch_biopsy_performed_by.toPlainText(),
            'report_date': self.batch_report_date.text(),
            'indication': self.batch_indication.toPlainText()
        }
        
        e_data = []
        for editor in self.batch_embryo_editors:
            e_data.append({
                'embryo_id': editor['embryo_id'].text(),
                # result_summary, result_description, sex_chromosomes are now combo boxes
                'result_summary': editor['result_summary'].currentText(),
                'interpretation': editor['interpretation'].currentText(),
                'result_description': editor['result_description'].currentText(),
                'autosomes': editor['autosomes'].text(),
                'sex_chromosomes': editor['sex_chromosomes'].currentText(),
                'mtcopy': editor['mtcopy'].text(),
                'cnv_image_path': editor['image_path'],
                'chromosome_statuses': {ch: inp['status'].currentText() for ch, inp in editor.get('chr_inputs', {}).items()},
                'mosaic_percentages': {ch: inp['mosaic'].text() for ch, inp in editor.get('chr_inputs', {}).items()}
            })

        import tempfile
        temp_pdf = os.path.join(tempfile.gettempdir(), f"batch_preview_{self.current_batch_index}.pdf")
        
        # Run in worker
        if hasattr(self, 'batch_preview_worker') and self.batch_preview_worker.isRunning():
            return
            
        show_logo = self.bulk_logo_combo.currentText() == "With Logo"
        self.batch_preview_worker = PreviewWorker(p_data, e_data, temp_pdf, show_logo=show_logo)
        self.batch_preview_worker.finished.connect(lambda path: self.on_batch_preview_generated(path))
        self.batch_preview_worker.start()

    def on_batch_preview_generated(self, pdf_path):
        """Load generated batch PDF into viewer"""
        if QPdfDocument and hasattr(self, 'batch_pdf_document') and os.path.exists(pdf_path):
            self.batch_pdf_document.load(pdf_path)
            self.batch_pdf_view.setZoomMode(QPdfView.ZoomMode.FitInView)

    def preview_batch_patient_pdf(self):
        """Preview PDF for current batch patient"""
        if not hasattr(self, 'current_batch_index'):
            return
    
        # Save current edits first - wrapped in try cache to prevent crash on partial data
        try:
            self.save_batch_edits()
        except Exception as e:
            QMessageBox.warning(self, "Save Error", f"Could not save current edits before preview: {e}")
            return
    
        idx = self.current_batch_index
        data = self.bulk_patient_data_list[idx]
    
        # Generate temp PDF
        import tempfile
        import shutil
        temp_dir = tempfile.mkdtemp()
        temp_pdf = os.path.join(temp_dir, "preview.pdf")
    
        try:
            template = PGTAReportTemplate(assets_dir="assets/pgta")
            # Check for show_logo preference, default to True if combo not found
            show_logo = True
            if hasattr(self, 'bulk_logo_combo'):
                show_logo = self.bulk_logo_combo.currentText() == "With Logo"
                
            template.generate_pdf(temp_pdf, data['patient_info'], data['embryos'], show_logo=show_logo)
        
            # Show in PDF viewer dialog
            self.show_pdf_preview(temp_pdf, data['patient_info']['patient_name'])
        except Exception as e:
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Preview Error", f"Failed to generate preview:\\n{str(e)}")

    def show_pdf_preview(self, pdf_path, patient_name):
        """Show PDF preview dialog"""
        dialog = QDialog(self)
        dialog.setWindowTitle(f"PDF Preview - {patient_name}")
        dialog.resize(800, 1000)
    
        layout = QVBoxLayout()
        dialog.setLayout(layout)
    
        if QPdfDocument and QPdfView:
            # Use Qt PDF viewer with parent passed to constructor
            pdf_doc = QPdfDocument(dialog)
            pdf_view = QPdfView(dialog)
            pdf_doc.load(pdf_path)
            pdf_view.setDocument(pdf_doc)
            pdf_view.setPageMode(QPdfView.PageMode.MultiPage)
            layout.addWidget(pdf_view)
        else:
            # Fallback: show message
            label = QLabel(f"PDF generated at:\\n{pdf_path}\\n\\nPlease open with external viewer.")
            label.setWordWrap(True)
            layout.addWidget(label)
        
            open_btn = QPushButton("Open with System Viewer")
            open_btn.clicked.connect(lambda: os.startfile(pdf_path) if os.name == 'nt' else os.system(f'xdg-open "{pdf_path}"'))
            layout.addWidget(open_btn)
    
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.close)
        layout.addWidget(close_btn)
    
        dialog.exec()

    def generate_single_batch_report(self):
        """Generate report for current batch patient only"""
        if not hasattr(self, 'current_batch_index'):
            return
    
        output_dir = self.bulk_output_label.text()
        if output_dir == "No folder selected":
            QMessageBox.warning(self, "No Output Directory", "Please select an output directory first")
            return
    
        # Save current edits first
        self.save_batch_edits()
    
        idx = self.current_batch_index
        data = self.bulk_patient_data_list[idx]
    
        try:
            template = PGTAReportTemplate(assets_dir="assets/pgta")
            show_logo = self.bulk_logo_combo.currentText() == "With Logo"
            logo_suffix = "_withlogo" if show_logo else "_withoutlogo"
            p_name = data['patient_info']['patient_name'].replace(' ', '_')
            
            # Generate both PDF and DOCX if requested in settings
            if self.generate_pdf_check.isChecked():
                pdf_path = os.path.join(output_dir, f"{p_name}_PGTA_Report{logo_suffix}.pdf")
                template.generate_pdf(pdf_path, data['patient_info'], data['embryos'], show_logo=show_logo)
                self.statusBar().showMessage(f"Generated PDF for {data['patient_info']['patient_name']}")
                
            if self.generate_docx_check.isChecked():
                from pgta_docx_generator import PGTADocxGenerator
                docx_gen = PGTADocxGenerator(assets_dir="assets/pgta")
                docx_path = os.path.join(output_dir, f"{p_name}_PGTA_Report{logo_suffix}.docx")
                docx_gen.generate_docx(docx_path, data['patient_info'], data['embryos'], show_logo=show_logo)
                self.statusBar().showMessage(f"Generated DOCX for {data['patient_info']['patient_name']}")
        
            QMessageBox.information(self, "Success", f"Reports generated for {data['patient_info']['patient_name']}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate report:\\n{str(e)}")

    def save_bulk_draft(self):
        """Save batch to JSON file"""
        if not self.bulk_patient_data_list:
            QMessageBox.warning(self, "No Data", "No batch data to save")
            return
        
        path, _ = QFileDialog.getSaveFileName(self, "Save Batch Draft", "", "JSON Files (*.json)")
        if path:
            try:
                with open(path, 'w') as f:
                    json.dump(self.bulk_patient_data_list, f, indent=4)
                QMessageBox.information(self, "Success", f"Batch draft saved to {os.path.basename(path)}")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def load_bulk_draft(self):
        """Load batch from JSON file"""
        path, _ = QFileDialog.getOpenFileName(self, "Load Batch Draft", "", "JSON Files (*.json)")
        if path:
            try:
                with open(path, 'r') as f:
                    self.bulk_patient_data_list = json.load(f)
                
                # Populate batch list
                self.batch_list_widget.clear()
                for i, data in enumerate(self.bulk_patient_data_list):
                    p_name = data['patient_info']['patient_name']
                    e_count = len(data['embryos'])
                    item = QListWidgetItem(f"{p_name} ({e_count} embryos)")
                    item.setData(Qt.ItemDataRole.UserRole, i)
                    self.batch_list_widget.addItem(item)
                
                self.update_data_summary()
                QMessageBox.information(self, "Success", f"Loaded {len(self.bulk_patient_data_list)} patients")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def generate_all_batch_reports(self):
        """Generate reports for all patients in batch"""
        if not self.bulk_patient_data_list:
            QMessageBox.warning(self, "No Data", "No batch data to generate reports from")
            return
    
        output_dir = self.bulk_output_label.text()
        if output_dir == "No folder selected":
            QMessageBox.warning(self, "No Output Directory", "Please select an output directory first")
            return
    
        # Use existing report generation worker
        self.progress_bar.setVisible(True)
        self.progress_label.setVisible(True)
        self.generate_btn.setEnabled(False)
    
        self.worker = ReportGeneratorWorker(
            self.bulk_patient_data_list,
            output_dir,
            self.generate_pdf_check.isChecked(),
            self.generate_docx_check.isChecked(),
            "PGT-A",
            show_logo=(self.bulk_logo_combo.currentText() == "With Logo")
        )
    
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.generation_finished)
        self.worker.error.connect(self.generation_error)
        self.worker.start()

    def load_bulk_data(self):
        """Load data from bulk file with robust extraction logic"""
        file_path = self.bulk_file_label.text()
        
        if file_path == "No file selected":
            QMessageBox.warning(self, "Warning", "Please select a file first")
            return
        
        try:
            # Helper to find column regardless of case/space
            def get_col_name(df, possible_names):
                cols = [c.strip().lower() for c in df.columns]
                for name in possible_names:
                    n = name.strip().lower()
                    if n in cols:
                        return df.columns[cols.index(n)]
                return None

            # Helper to clean value (remove nan, strip)
            def clean_val(val, default=""):
                if pd.isna(val) or str(val).lower().strip() == "nan":
                    return default
                return str(val).strip()

            # Load file forcing ALL as string to preserve leading zeros
            if file_path.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path, dtype=str)
            else:
                df = pd.read_csv(file_path, sep='\t' if file_path.endswith('.tsv') else ',', dtype=str)
            
            # Pre-fill all NaN cells to avoid grouping issues
            df = df.fillna("")
            
            # Find key columns
            c_sample = get_col_name(df, ['Sample_Number', 'Sample Number', 'SampleID', 'Sample_ID'])
            c_patient = get_col_name(df, ['Patient_Name', 'Patient Name', 'Patient'])
            c_embryo = get_col_name(df, ['Embryo_ID', 'Embryo ID', 'Embryo'])
            
            if not c_sample or not c_patient:
                raise ValueError(f"Missing required columns (Sample Number, Patient Name). Found: {', '.join(df.columns)}")

            # Group by Sample Number
            grouped = df.groupby(c_sample)
            self.bulk_patient_data_list = []
            
            for sample_num, group in grouped:
                first_row = group.iloc[0]
                
                # Full patient info mapping
                patient_info = {
                    'patient_name': clean_val(first_row.get(get_col_name(df, ['Patient_Name', 'Patient Name']))),
                    'spouse_name': clean_val(first_row.get(get_col_name(df, ['Spouse_Name', 'Spouse Name']))),
                    'pin': clean_val(first_row.get(get_col_name(df, ['PIN']))),
                    'age': clean_val(first_row.get(get_col_name(df, ['Age']))),
                    'sample_number': clean_val(sample_num),
                    'referring_clinician': clean_val(first_row.get(get_col_name(df, ['Referring_Clinician', 'Referring Clinician', 'Clinician']))),
                    'biopsy_date': clean_val(first_row.get(get_col_name(df, ['Biopsy_Date', 'Biopsy Date']))),
                    'hospital_clinic': clean_val(first_row.get(get_col_name(df, ['Hospital_Clinic', 'Hospital/Clinic', 'Hospital']))),
                    'sample_collection_date': clean_val(first_row.get(get_col_name(df, ['Sample_Collection_Date', 'Sample Collection Date']))),
                    'specimen': clean_val(first_row.get(get_col_name(df, ['Specimen'])), "Day 6 Trophectoderm Biopsy"),
                    'sample_receipt_date': clean_val(first_row.get(get_col_name(df, ['Sample_Receipt_Date', 'Sample Receipt Date']))),
                    'biopsy_performed_by': clean_val(first_row.get(get_col_name(df, ['Biopsy_Performed_By', 'Biopsy Performed By']))),
                    'report_date': clean_val(first_row.get(get_col_name(df, ['Report_Date', 'Report Date'])), datetime.now().strftime("%d-%m-%Y")),
                    'indication': clean_val(first_row.get(get_col_name(df, ['Indication'])))
                }
                
                embryos = []
                for _, row in group.iterrows():
                    embryo_id = clean_val(row.get(c_embryo))
                    
                    # Image matching
                    cnv_image_path = None
                    if embryo_id in self.uploaded_images and self.uploaded_images[embryo_id]:
                        cnv_image_path = self.uploaded_images[embryo_id][0]
                    
                    # Chromosome statuses (1-22) and Mosaics
                    chr_statuses = {}
                    mosaic_percentages = {}
                    
                    for i in range(1, 23):
                        s_i = str(i)
                        # Status
                        c_status = get_col_name(df, [s_i])
                        chr_statuses[s_i] = clean_val(row.get(c_status) if c_status else None, 'N')
                        
                        # Mosaic
                        c_mosaic = get_col_name(df, [f"{s_i}_Mosaic", f"{s_i} Mosaic", f"Chr{s_i}_Mosaic"])
                        mos_val = clean_val(row.get(c_mosaic) if c_mosaic else None)
                        if mos_val:
                            mosaic_percentages[s_i] = mos_val

                    embryos.append({
                        'embryo_id': embryo_id,
                        'cnv_image_path': cnv_image_path,
                        'result_summary': clean_val(row.get(get_col_name(df, ['Result_Summary', 'Result Summary', 'Result']))),
                        'result_description': clean_val(row.get(get_col_name(df, ['Result_Description', 'Result Description', 'Conclusion']))),
                        'autosomes': clean_val(row.get(get_col_name(df, ['Autosomes']))),
                        'sex_chromosomes': clean_val(row.get(get_col_name(df, ['Sex_Chromosomes', 'Sex Chromosomes', 'Sex'])), "Normal"),
                        'interpretation': clean_val(row.get(get_col_name(df, ['Interpretation']))),
                        'mtcopy': clean_val(row.get(get_col_name(df, ['MTcopy', 'MT copy'])), 'NA'),
                        'chromosome_statuses': chr_statuses,
                        'mosaic_percentages': mosaic_percentages
                    })
                
                self.bulk_patient_data_list.append({'patient_info': patient_info, 'embryos': embryos})
            
            # Final stats for confirmation
            total_embryos = sum(len(p['embryos']) for p in self.bulk_patient_data_list)
            QMessageBox.information(self, "Success", f"Loaded {len(self.bulk_patient_data_list)} patients with {total_embryos} embryos in total.")
            self.statusBar().showMessage(f"Bulk data loaded: {total_embryos} samples found.")
            self.update_data_summary() 
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Error", f"Failed to load data: {str(e)}")
            
    
    
    def add_images_with_embryo_id(self):
        """Add CNV chart images with embryo ID assignment"""
        from PyQt6.QtWidgets import QVBoxLayout as QVBoxLayoutDialog
        
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select CNV Chart Images",
            "",
            "Image Files (*.png *.jpg *.jpeg)"
        )
        
        if not file_paths:
            return
        
        # For each image, ask for embryo ID
        for path in file_paths:
            # Create dialog to get embryo ID
            dialog = QDialog(self)
            dialog.setWindowTitle(f"Assign Embryo ID - {os.path.basename(path)}")
            dialog_layout = QVBoxLayoutDialog()
            
            # Instructions
            instruction_label = QLabel(f"Assign an Embryo ID for:\n{os.path.basename(path)}")
            instruction_label.setWordWrap(True)
            dialog_layout.addWidget(instruction_label)
            
            # Embryo ID input
            embryo_id_layout = QHBoxLayout()
            embryo_id_layout.addWidget(QLabel("Embryo ID:"))
            embryo_id_input = QLineEdit()
            embryo_id_input.setPlaceholderText("e.g., PS1, PS2, PS4, PS9")
            
            # Try to auto-suggest from filename
            filename = os.path.basename(path)
            for possible_id in ['PS1', 'PS2', 'PS3', 'PS4', 'PS5', 'PS6', 'PS7', 'PS8', 'PS9', 'PS10']:
                if possible_id in filename.upper():
                    embryo_id_input.setText(possible_id)
                    break
            
            embryo_id_layout.addWidget(embryo_id_input)
            dialog_layout.addLayout(embryo_id_layout)
            
            # Buttons
            button_box = QDialogButtonBox(
                QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
            )
            button_box.accepted.connect(dialog.accept)
            button_box.rejected.connect(dialog.reject)
            dialog_layout.addWidget(button_box)
            
            dialog.setLayout(dialog_layout)
            
            # Show dialog
            if dialog.exec() == QDialog.DialogCode.Accepted:
                embryo_id = embryo_id_input.text().strip()
                
                if not embryo_id:
                    QMessageBox.warning(self, "Warning", "Embryo ID cannot be empty!")
                    continue
                
                # Add to table
                row_position = self.image_table.rowCount()
                self.image_table.insertRow(row_position)
                
                # Embryo ID (editable)
                embryo_id_item = QTableWidgetItem(embryo_id)
                self.image_table.setItem(row_position, 0, embryo_id_item)
                
                # Filename
                filename_item = QTableWidgetItem(os.path.basename(path))
                filename_item.setFlags(filename_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.image_table.setItem(row_position, 1, filename_item)
                
                # Full path (hidden but stored)
                path_item = QTableWidgetItem(path)
                path_item.setFlags(path_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.image_table.setItem(row_position, 2, path_item)
                
                # Remove button
                remove_btn = QPushButton()
                remove_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon))
                remove_btn.setMaximumWidth(40)
                remove_btn.clicked.connect(lambda checked, row=row_position: self.remove_image_row(row))
                self.image_table.setCellWidget(row_position, 3, remove_btn)
                
                # Store in uploaded_images dict with embryo ID as key
                if embryo_id not in self.uploaded_images:
                    self.uploaded_images[embryo_id] = []
                self.uploaded_images[embryo_id].append(path)
        
        self.update_image_summary()
        self.statusBar().showMessage(f"Added {len(file_paths)} image(s)")
    
    def remove_image_row(self, row):
        """Remove a specific image row"""
        # Get embryo ID and path before removing
        embryo_id = self.image_table.item(row, 0).text()
        path = self.image_table.item(row, 2).text()
        
        # Remove from uploaded_images dict
        if embryo_id in self.uploaded_images and path in self.uploaded_images[embryo_id]:
            self.uploaded_images[embryo_id].remove(path)
            if not self.uploaded_images[embryo_id]:  # Remove key if empty
                del self.uploaded_images[embryo_id]
        
        # Remove row from table
        self.image_table.removeRow(row)
        
        # Update all remove button connections (row indices changed)
        for i in range(self.image_table.rowCount()):
            remove_btn = self.image_table.cellWidget(i, 3)
            if remove_btn:
                remove_btn.clicked.disconnect()
                remove_btn.clicked.connect(lambda checked, r=i: self.remove_image_row(r))
        
        self.update_image_summary()
        self.statusBar().showMessage("Image removed")
    
    def clear_all_images(self):
        """Clear all images"""
        reply = QMessageBox.question(
            self, 'Confirm Clear',
            'Are you sure you want to clear all images?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.image_table.setRowCount(0)
            self.uploaded_images.clear()
            self.update_image_summary()
            self.statusBar().showMessage("All images cleared")
    
    def update_image_summary(self):
        """Update image summary label"""
        total_images = self.image_table.rowCount()
        unique_embryos = len(self.uploaded_images)
        
        if total_images == 0:
            self.image_summary_label.setText("No images uploaded")
        else:
            self.image_summary_label.setText(
                f"Total: {total_images} image(s) for {unique_embryos} embryo(s)"
            )
    
    def get_embryo_images(self):
        """Get dictionary of embryo ID to image paths"""
        embryo_images = {}
        
        for row in range(self.image_table.rowCount()):
            embryo_id = self.image_table.item(row, 0).text()
            image_path = self.image_table.item(row, 2).text()
            
            if embryo_id not in embryo_images:
                embryo_images[embryo_id] = []
            embryo_images[embryo_id].append(image_path)
        
        return embryo_images

    
    def browse_output_dir(self):
        """Browse for output directory"""
        dir_path = QFileDialog.getExistingDirectory(
            self,
            "Select Output Directory",
            self.settings.value('last_output_dir', '')
        )
        
        if dir_path:
            self.output_dir_label.setText(dir_path)
            self.settings.setValue('last_output_dir', dir_path)
    
    
    def update_progress(self, value, message):
        """Update progress bar"""
        self.progress_bar.setValue(value)
        self.progress_label.setText(message)
        self.statusBar().showMessage(message)
    
    def generation_finished(self, success_reports, failed_reports):
        """Handle generation completion"""
        self.progress_bar.setVisible(False)
        self.progress_label.setVisible(False)
        self.generate_btn.setEnabled(True)
        
        message = f"Successfully generated {len(success_reports)} report(s)"
        if failed_reports:
            message += f"\nFailed: {len(failed_reports)}"
        
        QMessageBox.information(self, "Complete", message)
        
        # Offer to open output folder
        reply = QMessageBox.question(
            self, 'Open Folder',
            'Would you like to open the output folder?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            import subprocess
            import platform
            
            output_dir = self.output_dir_label.text()
            if platform.system() == 'Windows':
                os.startfile(output_dir)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.Popen(['open', output_dir])
            else:  # Linux
                subprocess.Popen(['xdg-open', output_dir])
    
    def generation_error(self, error_message):
        """Handle generation error"""
        QMessageBox.critical(self, "Error", error_message)
    
    def load_settings(self):
        """Load saved settings"""
        # Restore window geometry
        geometry = self.settings.value('geometry')
        if geometry:
            self.restoreGeometry(geometry)
        
        # Restore last output directory
        last_output = self.settings.value('last_output_dir')
        if last_output:
            self.output_dir_label.setText(last_output)
    
    def closeEvent(self, event):
        """Save settings on close"""
        self.settings.setValue('geometry', self.saveGeometry())
        event.accept()


def main():
    """Main entry point"""
    app = QApplication(sys.argv)
    app.setApplicationName("PGT-A Report Generator")
    app.setOrganizationName("PGTA")

    window = PGTAReportGeneratorApp()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()

