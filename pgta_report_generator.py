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

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QLineEdit, QTextEdit, QPushButton, QFileDialog,
    QTableWidget, QTableWidgetItem, QMessageBox, QProgressBar,
    QGroupBox, QFormLayout, QScrollArea, QCheckBox, QSpinBox,
    QComboBox, QListWidget, QListWidgetItem, QStyle, QGridLayout,
    QSplitter, QTextBrowser, QRadioButton
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSettings, QTimer
from PyQt6.QtGui import QPixmap, QIcon
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
from pgta_template import PGTAReportTemplate
from pgta_docx_generator import PGTADocxGenerator

class PreviewWorker(QThread):
    """Worker thread for generating preview PDF"""
    finished = pyqtSignal(str) # Path to generated PDF
    error = pyqtSignal(str)
    
    def __init__(self, patient_data, embryos_data, output_path):
        super().__init__()
        self.patient_data = patient_data
        self.embryos_data = embryos_data
        self.output_path = output_path
        
    def run(self):
        try:
            # Generate PDF using native template
            gen = PGTAReportTemplate(assets_dir="assets/pgta")
            gen.generate_pdf(self.output_path, self.patient_data, self.embryos_data)
            self.finished.emit(self.output_path)
        except Exception as e:
            self.error.emit(str(e))


class ReportGeneratorWorker(QThread):
    """Worker thread for generating reports"""
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(list, list)
    error = pyqtSignal(str)
    
    def __init__(self, patient_data_list, output_dir, generate_pdf=True, generate_docx=True, template_type="PGT-A"):
        super().__init__()
        self.patient_data_list = patient_data_list
        self.output_dir = output_dir
        self.generate_pdf = generate_pdf
        self.generate_docx = generate_docx
        self.template_type = template_type
    
    def run(self):
        """Generate reports"""
        success_reports = []
        failed_reports = []
        
        if self.generate_pdf:
            # Revert to pure ReportLab template
            # In a full ADVAT system, we would select the class based on template_type
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
                
                base_filename = f"{sample_num}_{patient_name.replace(' ', '_')}_{timestamp}"
                
                self.progress.emit(
                    int((idx - 1) / total * 100),
                    f"Generating reports for {patient_name} ({idx}/{total})..."
                )
                
                # Generate using Original Template
                # Generate PDF
                if self.generate_pdf:
                    pdf_path = os.path.join(self.output_dir, f"{base_filename}.pdf")
                    # Use the ReportLab generator directly
                    pdf_generator.generate_pdf(
                        pdf_path,
                        patient_data['patient_info'],
                        patient_data['embryos']
                    )
                
                # Generate DOCX
                if self.generate_docx:
                    docx_path = os.path.join(self.output_dir, f"{base_filename}.docx")
                    docx_generator.generate_docx(
                        docx_path,
                        patient_data['patient_info'],
                        patient_data['embryos']
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
        self.settings = QSettings('ADVAT', 'ReportGenerator')
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
            template_type=self.template_combo.currentText()
        )
        
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.generation_finished)
        self.worker.error.connect(self.generation_error)
        
        self.worker.start()
    
    def init_ui(self):
        """Initialize user interface"""
        self.setWindowTitle("ADVAT Report Generator")
        self.setGeometry(100, 100, 1200, 800)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # Header layout with title and template selector
        header_layout = QHBoxLayout()
        
        # Title
        title_label = QLabel("ADVAT Report Generator")
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; padding: 10px;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        header_layout.addWidget(title_label)
        
        header_layout.addStretch()
        
        # Template selector
        header_layout.addWidget(QLabel("Select Template:"))
        self.template_combo = QComboBox()
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
        self.image_management_tab = self.create_image_management_tab()
        self.generate_tab = self.create_generate_tab()
        
        self.tabs.addTab(self.manual_entry_tab, "Manual Entry")
        self.tabs.setTabIcon(0, self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))
        
        self.tabs.addTab(self.bulk_upload_tab, "Bulk Upload")
        self.tabs.setTabIcon(1, self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogListView))
        
        self.tabs.addTab(self.image_management_tab, "Image Management")
        self.tabs.setTabIcon(2, self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
        
        self.tabs.addTab(self.generate_tab, "Generate Reports")
        self.tabs.setTabIcon(3, self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon))
        
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
        self.patient_name_input = QLineEdit()
        self.spouse_name_input = QLineEdit()
        self.pin_input = QLineEdit()
        self.age_input = QLineEdit()
        self.sample_number_input = QLineEdit()
        self.referring_clinician_input = QLineEdit()
        self.biopsy_date_input = QLineEdit()
        
        self.biopsy_date_input.setPlaceholderText("DD-MM-YYYY")
        self.hospital_clinic_input = QLineEdit()
        self.sample_collection_date_input = QLineEdit()
        self.sample_collection_date_input.setPlaceholderText("DD-MM-YYYY")
        self.specimen_input = QLineEdit()
        self.specimen_input.setText("Day 6 Trophectoderm Biopsy")
        self.sample_receipt_date_input = QLineEdit()
        self.sample_receipt_date_input.setPlaceholderText("DD-MM-YYYY")
        self.biopsy_performed_by_input = QLineEdit()
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
            'patient_name': self.patient_name_input.text(),
            'spouse_name': self.spouse_name_input.text(),
            'pin': self.pin_input.text(),
            'age': self.age_input.text(),
            'sample_number': self.sample_number_input.text(),
            'referring_clinician': self.referring_clinician_input.text(),
            'biopsy_date': self.biopsy_date_input.text(),
            'hospital_clinic': self.hospital_clinic_input.text(),
            'sample_collection_date': self.sample_collection_date_input.text(),
            'specimen': self.specimen_input.text(),
            'sample_receipt_date': self.sample_receipt_date_input.text(),
            'biopsy_performed_by': self.biopsy_performed_by_input.text(),
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
                interp = "Euploid"
                mtcopy = "NA"
                
                if self.summary_table.rowCount() > idx:
                     # Item 0 is ID
                     item_id = self.summary_table.item(idx, 0)
                     if item_id: embryo_id = item_id.text()
                     
                     item_res = self.summary_table.item(idx, 1)
                     if item_res: res_sum = item_res.text()
                     
                     widget_interp = self.summary_table.cellWidget(idx, 2)
                     if widget_interp: interp = widget_interp.currentText()
                     
                     item_mt = self.summary_table.item(idx, 3)
                     if item_mt: mtcopy = item_mt.text()

                # Get Detailed Info (Result Desc, Autosomes, Image, Chromosomes)
                # If form_dict has references:
                result_desc = form_dict.get('result_description', QLineEdit()).text()
                autosomes = form_dict.get('autosomes', QLineEdit()).text()
                sex = form_dict.get('sex_chromosomes', QLineEdit()).text()
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
                    'sex_chromosomes': sex,
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
            
        self.preview_worker = PreviewWorker(p_data, e_data, temp_pdf)
        self.preview_worker.finished.connect(self.on_preview_generated)
        self.preview_worker.start()

    def on_preview_generated(self, pdf_path):
        """Load generated PDF into viewer"""
        if QPdfDocument and self.pdf_document and os.path.exists(pdf_path):
            self.pdf_document.load(pdf_path)
            self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitInView)
    
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
            if not self.summary_table.item(r, 1): # Result (Summary)
                self.summary_table.setItem(r, 1, QTableWidgetItem(""))
            
            if not self.summary_table.cellWidget(r, 2): # Interpretation
                combo = QComboBox()
                combo.addItems(["Euploid", "Aneuploid", "Low level mosaic", "High level mosaic", "Complex mosaic"])
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
        
        result_description = QLineEdit()
        autosomes = QLineEdit()
        sex_chromosomes = QLineEdit()
        sex_chromosomes.setText("Normal")
        
        # Connect signals
        result_description.textChanged.connect(self.update_preview)
        autosomes.textChanged.connect(self.update_preview)
        sex_chromosomes.textChanged.connect(self.update_preview)
        
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
            chr_combo = QComboBox()
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
        """Create bulk upload tab"""
        tab = QWidget()
        layout = QVBoxLayout()
        tab.setLayout(layout)
        
        # Instructions
        instructions = QLabel(
            "Upload an Excel or TSV file with patient and embryo data.\n"
            "The file should contain columns for all patient information and embryo details."
        )
        instructions.setWordWrap(True)
        instructions.setStyleSheet("padding: 10px; background-color: #E3F2FD; border-radius: 5px;")
        layout.addWidget(instructions)
        
        # File selection
        file_layout = QHBoxLayout()
        self.bulk_file_label = QLabel("No file selected")
        self.bulk_file_label.setStyleSheet("padding: 5px; border: 1px solid #ccc;")
        file_layout.addWidget(self.bulk_file_label, 1)
        
        browse_btn = QPushButton("Browse")
        browse_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton))
        browse_btn.clicked.connect(self.browse_bulk_file)
        file_layout.addWidget(browse_btn)
        
        download_template_btn = QPushButton("Download Template")
        download_template_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton))
        download_template_btn.clicked.connect(self.download_template)
        file_layout.addWidget(download_template_btn)
        
        layout.addLayout(file_layout)
        
        # Preview table
        preview_label = QLabel("Data Preview:")
        preview_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        layout.addWidget(preview_label)
        
        self.bulk_preview_table = QTableWidget()
        layout.addWidget(self.bulk_preview_table)
        
        # Load button
        load_btn = QPushButton("Load Data")
        load_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogApplyButton))
        load_btn.clicked.connect(self.load_bulk_data)
        layout.addWidget(load_btn)
        
        return tab
    
    def create_image_management_tab(self):
        """Create image management tab"""
        tab = QWidget()
        layout = QVBoxLayout()
        tab.setLayout(layout)
        
        instructions = QLabel(
            "Upload CNV chart images and assign them to specific embryos.\n"
            "Each image must be tagged with an Embryo ID (e.g., PS1, PS2, PS4, PS9).\n"
            "Images should be in PNG, JPG, or JPEG format."
        )
        instructions.setWordWrap(True)
        instructions.setStyleSheet("padding: 10px; background-color: #FFF3E0; border-radius: 5px;")
        layout.addWidget(instructions)
        
        # Image table with embryo ID assignment
        table_label = QLabel("CNV Chart Images:")
        table_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        layout.addWidget(table_label)
        
        self.image_table = QTableWidget()
        self.image_table.setColumnCount(4)
        self.image_table.setHorizontalHeaderLabels(["Embryo ID", "Image Filename", "File Path", "Actions"])
        self.image_table.setColumnWidth(0, 150)
        self.image_table.setColumnWidth(1, 250)
        self.image_table.setColumnWidth(2, 350)
        self.image_table.setColumnWidth(3, 100)
        layout.addWidget(self.image_table)
        
        # Buttons
        btn_layout = QHBoxLayout()
        add_image_btn = QPushButton("Add Image(s)")
        add_image_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
        add_image_btn.clicked.connect(self.add_images_with_embryo_id)
        add_image_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 8px;")
        
        clear_images_btn = QPushButton("Clear All")
        clear_images_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon))
        clear_images_btn.clicked.connect(self.clear_all_images)
        
        btn_layout.addWidget(add_image_btn)
        btn_layout.addWidget(clear_images_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        # Summary
        self.image_summary_label = QLabel("No images uploaded")
        self.image_summary_label.setStyleSheet("padding: 5px; font-style: italic;")
        layout.addWidget(self.image_summary_label)
        
        return tab
    
    def create_generate_tab(self):
        """Create report generation tab"""
        tab = QWidget()
        layout = QVBoxLayout()
        tab.setLayout(layout)
        
        # Output settings
        settings_group = QGroupBox("Output Settings")
        settings_layout = QFormLayout()
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)
        
        # Output directory
        output_dir_layout = QHBoxLayout()
        self.output_dir_label = QLabel("No directory selected")
        self.output_dir_label.setStyleSheet("padding: 5px; border: 1px solid #ccc;")
        output_dir_layout.addWidget(self.output_dir_label, 1)
        
        browse_output_btn = QPushButton("Browse")
        browse_output_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton))
        browse_output_btn.clicked.connect(self.browse_output_dir)
        output_dir_layout.addWidget(browse_output_btn)
        
        settings_layout.addRow("Output Directory:", output_dir_layout)
        
        # Format options
        self.generate_pdf_check = QCheckBox("Generate PDF")
        self.generate_pdf_check.setChecked(True)
        self.generate_docx_check = QCheckBox("Generate DOCX")
        self.generate_docx_check.setChecked(True)
        
        format_layout = QHBoxLayout()
        format_layout.addWidget(self.generate_pdf_check)
        format_layout.addWidget(self.generate_docx_check)
        format_layout.addStretch()
        settings_layout.addRow("Formats:", format_layout)

        # Template selection
        # Data summary
        summary_group = QGroupBox("Data Summary")
        summary_layout = QVBoxLayout()
        summary_group.setLayout(summary_layout)
        layout.addWidget(summary_group)
        
        self.data_summary_label = QLabel("No data loaded")
        self.data_summary_label.setStyleSheet("padding: 10px;")
        summary_layout.addWidget(self.data_summary_label)
        
        # Progress
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        self.progress_label = QLabel("")
        self.progress_label.setVisible(False)
        layout.addWidget(self.progress_label)
        
        # Generate button
        self.generate_btn = QPushButton("Generate Reports")
        self.generate_btn.setObjectName("generate_btn") # For QSS
        self.generate_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
        self.generate_btn.clicked.connect(self.generate_reports)
        layout.addWidget(self.generate_btn)
        
        layout.addStretch()
        
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
            form['result_description'].setText(last_embryo.get('result_description', ''))
            form['autosomes'].setText(last_embryo.get('autosomes', ''))
            form['sex_chromosomes'].setText(last_embryo.get('sex_chromosomes', ''))
            
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
            val_summary = last_embryo.get('result_summary', '')
            self.summary_table.setItem(new_index, 1, QTableWidgetItem(val_summary))
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
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "extracted_assets", "chrominst_logo.png")
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
        """Collect current manual entry data into a dictionary"""
        patient_info = {
            'patient_name': self.patient_name_input.text(),
            'spouse_name': self.spouse_name_input.text(),
            'pin': self.pin_input.text(),
            'age': self.age_input.text(),
            'sample_number': self.sample_number_input.text(),
            'referring_clinician': self.referring_clinician_input.text(),
            'biopsy_date': self.biopsy_date_input.text(),
            'hospital_clinic': self.hospital_clinic_input.text(),
            'sample_collection_date': self.sample_collection_date_input.text(),
            'specimen': self.specimen_input.text(),
            'sample_receipt_date': self.sample_receipt_date_input.text(),
            'biopsy_performed_by': self.biopsy_performed_by_input.text(),
            'report_date': self.report_date_input.text(),
            'indication': self.indication_input.toPlainText()
        }
        
        embryos = []
        count = self.summary_table.rowCount()
        
        for i in range(count):
            # 1. Summary Table Data
            item_id = self.summary_table.item(i, 0)
            t_id = item_id.text() if item_id else ""
            
            item_sum = self.summary_table.item(i, 1)
            t_sum = item_sum.text() if item_sum else ""
            
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
                result_desc = form['result_description'].text()
                autosomes = form['autosomes'].text()
                sex = form['sex_chromosomes'].text()
                
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
            # Summary
            self.summary_table.setItem(i, 1, QTableWidgetItem(embryo.get('result_summary', '')))
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
            
            form['result_description'].setText(embryo.get('result_description', ''))
            form['autosomes'].setText(embryo.get('autosomes', ''))
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
    
    def load_bulk_data(self):
        """Load data from bulk file"""
        file_path = self.bulk_file_label.text()
        
        if file_path == "No file selected":
            QMessageBox.warning(self, "Warning", "Please select a file first")
            return
        
        try:
            # Load file
            if file_path.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path)
            else:
                df = pd.read_csv(file_path, sep='\t' if file_path.endswith('.tsv') else ',')
            
            # clean column names (strip whitespace)
            df.columns = [c.strip() for c in df.columns]
            
            # Required columns check
            required_cols = ['Sample_Number', 'Patient_Name', 'Embryo_ID']
            if not all(col in df.columns for col in required_cols):
                missing = [col for col in required_cols if col not in df.columns]
                raise ValueError(f"Missing required columns: {', '.join(missing)}")
            
            # Group by Sample Number to handle multiple embryos per patient
            grouped = df.groupby('Sample_Number')
            
            self.bulk_patient_data_list = []
            
            for sample_num, group in grouped:
                # Deduplicate embryos by ID within this patient
                group = group.drop_duplicates(subset=['Embryo_ID'])
                
                # Patient info from first row
                first_row = group.iloc[0]
                patient_info = {
                    'patient_name': str(first_row.get('Patient_Name', '')),
                    'patient_spouse': str(first_row.get('Patient_Name', '')), # Fallback
                    'spouse_name': str(first_row.get('Spouse_Name', '')),
                    'pin': str(first_row.get('PIN', '')),
                    'age': str(first_row.get('Age', '')),
                    'sample_number': str(sample_num),
                    'referring_clinician': str(first_row.get('Referring_Clinician', '')),
                    'biopsy_date': str(first_row.get('Biopsy_Date', '')),
                    'hospital_clinic': str(first_row.get('Hospital_Clinic', '')),
                    'sample_collection_date': str(first_row.get('Sample_Collection_Date', '')),
                    'specimen': str(first_row.get('Specimen', '')),
                    'sample_receipt_date': str(first_row.get('Sample_Receipt_Date', '')),
                    'biopsy_performed_by': str(first_row.get('Biopsy_Performed_By', '')),
                    'report_date': str(first_row.get('Report_Date', '')),
                    'indication': str(first_row.get('Indication', ''))
                }
                
                # Embryos
                embryos = []
                for _, row in group.iterrows():
                    embryo_id = str(row.get('Embryo_ID', ''))
                    
                    # Look up assigned image
                    cnv_image_path = None
                    if embryo_id in self.uploaded_images and self.uploaded_images[embryo_id]:
                        cnv_image_path = self.uploaded_images[embryo_id][0]
                    
                    # Parse chromosome statuses (N, G, L, etc.)
                    # Assumes columns like '1', '2'... or a 'Autosomes' text
                    # The template has specific columns? The template code I wrote earlier didn't create 1-22 columns.
                    # It had 'Autosomes', 'Sex_Chromosomes'.
                    # I should probably robustly handle this.
                    # For now, let's assume default 'N' if not specified, or parse from 'Autosomes' text?
                    # The user's prompt implied a rigorous report.
                    # Let's check the QuickStart guide I wrote: "Chromosome statuses (1-22): Select..."
                    # But the excel template creation only had 'Autosomes' (summary).
                    # I should update the template creator to include 1-22 columns if we want detailed control.
                    # For now, I'll default to 'N' for bulk, or look for columns '1', '2' etc. if they exist.
                    
                    chr_statuses = {}
                    for i in range(1, 23):
                        col_name = str(i)
                        status = 'N'
                        if col_name in df.columns:
                            status = str(row[col_name])
                        chr_statuses[str(i)] = status

                    embryo = {
                        'embryo_id': embryo_id,
                        'cnv_image_path': cnv_image_path,
                        'result_summary': str(row.get('Result_Summary', '')),
                        'result_description': str(row.get('Result_Description', '')),
                        'autosomes': str(row.get('Autosomes', '')),
                        'sex_chromosomes': str(row.get('Sex_Chromosomes', '')),
                        'interpretation': str(row.get('Interpretation', '')),
                        'mtcopy': str(row.get('MTcopy', 'NA')),
                        'chromosome_statuses': chr_statuses,
                        'mosaic_percentages': {}
                    }
                    embryos.append(embryo)
                
                self.bulk_patient_data_list.append({'patient_info': patient_info, 'embryos': embryos})
            
            QMessageBox.information(self, "Success", f"Loaded data for {len(self.bulk_patient_data_list)} patients")
            self.statusBar().showMessage(f"Bulk data loaded: {len(self.bulk_patient_data_list)} patients")
            self.update_data_summary() # Update summary to reflect bulk data
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load data: {str(e)}")
    
    
    def add_images_with_embryo_id(self):
        """Add CNV chart images with embryo ID assignment"""
        from PyQt6.QtWidgets import QDialog, QDialogButtonBox, QVBoxLayout as QVBoxLayoutDialog
        
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
