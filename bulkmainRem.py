import sys
import json
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout,
    QLabel, QFileDialog, QCheckBox, QLineEdit, QMessageBox
)
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Image, Spacer, PageTemplate, Frame
)
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.pdfgen.canvas import Canvas

COMPLIANCE_MAPPING = {
    'hipaa': 'HIPAA',
    'nist_800_53': 'NIST',
    'pci_dss': 'PCI DSS',
    'gdpr_IV': 'GDPR'
}

class NumberedCanvas(Canvas):
    def __init__(self, *args, leftMargin=0, rightMargin=0, bottomMargin=0, _pdf=None, **kwargs):
        Canvas.__init__(self, *args, **kwargs)
        self._pdf = _pdf  # store the pdf object as an instance variable
        self._leftMargin = leftMargin
        self._rightMargin = rightMargin
        self._bottomMargin = bottomMargin
        self._pageNumber = 1  # initialize the page number

        # Draw static footer elements
        self._draw_static_footer()

    def showPage(self):
        # Draw dynamic footer elements
        self._draw_dynamic_footer()
        # Draw header for the first page
        if self._pageNumber == 1:
            self._draw_header()
        # Show the page
        Canvas.showPage(self)

    def _draw_static_footer(self):
        footer_text = f'https://cns4u.com - Report generated on {datetime.now().strftime("%m-%d-%Y at %H:%M")}'
        self.setFont('Helvetica', 10)
        self.drawString(self._leftMargin, self._bottomMargin - 30, footer_text)

    def _draw_dynamic_footer(self):
        # Reset the font settings to ensure consistent appearance
        self.setFont('Helvetica', 10)
        page_number_text = f'Page {self._pageNumber}'
        self.drawRightString(self._pagesize[0] - self._rightMargin, self._bottomMargin - 30, page_number_text)

    def _draw_header(self):
        logo = Image(self._pdf.logo_path)
        width, height = logo.drawWidth, logo.drawHeight
        self.drawImage(self._pdf.logo_path, self._leftMargin, self._pagesize[1] - height, width=width, height=height)
        title_width = self.stringWidth(self._pdf.title, 'Helvetica-Bold', 24)
        title_x = self._leftMargin + width + ((self._pagesize[0] - self._leftMargin - width - title_width) / 2)
        self.setFont('Helvetica-Bold', 24)
        self.drawString(title_x, self._pagesize[1] - height/2 - 12, self._pdf.title)  # Adjust y position

def generate_pdf(data, document_title='', logo_path='', file_name='report.pdf'):
    pdf = SimpleDocTemplate(file_name, pagesize=landscape(letter))
    styles = getSampleStyleSheet()
    styleN = styles['Normal']
    styleN.alignment = 1  # Center alignment

    logo = Image(logo_path)
    logo.drawHeight = min(150, logo.drawHeight)  # limit height to 150 or its original height, whichever is smaller
    logo.drawWidth = logo.drawWidth * (logo.drawHeight / logo.imageHeight)  # maintain aspect ratio

    pdf.logo_path = logo_path  # assign string value to pdf.logo_path instead of pdf.logo
    pdf.title = document_title  # assign string value to pdf.title
    pdf.date = datetime.now()

    pdf_data = data.drop(columns='Remediation')
    table_data = [[Paragraph(cell, styleN) for cell in pdf_data.columns.values.tolist()]]  # Header row
    table_data.extend([[Paragraph(str(cell), styleN) for cell in row] for row in pdf_data.values.tolist()])  # Data rows
    col_widths = [1.5*inch, 2.5*inch, 2*inch, 2*inch, 1*inch]  # Adjust column widths
    table = Table(table_data, colWidths=col_widths)

    style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1,-1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('SIZE', (0, 0), (-1, -1), 10)
    ]
    for row_idx, row in enumerate(data.values.tolist(), start=1):
        cell_color = colors.green if row[4] == 'PASSED' else colors.red  # Adjust the index if necessary
        style_commands.append(('BACKGROUND', (4, row_idx), (4, row_idx), cell_color))

    style = TableStyle(style_commands)
    table.setStyle(style)

    frame = Frame(
        pdf.leftMargin, pdf.bottomMargin + 1.5 * inch, pdf.width, pdf.height - 2.5 * inch,
        id='content_frame', showBoundary=0  # hide frame border
    )
    
    main_page_template = PageTemplate(
        id='MainTemplate',
        frames=[frame]
    )

    pdf.addPageTemplates([main_page_template])
    pdf.build([table], canvasmaker=lambda *args, **kwargs: NumberedCanvas(
        *args, 
        leftMargin=pdf.leftMargin, 
        rightMargin=pdf.rightMargin, 
        bottomMargin=pdf.bottomMargin,
        _pdf=pdf,  # pass the pdf object to the canvas class
        **{k: v for k, v in kwargs.items() if k != '_pdf'}  # filter out the _pdf keyword argument
    ))

def generate_remediation_file(data, document_title, folder_path):
    file_name = f"{document_title}_Remediation.txt"
    file_path = os.path.join(folder_path, file_name)
    with open(file_path, 'w') as file:
        for index, row in data.iterrows():
            if row['Status'] == 'FAILED':
                file.write(f"Compliance Standard: {row['Compliance Standard']}\n")
                file.write(f"Rule Information: {row['Rule Information']}\n")
                file.write(f"Remediation: {row['Remediation']}\n")
                file.write('-' * 80 + '\n')

def process_folder(folder_path, selected_standards, document_title='', logo_path=''):
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.csv'):
            file_path = os.path.join(folder_path, file_name)
            processed_data = process_csv(file_path, selected_standards)
            pdf_file_name = os.path.splitext(file_name)[0] + '.pdf'
            pdf_file_path = os.path.join(folder_path, pdf_file_name)
            generate_pdf(processed_data, document_title, logo_path, pdf_file_path)
            generate_remediation_file(processed_data, document_title, folder_path)

def process_csv(file_path, selected_standards):
    data = pd.read_csv(file_path)
    computer_name = os.path.basename(file_path).split('.')[0]
    processed_data = []
    for _, row in data.iterrows():
        compliance_data = json.loads(row['Compliance'].replace("'", "\""))
        for item in compliance_data:
            if item['key'] in selected_standards:
                if row['Result'].upper().strip() != 'NOT APPLICABLE':
                    processed_data.append({
                        'Computer Name': computer_name,
                        'Rule Information': row['Title'],
                        'Compliance Standard': COMPLIANCE_MAPPING.get(item['key'], item['key']),
                        'Compliance Rule': item['value'],
                        'Status': row['Result'].upper(),
                        'Remediation': row['Remediation']  # Added Remediation column
                    })
    return pd.DataFrame(processed_data)

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('CRG x Steven')
        self.setFixedSize(400, 300)
        layout = QVBoxLayout()

        self.label = QLabel('Select CSV Folder:', self)
        layout.addWidget(self.label)

        self.btn_select_folder = QPushButton('Select Folder', self)
        self.btn_select_folder.clicked.connect(self.select_folder)
        layout.addWidget(self.btn_select_folder)

        self.logo_label = QLabel('Select Logo:', self)
        layout.addWidget(self.logo_label)

        self.btn_select_logo = QPushButton('Select Logo', self)
        self.btn_select_logo.clicked.connect(self.select_logo)
        layout.addWidget(self.btn_select_logo)

        self.title_label = QLabel('Document Title:', self)
        layout.addWidget(self.title_label)

        self.title_input = QLineEdit(self)
        layout.addWidget(self.title_input)

        self.checkbox_hipaa = QCheckBox('HIPAA', self)
        layout.addWidget(self.checkbox_hipaa)

        self.checkbox_nist = QCheckBox('NIST', self)
        layout.addWidget(self.checkbox_nist)

        self.checkbox_pci_dss = QCheckBox('PCI DSS', self)
        layout.addWidget(self.checkbox_pci_dss)

        self.checkbox_gdpr = QCheckBox('GDPR', self)
        layout.addWidget(self.checkbox_gdpr)

        self.btn_generate_report = QPushButton('Generate Report', self)
        self.btn_generate_report.clicked.connect(self.generate_report)
        layout.addWidget(self.btn_generate_report)

        self.setLayout(layout)

    def select_folder(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ShowDirsOnly
        folder_name = QFileDialog.getExistingDirectory(self, "Select Folder", "", options=options)
        if folder_name:
            self.label.setText(f'Selected Folder: {folder_name}')

    def select_logo(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Logo File", "", "Image Files (*.png *.jpg *.jpeg *.bmp);;All Files (*)", options=options)
        if file_name:
            self.logo_label.setText(f'Selected Logo: {file_name}')

    def generate_report(self):
        folder_path = self.label.text().replace('Selected Folder: ', '')
        document_title = self.title_input.text()
        logo_path = self.logo_label.text().replace('Selected Logo: ', '')

        if not folder_path or not any([self.checkbox_hipaa.isChecked(),
                                       self.checkbox_nist.isChecked(),
                                       self.checkbox_pci_dss.isChecked(),
                                       self.checkbox_gdpr.isChecked()]):
            QMessageBox.warning(self, 'Input Error', 'Please select a folder and at least one compliance standard.')
            return

        selected_standards = []
        if self.checkbox_hipaa.isChecked():
            selected_standards.append('hipaa')
        if self.checkbox_nist.isChecked():
            selected_standards.append('nist_800_53')
        if self.checkbox_pci_dss.isChecked():
            selected_standards.append('pci_dss')
        if self.checkbox_gdpr.isChecked():
            selected_standards.append('gdpr_IV')

        process_folder(folder_path, selected_standards, document_title, logo_path)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())