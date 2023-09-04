import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch, cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import csv
from PIL import Image as PILImage, ImageOps, ImageChops
from datetime import datetime

class PDFGeneratorApp:
    
    def __init__(self, root):
        self.root = root
        self.root.title("CIS Benchmark | Made with ❤️ by Steven Olsen")
        
        # CSV Selection
        self.csv_frame = ttk.LabelFrame(root, text="CSV Selection")
        self.csv_frame.grid(row=0, column=0, padx=10, pady=5, sticky="ew")
        self.csv_path = tk.StringVar()
        ttk.Label(self.csv_frame, text="CSV Path:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(self.csv_frame, textvariable=self.csv_path).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.csv_frame, text="Browse", command=self.browse_csv).grid(row=0, column=2, padx=5, pady=5)
        
        # Logo Selection
        self.logo_frame = ttk.LabelFrame(root, text="Logo Selection")
        self.logo_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        self.logo_path = tk.StringVar()
        ttk.Label(self.logo_frame, text="Logo Path:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(self.logo_frame, textvariable=self.logo_path).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.logo_frame, text="Browse", command=self.browse_logo).grid(row=0, column=2, padx=5, pady=5)
        
        # Additional Info
        self.info_frame = ttk.LabelFrame(root, text="Additional Info")
        self.info_frame.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        self.report_title = tk.StringVar()
        ttk.Label(self.info_frame, text="Report Title:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(self.info_frame, textvariable=self.report_title).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.report_url = tk.StringVar()
        ttk.Label(self.info_frame, text="Footer URL:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(self.info_frame, textvariable=self.report_url).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.output_path = tk.StringVar()
        ttk.Label(self.info_frame, text="Output Path:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(self.info_frame, textvariable=self.output_path).grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.info_frame, text="Browse", command=self.browse_output).grid(row=2, column=2, padx=5, pady=5)
        
        # Generate PDF
        self.gen_frame = ttk.LabelFrame(root, text="Generate PDF")
        self.gen_frame.grid(row=3, column=0, padx=10, pady=5, sticky="ew")
        self.status_label = ttk.Label(self.gen_frame, text="Status: Awaiting Input")
        self.status_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        ttk.Button(self.gen_frame, text="Generate PDF", command=self.generate_pdf).grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(self.gen_frame, text="Exit", command=root.quit).grid(row=1, column=1, padx=5, pady=5)
        
    def browse_csv(self):
        file_path = filedialog.askopenfilename(title="Select CSV", filetypes=[("CSV Files", "*.csv")])
        self.csv_path.set(file_path)
    
    def browse_logo(self):
        file_path = filedialog.askopenfilename(title="Select Logo", filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
        self.logo_path.set(file_path)
        
    def browse_output(self):
        file_path = filedialog.asksaveasfilename(title="Save PDF As", defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        self.output_path.set(file_path)

    def generate_pdf(self):
        success = generate_pdf_from_csv(self.csv_path.get(), self.logo_path.get(), self.report_title.get(), self.report_url.get(), self.output_path.get())
        if success:
            self.status_label.config(text="Status: PDF Generation Successful!")
        else:
            self.status_label.config(text="Status: PDF Generation Failed!")

def generate_pdf_from_csv(csv_path, logo_path, report_title, footer_url, output_path):
    doc = SimpleDocTemplate(output_path, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()
    title_style = styles['Heading1']
    desc_style = styles['BodyText']
    ref_style = ParagraphStyle('ReferenceStyle', parent=styles['BodyText'], textColor=colors.blue)
    rationale_style = styles['Italic']
    
    elements.append(Paragraph(report_title, title_style))
    
    if logo_path:
        try:
            pil_img = PILImage.open(logo_path)
            w, h = pil_img.size
            aspect_ratio = w / h
            new_width = min(150, w)
            new_height = new_width / aspect_ratio
            pil_img = pil_img.resize((int(new_width), int(new_height)), PILImage.LANCZOS)
            logo_path_resized = "temp_resized_logo.png"
            pil_img.save(logo_path_resized, "PNG")
            logo = Image(logo_path_resized)
            elements.append(logo)
        except Exception as e:
            print(f"Error processing logo: {e}")
    
    with open(csv_path, 'r', newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            elements.append(Paragraph(f"Failed: {row['Title']}", title_style))
            elements.append(Paragraph(f"Description: {row['Description']}", desc_style))
            elements.append(Paragraph(f"Reference: {row['References']}", ref_style))
            elements.append(Spacer(1, 12))  
            elements.append(Paragraph(f"Rationale: {row['Rationale']}", rationale_style))
            elements.append(Spacer(1, inch))

    def footer(canvas, doc):
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        footer_text = f"{footer_url} - Generated on: {timestamp}"
        canvas.saveState()
        canvas.drawString(inch, 0.75 * inch, footer_text)
        canvas.drawRightString(10 * inch - 2*inch, 0.75 * inch, f"Page {canvas.getPageNumber()}")
        canvas.restoreState()
        
    doc.build(elements, onFirstPage=footer, onLaterPages=footer)
    
    return True

root = tk.Tk()
root.geometry("450x350")  # 800 pixels wide and 500 pixels tall
app = PDFGeneratorApp(root)
root.mainloop()