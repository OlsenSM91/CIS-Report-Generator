# CIS Benchmark PDF Generator

This tool provides a graphical user interface to generate PDF reports from CSV files. The tool is primarily designed to be used with CIS Benchmark data, allowing users to insert logos and provide additional information to customize the output report.

## Features:

- **CSV Selection**: Allows users to browse and select a CSV file as input.
- **Logo Selection**: Users can select a logo to be included in the PDF report.
- **Additional Info**: Users can provide a custom title for the report, a footer URL, and specify an output path for the generated PDF.
- **PDF Generation**: With a click of a button, generate a PDF report based on the provided CSV file.

## How to Use:

1. **Launch the Application**:
   - Run the application to open the GUI.
   
2. **Select a CSV File**:
   - Click on the "Browse" button under "CSV Selection" to choose a CSV file.

3. **Choose a Logo (Optional)**:
   - Click on the "Browse" button under "Logo Selection" to select an image file. The tool supports PNG, JPG, and JPEG formats.

4. **Provide Additional Information**:
   - Fill in the "Report Title" field to set a custom title for your report.
   - Specify a "Footer URL" if you wish to have a URL displayed at the bottom of each page.
   - Choose an "Output Path" by clicking on the "Browse" button. This is where the generated PDF will be saved.

5. **Generate PDF**:
   - Click on the "Generate PDF" button. The status will be updated once the PDF generation is complete.

6. **Exit**:
   - Click on the "Exit" button to close the application.

## Dependencies:

- `tkinter`: For the graphical user interface.
- `reportlab`: For generating PDF reports.
- `csv`: For reading CSV files.
- `PIL` from `Pillow`: For processing image files.

To install the dependencies (if running the source script):
`pip install tkinter reportlab Pillow`

## Notes:

- Make sure the CSV file has the following columns: `Title`, `Description`, `References`, and `Rationale`.
- The logo image is automatically resized to a maximum width of 150 pixels while maintaining its aspect ratio.

## Author:

Made with ❤️ by Steven Olsen
