import os
import glob
from win32com import client

def scale_excel_and_print_to_pdf(excel_file, pdf_output_file):
    # Load the Excel workbook
    excel = client.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Open(os.path.abspath(excel_file))

    # Iterate through all sheets and set the print area and fit to page settings
    for sheet in workbook.Sheets:
        sheet.PageSetup.PrintArea = sheet.UsedRange.Address
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = 1
        sheet.PageSetup.PaperSize = 9  # xlPaperA4
        sheet.PageSetup.Orientation = 2  # xlPortrait
        sheet.PageSetup.Zoom = False  # Disable automatic scaling

    # Export the workbook to a PDF file
    workbook.ExportAsFixedFormat(0, os.path.abspath(pdf_output_file))

    # Close the workbook and quit Excel
    workbook.Close(False)
    excel.Quit()

def convert_folder(input_folder, output_folder):
    # Get all Excel files in the input folder
    excel_files = glob.glob(os.path.join(input_folder, "*.xlsx"))

    # Convert each Excel file to a PDF
    for excel_file in excel_files:
        # Define the PDF output file path
        pdf_output_file = os.path.join(
            output_folder,
            os.path.splitext(os.path.basename(excel_file))[0] + ".pdf"
        )

        # Convert the Excel file to a PDF
        scale_excel_and_print_to_pdf(excel_file, pdf_output_file)
        print(f"Converted {excel_file} to {pdf_output_file}")

if __name__ == "__main__":
    input_folder = 'Input/'
    output_folder = 'Output/'

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Convert all Excel files in the input folder to PDFs in the output folder
    convert_folder(input_folder, output_folder)