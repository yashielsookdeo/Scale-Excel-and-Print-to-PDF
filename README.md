# Excel to PDF Converter

This script converts all Excel files in a specified input folder to PDF files in a specified output folder. Each Excel sheet is scaled to fit on one A4 page.

## Requirements

- Python 3.x
- Windows OS
- Microsoft Excel installed

## Installation

1. **Clone the repository or download the script.**

2. **Install the required Python libraries:**

    ```bash
    pip install pandas openpyxl pywin32
    ```

## Usage

1. **Modify the script:**

    Open the script and set the `input_folder` and `output_folder` variables to the paths of your input and output folders, respectively.

    ```python
    input_folder = 'path_to_your_input_folder'
    output_folder = 'path_to_your_output_folder'
    ```

2. **Run the script:**

    ```bash
    python convert_excel_to_pdf.py
    ```

## Script Details

### `scale_excel_and_print_to_pdf(excel_file, pdf_output_file)`

This function opens an Excel file, scales each sheet to fit on one A4 page, and exports the file as a PDF.

- **Parameters:**
  - `excel_file`: Path to the Excel file.
  - `pdf_output_file`: Path to the output PDF file.

### `convert_folder(input_folder, output_folder)`

This function retrieves all Excel files in the input folder and converts each to a PDF, saving them in the output folder.

- **Parameters:**
  - `input_folder`: Path to the folder containing Excel files.
  - `output_folder`: Path to the folder where PDF files will be saved.

### Main Execution Block

- Ensures the output folder exists.
- Calls `convert_folder` to process all Excel files in the input folder.

```python
if __name__ == "__main__":
    input_folder = 'path_to_your_input_folder'
    output_folder = 'path_to_your_output_folder'

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Convert all Excel files in the input folder to PDFs in the output folder
    convert_folder(input_folder, output_folder)

```
## Notes
- The script uses the win32com.client library, which requires Windows and Microsoft Excel to be installed.
- The script sets the paper size to A4 and orientation to portrait.
- All content in each Excel sheet will be scaled to fit on one A4 page.