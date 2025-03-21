# CertifyLite: Certificate Generator

Welcome to **CertifyLite**, a streamlined Python tool that makes certificate creation a breeze. Got a list of names and a PowerPoint template? This script pulls names from an Excel file, updates your template, and delivers crisp PDFs—no manual hassle required. Ideal for awards, events, or any bulk certificate task.

## What It Does
- Reads names from an Excel file (`names.xlsx`).
- Replaces a placeholder in a PowerPoint template (`cert.pptx`).
- Converts each certificate to PDF and removes temporary files.
- Saves the results in a `Certificates` folder.

## Requirements
- **Python 3.12**
- **Libraries**:
  - `python-pptx` (for editing PowerPoint files)
  - `pandas` and `openpyxl` (for reading Excel)
  - `comtypes` (for PDF conversion, Windows only)
- **Microsoft PowerPoint** installed (for PDF export)
- An Excel file with a "NAME" column
- A PowerPoint template with "Arnab Aich" as the placeholder text

## How to Use
1. **Install Dependencies**:
   ```bash
   pip install python-pptx pandas openpyxl comtypes
   ```
2. **Set Up Files**:
   - Place `names.xlsx` and `cert.pptx` in `E:/Projects/certifylite/` (or your project folder).
   - Update the paths in `main.py` if needed.
3. **Run the Script**:
   ```bash
   & "C:/Users/YourName/AppData/Local/Programs/Python/Python312/python.exe" e:/Projects/certifylite/main.py
   ```
4. **Check Output**:
   - PDFs will land in `E:/Projects/certifylite/Certificates/`.

## Code Overview
```python
# main.py (simplified)
for name in names_list:
    prs = Presentation("cert.pptx")
    # Replace "Arnab Aich" with each name
    # Save as PPTX, convert to PDF, delete PPTX
    print(f"PDF saved: Certificates/Certificate_{name}.pdf")
```

## Notes
- **Windows Only**: PDF conversion uses PowerPoint’s COM interface.
- **Placeholder**: If your template’s placeholder isn’t "Arnab Aich", adjust the code.
- **Troubleshooting**: Double-check file paths and ensure PowerPoint is installed.

## Contributing
Got suggestions or fixes? Fork the repo, make your changes, and send a pull request. Let’s improve **CertifyLite** together!

## Credits
Created by Obidur Rahman with a nudge from xAI’s Grok 3. Built to simplify life, one certificate at a time.
