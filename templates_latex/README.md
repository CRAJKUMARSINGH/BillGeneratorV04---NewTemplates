# Government Templates - LaTeX Files

## About These Templates
These LaTeX files were converted from governmental HTML templates for professional document generation.
Each file is ready to compile into a high-quality PDF document with government-standard formatting.

## Prerequisites
Install a LaTeX distribution on your system:
- **Windows**: MiKTeX (https://miktex.org/) or TeX Live (https://www.tug.org/texlive/)
- **macOS**: MacTeX (https://www.tug.org/mactex/)
- **Linux**: TeX Live (install via package manager: `sudo apt-get install texlive-full`)

## Required LaTeX Packages
The templates use these LaTeX packages (usually included in full distributions):
- geometry (page layout and margins)
- array (enhanced table formatting)
- booktabs (professional table rules)
- longtable (tables that span multiple pages)
- xcolor (color support)
- fontenc, inputenc (font and input encoding)
- helvet (Helvetica font family)
- fancyhdr (headers and footers)
- amsmath, amsfonts (mathematical symbols)

## Compilation Instructions

### Method 1: Command Line
For each .tex file, run these commands in your terminal:
```bash
pdflatex filename.tex
pdflatex filename.tex  # Run twice for proper cross-references and formatting
```

### Method 2: LaTeX Editor
1. Open the .tex file in a LaTeX editor (TeXmaker, TeXstudio, Overleaf, etc.)
2. Set the compiler to pdflatex
3. Click the "Build" or "Compile" button
4. Compile twice for best results

## Template Descriptions
- **bill_template.tex**: Professional billing template with tabular data
- **certificate_ii.tex**: Official government certificate (Type II)
- **certificate_iii.tex**: Official government certificate (Type III)
- **deviation_statement.tex**: Project deviation documentation
- **extra_items.tex**: Additional items and supplementary information
- **first_page.tex**: Cover page template for government documents
- **note_sheet.tex**: Note sheet for official communications

## Output
Compilation will generate PDF files with:
- Professional government-standard formatting
- Proper margins and spacing
- Government-appropriate fonts (Helvetica/Arial family)
- Professional table layouts
- Headers and footers with document information

## Customization
You can modify the .tex files to:
- Add your specific content
- Adjust formatting as needed
- Include organization logos or letterheads
- Modify colors and styling

## Troubleshooting
- If compilation fails, ensure all required packages are installed
- Check that your LaTeX distribution is up to date
- Some templates use landscape orientation for better table display
- Error messages usually indicate missing packages or syntax issues

## Support
These templates were generated automatically from HTML sources.
For LaTeX-specific issues, consult LaTeX documentation or community forums.

---
Generated on: /home/runner/workspace
Conversion tool: Direct Government Template to LaTeX Converter
