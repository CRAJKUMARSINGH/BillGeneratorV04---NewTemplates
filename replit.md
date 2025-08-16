# Bill Generator Application

## Overview

This is a Streamlit-based government contractor bill generator application designed to create standardized government contractor bills and related documentation. The application processes bill data through HTML templates using Jinja2 templating and converts them to PDF format for official documentation purposes.

## User Preferences

Preferred communication style: Simple, everyday language.

## System Architecture

### Frontend Architecture
- **Framework**: Streamlit web application framework
- **UI Design**: Single-page application with wide layout configuration
- **Branding**: Dynamic logo resolution system that checks local files first, then falls back to remote GitHub repository URLs
- **Page Configuration**: Custom page title "Bill Generator" with dynamic favicon support

### Backend Architecture
- **Template Engine**: Jinja2 for HTML template rendering
- **PDF Generation**: Multiple PDF generation libraries (pdfkit, xhtml2pdf) for document conversion
- **Document Processing**: Support for both DOCX and PDF output formats
- **Data Processing**: Pandas for spreadsheet data handling and manipulation
- **File Management**: Temporary file handling with zipfile support for batch operations

### Document Generation System
- **Template Structure**: Modular HTML templates for different bill sections:
  - `first_page.html`: Main contractor bill with itemized work details
  - `certificate_ii.html`: Work certification and signatures
  - `certificate_iii.html`: Payment memorandum
  - `deviation_statement.html`: Work deviations documentation
  - `extra_items.html`: Additional work items
  - `note_sheet.html`: Bill scrutiny documentation
- **Styling**: Consistent CSS styling with A4 page formatting (both portrait and landscape)
- **Data Binding**: Dynamic content injection through Jinja2 template variables

### Data Processing Pipeline
- **Input Processing**: Excel/CSV file upload and parsing
- **Data Transformation**: Automatic calculation of totals, amounts, and numeric conversions
- **Text Processing**: Number-to-words conversion using num2words library
- **Date Formatting**: Automatic date format conversion and validation

### File Management System
- **Asset Management**: Organized attachment system with timestamped file naming
- **Template Directory**: Centralized template storage in `/templates` directory
- **Output Generation**: PDF compilation with proper page sequencing
- **Archive Support**: Compression and bundling capabilities for document packages

## External Dependencies

### Core Python Libraries
- **streamlit**: Web application framework for the user interface
- **pandas**: Data manipulation and analysis for spreadsheet processing
- **jinja2**: Template engine for HTML document generation
- **num2words**: Number-to-text conversion for financial amounts

### PDF Generation Stack
- **pdfkit**: Primary PDF generation from HTML (requires wkhtmltopdf system dependency)
- **xhtml2pdf**: Alternative PDF generation library
- **pypdf**: PDF manipulation and merging capabilities

### Document Processing
- **python-docx**: Microsoft Word document generation and formatting
- **openpyxl**: Excel file reading and writing (via pandas)

### System Integration
- **requests**: HTTP client for remote logo/asset fetching
- **subprocess**: System command execution for external tools
- **tempfile**: Temporary file management for processing pipelines

### File Handling
- **zipfile**: Archive creation for document bundles
- **tarfile**: Alternative archive format support
- **io**: In-memory file operations

### Platform Dependencies
- **wkhtmltopdf**: System-level PDF rendering engine (required for pdfkit)
- Cross-platform support for Windows, macOS, and Linux environments

### Remote Resources
- GitHub repository integration for logo and asset fetching with fallback mechanisms
- Multiple URL candidates for robust asset loading