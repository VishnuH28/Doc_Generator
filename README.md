# Document Generator

A Python-based tool that generates personalized PDF and Word documents from Excel data. This tool can automatically create formatted documents for each row in your Excel sheet, making it perfect for generating employee documents, certificates, or any other type of personalized documentation.

## Features

- **Excel Data Processing**: Reads and validates Excel files with employee information
- **Multiple Output Formats**: Generate documents in both PDF and Word formats
- **Company Logo Support**: Option to add company logo to generated documents
- **User Interface**: Streamlit-based web interface for easy file upload and processing
- **Error Handling**: Robust error checking and logging
- **Customizable Output**: Generated documents are automatically named and organized

## Prerequisites

- Python 3.7 or higher
- Required Python packages:
```bash
pandas
openpyxl
python-docx
fpdf
Pillow
streamlit
```

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/document-generator.git
cd document-generator
```

2. Install required packages:
```bash
pip install pandas openpyxl python-docx fpdf Pillow streamlit
```

## Usage

### Using the Streamlit Interface

1. Start the web interface:
```bash
streamlit run document_generator.py
```

2. Access the interface in your web browser (typically http://localhost:8501)

3. Upload your Excel file through the web interface

4. (Optional) Upload a company logo

5. Select your desired output format (PDF, Word, or both)

6. Click "Generate Documents" to process the files

### Excel File Format

Your Excel file should contain the following columns:
- Name
- Email
- Company Name
- Position
- Joining Date

Example Excel format:
| Name | Email | Company Name | Position | Joining Date |
|------|--------|--------------|-----------|--------------|
| John Doe | john.doe@example.com | Tech Corp | Software Engineer | 2024-01-15 |
| Jane Smith | jane.smith@example.com | Tech Corp | Product Manager | 2024-02-01 |

### Output Format

Generated documents will be saved in the `output` directory with the following naming convention:
```
EmployeeName_CompanyName.pdf
EmployeeName_CompanyName.docx
```

## Project Structure

```
document-generator/
│
├── document_generator.py    # Main application file
├── README.md               # Documentation
├── requirements.txt        # Package dependencies
├── output/                 # Generated documents
└── sample_data.xlsx        # Example Excel file
```

## Generated Document Features

### PDF Documents
- Company logo (if provided)
- Professional formatting
- Employee details in a structured layout
- Company branding elements

### Word Documents
- Company logo (if provided)
- Editable format
- Professional template
- Consistent formatting

## Error Handling

The application includes comprehensive error handling for:
- Missing or invalid Excel files
- Incorrect data formats
- Missing required columns
- File system errors
- Image processing errors

## Logging

All operations are logged with timestamps and error details for troubleshooting.

## Limitations

- Supported image formats for logo: PNG, JPG, JPEG
- Excel file must contain all required columns
- Company logo should be of reasonable size (<5MB)

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## Support

For support and questions, please open an issue in the GitHub repository.