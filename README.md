# Form Generator

A Python-based tool that automatically generates personalized PDF forms by filling in template PDFs with data from Excel spreadsheets. This project is particularly useful for creating consent forms, certificates, or any document that requires bulk personalization.

## Features

- **Bulk PDF Generation**: Automatically fills multiple PDF forms from Excel data
- **Multi-page Support**: Handles complex PDF templates with multiple pages
- **Signature Integration**: Generates and embeds handwritten-style signature images
- **Custom Positioning**: Precise control over text placement on PDF pages
- **Date Formatting**: Supports multiple date formats (MM-DD-YYYY, Indian format)
- **Batch Processing**: Processes multiple records efficiently

## Project Structure

```
Form_generator/
├── generate_forms.py       # Main script for PDF form generation
├── Signatures_create.py    # Script to generate signature images
├── p1.xlsx                # Excel file with participant data
├── sample.pdf             # Template PDF form
├── Dancing-Script.ttf     # Font file for signatures
├── filled_forms/          # Output directory for generated PDFs
└── signatures/            # Directory for generated signature images
```

## Prerequisites

Make sure you have Python installed on your system. This project requires the following Python packages:

```bash
pip install pandas PyPDF2 reportlab pillow openpyxl
```

## Required Files

1. **Excel File (`p1.xlsx`)**: Should contain columns:
   - `Name`: Participant's full name
   - `Email`: Email address
   - `Date`: Date in any standard format
   - `Unique Batch ID`: Unique identifier for each form

2. **Template PDF (`sample.pdf`)**: A blank PDF form that will be filled
3. **Font File (`Dancing-Script.ttf`)**: Handwriting-style font for signatures

## Installation

1. Clone this repository:
```bash
git clone https://github.com/vishal-04-singh/Form_generator.git
cd Form_generator
```

2. Install required dependencies:
```bash
pip install pandas PyPDF2 reportlab pillow openpyxl
```

3. Download the Dancing Script font:
   - Download `Dancing-Script.ttf` from Google Fonts
   - Place it in the project root directory

## Usage

### Step 1: Generate Signature Images

First, create signature images for all participants:

```bash
python Signatures_create.py
```

This will:
- Read names from `p1.xlsx`
- Generate handwritten-style signature images
- Save them in the `signatures/` folder

### Step 2: Generate Filled Forms

Then, generate the personalized PDF forms:

```bash
python generate_forms.py
```

This will:
- Read participant data from Excel
- Fill the template PDF with personalized information
- Add signatures where specified
- Save completed forms in the `filled_forms/` folder

## Customization

### Adjusting Text Positions

In `generate_forms.py`, modify the coordinates in the `create_overlay()` function:

```python
# Page 1 - Name position
can.drawString(445, 546.5, name)  # Adjust x, y coordinates

# Page 3 - Multiple fields
can.drawString(184, 476.5, name)   # Name position
can.drawString(60, 446.5, email)   # Email position
can.drawString(60, 388.5, date)    # Date position
```

### Signature Settings

Modify signature appearance in `Signatures_create.py`:

```python
font_size = 48  # Adjust font size
img = Image.new('RGBA', (400, 100), (255, 255, 255, 0))  # Image dimensions
```

### Date Format

The script automatically formats dates to MM-DD-YYYY format. To change this, modify:

```python
date = pd.to_datetime(date_raw).strftime('%m-%d-%Y')  # Change format here
```

## Output

- **Individual PDFs**: Each participant gets a personalized PDF named with their Batch ID
- **Signature Images**: PNG files for each participant's signature
- **Organized Structure**: All outputs are saved in respective folders

## Troubleshooting

### Common Issues

1. **Font not found error**: Ensure `Dancing-Script.ttf` is in the project directory
2. **PDF positioning issues**: Adjust coordinates in the `create_overlay()` function
3. **Excel reading errors**: Verify column names match exactly in your Excel file
4. **Missing signatures**: Check if signature images exist in the `signatures/` folder

### Error Messages

- `⚠️ Could not place signature for [Name]`: Signature image not found or corrupted
- `✅ All consent forms generated successfully`: Process completed without errors

## Requirements.txt

```
pandas>=1.3.0
PyPDF2>=2.0.0
reportlab>=3.6.0
Pillow>=8.0.0
openpyxl>=3.0.0
```

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Commit your changes: `git commit -am 'Add some feature'`
4. Push to the branch: `git push origin feature-name`
5. Submit a pull request

## License

This project is open source and available under the [MIT License](LICENSE).

## Contact

Created by [@vishal-04-singh](https://github.com/vishal-04-singh)

---

**Note**: This tool is designed for legitimate document generation purposes. Ensure you have proper authorization before generating documents on behalf of others.