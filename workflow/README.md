# Bonus Bot - Document Generator

A Windows desktop application that generates personalized bonus documents from Excel data and Word templates.

## Features

- 🖥️ **User-friendly GUI** - No technical knowledge required
- 📊 **Excel Integration** - Load employee data from .xlsx files
- 📄 **Template Processing** - Use custom Word templates with placeholders
- 📑 **PDF Generation** - Automatically converts to PDF format
- 📁 **Organized Output** - Creates folders by month/year
- ⚡ **Fast Processing** - Handles multiple employees quickly

## Download

Download the latest Windows executable from the [Releases](../../releases) page.

**No Python installation required!** Just download and run the .exe file.

## How to Use

1. **Run** `BonusBot.exe`
2. **Select Excel file** with employee data
3. **Select Word template** with placeholders
4. **Choose output folder** for generated documents
5. **Click "Generate Documents"** and wait for completion

## Excel File Format

Your Excel file should contain these columns:

| Column Name | Description | Required |
|-------------|-------------|----------|
| `full_name` | Employee full name | ✅ Yes |
| `first_name` | Employee first name | ❌ No |
| `job_description` | Job title | ❌ No |
| `branch` | Branch/Department | ❌ No |
| `branch_grams_0000` | Branch grams | ❌ No |
| `personal_grams_0000` | Personal grams | ❌ No |
| `dinar` | Dinar amount | ✅ Yes |
| `value_18ct` | 18ct value | ❌ No |
| `value_21ct` | 21ct value | ❌ No |

## Template Placeholders

Use these placeholders in your Word template:

- `{{current_date}}` - Today's date (DD/MM/YYYY)
- `{{full_name}}` - Employee full name
- `{{first_name}}` - Employee first name
- `{{job_description}}` - Job description
- `{{branch}}` - Branch name
- `{{branch_grams_0000}}` - Branch grams amount
- `{{personal_grams_0000}}` - Personal grams amount
- `{{dinar}}` - Dinar amount
- `{{value_18ct}}` - 18ct gold value
- `{{value_21ct}}` - 21ct gold value

## System Requirements

- **Windows 10 or later**
- **Microsoft Word** (for best PDF conversion) or **LibreOffice** (free alternative)

## Troubleshooting

### PDF Conversion Issues
If PDF conversion fails:
1. Install [LibreOffice](https://www.libreoffice.org/download/download/) (free)
2. Or ensure Microsoft Word is properly installed

### Missing Data
- Check that your Excel file has the required columns (`full_name`, `dinar`)
- Ensure template placeholders match exactly (including double curly braces)

### File Permissions
- Run as Administrator if you get permission errors
- Ensure output folder is writable

## Development

This app is built with:
- **Python 3.11**
- **Tkinter** (GUI)
- **pandas** (Excel processing)
- **python-docx** (Word document handling)
- **PyInstaller** (Executable creation)

### Building from Source

1. Install Python 3.11+
2. Install dependencies: `pip install -r requirements.txt`
3. Build executable: `pyinstaller bonus_bot.spec`

## License

MIT License - See LICENSE file for details.

## Support

For issues or questions, please create an issue in this repository.