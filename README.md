# JSON to Excel Converter

A GUI application that converts JSON arrays to Excel files. The application allows users to select a JSON file, choose which array to convert, and save the result as an Excel file with keys as columns.

**Author:** Michael Dehne  
**License:** MIT License  
**Version:** 1.0.0

## Features

- **File Selection**: Browse and select JSON files
- **Array Detection**: Automatically finds all arrays in the JSON structure
- **Array Selection**: Choose which array to convert to Excel
- **Excel Export**: Convert selected arrays to Excel format with keys as columns
- **Auto-Open**: Automatically opens the converted Excel file for immediate review
- **User-Friendly GUI**: Simple and intuitive interface

## Installation

1. Make sure you have Python 3.7+ installed
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

1. Run the script:
```bash
python json_to_excel.py
```

2. **Step 1**: Click "Browse" to select a JSON file
3. **Step 2**: Choose which array to convert from the dropdown
4. **Step 3**: Set the output location:
   - Click "Browse" to select the save folder
   - Enter a filename (auto-generated based on selected array)
5. **Step 4**: Click "Convert to Excel" to process the conversion
6. **Auto-Open**: The Excel file automatically opens in your default spreadsheet application

## Requirements

- Python 3.7+
- pandas
- openpyxl
- tkinter (usually comes with Python)

## How it Works

1. **JSON Parsing**: The script loads and parses the JSON file
2. **Array Detection**: It recursively searches for all arrays in the JSON structure
3. **Array Selection**: Users can choose which array to convert
4. **Output Configuration**: Users set the save folder and filename before conversion
5. **Excel Conversion**: The selected array is converted to a pandas DataFrame and saved as an Excel file
6. **Column Mapping**: Object keys become Excel columns
7. **Auto-Open**: The converted Excel file automatically opens for immediate review

## Example JSON Structure

The script can handle various JSON structures:

```json
{
  "users": [
    {"name": "John", "age": 30, "city": "New York"},
    {"name": "Jane", "age": 25, "city": "Los Angeles"}
  ],
  "products": [
    {"id": 1, "name": "Product A", "price": 100},
    {"id": 2, "name": "Product B", "price": 200}
  ]
}
```

In this example, the script would detect two arrays: "users" and "products", and you could choose which one to convert to Excel.

## Notes

- The script only converts arrays that contain objects (dictionaries)
- Each object's keys become columns in the Excel file
- The script supports nested JSON structures
- Excel files are saved with .xlsx extension

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Author

**Michael Dehne** - [GitHub Profile](https://github.com/micdehne)

---

*If you find this tool useful, please consider giving it a star on GitHub!*
