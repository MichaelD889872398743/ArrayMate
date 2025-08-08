# JSON Data Converter

A GUI application that converts JSON arrays to multiple formats (Excel, CSV, JSON). The application allows users to select a JSON file or paste JSON directly, choose which array to convert, and save the result in various formats with keys as columns.

**Author:** Michael Dehne  
**License:** MIT License  
**Version:** 1.0.0

## Features

- **Multiple JSON Sources**: Browse JSON files OR paste JSON directly (e.g., from Postman, API responses)
- **Array Detection**: Automatically finds all arrays in the JSON structure
- **Array Selection**: Choose which array to convert
- **Multiple Output Formats**: Export to Excel (.xlsx), CSV (.csv), or JSON (.json)
- **Auto-Open**: Automatically opens converted files or shows file location
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

2. **Step 1**: Choose JSON source:
   - **Option A**: Click "Browse" to select a JSON file
   - **Option B**: Click "Open JSON Input" to paste JSON directly (e.g., from Postman)
3. **Step 2**: Choose which array to convert from the dropdown
4. **Step 3**: Set the output location:
   - Select output format (Excel, CSV, or JSON)
   - Click "Browse" to select the save folder
   - Enter a filename (auto-generated based on selected array)
5. **Step 4**: Click "Convert to File" to process the conversion
6. **Auto-Open**: Files automatically open or show file location

## Requirements

- Python 3.7+
- pandas
- openpyxl
- tkinter (usually comes with Python)

## How it Works

1. **JSON Input**: Load JSON from file or paste directly from clipboard
2. **Array Detection**: It recursively searches for all arrays in the JSON structure
3. **Array Selection**: Users can choose which array to convert
4. **Format Selection**: Choose output format (Excel, CSV, or JSON)
5. **Output Configuration**: Users set the save folder and filename before conversion
6. **Data Conversion**: The selected array is converted to a pandas DataFrame
7. **File Export**: Save in the chosen format with proper formatting
8. **Auto-Open**: Files automatically open or show file location

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
