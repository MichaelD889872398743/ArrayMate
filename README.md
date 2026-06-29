# ArrayMate

ArrayMate is a local desktop app for turning JSON arrays into table files. It is built for the common "I got API/report data as JSON, but need to send someone an Excel or CSV file" workflow.

It can load JSON from a file or from pasted text, detect table-like arrays, preview the selected data, apply quick transforms, and export the result.

**Author:** Michael Dehne  
**License:** MIT License  
**Version:** 2.0.1

## Features

- **Local desktop app**: Runs on your machine; no web upload step.
- **No-install portable release**: Built as a portable Windows app folder.
- **File or paste input**: Load a JSON file, or paste JSON directly from tools like Postman.
- **Automatic array detection**: Finds top-level and nested arrays.
- **Table preview**: Shows detected columns and sample rows before export.
- **Nested array handling**: Select nested candidates such as `orders[*].items`, include parent metadata, or unfold nested arrays from a parent table.
- **Transform options**: Stringify all values, stringify spreadsheet formulas, set per-column types, and perform per-column find/replace.
- **Exports**: Excel `.xlsx`, CSV `.csv`, and JSON `.json`.
- **Modern UI**: PySide6/Qt interface with app and tray icons.

## Download

1. Go to the [Releases](https://github.com/MichaelD889872398743/ArrayMate/releases) page.
2. Download the latest `ArrayMate-v*-Windows-PortablePython.zip`.
3. Extract the zip file.
4. Run `ArrayMate/Run ArrayMate.bat`.

### Windows Security Notice

ArrayMate is currently unsigned, so Windows may show an "Unknown publisher" or "Windows protected your PC" warning the first time you run it.

This is expected for a new open-source tool without a code-signing certificate. If Windows SmartScreen appears, click **More info**, then **Run anyway**.

The release is built from the source code in this repository, so you can inspect or build it yourself.

## Run From Source

Use Python 3.10 or newer.

```bash
pip install -r requirements.txt
python app.py
```

## Basic Workflow

1. Start ArrayMate.
2. Paste JSON into the JSON input, or click **Load JSON File**.
3. Pick an array from the parsed structure pane.
4. Review the table preview.
5. Optional: choose transform options.
6. Pick the output format, file name, and save folder.
7. Click **Convert to File**.

Loaded JSON files are copied into the JSON input. After that, pasted JSON and file-based JSON use the same parse path.

## Nested Arrays

ArrayMate lists arrays by path. For repeated nested arrays it groups compatible paths with a wildcard:

- `users`: a top-level array.
- `orders`: a top-level order table.
- `orders[*].items`: all `items` arrays found inside the `orders` records.
- `orders[*].items[*].descriptions`: nested arrays below order items.

For nested arrays, you can export the nested table on its own or include parent metadata where useful. For parent tables, the unfold option lets you expand a nested child array into the current preview.

## Example

```json
{
  "orders": [
    {
      "order_id": "ORD001",
      "status": "Completed",
      "items": [
        { "product_id": "P001", "quantity": 1, "price": 999.99 },
        { "product_id": "P002", "quantity": 1, "price": 29.99 }
      ]
    }
  ]
}
```

ArrayMate can detect both:

- `orders`
- `orders[*].items`

If you export `orders[*].items` with parent metadata, the item rows can include information from the order they came from.

## Transform Options

Quick options:

- **Stringify everything**: Export every value as text.
- **Stringify formulas**: Treat spreadsheet formula-looking values as text.
- **Include parent metadata**: Add parent object fields to nested array rows where supported.

Advanced column actions:

- Set a selected column to text, number, integer, or boolean.
- Replace text inside one selected column.

These transforms are applied to the exported data and preview.

## Building Portable Releases

Install runtime and build dependencies:

```bash
pip install -r requirements.txt
pip install -r requirements-build.txt
```

For managed Windows laptops, build the portable Python package:

```bash
build.bat
.\build.ps1
python build_exe.py portable-python
```

This creates `release/ArrayMate/Run ArrayMate.bat`. The launcher starts ArrayMate through the bundled `pythonw.exe` runtime instead of a generated `ArrayMate.exe`.

This is intentional. Some company devices block unsigned generated executables through Defender Exploit Guard or similar policy. Running through a trusted Python runtime can be easier to allow while still keeping the app portable and installation-free.

There is also a PyInstaller target:

```bash
python build_exe.py pyinstaller
```

The PyInstaller build uses `ArrayMate.spec` as the canonical configuration. It expects these files to exist:

- `app.py`
- `ArrayMate.spec`
- `version.txt`
- `icon.ico`
- `assets/arraymate_icon.png`
- `assets/arraymate_tray_icon.png`

Successful builds place the portable app in `release/ArrayMate/` and create `ArrayMate-v*-Windows-*.zip`.

## Publishing A GitHub Release

Releases are created from version tags. Generated zip files should not be committed to git.

1. Update the version in:
   - `arraymate/__init__.py`
   - `setup.py`
   - `version.txt`
   - `README.md`
   - `CHANGELOG.md`
2. Commit the version change.
3. Create and push a matching tag:

```bash
git tag v1.0.11
git push origin rewrite/v2
git push origin v1.0.11
```

The GitHub Actions workflow builds `ArrayMate-v*-Windows-PortablePython.zip` and attaches it to a new release under [Releases](https://github.com/MichaelD889872398743/ArrayMate/releases).

## Project Structure

- `app.py`: Application entry point.
- `arraymate/core.py`: JSON discovery, table extraction, transforms, and export logic.
- `arraymate/service.py`: UI-independent workflow layer.
- `arraymate/qt_desktop.py`: PySide6 desktop UI.
- `tests/`: Unit tests for core and service behavior.
- `assets/`: UI mockup and icon assets.

## Notes

- The main use case is converting arrays of objects into tables.
- Object keys become table columns.
- Empty arrays can be detected, but they do not produce table rows.
- Spreadsheet formula protection is optional because some users intentionally export formula values.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.

## Author

**Michael Dehne** - [GitHub Profile](https://github.com/MichaelD889872398743)
