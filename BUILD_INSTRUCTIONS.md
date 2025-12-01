# Building the Executable

This guide explains how to convert the Python project into an executable (.exe) file.

## Prerequisites

1. **Python 3.9+** installed
2. **Virtual environment** activated (recommended)
3. **All dependencies** installed:
   ```bash
   pip install -r requirements.txt
   ```
4. **Playwright browsers** installed:
   ```bash
   playwright install chromium
   ```
   This creates the `.playwright-browsers` folder in your project root.

5. **All required files** present in the project root:
   - `asins.csv`
   - `keywords.csv`
   - `spreadsheetIDs.csv`
   - `weighty-vertex-464012-u4-7cd9bab1166b.json`
   - (Optional) `.env` file

## Building Methods

### Method 1: Using the Build Script (Recommended)

The easiest way is to use the provided build script:

```bash
python build_exe.py
```

This script will:
1. Clean previous builds
2. Run PyInstaller with the correct settings
3. Copy the `.playwright-browsers` folder to the output
4. Verify all required files are present

### Method 2: Manual Build

If you prefer to build manually:

```bash
pyinstaller app.spec --clean
```

Then manually copy the `.playwright-browsers` folder:
```bash
# Windows PowerShell
Copy-Item -Path ".playwright-browsers" -Destination "dist\AmazonRankingTool\.playwright-browsers" -Recurse

# Windows CMD
xcopy /E /I ".playwright-browsers" "dist\AmazonRankingTool\.playwright-browsers"
```

### Method 3: Using PyInstaller Directly

You can also use PyInstaller directly, but the spec file is recommended:

```bash
pyinstaller --name AmazonRankingTool --windowed --add-data "asins.csv;." --add-data "keywords.csv;." --add-data "spreadsheetIDs.csv;." --add-data "weighty-vertex-464012-u4-7cd9bab1166b.json;." app.py
```

Then manually copy `.playwright-browsers` as in Method 2.

## Output Structure

After building, you'll find your executable in:

```
dist/
└── AmazonRankingTool/
    ├── AmazonRankingTool.exe
    ├── .playwright-browsers/
    │   └── chromium-xxxx/
    │       └── chrome-win/
    │           └── chrome.exe
    ├── asins.csv
    ├── keywords.csv
    ├── spreadsheetIDs.csv
    ├── weighty-vertex-464012-u4-7cd9bab1166b.json
    ├── .env (if included)
    └── [various DLL and Python library files]
```

## Distributing Your Application

To distribute the application:

1. **Copy the entire `AmazonRankingTool` folder** (not just the exe)
2. Ensure all files remain in the same folder:
   - The `.exe` file
   - The `.playwright-browsers` folder
   - All CSV files
   - The JSON credentials file
   - (Optional) `.env` file

**Important Notes:**
- Users can edit the CSV files directly in the folder
- The folder structure must be maintained
- All files must be in the same directory as the exe

## Troubleshooting

### "No Chromium browser found" Error

**Problem:** The `.playwright-browsers` folder is missing or not copied correctly.

**Solution:**
1. Make sure you ran `playwright install chromium` before building
2. Verify the folder exists in your project root
3. Check that the build script copied it correctly
4. Manually copy the folder if needed

### Missing Data Files

**Problem:** CSV or JSON files are not found at runtime.

**Solution:**
1. Verify the files are included in the `app.spec` file's `datas` list
2. Check that files exist in the `dist/AmazonRankingTool/` folder
3. Manually copy missing files to the exe directory

### Import Errors

**Problem:** Module not found errors when running the exe.

**Solution:**
1. Add missing modules to `hiddenimports` in `app.spec`
2. Rebuild the application
3. Check that all dependencies are in `requirements.txt`

### Large Executable Size

**Problem:** The output folder is very large (several hundred MB).

**Solution:**
- This is normal due to:
  - Python runtime
  - All dependencies (Playwright, Google APIs, etc.)
  - Playwright browsers
- Consider using UPX compression (not recommended as it can cause issues)

## Advanced Configuration

### Changing the Executable Name

Edit `app.spec` and change:
```python
name='AmazonRankingTool',
```

### Adding an Icon

1. Create or obtain an `.ico` file
2. Edit `app.spec` and change:
   ```python
   icon='path/to/your/icon.ico',
   ```

### Console Mode (for Debugging)

To see console output for debugging, change in `app.spec`:
```python
console=True,  # Show console window
```

## File Structure Requirements

The application expects the following structure at runtime:

```
[Application Folder]/
├── AmazonRankingTool.exe (or your custom name)
├── .playwright-browsers/
│   └── chromium-*/
│       └── chrome-win/
│           └── chrome.exe
├── asins.csv
├── keywords.csv
├── spreadsheetIDs.csv
├── weighty-vertex-464012-u4-7cd9bab1166b.json
├── .env (optional)
└── [Python runtime files]
```

The code in `scrap.py` automatically detects if it's running from an exe and adjusts paths accordingly using:
```python
if getattr(sys, 'frozen', False):
    base_dir = Path(sys.executable).parent
else:
    base_dir = Path(__file__).parent
```

This ensures all files are found correctly whether running as a script or as an executable.

