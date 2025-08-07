# SAP Statement GUI - Build Instructions

This guide explains how to create a standalone Windows executable (`.exe`) for the SAP Statement GUI using PyInstaller.

## Prerequisites

- **Python 3.x** installed on your system
- **PyInstaller** installed  
  Install with:  
  ```
  pip install pyinstaller
  ```
- Ensure your working directory contains:
  - `SapStatementGUI.py`
  - The `icons` folder with `Loblaws.ico` inside

## Building the Executable

1. **Open Command Prompt** and navigate to your project directory:
   ```
   cd path\to\your\project
   ```

2. **Run PyInstaller** with the following command:
   ```
   pyinstaller -w --icon="retrieve_sap_statement\resources\icons\loblaw.ico" --add-data="retrieve_sap_statement;retrieve_sap_statement" --name "SAP Statement" SapStatementGUI.py
   ```
   - `-w` : No console window --windows and --noconsole can be used too (for GUI apps)
   - `--icon="retrieve_sap_statement\resources\icons\loblaw.ico" ` : Sets the application icon
   - `--add-data="retrieve_sap_statement;retrieve_sap_statement"` : Includes the `icons` folder in the build
   - `--name "SAP Statement"` : Sets the output executable name
   - `SapStatementGUI.py` : Your main Python script

3. **Find your executable** in the `dist\SAP Statement` folder.

## Notes

- Always use the correct path separator for `--add-data`:
  - On **Windows**: `--add-data="source;dest"`
  - On **macOS/Linux**: `--add-data="source:dest"`
- If you run the `.exe` from a network share, it may be slow to start. For best performance, copy it to your local drive before running.
- If you update your code or resources, repeat the build process.

## Troubleshooting

- If the icon does not appear, ensure `Loblaws.ico` exists in the `icons` folder and the path is correct.
- If you get missing file errors, check that `--add-data` is set correctly and the `icons` folder is present.