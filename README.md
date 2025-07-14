# Excel Discrepancy Checker

A portable Streamlit app to compare two Excel files using a customizable rubric. Designed for auditors and analysts to quickly spot discrepancies in key sales and unit data.

## Features
- Upload two Excel files (Subway and Sunthesis)
- Automatically detects and compares key ranges/labels as defined in `Descrepncy Rubrik.txt`
- Handles custom day/date headers for each label
- Displays discrepancies in a clear, sortable table
- Export results to Excel
- Minimal, focused UI
- Portable: runs on Windows and Mac with double-click scripts

## Quick Start

### 1. Install Python (if not already installed)
- [Download Python 3.x](https://www.python.org/downloads/)
- Make sure to check "Add Python to PATH" during installation (Windows)

### 2. Install Requirements
- Open a terminal/command prompt in the app directory
- Run:
  ```
  pip install -r requirements.txt
  ```
  *(This step is automatic if you use the provided launch scripts!)*

### 3. Launch the App

#### **Windows**
- Double-click `Run Excel Discrepancy Checker.bat`
- The app will open in your default browser

#### **Mac**
- Open Terminal, run:
  ```
  chmod +x run_excel_discrepancy_checker.command
  ```
- Double-click `run_excel_discrepancy_checker.command` in Finder
- The app will open in your default browser

### 4. Usage
- Upload your two Excel files (must contain a sheet named `ControlSheet`)
- Review discrepancies in the table
- Download the report if needed

## Troubleshooting
- **Missing dependencies:** The launch scripts will auto-install required packages if missing.
- **Python not found:** Ensure Python 3.x is installed and available in your PATH.
- **App doesn't open:** Try running the launch script from a terminal to see error messages.
- **Excel file issues:** Ensure your files match the expected structure (see `Descrepncy Rubrik.txt`).

## Further Improvements & Suggestions
- **True Standalone Packaging:** Use [PyInstaller](https://pyinstaller.org/) or [Briefcase](https://beeware.org/project/projects/tools/briefcase/) to bundle Python and dependencies into a single executable (advanced, larger file size).
- **Cross-platform GUI Launcher:** Create a small native launcher (e.g., with Electron or Tauri) for a more polished experience.
- **Docker Support:** Add a `Dockerfile` for containerized deployment.
- **Automatic Updates:** Integrate with GitHub Releases or a private server for easy updates.
- **Customizable Rubric UI:** Allow rubric editing from the app interface.
- **User Authentication:** Add login for multi-user or sensitive environments.
- **Logging & Error Reporting:** Improve error logs and optionally send error reports.
- **Localization:** Add support for multiple languages.
- **Mobile/Tablet UI:** Optimize layout for smaller screens.

## License
MIT (or specify your license here)

---

*For questions or support, contact the developer or open an issue in your repository.* 