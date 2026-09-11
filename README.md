# 🎉 Initial Release (V1.0.0)

Welcome to the first official release of **Folder Indexer**! This is a lightweight C# console tool designed to quickly scan a local directory, alphabetize its subfolders, and output a clean, professionally styled Excel spreadsheet.

---

### ✨ What's New & Features
* **Interactive CLI:** Step-by-step prompts guide you through setting up your index preferences.
* **Custom Headers & Naming:** Option to add a custom column header and name your final Excel output file.
* **Smart Path Sanitization:** Automatically cleans up extra quotes and spaces if you use Windows' "Copy as path" feature.
* **Hidden Folder Filtering:** Automatically ignores hidden system folders starting with a dot (like `.git` or `.vs`).
* **Polished Excel Output:** Generates a report-style spreadsheet with a light blue theme, custom borders, hidden default gridlines, and auto-adjusted column widths.

### 🛡️ Reliability & Fixes
* **Empty Directory Safety:** Gracefully handles directories that contain no subfolders without throwing formatting exceptions.
* **Input Safeguards:** Built-in retry loops and validation for directory paths and user prompts to prevent accidental crashes.

---

### 📦 How to Use
1. Download the standalone `FolderIndexer.exe` file from the assets below.
2. Double-click or run it from your terminal.
3. Follow the on-screen prompts to point it at your target directory and save your index!
