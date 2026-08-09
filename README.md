# PrimeWash Excel Cleanup & Automation

A Python-based utility project designed to streamline data processing, automate Excel data cleaning, generate visual charts, and build pivot tables for business analytics.

---

## Project Structure

```text
PRIMEWASH-EXCEL-CLEANUP/
│
├── raw_data/                 # Directory for raw input files
├── venv/                     # Python virtual environment
├── clean.py                  # Script for data cleaning and preprocessing
├── create_excel_charts.py    # Generates visual charts from processed data
├── create_excel_chartsapp.py # Application interface / runner for chart generation
├── pivot_tables.py           # Script for aggregating data and generating pivot tables
└── README.md

Technologies Used
Language: Python

Tooling: Git & VS Code

Features
Automated Data Cleaning (clean.py): Parses raw files and formats datasets for analysis.

Pivot Table Generation (pivot_tables.py): Aggregates and summarizes key metrics into structural pivot tables.

Data Visualization (create_excel_charts.py & create_excel_chartsapp.py): Automatically builds and exports visual charts tied to the dataset.

Setup & Installation Instructions
Follow these steps to set up and run the project locally:

Clone the repository:

Bash
git clone [https://github.com/your-username/PRIMEWASH-EXCEL-CLEANUP.git](https://github.com/your-username/PRIMEWASH-EXCEL-CLEANUP.git)
Navigate to the project directory:

Bash
cd PRIMEWASH-EXCEL-CLEANUP
Activate the virtual environment:

On Windows (Git Bash):

Bash
source venv/Scripts/activate
On macOS / Linux:

Bash
source venv/bin/activate
Usage
Run the scripts in your terminal in the following workflow order:

Clean the raw data:

Bash
python clean.py
Generate Pivot Tables:

Bash
python pivot_tables.py
Generate Charts:

Bash
python create_excel_charts.py

### Next steps:
1. Paste this block into your open `README.md` file in VS Code.
2. Save the file (`Ctrl + S` or `Cmd + S`).
3. Run your Git commands in the terminal to push it:
   ```bash
   git add README.md
   git commit -m "Add project README.md"
   git push origin main