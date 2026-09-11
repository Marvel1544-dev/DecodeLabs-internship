# Project 3 — Data Analytics

Description
-----------
Data Analytics projects and tasks completed during my internship at Decode Labs.

This project contains a dataset and an SQL analysis file used to explore and analyze the data for Project 3. The analysis focuses on data-cleaning, exploratory data analysis (EDA), and SQL-based aggregation/insights.

Repository contents (Project3/)
-------------------------------
- `Dataset for Data Analytics.xlsx`  
  The primary dataset (Excel workbook) used for analysis. Open it in Excel or read it with pandas (`pd.read_excel`) for programmatic analysis.

- `DecodeLabs_Project3_SQL_Analysis.sql`  
  A collection of SQL queries used for analysis and reporting. Run these against your chosen SQL engine after loading the dataset into a database, or use a local SQLite database for quick testing.

- `gitkeep`  
  Placeholder file to keep the directory present in the repository.

Getting started
---------------
1. Clone this repository
   ```bash
   git clone https://github.com/Marvel1544-dev/DecodeLabs-internship.git
   cd DecodeLabs-internship/Project3
   ```

2. Recommended Python environment
   - Python 3.8+
   - Create and activate a virtual environment:
     ```bash
     python -m venv .venv
     source .venv/bin/activate   # macOS / Linux
     .venv\Scripts\activate      # Windows
     ```

3. Install suggested packages
   Create a `requirements.txt` with these packages, or install directly:
   ```
   pandas
   numpy
   matplotlib
   seaborn
   openpyxl
   sqlalchemy
   sqlite3   # builtin with Python; listed for clarity
   ```
   Install with:
   ```bash
   pip install pandas numpy matplotlib seaborn openpyxl sqlalchemy
   ```

Quick exploration (example)
---------------------------
Use pandas to preview and inspect the Excel dataset:

```python
import pandas as pd

# adjust sheet_name if needed
df = pd.read_excel("Dataset for Data Analytics.xlsx", sheet_name=0, engine="openpyxl")
print(df.shape)
print(df.info())
print(df.head())
```

Running the SQL analysis
------------------------
The SQL file includes queries and aggregations. To run them you must first load the data into a database. Two common workflows:

- SQLite (quick local test)
  1. Convert the Excel to CSV(s) or use a script to import into SQLite.
  2. Run the SQL file:
     ```bash
     sqlite3 project3.db < DecodeLabs_Project3_SQL_Analysis.sql
     ```
  Note: The SQL file may assume specific table names — adjust the import step accordingly.

- MySQL / Postgres
  1. Import the data into a table in your DB.
  2. Run queries using your DB client or a GUI (DBeaver, TablePlus, etc.).

If you want, I can help create a short Python script to convert the Excel workbook into a SQLite database with appropriate table names so the SQL file can run directly.

Recommendations & next steps
----------------------------
- Add a `requirements.txt` and (optionally) a small `scripts/` folder:
  - `scripts/convert_to_sqlite.py` — convert the Excel to a SQLite DB.
  - `scripts/eda.ipynb` — Jupyter notebook for the exploratory analysis and visualizations.

- Document the dataset schema or include a data dictionary in this README (column names, types, descriptions). I can generate a draft data dictionary after inspecting the spreadsheet.

Contributing
------------
Contributions, bug reports, and enhancements are welcome. Please open an issue or submit a pull request describing the change.

License
-------
This repository does not include a license file. If you want to make it open-source, consider adding an MIT or Apache-2.0 license.

Contact
-------
Author: Marvel1544-dev  
Project: DecodeLabs internship — Project 3
