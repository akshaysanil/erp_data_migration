* [ ] 

# ERP Data Migration

This tool copies data from an old ERP Excel file into a new Excel file that uses different column names.

You write down which old column goes into which new column in `data_migration_helper.xlsx`. The tool then creates the new file and copies every value exactly as it is: dates stay dates, codes keep their leading zeros, and numbers are not rounded. Before it saves the new file, it checks every copied cell against the old file.

The old file is never changed.


## What is in this folder

- `main.py`: the tool itself. Do not change it.
- `data_migration_helper.xlsx`: the list of which old column goes into which new column. You fill it in.
- `requirements.txt`: the list of packages the tool needs.
- `test_data`: practice files with invented data.


## One-time setup (on each computer)

### 1. Check if Python is installed

Open a Command Prompt: click **Start**, type `cmd` and press Enter. A black window opens. Type this and press Enter:

```
python --version
```

- If it shows `Python 3.8` or a higher number (for example `Python 3.13.5`), Python is installed. Go to step 3.
- If it says `'python' is not recognized`, or the Microsoft Store opens, or the number is lower than 3.8, install Python in step 2.

### 2. Install Python (only if step 1 says it is missing)

1. Go to https://www.python.org/downloads/windows/
2. Under the newest **Stable Release**, click **Windows installer (64-bit)**. Do not choose the "embeddable package", because it does not work for this tool.
3. Open the downloaded file.
4. On the first screen, **tick "Add python.exe to PATH"** at the bottom. This is important. Then click **Install Now**. pip, the program that installs packages, is installed together with Python.
5. When it is finished, click **Close**. If it offers **Disable path length limit**, click that first.
6. **Close the Command Prompt and open a new one.** Type `python --version` again. It must now show the version number.

If the installer asks for an administrator password that you do not have, ask your IT team.

If `python --version` still does not work but `py --version` does, Python was installed without the PATH box. That is fine: type `py -m venv .venv` instead of `python -m venv .venv` in step 5. Everything after that works the same.

### 3. Unzip the tool's folder

Right-click the zip file and choose **Extract All**. A short location is easiest to type later, for example `C:\erp_data_migration`.

### 4. Open a Command Prompt in the tool's folder

In File Explorer, open the tool's folder, click the address bar at the top, type `cmd` and press Enter. A black window opens in that folder.

### 5. Create the environment and install the packages

Type these lines one by one, and press Enter after each one:

```
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
```

- After the second line, the line in the window starts with `(.venv)`. This means the environment is turned on.
- The third line downloads the packages, so it needs an internet connection. It ends with `Successfully installed ... openpyxl ...`.

Create the environment on each computer. A `.venv` folder copied from another computer does not work.

### If pip does not work

- **"No module named pip"**: type `python -m ensurepip --upgrade`, then type the install line again: `python -m pip install -r requirements.txt`.
- **"'pip' is not recognized"**: type `python -m pip` instead of `pip`, as in step 5. Check that the line starts with `(.venv)`. If it does not, type `.venv\Scripts\activate` first.
- **The install stops with a connection, proxy or SSL error**: the computer cannot download the packages. Check the internet connection. On a company network, ask your IT team.
- **"A new release of pip is available"**: this is only a notice. You can ignore it.


## Run a migration

### Step 1: Fill in the mapping

Open `data_migration_helper.xlsx` and go to the **Mapping** sheet. Each row becomes one column of the new file, in the same order:

- **New Column**: the column name that the new system needs. Each name may appear only once.
- **Old Column**: the column name exactly as it is in the old file. Upper/lower case and extra spaces do not matter. Leave it empty to create an empty column in the new file. You can use the same old column for more than one new column.
- **Notes**: for yourself. The tool ignores this column.

Delete the EXAMPLE rows before your first real run. Do not rename the Mapping sheet or the headers in row 1. Save and close the file.

### Step 2: Run the tool

1. Close the old Excel file.
2. Open a Command Prompt in the tool's folder (see step 4 of the setup) and turn on the environment. You need to do this in every new window:

   ```
   .venv\Scripts\activate
   ```

3. Run the tool in one of these two ways.

**Way A: answer questions.** Type:

```
python main.py
```

The tool asks for the old file, the sheet (only if there is more than one), the row with the column names, the mapping file and the name of the new file. To give a file, drag it from File Explorer into the black window and press Enter. To accept the suggestion shown in [brackets], just press Enter. At the end, it shows your answers and waits for Enter before it starts.

**Way B: give the paths in the command.** Type everything in one line:

```
python main.py --old "C:\Data\old_data.xlsx" --output "C:\Data\new_data.xlsx"
```

- `--old` is the old file with the data.
- `--output` is the new file to create.
- Put "quotes" around every path.
- To get a file's path, right-click the file in File Explorer and choose **Copy as path**. On older Windows, hold Shift while you right-click. Then paste it into the black window with a right-click or Ctrl+V.

Add these only when needed:

- `--sheet "Customers"`: the old file has more than one sheet. Give the name of the sheet (the tab at the bottom in Excel) that has the data.
- `--header-row 3`: the column names in the old file are not in row 1. Give the row number.
- `--helper "C:\Data\other_mapping.xlsx"`: use another mapping file instead of `data_migration_helper.xlsx` in the tool's folder.
- `--new-sheet "Import"`: the name of the sheet in the new file. The default is `Sheet1`.
- `--overwrite`: replace the new file if it already exists. Without this, the tool stops instead of replacing it.

For example, take the data from the sheet "Customers", whose column names are in row 3:

```
python main.py --old "C:\Data\old_data.xlsx" --sheet "Customers" --header-row 3 --output "C:\Data\new_data.xlsx"
```

To see every option explained, type `python main.py --help`.

**Do not create the new file yourself. The tool creates it.**

### Step 3: Check the result

When it is finished, the window shows a summary like this:

```
Old file     : C:\Data\old_data.xlsx (header row 1)
Columns      : 13 in new file, 12 filled from old file, 1 left empty
Rows copied  : 15
Verified     : 195 cells match the old file exactly
Saved to     : C:\Data\new_data.xlsx

New columns left empty (no Old Column in the helper): branch

Old columns NOT migrated (not in the helper): Sales Rep, Internal Notes
```

- **Always read the "NOT migrated" list.** These old columns are not in the new file because they are not in the mapping. Make sure that is what you want.
- Open the new file and look at a few rows.


## If you see an error

When the tool stops with an error, it creates nothing. Fix the problem and run it again.

- **"'python' is not recognized" or the Microsoft Store opens**: the environment is not turned on. Type `.venv\Scripts\activate` first. If that also fails, check that you are in the tool's folder and that the setup was done on this computer.
- **"No module named 'openpyxl'"**: the environment is not turned on. Type `.venv\Scripts\activate` first.
- **"running scripts is disabled on this system"**: the window is PowerShell, not Command Prompt. Type `cmd`, press Enter, and try again.
- **"... already exists; use --overwrite"**: a file with the new file's name already exists. Choose another name or add `--overwrite`.
- **"old column '...' is not in row 1 of the old file. Did you mean '...'?"**: the Old Column name in the mapping is spelled differently from the old file. Fix it in the mapping. If the column names in the old file are not in row 1, use `--header-row`.
- **"old column '...' appears more than once in the old file"**: two columns in the old file have the same name. Rename one of them in a copy of the old file.
- **"new column '...' is already used in row N"**: the same New Column name is written twice in the mapping.
- **"... has N sheets ...; choose one with --sheet"**: add `--sheet "SheetName"`.
- **"cannot open ... as an .xlsx file"**: the file is an old `.xls` file or is damaged. Open it in Excel and use **Save As** with the type **Excel Workbook (.xlsx)**.
- **"formulas whose result was never saved"**: open the old file in Excel, save it, close it, and run again.
- **"Permission denied ... Close the file if it is open in Excel"**: the new file is open in Excel. Close it and run again.


## Practice with the test files

The `test_data` folder has practice files with **invented** data. In the Command Prompt, with the environment turned on, type:

```
python main.py --old test_data\old_erp_data.xlsx --helper test_data\test_mapping_helper.xlsx --output test_data\new_erp_data.xlsx
```

Open `test_data\new_erp_data.xlsx` next to `test_data\expected_new_erp_data.xlsx`. Both should have the same data. The summary should say 15 rows, `branch` left empty, and `Sales Rep, Internal Notes` NOT migrated. The Notes column of `test_mapping_helper.xlsx` explains what each row tests.

To run the practice again, delete `test_data\new_erp_data.xlsx` first, or add `--overwrite`.


## Safety

- The old file and the mapping file are only read, never changed.
- The new file is checked cell by cell against the old file before it is saved. If anything is different, no new file is created.
- An existing file is never replaced unless you add `--overwrite`, or answer yes when the tool asks.
- Messages show only column names and cell positions (for example `C14`), never the data itself.
- Blank rows inside the data are kept, so row 10 of the old file is still row 10 of the new file. Blank rows at the very end are left out.


## Linux or macOS

Setup: `python3 -m venv .venv`, then `source .venv/bin/activate`, then `python -m pip install -r requirements.txt`. Run with `python main.py` in the same way as on Windows, using paths like `/home/me/old_data.xlsx`.

If Python is missing on Ubuntu or Debian, or `python3 -m venv .venv` says `ensurepip is not available`, install it with `sudo apt install python3 python3-venv`. On macOS, install Python from https://www.python.org/downloads/macos/
