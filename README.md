# ERP Data Migration

This tool copies data from an old ERP file into a new Excel file that uses different column names. The old file can be an Excel file (`.xlsx`) or a text file (`.txt`) exported from the old ERP.

You write down which old column goes into which new column in `data_migration_helper.xlsx`. The tool then creates the new file and copies every value exactly as it is: dates stay dates, codes keep their leading zeros, and numbers are not rounded. Before it saves the new file, it checks every copied cell against the old file.

The old file is never changed.


## What is in this folder

- `main.py`: the tool itself. Do not change it.
- `data_migration_helper.xlsx`: the list of which old column goes into which new column. You fill it in.
- `requirements.txt`: the list of packages the tool needs.
- `test_data`: practice files with invented data.
- `README.md` and `ERP_Data_Migration_User_Guide.pdf`: the instructions. Each update brings the newest version.
- `.git` (a hidden folder) and `.gitignore`: Git uses these to keep track of updates. Do not delete or change them.


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

If `python --version` still does not work but `py --version` does, Python was installed without the PATH box. That is fine: type `py -m venv .venv` instead of `python -m venv .venv` in step 6. Everything after that works the same.

### 3. Install Git

Git is a free program. It downloads the tool from GitHub, the website where the tool is kept, and later brings you updates with one command. You install it with `winget`, the app installer that comes with Windows.

1. In the Command Prompt, check that winget is there:

   ```
   winget --version
   ```

   If it shows a version number, for example `v1.9.25200`, go on. If it says `'winget' is not recognized`, see below.

2. Install Git:

   ```
   winget install --id Git.Git -e --source winget
   ```

   This means: install exactly the official Git package from Microsoft's list of apps.

   - If it asks `Do you agree to all the source agreements terms?`, type `Y` and press Enter.
   - If Windows asks "Do you want to allow this app to make changes to your device?", click **Yes**. If it asks for an administrator password that you do not have, ask your IT team.
   - Wait until it says `Successfully installed`. This can take a minute or two.

3. **Close the Command Prompt and open a new one.** A window that was already open does not know about Git yet.
4. Check that Git works:

   ```
   git --version
   ```

   It must show a version number, for example `git version 2.47.1.windows.1`.

If `winget` is not recognized: open the **Microsoft Store**, search for **App Installer**, click **Update** (or **Get**), then open a new Command Prompt and try again. If the Store is blocked on your computer, ask your IT team to install Git, or download it from https://git-scm.com/downloads/win and click **Next** on every screen.

If `git --version` says `'git' is not recognized`: close every Command Prompt window and open a new one. If it still fails, restart the computer. If it still fails after that, the install did not finish: type the `winget install` line again.

### 4. Download the tool

"Cloning" means downloading your own copy of the tool from GitHub. Do it once on each computer. After that, you get updates with `git pull` (see [Get updates](#get-updates)).

1. Open a new Command Prompt. It opens in your personal folder, for example `C:\Users\YourName`.
2. Type:

   ```
   git clone https://github.com/akshaysanil/erp_data_migration.git
   ```

3. Go into the new folder:

   ```
   cd erp_data_migration
   ```

What happens: Git creates a folder called `erp_data_migration` in your personal folder and downloads all the tool's files into it. You see lines like `Cloning into 'erp_data_migration'...` and `Receiving objects: 100%`. When the cursor comes back, it is finished. The folder stays linked to GitHub, which is how updates reach you later. You do not need a GitHub account.

The tool's folder is now `C:\Users\YourName\erp_data_migration`. To open it in File Explorer, type `%USERPROFILE%\erp_data_migration` in the address bar and press Enter.

- **"destination path 'erp_data_migration' already exists"**: the tool is already on this computer. Do not clone it again. Go to step 5. To get the newest version, use `git pull`.
- **You used the zip version before**: a folder from the zip file cannot get updates. Clone a new copy as above. If you already filled in your mapping, copy `data_migration_helper.xlsx` from the old folder into the new one and replace the file. When the tool works from the new folder, you can delete the old one.

### 5. Open a Command Prompt in the tool's folder

If you just did step 4, you are already there. Go to step 6. At any other time, use one of these:

- In File Explorer, open the tool's folder, click the address bar at the top, type `cmd` and press Enter. A black window opens in that folder.
- Or open a Command Prompt from **Start** and type `cd erp_data_migration`.

### 6. Create the environment and install the packages

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
- **"'pip' is not recognized"**: type `python -m pip` instead of `pip`, as in step 6. Check that the line starts with `(.venv)`. If it does not, type `.venv\Scripts\activate` first.
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
2. Open a Command Prompt in the tool's folder (see step 5 of the setup) and turn on the environment. You need to do this in every new window:

   ```
   .venv\Scripts\activate
   ```

3. Run the tool in one of these two ways.

**Way A: answer questions.** Type:

```
python main.py
```

The tool asks for the old file (`.xlsx` or `.txt`), the sheet (only if an Excel file has more than one), the row with the column names, the mapping file and the name of the new file. To give a file, drag it from File Explorer into the black window and press Enter. To accept the suggestion shown in [brackets], just press Enter. At the end, it shows your answers and waits for Enter before it starts.

**Way B: give the paths in the command.** Type everything in one line:

```
python main.py --old "C:\Data\old_data.xlsx" --output "C:\Data\new_data.xlsx"
```

- `--old` is the old file with the data: an Excel file (`.xlsx`) or a text file (`.txt`). For a text file, see [If the old file is a text file](#if-the-old-file-is-a-text-file-txt).
- `--output` is the new file to create.
- Put "quotes" around every path.
- To get a file's path, right-click the file in File Explorer and choose **Copy as path**. On older Windows, hold Shift while you right-click. Then paste it into the black window with a right-click or Ctrl+V.

Add these only when needed:

- `--sheet "Customers"`: the old Excel file has more than one sheet. Give the name of the sheet (the tab at the bottom in Excel) that has the data. A text file has no sheets, so leave this out.
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

### If the old file is a text file (.txt)

The old ERP can also give its data as a text file, for example `CUSTOMER.TXT`. Use it like an Excel file: drag it into the window when the tool asks for the old file, or give it with `--old`:

```
python main.py --old "C:\Data\CUSTOMER.TXT" --output "C:\Data\new_customers.xlsx"
```

The text file must be like this:

- The values on each line are separated by tabs.
- The column names are on the first line. If they are on another line, use `--header-row`, as with an Excel file.
- A value in "double quotes" can contain tabs and line breaks, for example an address on several lines. It is still one value, and its row is still one row.
- The file is saved as UTF-8. Excel's "Unicode Text" files work too. If the tool says the file is not saved as UTF-8, see [If you see an error](#if-you-see-an-error).

What to expect in the new file:

- Every value is copied as text, exactly as it is written in the text file. Codes keep their leading zeros. Numbers and dates are also copied as text, because a text file does not say which values are numbers or dates. Excel may show a small green triangle on numbers stored as text. That is expected.
- Line breaks inside a value are saved as normal Excel line breaks. The summary then has an extra line: `Note: 1 copied cell has text on several lines (for example an address); the line breaks were saved as Excel line breaks. The text itself is unchanged.`
- Messages give positions as for an Excel file: `C14` is the third value (column C) of row 14. Messages about a broken row also give its line number. To find that line in Notepad, press Ctrl+G, type the number and press Enter.


## If you see an error

When the tool stops with an error, it creates nothing. Fix the problem and run it again. For messages from `git pull`, see [Get updates](#get-updates).

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

Only with a text file (`.txt`):

- **"... is not saved as UTF-8 text"**: open the text file in Notepad, choose **File > Save as**, set **Encoding** (at the bottom) to **UTF-8**, save it under a new name, for example `CUSTOMER_utf8.txt`, and run again with the new file.
- **"these rows do not have the same number of values as row 1"**: a value in those rows has a tab or a line break without "quotes" around it, so its row was cut into pieces or its values moved to the wrong columns. The message gives the line numbers. Export the file again from the old ERP, or fix those lines in Notepad.
- **"cannot read the row that starts on line N"**: a value in that row starts with a `"` quote but does not end with one. Look at that line in Notepad.
- **"have more than 32,767 characters"** or **"contain control characters"**: a value is too long for an Excel cell, or has invisible special characters that Excel cannot store. Fix that value in the text file, or leave its column out of the mapping.
- **"more rows than an Excel sheet can hold"**: one Excel sheet holds at most 1,048,576 rows. Split the text file into smaller files and migrate each one.
- **"--sheet is only for Excel files"**: leave out `--sheet` when the old file is a text file.


## Practice with the test files

The `test_data` folder has practice files with **invented** data. In the Command Prompt, with the environment turned on, type:

```
python main.py --old test_data\old_erp_data.xlsx --helper test_data\test_mapping_helper.xlsx --output test_data\practice_result.xlsx
```

Open `test_data\practice_result.xlsx` next to `test_data\expected_new_erp_data.xlsx`. Both should have the same data. The summary should say 15 rows, `branch` left empty, and `Sales Rep, Internal Notes` NOT migrated. The Notes column of `test_mapping_helper.xlsx` explains what each row tests.

To run the practice again, delete `test_data\practice_result.xlsx` first, or add `--overwrite`.


## Get updates

The official copy of the tool is on GitHub. When you need a change, ask for it. When it is ready, one command brings it to your computer. Your own files and your setup stay as they are.

How it works:

1. **Ask for the change.** Tell the tool's maintainer what you need: a new feature, something that should work differently, or an error you saw. Say what you were doing and copy the exact message from the black window.
2. **Wait.** The maintainer makes the change, uploads it to GitHub and tells you when it is ready.
3. **Get it.** Run `git pull` on each computer that uses the tool, as shown below.

**Never send real ERP data.** Do not attach real ERP files, and do not send screenshots that show real data. The tool's messages show only column names and cell positions, so copying a message's text is safe. If an example file is needed, use the files in `test_data` or a small file with invented data.

### Get the newest version

1. Close the tool's files if they are open in Excel, especially `data_migration_helper.xlsx`.
2. Open a Command Prompt in the tool's folder (see step 5 of the setup) and type:

   ```
   git pull
   ```

3. Turn on the environment and run the install line again. If nothing new is needed, it only prints `Requirement already satisfied`, which is fine:

   ```
   .venv\Scripts\activate
   python -m pip install -r requirements.txt
   ```

What `git pull` shows:

- `Already up to date.`: you already have the newest version. Nothing changed.
- A list of files that ends with a line like `1 file changed, 12 insertions(+), 6 deletions(-)`: the update was downloaded, and the listed files now have their newest version. For example:

```
From https://github.com/akshaysanil/erp_data_migration
   bf70b4d..3c9e1a2  main       -> origin/main
Updating bf70b4d..3c9e1a2
Fast-forward
 main.py | 18 ++++++++++++------
 1 file changed, 12 insertions(+), 6 deletions(-)
```

`git pull` only replaces the tool's own files that were updated, such as `main.py` or the instructions. It does not touch your `.venv` folder or files you created yourself. Your filled-in `data_migration_helper.xlsx` is kept too, unless the update changes that same file (see below). `git pull` only downloads. Nothing on your computer is uploaded.

### If git pull stops with "would be overwritten"

You may see:

```
error: Your local changes to the following files would be overwritten by merge:
        data_migration_helper.xlsx
Please commit your changes or stash them before you merge.
Aborting
```

The update changes a file that you also changed on this computer, usually your filled-in mapping. Git stopped so that it does not delete your work. Nothing was changed. To fix it:

1. In File Explorer, **move** each file named in the message out of the tool's folder, for example to your Desktop. Move it, do not copy it.
2. Type `git pull` again. Now it finishes and downloads the newest version of the file.
3. If you moved your mapping file, open it next to the new `data_migration_helper.xlsx`, copy your rows into the new file and save it. Copy the rows instead of replacing the whole file, because the update may have changed it.

### Other git messages

- **"'git' is not recognized"**: Git is not installed on this computer, or the window was opened before Git was installed. Open a new Command Prompt. If it still fails, do step 3 of the setup.
- **"fatal: not a git repository"**: you are not in the tool's folder, or the folder came from the old zip file. Open the Command Prompt in the tool's folder (step 5 of the setup). If the folder came from the zip, do step 4 of the setup to download a new copy.
- **"Could not resolve host: github.com" or "unable to access"**: the computer cannot reach GitHub. Check the internet connection. On a company network, ask your IT team.
- **"Unlink of file ... failed. Should I try again? (y/n)"**: that file is open in Excel. Close it, then type `y` and press Enter.
- **Anything else**: do not try other git commands. Copy the message and send it to the tool's maintainer.

Two habits keep updates easy: use only `git clone` (once per computer) and `git pull` (for updates), and keep real ERP files and the new files you create outside the tool's folder, for example in `C:\Data`.


## Safety

- The old file and the mapping file are only read, never changed.
- The new file is checked cell by cell against the old file before it is saved. If anything is different, no new file is created.
- An existing file is never replaced unless you add `--overwrite`, or answer yes when the tool asks.
- Messages show only column names and cell positions (for example `C14`), never the data itself.
- Blank rows inside the data are kept, so row 10 of the old file is still row 10 of the new file. Blank rows at the very end are left out.


## Linux or macOS

Setup: install Git (on Ubuntu or Debian, `sudo apt install git`), then `git clone https://github.com/akshaysanil/erp_data_migration.git` and `cd erp_data_migration`. Then `python3 -m venv .venv`, then `source .venv/bin/activate`, then `python -m pip install -r requirements.txt`. Run with `python main.py` in the same way as on Windows, using paths like `/home/me/old_data.xlsx`.

If Python is missing on Ubuntu or Debian, or `python3 -m venv .venv` says `ensurepip is not available`, install it with `sudo apt install python3 python3-venv`. On macOS, install Python from https://www.python.org/downloads/macos/
