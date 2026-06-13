# Brock & Scott Foreclosure Sales Scraper

> A desktop automation tool for collecting foreclosure sale data and exporting structured records to Excel.

---

## Overview

Built with **Python**, **Selenium**, **Tkinter**, **OpenPyXL**, and **Pandas** — this tool connects directly to the Brock & Scott foreclosure sales website, navigates through listings automatically, and delivers a clean Excel file without any copy-pasting.

---

## Features

| Category | Details |
|---|---|
| **Interface** | Desktop GUI with live activity log and record counter |
| **Input** | Manual URL entry or bulk import from Excel / CSV |
| **Automation** | Headless Chrome, automatic pagination, no user intervention needed |
| **Controls** | Pause, Resume, and Stop at any point without data loss |
| **Output** | Auto-generated Excel file with professional column formatting |

---

## Data Collected

Each record captures the following fields from the Brock & Scott listings:

- County
- Sale Date
- State
- Court SP Number
- Case Number
- Property Address
- Book / Page

---

## How to Use

### Step 1 — Get Your URLs

Open [brockandscott.com/foreclosure-sales](https://brockandscott.com/foreclosure-sales/) and apply any filters:

- State
- County
- Case Number
- Court SP Number
- Sale Date Range

Once results load, copy the URL from your browser address bar.

---

### Step 2 — Launch the Scraper

Open the application and paste your URLs into the input box — one per line:

```
https://brockandscott.com/foreclosure-sales/?state=NC&county=Wake
https://brockandscott.com/foreclosure-sales/?state=NC&county=Mecklenburg
https://brockandscott.com/foreclosure-sales/?state=VA&county=Fairfax
```

---

### Step 3 — Choose Output Location

Select the folder and filename where the Excel file will be saved.

---

### Step 4 — Start

Click **START**. The scraper will:

1. Launch Chrome automatically
2. Visit each URL
3. Collect all listings
4. Page through results until complete
5. Write everything to Excel

---

## Importing URLs from a File

Prefer working from a spreadsheet? Click **Upload Excel** to load URLs in bulk.

**Supported formats:** `.xlsx` `.xls` `.csv`

The scraper reads URLs from the file and loads them into the input field automatically.

---

## Scraper Controls

| Button | Behaviour |
|---|---|
| **START** | Begin scraping all URLs |
| **PAUSE** | Hold at the current page — no data is lost |
| **RESUME** | Pick up exactly where it paused |
| **STOP** | End the session and save all collected records immediately |

---

## Setup

### System Requirements

- Python 3.9 or newer — [python.org/downloads](https://www.python.org/downloads/)
- Google Chrome installed
- Windows operating system
- Active internet connection

---

### Dependencies

| Library | Version | Purpose |
|---|---|---|
| `selenium` | >= 4.0.0 | Controls Chrome to navigate and extract data from the website |
| `openpyxl` | >= 3.0.0 | Creates and formats the Excel output file |
| `pandas` | >= 1.3.0 | Handles data processing and CSV / Excel URL imports |
| `tkinter` | Bundled with Python | Powers the desktop GUI — no separate install needed |

---

### Installing Dependencies in PyCharm

**Step 1 — Open the Terminal**

At the bottom of the PyCharm window, click the **Terminal** tab.
If it is not visible, go to: `View` → `Tool Windows` → `Terminal`

**Step 2 — Confirm you are in the project folder**

The terminal should show the path to your project directory. If not, navigate there:

```bash
cd path\to\your\project
```

**Step 3 — Install the libraries**

```bash
pip install -r requirements.txt
```

**Step 4 — Run the application**

```bash
python Foreclosure scraper.py
```

---

### Installing Dependencies in Visual Studio Code

**Step 1 — Open the Terminal**

Go to the top menu: `Terminal` → `New Terminal`

A terminal panel will open at the bottom of the screen.

**Step 2 — Confirm you are in the project folder**

VS Code usually opens the terminal in your project folder automatically. Verify the path shown matches your project. If not:

```bash
cd path\to\your\project
```

**Step 3 — Install the libraries**

```bash
pip install -r requirements.txt
```

**Step 4 — Run the application**

```bash
python Foreclosure scraper.py
```

---

### Installing Libraries One by One (Optional)

If you prefer to install each library individually rather than using the requirements file:

```bash
pip install selenium
pip install openpyxl
pip install pandas
```

---

### Verifying the Installation

To confirm all libraries installed correctly, run:

```bash
pip show selenium openpyxl pandas
```

Each library should display its name and version. If any are missing, re-run the install command for that specific package.

---

## Output

The exported Excel file contains one row per listing with these columns:

| Column | Description |
|---|---|
| County | County where the sale is listed |
| Sale Date | Scheduled date of foreclosure sale |
| State | State abbreviation |
| Court SP | Court SP number |
| Case | Case number |
| Address | Full property address |
| Book Page | Book and page reference |

---

## Troubleshooting

**Chrome does not open**

```bash
pip install --upgrade selenium
```

**Excel file cannot be saved**

Close the file in Excel before running the scraper. Excel locks files that are open.

**Missing module error**

```bash
pip install -r requirements.txt
```

---

## Purpose

Manual data collection from foreclosure listings is slow and error-prone. This tool removes that work entirely — run it once and receive a structured, ready-to-use spreadsheet.
