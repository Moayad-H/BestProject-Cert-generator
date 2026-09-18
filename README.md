# Certificate Forge - Bulk PowerPoint & Excel Certificate Generator

A lightweight, single-page web application to generate certificates in bulk directly inside your browser from a PowerPoint (`.pptx`) template and an Excel (`.xlsx`) nomination sheet.

**100% Client-Side & Private**: All file parsing, name cleaning, course matching, and PowerPoint XML generation run directly in the browser. Zero data is sent to or stored on any server.

---

## Features

- **Step 1: PowerPoint Template Upload**:
  - Upload any `.pptx` template containing `{Student_Name}` and `{Course_Name}` text placeholders.
  - Or click **Use Included Template** to quick-start with `Best Project Placeholder.pptx`.
  
- **Step 2: Excel Nomination Sheet Upload & Parser**:
  - Upload your `.xlsx` nomination workbook.
  - Automatically filters rows where Column A (Course Name) is non-empty and Column F (Certificates) is `TRUE`.
  - Cleans messy Column C cells: extracts names, removes IDs, digits, labels (`NAME:`, `ID:`, `Project #`), emails, and phone numbers.
  - Automatically detects Arabic script (`\u0600-\u06FF`) and routes them to a dedicated **Skipped Arabic** drawer for manual review and translation.
  - Shortens names to **First + Last** while preserving compound first names (e.g. `Abd Allah`, `Menna Allah`) and compound last names (e.g. `El Sayed`, `Alaa Eldin`, `Zein-Eldeen`, `El-Gazairy`).
  - Title-cases names and preserves hyphenated parts.
  - Matches course names against the official `Courses code & names` sheet; flags unconfident matches in yellow.

- **Step 3: Interactive Review Table**:
  - Live row counter with summary pills (`Students Ready`, `Flagged`, `Skipped Arabic`).
  - Search and filter table in real time.
  - In-place editable cells for Course Name and Student Name.
  - Visual badges for long names indicating 90% box widening (> 17 chars) and scaled font size (> 20 chars).
  - Add new students or delete existing rows.
  - 1-click clipboard copy and transcribe actions in the Skipped Arabic modal.

- **Step 4: PPTX Generator & Export**:
  - Duplicates slide1 and its relationships directly inside the `.pptx` ZIP package using `JSZip`.
  - Dynamically centers and widens the text box to 90% of slide width for names longer than 17 characters.
  - Dynamically scales font size down for names longer than 20 characters.
  - Generates `Best_Project_Certificates.pptx` with zero OpenXML repair warnings.
  - Includes instructions to export as a single combined PDF via PowerPoint: **File → Save As → PDF**.

---

## How to Run Locally

You can serve the folder with any static web server:

```bash
# Using Python 3
python3 -m http.server 8088
```

Then open your browser to [http://localhost:8088](http://localhost:8088).

---

## Project Structure

```
cert-generator/
├── index.html            # Main web app interface
├── css/
│   └── style.css         # Modern dark-mode styling
├── js/
│   ├── app.js            # Main controller & UI state
│   ├── parser.js         # Excel parsing, cleaning & course matching
│   └── pptx.js           # Client-side PPTX XML manipulation
├── vendor/
│   ├── jszip.min.js      # JSZip library (bundled offline)
│   └── xlsx.full.min.js  # SheetJS library (bundled offline)
├── samples/              # Included sample templates & data
│   ├── Best Project Placeholder.pptx
│   └── Best Project Spring 2026.xlsx
└── README.md
```
