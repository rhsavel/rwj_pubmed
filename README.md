# RWJBarnabas Health Monthly PubMed Publications Finder

A Python/Jupyter Notebook workflow for identifying PubMed-indexed publications associated with RWJBarnabas Health hospitals and related institutions, reviewing the results, and preparing a monthly academic publications newsletter.

The notebook uses the **NCBI PubMed E-utilities API** to search author affiliation information. It produces an **Excel review workbook** and a **draft Word newsletter**, then supports creating a **final Word newsletter** after human review.

> **Important:** This is an affiliation-based publications discovery and editorial review tool, not a validated or exhaustive institutional bibliography. Matches should be checked before publication or distribution.

## Features

- Searches PubMed for a specified calendar month using institution and hospital affiliation aliases.
- Covers the 12 RWJBarnabas Health acute-care hospitals, plus selected related institutions and historical names configured in the notebook.
- Optionally includes system-wide RWJBarnabas Health affiliation aliases.
- Retrieves PubMed IDs (PMIDs), then article metadata including titles, authors, journals, dates, affiliations, DOIs, and PubMed links.
- Identifies confirmed affiliation matches and separates uncertain or incomplete records for review.
- Handles publications associated with more than one institution without treating them as separate unique articles.
- Creates a formatted Excel workbook with review fields and institution dropdowns.
- Generates a Word draft organized by institution, including a publication-count summary, clickable PubMed/DOI hyperlinks, and page numbering.
- Generates a separately versioned final Word document from Excel review decisions, without repeating the PubMed search.
- Uses run-date and version identifiers to avoid overwriting prior outputs.

## Notebook

**`RWJBH_PubMed_Notebook_v18(2).ipynb`** — Version 18 of the monthly PubMed publications finder.

## Requirements

- Python 3
- Jupyter Notebook or Visual Studio Code with Jupyter support
- Internet connection for the PubMed search
- Python packages: `requests`, `pandas`, `python-docx`, `xlsxwriter`, and `openpyxl`

Install the required packages:

```bash
python -m pip install requests pandas python-docx xlsxwriter openpyxl notebook
```

The notebook also uses standard Python libraries, including `pathlib`, `datetime`, `calendar`, `re`, and `xml.etree.ElementTree`.

## Quick start

1. Download or clone this repository and open `RWJBH_PubMed_Notebook_v18(2).ipynb` in VS Code or Jupyter.
2. Locate the **Settings** code cell near the top of the notebook.
3. Set the reporting month, for example:

   ```python
   REPORT_MONTH = "2026-09"
   ```

4. Optionally adjust the newsletter title, appearance, PubMed date field, and contact email.
5. Select **Run All** to execute the PubMed search and generate the Excel workbook and draft Word document.
6. Check the notebook output for the exact directory and filenames. By default, files are saved in `Path.cwd()` (the Python working directory).
7. Open the Excel workbook and review the results, especially **Needs Review**.
8. Save and close Excel. Run **only the notebook's last code cell** to create the reviewed final Word document.

**Do not use Run All after reviewing the Excel workbook unless you intend to perform a new search and create a new set of output files.**

## Configuration

Most months, only `REPORT_MONTH` needs to change.

| Setting | Purpose |
| --- | --- |
| `REPORT_MONTH` | Month to search, formatted as `YYYY-MM`. |
| `DATE_TYPE` | `"edat"` for records entered into PubMed during the month (default); `"pdat"` for PubMed publication-date filtering. |
| `NEWSLETTER_TITLE` | Title displayed in both Word documents. |
| `FONT_NAME`, `FONT_SIZE` | Word document font settings. |
| `PAGE_MARGIN_INCHES` | Word document margin setting. |
| `OUTPUT_FOLDER` | Location for generated Excel and Word files; defaults to `Path.cwd()`. |
| `CONTACT_EMAIL` | Optional contact email included in NCBI requests. |
| `INCLUDE_SYSTEM_WIDE_ALIASES` | Whether to search RWJBarnabas Health system-wide affiliation names. |
| `MAX_RESULTS` | Safety limit that stops an unexpectedly large search instead of silently truncating it. |

Institution names and alternative spellings can be maintained in `HOSPITAL_ALIASES`, `SYSTEM_ALIASES`, and related matching rules.

## Output files

A typical search produces:

| File | Description |
| --- | --- |
| `RWJBH_Publications_YYYY-MM_YYYY-MM-DD_v01.xlsx` | Excel workbook for auditing and reviewing matches. |
| `RWJBH_Publications_YYYY-MM_YYYY-MM-DD_v01_DRAFT.docx` | Automatically generated draft newsletter. |
| `RWJBH_Publications_YYYY-MM_YYYY-MM-DD_v01_FINAL_v01.docx` | Final newsletter generated after spreadsheet review. |

The Excel workbook contains these sheets:

- **By Hospital** — publications organized by assigned institution.
- **Matched Articles** — publications with identified institutional affiliation matches.
- **Needs Review** — possible matches and records requiring a decision; includes `Include?` and `Hospital to Assign` fields.
- **All Search Results** — the broader retrieved results for auditing.

The final export reads the relevant Excel workbook and applies the manual review decisions. For articles you elect to include from **Needs Review**, enter `Yes` in **Include?** and select or enter the appropriate institution in **Hospital to Assign**. Other records may be marked `No` or left blank. Save and close the workbook before running the final export cell.

If multiple versions exist, pay attention to the workbook filename displayed by the final cell so that you know which review is being used.

## How the PubMed search works

1. Build an `OR` query from configured hospital and institution names, limited to PubMed's affiliation field (`[ad]`, shorthand for `[Affiliation]`).
2. Call `esearch.fcgi` to identify matching PMIDs for the selected month, retrieving paginated results as necessary.
3. Fetch detailed PubMed article records using `efetch.fcgi` and parse the returned XML.
4. Inspect affiliation strings to assign matching hospitals and flag ambiguous records.
5. Export the results for manual review and newsletter preparation.

The default `DATE_TYPE = "edat"` is designed to find **newly added PubMed records** in a given month. It does **not** mean that every article was originally published during that same month.

## Limitations and quality assurance

- PubMed affiliation data can be missing, incomplete, or inconsistent; some relevant articles may not appear.
- A hospital name in an affiliation does not independently establish current employment or institutional publication ownership.
- Certain generic institution names require extra location/context checks.
- Rutgers-only affiliations are **not** automatically assumed to be RWJBarnabas Health publications.
- Historical institution names and manually maintained aliases require periodic verification.
- Results for an earlier month can change if PubMed records are added or revised.
- The draft and final documents should be checked for completeness, duplicate presentation, institution assignment, and citation accuracy before circulation.

For an auditable monthly process, retain the reviewed Excel files and final documents, and consider checking for repeated PMIDs across months.

## Data and access

This workflow retrieves bibliographic metadata from the public PubMed service. It does not require patient data or hospital clinical-system access. Avoid uploading internal-only review annotations or unpublished organizational information to a public repository.

## References

- [NCBI E-utilities documentation](https://www.ncbi.nlm.nih.gov/books/NBK25499/)
- [PubMed](https://pubmed.ncbi.nlm.nih.gov/)
- [RWJBarnabas Health](https://www.rwjbh.org/)

## Project status

Version 18. Designed as a practical, human-reviewed monthly academic-publications newsletter workflow. The generated newsletter is a draft until reviewed and approved for distribution.
