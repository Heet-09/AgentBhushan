# Jira Automation

Turn failed test cases in an Excel test sheet into Jira bugs.

1. `parse_excel.py` reads the Excel workbook and extracts every row whose **Chrome** column is `Fail`.
2. `create_jira_bugs.py` creates one Jira Bug for each of those rows.

## Requirements

- Python 3.10+
- Packages: `pandas`, `openpyxl`, `requests`

```powershell
pip install pandas openpyxl requests
```

## Step 1: Parse the Excel file

Print every sheet in a readable form:

```powershell
python parse_excel.py "C:\path\to\TestSheet.xlsx"
python parse_excel.py "C:\path\to\TestSheet.xlsx" -o output.txt        # save to a text file
python parse_excel.py "C:\path\to\TestSheet.xlsx" --sheet "Login"      # only one sheet
python parse_excel.py "C:\path\to\TestSheet.xlsx" --max-rows 20        # limit rows per sheet
```

The header row is found automatically (it does not have to be row 1). Single-cell rows such as
"SCAD-2 Search filters" are printed as section titles.

Extract the failed test cases (rows where `Chrome` is `Fail`):

```powershell
python parse_excel.py "C:\path\to\TestSheet.xlsx" --failed -o failed.json   # list of dicts
python parse_excel.py "C:\path\to\TestSheet.xlsx" --failed -o failed.xlsx   # Excel file
```

Each item in `failed.json` is a dictionary with all the data of the row, plus the `Sheet` name and
the `Excel row` number.

## Step 2: Create Jira bugs

### Configure

Create a `.env` file next to the scripts (one `KEY=value` per line, no quotes):

```
JIRA_URL=https://your-site.atlassian.net
JIRA_EMAIL=you@example.com
JIRA_API_TOKEN=your-api-token
JIRA_PROJECT=SCRUM
```

- **JIRA_URL**: your Jira site, with no trailing path.
- **JIRA_EMAIL**: your Atlassian account email (Jira Cloud). Leave it out if you use a personal access
  token on Jira Server / Data Center.
- **JIRA_API_TOKEN**: create one at <https://id.atlassian.com/manage-profile/security/api-tokens>.
- **JIRA_PROJECT**: the project key. It is the part before the dash in a ticket number (`SCRUM-12`
  means `SCRUM`), or open `https://your-site.atlassian.net/rest/api/2/project` in the browser while
  logged in and read the `key` fields.

> **Never commit `.env`.** It contains your token. Add `.env` to `.gitignore`.

### Run

```powershell
python create_jira_bugs.py failed.json            # dry run: prints what would be created
python create_jira_bugs.py failed.json --create   # really create the bugs
```

You can also pass the project on the command line: `--project SCRUM`.

If the project does not exist yet, `--create` creates it first (a company-managed Scrum project, which
includes the Bug issue type). This needs Jira admin rights. The project name comes from
`JIRA_PROJECT_NAME` in `.env`, or defaults to the key. To create the bugs again in a different project,
delete `created_bugs.json` first, otherwise those rows are skipped.

### Field mapping

| Jira field  | Source                        |
|-------------|-------------------------------|
| Summary     | `Test Case`                   |
| Description | `Steps` + `Expected Result`   |
| Issue type  | Bug                           |
| Priority    | `Priority`                    |

Created bugs are saved in `created_bugs.json` (sheet and row mapped to the Jira key), so running the
script again skips rows that already have a bug and does not create duplicates.

### See the created bugs

Open the board at `https://your-site.atlassian.net/jira/software/projects/SCRUM/boards/1`, or search
with JQL:

```
project = SCRUM AND issuetype = Bug ORDER BY created DESC
```

## Troubleshooting

| Error | Fix |
|-------|-----|
| `Specify a valid issue type` | The project has no **Bug** issue type. Add it under Project settings → Work types (Issue types). |
| `401` | Wrong email or token. Create a new token and check `.env` for stray spaces or line breaks. |
| `404` | Wrong `JIRA_URL` or project key. |
| Priority error | The sheet's priority names (High / Medium / Low) must exist in your Jira. |
| `Set the project key...` | Set `JIRA_PROJECT` in `.env` or pass `--project`. |

## Files

| File | Purpose |
|------|---------|
| `parse_excel.py` | Parse the workbook, print sheets, extract failed rows |
| `create_jira_bugs.py` | Create Jira bugs from `failed.json` |

