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

# Defect Grouping Logic
## 1. Objective

Prevent duplicate Jira bug creation when multiple failed test cases are caused by the same underlying defect.

A common Excel column named `Defect Group` will be used to identify test cases that belong to the same defect.

**Core rule:** One defect group should generate one Jira bug candidate, regardless of how many test cases fail within that group.

## 2. Excel Input

The Excel test execution report should contain these columns:

| Column | Description |
|---|---|
| Test Case ID | Unique test case identifier |
| Test Case | Test case title or description |
| Status | Pass, Fail, Blocked, or Not Run |
| Expected Result | Expected application behavior |
| Actual Result | Actual behavior observed during execution |
| Severity | Severity of the defect |
| Defect Group | Common identifier for failures caused by the same defect |
| Existing Jira ID | Jira issue key if the defect has already been reported |

## 3. Example

| Test Case ID | Status | Actual Result | Defect Group | Existing Jira ID |
|---|---|---|---|---|
| TC001 | Fail | Formula displays 150 instead of 150.00 | DG-001 | |
| TC002 | Fail | Formula displays 250 instead of 250.00 | DG-001 | |
| TC003 | Fail | Subtotal is calculated incorrectly | DG-002 | |
| TC004 | Fail | Formula value has missing decimal places | DG-001 | |
| TC005 | Pass | Search results are displayed correctly | | |

### Expected Output

- `DG-001`: One Jira bug candidate containing TC001, TC002, and TC004.
- `DG-002`: One Jira bug candidate containing TC003.
- TC005: No bug candidate because the test case passed.

**Result: 4 failed test cases produce 2 Jira bug candidates.**

## 4. Grouping Rules

1. Only test cases with status `Fail` are eligible for bug generation.
2. Failed test cases with the same non-empty `Defect Group` must be combined into one bug candidate.
3. Different defect group identifiers must produce separate bug candidates.
4. Failed test cases with a blank `Defect Group` must remain separate candidates. Blank groups must never be combined with each other automatically.
5. Passed, Blocked, and Not Run test cases must not generate bug candidates.
6. If an `Existing Jira ID` is provided, validate the issue before linking the related failures. Do not create another bug for the same group by default.
7. If a group contains conflicting Jira issue IDs, flag it for manual review.
8. The tester must review and approve bug candidates before Jira tickets are created.
9. The tool must prevent repeated submissions from creating duplicate tickets.
10. If the same defect group contains different severity values, flag the conflict or apply an explicitly approved severity rule.

## 5. Processing Flow

```text
Upload Excel Report
        |
        v
Validate Required Columns
        |
        v
Filter Failed Test Cases
        |
        v
Separate Grouped and Ungrouped Failures
        |
        v
Group by Defect Group
        |
        v
Check Existing Jira ID
        |
        v
Generate Bug Candidates
        |
        v
Tester Reviews and Approves
        |
        v
Create Jira Bug or Link Existing Issue
        |
        v
Generate Processing Report
```

## 6. Important Considerations

- `Defect Group` is the primary grouping mechanism in the initial version.
- Testers are responsible for assigning the same group only when failures are believed to originate from the same underlying defect.
- Similar error messages do not necessarily mean that failures have the same root cause.
- The tool should preserve individual test case IDs, steps, expected results, and actual results within the consolidated bug description.
- The tool must not silently discard failures when grouping them.
- Jira creation results should be recorded so that a repeated upload can be handled safely.

## 7. Future Enhancement

In a future version, the tool can suggest defect groups automatically using test case descriptions, failed steps, actual results, error messages, and module information.

The tester should be able to confirm or modify these suggestions before bug creation.

Automatic grouping should be introduced only after the initial Defect Group-based approach works reliably.
