# ITU-T Rapporteur's status report generator

Generates pre-populated status reports for ITU-T Study Group 12 Rapporteurs as Word documents.

**Latest reports: https://ludomal.github.io/rapporteur_helper/**

Meeting details (place, dates) are automatically fetched from the ITU-T website. The tool scrapes:

- Question title
- Rapporteur / co-rapporteur contact details
- List of contributions
- List of temporary documents (TDs)
- Work programme

## Generate reports via GitHub Actions (recommended)

The easiest way to generate reports is through the GitHub Actions workflow:

1. Go to **Actions** → **Generate reports**
2. Click **Run workflow**
3. Fill in the parameters:
   - **Questions**: which questions to generate (e.g. `1-20`, `1,2,7`, or `5`)
   - **Study Group**: the SG number (default: `12`)
4. Click **Run workflow**

The workflow automatically fetches the current meeting details from the ITU-T SG page and generates one `.docx` file per question.

Once complete, download the reports from the **Artifacts** section of the workflow run. The artifact name includes the questions and study group for traceability (e.g. `status-reports-Q1-20-SG12`).

## Generate reports locally

1. Install dependencies:

```shell
uv sync
```

2. Run the tool:

```shell
uv run rapporteur_helper
```

This generates reports for all questions (1–20) using the current meeting info from the ITU-T website.

### Options

```
-q, --questions TEXT        Questions: range '1-20', list '1,2,7', or single '5'
-s, --study-group INTEGER   Study Group number (default: 12)
-d, --meeting-date TEXT     Override meeting start date (YYMMDD)
-p, --meeting-place TEXT    Override meeting location
    --meeting-end-date TEXT Override meeting end date (YYMMDD)
    --add-qall / --no-add-qall  Include QALL documents (default: False)
-o, --output-dir PATH       Output directory (default: current directory)
-v, --verbose               Enable verbose output
```

Example — generate only questions 1 and 7:

```shell
uv run rapporteur_helper -q "1,7" -v
```

Reports are saved to a directory named by the meeting start date (e.g. `./260609/`).

## Development

```shell
uv sync              # install dependencies
make check           # run linting and quality checks
make test            # run tests
```
