# High School Application Dashboard Generator (Mokil Middle School)

This project automates the generation of dashboards and reports for high school application results. It processes data from Google Sheets and outputs visual HTML dashboards and Excel spreadsheets.

## Features

- **Data Integration**: Fetches student application data directly from Google Sheets.
- **Dual Reporting**: Generates reports for both Early Admissions (전기고) and Late Admissions (후기고).
- **Visual Dashboards**: Creates interactive HTML cards for easy visualization of results.
- **Excel Exports**: Produces structured Excel files for administrative use.
- **Privacy Focused**: Report files containing PII are excluded from version control.

## Project Structure

```text
├── generators/                 # Core Python scripts
│   ├── generate_dashboard.py   # Main dashboard generator (Auth required)
│   └── mokil_high_school_results_gen.py # Result report generator
├── reports/                    # Generated output files (Ignored by Git)
├── requirements.txt            # Python dependencies
├── service_key.json            # Google Service Account Key (Ignored by Git)
├── dev_log.md                  # Development history log
└── gemini.md                   # AI Assistant guidelines
```

## Setup

1.  **Clone the repository**
2.  **Install Dependencies**:
    ```bash
    pip install -r requirements.txt
    ```
3.  **Google Auth (Optional for some scripts)**:
    - Place your Google Service Account key as `service_key.json` in the root directory.

## Usage

### Generate Application Results Report
This script uses public CSV exports and does not require the service key.
```bash
python generators/mokil_high_school_results_gen.py
```
- Follow the interactive prompts to set the reference date.
- Generates HTML and Excel reports in the `reports/` directory.

### Generate Dashboard
This script requires `service_key.json` with appropriate permissions to the target Google Sheet.
```bash
python generators/generate_dashboard.py
```

## Security Note

- `reports/` directory is git-ignored to protect student privacy.
- `service_key.json` and other secrets are git-ignored.
