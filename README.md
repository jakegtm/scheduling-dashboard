# GTM Scheduling Analyzer

A comprehensive Streamlit-based dashboard for analyzing project scheduling, budget allocation, resource utilization, and actual-vs-scheduled variance tracking.

## Overview

The **GTM Scheduling Analyzer** is designed to help teams:
- **Track project status** and identify missing rates or issues
- **Analyze budget-to-actual comparisons** and flag over-budget projects
- **Monitor resource utilization** by month and role
- **Detect scheduling variances** by comparing planned hours against actual hours from OpenAir
- **Send consolidated email reports** to stakeholders with personalized insights

## Features

### 📋 Project Tracker
Monitor known projects and identify issues:
- View all projects with missing rates or status concerns
- Track TBD (To Be Determined) and pending SOW (Statement of Work) projects
- Identify project owners and their responsibilities
- Flag projects requiring immediate attention

### 💰 Budget to Actual
Analyze project financials:
- Compare budgeted amounts to actual spending
- Flag projects over budget (negative remaining)
- Identify projects with significant unscheduled remaining budget
- Configurable thresholds for cost-based alerts

### 📈 Utilization Dashboard
Monitor resource allocation:
- View monthly utilization rates by role and individual
- Track chargeable hours, holidays, and PTO
- Compare actual utilization against organizational goals
- Identify staffing gaps and over-allocation

### 📊 Variance Analysis
Detect scheduling anomalies:
- Compare scheduled hours to actual hours from OpenAir reports
- Flag staff members with significant variances (over/under-work)
- View per-project and per-period breakdowns
- Customizable thresholds for variance detection
- Support for half-month periods throughout the year

### 📧 Combined Email Reports
Automated stakeholder communication:
- Generate personalized HTML email reports per team member
- Include relevant project tracker issues, budget concerns, and variance alerts
- Support for interns (receive only their own variance rows)
- Support for project owners (receive all issues across their projects)
- Batch email sending via SendGrid

## Technology Stack

- **Python 3.11+**
- **Streamlit** — Interactive web dashboard framework
- **OpenPyXL** — Excel/XLSX file parsing
- **SendGrid** — Email delivery (optional)

## Installation

### Prerequisites
- Python 3.11 or higher
- pip (Python package manager)

### Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/jakegtm/scheduling-dashboard.git
   cd scheduling-dashboard
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **(Optional) Configure SendGrid for email:**
   Create or update `.streamlit/secrets.toml`:
   ```toml
   [email]
   sendgrid_api_key = "your-sendgrid-api-key"
   from_email = "your-sender@example.com"

   [auth]
   username = "gtmtas"
   password_hash = "your-sha256-password-hash"
   ```

4. **Update configuration:**
   Edit `config.py` to match your organization:
   - `EMAIL_LOOKUP` — Staff email addresses
   - `PERSON_ROLE` — Employee role codes (MD, DR, SM, MR, SR, AN, IN)
   - `FIRST_NAMES` — First names for email greetings
   - `INTERN_NAMES` — Staff who see only their own rows
   - `STAFF_NAMES` — Staff who are not project managers

## Usage

### Running the App

```bash
streamlit run app.py
```

The app will open in your browser at `http://localhost:8501`

### File Format Requirements

#### Schedule File (.xlsx)
- **Budget to Actual sheet** — Contains:
  - Column H: Project Owner name
  - Budget and remaining columns for financial tracking
  
- **Project Tracker sheet** — Contains:
  - Client name, project codes, status, owner, budget
  - Rate information and SOW details

- **Month sheets** (e.g., "January", "February", etc.) — Contains:
  - Row 2: Staff names (starting from column G)
  - Columns G-AO: Scheduled hours by project for each person
  - Support for purple-filled cells to mark "done" projects

- **Utilization by Month sheet** — Contains:
  - Person name, role code
  - Chargeable hours, holidays, PTO, utilization %

#### OpenAir Report (.csv or .xlsx)
- CSV format with person, project, actual hours by period
- Used for variance calculations against scheduled hours
- Optional — if not provided, variance shows scheduled vs. 0 actual

### Workflow

1. **Sign In** — Enter credentials (default: `gtmtas`)
2. **Upload Files** — Choose schedule file and optionally OpenAir report
3. **Run Analysis** — Click to parse and process all sheets
4. **Review Tabs**:
   - Project Tracker issues
   - Budget-to-Actual flags
   - Utilization rates
   - Variance analysis with customizable thresholds
5. **Configure Email Recipients** — Select which team members receive reports
6. **Preview & Send** — Review HTML previews and send batch emails

### Settings & Thresholds

**Budget Thresholds:**
- "Flag unscheduled remaining over" — Minimum budget gap to flag
- "Flag negative budgets below" — Maximum negative budget to flag

**Variance Thresholds:**
- "Scheduled but not actual" — Hours under-worked threshold
- "Actual but not scheduled" — Hours over-worked threshold

## Project Structure

```
scheduling-dashboard/
├── app.py                   # Main Streamlit application
├── config.py                # Configuration (emails, roles, thresholds)
├── email_utils.py           # Email building and SendGrid integration
├── requirements.txt         # Python dependencies
├── .python-version          # Python 3.11 specification
├── assets/                  # Logo and static files
├── processors/              # Analysis modules
│   ├── budget_actual.py     # Budget-to-actual analysis
│   ├── project_tracker.py   # Project tracking
│   ├── utilization.py       # Resource utilization
│   └── variance.py          # Variance calculations
├── .devcontainer/           # Dev container configuration
└── .github/                 # GitHub workflows
```

## Configuration Guide

### EMAIL_LOOKUP (config.py)
Map staff names to email addresses. Keys must match exactly as they appear in your schedule file:
```python
EMAIL_LOOKUP = {
    "J. O'Donnell": "jodonnell@example.com",
    "S. O'Donnell": "sodonnell@example.com",
    # Add all team members
}
```

### PERSON_ROLE (config.py)
Define role codes for organizational hierarchy:
```python
PERSON_ROLE = {
    "Sorrentino":   "MD",   # Managing Director
    "Colonna":      "DR",   # Director
    "Hendrickson":  "SM",   # Senior Manager
    # ... etc
}
```

### Role-Based Email Routing
- **Project Owners** (not in `STAFF_NAMES`) receive all variance rows for their projects
- **Staff Members** (in `STAFF_NAMES`) receive only their own variance rows
- **Interns** (in `INTERN_NAMES`) receive minimal variance information

## Troubleshooting

### "Could not open schedule file"
- Ensure file is valid XLSX format
- Check that required sheets exist (Budget to Actual, Project Tracker, or month tabs)

### "Email credentials not configured"
- Add SendGrid API key to `.streamlit/secrets.toml`
- Email sending will be disabled until configured

### Missing staff in email list
- Verify names in `EMAIL_LOOKUP` match exactly how they appear in the schedule
- Check `FIRST_NAMES` for email greeting customization

### No variance data showing
- Ensure OpenAir report is uploaded if you want actual hours
- Check that selected periods have data in the schedule

## Security Notes

- **Authentication** is handled via username/password hash (SHA-256)
- **Streamlit Secrets** store sensitive credentials (SendGrid API keys)
- **Credentials are NOT stored in the repository** — configure via environment
- Fallback credentials exist for local development (see `app.py` lines 104-117)

## Performance Optimization

The app uses Streamlit's caching to optimize performance:
- `@st.cache_resource` — Caches workbooks and complex objects
- `@st.cache_data` — Caches serializable results (lists, tuples)
- Memory cleanup via `gc.collect()` after large operations

## Development

### Dependencies
- See `requirements.txt` for all packages
- Python 3.11 recommended (see `.python-version`)

### Contributing
When adding new features:
1. Update relevant processor modules in `processors/`
2. Add configuration options to `config.py`
3. Test with sample schedule files
4. Update this README with new functionality

## License

No license specified. Check with repository owner for usage rights.

## Support

For issues or questions:
- Check the **Troubleshooting** section above
- Review `app.py` comments for implementation details
- Ensure all configuration in `config.py` is correct
- Verify input file formats match the specifications

---

**Version**: 1.0  
**Last Updated**: May 2026
