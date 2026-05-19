# HireFlow AI — Automated Job Application Email Sender

HireFlow AI is a Python automation tool that sends personalized job application emails (with resume attachments) to company and HR emails extracted from an Excel sheet. It supports two role profiles — **Backend/Python Developer** and **Data Science/ML Engineer** — and tracks the delivery status back into the same Excel file.

---

## Features

- Sends emails to both company-level and HR/careers addresses
- Two email templates: Backend Developer and ML/DS roles
- Attaches the appropriate resume per template
- Deduplicates emails across columns before sending
- Tracks status per row: `SENT`, `PARTIAL`, `FAILED`, `NO EMAIL`
- Test mode to redirect all emails to a single test address
- Configurable batch size and delay between sends
- Auto-saves progress to Excel after each company

---

## Project Structure

```
hireflow-ai/
├── main.py               # Main script — reads Excel, sends emails, updates status
├── email_sender.py       # Low-level Gmail SMTP sender with attachment support
├── email_templates.py    # Email subject/body templates for Backend and ML/DS roles
├── hr_email.xlsx         # Input Excel file with company and HR emails (gitignored)
├── resumes/              # Resume PDFs (gitignored)
│   ├── Muhammad_Zaid.pdf     # Resume for Backend role
│   └── MuhammadZaid.pdf      # Resume for ML/DS role
├── .env                  # Credentials (gitignored)
└── requirements.txt      # Python dependencies
```

---

## Setup

### 1. Clone the repository

```bash
git clone https://github.com/your-username/hireflow-ai.git
cd hireflow-ai
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

Or manually:

```bash
pip install pandas openpyxl python-dotenv
```

### 3. Create a `.env` file

```env
SENDER_EMAIL=your_gmail@gmail.com
APP_PASSWORD=your_google_app_password
TEST_EMAIL=test_recipient@gmail.com
```

> **Note:** Use a [Google App Password](https://myaccount.google.com/apppasswords), not your regular Gmail password. Enable 2-Step Verification first.

### 4. Prepare the Excel file

Name it `hr_email.xlsx` and place it in the project root. Required columns:

| Column | Description |
|--------|-------------|
| `Company` | Company name (optional, used for personalization) |
| `Emails` | Direct company email addresses |
| `"hr@" + Domain` | Auto-generated HR email |
| `"careers@" + Domain` | Auto-generated careers email |

Multiple emails in a cell can be separated by commas, semicolons, newlines, or spaces.

### 5. Add your resumes

Place your resume PDFs in the `resumes/` folder:

```
resumes/Muhammad_Zaid.pdf     ← Backend role resume
resumes/MuhammadZaid.pdf      ← ML/DS role resume
```

---

## Configuration

Edit the config block at the top of `main.py`:

```python
TEST_MODE = False              # Set True to redirect all emails to TEST_EMAIL
MAX_COMPANIES_PER_RUN = 5     # Max companies to process per run (None = unlimited)
DELAY_SECONDS = 5             # Seconds to wait between emails
```

---

## Usage

```bash
python main.py
```

The script will:
1. Load `hr_email.xlsx` and skip rows already marked `SENT`
2. For each company, send a Backend and ML/DS email to every unique HR and company email
3. Update the Excel file with status columns after each company

### Status values written back to Excel

| Status | Meaning |
|--------|---------|
| `SENT` | All emails sent successfully |
| `PARTIAL` | Some emails sent, some failed |
| `FAILED` | All email attempts failed |
| `NO EMAIL` | No valid email address found in the row |

---

## Email Templates

Templates are defined in `email_templates.py` and personalized with the company name.

- **`backend_email(company)`** — applies for Backend / Python / Django / REST API roles
- **`ml_ds_email(company)`** — applies for Data Science / ML / AI Engineer roles

To customize the templates, edit the subject and body strings in `email_templates.py`.

---

## Security Notes

- `.env`, `resumes/`, and `hr_email.xlsx` are listed in `.gitignore` and will not be committed
- Never commit your App Password or email credentials
- Use Test Mode (`TEST_MODE = True`) to verify emails before a real send

---

## Requirements

- Python 3.8+
- Gmail account with App Password enabled
- `pandas`, `openpyxl`, `python-dotenv`
