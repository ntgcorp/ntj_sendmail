# NTJ SendMail

**Batch email sender for ntjobsOS** — Send multiple emails via Outlook, SMTP, or API (Resend/SendGrid) using a simple JSON configuration file.

## Overview

NTJ SendMail is a module of the **NTJOBS.APP v1** system, designed to run on the **ntjobsOS** batch orchestrator. It sends one or more emails defined in a JSON file using three interchangeable transport engines:

| Engine | Script | Use Case |
|--------|--------|----------|
| **Outlook COM** | `ntj_sendmail_olk.py` | Windows environments with Outlook installed |
| **SMTP** | `ntj_sendmail_smtp.py` | Any SMTP server (Gmail, Office365, custom) |
| **API** (planned) | `ntj_sendmail_api.py` | Resend, SendGrid, and other HTTP APIs |

All engines share the same JSON input format and validation logic — only the transport layer differs.

## Quick Start

### Prerequisites

- Python 3.10+
- `pywin32` for Outlook engine: `pip install pywin32`
- ntjobsOS environment (or standalone with `pyn.cmd` launcher)

### 1. Configure SMTP (Environment Variables)

```cmd
set SMTP_SERVER=smtp.gmail.com
set SMTP_USER=your_email@gmail.com
set SMTP_PASSWORD=your_app_password
set SMTP_PORT=587
set SMTP_SSL=FALSE
```

### 2. Prepare JSON Mail File

```json
{
  "MAIL_01": {
    "TO": ["recipient@example.com"],
    "CC": ["cc@example.com"],
    "SUBJECT": "Test Email",
    "FORMAT": "html",
    "BODY": "<h1>Hello World</h1>",
    "ATTACH": []
  }
}
```

### 3. Run with Launcher

```cmd
# Outlook engine
ntj_sendmail_olk.cmd ntjobs_sendmail_olk.ini

# SMTP engine
ntj_sendmail_smtp.cmd ntjobs_sendmail_smtp.ini
```

## Documentation

📖 **[Full User Manual](ntj_sendmail_man_en.md)** — Complete guide covering:
- Installation & prerequisites
- INI configuration (Outlook, SMTP, API)
- JSON mail format with examples
- Job commands & execution flow
- Environment variable usage
- Logging & troubleshooting
- Security & validation rules

## Package Contents

```
ntj_sendmail/
├── Core Scripts
│   ├── ntj_sendmail_olk.py      # Outlook COM engine
│   ├── ntj_sendmail_smtp.py     # SMTP engine
│   ├── ntj_sendmail_api.py      # API engine (planned)
│   ├── ncMailSimple.py          # SMTP wrapper class
│   ├── acJobsApp.py             # ntJobsApp orchestrator
│   └── aiSys.py                 # Support library (timestamp, config, I/O)
│
├── Configuration Examples
│   ├── ntjobs_sendmail_olk.ini      # Outlook INI example
│   ├── ntjobs_sendmail_smtp.ini     # SMTP INI example
│   ├── ntjobs_sendmail_resend.ini   # Resend API INI example
│   ├── test_maildemo_olk.ini        # Test INI (Outlook)
│   ├── test_maildemo_smtp.ini       # Test INI (SMTP)
│   └── test_maildemo_api.ini        # Test INI (API)
│
├── Launchers (Windows)
│   ├── ntj_sendmail_olk.cmd         # Outlook launcher
│   ├── ntj_sendmail_smtp.cmd        # SMTP launcher
│   ├── test_sendmail_olk.cmd        # Test launcher (Outlook)
│   ├── test_sendmail_smtp.cmd       # Test launcher (SMTP)
│   ├── olk_start.cmd                # Outlook start helper
│   └── olk_end.cmd                  # Outlook end helper
│
├── Documentation
│   ├── ntj_sendmail_man_en.md       # User manual (this repo)
│   ├── Prompt_sendmail.md           # Functional specification
│   ├── AGENTS.md                    # Agent instructions
│   └── opencode.json                # OpenCode agent config
│
├── Data
│   └── Json/
│       ├── maildemo.json            # Example mail JSON
│       └── ntrobot_exist.json       # Example mail JSON
│
└── Logs
    └── Log/                         # Runtime logs
```

## Architecture

```
┌─────────────────────────────────────┐
│         ntjobsOS Orchestrator       │
│    (searches jobs*.ini in USER_PATH)│
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│         acJobsApp Framework         │
│    (parses INI, expands $VARS,      │
│     validates, runs JOB_* sections) │
└──────────────┬──────────────────────┘
               │
       ┌───────┴───────┐
       ▼               ▼
┌─────────────┐ ┌─────────────┐
│ Outlook     │ │ SMTP        │
│ ntj_sendmail│ │ ntj_sendmail│
│ _olk.py     │ │ _smtp.py    │
└──────┬──────┘ └──────┬──────┘
       │               │
       ▼               ▼
┌─────────────┐ ┌─────────────┐
│ win32com    │ │ ncMailSimple│
│ .client     │ │ + smtplib   │
└─────────────┘ └─────────────┘
```

## Key Features

- **Multi-engine**: Switch transport via single `MAIL.ENGINE` parameter
- **Batch sending**: Multiple emails from one JSON file
- **Attachments**: File validation (existence + 5MB size limit)
- **Rate limiting**: Configurable `WAIT.START` and `WAIT.MAIL` delays
- **Variable expansion**: `$ENV.*`, `$SYS.*`, `$VAR` in INI files
- **Comprehensive logging**: Console + file with timestamps
- **Error handling**: Detailed error messages with procedure context

## License

See [LICENSE](LICENSE) file.

## Related Projects

- **ntjobsOS** — Batch orchestrator (parent system)
- **ntj_finance** — Financial quotes from Yahoo Finance/OpenFigi
- **ntj_fconvert** — File format conversion (PDF, Excel, JSON, YAML, etc.)

---

*Part of the ntjobsOS ecosystem — Modular, extensible, zero-dependency core*