# NTJ SendMail — User Manual

## Overview

**NTJ SendMail** is a module of the **NTJOBS.APP v1** system for batch email sending on **ntjobsOS** infrastructure. It runs as an `ntJobsApp` orchestrated by `acJobsApp` (`acJobsApp.py`) and supported by `aiSys` (`aiSys.py`).

The application sends one or more emails read from a **JSON** parameter file, with three distinct engines:

| Script | Engine | INI Parameter `MAIL.ENGINE` | Library |
|--------|--------|----------------------------|---------|
| `ntj_sendmail_olk.py` | Outlook COM | `OLK` | `win32com.client` |
| `ntj_sendmail_smtp.py` | Generic SMTP | `SMTP` | `ncMailSimple.py` + `smtplib` |
| `ntj_sendmail_api.py` (planned) | External APIs | `RESEND` / `SENDGRID` | `resend` / `sendgrid` |

All apps share the same JSON file and validation/sending logic; only the transport layer changes.

---

## Prerequisites

### Environment Variables

Before running the scripts, ensure these environment variables are set:

| Variable | Description | Example |
|----------|-------------|---------|
| `PYTHON` | Path to Python launcher (`pyn.cmd`) | `k:\tools\` or `x pyn.cmd` |
| `NTJ_USER` | Current ntjobsOS user (injected by orchestrator) | `admin` |
| `NTJ_USERG` | User groups (comma-separated) | `ADMIN,OPERATORS` |

### Python Dependencies

- **Python 3.10+**
- **Outlook Engine**: `pywin32` (`pip install pywin32`)
- **SMTP Engine**: Standard library only (`smtplib`, `email`)
- **API Engine** (planned): `resend` or `sendgrid` packages

---

## Installation

1. Copy the `ntj_sendmail` folder to your `Apps/` directory in ntjobsOS:
   ```
   ntjobsOS/Apps/ntj_sendmail/
   ```

2. Ensure the following files are present:
   ```
   ntj_sendmail_olk.py
   ntj_sendmail_smtp.py
   ncMailSimple.py
   acJobsApp.py
   aiSys.py
   ```

3. Install `pywin32` for Outlook support:
   ```cmd
   pyn.cmd x pip install pywin32
   ```

---

## Launching with `pyn.cmd` and `.ini` Script

### Syntax

```cmd
pyn.cmd ntj_sendmail_olk.py  ntjobs_sendmail_olk.ini
pyn.cmd ntj_sendmail_smtp.py ntjobs_sendmail_smtp.ini
```

### Windows Launcher Scripts (`.cmd`)

The package includes ready-to-use launcher scripts:

**`ntj_sendmail_olk.cmd`**
```cmd
@ECHO OFF
ECHO Launcher Python Script %1
REM ----- Configurazione -----
set "PYTHON=k:\tools\"
REM ----- Ricava nome script python con PATH completo -----
set "SCRIPT=%~dp0%~n0.py"
REM ----- Esecuzione -----
ECHO Sintassi filein.json
call "%PYTHON%pyn.cmd" "%SCRIPT%" %*
```

**`ntj_sendmail_smtp.cmd`**
```cmd
REM @echo off
REM ----- Configurazione -----
set "PYTHON=x pyn.cmd"
REM ----- Ricava nome script python con PATH completo -----
set "SCRIPT=%~dp0%~n0.py"
REM ----- Esecuzione -----
echo Sintassi filein.json
call "%PYTHON%" "%SCRIPT%" %*
```

### Usage Examples

```cmd
# Outlook engine
ntj_sendmail_olk.cmd ntjobs_sendmail_olk.ini

# SMTP engine
ntj_sendmail_smtp.cmd ntjobs_sendmail_smtp.ini

# With custom INI path
pyn.cmd ntj_sendmail_olk.py "C:\MyConfigs\my_sendmail.ini"
```

---

## Configuration File (`.ini`)

### Common `[CONFIG]` Section (All Engines)

| Key | Variable | Type | Default | Notes |
|-----|----------|------|---------|-------|
| `MAIL.ENGINE` | `sMailEngine` | `str` | — | `OLK` \| `SMTP` \| `RESEND` \| `SENDGRID` — validated in `app_start` |
| `FILE.IN.JSON` | `sFileIn` | `str` | — | **Required**. Path to JSON mail file. Verified with `os.path.isfile` |
| `WAIT.START` | `nWaitStart` | `int` | `3` | `aiSys.StringToNum()`, if 0 → 3 |
| `WAIT.MAIL` | `nWaitMail` | `int` | `5` | `aiSys.StringToNum()`, if 0 → 5 |
| `DELIMITER` | `sDelimiter` | `str` | `,` | Delimiter for `TO/CC/CCN/ATTACH` lists when strings in JSON |
| `TYPE` | — | `str` | `NTJOBS.APP.1` | Validated by `acJobsApp.Start()` |
| `NAME` | — | `str` | — | Required |
| `LOG` | — | `str` | — | Log path (`jData.jLog`) |
| `EXIT` | — | `bool` | `TRUE` | If `TRUE` stops `Run()` on first error |

> `WAIT.START`/`WAIT.MAIL` are numeric strings in INI, converted to numbers in `app_start`. `DELIMITER` defaults to `,` if empty.

### SMTP-Specific Parameters (`ntj_sendmail_smtp.py`)

| Key | Type | Required | Default | Conversion |
|-----|------|----------|---------|------------|
| `SMTP.SERVER` | `str` | **Yes** | — | — |
| `SMTP.USER` | `str` | **Yes** | — | — |
| `SMTP.PASSWORD` | `str` | **Yes** | — | — |
| `SMTP.PORT` | `int` | No | `25` | `aiSys.StringToNum()` in `app_start_smtp` |
| `SMTP.SSL` | `bool` | No | `FALSE` | `aiSys.StringBool()` |
| `SMTP.TLS` | `bool` | No | — | `aiSys.StringBool()` (documented, `ncMailSimple` uses STARTTLS if `bSSL=False`) |

### API-Specific Parameters (`ntj_sendmail_api.py` - planned)

| Key | Type | Notes |
|-----|------|-------|
| `MAIL.API` | `str` | API key for `resend` or `sendgrid` (`jData.Config("MAIL.API")`) |

---

## INI Examples

### Outlook Engine (`ntjobs_sendmail_olk.ini`)

```ini
[CONFIG]
; General ntJobsApp Configuration
TYPE=NTJOBS.APP.1
NAME=SENDMAIL_NTJOBSOS
LOG=Log/ntjobsos_sendmail_smtp.log
EXIT=TRUE
; TO BE SPECIFIED
SENDMAIL_PATH=G:\Il mio Drive\Progetti\ntJobsAi2\Apps\ntj_sendmail

; --- Application Specific Configuration ---
; Defines the engine as required
MAIL.ENGINE=OLK

; Name of JSON file (must be in same folder or full path)
FILE.IN.JSON=$SENDMAIL_PATH/Json/ntjobsos.json

; Delimiter for email lists in JSON (default comma)
DELIMITER=,

; --- Wait Parameters ---
; Initial wait before starting (seconds)
WAIT.START=3
; Wait between emails (seconds)
WAIT.MAIL=5

[JOB_01_AVVIO]
; Explicitly starts Outlook (or attaches if already open)
COMMAND=OLK.START

[JOB_02_INVIO]
; Executes JSON read cycle and mail sending
COMMAND=SEND

; [JOB_03_CHIUSURA]
; Optional: Uncomment below to close Outlook at end
; COMMAND=OLK.KILL
```

**Environment Variable Usage**: The `$SENDMAIL_PATH` variable is expanded by `acJobsApp` using `$VAR` syntax from the same CONFIG section.

### SMTP Engine (`ntjobs_sendmail_smtp.ini`)

```ini
[CONFIG]
; General ntJobsApp Configuration
TYPE=NTJOBS.APP.1
NAME=SENDMAIL_NTJOBSOS
LOG=Log/ntjobsos_sendmail_smtp.log
EXIT=TRUE

; --- Application Specific Configuration ---
; Defines the SMTP engine as required
MAIL.ENGINE=SMTP

; Name of JSON file (must be in same folder or full path)
FILE.IN.JSON=Json/ntjobsos.json

; Delimiter for email lists in JSON (default comma)
DELIMITER=,

; --- Wait Parameters ---
; Initial wait before starting (seconds)
WAIT.START=3
; Wait between emails (seconds)
WAIT.MAIL=5

; --- SMTP Configuration ---
SMTP.SERVER=smtp.falcricrv.org
SMTP.USER=stefano.petrone@falcricrv.org
SMTP.PASSWORD=***
SMTP.PORT=587
SMTP.SSL=FALSE

[JOB_INVIO]
; Command that triggers app_send_mail function
COMMAND=SEND
```

### SMTP with Environment Variables (`test_maildemo_smtp.ini`)

```ini
[CONFIG]
TYPE=NTJOBS.APP.1
NAME=SENDMAIL_SMTP_DEMO
LOG=sendmail_smtp.log
EXIT=TRUE

MAIL.ENGINE=SMTP
FILE.IN.JSON=Json\maildemo.json
DELIMITER=,
WAIT.START=3
WAIT.MAIL=5

; SMTP credentials via environment variables
SMTP.SERVER=$ENV.SMTP_SERVER
SMTP.USER=$ENV.SMTP_USER
SMTP.PASSWORD=$ENV.SMTP_PASSWORD
SMTP.PORT=$ENV.SMTP_PORT
SMTP.SSL=$ENV.SMTP_SSL

[JOB_INVIO]
COMMAND=SEND
```

**Set environment variables before launch:**
```cmd
set SMTP_SERVER=smtp.gmail.com
set SMTP_USER=myemail@gmail.com
set SMTP_PASSWORD=myapppassword
set SMTP_PORT=587
set SMTP_SSL=FALSE
ntj_sendmail_smtp.cmd test_maildemo_smtp.ini
```

### Outlook with Environment Variables (`test_maildemo_olk.ini`)

```ini
[CONFIG]
TYPE=NTJOBS.APP.1
NAME=OLK_DEMO
LOG=olk_process.log
EXIT=TRUE

MAIL.ENGINE=OLK
FILE.IN.JSON=json\maildemo.json
DELIMITER=,
WAIT.START=3
WAIT.MAIL=5

[JOB_01_AVVIO]
COMMAND=OLK.START

[JOB_02_INVIO]
COMMAND=SEND

; [JOB_03_CHIUSURA]
; COMMAND=OLK.KILL
```

---

## JSON Mail File Format

### Structure

Each top-level key (except `CONFIG`) is the **unique mail ID**. Value = dictionary with case-insensitive fields (normalized to `UPPER` in `app_argread`).

```json
{
  "mail_1": {
    "TO": ["ntgcorp@gmail.com"],
    "CC": [],
    "CCN": [],
    "SUBJECT": "Mail from ntjobsos",
    "FORMAT": "txt",
    "BODY": "Mail example",
    "ATTACH": []
  },
  "MAIL_02": {
    "TO": "a@b.it, c@d.it",
    "SUBJECT": "Report",
    "FORMAT": "html",
    "BODY": "<strong>Hello</strong>",
    "ATTACH": "C:/tmp/a.pdf, C:/tmp/b.pdf"
  }
}
```

### Fields per Record

| Field | JSON Type | Required | Description |
|-------|-----------|----------|-------------|
| `TO` | `array` or `string` | **Yes** | Recipients. If string → `split(sDelimiter)` |
| `CC` | `array` or `string` | No | CC recipients |
| `CCN` | `array` or `string` | No | BCC/CCN recipients |
| `SUBJECT` | `str` | No | Email subject |
| `FORMAT` | `str` | No | `txt` (default) or `html` |
| `BODY` | `str` | No | Text or HTML (already JSON-escaped) |
| `ATTACH` | `array` or `string` | No | Attachment file paths. If string → `split(sDelimiter)` |

### Normalization Rules (`app_argread`)

- First-level keys + internal keys → `strip().upper()` (e.g., `to` → `TO`, `mail_1` → `MAIL_1`)
- String values → `strip().strip('"')` (removes outer quotes if present)
- If `TO`/`CC`/`CCN`/`ATTACH` is string → `[x.strip() for x in v.split(sDelimiter) if x.strip()]`
- If already list → `strip()` on each element if string
- Result: `dictMail = {"MAIL_01": {"TO": [...], ...}, "MAIL_02": {...}}`
- Updates `sTimestamp = aiSys.Timestamp()` after successful read

### Example JSON (`Json/maildemo.json`)

```json
{
  "mail_1": {
    "id": "mail_001",
    "to": ["ntgcorp@gmail.com"],
    "cc": ["stefano.petrone@falcricrv.org"],
    "ccn": [],
    "subject": "First test email - Nature poem",
    "format": "txt",
    "body": "The wind whispers through the trees,\nthe leaves dance in the blue sky,\nthe sun sets behind the mountains,\nand the moon rises slowly.\n\nNature speaks to us in silence,\nwith its infinite beauty,\nreminds us how precious\nis every moment of this life.",
    "attach": []
  },
  "mail_2": {
    "id": "mail_002",
    "to": ["stefano.petrone@falcricrv.org", "ntgeshop@gmail.com"],
    "cc": [],
    "ccn": ["ntgcorp@gmail.com"],
    "subject": "Second test email - Sea poem",
    "format": "html",
    "body": "<html><body><h2>The Sea</h2><p><i>Waves break on the shore,<br/>bringing distant stories,<br/>salt in the air scents summer,<br/>while seagulls fly serenely.</i></p><p><i>The infinite horizon calls us,<br/>towards unknown lands and adventures,<br/>the sea is freedom and mystery,<br/>keeper of treasures and fears.</i></p></body></html>",
    "attach": []
  }
}
```

---

## Job Commands (INI `[JOB_*]` Sections)

| COMMAND | Action | Script |
|---------|--------|--------|
| `SEND` | `sResult = app_sendmail()` | All |
| `OLK.START` | `sResult = app_olk_start()` | Outlook only |
| `OLK.KILL` | `sResult = app_olk_kill()` | Outlook only |
| Other | `aiSys.ErrorProc("Unrecognized command ...", sProc)` | — |

The dispatcher `cbCommand` / `cbCommands` reads `COMMAND` from each `JOB_*` section.

---

## Execution Flow

```
1. Launch: pyn.cmd ntj_sendmail_*.py script.ini
   ↓
2. app_start()
   ├─ jData.Start() → reads .ini, expands $SYS.*/$ENV.*/$VAR, validates CONFIG
   ├─ app_argread() → reads JSON → dictMail
   ├─ app_checkmail() → validates TO/CC/CCN/ATTACH
   ├─ (SMTP) app_start_smtp() → jMail.Start() → SMTP login
   └─ jData.Run(cbCommand)
       ├─ JOB_01: COMMAND=OLK.START → app_olk_start()  [Outlook only]
       ├─ JOB_02: COMMAND=SEND → app_sendmail()
       │     ├─ sleep(WAIT.START) once
       │     ├─ for each mail in dictMail (skip CONFIG)
       │     │   ├─ app_sendmail_olk() or app_sendmail_smtp() or app_sendmail_api()
       │     │   └─ sleep(WAIT.MAIL) between emails
       │     └─ log "Mails sent: N | Total time: HH:MM:SS"
       └─ JOB_03: COMMAND=OLK.KILL → app_olk_kill()  [Optional Outlook]
   ↓
3. jData.End(sResult) → writes .end file (RETURN.TYPE=S|E, RETURN.VALUE, TS.START/TS.END)
```

---

## Logging

- **Console**: `print(f"{aiSys.Timestamp()} - Executing {sProc}")` at start of each `app_*`; errors via `aiSys.ErrorProc`
- **File**: `jData.jLog` (`acLog`) writes to `LOG/` defined by `LOG=` in INI; format `YYYYMMDD:HHMMSS:TYPE: message`
- **Timestamp**: Always `aiSys.Timestamp()` format `YYYYMMDD:HHMMSS` (e.g., `20260722:143055`); diff via `aiSys.TimestampDiff()` + `format_timedelta()`

---

## Security & Validation

- **Email validation**: `aiSys.isEmail(email)` for every `TO`/`CC`/`CCN` entry
- **Attachment validation**: existence + size `<= nMaxSizeAttach` (5 MB); path normalized
- **INI validation**: `acJobsApp.Start()` verifies `TYPE`, `NAME`, `CONFIG`, reserved keys `RETURN.*`, `TS.*`, `FILE.*`
- **Password**: removed from `dictJobs["CONFIG"]` after `Start()` (not written to `.end`)
- **Errors**: always `aiSys.ErrorProc(sResult, sProc)` with procedure prefix

---

## Troubleshooting

### Common Issues

| Issue | Cause | Solution |
|-------|-------|----------|
| `File json non presente` | `FILE.IN.JSON` path wrong | Verify path, use `$VAR` or `$ENV.` expansion |
| `Motore Outlook non avviato` | Missing `OLK.START` job | Add `JOB_01` with `COMMAND=OLK.START` |
| `SMTP non avviato` | `app_start_smtp` failed | Check `SMTP.SERVER`, `USER`, `PASSWORD`, `PORT`, `SSL` |
| `Email non valida` | Malformed email in JSON | Validate with `aiSys.isEmail()` format |
| `File attach non presente` | Attachment path wrong | Use `NormalizePath()`, verify file exists |
| `pywin32` not installed | Outlook engine requires it | Run `pyn.cmd x pip install pywin32` |

### Debug Mode

Add `EXIT=FALSE` in INI to continue on errors and see all job results.

---

## File List (Package Contents)

Based on `pyt_robocopy_sendmail_github.json`:

```
ntj_sendmail/
├── ntj_sendmail_olk.py          # Outlook engine app
├── ntj_sendmail_smtp.py         # SMTP engine app
├── ntj_sendmail_api.py          # API engine app (planned)
├── ncMailSimple.py              # SMTP wrapper class (used by SMTP)
├── acJobsApp.py                 # Orchestrator (local copy)
├── aiSys.py                     # Support library (local copy)
├── Prompt_sendmail.md           # Functional specification (source of truth)
├── AGENTS.md                    # Agent instructions
├── opencode.json                # Agent config (model, language)
├── ntjobs_sendmail_olk.ini      # Example INI Outlook
├── ntjobs_sendmail_smtp.ini     # Example INI SMTP
├── ntjobs_sendmail_resend.ini   # Example INI API Resend
├── ntj_sendmail_olk.cmd         # Windows launcher (Outlook)
├── ntj_sendmail_smtp.cmd        # Windows launcher (SMTP)
├── test_sendmail_olk.cmd        # Test launcher (Outlook)
├── test_sendmail_smtp.cmd       # Test launcher (SMTP)
├── test_maildemo_olk.ini        # Test INI (Outlook)
├── test_maildemo_smtp.ini       # Test INI (SMTP)
├── test_maildemo_api.ini        # Test INI (API)
├── olk_start.cmd                # Outlook start helper
├── olk_end.cmd                  # Outlook end helper
├── Json/
│   ├── maildemo.json            # Example JSON mail
│   ├── ntrobot_exist.json       # Example JSON
│   └── file_tracker*.json       # Tracking files
└── Log/
    └── *.log                    # Operation logs
```

---

## Support & References

1. `Prompt_sendmail.md` — Complete functional specification (primary source)
2. `acJobsApp.py` — ntJobsApp orchestrator (Start/Run/End/Config/Expand/NormalizePath)
3. `aiSys.py` — Timestamp, isEmail, StringToNum/StringBool, ErrorProc, ExpandDict
4. `ncMailSimple.py` — SMTP wrapper (Start/Send/Logoff/VerifyConnect)
5. `ntjobsOS/AGENTS.md` — Global orchestrator (Inbox/Archive flow, ntjobs_config.ini)
6. API Docs: [Resend](https://resend.com/docs) · [SendGrid](https://docs.sendgrid.com/)

---

*Manual generated for NTJ SendMail v1.0 — Part of ntjobsOS ecosystem*