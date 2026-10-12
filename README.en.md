# Schedule Deviation Early-Warning System (AI Edition)

An automatic **schedule-deviation early-warning** tool for construction and engineering projects. It reads two Excel workbooks — the **master schedule plan** and the **daily site reports** — automatically compares planned vs. actual progress for both parent tasks and sub-tasks, and produces a three-level (red / amber / green) warning report. The report can be emailed to the project team in one step.

The command-line edition additionally integrates **LLM analysis**: it feeds the day's structured data to any OpenAI-compatible endpoint, which generates a problem diagnosis and corrective recommendations — removing the manual effort of writing the analysis by hand. When no LLM is configured, the tool automatically falls back to rule-based analysis, so every feature remains available.

---

## Features

- **Parent-task comparison**: compares the daily report's "Progress Ledger" section 1.1 summary row against the master schedule, computing planned completion %, actual completion %, deviation, and deviation in days
- **Sub-task breakdown**: covers all categorized sub-tasks in sections 1.2 – 1.6 of the daily report, flagging each red / amber / green by schedule deviation
- **Crew-and-zone comparison**: aggregates completion rates against deadlines, median work intensity, and other comparable metrics per zone (A / B / C / D …) and per crew
- **Key milestone reminders**: highlights milestones falling within the next two weeks, and separately flags milestones already overdue
- **Upcoming-week task start-up warnings**: gives advance warning about tasks starting next week that may be constrained by unmet prerequisites
- **LLM AI analysis** (command-line edition): automatically produces a problem list with actionable recommendations
- **Email delivery**: generates an HTML warning report and sends the body plus the report as an attachment
- **Flask web edition**: upload files and view results in a browser — suitable for colleagues who are not comfortable with the command line

## Project Structure

```
.
├── schedule_warning.py        # command-line entry point (includes LLM analysis)
├── run_warning.bat            # one-click runner for the CLI edition (Windows)
├── requirements.txt           # Python dependencies
├── .env.example               # environment variable template (copy to .env and fill in)
├── .gitignore
├── master_plan/               # put the master schedule workbook here (input)
├── key_nodes/                 # put the key milestone workbook here (input, optional)
├── daily_reports/             # put daily site report workbooks here (input)
├── warning_output/            # generated reports and email previews (created automatically)
└── web/                       # Flask web edition (rule-based, no LLM)
    ├── app.py                 # Flask backend
    ├── schedule_warning_core.py  # core comparison logic module
    ├── templates/index.html   # front-end page
    └── run_web.bat            # one-click launcher for the web service (Windows)
```

## Requirements

- Python 3.8+
- Dependencies are listed in `requirements.txt`

```bash
pip install -r requirements.txt
```

The command-line edition only needs `openpyxl` plus the Python standard library; the web edition additionally needs `flask`.

---

## Quick Start

### 1. Prepare the input files

Place the three types of Excel workbook into their directories (file names are configurable):

| Directory | Contents | Required |
|-----------|----------|----------|
| `master_plan/` | Master schedule plan workbook | Yes |
| `daily_reports/` | Daily site report (file name contains a date, e.g. `工程日报 2026.10.11.xlsx`) | Yes |
| `key_nodes/` | Key milestone summary workbook | Optional |

> By default the tool looks in `daily_reports/` for "yesterday's" report based on its date. If it cannot find one, it falls back to the most recent file in the directory.
> The daily report file name must contain one of the keywords listed in `DAILY_FILE_KEYWORDS` (default: `工程日报`, Chinese for "site daily report"). You can change this to match your own naming.

### 2. Configure environment variables

Copy `.env.example` to `.env` (or set the variables directly in your system environment) and fill in what you need:

- **Required** (only if you want email delivery): `SMTP_USER`, `SMTP_PASS` (an app-specific authorization code, not your login password), `EMAIL_RECIPIENTS`
- **Optional** (to enable AI): `LLM_API_KEY`, `LLM_API_URL`, `LLM_MODEL`
- The rest cover paths, project name, and the structural mapping of your daily report / plan workbook. The defaults are usually fine.

> The repository contains **no keys and no real paths** — everything is injected through environment variables. Without SMTP configured the tool only generates the report and skips email; without `LLM_API_KEY` it disables the LLM and falls back to rule-based analysis.

> **Note on Chinese defaults:** the built-in default values for the project name, sheet name, file-name keyword, and the default input workbook names are in Chinese, because the tool was originally built for Chinese-language site reports. Override them via `PROJECT_NAME`, `DAILY_SHEET_NAME`, `DAILY_FILE_KEYWORDS`, `PLAN_FILE`, and `KEYNODE_FILE` to match your own convention.

### 3. Run the command-line edition

```bash
python schedule_warning.py
```

On Windows you can also simply double-click `run_warning.bat`. After a run:

- The report is written to `warning_output/工期预警报表_<date>.xlsx`
- If email is configured, the report is pushed automatically
- Setting `DUMP_HTML=1` additionally exports an HTML email preview into `warning_output/`

### 4. Run the web edition

```bash
cd web
python app.py
```

Or double-click `web/run_web.bat`, then open <http://127.0.0.1:8080> in your browser and upload the schedule plan and daily report as prompted.

---

## Expected Input Workbook Formats

### Daily site report (`daily_reports/*.xlsx`)

- Default worksheet name: `进行台账` (Chinese for "progress ledger"), overridable with `DAILY_SHEET_NAME`; falls back to the first worksheet if not found
- The row whose column B equals `序号` ("No.") is treated as the header row; sections are numbered `1.1 / 1.2 …`
- Main columns (as expected by `parse_daily` in the command-line edition):

  | Column | Field |
  |--------|-------|
  | E | Task name |
  | F | Total quantity |
  | G | Unit |
  | H / I | Planned start / planned finish |
  | J / K | Actual start / actual finish |
  | L | Status |
  | O | Completion percentage |
  | P | Remarks |

### Master schedule plan (`master_plan/*.xlsx`)

The command-line edition's `load_plan` auto-detects two layouts:

- **New layout**: the header row appears within the first 8 rows and contains both a "task name" column and "start time / finish time" columns. Hierarchy is inferred from the indentation of the task-name cell (indent 0 → L1, 1 → L2, …). The worksheet name can be pinned with `PLAN_SHEET_NAMES`; otherwise the first worksheet is used.
- **Legacy layout**: task name in column 5, hierarchy level in column 2, duration in column 6, start in column 7, finish in column 8.

### Key milestones (`key_nodes/*.xlsx`, optional)

Uses the first worksheet, starting from row 2: column 2 = milestone name, column 3 = date, column 4 = category.

---

## LLM Configuration

The command-line edition calls the model through an OpenAI-compatible `chat/completions` endpoint, so it works with DeepSeek, Qwen (Tongyi), Zhipu, Kimi, and self-hosted gateways alike:

| Variable | Default | Description |
|----------|---------|-------------|
| `LLM_API_URL` | `https://api.deepseek.com/v1/chat/completions` | Endpoint URL |
| `LLM_API_KEY` | empty (LLM disabled) | API key — **leave empty to fall back to rule-based analysis** |
| `LLM_MODEL` | `deepseek-chat` | Model name; a non-reasoning model is recommended |

> A **non-reasoning model** is recommended: a reasoning model's thinking tokens are unpredictable and can consume the entire `LLM_MAX_TOKENS` budget, leaving the response body empty. Timeouts, token limits, temperature, and related parameters live in the configuration block of `schedule_warning.py`.

---

## Command-Line Edition vs. Web Edition

| | CLI edition `schedule_warning.py` | Web edition `web/` |
|---|---|---|
| LLM AI analysis | ✅ Supported | ❌ Rule-based only |
| Interaction | Command line / scheduled job | Browser upload |
| Parent-task matching | Dynamically parses daily-report sections | Fixed row numbers via a `MAPPING` config |
| Best for | Automation, daily scheduled push | Ad-hoc manual comparison, demos |

> The web edition's `schedule_warning_core.py` relies on a fixed row-number `MAPPING` configuration tailored to one specific daily-report template; the command-line edition is the more general one. For further development, start from the command-line edition.

---

## Security Notes

- This repository contains no API keys, mailbox passwords, real project names, or local paths — everything is injected through environment variables.
- `.env` is listed in `.gitignore` — **never commit any file containing credentials**.
- The web edition listens only on `127.0.0.1` by default. If you expose it to a public network, tighten the CORS allow-list and add authentication first — otherwise anyone can upload files and trigger email sending.
- Use the **authorization code** your mail provider issues for `SMTP_PASS` (163 / QQ / 126 and similar require you to request one separately in the mailbox settings), not your login password.

## License

[MIT](LICENSE)
