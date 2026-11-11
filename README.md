# summeryCollectBot

Collects a process's daily summary emails from Outlook for a given week and
builds an Excel report with one row per day plus a week-total row.

## Setup

1. Outlook desktop must be open and signed in.
2. `pip install -r requirements.txt`
3. Edit `config/processes.json` — add one entry per bot/process:

```json
"example_process_label": {
  "display_name": "Example Process (Label Mode)",
  "folder": "Inbox",
  "match_text": "replace with text unique to this process's email body",
  "subject_contains": null,
  "sender_contains": null,
  "fields": [
    { "name": "Total", "label": "Total", "occurrence": "first" },
    { "name": "Business Exceptions", "label": "Business Exception", "occurrence": "first" },
    { "name": "System Exceptions", "label": "System Exception", "occurrence": "first" }
  ]
}
```

- `match_text`: a snippet that only appears in this process's emails (case-insensitive, checked against the email body). This is how the bot tells processes apart.
- `subject_contains` / `sender_contains`: optional extra filters if `match_text` alone isn't unique enough.
- `folder`: `"Inbox"`, a subfolder like `"Inbox/Reports"`, or a shared mailbox like `"sharedbox@company.com/Inbox"`.
- `fields`: the numbers to pull out of the email. Two modes:
  - **`"label"` (default)** — for "Total: 42" style text, where the label sits right next to its number. Set `label` to the text to search for. `occurrence` controls which match to use if the label appears more than once: `first` (default), `last`, or `sum`.
  - **`"table_column"`** — for a table with a header row and one row per item (e.g. Clinic, Total, Success, Bus Exception, Sys Exception listing 7 clinics). Set `column_header` to the header text, and `aggregate` to either:
    - `"total_row"` — read the value from the email's own grand-total row (the row whose first cell contains `total_row_keyword`, default `"total"`)
    - `"sum_rows"` — ignore any total row in the email and sum the per-item rows ourselves

  Example (clinic table, reading the email's own total row):
  ```json
  { "name": "Total", "mode": "table_column", "column_header": "Total", "aggregate": "total_row", "total_row_keyword": "total" }
  ```

## Running

```
python main.py --process example_process_label
```

Defaults to the current Mon–Sun week. Options:

```
python main.py --process example_process_label --date 2026-11-10        # week containing this date
python main.py --process example_process_label --start 2026-11-09 --end 2026-11-16   # explicit range
python main.py --process example_process_label --output reports\weekly_report.xlsx
```

Output: an `.xlsx` with columns `Date, Email Count, <your fields...>` and a
bold `Week Total` row at the bottom.

## Adding a new bot/process

Add a new key under `"processes"` in `config/processes.json` with its own
`match_text` and `fields`. No code changes needed as long as the numbers sit
next to their label text in an HTML table or as plain text (e.g. `Total: 42`).

**Fastest way:** open [AI_CONFIG_PROMPT.txt](AI_CONFIG_PROMPT.txt), copy its
whole contents, paste it into an AI chat, then paste one sample email from
the new bot right after it. The AI will reply with a ready-to-paste JSON
block to drop into `config/processes.json` — or, if the email's format isn't
one of the two supported shapes (label-value pairs, or a header+rows table),
it will reply `UNSUPPORTED_FORMAT: <reason>` instead of guessing.
