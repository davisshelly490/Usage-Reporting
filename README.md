# Weekly AI Usage Report

This project creates the employee-level and summary workbook from the current roster and the five source exports in this project folder. It does not read or fall back to `Weekly Usage Report - 9-25-26.xlsx`.

## Source files

Replace these files in the project root before each refresh, keeping the file names unchanged:

- `roster.xlsx` — employee roster and reporting attributes
- `users_chatgpt.csv` — ChatGPT Enterprise users; enabled accounts are treated as licensed
- `usage_chatgpt.csv` — ChatGPT two-week usage; supplies message counts and last active date
- `users_copilot.csv` — Microsoft users; records with the `Microsoft 365 Copilot` SKU are treated as licensed
- `usage_copilot.csv` — Copilot usage; supplies each user's last activity date

The report date is the ChatGPT `period_end`. The ChatGPT and Copilot windows may end on different dates because their exports can arrive at different times; the workbook prints both source windows on the `Summary` sheet.

Before building the report, the script validates the ChatGPT and Copilot source windows independently. Their dates and lengths may differ. Metrics labeled "2 weeks" use the trailing 14 days of each product's respective source window, and metrics labeled "4 weeks" use the trailing 28 days. Copilot prompt counts remain labeled as the source-period total.

## Run

From the project root:

```bash
python build_weekly_usage_report.py
```

The refreshed dated workbook contains `Summary`, `Employee Level Data`, and `Prompt Matrix` sheets.

## People and filters

Full outer joins retain roster employees and accounts present only in any of the five source exports. Accounts are matched by normalized email; source records without email cannot be matched and are excluded. `In Roster` identifies roster matches. Source-only accounts keep their source name and email, with missing roster attributes left blank.

The Summary dropdowns filter Job, BU, Org 1 Report Category, Location, and Alternate Title. Select `All` to include all values or `(Blank)` for missing attributes. The selectors combine and recalculate Summary counts and rates. Licensed-user and inactivity metrics include source-only accounts; employee headcount and employee license percentages use roster employees. License capacity remains the workspace total.

The Employee Level Data sheet also has column filters for all five fields. Summary selectors do not filter that sheet or the Prompt Matrix. The matrix includes source-only accounts, using email when no roster identifier is available.

`ChatGPT Created Date (UTC)` comes from the user export's `Created Time (UTC)`. `new_chatgpt_user` is 1 for accounts created within the trailing 14 inclusive dates ending on the latest ChatGPT usage export end date, and 0 for earlier accounts. Missing or post-report creation dates stay blank. This flag never excludes anyone from Employee Level Data, Summary, or Prompt Matrix. Account creation is a proxy for tenure; the export does not supply a license assignment date.

Each run also generates `ChatGPT Inactive 4 Weeks`, a worksheet listing currently licensed ChatGPT users with `new_chatgpt_user = 0` and no activity in the four-week (28-day) usage window. It includes qualifying accounts outside the roster. Unknown or post-report account ages are excluded only from this targeted list. Summary filters do not affect this full list.

Dated workbooks are saved locally and in the configured shared Reporting folder as `Weekly AI Usage Report - YYYY-MM-DD.xlsx`, using the ChatGPT reporting end date.
