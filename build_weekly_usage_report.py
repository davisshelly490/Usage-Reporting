from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from numbers import Real
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
PUBLISHED_OUTPUT_DIR = Path(r"C:\Users\DBShelly\OneDrive - The Haskell Company\Haskell Dysruptek - General\AI\Reporting")

# Live exports.  Replace these files in place before each refresh.
EMPLOYEE_ROSTER_FILE = BASE_DIR / "roster.xlsx"
COPILOT_USERS_FILE = BASE_DIR / "users_copilot.csv"
CHATGPT_USERS_FILE = BASE_DIR / "users_chatgpt.csv"
COPILOT_USAGE_FILE = BASE_DIR / "usage_copilot.csv"
CHATGPT_USAGE_FILE = BASE_DIR / "usage_chatgpt.csv"
CHATGPT_USAGE_FOUR_WEEKS_FILE = BASE_DIR / "usage_chatgpt_four_weeks.csv"

CHATGPT_LICENSE_CAPACITY = 1000
COPILOT_LICENSE_CAPACITY = 900
FILTER_FIELDS = ["Job", "BU", "Org 1 Report Category", "Location", "Alternate Title"]

BASE_EMPLOYEE_COLUMNS = [
    "Team Member Number",
    "Team Member Name",
    "Employee Email",
    "Original Hire",
    "Last Hire",
    "Date In Job",
    "Job",
    "BU",
    "Org 1 Report Category",
    "Location",
    "Alternate Title",
]

OUTPUT_EMPLOYEE_COLUMNS = BASE_EMPLOYEE_COLUMNS + [
    "Copilot",
    "ChatGPT",
    "AI License",
    "Last Copilot Activity",
    "More than 2 weeks (Copilot)",
    "Copilot Prompts (Source Period)",
    "Messages in 2 Weeks",
    "Active in 2 Weeks (ChatGPT)",
    "No Activity in 2 Weeks (ChatGPT)",
    "Both 2 Weeks",
    "Last Usage (ChatGPT)",
    "More than 2 Weeks (ChatGPT)",
    "Messages in 4 Weeks",
    "Active in 4 Weeks (ChatGPT)",
    "No Activity in 4 Weeks (ChatGPT)",
    "Both 4 Weeks",
    "Last Usage (ChatGPT, 4 Weeks)",
    "More than 4 Weeks (ChatGPT)",
    "In Roster",
    "ChatGPT Created Date (UTC)",
    "new_chatgpt_user",
]


def normalize_email(series: pd.Series) -> pd.Series:
    normalized = series.astype("string").str.strip().str.lower()
    return normalized.mask(normalized.eq(""), pd.NA)


def normalize_bu(value):
    """Store BU identifiers as text, without decimals added by numeric imports."""
    if pd.isna(value):
        return pd.NA
    if isinstance(value, Real) and float(value).is_integer():
        return str(int(value))
    text = str(value).strip()
    if "." in text:
        whole, fraction = text.split(".", 1)
        if whole.isdigit() and fraction and set(fraction) == {"0"}:
            text = whole
    return text if text else pd.NA


def unique_employee_count(df: pd.DataFrame) -> int:
    """Count employees by the roster's stable identifier, including people without email."""
    employee_ids = df["Team Member Number"].astype("string").str.strip()
    return int(employee_ids.mask(employee_ids.eq(""), pd.NA).nunique(dropna=True))


@dataclass(frozen=True)
class ReportingWindows:
    chatgpt_start: pd.Timestamp
    chatgpt_end: pd.Timestamp
    copilot_start: pd.Timestamp
    copilot_end: pd.Timestamp
    chatgpt_four_week_source_start: pd.Timestamp
    chatgpt_four_week_source_end: pd.Timestamp

    @property
    def max_chatgpt_report_date(self) -> pd.Timestamp:
        return max(self.chatgpt_end, self.chatgpt_four_week_source_end)

    @property
    def chatgpt_days(self) -> int:
        return (self.chatgpt_end - self.chatgpt_start).days + 1

    @property
    def copilot_days(self) -> int:
        return (self.copilot_end - self.copilot_start).days + 1

    @property
    def chatgpt_two_week_start(self) -> pd.Timestamp:
        return self.chatgpt_end - pd.Timedelta(days=13)

    @property
    def copilot_two_week_start(self) -> pd.Timestamp:
        return self.copilot_end - pd.Timedelta(days=13)

    @property
    def chatgpt_four_week_source_days(self) -> int:
        return (self.chatgpt_four_week_source_end - self.chatgpt_four_week_source_start).days + 1

    @property
    def chatgpt_four_week_start(self) -> pd.Timestamp:
        return self.chatgpt_four_week_source_end - pd.Timedelta(days=27)

    @property
    def chatgpt_four_week_end(self) -> pd.Timestamp:
        return self.chatgpt_four_week_source_end

    @property
    def copilot_four_week_start(self) -> pd.Timestamp:
        return self.copilot_end - pd.Timedelta(days=27)


def validate_reporting_window(
    chatgpt_usage: pd.DataFrame, copilot_usage: pd.DataFrame, chatgpt_four_week_usage: pd.DataFrame
) -> ReportingWindows:
    """Validate the two source windows without requiring matching dates or lengths."""
    required_chatgpt = {"period_start", "period_end"}
    required_copilot = {"Report Refresh Date", "Report Period", "Last Activity Date"}
    missing_chatgpt = required_chatgpt.difference(chatgpt_usage.columns)
    missing_copilot = required_copilot.difference(copilot_usage.columns)
    if missing_chatgpt:
        raise ValueError(f"ChatGPT usage export is missing columns: {sorted(missing_chatgpt)}")
    if missing_copilot:
        raise ValueError(f"Copilot usage export is missing columns: {sorted(missing_copilot)}")

    missing_chatgpt_four_week = required_chatgpt.difference(chatgpt_four_week_usage.columns)
    if missing_chatgpt_four_week:
        raise ValueError(
            f"ChatGPT four-week usage export is missing columns: {sorted(missing_chatgpt_four_week)}"
        )

    starts = pd.to_datetime(chatgpt_usage["period_start"], errors="coerce").dropna().dt.normalize().unique()
    ends = pd.to_datetime(chatgpt_usage["period_end"], errors="coerce").dropna().dt.normalize().unique()
    if len(starts) != 1 or len(ends) != 1:
        raise ValueError("ChatGPT usage export must contain exactly one valid period_start and period_end.")

    period_start = pd.Timestamp(starts[0])
    period_end = pd.Timestamp(ends[0])
    inclusive_days = (period_end - period_start).days + 1
    if inclusive_days != 14:
        raise ValueError(
            f"ChatGPT reporting window must contain exactly 14 inclusive days; "
            f"found {inclusive_days} days ({period_start.date()} through {period_end.date()})."
        )

    four_week_starts = (
        pd.to_datetime(chatgpt_four_week_usage["period_start"], errors="coerce").dropna().dt.normalize().unique()
    )
    four_week_ends = (
        pd.to_datetime(chatgpt_four_week_usage["period_end"], errors="coerce").dropna().dt.normalize().unique()
    )
    if len(four_week_starts) != 1 or len(four_week_ends) != 1:
        raise ValueError("ChatGPT four-week usage export must contain exactly one valid period_start and period_end.")
    chatgpt_four_week_start = pd.Timestamp(four_week_starts[0])
    chatgpt_four_week_end = pd.Timestamp(four_week_ends[0])
    chatgpt_four_week_days = (chatgpt_four_week_end - chatgpt_four_week_start).days + 1
    if chatgpt_four_week_days < 28:
        raise ValueError(
            "ChatGPT four-week usage export must contain at least 28 inclusive days; "
            f"found {chatgpt_four_week_days} days "
            f"({chatgpt_four_week_start.date()} through {chatgpt_four_week_end.date()})."
        )

    copilot_refresh_dates = (
        pd.to_datetime(copilot_usage["Report Refresh Date"], errors="coerce").dropna().dt.normalize().unique()
    )
    if len(copilot_refresh_dates) != 1:
        raise ValueError("Copilot usage export must contain exactly one valid Report Refresh Date.")
    copilot_end = pd.Timestamp(copilot_refresh_dates[0])
    copilot_periods = pd.to_numeric(copilot_usage["Report Period"], errors="coerce").dropna().unique()
    if len(copilot_periods) != 1 or not float(copilot_periods[0]).is_integer() or copilot_periods[0] <= 0:
        raise ValueError(
            "Copilot usage export must contain exactly one positive whole-number Report Period."
        )
    copilot_days = int(copilot_periods[0])
    copilot_start = copilot_end - pd.Timedelta(days=copilot_days - 1)
    copilot_activity_dates = pd.to_datetime(copilot_usage["Last Activity Date"], errors="coerce")
    future_activity_count = int(copilot_activity_dates.gt(copilot_end).sum())
    if future_activity_count:
        raise ValueError(
            f"Copilot usage export contains {future_activity_count} Last Activity Date value(s) "
            f"after the Copilot period end ({copilot_end.date()})."
        )

    return ReportingWindows(
        period_start, period_end, copilot_start, copilot_end, chatgpt_four_week_start, chatgpt_four_week_end
    )


def reporting_window() -> ReportingWindows:
    if not file_exists(CHATGPT_USAGE_FILE):
        raise FileNotFoundError(f"ChatGPT usage export not found: {CHATGPT_USAGE_FILE.name}")
    if not file_exists(COPILOT_USAGE_FILE):
        raise FileNotFoundError(f"Copilot usage export not found: {COPILOT_USAGE_FILE.name}")
    if not file_exists(CHATGPT_USAGE_FOUR_WEEKS_FILE):
        raise FileNotFoundError(f"ChatGPT four-week usage export not found: {CHATGPT_USAGE_FOUR_WEEKS_FILE.name}")
    chatgpt_usage = pd.read_csv(CHATGPT_USAGE_FILE, usecols=["period_start", "period_end"])
    copilot_usage = pd.read_csv(
        COPILOT_USAGE_FILE, usecols=["Report Refresh Date", "Report Period", "Last Activity Date"]
    )
    chatgpt_four_week_usage = pd.read_csv(CHATGPT_USAGE_FOUR_WEEKS_FILE, usecols=["period_start", "period_end"])
    return validate_reporting_window(chatgpt_usage, copilot_usage, chatgpt_four_week_usage)


def report_date() -> pd.Timestamp:
    """Return the ChatGPT reporting-period end date used as the report date."""
    return reporting_window().chatgpt_end


def report_filename(windows: ReportingWindows) -> str:
    return f"Weekly AI Usage Report - {windows.chatgpt_end.date()}.xlsx"


def ensure_dirs() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)


def file_exists(path: Path) -> bool:
    return path.exists() and path.is_file()


def load_employee_roster() -> pd.DataFrame:
    if file_exists(EMPLOYEE_ROSTER_FILE):
        if EMPLOYEE_ROSTER_FILE.suffix.lower() == ".csv":
            df = pd.read_csv(EMPLOYEE_ROSTER_FILE)
        else:
            df = pd.read_excel(EMPLOYEE_ROSTER_FILE)
        if not df.empty:
            return df

    raise FileNotFoundError(f"Employee roster not found: {EMPLOYEE_ROSTER_FILE.name}")


def load_copilot_users() -> pd.DataFrame:
    if not file_exists(COPILOT_USERS_FILE):
        return pd.DataFrame()
    return pd.read_csv(COPILOT_USERS_FILE)


def load_chatgpt_users() -> pd.DataFrame:
    if not file_exists(CHATGPT_USERS_FILE):
        return pd.DataFrame()
    return pd.read_csv(CHATGPT_USERS_FILE)


def load_copilot_usage() -> pd.DataFrame:
    if not file_exists(COPILOT_USAGE_FILE):
        return pd.DataFrame()
    return pd.read_csv(COPILOT_USAGE_FILE)


def load_chatgpt_usage() -> pd.DataFrame:
    if not file_exists(CHATGPT_USAGE_FILE):
        return pd.DataFrame()
    return pd.read_csv(CHATGPT_USAGE_FILE)


def load_chatgpt_four_week_usage() -> pd.DataFrame:
    if not file_exists(CHATGPT_USAGE_FOUR_WEEKS_FILE):
        return pd.DataFrame()
    return pd.read_csv(CHATGPT_USAGE_FOUR_WEEKS_FILE)


def derive_license_status(row: pd.Series, col_name: str, premium_value: str, no_value: str) -> str:
    val = row.get(col_name)
    if pd.isna(val):
        return no_value
    return premium_value if str(val).strip() == premium_value else no_value


def parse_activity_dates(df: pd.DataFrame, email_column: str, activity_column: str) -> pd.DataFrame:
    if df.empty:
        return df

    out = df.copy()
    out[email_column] = normalize_email(out[email_column])
    out[activity_column] = pd.to_datetime(out[activity_column], errors="coerce")
    out = out.dropna(subset=[email_column]).copy()
    out = out.sort_values([email_column, activity_column], ascending=[True, False], kind="mergesort")
    return out.drop_duplicates(subset=[email_column], keep="first")


def build_employee_level_data() -> pd.DataFrame:
    windows = reporting_window()
    employee_df = load_employee_roster().copy()
    employee_df.columns = [str(c).strip() for c in employee_df.columns]

    # The current roster does not include every optional reporting field.
    for column in BASE_EMPLOYEE_COLUMNS:
        if column not in employee_df.columns:
            employee_df[column] = pd.NA
    employee_df = employee_df.loc[:, BASE_EMPLOYEE_COLUMNS]

    if "Employee Email" in employee_df.columns:
        employee_df["Employee Email"] = normalize_email(employee_df["Employee Email"])

    final_df = employee_df.copy()
    final_df["In Roster"] = True
    # Seed the full outer union, including unlicensed source accounts. Source
    # rows without email cannot be matched safely to a person and are excluded.
    source_people = []
    for source, email_col, name_col in (
        (load_copilot_users(), "User principal name", "Display name"),
        (load_chatgpt_users(), "Email", "Name"),
        (load_copilot_usage(), "User Principal Name", "Display Name"),
        (load_chatgpt_usage(), "email", "name"),
        (load_chatgpt_four_week_usage(), "email", "name"),
    ):
        if source.empty:
            continue
        people = pd.DataFrame({"Employee Email": normalize_email(source[email_col])})
        people["Team Member Name"] = source[name_col] if name_col in source else pd.NA
        source_people.append(people.dropna(subset=["Employee Email"]))
    if source_people:
        people = pd.concat(source_people, ignore_index=True).drop_duplicates("Employee Email")
        final_df = final_df.merge(people, on="Employee Email", how="outer", suffixes=("", "_source"))
        final_df["Team Member Name"] = final_df["Team Member Name"].fillna(final_df.pop("Team Member Name_source"))
        final_df["In Roster"] = final_df["In Roster"].eq(True)
    final_df["Copilot"] = "No Premium License"
    final_df["ChatGPT"] = "No Enterprise License"
    final_df["AI License"] = "No AI License"
    final_df["Last Copilot Activity"] = pd.NaT
    final_df["More than 2 weeks (Copilot)"] = pd.NA
    final_df["Copilot Prompts (Source Period)"] = pd.NA
    final_df["Messages in 2 Weeks"] = pd.NA
    final_df["Active in 2 Weeks (ChatGPT)"] = pd.NA
    final_df["No Activity in 2 Weeks (ChatGPT)"] = pd.NA
    final_df["Both 2 Weeks"] = pd.NA
    final_df["Last Usage (ChatGPT)"] = pd.NaT
    final_df["More than 2 Weeks (ChatGPT)"] = pd.NA
    final_df["Messages in 4 Weeks"] = pd.NA
    final_df["Active in 4 Weeks (ChatGPT)"] = pd.NA
    final_df["No Activity in 4 Weeks (ChatGPT)"] = pd.NA
    final_df["Both 4 Weeks"] = pd.NA
    final_df["Last Usage (ChatGPT, 4 Weeks)"] = pd.NaT
    final_df["More than 4 Weeks (ChatGPT)"] = pd.NA
    final_df["ChatGPT Created Date (UTC)"] = pd.NaT

    # Copilot user export: licensing is identified by the Microsoft 365 Copilot SKU.
    copilot_users = load_copilot_users()
    if not copilot_users.empty:
        copilot_map = copilot_users.loc[
            copilot_users["Licenses"].astype("string").str.contains("Microsoft 365 Copilot", case=False, na=False),
            ["User principal name"],
        ].rename(columns={"User principal name": "Employee Email"})
        copilot_map["Employee Email"] = normalize_email(copilot_map["Employee Email"])
        copilot_map = copilot_map.dropna(subset=["Employee Email"]).drop_duplicates().assign(Copilot="Premium License")
        final_df = final_df.merge(copilot_map, on="Employee Email", how="outer", suffixes=("", "_source"))
        final_df["Copilot"] = final_df.pop("Copilot_source").fillna(final_df["Copilot"])

    # ChatGPT user export contains enabled Enterprise accounts only.
    chatgpt_users = load_chatgpt_users()
    if not chatgpt_users.empty:
        chatgpt_map = chatgpt_users[["Email", "Created Time (UTC)"]].rename(
            columns={"Email": "Employee Email", "Created Time (UTC)": "ChatGPT Created Date (UTC)"}
        )
        chatgpt_map["Employee Email"] = normalize_email(chatgpt_map["Employee Email"])
        chatgpt_map["ChatGPT Created Date (UTC)"] = pd.to_datetime(
            chatgpt_map["ChatGPT Created Date (UTC)"], errors="coerce", utc=True
        ).dt.tz_localize(None)
        chatgpt_map = (
            chatgpt_map.dropna(subset=["Employee Email"])
            .sort_values("ChatGPT Created Date (UTC)", kind="mergesort")
            .drop_duplicates("Employee Email")
            .assign(ChatGPT="Enterprise License")
        )
        final_df = final_df.drop(columns=["ChatGPT Created Date (UTC)"])
        final_df = final_df.merge(chatgpt_map, on="Employee Email", how="outer", suffixes=("", "_source"))
        final_df["ChatGPT"] = final_df.pop("ChatGPT_source").fillna(final_df["ChatGPT"])

    copilot_usage = load_copilot_usage()
    if not copilot_usage.empty:
        activity = copilot_usage[["User Principal Name", "Last Activity Date", "Prompts submitted for All Apps"]].copy()
        activity.columns = ["Employee Email", "Last Copilot Activity", "Copilot Prompts (Source Period)"]
        activity["Employee Email"] = normalize_email(activity["Employee Email"])
        activity["Last Copilot Activity"] = pd.to_datetime(activity["Last Copilot Activity"], errors="coerce")
        activity["Copilot Prompts (Source Period)"] = pd.to_numeric(activity["Copilot Prompts (Source Period)"], errors="coerce").fillna(0)
        activity = activity.sort_values(["Employee Email", "Last Copilot Activity"], ascending=[True, False], kind="mergesort")
        activity = activity.dropna(subset=["Employee Email"]).drop_duplicates(subset=["Employee Email"], keep="first")
        # Replace the placeholder rather than allowing pandas to create _x/_y columns.
        final_df = final_df.drop(columns=["Last Copilot Activity", "Copilot Prompts (Source Period)"])
        final_df = final_df.merge(activity, on="Employee Email", how="outer")

    chatgpt_usage = load_chatgpt_usage()
    if not chatgpt_usage.empty:
        usage = chatgpt_usage[["email", "last_day_active_in_period", "messages", "is_active"]].copy()
        usage.columns = ["Employee Email", "Last Usage (ChatGPT)", "Messages in 2 Weeks", "Active in 2 Weeks (ChatGPT)"]
        usage["Employee Email"] = normalize_email(usage["Employee Email"])
        usage["Last Usage (ChatGPT)"] = pd.to_datetime(usage["Last Usage (ChatGPT)"], errors="coerce")
        usage["Messages in 2 Weeks"] = pd.to_numeric(usage["Messages in 2 Weeks"], errors="coerce").fillna(0)
        usage["Active in 2 Weeks (ChatGPT)"] = pd.to_numeric(usage["Active in 2 Weeks (ChatGPT)"], errors="coerce").fillna(0).eq(1)
        usage = usage.sort_values(["Employee Email", "Last Usage (ChatGPT)"], ascending=[True, False], kind="mergesort")
        usage = usage.dropna(subset=["Employee Email"]).drop_duplicates(subset=["Employee Email"], keep="first")
        # Replace placeholders so the joined usage values retain their expected names.
        final_df = final_df.drop(columns=["Last Usage (ChatGPT)", "Messages in 2 Weeks", "Active in 2 Weeks (ChatGPT)"])
        final_df = final_df.merge(usage, on="Employee Email", how="outer")

    chatgpt_four_week_usage = load_chatgpt_four_week_usage()
    if not chatgpt_four_week_usage.empty:
        usage = chatgpt_four_week_usage[["email", "last_day_active_in_period", "messages", "is_active"]].copy()
        usage.columns = [
            "Employee Email",
            "Last Usage (ChatGPT, 4 Weeks)",
            "Messages in 4 Weeks",
            "Active in 4 Weeks (ChatGPT)",
        ]
        usage["Employee Email"] = normalize_email(usage["Employee Email"])
        usage["Last Usage (ChatGPT, 4 Weeks)"] = pd.to_datetime(usage["Last Usage (ChatGPT, 4 Weeks)"], errors="coerce")
        usage["Messages in 4 Weeks"] = pd.to_numeric(usage["Messages in 4 Weeks"], errors="coerce").fillna(0)
        usage["Active in 4 Weeks (ChatGPT)"] = usage["Last Usage (ChatGPT, 4 Weeks)"].ge(
            windows.chatgpt_four_week_start
        )
        usage = usage.sort_values(["Employee Email", "Last Usage (ChatGPT, 4 Weeks)"], ascending=[True, False], kind="mergesort")
        usage = usage.dropna(subset=["Employee Email"]).drop_duplicates(subset=["Employee Email"], keep="first")
        final_df = final_df.drop(
            columns=["Last Usage (ChatGPT, 4 Weeks)", "Messages in 4 Weeks", "Active in 4 Weeks (ChatGPT)"]
        )
        final_df = final_df.merge(usage, on="Employee Email", how="outer")

    if "Last Copilot Activity" in final_df.columns:
        final_df["Last Copilot Activity"] = pd.to_datetime(final_df["Last Copilot Activity"], errors="coerce")
        final_df["More than 2 weeks (Copilot)"] = (
            final_df["Last Copilot Activity"].notna()
            & final_df["Last Copilot Activity"].lt(windows.copilot_two_week_start)
        )

    if "Messages in 2 Weeks" in final_df.columns:
        final_df["Messages in 2 Weeks"] = pd.to_numeric(final_df["Messages in 2 Weeks"], errors="coerce").fillna(0)

    if "Active in 2 Weeks (ChatGPT)" in final_df.columns:
        final_df["Active in 2 Weeks (ChatGPT)"] = final_df["Active in 2 Weeks (ChatGPT)"].fillna(False).astype(bool)
        final_df["No Activity in 2 Weeks (ChatGPT)"] = ~final_df["Active in 2 Weeks (ChatGPT)"]

    if "Messages in 4 Weeks" in final_df.columns:
        final_df["Messages in 4 Weeks"] = pd.to_numeric(final_df["Messages in 4 Weeks"], errors="coerce").fillna(0)

    if "Active in 4 Weeks (ChatGPT)" in final_df.columns:
        final_df["Active in 4 Weeks (ChatGPT)"] = final_df["Active in 4 Weeks (ChatGPT)"].fillna(False).astype(bool)
        final_df["No Activity in 4 Weeks (ChatGPT)"] = ~final_df["Active in 4 Weeks (ChatGPT)"]

    if "Copilot Prompts (Source Period)" in final_df.columns:
        final_df["Copilot Prompts (Source Period)"] = pd.to_numeric(final_df["Copilot Prompts (Source Period)"], errors="coerce").fillna(0)

    if "Last Usage (ChatGPT)" in final_df.columns:
        final_df["Last Usage (ChatGPT)"] = pd.to_datetime(final_df["Last Usage (ChatGPT)"], errors="coerce")
        final_df["More than 2 Weeks (ChatGPT)"] = (
            final_df["Last Usage (ChatGPT)"].notna()
            & final_df["Last Usage (ChatGPT)"].lt(windows.chatgpt_two_week_start)
        )

    if "Last Usage (ChatGPT, 4 Weeks)" in final_df.columns:
        final_df["Last Usage (ChatGPT, 4 Weeks)"] = pd.to_datetime(
            final_df["Last Usage (ChatGPT, 4 Weeks)"], errors="coerce"
        )
        final_df["More than 4 Weeks (ChatGPT)"] = (
            final_df["Last Usage (ChatGPT, 4 Weeks)"].notna()
            & final_df["Last Usage (ChatGPT, 4 Weeks)"].lt(windows.chatgpt_four_week_start)
        )

    final_df["AI License"] = final_df.apply(
        lambda row: "AI License" if row.get("Copilot") == "Premium License" or row.get("ChatGPT") == "Enterprise License" else "No AI License",
        axis=1,
    )

    final_df["Both 2 Weeks"] = (
        (final_df["Last Copilot Activity"].isna() | final_df["More than 2 weeks (Copilot)"].fillna(False).astype(bool))
        & final_df["No Activity in 2 Weeks (ChatGPT)"].fillna(False).astype(bool)
    )
    final_df["Both 4 Weeks"] = (
        (final_df["Last Copilot Activity"].isna() | final_df["Last Copilot Activity"].lt(windows.copilot_four_week_start))
        & final_df["No Activity in 4 Weeks (ChatGPT)"].fillna(False).astype(bool)
    )

    final_df["Copilot"] = final_df.get("Copilot", "No Premium License").fillna("No Premium License")
    final_df["ChatGPT"] = final_df.get("ChatGPT", "No Enterprise License").fillna("No Enterprise License")

    for col in OUTPUT_EMPLOYEE_COLUMNS:
        if col not in final_df.columns:
            final_df[col] = pd.NA

    final_df["BU"] = final_df["BU"].map(normalize_bu).astype("string")
    created = pd.to_datetime(final_df["ChatGPT Created Date (UTC)"], errors="coerce")
    final_df["new_chatgpt_user"] = classify_new_chatgpt_users(created, windows.max_chatgpt_report_date)
    return final_df[OUTPUT_EMPLOYEE_COLUMNS]


def classify_new_chatgpt_users(created: pd.Series, report_end: pd.Timestamp) -> pd.Series:
    """1 for creation within the trailing 14 inclusive UTC dates; 0 for earlier dates."""
    dates = pd.to_datetime(created, errors="coerce", utc=True).dt.tz_localize(None).dt.normalize()
    return dates.between(report_end - pd.Timedelta(days=13), report_end).astype("Int64").mask(
        dates.isna() | dates.gt(report_end), pd.NA
    )


def established_inactive_chatgpt_users(employee_df: pd.DataFrame) -> pd.DataFrame:
    """Return licensed, known non-new users with no ChatGPT activity in 28 days."""
    selected = employee_df.loc[
        employee_df["ChatGPT"].eq("Enterprise License")
        & employee_df["new_chatgpt_user"].eq(0).fillna(False)
        & employee_df["No Activity in 4 Weeks (ChatGPT)"].eq(True)
    ].copy()
    return selected.sort_values(["Team Member Name", "Employee Email"], kind="mergesort")


def calculate_summary_metrics(df: pd.DataFrame) -> pd.DataFrame:
    employee_count = unique_employee_count(df)
    chatgpt_licensed = int((df["ChatGPT"] == "Enterprise License").sum())
    copilot_licensed = int((df["Copilot"] == "Premium License").sum())
    ai_licensed = int((df["AI License"] == "AI License").sum())
    both_licensed = int(((df["ChatGPT"] == "Enterprise License") & (df["Copilot"] == "Premium License")).sum())
    chatgpt_inactive = int(((df["ChatGPT"] == "Enterprise License") & (df["No Activity in 2 Weeks (ChatGPT)"] == True)).sum())
    copilot_inactive = int(
        ((df["Copilot"] == "Premium License") & (df["Last Copilot Activity"].isna() | (df["More than 2 weeks (Copilot)"] == True))).sum()
    )
    total_inactive = int(((df["AI License"] == "AI License") & (df["Both 2 Weeks"] == True)).sum())

    summary_rows = [
        {"Metric": "Haskell Employees", "ChatGPT": employee_count, "Copilot": None, "Total": None},
        {"Metric": "", "ChatGPT": None, "Copilot": None, "Total": None},
        {"Metric": "", "ChatGPT": "ChatGPT", "Copilot": "Copilot", "Total": "Total"},
        {"Metric": "Licenses", "ChatGPT": CHATGPT_LICENSE_CAPACITY, "Copilot": COPILOT_LICENSE_CAPACITY, "Total": CHATGPT_LICENSE_CAPACITY + COPILOT_LICENSE_CAPACITY},
        {"Metric": "Licensed Users", "ChatGPT": chatgpt_licensed, "Copilot": copilot_licensed, "Total": ai_licensed},
        {"Metric": "License Utilization", "ChatGPT": chatgpt_licensed / CHATGPT_LICENSE_CAPACITY, "Copilot": copilot_licensed / COPILOT_LICENSE_CAPACITY, "Total": "N/A"},
        {"Metric": "% of Employees with License", "ChatGPT": chatgpt_licensed / max(employee_count, 1), "Copilot": copilot_licensed / max(employee_count, 1), "Total": ai_licensed / max(employee_count, 1)},
        {"Metric": "% of AI Users with Both ChatGPT and Copilot Licenses", "ChatGPT": "N/A", "Copilot": "N/A", "Total": both_licensed / max(ai_licensed, 1)},
        {"Metric": "Users with no activity in 2 weeks", "ChatGPT": chatgpt_inactive, "Copilot": copilot_inactive, "Total": total_inactive},
        {"Metric": "% Users with no activity in 2 weeks", "ChatGPT": chatgpt_inactive / max(chatgpt_licensed, 1), "Copilot": copilot_inactive / max(copilot_licensed, 1), "Total": total_inactive / max(ai_licensed, 1)},
    ]

    return pd.DataFrame(summary_rows, columns=["Metric", "ChatGPT", "Copilot", "Total"])


def build_prompt_matrix(df: pd.DataFrame) -> pd.DataFrame:
    """Return a unique-employee 4 x 4 matrix of license and prompt bands."""
    matrix_df = df.copy()
    roster_ids = matrix_df["Team Member Number"].astype("string").str.strip()
    matrix_df["_employee_id"] = ("roster:" + roster_ids.mask(roster_ids.eq(""))).fillna(
        "email:" + matrix_df["Employee Email"].astype("string")
    )
    matrix_df = matrix_df.loc[matrix_df["_employee_id"].notna() & matrix_df["_employee_id"].ne("")].copy()
    matrix_df["_copilot_licensed"] = matrix_df["Copilot"].eq("Premium License")
    matrix_df["_chatgpt_licensed"] = matrix_df["ChatGPT"].eq("Enterprise License")
    matrix_df["_copilot_prompts"] = pd.to_numeric(matrix_df["Copilot Prompts (Source Period)"], errors="coerce").fillna(0)
    matrix_df["_chatgpt_prompts"] = pd.to_numeric(matrix_df["Messages in 2 Weeks"], errors="coerce").fillna(0)

    # Team Member Number is the unique-person key. Max avoids double-counting should
    # a future roster contain repeated records for the same employee.
    people = matrix_df.groupby("_employee_id", as_index=False).agg(
        _copilot_licensed=("_copilot_licensed", "max"),
        _chatgpt_licensed=("_chatgpt_licensed", "max"),
        _copilot_prompts=("_copilot_prompts", "max"),
        _chatgpt_prompts=("_chatgpt_prompts", "max"),
    )

    def prompt_band(licensed: pd.Series, prompts: pd.Series, no_license: str) -> pd.Series:
        bands = pd.Series(no_license, index=people.index, dtype="string")
        bands.loc[licensed & prompts.between(0, 9)] = "Low (0-9)"
        bands.loc[licensed & prompts.between(10, 50)] = "Medium (10-50)"
        bands.loc[licensed & prompts.gt(50)] = "High (51+)"
        return bands

    copilot_bands = ["No Copilot License", "Low (0-9)", "Medium (10-50)", "High (51+)"]
    chatgpt_bands = ["No ChatGPT License", "Low (0-9)", "Medium (10-50)", "High (51+)"]
    people["_copilot_band"] = prompt_band(people["_copilot_licensed"], people["_copilot_prompts"], copilot_bands[0])
    people["_chatgpt_band"] = prompt_band(people["_chatgpt_licensed"], people["_chatgpt_prompts"], chatgpt_bands[0])

    return pd.crosstab(people["_copilot_band"], people["_chatgpt_band"]).reindex(
        index=copilot_bands, columns=chatgpt_bands, fill_value=0
    )


def apply_sheet_styles(ws, row_count: int, col_count: int) -> None:
    header_fill = PatternFill(fill_type="solid", fgColor="D9E2F3")
    header_font = Font(bold=True)
    for row in ws.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=2, max_row=row_count):
        for cell in row:
            cell.alignment = Alignment(vertical="center")

    for col in range(1, col_count + 1):
        max_len = 0
        for row in ws.iter_rows(min_row=1, max_row=row_count, min_col=col, max_col=col):
            for cell in row:
                if cell.value is None:
                    val_len = 0
                else:
                    val_len = len(str(cell.value))
                if val_len > max_len:
                    max_len = val_len
        ws.column_dimensions[chr(64 + col)].width = max(12, max_len + 2)

    ws.freeze_panes = "A2"


def write_employee_sheet(df: pd.DataFrame, output_wb: Workbook) -> None:
    ws = output_wb.active
    ws.title = "Employee Level Data"
    ws.append(list(df.columns))
    for _, row in df.iterrows():
        ws.append([None if pd.isna(value) else value for value in row])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="808080")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in (
        "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB"
    ):
        ws.column_dimensions[col].width = 15

    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["G"].width = 24
    ws.column_dimensions["J"].width = 24
    ws.column_dimensions["K"].width = 22
    ws.column_dimensions["M"].width = 20
    ws.column_dimensions["N"].width = 20
    ws.column_dimensions["Q"].width = 16
    ws.column_dimensions["R"].width = 18
    ws.column_dimensions["S"].width = 18
    ws.column_dimensions["T"].width = 18
    ws.column_dimensions["U"].width = 18
    for field, width in (("ChatGPT Created Date (UTC)", 26), ("new_chatgpt_user", 24)):
        column = OUTPUT_EMPLOYEE_COLUMNS.index(field) + 1
        ws.column_dimensions[get_column_letter(column)].width = width
        ws.cell(1, column).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if field == "ChatGPT Created Date (UTC)":
            for row in range(2, len(df) + 2):
                ws.cell(row, column).number_format = "mm/dd/yyyy hh:mm"
    ws.row_dimensions[1].height = 32


def write_summary_sheet(df: pd.DataFrame, output_wb: Workbook, windows: ReportingWindows) -> None:
    ws = output_wb.create_sheet("Summary")

    employee_count = unique_employee_count(df)
    ws["A1"] = "Haskell Employees"
    ws["B1"] = employee_count
    ws["B3"] = "ChatGPT"
    ws["D3"] = "Copilot"
    ws["F3"] = "Total"
    copilot_four_week_inactive = int(
        (
            (df["Copilot"] == "Premium License")
            & (df["Last Copilot Activity"].isna() | df["Last Copilot Activity"].lt(windows.copilot_four_week_start))
        ).sum()
    )
    chatgpt_four_week_inactive = int(
        ((df["ChatGPT"] == "Enterprise License") & (df["No Activity in 4 Weeks (ChatGPT)"] == True)).sum()
    )
    total_four_week_inactive = int(((df["AI License"] == "AI License") & (df["Both 4 Weeks"] == True)).sum())
    # Minimal fallback table for environments without the reference workbook.
    summary_rows = [
        ("Licenses", CHATGPT_LICENSE_CAPACITY, COPILOT_LICENSE_CAPACITY, CHATGPT_LICENSE_CAPACITY + COPILOT_LICENSE_CAPACITY),
        ("Licensed Users", int((df["ChatGPT"] == "Enterprise License").sum()), int((df["Copilot"] == "Premium License").sum()), int((df["AI License"] == "AI License").sum())),
        ("License Utilization", (df["ChatGPT"] == "Enterprise License").sum() / CHATGPT_LICENSE_CAPACITY, (df["Copilot"] == "Premium License").sum() / COPILOT_LICENSE_CAPACITY, "N/A"),
        ("% of Employees with License", (df["ChatGPT"] == "Enterprise License").sum() / max(employee_count, 1), (df["Copilot"] == "Premium License").sum() / max(employee_count, 1), (df["AI License"] == "AI License").sum() / max(employee_count, 1)),
        ("% of AI Users with Both ChatGPT and Copilot Licenses", "N/A", "N/A", int(((df["ChatGPT"] == "Enterprise License") & (df["Copilot"] == "Premium License")).sum()) / max(int((df["AI License"] == "AI License").sum()), 1)),
        ("Users with no activity in 2 weeks", int(((df["ChatGPT"] == "Enterprise License") & (df["No Activity in 2 Weeks (ChatGPT)"] == True)).sum()), int(((df["Copilot"] == "Premium License") & (df["Last Copilot Activity"].isna() | (df["More than 2 weeks (Copilot)"] == True))).sum()), int(((df["AI License"] == "AI License") & (df["Both 2 Weeks"] == True)).sum())),
        ("% Users with no activity in 2 weeks", int(((df["ChatGPT"] == "Enterprise License") & (df["No Activity in 2 Weeks (ChatGPT)"] == True)).sum()) / max(int((df["ChatGPT"] == "Enterprise License").sum()), 1), int(((df["Copilot"] == "Premium License") & (df["Last Copilot Activity"].isna() | (df["More than 2 weeks (Copilot)"] == True))).sum()) / max(int((df["Copilot"] == "Premium License").sum()), 1), int(((df["AI License"] == "AI License") & (df["Both 2 Weeks"] == True)).sum()) / max(int((df["AI License"] == "AI License").sum()), 1)),
        ("Users with no activity in 4 weeks", chatgpt_four_week_inactive, copilot_four_week_inactive, total_four_week_inactive),
        ("% Users with no activity in 4 weeks", chatgpt_four_week_inactive / max(int((df["ChatGPT"] == "Enterprise License").sum()), 1), copilot_four_week_inactive / max(int((df["Copilot"] == "Premium License").sum()), 1), total_four_week_inactive / max(int((df["AI License"] == "AI License").sum()), 1)),
    ]
    for offset, (metric, chatgpt_val, copilot_val, total_val) in enumerate(summary_rows, start=4):
        row = offset
        ws.cell(row=row, column=1, value=metric)
        ws.cell(row=row, column=2, value=chatgpt_val)
        ws.cell(row=row, column=4, value=copilot_val)
        ws.cell(row=row, column=6, value=total_val)

    ws["A15"] = "ChatGPT data ended"
    ws["B15"] = windows.chatgpt_end.date()
    ws["A16"] = "Copilot data ended"
    ws["B16"] = windows.copilot_end.date()
    ws["A17"] = "2-week metric windows"
    ws["B17"] = (
        f"ChatGPT: {windows.chatgpt_two_week_start:%m/%d/%y}-{windows.chatgpt_end:%m/%d/%y} (14 days); "
        f"Copilot: {windows.copilot_two_week_start:%m/%d/%y}-{windows.copilot_end:%m/%d/%y} (14 days)"
    )
    ws["A18"] = "4-week metric windows"
    ws["B18"] = (
        f"ChatGPT: {windows.chatgpt_four_week_start:%m/%d/%y}-{windows.chatgpt_four_week_end:%m/%d/%y} (28 days); "
        f"Copilot: {windows.copilot_four_week_start:%m/%d/%y}-{windows.copilot_end:%m/%d/%y} (28 days)"
    )
    ws.merge_cells("B17:F17")
    ws.merge_cells("B18:F18")
    for cell in ("A1", "B3", "D3", "F3", "A15", "A16", "A17", "A18"):
        ws[cell].font = Font(bold=True)
    for row in (6, 7, 8, 10, 12):
        for col in (2, 4, 6):
            if isinstance(ws.cell(row=row, column=col).value, (int, float)):
                ws.cell(row=row, column=col).number_format = "0.0%"
    ws["B15"].number_format = ws["B16"].number_format = "mm/dd/yy"

    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 4
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 4
    ws.column_dimensions["F"].width = 20
    add_summary_filters(df, ws, output_wb, windows)


def add_summary_filters(df: pd.DataFrame, ws, output_wb: Workbook, windows: ReportingWindows) -> None:
    """Filter Summary calculations using five independent, exact-match selectors."""
    last_row = len(df) + 1

    def data_range(field: str) -> str:
        letter = get_column_letter(OUTPUT_EMPLOYEE_COLUMNS.index(field) + 1)
        return f"'Employee Level Data'!${letter}$2:${letter}${last_row}"

    masks = []
    ws["H1"] = "Summary filters"
    ws["H1"].font = Font(bold=True)
    ws["H8"] = "All includes accounts outside the roster."
    ws["H9"] = "Use (Blank) for missing roster attributes."
    ws["H10"] = "Filters apply to Summary only."
    for index, field in enumerate(FILTER_FIELDS):
        row = index + 2
        ws.cell(row, 8, field).font = Font(bold=True)
        ws.cell(row, 9, "All")
        ws.cell(row, 9).fill = PatternFill("solid", fgColor="FFF2CC")
        options = ["All", "(Blank)"] + sorted(
            {str(value) for value in df[field].dropna().unique() if str(value) != ""}
        )
        option_col = get_column_letter(27 + index)
        for option_row, value in enumerate(options, start=1):
            ws[f"{option_col}{option_row}"] = value
        ws.column_dimensions[option_col].hidden = True
        name = f"SummaryFilter{index}"
        output_wb.defined_names.add(DefinedName(name, attr_text=f"'Summary'!${option_col}$1:${option_col}${len(options)}"))
        validation = DataValidation(type="list", formula1=f"={name}")
        validation.errorTitle = "Choose a listed value"
        validation.error = "Select All, (Blank), or a value from the dropdown."
        validation.showErrorMessage = True
        ws.add_data_validation(validation)
        validation.add(ws.cell(row, 9))
        selector = f"$I${row}"
        values = data_range(field)
        masks.append(f'--((({selector}="All")+({values}=IF({selector}="(Blank)","",{selector})))>0)')

    def count(*conditions: str) -> str:
        return "SUMPRODUCT(" + ",".join(masks + [f"--({condition})" for condition in conditions]) + ")"

    cg = f'{data_range("ChatGPT")}="Enterprise License"'
    cp = f'{data_range("Copilot")}="Premium License"'
    ai = f'{data_range("AI License")}="AI License"'
    roster = f'{data_range("In Roster")}=TRUE'
    # Roster headcount and employee penetration exclude source-only accounts.
    ws["B1"] = "=" + count(roster)
    for col, condition in (("B", cg), ("D", cp), ("F", ai)):
        ws[f"{col}5"] = "=" + count(condition)
        ws[f"{col}7"] = "=" + count(condition, roster) + "/MAX($B$1,1)"
    ws["B6"] = "=B5/B4"
    ws["D6"] = "=D5/D4"
    ws["F8"] = "=" + count(cg, cp) + "/MAX(F5,1)"

    for weeks, count_row, rate_row in ((2, 9, 10), (4, 11, 12)):
        ws[f"B{count_row}"] = "=" + count(cg, f'{data_range(f"No Activity in {weeks} Weeks (ChatGPT)")}=TRUE')
        start = windows.copilot_end - pd.Timedelta(days=weeks * 7 - 1)
        dates = data_range("Last Copilot Activity")
        cutoff = f"DATE({start.year},{start.month},{start.day})"
        ws[f"D{count_row}"] = "=" + count(cp, f'(({dates}="")+({dates}<{cutoff}))>0')
        ws[f"F{count_row}"] = "=" + count(ai, f'{data_range(f"Both {weeks} Weeks")}=TRUE')
        for col in ("B", "D", "F"):
            ws[f"{col}{rate_row}"] = f"={col}{count_row}/MAX({col}5,1)"
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 38
    ws.column_dimensions["I"].width = 35


def write_prompt_matrix_sheet(df: pd.DataFrame, output_wb: Workbook) -> None:
    ws = output_wb.create_sheet("Prompt Matrix")
    matrix = build_prompt_matrix(df)

    ws["A1"] = "AI License and Prompt Matrix"
    ws["A2"] = "Low: 0-9 prompts | Medium: 10-50 prompts | High: 51+ prompts"
    ws["A4"] = "Copilot \\ ChatGPT"

    for column_index, category in enumerate(matrix.columns, start=2):
        ws.cell(row=4, column=column_index, value=category)
    for row_index, (category, values) in enumerate(matrix.iterrows(), start=5):
        ws.cell(row=row_index, column=1, value=category)
        for column_index, value in enumerate(values, start=2):
            ws.cell(row=row_index, column=column_index, value=int(value))

    ws["A10"] = "People counted"
    ws["B10"] = int(matrix.to_numpy().sum())
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"].font = Font(italic=True)
    for cell in ws[4]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="808080")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=5, max_row=8, min_col=1, max_col=1):
        row[0].font = Font(bold=True)
        row[0].fill = PatternFill("solid", fgColor="D9E2F3")
    for row in ws.iter_rows(min_row=5, max_row=8, min_col=2, max_col=5):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.number_format = "#,##0"
    ws.column_dimensions["A"].width = 24
    for column in ("B", "C", "D", "E"):
        ws.column_dimensions[column].width = 22
    ws.sheet_view.showGridLines = False


def write_inactive_chatgpt_sheet(employee_df: pd.DataFrame, output_wb: Workbook, windows: ReportingWindows) -> None:
    ws = output_wb.create_sheet("ChatGPT Inactive 4 Weeks")
    selected = established_inactive_chatgpt_users(employee_df)
    columns = [
        "Team Member Name", "Employee Email", "Job", "BU", "Org 1 Report Category", "Location",
        "Alternate Title", "In Roster", "ChatGPT", "ChatGPT Created Date (UTC)", "new_chatgpt_user",
        "Last Usage (ChatGPT, 4 Weeks)", "No Activity in 4 Weeks (ChatGPT)",
    ]
    ws.append(columns)
    for values in selected[columns].itertuples(index=False, name=None):
        ws.append([None if pd.isna(value) else value for value in values])
    ws.auto_filter.ref = f"A1:M{len(selected) + 1}"
    ws.freeze_panes = "C2"
    ws.row_dimensions[1].height = 42
    for column, field in enumerate(columns, start=1):
        letter = get_column_letter(column)
        ws.column_dimensions[letter].width = 32 if field == "Employee Email" else 24
        header = ws.cell(1, column)
        header.font = Font(bold=True, color="FFFFFF")
        header.fill = PatternFill("solid", fgColor="808080")
        header.alignment = Alignment(wrap_text=True, vertical="center")
        if field in ("ChatGPT Created Date (UTC)", "Last Usage (ChatGPT, 4 Weeks)"):
            for row in range(2, len(selected) + 2):
                ws.cell(row, column).number_format = "mm/dd/yyyy"
    ws["O1"] = "Selection"
    ws["O1"].font = Font(bold=True)
    ws["O2"] = f"New-user reference date: {windows.max_chatgpt_report_date.date()}"
    ws["O3"] = f"New-user window: {(windows.max_chatgpt_report_date - pd.Timedelta(days=13)).date()} through {windows.max_chatgpt_report_date.date()} (14 days)"
    ws["O4"] = f"No-activity window: {windows.chatgpt_four_week_start.date()} through {windows.chatgpt_four_week_end.date()} (28 days)"
    ws["O5"] = "Licensed ChatGPT users with new_chatgpt_user = 0 and no activity in 4 weeks."
    ws["O6"] = "Unknown or post-report creation dates are excluded from this list."
    ws["O7"] = "Account creation date is a proxy for license tenure."
    ws["O8"] = f"Users listed: {len(selected)}"
    ws.column_dimensions["O"].width = 95
    ws.sheet_view.showGridLines = False


def write_output_workbook(employee_df: pd.DataFrame, windows: ReportingWindows, output_file: Path) -> None:
    output_wb = Workbook()
    write_employee_sheet(employee_df, output_wb)
    write_summary_sheet(employee_df, output_wb, windows)
    write_prompt_matrix_sheet(employee_df, output_wb)
    write_inactive_chatgpt_sheet(employee_df, output_wb, windows)

    output_wb.save(output_file)
    print(f"Workbook written to: {output_file}")


def main() -> None:
    ensure_dirs()
    try:
        windows = reporting_window()
        employee_df = build_employee_level_data()
    except (FileNotFoundError, ValueError) as exc:
        raise SystemExit(str(exc)) from exc

    local_output_file = OUTPUT_DIR / report_filename(windows)
    write_output_workbook(employee_df, windows, local_output_file)
    PUBLISHED_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    published_output_file = PUBLISHED_OUTPUT_DIR / report_filename(windows)
    write_output_workbook(employee_df, windows, published_output_file)
    print(f"Established ChatGPT users inactive for 4 weeks: {len(established_inactive_chatgpt_users(employee_df))}")

    print(f"Report date: {windows.chatgpt_end.date()}")
    print(
        f"ChatGPT reporting window: {windows.chatgpt_start.date()} through {windows.chatgpt_end.date()} "
        f"({windows.chatgpt_days} days)"
    )
    print(
        f"Copilot reporting window: {windows.copilot_start.date()} through {windows.copilot_end.date()} "
        f"({windows.copilot_days} days)"
    )


if __name__ == "__main__":
    main()
