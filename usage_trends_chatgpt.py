"""Generate six ChatGPT weekly or monthly charts and enabled-user tables.

Run: python usage_trends_chatgpt.py
Monthly (including month-to-date): python usage_trends_chatgpt.py --frequency monthly
Completed months only: python usage_trends_chatgpt.py --frequency monthly --completed-months-only
Only the creation-date table: python usage_trends_chatgpt.py --enabled-only

Chart metrics and Notes come from input/chatgpt_weekly_metrics_with_admin_usage.csv.
Actions are a proxy: ChatGPT Messages + Tool Messages Received. Missing ratios
and active-user counts remain missing; chart populations follow the source.
Monthly charts combine daily Chat messages with chatgpt_monthly_active_users.csv.
Monthly active users are the distinct admin counts supplied in that file.
Messages per active user = monthly Chat messages / distinct monthly active users.
The active-user export has no product breakdown; its supplied scope is retained.
Monthly boundaries are UTC to match the daily exports; weekly boundaries remain
Eastern. Enabled users are current-export creation-date proxies in both cases.
Dependencies: pandas, matplotlib (same chart renderer as usage_trends.py).
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from usage_trends import make_chart

ROOT = Path(__file__).resolve().parent
SOURCE_DEFAULT = ROOT / 'input' / 'chatgpt_weekly_metrics_with_admin_usage.csv'
OUTPUT_DEFAULT = ROOT / 'chatgpt_weekly_charts'
METRICS = ['Enabled Users', 'Active Users (Admin)', '% Active', 'Total Actions (Proxy)',
           'Actions per Enabled User', 'Actions per Active User']


def creation_timestamps(source: Path) -> pd.Series:
    """Validate one current user per email and return UTC creation instants."""
    users = pd.read_csv(source, usecols=['Email', 'Created Time (UTC)'])
    emails = users['Email'].str.strip().str.lower()
    if emails.isna().any() or emails.eq('').any() or emails.duplicated().any():
        raise ValueError('ChatGPT users must have nonempty, unique email addresses.')
    created = pd.to_datetime(users['Created Time (UTC)'], utc=True, errors='raise', format='ISO8601')
    if created.isna().any() or created.empty:
        raise ValueError('ChatGPT users must have valid creation timestamps and at least one row.')
    return created


def completed_sunday(as_of: str | None = None) -> pd.Timestamp:
    """Latest Sunday completed before the given Eastern calendar date."""
    report_date = (pd.Timestamp(as_of) if as_of is not None
                   else pd.Timestamp.now(tz='America/New_York').tz_localize(None)).normalize()
    return report_date - pd.Timedelta(days=report_date.weekday() + 1)


def chatgpt_weekly_enabled_users(source: Path, as_of: str | None = None) -> pd.DataFrame:
    """Count current-export users created by each completed Sunday in Eastern time.

    Creation is a proxy for enablement: the export cannot reconstruct removed
    users, disablements or the actual license assignment date. The as-of date
    is exclusive, so a Sunday becomes complete on the following Monday.
    """
    created = creation_timestamps(source)

    # Convert UTC instants before assigning local dates, including DST changes.
    local_dates = created.dt.tz_convert('America/New_York').dt.tz_localize(None).dt.normalize()
    first_date = local_dates.min()
    first_sunday = first_date + pd.Timedelta(days=6 - first_date.weekday())
    weeks = pd.date_range(first_sunday, completed_sunday(as_of), freq='W-SUN')
    # Exclusive Monday cutoff includes all of Sunday, even subsecond timestamps.
    counts = local_dates.sort_values().searchsorted(weeks + pd.Timedelta(days=1), side='left')
    return pd.DataFrame({'Week ending': weeks, 'Enabled Users': counts})


def chatgpt_monthly_enabled_users(source: Path, as_of: str | None = None) -> pd.DataFrame:
    """Cumulative current-export users at completed UTC month ends.

    UTC matches the analytics daily buckets. Weekly output keeps Eastern time.
    Historical removals and enable/disable dates remain unavailable.
    """
    created = creation_timestamps(source)
    report_date = (pd.Timestamp(as_of) if as_of is not None
                   else pd.Timestamp.now(tz='America/New_York').tz_localize(None)).normalize()
    first_month = created.min().tz_localize(None).to_period('M').start_time
    last_end = report_date.to_period('M').start_time - pd.Timedelta(days=1)
    months = pd.date_range(first_month, last_end, freq='ME')
    cutoffs = (months + pd.Timedelta(days=1)).tz_localize('UTC')
    counts = created.sort_values().searchsorted(cutoffs, side='left')
    return pd.DataFrame({'Month ending': months, 'Enabled Users': counts})


def load_daily_analytics(source: Path) -> pd.DataFrame:
    """Read daily UTC product counts, rejecting missing or duplicate buckets."""
    daily = pd.read_csv(source, usecols=['Start Time', 'End Time', 'Chat', 'Codex', 'Work'])
    for column in ['Start Time', 'End Time']:
        daily[column] = pd.to_datetime(daily[column], utc=True, errors='raise', format='ISO8601')
    starts, ends = daily['Start Time'], daily['End Time']
    if (daily.empty or starts.isna().any() or ends.isna().any() or starts.duplicated().any()
            or starts.ne(starts.dt.normalize()).any()
            or (ends - starts).ne(pd.Timedelta(days=1)).any()):
        raise ValueError(f'{source.name}: expected unique full daily UTC intervals.')
    for column in ['Chat', 'Codex', 'Work']:
        values = pd.to_numeric(daily[column], errors='raise')
        if values.isna().any() or values.lt(0).any() or values.mod(1).ne(0).any():
            raise ValueError(f'{source.name}: {column} must contain nonnegative integer counts.')
        daily[column] = values.astype('int64')
    return daily.sort_values('Start Time').reset_index(drop=True)


def find_analytics_export(kind: str) -> Path:
    matches = sorted((ROOT / 'input').glob(f'analytics-overview-{kind}_*.csv'))
    if len(matches) != 1:
        raise ValueError(f'Expected one {kind} export in input, found {len(matches)}; specify its path explicitly.')
    return matches[0]


def load_monthly_active_users(source: Path) -> pd.DataFrame:
    """Read distinct monthly counts and their inclusive UTC coverage dates."""
    active = pd.read_csv(source, usecols=['Month', 'Period Start (UTC)', 'Period End (UTC)',
                                        'Active Users (Admin, Distinct)'])
    active['Month'] = pd.to_datetime(active['Month'], format='%Y-%m', errors='raise').dt.to_period('M')
    if active.empty or active['Month'].isna().any() or active['Month'].duplicated().any():
        raise ValueError('Monthly active-user records must have unique, nonempty months.')
    for column in ['Period Start (UTC)', 'Period End (UTC)']:
        active[column] = pd.to_datetime(active[column], utc=True, format='mixed', errors='raise')
        if active[column].isna().any() or active[column].ne(active[column].dt.normalize()).any():
            raise ValueError('Monthly active-user coverage must use nonempty UTC calendar dates.')
    for _, row in active.iterrows():
        start = row['Month'].start_time.tz_localize('UTC')
        next_month = (row['Month'] + 1).start_time.tz_localize('UTC')
        if row['Period Start (UTC)'] != start or not start <= row['Period End (UTC)'] < next_month:
            raise ValueError('Monthly active-user coverage must start on day 1 and end within its labeled month.')
    counts = pd.to_numeric(active['Active Users (Admin, Distinct)'], errors='raise')
    if counts.isna().any() or counts.lt(0).any() or counts.mod(1).ne(0).any():
        raise ValueError('Distinct monthly active users must be nonnegative integer counts.')
    active['Active Users'] = counts.astype('int64')
    missing = set(pd.period_range(active['Month'].min(), active['Month'].max(), freq='M')) - set(active['Month'])
    if missing:
        raise ValueError(f'Missing distinct active-user counts for entire months: {sorted(missing)}')
    return active.sort_values('Month').reset_index(drop=True)


def monthly_from_exports(messages_source: Path, active_source: Path, users_source: Path,
                         as_of: str | None = None, include_partial: bool = False) -> pd.DataFrame:
    """Join daily Chat messages to supplied distinct active counts for exact UTC periods.

    Each included period needs all daily message buckets. An MTD active count
    cannot be truncated to an earlier as-of date or extended to later messages.
    """
    messages = load_daily_analytics(messages_source)
    active = load_monthly_active_users(active_source)
    report_date = (pd.Timestamp(as_of) if as_of is not None
                   else pd.Timestamp.now(tz='America/New_York').tz_localize(None)).normalize()
    created = creation_timestamps(users_source).sort_values()
    rows = []
    for _, record in active.iterrows():
        month = record['Month']
        start = record['Period Start (UTC)']
        cutoff = record['Period End (UTC)'] + pd.Timedelta(days=1)
        complete = cutoff == (month + 1).start_time.tz_localize('UTC')
        if cutoff > report_date.tz_localize('UTC') or (not complete and not include_partial):
            continue
        group = messages.loc[messages['Start Time'].ge(start) & messages['End Time'].le(cutoff)]
        expected_days = (cutoff - start).days
        if len(group) != expected_days:
            raise ValueError(f'{month}: missing daily message records for the active-user coverage period.')
        enabled = int(created.searchsorted(cutoff, side='left'))
        active_users = int(record['Active Users'])
        total_messages = int(group['Chat'].sum())
        notes = ['Active users are distinct monthly admin counts as supplied (no product breakdown).',
                 'Messages use the Chat column only.',
                 'Enabled users use current-export creation dates and UTC cutoffs.']
        if not complete:
            notes.append(f'Month to date through {record["Period End (UTC)"]:%Y-%m-%d} UTC.')
        if active_users > enabled:
            notes.append('Distinct active users exceed the enabled-user proxy; source populations differ.')
        rows.append({
            'Month ending': month.end_time.normalize(),
            'Data through (UTC)': record['Period End (UTC)'].tz_localize(None),
            'Complete month': complete,
            'Days covered': expected_days,
            'Enabled Users': enabled,
            'Active Users': active_users,
            '% Active': active_users / enabled if enabled else float('nan'),
            'ChatGPT Messages': total_messages,
            'Messages per Enabled User': total_messages / enabled if enabled else float('nan'),
            'Messages per Active User': total_messages / active_users if active_users else float('nan'),
            'Notes': ' '.join(notes),
        })
    if not rows:
        raise ValueError('No covered calendar months available; partial months require --include-partial-month.')
    return pd.DataFrame(rows)


def export_monthly(args: argparse.Namespace) -> None:
    """Write monthly tables and six charts using distinct active users."""
    enabled = chatgpt_monthly_enabled_users(args.users, args.as_of)
    monthly = None
    if not args.enabled_only:
        monthly = monthly_from_exports(args.messages or find_analytics_export('messages'),
                                     args.active_users or ROOT / 'input' / 'chatgpt_monthly_active_users.csv',
                                     args.users, args.as_of, args.include_partial_month)
        # The MTD snapshot uses the same cutoff as the supplied active-user count.
        enabled = monthly[['Month ending', 'Data through (UTC)', 'Complete month', 'Enabled Users']]
    args.output_dir.mkdir(parents=True, exist_ok=True)
    enabled.to_csv(args.output_dir / 'chatgpt_monthly_enabled_users.csv', index=False, date_format='%Y-%m-%d')
    if monthly is not None:
        monthly.to_csv(args.output_dir / 'chatgpt_monthly_metrics.csv', index=False, date_format='%Y-%m-%d')
        charts = [
            ('Enabled Users', 'Enabled ChatGPT Users by Month', 'enabled_users.png', 'integer'),
            ('Active Users', 'Distinct Monthly Active ChatGPT Users', 'monthly_active_users.png', 'integer'),
            ('% Active', '% Active ChatGPT Users by Month', 'monthly_active_rate.png', 'percent'),
            ('ChatGPT Messages', 'Total ChatGPT Messages by Month', 'total_chatgpt_messages.png', 'integer'),
            ('Messages per Enabled User', 'Monthly ChatGPT Messages per Enabled User', 'messages_per_enabled_user.png', 'decimal'),
            ('Messages per Active User', 'Monthly ChatGPT Messages per Active User', 'messages_per_active_user.png', 'decimal'),
        ]
        for column, title, filename, style in charts:
            if not monthly['Complete month'].all():
                title += ' (Includes MTD)'
            make_chart(monthly, column, title, '', filename, args.output_dir, style,
                       decimal_tick_step=None, date_column='Month ending', period_label='Month (UTC)')
        # Remove the superseded chart from the earlier average-daily analysis.
        obsolete = args.output_dir / 'messages_per_active_user_day.png'
        if obsolete.exists():
            obsolete.unlink()
        print(monthly.drop(columns='Notes').to_string(index=False))
    else:
        print(enabled.to_string(index=False))
    print('UTC months; messages use Chat only. Active users = supplied distinct monthly admin counts.')
    print('The active-user file has no product breakdown; its supplied scope is retained.')
    print('Active rate = distinct monthly active users / end-of-period enabled users.')
    print('Messages per active user = total Chat messages / distinct monthly active users.')
    print('Enabled counts use current-export creation dates; historical removals/disablements unavailable.')
    print(f'Files: {args.output_dir.resolve()}')


def load_weekly_metrics(source: Path, as_of: str | None = None) -> pd.DataFrame:
    """Read the supplied weekly metrics without filling unavailable values."""
    weekly = pd.read_csv(source)
    missing = set(['Week ending', *METRICS]) - set(weekly.columns)
    if missing:
        raise ValueError(f'Missing ChatGPT columns: {sorted(missing)}')
    weekly['Week ending'] = pd.to_datetime(weekly['Week ending'], format='%Y-%m-%d', errors='raise')
    dates = weekly['Week ending']
    if dates.isna().any() or dates.duplicated().any() or dates.dt.weekday.ne(6).any():
        raise ValueError('ChatGPT week-ending dates must be unique Sundays.')
    for column in METRICS:
        weekly[column] = pd.to_numeric(weekly[column], errors='raise')
        if weekly[column].lt(0).any():
            raise ValueError(f'Negative ChatGPT metric: {column}')
    weekly = weekly.loc[dates.le(completed_sunday(as_of))].sort_values('Week ending').reset_index(drop=True)
    if weekly.empty:
        raise ValueError('No completed Sunday-ending weeks in the ChatGPT metrics file.')
    return weekly


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--source', type=Path, default=SOURCE_DEFAULT, help='Weekly ChatGPT metrics CSV')
    parser.add_argument('--users', '--chatgpt-users', type=Path, default=ROOT / 'users_chatgpt.csv',
                        help='User export for the creation-date enabled-user table')
    parser.add_argument('--frequency', choices=['weekly', 'monthly'], default='weekly')
    parser.add_argument('--output-dir', type=Path, help='Defaults to chatgpt_weekly_charts or chatgpt_monthly_charts')
    parser.add_argument('--messages', type=Path, help='Monthly mode: daily messages export (auto-detected in input)')
    parser.add_argument('--active-users', type=Path, help='Monthly mode: distinct monthly active-users CSV (default input/chatgpt_monthly_active_users.csv)')
    coverage = parser.add_mutually_exclusive_group()
    coverage.add_argument('--include-partial-month', action='store_true', default=None,
                          help='Monthly mode: include labeled month-to-date data (default for monthly charts)')
    coverage.add_argument('--completed-months-only', action='store_false', dest='include_partial_month',
                          help='Monthly mode: exclude the current incomplete month')
    parser.add_argument('--as-of', '--chatgpt-as-of', help='YYYY-MM-DD; exclusive reporting cutoff date')
    parser.add_argument('--enabled-only', '--chatgpt-only', action='store_true',
                        help='Only generate the creation-date enabled-user table')
    args = parser.parse_args()
    if args.include_partial_month is None:
        args.include_partial_month = args.frequency == 'monthly' and not args.enabled_only
    if args.include_partial_month and (args.frequency != 'monthly' or args.enabled_only):
        parser.error('--include-partial-month requires --frequency monthly and the monthly active-users export')
    if args.output_dir is None:
        args.output_dir = ROOT / f'chatgpt_{args.frequency}_charts'
    if args.frequency == 'monthly':
        export_monthly(args)
        return

    enabled = chatgpt_weekly_enabled_users(args.users, args.as_of)
    weekly = None if args.enabled_only else load_weekly_metrics(args.source, args.as_of)
    args.output_dir.mkdir(parents=True, exist_ok=True)
    enabled.to_csv(args.output_dir / 'chatgpt_weekly_enabled_users.csv', index=False, date_format='%Y-%m-%d')
    print('Creation-date enabled-user proxy (America/New_York); historical removals/disablements unavailable.')
    print(enabled.to_string(index=False))
    if weekly is not None:
        weekly.to_csv(args.output_dir / 'chatgpt_weekly_metrics.csv', index=False, date_format='%Y-%m-%d')
        charts = [
            ('Enabled Users', 'Enabled ChatGPT Users', 'enabled_users.png', 'integer'),
            ('Active Users (Admin)', 'Active ChatGPT Users', 'weekly_active_users.png', 'integer'),
            ('% Active', '% Active ChatGPT Users', 'weekly_active_rate.png', 'percent'),
            ('Total Actions (Proxy)', 'Total ChatGPT Actions (Proxy)', 'total_chatgpt_actions.png', 'integer'),
            ('Actions per Enabled User', 'ChatGPT Actions per Enabled User (Proxy)', 'actions_per_enabled_user.png', 'decimal'),
            ('Actions per Active User', 'ChatGPT Actions per Active User (Proxy)', 'actions_per_active_user.png', 'decimal'),
        ]
        for column, title, filename, style in charts:
            make_chart(weekly, column, title, '', filename, args.output_dir, style,
                       decimal_tick_step=None)
        print(f'Chart source: {args.source}')
        print('Actions proxy = ChatGPT Messages + Tool Messages Received. Missing metrics remain gaps.')
        if 'Notes' in weekly.columns:
            notes = weekly.loc[weekly['Notes'].fillna('').ne(''), ['Week ending', 'Notes']]
            if not notes.empty:
                print(notes.to_string(index=False))
        print(f'Chart weeks: {len(weekly)}, {weekly["Week ending"].min():%Y-%m-%d} through {weekly["Week ending"].max():%Y-%m-%d}')
    print(f'Files: {args.output_dir.resolve()}')


if __name__ == '__main__':
    main()
