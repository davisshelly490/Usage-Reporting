"""Reproduce six Haskell Copilot weekly or monthly charts from a Viva Insights export.

Example:
    python usage_trends.py \
        --source 'Person_Davis.Shelly_Oct21_2026_1417Hours (2).zip' \
        --output-dir 'copilot_weekly_charts_oct2026'

Monthly (includes month-to-date):
    python usage_trends.py --frequency monthly
Completed months only:
    python usage_trends.py --frequency monthly --completed-months-only

Dependencies:
    pip install pandas matplotlib

The October 2026 Person_... CSV is a DAILY export, not the same weekly grain as
Copilot_dashboard_metric_Sep28_2026_1548Hours.csv. This script aggregates daily
per-person records into SUNDAY-ENDING full weeks. It does not treat each daily
row as a separate weekly user. The end-of-week enabled population is the
number of licensed users as of Sunday; active users are those same licensed
people with at least one Copilot action during the prior seven days.
Monthly mode uses the same population rule at month end (or the latest covered
day for MTD), with each licensed person counted once across the entire month.
Actions include those users' actions throughout the reporting period. Monthly
dates follow MetricDate's calendar. Leading partial months are excluded; the
last month is included and labeled MTD when the export stops before month end.

The historical results may differ from earlier dashboard exports because the
report population/granularity and historical metric values differ.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import zipfile

import matplotlib
matplotlib.use('Agg')
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.ticker import FuncFormatter, MaxNLocator, MultipleLocator
import pandas as pd

SOURCE_DEFAULT = Path.home() / 'Downloads' / 'Person_Davis.Shelly_Oct21_2026_1417Hours (2).zip'
OUTPUT_DEFAULT = Path(__file__).with_name('copilot_weekly_charts_oct2026')
NEEDED = ['PersonId', 'MetricDate', 'Copilot Enabled User', 'Total Copilot actions taken']
WORKSHOP_WEEKS = pd.to_datetime(['2026-06-07','2026-06-14','2026-06-21'])
BLUE = '#1D517C'
ORANGE = '#EF7A2C'
#SHADE = '#FAEADD'


def load_source(path: Path) -> pd.DataFrame:
    """Read only the source columns, from a plain CSV or the CSV inside a ZIP."""
    if path.suffix.lower() == '.zip':
        with zipfile.ZipFile(path) as archive:
            csv_names = [n for n in archive.namelist() if n.lower().endswith('.csv') and not n.startswith('__MACOSX/')]
            if len(csv_names) != 1:
                raise ValueError(f'Expected exactly one CSV in ZIP, found: {csv_names}')
            with archive.open(csv_names[0]) as source:
                df = pd.read_csv(source, usecols=NEEDED, encoding='utf-8-sig', low_memory=False)
    else:
        df = pd.read_csv(path, usecols=NEEDED, encoding='utf-8-sig', low_memory=False)
    df['MetricDate'] = pd.to_datetime(df['MetricDate'], errors='raise').dt.normalize()
    df['Copilot Enabled User'] = pd.to_numeric(df['Copilot Enabled User'], errors='coerce').fillna(0)
    df['Total Copilot actions taken'] = pd.to_numeric(df['Total Copilot actions taken'], errors='coerce').fillna(0)
    if df['PersonId'].isna().any() or df['MetricDate'].isna().any():
        raise ValueError('PersonId and MetricDate are required for unique counts.')
    if df.duplicated(['PersonId', 'MetricDate']).any():
        raise ValueError('Duplicate PersonId/MetricDate rows detected. Check the source export before charting.')
    return df


def weekly_from_daily(df: pd.DataFrame) -> pd.DataFrame:
    """Convert daily records into complete Sunday-ending weekly population metrics."""
    # Sunday has weekday=6; the following expression maps every date to its Sunday.
    df = df.copy()
    df['Week ending'] = df['MetricDate'] + pd.to_timedelta(6 - df['MetricDate'].dt.weekday, unit='D')
    # Drop partial boundary weeks (rather than showing deceptive drop-offs).
    covered_days = df.groupby('Week ending')['MetricDate'].nunique()
    valid_weeks = covered_days.loc[covered_days.eq(7)].index
    df = df[df['Week ending'].isin(valid_weeks)].copy()
    if df.empty:
        raise ValueError('No complete seven-day Sunday-ending weeks found in the export.')

    # Each user's actions throughout the full week (sum of seven daily values).
    by_user = (
        df.groupby(['Week ending', 'PersonId'], as_index=False)
        ['Total Copilot actions taken'].sum()
        .rename(columns={'Total Copilot actions taken': 'Weekly actions'})
    )
    # A snapshot of who had an M365 Copilot license at the END of that week.
    sunday = df.loc[df['MetricDate'].dt.weekday.eq(6),
                    ['Week ending', 'PersonId', 'Copilot Enabled User']]
    by_user = by_user.merge(sunday, how='left', on=['Week ending', 'PersonId'], validate='one_to_one')
    by_user['Enabled'] = by_user['Copilot Enabled User'].fillna(0).gt(0)
    by_user['Active'] = by_user['Enabled'] & by_user['Weekly actions'].gt(0)
    by_user['Licensed actions'] = by_user['Weekly actions'].where(by_user['Enabled'], 0)

    weekly = (
        by_user.groupby('Week ending', as_index=False)
        .agg(**{
            'Enabled Users': ('Enabled', 'sum'),
            'Active Users': ('Active', 'sum'),
            'Total Copilot Actions': ('Licensed actions', 'sum'),
        })
        .sort_values('Week ending')
    )
    weekly['% Active'] = weekly['Active Users'] / weekly['Enabled Users'].replace(0, float('nan'))
    weekly['Actions per Enabled User'] = weekly['Total Copilot Actions'] / weekly['Enabled Users'].replace(0, float('nan'))
    weekly['Actions per Active User'] = weekly['Total Copilot Actions'] / weekly['Active Users'].replace(0, float('nan'))
    for name in ['Enabled Users', 'Active Users', 'Total Copilot Actions']:
        weekly[name] = weekly[name].astype(int)
    return weekly


def monthly_from_daily(df: pd.DataFrame, as_of: str | None = None,
                       include_partial: bool = True) -> pd.DataFrame:
    """Aggregate distinct users and actions against end-of-period license snapshots."""
    report_date = (pd.Timestamp(as_of) if as_of is not None
                   else pd.Timestamp.now(tz='America/New_York').tz_localize(None)).normalize()
    daily = df.loc[df['MetricDate'].lt(report_date)].copy()
    if daily.empty:
        raise ValueError('No Copilot records before the reporting cutoff.')
    daily['Month'] = daily['MetricDate'].dt.to_period('M')
    missing = set(pd.period_range(daily['Month'].min(), daily['Month'].max(), freq='M')) - set(daily['Month'])
    if missing:
        raise ValueError(f'Missing Copilot daily records for entire months: {sorted(missing)}')
    rows = []
    for month, group in daily.groupby('Month', sort=True):
        last_date = group['MetricDate'].max()
        complete = last_date == month.end_time.normalize()
        if group['MetricDate'].min() != month.start_time or (not complete and not include_partial):
            continue
        days = (last_date - month.start_time).days + 1
        if group['MetricDate'].nunique() != days:
            raise ValueError(f'{month}: missing daily Copilot records; cannot calculate monthly metrics.')
        # One snapshot per person, using the final covered day of the period.
        licensed_ids = group.loc[(group['MetricDate'] == last_date)
                                 & group['Copilot Enabled User'].gt(0), 'PersonId']
        actions = group.loc[group['PersonId'].isin(licensed_ids)].groupby('PersonId')[
            'Total Copilot actions taken'].sum()
        enabled = int(licensed_ids.nunique())
        active = int(actions.gt(0).sum())
        total = int(actions.sum())
        rows.append({
            'Month ending': month.end_time.normalize(),
            'Data through': last_date,
            'Complete month': complete,
            'Days covered': days,
            'Enabled Users': enabled,
            'Active Users': active,
            '% Active': active / enabled if enabled else float('nan'),
            'Total Copilot Actions': total,
            'Actions per Enabled User': total / enabled if enabled else float('nan'),
            'Actions per Active User': total / active if active else float('nan'),
        })
    if not rows:
        raise ValueError('No covered calendar months available in the Copilot export.')
    return pd.DataFrame(rows)


def make_chart(weekly: pd.DataFrame, col: str, title: str, ylabel: str,
               filename: str, outdir: Path, value_style: str,
               decimal_tick_step: float | None = 5,
               date_column: str = 'Week ending',
               period_label: str = 'Week ending (Sunday)') -> None:
    x = weekly[date_column]
    y = weekly[col]
    last_valid = y.last_valid_index()
    last_value = y.loc[last_valid] if last_valid is not None else float('nan')
    fig, ax = plt.subplots(figsize=(14.5, 6.75), dpi=180)
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')

    # Use the same blue/orange workshop motif as the original six charts.
    # ax.axvspan(WORKSHOP_WEEKS[0] - pd.Timedelta(days=3.5),
    #            WORKSHOP_WEEKS[-1] + pd.Timedelta(days=3.5),
    #            color=SHADE, alpha=0.6, lw=0, zorder=0)
    ax.plot(x, y, color=BLUE, marker='o', markersize=4.3,
            lw=2.2, zorder=2)
    # highlighted = weekly[weekly['Week ending'].isin(WORKSHOP_WEEKS)]
    # if not highlighted.empty:
    #     ax.plot(highlighted['Week ending'], highlighted[col], color=ORANGE,
    #             marker='o', lw=3.1, markersize=6.4, zorder=4)

    ax.set_title(title, fontsize=17, fontweight='bold', pad=16)
    ax.set_ylabel(ylabel, fontsize=12)
    ax.set_xlabel(period_label, fontsize=11)
    ax.grid(axis='y', color='#DDDDDD', lw=0.85)
    ax.set_axisbelow(True)
    ax.spines[['top','right']].set_visible(False)
    ax.spines[['left','bottom']].set_color('#AEB8C0')

    # Weekly points, approximately one date label every three weeks.
    monthly = date_column == 'Month ending'
    tick_ix = list(range(0, len(x), 1 if monthly else 3))
    if tick_ix[-1] != len(x) - 1:
        tick_ix.append(len(x) - 1)
    ax.set_xticks(x.iloc[tick_ix])
    labels = ([f'{d:%b %Y}' for d in x.iloc[tick_ix]] if monthly
              else [f'{d:%b} {d.day}' for d in x.iloc[tick_ix]])
    if monthly and 'Complete month' in weekly:
        labels = [label if weekly['Complete month'].iloc[i] else f'{label} (MTD)'
                  for label, i in zip(labels, tick_ix)]
    ax.set_xticklabels(labels, rotation=38, ha='right', fontsize=9)
    ax.set_xlim(x.iloc[0] - pd.Timedelta(days=8), x.iloc[-1] + pd.Timedelta(days=17))

    if value_style == 'percent':
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, pos: f'{v:.1%}' if monthly else f'{v:.0%}'))
        last_label = f'{float(last_value):.1%}'
    elif value_style == 'integer':
        ax.yaxis.set_major_locator(MaxNLocator(nbins=7, integer=True))
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, pos: f'{v:,.0f}'))
        last_label = f'{float(last_value):,.0f}'
    else:
        if col in ('Actions per Enabled User', 'Actions per Active User'):
            ax.yaxis.set_major_locator(MultipleLocator(decimal_tick_step) if decimal_tick_step
                                      else MaxNLocator(nbins=7))
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, pos: f'{v:,.1f}' if monthly else f'{v:,.0f}'))
        last_label = f'{float(last_value):,.1f}'

    if last_valid is not None:
        ax.annotate(last_label, xy=(x.loc[last_valid], last_value),
                    xytext=(8, 0), textcoords='offset points',
                    ha='left', va='center', fontsize=11, fontweight='bold', color='#1A2A35')

    #line = Line2D([0],[0],color=ORANGE, marker='o', lw=3, label='eGroup workshop weeks')
    #ax.legend(handles=[line], loc='upper left', frameon=False, fontsize=10)
    # fig.text(0.01, 0.012,
    #          'Source: Copilot Dashboard person-level DAILY export; full 7-day weeks, ending Sundays. June workshop weeks highlighted.',
    #          color='#52616C', fontsize=8.3)
    fig.subplots_adjust(left=0.08, right=0.93, bottom=0.20, top=0.88)
    fig.savefig(outdir / filename, dpi=180, bbox_inches='tight', facecolor='white')
    plt.close(fig)

def main() -> None:
    parser=argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--source',type=Path,default=SOURCE_DEFAULT,help='CSV file or ZIP containing a single CSV')
    parser.add_argument('--frequency', choices=['weekly', 'monthly'], default='weekly')
    parser.add_argument('--output-dir',type=Path,help='Destination for six PNGs and metrics CSV (separate weekly/monthly defaults)')
    parser.add_argument('--as-of', help='YYYY-MM-DD; exclusive reporting cutoff date')
    coverage = parser.add_mutually_exclusive_group()
    coverage.add_argument('--include-partial-month', action='store_true', default=None,
                          help='Monthly mode: include the last month-to-date period (default)')
    coverage.add_argument('--completed-months-only', action='store_false', dest='include_partial_month',
                          help='Monthly mode: exclude the last incomplete month')
    args=parser.parse_args()
    if args.include_partial_month is None:
        args.include_partial_month = args.frequency == 'monthly'
    if args.include_partial_month and args.frequency != 'monthly':
        parser.error('--include-partial-month requires --frequency monthly')
    if args.output_dir is None:
        args.output_dir = (OUTPUT_DEFAULT if args.frequency == 'weekly'
                           else Path(__file__).with_name('copilot_monthly_charts'))
    args.output_dir.mkdir(parents=True,exist_ok=True)

    data = load_source(args.source)
    if args.as_of is not None:
        data = data.loc[data['MetricDate'].lt(pd.Timestamp(args.as_of))]
    monthly = args.frequency == 'monthly'
    metrics = (monthly_from_daily(data, args.as_of, args.include_partial_month) if monthly
               else weekly_from_daily(data))
    metrics.to_csv(args.output_dir / f'copilot_{args.frequency}_metrics.csv', index=False,
                   date_format='%Y-%m-%d', float_format='%.6f')

    charts = [
        ('Enabled Users','Enabled Copilot Premium Users','','enabled_users.png','integer'),
        ('Active Users','Active Copilot Premium Users','','weekly_active_users.png','integer'),
        ('% Active','% Active Copilot Premium Users','','weekly_active_rate.png','percent'),
        ('Total Copilot Actions','Total Copilot Premium Actions','','total_copilot_actions.png','integer'),
        ('Actions per Enabled User','Copilot Actions per Enabled Premium User','','actions_per_enabled_user.png','decimal'),
        ('Actions per Active User','Copilot Actions per Active Premium User','','actions_per_active_user.png','decimal'),
    ]
    for col,title,ylabel,filename,fmt in charts:
        if monthly:
            title += ' by Month'
            if not metrics['Complete month'].all():
                title += ' (Includes MTD)'
            filename = filename.replace('weekly_', 'monthly_')
            make_chart(metrics, col, title, ylabel, filename, args.output_dir, fmt,
                       decimal_tick_step=None, date_column='Month ending', period_label='Month')
        else:
            make_chart(metrics,col,title,ylabel,filename,args.output_dir,fmt)
    print(f'Export: {args.source.name}')
    print(f'Raw date range: {data.MetricDate.min():%Y-%m-%d} through {data.MetricDate.max():%Y-%m-%d}')
    date_column = 'Month ending' if monthly else 'Week ending'
    print(f'{args.frequency.capitalize()} periods: {len(metrics)}, {metrics[date_column].min():%Y-%m-%d} through {metrics[date_column].max():%Y-%m-%d}')
    print(metrics.tail(5).to_string(index=False))
    if monthly:
        print('Distinct active users and actions include only users licensed at the end of each reporting period.')
        print('Monthly dates follow the source MetricDate calendar; MTD ends on the last covered source day.')
    print(f'Files: {args.output_dir.resolve()}')

if __name__ == '__main__':
    main()
