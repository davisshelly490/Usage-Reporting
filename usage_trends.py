"""Reproduce the six Haskell Copilot weekly usage charts from a Viva Insights export.

Example:
    python usage_trends.py \
        --source 'Person_Davis.Shelly_Oct21_2026_1417Hours (2).zip' \
        --output-dir 'copilot_weekly_charts_oct2026'

Dependencies:
    pip install pandas matplotlib

The October 2026 Person_... CSV is a DAILY export, not the same weekly grain as
Copilot_dashboard_metric_Sep28_2026_1548Hours.csv. This script aggregates daily
per-person records into SUNDAY-ENDING full weeks. It does not treat each daily
row as a separate weekly user. The end-of-week enabled population is the
number of licensed users as of Sunday; active users are those same licensed
people with at least one Copilot action during the prior seven days.

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


def make_chart(weekly: pd.DataFrame, col: str, title: str, ylabel: str,
               filename: str, outdir: Path, value_style: str) -> None:
    x = weekly['Week ending']
    y = weekly[col]
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
    ax.set_xlabel('Week ending (Sunday)', fontsize=11)
    ax.grid(axis='y', color='#DDDDDD', lw=0.85)
    ax.set_axisbelow(True)
    ax.spines[['top','right']].set_visible(False)
    ax.spines[['left','bottom']].set_color('#AEB8C0')

    # Weekly points, approximately one date label every three weeks.
    tick_ix = list(range(0, len(x), 3))
    if tick_ix[-1] != len(x) - 1:
        tick_ix.append(len(x) - 1)
    ax.set_xticks(x.iloc[tick_ix])
    ax.set_xticklabels([f'{d:%b} {d.day}' for d in x.iloc[tick_ix]], rotation=38, ha='right', fontsize=9)
    ax.set_xlim(x.iloc[0] - pd.Timedelta(days=8), x.iloc[-1] + pd.Timedelta(days=17))

    if value_style == 'percent':
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, pos: f'{v:.0%}'))
        last_label = f'{float(y.iloc[-1]):.1%}'
    elif value_style == 'integer':
        ax.yaxis.set_major_locator(MaxNLocator(nbins=7, integer=True))
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, pos: f'{v:,.0f}'))
        last_label = f'{float(y.iloc[-1]):,.0f}'
    else:
        if col in ('Actions per Enabled User', 'Actions per Active User'):
            ax.yaxis.set_major_locator(MultipleLocator(5))
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, pos: f'{v:,.0f}'))
        last_label = f'{float(y.iloc[-1]):,.1f}'

    ax.annotate(last_label, xy=(x.iloc[-1], y.iloc[-1]),
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
    parser.add_argument('--output-dir',type=Path,default=OUTPUT_DEFAULT,help='Destination for six PNGs and weekly CSV')
    args=parser.parse_args()
    args.output_dir.mkdir(parents=True,exist_ok=True)

    data = load_source(args.source)
    weekly = weekly_from_daily(data)
    weekly.to_csv(args.output_dir / 'copilot_weekly_metrics.csv', index=False, date_format='%Y-%m-%d', float_format='%.6f')

    charts = [
        ('Enabled Users','Enabled Copilot Premium Users','','enabled_users.png','integer'),
        ('Active Users','Active Copilot Premium Users','','weekly_active_users.png','integer'),
        ('% Active','% Active Copilot Premium Users','','weekly_active_rate.png','percent'),
        ('Total Copilot Actions','Total Copilot Premium Actions','','total_copilot_actions.png','integer'),
        ('Actions per Enabled User','Copilot Actions per Enabled Premium User','','actions_per_enabled_user.png','decimal'),
        ('Actions per Active User','Copilot Actions per Active Premium User','','actions_per_active_user.png','decimal'),
    ]
    for col,title,ylabel,filename,fmt in charts:
        make_chart(weekly,col,title,ylabel,filename,args.output_dir,fmt)
    print(f'Export: {args.source.name}')
    print(f'Raw date range: {data.MetricDate.min():%Y-%m-%d} through {data.MetricDate.max():%Y-%m-%d}')
    print(f'Complete weeks: {len(weekly)}, {weekly["Week ending"].min():%Y-%m-%d} through {weekly["Week ending"].max():%Y-%m-%d}')
    print(weekly.tail(5).to_string(index=False))
    print(f'Files: {args.output_dir.resolve()}')

if __name__ == '__main__':
    main()
