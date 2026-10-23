"""Verify monthly time boundaries, aggregation and source completeness."""

from pathlib import Path
import unittest
from unittest.mock import patch

import pandas as pd

from usage_trends_chatgpt import (
    chatgpt_monthly_enabled_users,
    load_daily_analytics,
    load_monthly_active_users,
    monthly_from_exports,
)


class MonthlyMetricsTests(unittest.TestCase):
    def setUp(self):
        self.users = Path('users.csv')
        users = pd.DataFrame({
            'Email': ['a@example.com', 'b@example.com', 'c@example.com'],
            'Created Time (UTC)': ['2026-01-01T00:00:00Z', '2026-01-31T23:59:59.999999Z',
                                   '2026-02-01T00:00:00Z'],
        })
        self.messages = Path('messages.csv')
        self.active = Path('active.csv')
        starts = pd.date_range('2026-01-01', '2026-02-03', tz='UTC')
        self.daily = pd.DataFrame({
            'Start Time': starts, 'End Time': starts + pd.Timedelta(days=1),
            'Chat': 1, 'Codex': 100, 'Work': 200,
        })
        monthly_active = pd.DataFrame({
            'Month': ['2026-01', '2026-02'],
            'Period Start (UTC)': ['1/1/2026', '2/1/2026'],
            'Period End (UTC)': ['1/31/2026', '2/3/2026'],
            'Active Users (Admin, Distinct)': [2, 3],
        })
        self.sources = {self.users: users, self.active: monthly_active,
                        self.messages: self.daily.assign(Chat=10)}
        reader = patch('usage_trends_chatgpt.pd.read_csv', side_effect=self.read_source)
        reader.start()
        self.addCleanup(reader.stop)

    def read_source(self, path, **kwargs):
        return self.sources[path][kwargs['usecols']].copy()

    def test_utc_month_end_and_incomplete_month_exclusion(self):
        enabled = chatgpt_monthly_enabled_users(self.users, '2026-02-04')
        self.assertEqual(enabled['Enabled Users'].tolist(), [2])
        result = monthly_from_exports(self.messages, self.active, self.users, '2026-02-04')
        self.assertEqual(len(result), 1)
        january = result.iloc[0]
        self.assertEqual(january['ChatGPT Messages'], 310)
        self.assertEqual(january['Active Users'], 2)
        self.assertEqual(january['Messages per Active User'], 155)
        self.assertEqual(january['Messages per Enabled User'], 155)
        self.assertEqual(january['% Active'], 1)
        self.assertNotIn('Average Daily Active Users', result.columns)

    def test_partial_month_and_as_of_exclusive_cutoff(self):
        self.sources[self.active].loc[1, 'Period End (UTC)'] = '2/2/2026'
        result = monthly_from_exports(self.messages, self.active, self.users,
                                    '2026-02-03', include_partial=True)
        february = result.iloc[-1]
        self.assertFalse(february['Complete month'])
        self.assertEqual(february['Days covered'], 2)
        self.assertEqual(february['Data through (UTC)'], pd.Timestamp('2026-02-02'))
        self.assertEqual(february['Enabled Users'], 3)
        self.assertEqual(february['ChatGPT Messages'], 20)
        self.assertEqual(february['Active Users'], 3)
        self.assertAlmostEqual(february['Messages per Active User'], 20 / 3)

    def test_active_snapshot_cannot_be_truncated_to_earlier_as_of(self):
        result = monthly_from_exports(self.messages, self.active, self.users,
                                      '2026-02-03', include_partial=True)
        self.assertEqual(len(result), 1)

    def test_messages_after_active_snapshot_are_excluded(self):
        self.sources[self.active].loc[1, 'Period End (UTC)'] = '2/1/2026'
        result = monthly_from_exports(self.messages, self.active, self.users,
                                      '2026-02-04', include_partial=True)
        self.assertEqual(result.iloc[-1]['ChatGPT Messages'], 10)
        self.assertEqual(result.iloc[-1]['Days covered'], 1)

    def test_missing_calendar_day_is_rejected(self):
        self.sources[self.messages] = self.daily.drop(index=12).assign(Chat=10)
        with self.assertRaisesRegex(ValueError, 'missing daily message'):
            monthly_from_exports(self.messages, self.active, self.users, '2026-02-04')

    def test_duplicate_monthly_active_counts_are_rejected(self):
        active = self.sources[self.active]
        self.sources[self.active] = pd.concat([active, active.iloc[[0]]])
        with self.assertRaisesRegex(ValueError, 'unique, nonempty months'):
            load_monthly_active_users(self.active)

    def test_missing_entire_month_is_rejected(self):
        self.sources[self.active].loc[1, ['Month', 'Period Start (UTC)', 'Period End (UTC)']] = [
            '2026-03', '3/1/2026', '3/31/2026']
        with self.assertRaisesRegex(ValueError, 'entire months'):
            load_monthly_active_users(self.active)

    def test_monthly_coverage_must_start_on_first_day(self):
        self.sources[self.active].loc[0, 'Period Start (UTC)'] = '1/2/2026'
        with self.assertRaisesRegex(ValueError, 'start on day 1'):
            load_monthly_active_users(self.active)

    def test_duplicate_daily_buckets_are_rejected(self):
        self.sources[self.messages] = pd.concat([self.daily, self.daily.iloc[[0]]])
        with self.assertRaisesRegex(ValueError, 'unique full daily UTC intervals'):
            load_daily_analytics(self.messages)

    def test_zero_active_users_preserves_unavailable_ratio(self):
        self.sources[self.active]['Active Users (Admin, Distinct)'] = 0
        result = monthly_from_exports(self.messages, self.active, self.users, '2026-02-04')
        self.assertEqual(result.iloc[0]['% Active'], 0)
        self.assertTrue(pd.isna(result.iloc[0]['Messages per Active User']))


if __name__ == '__main__':
    unittest.main()
