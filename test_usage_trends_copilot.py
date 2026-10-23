"""Check monthly distinct-user counts, license snapshots and calendar coverage."""

import unittest

import pandas as pd

from usage_trends import monthly_from_daily


class CopilotMonthlyTests(unittest.TestCase):
    def setUp(self):
        rows = []
        for day in pd.date_range('2026-01-01', '2026-02-03'):
            for person in ['a', 'b', 'c', 'd']:
                enabled = int(person != 'c' or day < pd.Timestamp('2026-01-31'))
                actions = int(person == 'a')
                if person == 'b' and day == pd.Timestamp('2026-01-05'):
                    actions = 4
                if person == 'c' and day == pd.Timestamp('2026-01-15'):
                    actions = 100
                rows.append({'PersonId': person, 'MetricDate': day,
                             'Copilot Enabled User': enabled, 'Total Copilot actions taken': actions})
        self.daily = pd.DataFrame(rows)

    def test_distinct_active_users_and_period_end_licensed_population(self):
        january = monthly_from_daily(self.daily, '2026-02-04').iloc[0]
        self.assertEqual(january['Enabled Users'], 3)
        self.assertEqual(january['Active Users'], 2)  # Repeated daily activity counts once.
        self.assertEqual(january['Total Copilot Actions'], 35)  # Excludes removed c's actions.
        self.assertAlmostEqual(january['% Active'], 2 / 3)
        self.assertAlmostEqual(january['Actions per Enabled User'], 35 / 3)
        self.assertEqual(january['Actions per Active User'], 17.5)

    def test_mtd_included_by_default_with_exclusive_as_of(self):
        february = monthly_from_daily(self.daily, '2026-02-03').iloc[-1]
        self.assertFalse(february['Complete month'])
        self.assertEqual(february['Days covered'], 2)
        self.assertEqual(february['Data through'], pd.Timestamp('2026-02-02'))
        self.assertEqual(february['Total Copilot Actions'], 2)
        self.assertEqual(february['Active Users'], 1)

    def test_completed_months_only(self):
        result = monthly_from_daily(self.daily, '2026-02-04', include_partial=False)
        self.assertEqual(len(result), 1)
        self.assertTrue(result.iloc[0]['Complete month'])

    def test_leading_partial_month_is_excluded(self):
        result = monthly_from_daily(self.daily.loc[self.daily.MetricDate.ge('2026-01-10')], '2026-02-04')
        self.assertEqual(len(result), 1)
        self.assertEqual(result.iloc[0]['Month ending'], pd.Timestamp('2026-02-28'))

    def test_missing_calendar_day_is_rejected(self):
        daily = self.daily.loc[self.daily.MetricDate.ne('2026-01-12')]
        with self.assertRaisesRegex(ValueError, 'missing daily Copilot'):
            monthly_from_daily(daily, '2026-02-04')

    def test_zero_enabled_population_keeps_ratios_unavailable(self):
        self.daily.loc[self.daily.MetricDate.eq('2026-01-31'), 'Copilot Enabled User'] = 0
        january = monthly_from_daily(self.daily, '2026-02-04').iloc[0]
        self.assertEqual(january['Enabled Users'], 0)
        self.assertEqual(january['Active Users'], 0)
        self.assertEqual(january['Total Copilot Actions'], 0)
        self.assertTrue(pd.isna(january['% Active']))
        self.assertTrue(pd.isna(january['Actions per Active User']))


if __name__ == '__main__':
    unittest.main()
