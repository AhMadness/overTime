import unittest

from overtime_core import calculate_rates, calculate_summary


class RateCalculationTests(unittest.TestCase):
    def test_calculates_daily_and_hourly_rates(self):
        daily, hourly = calculate_rates(3100, 31)
        self.assertEqual(daily, 100)
        self.assertEqual(hourly, 12.5)

    def test_calculates_overtime_summary_with_multiplier(self):
        entries = [{"hours": 2}, {"hours": 6.5}]
        hours, days, amount = calculate_summary(entries, 3100, 31, 1.5)
        self.assertEqual(hours, 8.5)
        self.assertEqual(days, 1.0625)
        self.assertEqual(amount, 159.375)

    def test_handles_empty_entries(self):
        self.assertEqual(calculate_summary([], 3000, 30, 2), (0, 0, 0))

    def test_rejects_invalid_month_length(self):
        with self.assertRaises(ValueError):
            calculate_rates(3000, 0)


if __name__ == "__main__":
    unittest.main()
