def calculate_rates(salary, days_in_month, hours_per_day=8):
    """Calculate daily and hourly rates for a monthly salary."""
    if days_in_month <= 0 or hours_per_day <= 0:
        raise ValueError("Days and hours must be positive")

    daily_rate = float(salary) / days_in_month
    return daily_rate, daily_rate / hours_per_day


def calculate_summary(entries, salary, days_in_month, multiplier=1):
    """Calculate total overtime hours, days, and payment amount."""
    total_hours = sum(float(entry['hours']) for entry in entries)
    _, hourly_rate = calculate_rates(salary, days_in_month)
    return total_hours, total_hours / 8, total_hours * hourly_rate * multiplier
