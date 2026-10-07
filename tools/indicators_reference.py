"""Python reference of the financial indicators of modIndicators (the same formulas, used by the tests).
Each returns None when the figure has no meaning (nothing to divide by); the dashboard then shows "-"."""


def gross_margin(sales, cost):
    """(net sales - cost of sales) / net sales."""
    return None if sales <= 0 else (sales - cost) / sales


def stock_turnover(year_cost, stock_start, stock_end):
    """Cost of sales of the last 365 days / the average stock (start and end of those days)."""
    average = (stock_start + stock_end) / 2
    return None if average <= 0 or year_cost <= 0 else year_cost / average


def stock_days(turnover):
    """How many days the stock lasts at that pace."""
    return None if not turnover else 365 / turnover


def collection_days(receivables, credit_sales, days=90):
    """Customer debts / credit sales per day over the last `days` days."""
    return None if credit_sales <= 0 else max(receivables, 0) / (credit_sales / days)


def current_ratio(current_assets, current_liabilities):
    """Current assets / current liabilities."""
    return None if current_liabilities <= 0 else current_assets / current_liabilities
