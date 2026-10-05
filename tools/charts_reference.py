"""Reference for the pure helpers of modCharts (checked through LibreOffice)."""
import math
from datetime import date


def nice_max(value: float) -> float:
    """Smallest 1, 2, 2.5 or 5 x 10^n that is >= value."""
    if value <= 0:
        return 1.0
    p = 10.0 ** math.floor(math.log10(value))
    for step in (1, 2, 2.5, 5, 10):
        if step * p >= value * 0.999999:
            return step * p
    return 10 * p


def chart_months(end: date) -> str:
    out = []
    for i in range(11, -1, -1):
        y, m = end.year, end.month - i
        while m < 1:
            m += 12
            y -= 1
        out.append(str(y * 100 + m))
    return ",".join(out)
