"""Python mirror of the rules of modProductUnits (docs/72, multiple units per product).

A line in a bigger unit (a carton = factor pieces) keeps its quantity and prices in that unit; the stock moves
quantity x factor pieces at the cost of one piece, and the cost of one unit of the line is that cost x factor.
"""
from decimal import ROUND_HALF_UP, Decimal
from typing import Optional, Tuple


def money(x, digits=2) -> Decimal:
    return Decimal(str(x)).quantize(Decimal(1).scaleb(-digits), rounding=ROUND_HALF_UP)


def unit_factor(base_unit: int, unit: Optional[int], units: dict) -> Decimal:
    """1 for the smallest unit (or none), the factor of a unit of the product, 0 when it is not one of them."""
    if unit is None or unit == 0 or unit == base_unit:
        return Decimal(1)
    return Decimal(str(units.get(unit, 0)))


def unit_sale_price(piece_price, factor, own_price=0) -> Decimal:
    """UnitSalePrice: the piece price, else the unit's own price, else the piece price x factor."""
    if Decimal(str(factor)) == 1:
        return money(piece_price)
    if Decimal(str(own_price)) > 0:
        return money(own_price)
    return money(Decimal(str(piece_price)) * Decimal(str(factor)))


def unit_purchase_price(piece_cost, factor, own_price=0) -> Decimal:
    if Decimal(str(own_price)) > 0:
        return money(own_price, 4)
    return money(Decimal(str(piece_cost)) * Decimal(str(factor)), 4)


def sale_stock_move(qty, factor, average_cost) -> Tuple[Decimal, Decimal, Decimal]:
    """(pieces out, cost of one piece, UnitCost of the line) of a sale line."""
    f = Decimal(str(factor))
    return -Decimal(str(qty)) * f, money(average_cost, 4), money(Decimal(str(average_cost)) * f, 4)


def purchase_stock_move(qty, factor, net_amount) -> Tuple[Decimal, Decimal]:
    """(pieces in, cost of one piece) of a purchase line (PurchaseUnitCost over the pieces)."""
    pieces = Decimal(str(qty)) * Decimal(str(factor))
    return pieces, money(Decimal(str(net_amount)) / pieces, 4)
