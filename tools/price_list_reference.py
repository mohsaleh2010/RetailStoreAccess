"""Python mirror of the rules of modPriceLists (docs/73, price lists).

A customer has one price list. The price of a new sale line for him:
  1. the item of the line's unit in his list, valid that day (the valid item that starts last wins);
  2. a bigger unit without an item: the item of the smallest unit x the factor;
  3. else the usual price (rep price, the unit's own price, the piece price x factor) less the list discount.
"""
from datetime import date
from decimal import ROUND_HALF_UP, Decimal
from typing import Iterable, NamedTuple, Optional


class Item(NamedTuple):
    item_id: int
    list_id: int
    product: int
    unit: Optional[int]            # None = the smallest unit
    price: Decimal
    valid_from: Optional[date] = None
    valid_to: Optional[date] = None


def money(x) -> Decimal:
    return Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def item_price(items: Iterable[Item], list_id: int, product: int, unit: Optional[int], day: date) -> Optional[Decimal]:
    """PriceListItemPrice: ORDER BY ValidFrom DESC (Null last), PriceListItemID DESC."""
    valid = [i for i in items if i.list_id == list_id and i.product == product and i.unit == unit
             and (i.valid_from is None or i.valid_from <= day) and (i.valid_to is None or i.valid_to >= day)]
    if not valid:
        return None
    valid.sort(key=lambda i: (i.valid_from is not None, i.valid_from or date.min, i.item_id), reverse=True)
    return Decimal(str(valid[0].price))


def customer_price(items, list_id, discount, product, unit, factor, usual, day: date) -> Decimal:
    """UnitSalePrice for a customer whose active list is list_id (0 = none); discount is a rate (0.1 = 10 %)."""
    if not list_id:
        return money(usual)
    v = item_price(items, list_id, product, unit, day)
    if v is not None:
        return money(v)
    if unit is not None:
        v = item_price(items, list_id, product, None, day)
        if v is not None:
            return money(v * Decimal(str(factor)))
    if discount:
        return money(Decimal(str(usual)) * (1 - Decimal(str(discount))))
    return money(usual)


def item_problem(items, list_id, product, unit, price, valid_from, valid_to, item_id=0) -> str:
    """PriceListItemProblem (without the unit check): '' when the price can be saved."""
    if not list_id or not product:
        return "choose"
    if not price or Decimal(str(price)) <= 0:
        return "price"
    if valid_from and valid_to and valid_to < valid_from:
        return "dates"
    if any(i.list_id == list_id and i.product == product and i.unit == unit and i.valid_from == valid_from
           and i.item_id != item_id for i in items):
        return "same start"
    return ""
