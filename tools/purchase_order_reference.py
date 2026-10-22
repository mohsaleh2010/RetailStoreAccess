"""Python mirror of the rules of modPurchaseOrders (docs/74, purchase orders and the purchase suggestion)."""
from decimal import Decimal


def suggest_qty(current, on_order, minimum, reorder) -> Decimal:
    """SuggestQty / PurchaseSuggestQuery: 0 above the reorder level, else the reorder quantity, else what brings
    the stock (with what is on order) to twice the level."""
    current, on_order, minimum, reorder = (Decimal(str(x)) for x in (current, on_order, minimum, reorder))
    if minimum <= 0 or current + on_order > minimum:
        return Decimal(0)
    if reorder > 0:
        return reorder
    return 2 * minimum - current - on_order


def status_for(received, invoiced, remaining) -> str:
    """PurchaseOrderStatusFor: the status of an open order from its receipt notes."""
    if received == 0:
        return "APPROVED"
    if remaining > 0:
        return "PARTIAL"
    if invoiced >= received:
        return "INVOICED"
    return "RECEIVED"
