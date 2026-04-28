"""Mixin modules for the Qt legacy backend bridge."""

from .billing import BillingBridgeMixin
from .dashboards import DashboardBridgeMixin
from .planning import PlanningBridgeMixin
from .purchasing import PurchasingBridgeMixin
from .quotes import QuotesBridgeMixin
from .shipping import ShippingBridgeMixin
from .transport import TransportBridgeMixin

__all__ = [
    "BillingBridgeMixin",
    "DashboardBridgeMixin",
    "PlanningBridgeMixin",
    "PurchasingBridgeMixin",
    "QuotesBridgeMixin",
    "ShippingBridgeMixin",
    "TransportBridgeMixin",
]
