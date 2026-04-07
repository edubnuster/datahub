# -*- coding: utf-8 -*-
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Optional


def format_decimal_br(value: Any) -> str:
    if value in (None, ""):
        return "0,00"
    try:
        num = float(value)
        return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(value)


def format_date_br(value: Any) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y %H:%M")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    return str(value)


@dataclass
class CustomerRow:
    customer_id: Any
    customer_code: Any
    customer_name: str
    last_purchase_date: Optional[Any]
    last_purchase_company: Optional[str]
    account_name: Optional[str]
    customer_status: str
    has_account: bool = False
    credit_limit: Any = 0
    selected: bool = False

    def checkbox(self) -> str:
        return "☑" if self.selected else "☐"

    def last_purchase_date_display(self) -> str:
        value = self.last_purchase_date
        if value in (None, ""):
            return "Sem compra"
        return format_date_br(value)

    def credit_limit_display(self) -> str:
        return format_decimal_br(self.credit_limit)


@dataclass
class InvoiceRow:
    invoice_id: Any
    company: str
    customer_id: Any
    customer_code: Any
    customer_name: str
    motive_code: Any
    motive_name: str
    account_code: str
    account_name: str
    issue_date: Optional[Any]
    due_date: Optional[Any]
    amount: Any
    discount_amount: Any
    paid_amount: Any
    open_balance: Any

    def issue_date_display(self) -> str:
        return format_date_br(self.issue_date)

    def due_date_display(self) -> str:
        return format_date_br(self.due_date)

    def amount_display(self) -> str:
        return format_decimal_br(self.amount)

    def discount_amount_display(self) -> str:
        return format_decimal_br(self.discount_amount)

    def paid_amount_display(self) -> str:
        return format_decimal_br(self.paid_amount)

    def open_balance_display(self) -> str:
        return format_decimal_br(self.open_balance)
