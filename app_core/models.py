# -*- coding: utf-8 -*-
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Optional


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

    def last_purchase_date_display(self) -> str:
        value = self.last_purchase_date
        if value in (None, ""):
            return "Sem compra"
        if isinstance(value, datetime):
            return value.strftime("%d/%m/%Y %H:%M")
        if isinstance(value, date):
            return value.strftime("%d/%m/%Y")
        return str(value)

    def credit_limit_display(self) -> str:
        value = self.credit_limit
        if value in (None, ""):
            return "0,00"
        try:
            num = float(value)
            return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return str(value)


@dataclass
class InvoiceRow:
    invoice_id: Any
    company: str
    customer_id: Any
    customer_code: Any
    customer_name: str
    motive_code: Any = ""
    motive_name: str = ""
    account_code: str = ""
    account_name: str = ""
    issue_date: Optional[Any] = None
    due_date: Optional[Any] = None
    amount: Any = 0
    discount_amount: Any = 0
    paid_amount: Any = 0
    open_balance: Any = 0
    customer_email: str = ""

    def issue_date_display(self) -> str:
        value = self.issue_date
        if value in (None, ""):
            return ""
        if isinstance(value, datetime):
            return value.strftime("%d/%m/%Y")
        if isinstance(value, date):
            return value.strftime("%d/%m/%Y")
        return str(value)

    def due_date_display(self) -> str:
        value = self.due_date
        if value in (None, ""):
            return ""
        if isinstance(value, datetime):
            return value.strftime("%d/%m/%Y")
        if isinstance(value, date):
            return value.strftime("%d/%m/%Y")
        return str(value)

    def amount_display(self) -> str:
        return self._money_display(self.amount)

    def discount_amount_display(self) -> str:
        return self._money_display(self.discount_amount)

    def paid_amount_display(self) -> str:
        return self._money_display(self.paid_amount)

    def open_balance_display(self) -> str:
        return self._money_display(self.open_balance)

    @staticmethod
    def _money_display(value: Any) -> str:
        if value in (None, ""):
            return "0,00"
        try:
            num = float(value)
            return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return str(value)
