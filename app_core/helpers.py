# -*- coding: utf-8 -*-
import re
import calendar
from datetime import date

class AppError(Exception):
    pass

def normalize_document(value: str) -> str:
    return re.sub(r"\D", "", str(value or ""))

def is_valid_cnpj(value: str) -> bool:
    cnpj = normalize_document(value)
    if len(cnpj) != 14:
        return False
    if cnpj == cnpj[0] * 14:
        return False

    def calc_digit(base: str) -> str:
        if len(base) == 12:
            weights = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        else:
            weights = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        total = sum(int(d) * w for d, w in zip(base, weights))
        remainder = total % 11
        return '0' if remainder < 2 else str(11 - remainder)

    first_digit = calc_digit(cnpj[:12])
    second_digit = calc_digit(cnpj[:12] + first_digit)
    return cnpj[-2:] == first_digit + second_digit

def add_months(date_obj: date, months: int) -> date:
    total_month = date_obj.month - 1 + months
    year = date_obj.year + total_month // 12
    month = total_month % 12 + 1
    day = min(date_obj.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)
