# -*- coding: utf-8 -*-
import calendar
import hashlib
import hmac
import json
import re
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

LICENSE_SECRET = "DATABREV-LICENSE-2026"


def normalize_document(value: str) -> str:
    return re.sub(r"\D", "", str(value or ""))


def normalize_text_for_filename(value: str) -> str:
    value = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode("ascii")
    value = re.sub(r"[^A-Za-z0-9]+", "_", value).strip("_")
    return value or "cliente"


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
        return "0" if remainder < 2 else str(11 - remainder)

    digit_1 = calc_digit(cnpj[:12])
    digit_2 = calc_digit(cnpj[:12] + digit_1)
    return cnpj[-2:] == digit_1 + digit_2


def add_months(date_obj: date, months: int) -> date:
    total_month = date_obj.month - 1 + months
    year = date_obj.year + total_month // 12
    month = total_month % 12 + 1
    day = min(date_obj.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def license_signature(customer_document: str, expires_at: str) -> str:
    payload = f"{normalize_document(customer_document)}|{expires_at}".encode("utf-8")
    return hmac.new(LICENSE_SECRET.encode("utf-8"), payload, hashlib.sha256).hexdigest()


def build_license_filename(customer_name: str, issue_dt: datetime, directory: Path) -> Path:
    base_name = normalize_text_for_filename(customer_name)
    timestamp = issue_dt.strftime("%Y%m%d_%H%M%S")
    filename = f"{base_name}_{timestamp}.key"
    path = directory / filename

    if not path.exists():
        return path

    counter = 1
    while True:
        path = directory / f"{base_name}_{timestamp}_{counter:02d}.key"
        if not path.exists():
            return path
        counter += 1


class KeyGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gerador de chave DataHub")
        self.geometry("600x330")
        self.resizable(False, False)
        self._build()

    def _build(self):
        wrapper = ttk.Frame(self, padding=18)
        wrapper.pack(fill="both", expand=True)

        ttk.Label(wrapper, text="Gerador de chave", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 8))
        ttk.Label(
            wrapper,
            text=(
                "Informe o nome, o CNPJ do cliente e a validade da licença. "
                "O arquivo será salvo com nome único usando empresa e data/hora, sem sobrescrever chaves antigas."
            ),
            wraplength=520,
            justify="left",
        ).pack(anchor="w", pady=(0, 14))

        form = ttk.Frame(wrapper)
        form.pack(fill="x")

        ttk.Label(form, text="Cliente").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=6)
        self.client_entry = ttk.Entry(form, width=40)
        self.client_entry.grid(row=0, column=1, sticky="ew", pady=6)

        ttk.Label(form, text="CNPJ").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=6)
        self.doc_entry = ttk.Entry(form, width=40)
        self.doc_entry.grid(row=1, column=1, sticky="ew", pady=6)

        ttk.Label(form, text="Validade").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=6)
        self.months_var = tk.StringVar(value="3 meses")
        ttk.Combobox(
            form,
            textvariable=self.months_var,
            values=["1 minuto (teste)", "3 meses", "6 meses", "12 meses"],
            state="readonly",
            width=18,
        ).grid(row=2, column=1, sticky="w", pady=6)

        form.columnconfigure(1, weight=1)

        btns = ttk.Frame(wrapper)
        btns.pack(fill="x", pady=(18, 0))
        ttk.Button(btns, text="Gerar chave", command=self.generate).pack(side="right")

    def generate(self):
        client_name = self.client_entry.get().strip()
        customer_document = normalize_document(self.doc_entry.get())
        validity_label = self.months_var.get().strip()

        if not client_name:
            messagebox.showerror("DataHub", "Informe o nome do cliente.", parent=self)
            return

        if not is_valid_cnpj(customer_document):
            messagebox.showerror("DataHub", "Informe um CNPJ válido com 14 dígitos e DV correto.", parent=self)
            return

        issue_dt = datetime.now()

        if validity_label == "1 minuto (teste)":
            expires_dt = issue_dt + timedelta(minutes=1)
            validity_value = "1_minuto"
        elif validity_label == "3 meses":
            expires_dt = datetime.combine(add_months(issue_dt.date(), 3), datetime.max.time().replace(microsecond=0))
            validity_value = "3_meses"
        elif validity_label == "6 meses":
            expires_dt = datetime.combine(add_months(issue_dt.date(), 6), datetime.max.time().replace(microsecond=0))
            validity_value = "6_meses"
        elif validity_label == "12 meses":
            expires_dt = datetime.combine(add_months(issue_dt.date(), 12), datetime.max.time().replace(microsecond=0))
            validity_value = "12_meses"
        else:
            messagebox.showerror("DataHub", "Selecione uma validade válida.", parent=self)
            return

        expires_at = expires_dt.strftime("%Y-%m-%d %H:%M:%S")
        payload = {
            "customer_name": client_name,
            "customer_document": customer_document,
            "issue_date": issue_dt.strftime("%Y-%m-%d %H:%M:%S"),
            "expires_at": expires_at,
            "validity": validity_value,
            "validity_label": validity_label,
        }
        payload["signature"] = license_signature(customer_document, expires_at)

        path = build_license_filename(client_name, issue_dt, Path.cwd())

        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=4)

        messagebox.showinfo(
            "DataHub",
            f"Chave gerada com sucesso:\n{path.name}\n\nCaminho completo:\n{path}\n\nValidade: {validity_label}",
            parent=self,
        )


if __name__ == "__main__":
    app = KeyGeneratorApp()
    app.mainloop()
