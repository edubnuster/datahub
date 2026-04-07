# -*- coding: utf-8 -*-
from copy import deepcopy
import smtplib
import ssl
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from typing import Any, Dict, List, Optional
import re
import tkinter as tk
from tkinter import ttk, messagebox
from email.message import EmailMessage
from app_core.audit import AuditLogger
from app_core.auth import UserManager
from app_core.config_manager import ConfigManager
from app_core.constants import APP_TITLE, CONFIG_PATH, LICENSE_FILENAME, MASTER_PASSWORD, MASTER_USERNAME
from app_core.database import Database
from app_core.helpers import AppError
from app_core.license_manager import LicenseManager
from app_core.models import CustomerRow, InvoiceRow


DATE_INPUT_FORMAT = "%d/%m/%Y"


def format_date_br(value: date) -> str:
    return value.strftime(DATE_INPUT_FORMAT)


def parse_flexible_date(value: str):
    value = (value or "").strip()
    if not value:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def bind_date_entry_shortcuts(entry):
    def _set_entry_text(value: str):
        entry.delete(0, "end")
        entry.insert(0, value)
        entry.icursor("end")

    def _apply_shortcut(action: str):
        if action == "today":
            new_value = date.today()
        else:
            current_value = parse_flexible_date(entry.get()) or date.today()
            delta_days = -1 if action == "minus" else 1
            new_value = current_value + timedelta(days=delta_days)

        _set_entry_text(format_date_br(new_value))

    def _format_typed_date(event=None):
        keysym = getattr(event, "keysym", "")
        if keysym in (
            "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R",
            "Left", "Right", "Up", "Down", "Home", "End", "Tab", "Return",
            "Escape"
        ):
            return None

        digits = re.sub(r"\D", "", entry.get())[:8]
        if not digits:
            if entry.get():
                _set_entry_text("")
            return None

        if len(digits) <= 2:
            formatted = digits
        elif len(digits) <= 4:
            formatted = f"{digits[:2]}/{digits[2:]}"
        else:
            formatted = f"{digits[:2]}/{digits[2:4]}/{digits[4:]}"

        if entry.get() != formatted:
            _set_entry_text(formatted)
        return None

    def _handle_keypress(event=None):
        keysym = getattr(event, "keysym", "")
        char = getattr(event, "char", "")

        if keysym in ("equal", "KP_Equal") or char == "=":
            _apply_shortcut("today")
            return "break"

        if keysym in ("minus", "KP_Subtract") or char == "-":
            _apply_shortcut("minus")
            return "break"

        if keysym in ("plus", "KP_Add") or char == "+":
            _apply_shortcut("plus")
            return "break"

        return None

    entry.bind("<KeyPress>", _handle_keypress, add="+")
    entry.bind("<KeyRelease>", _format_typed_date, add="+")

def money_br(value: Any) -> str:
    if value in (None, ""):
        return "0,00"
    try:
        num = float(value)
        return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(value)


def _pdf_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def build_text_pdf_bytes(lines: List[str], title: str = "Boleto") -> bytes:
    page_width = 595
    page_height = 842
    start_x = 40
    start_y = 800
    line_height = 16

    content_lines = ["BT", "/F1 11 Tf", f"{start_x} {start_y} Td"]
    first = True
    for raw_line in [title, ""] + list(lines):
        line = _pdf_escape(str(raw_line))
        if first:
            content_lines.append(f"({line}) Tj")
            first = False
        else:
            content_lines.append(f"0 -{line_height} Td")
            content_lines.append(f"({line}) Tj")
    content_lines.append("ET")
    stream = "\n".join(content_lines).encode("latin-1", errors="replace")

    objects = []
    objects.append(b"1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n")
    objects.append(b"2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n")
    objects.append(
        f"3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {page_width} {page_height}] "
        f"/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n".encode("latin-1")
    )
    objects.append(b"4 0 obj\n<< /Length " + str(len(stream)).encode("ascii") + b" >>\nstream\n" + stream + b"\nendstream\nendobj\n")
    objects.append(b"5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n")

    pdf = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for obj in objects:
        offsets.append(len(pdf))
        pdf.extend(obj)

    xref_start = len(pdf)
    pdf.extend(f"xref\n0 {len(objects)+1}\n".encode("ascii"))
    pdf.extend(b"0000000000 65535 f \n")
    for off in offsets[1:]:
        pdf.extend(f"{off:010d} 00000 n \n".encode("ascii"))
    pdf.extend(
        f"trailer\n<< /Size {len(objects)+1} /Root 1 0 R >>\nstartxref\n{xref_start}\n%%EOF".encode("ascii")
    )
    return bytes(pdf)


def build_boleto_pdf_bytes(boleto_data: Dict[str, Any], invoice_row: "InvoiceRow") -> bytes:
    company = invoice_row.company or ""
    account_display = (f"{invoice_row.account_code or ''} - {invoice_row.account_name or ''}").strip(" -")
    lines = [
        f"Empresa: {company}",
        f"Cliente: {invoice_row.customer_name}",
        f"Código do cliente: {invoice_row.customer_code}",
        f"Conta: {account_display}",
        "",
        "Dados do boleto",
        f"Documento: {boleto_data.get('documento', '')}",
        f"Vencimento: {boleto_data.get('vencto_display', '')}",
        f"Valor: {boleto_data.get('valor_display', '')}",
        f"Nosso número: {boleto_data.get('nosso_numero', '')}",
        f"Linha digitável: {boleto_data.get('linha_digitavel', '')}",
        f"Código de barras: {boleto_data.get('codigo_barra', '')}",
        "",
        "Sacado",
        f"Nome: {boleto_data.get('sacado_nome', '')}",
        f"Documento: {boleto_data.get('sacado_inscricao', '')}",
        f"Endereço: {boleto_data.get('sacado_endereco', '')}",
        f"Cidade/UF: {boleto_data.get('sacado_cidade_uf', '')}",
        "",
        "Portador",
        f"Nome: {boleto_data.get('portador_nome', '')}",
        f"Código: {boleto_data.get('portador_codigo', '')}",
        f"Carteira: {boleto_data.get('portador_carteira', '')}",
        f"Convênio: {boleto_data.get('portador_convenio', '')}",
        f"Conta corrente: {boleto_data.get('portador_conta_corrente', '')}",
    ]
    mensagem = str(boleto_data.get("mensagem", "") or "").strip()
    if mensagem:
        lines.extend(["", "Instruções", mensagem])
    return build_text_pdf_bytes(lines, title="Boleto para pagamento")
class SimpleDialog(tk.Toplevel):
    def __init__(self, master, title: str, size: str):
        super().__init__(master)
        self.title(title)
        self.geometry(size)
        self.resizable(True, True)
        self.transient(master)
        self.grab_set()
        self.lift()
        self.focus_force()
        self._center()
    def _center(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - self.winfo_width()) // 2
        y = (sh - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 50)}+{max(y, 50)}")
class ConfigWindow(SimpleDialog):
    def __init__(self, master, config_data: Dict[str, Any], on_save):
        self.config_data = deepcopy(config_data)
        self.on_save = on_save
        super().__init__(master, "Configuração local", "820x760")
        self.minsize(820, 760)
        self._build()

    def _build(self):
        outer = ttk.Frame(self, padding=14)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        header = ttk.Frame(outer)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(
            header,
            text="Configuração do sistema",
            font=("Segoe UI", 11, "bold"),
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Label(
            header,
            text="Informe os dados de conexão com o banco e o SMTP para envio de e-mails.",
            wraplength=760,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(0, 12))

        content = ttk.Frame(outer)
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)

        self.entries = {}

        db_box = ttk.LabelFrame(content, text="Banco de dados", padding=12)
        db_box.grid(row=0, column=0, sticky="ew")
        db_box.columnconfigure(0, weight=1)

        form = ttk.Frame(db_box)
        form.grid(row=0, column=0, sticky="ew")
        form.columnconfigure(1, weight=1)

        conn = self.config_data["connection"]
        fields = [
            ("Host", "host"),
            ("Porta", "port"),
            ("Banco", "dbname"),
            ("Usuário", "user"),
            ("Senha", "password"),
            ("Encoding", "client_encoding"),
        ]
        for i, (label, key) in enumerate(fields):
            ttk.Label(form, text=label).grid(row=i, column=0, sticky="w", padx=(0, 8), pady=5)
            entry = ttk.Entry(form, show="*" if key == "password" else None)
            entry.grid(row=i, column=1, sticky="ew", pady=5)
            entry.insert(0, str(conn.get(key, "")))
            self.entries[key] = entry

        smtp_box = ttk.LabelFrame(content, text="SMTP", padding=12)
        smtp_box.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        smtp_box.columnconfigure(0, weight=1)

        smtp_form = ttk.Frame(smtp_box)
        smtp_form.grid(row=0, column=0, sticky="ew")
        smtp_form.columnconfigure(1, weight=1)

        smtp_cfg = self.config_data.get("smtp", {})
        smtp_fields = [
            ("E-mail remetente", "smtp_email"),
            ("Servidor SMTP", "smtp_host"),
            ("Senha", "smtp_password"),
            ("Porta", "smtp_port"),
        ]
        for i, (label, key) in enumerate(smtp_fields):
            ttk.Label(smtp_form, text=label).grid(row=i, column=0, sticky="w", padx=(0, 8), pady=5)
            entry = ttk.Entry(smtp_form, show="*" if key == "smtp_password" else None)
            entry.grid(row=i, column=1, sticky="ew", pady=5)
            entry.insert(0, str(smtp_cfg.get(key.replace("smtp_", ""), smtp_cfg.get(key, ""))))
            self.entries[key] = entry

        buttons = ttk.Frame(self, padding=(14, 0, 14, 14))
        buttons.pack(side="bottom", fill="x")
        ttk.Button(buttons, text="Testar conexão", command=self._test).pack(side="left")
        ttk.Button(buttons, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="Salvar conexão", command=self._save).pack(side="right", padx=(0, 8))

    def _collect(self):
        cfg = deepcopy(self.config_data)
        try:
            port = int(self.entries["port"].get().strip())
        except Exception:
            raise AppError("A porta deve ser numérica.")
        cfg["connection"]["host"] = self.entries["host"].get().strip()
        cfg["connection"]["port"] = port
        cfg["connection"]["dbname"] = self.entries["dbname"].get().strip()
        cfg["connection"]["user"] = self.entries["user"].get().strip()
        cfg["connection"]["password"] = self.entries["password"].get().strip()
        cfg["connection"]["client_encoding"] = self.entries["client_encoding"].get().strip()

        smtp_port_raw = self.entries["smtp_port"].get().strip()
        if smtp_port_raw:
            try:
                smtp_port = int(smtp_port_raw)
            except Exception:
                raise AppError("A porta SMTP deve ser numérica.")
        else:
            smtp_port = 587

        cfg["smtp"] = {
            "email": self.entries["smtp_email"].get().strip(),
            "host": self.entries["smtp_host"].get().strip(),
            "password": self.entries["smtp_password"].get().strip(),
            "port": smtp_port,
        }

        if not cfg["connection"]["host"]:
            raise AppError("Informe o host.")
        if not cfg["connection"]["dbname"]:
            raise AppError("Informe o nome do banco.")
        if not cfg["connection"]["user"]:
            raise AppError("Informe o usuário.")
        return cfg

    def _test(self):
        try:
            cfg = self._collect()
            Database(cfg).test_connection()
            messagebox.showinfo(APP_TITLE, "Conexão realizada com sucesso.", parent=self)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Falha ao testar conexão:\n\n{e}", parent=self)

    def _save(self):
        try:
            cfg = self._collect()
            ConfigManager.save(cfg)
            self.on_save(cfg)
            messagebox.showinfo(APP_TITLE, f"Configuração salva em:\n{CONFIG_PATH}", parent=self)
            self.destroy()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Erro ao salvar configuração:\n\n{e}", parent=self)


class EmailComposeWindow(SimpleDialog):

    def __init__(self, master, config_data: Dict[str, Any], current_user: str, invoice_row: InvoiceRow, customer_email: str):
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.invoice_row = invoice_row
        self.customer_email = customer_email or ""
        self.boleto_data = {}
        self.attachment_bytes = None
        self.attachment_name = ""
        self.boleto_status_text = ""
        self._prepare_boleto_attachment()
        super().__init__(master, "Enviar fatura por e-mail", "760x660")
        self._build()

    def _prepare_boleto_attachment(self):
        try:
            payload = Database(self.config_data).get_boleto_email_payload(self.invoice_row.invoice_id)
        except Exception:
            payload = {"exists": False, "email_note": "Observação: não foi possível consultar os dados do boleto neste momento."}

        self.boleto_data = payload or {}
        if not self.boleto_data.get("exists"):
            self.boleto_status_text = "Boleto ainda não gerado"
            return

        attachment_data = self.boleto_data.get("attachment_data")
        filename = self.boleto_data.get("filename") or f"boleto_{self.invoice_row.invoice_id}.pdf"
        if attachment_data:
            self.attachment_bytes = attachment_data
            self.attachment_name = filename
            self.boleto_status_text = "Boleto localizado e anexado"
            return

        try:
            generated = build_boleto_pdf_bytes(self.boleto_data, self.invoice_row)
            self.attachment_bytes = generated
            self.attachment_name = filename
            self.boleto_status_text = "Boleto gerado pelo app e anexado"
        except Exception:
            self.boleto_status_text = "Boleto localizado, mas sem anexo automático"

    def _default_subject(self) -> str:
        due_text = self.invoice_row.due_date_display()
        return f"Fatura a receber - {self.invoice_row.customer_name} - vencimento {due_text}"

    def _default_body(self) -> str:
        company = self.invoice_row.company or ""
        account_display = (f"{self.invoice_row.account_code or ''} - {self.invoice_row.account_name or ''}").strip(" -")
        note = str(self.boleto_data.get("email_note", "") or "").strip()
        if not note:
            if self.boleto_data.get("exists"):
                if self.attachment_bytes:
                    note = "Observação: o boleto segue em anexo."
                else:
                    note = "Observação: foi localizado um boleto, mas não foi possível anexá-lo automaticamente."
            else:
                note = "Observação: o boleto ainda não foi gerado."
        return (
            f"Prezado(a),\n\n"
            f"Segue abaixo os dados da fatura para conferência e programação do pagamento.\n\n"
            f"Empresa: {company}\n"
            f"Cliente: {self.invoice_row.customer_name}\n"
            f"Código do cliente: {self.invoice_row.customer_code}\n"
            f"Conta: {account_display}\n"
            f"Emissão: {self.invoice_row.issue_date_display()}\n"
            f"Vencimento: {self.invoice_row.due_date_display()}\n"
            f"Valor original: {self.invoice_row.amount_display()}\n"
            f"Desconto: {self.invoice_row.discount_amount_display()}\n"
            f"Saldo em aberto: {self.invoice_row.open_balance_display()}\n\n"
            f"{note}\n\n"
            f"Em caso de dúvidas, ficamos à disposição."
        )

    def _build(self):
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Envio de fatura por e-mail", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 8))
        info = ttk.Frame(frm)
        info.pack(fill="x", pady=(0, 10))
        ttk.Label(info, text=f"Cliente: {self.invoice_row.customer_name}").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Label(info, text=f"Vencimento: {self.invoice_row.due_date_display()}").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Label(info, text=f"Saldo em aberto: {self.invoice_row.open_balance_display()}").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Label(info, text=f"Status do boleto: {self.boleto_status_text}").grid(row=3, column=0, sticky="w", pady=2)

        form = ttk.Frame(frm)
        form.pack(fill="x", pady=(0, 10))
        ttk.Label(form, text="Para").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.to_var = tk.StringVar(value=self.customer_email)
        ttk.Entry(form, textvariable=self.to_var, width=72).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(form, text="Assunto").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.subject_var = tk.StringVar(value=self._default_subject())
        ttk.Entry(form, textvariable=self.subject_var, width=72).grid(row=1, column=1, sticky="ew", pady=4)
        form.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Mensagem").pack(anchor="w")
        self.body_text = tk.Text(frm, wrap="word", height=18)
        self.body_text.pack(fill="both", expand=True, pady=(4, 0))
        self.body_text.insert("1.0", self._default_body())

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(12, 0))
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Enviar", command=self._send_email).pack(side="right", padx=(0, 8))

    def _send_email(self):
        to_email = self.to_var.get().strip()
        subject = self.subject_var.get().strip()
        body = self.body_text.get("1.0", "end").strip()
        smtp_cfg = self.config_data.get("smtp", {})
        smtp_email = str(smtp_cfg.get("email", "")).strip()
        smtp_host = str(smtp_cfg.get("host", "")).strip()
        smtp_password = str(smtp_cfg.get("password", "")).strip()
        smtp_port = int(smtp_cfg.get("port", 587) or 587)

        if not to_email:
            messagebox.showwarning(APP_TITLE, "Informe o e-mail do destinatário.", parent=self)
            return
        if not subject:
            messagebox.showwarning(APP_TITLE, "Informe o assunto do e-mail.", parent=self)
            return
        if not body:
            messagebox.showwarning(APP_TITLE, "Informe a mensagem do e-mail.", parent=self)
            return
        if not smtp_email or not smtp_host or not smtp_password or not smtp_port:
            messagebox.showwarning(APP_TITLE, "Configure o SMTP na tela de Configuração antes de enviar e-mails.", parent=self)
            return

        msg = EmailMessage()
        msg["From"] = smtp_email
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.set_content(body)

        if self.attachment_bytes:
            msg.add_attachment(
                self.attachment_bytes,
                maintype="application",
                subtype="pdf",
                filename=self.attachment_name or f"boleto_{self.invoice_row.invoice_id}.pdf",
            )

        try:
            if smtp_port == 465:
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context, timeout=20) as server:
                    server.login(smtp_email, smtp_password)
                    server.send_message(msg)
            else:
                with smtplib.SMTP(smtp_host, smtp_port, timeout=20) as server:
                    server.ehlo()
                    try:
                        server.starttls(context=ssl.create_default_context())
                        server.ehlo()
                    except Exception:
                        pass
                    server.login(smtp_email, smtp_password)
                    server.send_message(msg)
            AuditLogger.write(
                self.current_user,
                "enviar_email_fatura",
                f"cliente={self.invoice_row.customer_name};para={to_email};invoice={self.invoice_row.invoice_id};anexo_pdf={'sim' if self.attachment_bytes else 'nao'}"
            )
            messagebox.showinfo(APP_TITLE, "E-mail enviado com sucesso.", parent=self)
            self.destroy()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Falha ao enviar e-mail:\n\n{e}", parent=self)
class CreateUserWindow(SimpleDialog):
    def __init__(self, master, config_data: Dict[str, Any], current_user: str, on_save):
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.on_save = on_save
        super().__init__(master, "Cadastrar usuário", "500x320")
        self._build()
    def _build(self):
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Cadastrar novo usuário", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 6))
        ttk.Label(frm, text="Cadastre um novo usuário local para acesso ao sistema.", wraplength=450, justify="left").pack(anchor="w", pady=(0, 14))
        form = ttk.Frame(frm)
        form.pack(fill="x")
        ttk.Label(form, text="Usuário").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.user_entry = ttk.Entry(form, width=34)
        self.user_entry.grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Label(form, text="Senha").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        self.pass_entry = ttk.Entry(form, width=34, show="*")
        self.pass_entry.grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Label(form, text="Confirmar senha").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)
        self.confirm_entry = ttk.Entry(form, width=34, show="*")
        self.confirm_entry.grid(row=2, column=1, sticky="ew", pady=5)
        form.columnconfigure(1, weight=1)
        ttk.Frame(frm).pack(fill="both", expand=True)
        btns = ttk.Frame(frm)
        btns.pack(fill="x", side="bottom", pady=(20, 8))
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Cadastrar", command=self._create).pack(side="right", padx=(0, 8))
    def _create(self):
        try:
            username = self.user_entry.get().strip()
            password = self.pass_entry.get()
            confirm = self.confirm_entry.get()
            if not username:
                raise AppError("Informe o usuário.")
            if not password:
                raise AppError("Informe a senha.")
            if password != confirm:
                raise AppError("A confirmação da senha não confere.")
            UserManager.add_user(self.config_data, username, password)
            ConfigManager.save(self.config_data)
            self.on_save(self.config_data)
            AuditLogger.write(self.current_user, "cadastrar_usuario", f"alvo={username}")
            messagebox.showinfo(APP_TITLE, "Usuário cadastrado com sucesso.", parent=self)
            self.destroy()
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e), parent=self)
class ChangeOwnPasswordWindow(SimpleDialog):
    def __init__(self, master, config_data: Dict[str, Any], current_user: str, on_save):
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.on_save = on_save
        super().__init__(master, "Alterar senha", "500x300")
        self._build()
    def _build(self):
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Alterar senha", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 6))
        ttk.Label(frm, text=f"Usuário: {self.current_user}", wraplength=450, justify="left").pack(anchor="w", pady=(0, 14))
        form = ttk.Frame(frm)
        form.pack(fill="x")
        ttk.Label(form, text="Nova senha").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.pass_entry = ttk.Entry(form, width=34, show="*")
        self.pass_entry.grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Label(form, text="Confirmar senha").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        self.confirm_entry = ttk.Entry(form, width=34, show="*")
        self.confirm_entry.grid(row=1, column=1, sticky="ew", pady=5)
        form.columnconfigure(1, weight=1)
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(20, 0))
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Salvar alteração", command=self._save).pack(side="right", padx=(0, 8))
    def _save(self):
        try:
            password = self.pass_entry.get()
            confirm = self.confirm_entry.get()
            if not password:
                raise AppError("Informe a nova senha.")
            if password != confirm:
                raise AppError("A confirmação da senha não confere.")
            if not UserManager.find_user(self.config_data, self.current_user):
                raise AppError("Usuário não encontrado.")
            UserManager.update_user_password(self.config_data, self.current_user, password)
            ConfigManager.save(self.config_data)
            self.on_save(self.config_data)
            AuditLogger.write(self.current_user, "alterar_senha", f"alvo={self.current_user}")
            messagebox.showinfo(APP_TITLE, "Senha alterada com sucesso.", parent=self)
            self.destroy()
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e), parent=self)
class MasterManageUsersWindow(SimpleDialog):
    def __init__(self, master, config_data: Dict[str, Any], current_user: str, on_save):
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.on_save = on_save
        self.selected_username = None
        self.users_by_item = {}
        super().__init__(master, "Alterar usuários", "620x430")
        self._build()
        self._load_users()
    def _build(self):
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Alterar usuários", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 6))
        ttk.Label(frm, text="Selecione um usuário para alterar a senha ou excluir. O usuário master não pode excluir a si próprio.", wraplength=580, justify="left").pack(anchor="w", pady=(0, 12))
        mid = ttk.Frame(frm)
        mid.pack(fill="both", expand=True)
        self.tree_users = ttk.Treeview(mid, columns=("username",), show="headings", height=10)
        self.tree_users.heading("username", text="Usuário")
        self.tree_users.column("username", width=260, anchor="w")
        self.tree_users.grid(row=0, column=0, sticky="nsew")
        self.tree_users.bind("<<TreeviewSelect>>", self._on_select)
        yscroll = ttk.Scrollbar(mid, orient="vertical", command=self.tree_users.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.tree_users.configure(yscrollcommand=yscroll.set)
        mid.rowconfigure(0, weight=1)
        mid.columnconfigure(0, weight=1)
        form = ttk.Frame(frm)
        form.pack(fill="x", pady=(14, 0))
        ttk.Label(form, text="Usuário selecionado").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.username_var = tk.StringVar(value="")
        ttk.Entry(form, textvariable=self.username_var, width=34, state="readonly").grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Label(form, text="Nova senha").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        self.pass_entry = ttk.Entry(form, width=34, show="*")
        self.pass_entry.grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Label(form, text="Confirmar senha").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)
        self.confirm_entry = ttk.Entry(form, width=34, show="*")
        self.confirm_entry.grid(row=2, column=1, sticky="ew", pady=5)
        form.columnconfigure(1, weight=1)
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(18, 0))
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Salvar alteração", command=self._save).pack(side="right", padx=(0, 8))
        ttk.Button(btns, text="Excluir usuário", command=self._remove).pack(side="left")
    def _load_users(self):
        for item in self.tree_users.get_children():
            self.tree_users.delete(item)
        self.users_by_item.clear()
        for user in UserManager.list_users(self.config_data):
            username = user.get("username", "")
            item = self.tree_users.insert("", "end", values=(username,))
            self.users_by_item[item] = username
    def _on_select(self, event=None):
        selected = self.tree_users.selection()
        if not selected:
            self.selected_username = None
            self.username_var.set("")
            return
        item = selected[0]
        self.selected_username = self.users_by_item.get(item)
        self.username_var.set(self.selected_username or "")
    def _require_selected(self):
        username = (self.selected_username or "").strip()
        if not username:
            raise AppError("Selecione um usuário.")
        return username
    def _save(self):
        try:
            username = self._require_selected()
            password = self.pass_entry.get()
            confirm = self.confirm_entry.get()
            if not password:
                raise AppError("Informe a nova senha.")
            if password != confirm:
                raise AppError("A confirmação da senha não confere.")
            if not UserManager.find_user(self.config_data, username):
                raise AppError("Usuário não encontrado.")
            UserManager.update_user_password(self.config_data, username, password)
            ConfigManager.save(self.config_data)
            self.on_save(self.config_data)
            AuditLogger.write(self.current_user, "editar_usuario", f"alvo={username}")
            self.pass_entry.delete(0, "end")
            self.confirm_entry.delete(0, "end")
            messagebox.showinfo(APP_TITLE, "Senha alterada com sucesso.", parent=self)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e), parent=self)
    def _remove(self):
        try:
            username = self._require_selected()
            if username.lower() == MASTER_USERNAME.lower():
                raise AppError("O usuário master não pode ser excluído.")
            if not UserManager.find_user(self.config_data, username):
                raise AppError("Usuário não encontrado.")
            if messagebox.askyesno(APP_TITLE, f"Deseja excluir o usuário '{username}'?", parent=self):
                UserManager.remove_user(self.config_data, username)
                ConfigManager.save(self.config_data)
                self.on_save(self.config_data)
                AuditLogger.write(self.current_user, "excluir_usuario", f"alvo={username}")
                self.selected_username = None
                self.username_var.set("")
                self.pass_entry.delete(0, "end")
                self.confirm_entry.delete(0, "end")
                self._load_users()
                messagebox.showinfo(APP_TITLE, "Usuário excluído com sucesso.", parent=self)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e), parent=self)
class InactiveCustomersWindow(tk.Toplevel):
    FILTER_OPTIONS = {
        "Todos": None,
        "Ativos": "Ativo",
        "Inativos": "Inativo",
        "Deletados": "Deletado",
        "Com limite": "__HAS_CREDIT__",
        "Possui conta": "__HAS_ACCOUNT__",
    }
    def __init__(self, master, config_data: Dict[str, Any], current_user: str, on_config_saved):
        super().__init__(master)
        self.master_app = master
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.on_config_saved = on_config_saved
        self.rows: List[CustomerRow] = []
        self.filtered_rows: List[CustomerRow] = []
        self.tree_items: Dict[str, CustomerRow] = {}
        self.filter_var = tk.StringVar(value="Todos")
        self.sort_column: Optional[str] = None
        self.sort_reverse = False
        self.status_var = tk.StringVar(value="Pronto.")
        self.title(f"{APP_TITLE} - Clientes inativos")
        self.geometry("1360x720")
        self.minsize(1260, 680)
        self.transient(master)
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._setup_style()
        self._build_ui()
        self._center_window()
        self.load_data()
    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("Treeview", rowheight=24, font=("Segoe UI", 10), background="#ffffff", fieldbackground="#ffffff")
        self.configure(background="#f5f7fb")
    def _center_window(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - self.winfo_width()) // 2
        y = (sh - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 20)}+{max(y, 20)}")
    def _build_ui(self):
        top = ttk.Frame(self, padding=(12, 12, 12, 8))
        top.pack(fill="x")
        left_actions = ttk.Frame(top)
        left_actions.pack(side="left")
        ttk.Button(left_actions, text="Atualizar lista", command=self.load_data).pack(side="left")
        ttk.Button(left_actions, text="Marcar todos", command=self.mark_all).pack(side="left", padx=(8, 0))
        ttk.Button(left_actions, text="Desmarcar todos", command=self.unmark_all).pack(side="left", padx=(8, 0))
        ttk.Button(left_actions, text="Voltar ao início", command=self._close).pack(side="left", padx=(16, 0))
        filter_box = ttk.Frame(top)
        filter_box.pack(side="left", padx=(18, 0))
        ttk.Label(filter_box, text="Mostrar:").pack(side="left", padx=(0, 6))
        filtro = ttk.Combobox(filter_box, textvariable=self.filter_var, values=list(self.FILTER_OPTIONS.keys()), state="readonly", width=12)
        filtro.pack(side="left")
        filtro.bind("<<ComboboxSelected>>", lambda e: self.apply_filter())
        actions = ttk.Frame(top)
        actions.pack(side="right")
        ttk.Button(actions, text="Inativar cliente", command=lambda: self.run_action("inactivate_customer_sql", "Inativar cliente", "Inativo")).pack(side="left")
        ttk.Button(actions, text="Excluir cliente", command=lambda: self.run_action("delete_customer_sql", "Excluir cliente", "Deletado")).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Inativar vendas a prazo", command=lambda: self.run_action("disable_credit_sql", "Inativar vendas a prazo", None)).pack(side="left", padx=(8, 0))
        middle = ttk.Frame(self, padding=(12, 0, 12, 0))
        middle.pack(fill="both", expand=True)
        columns = ("company", "code", "name", "account", "credit_limit", "last_date", "status")
        self.tree = ttk.Treeview(middle, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("company", text="Empresa da última compra", command=lambda: self.sort_by("company"))
        self.tree.heading("code", text="Código", command=lambda: self.sort_by("code"))
        self.tree.heading("name", text="Cliente", command=lambda: self.sort_by("name"))
        self.tree.heading("account", text="Conta", command=lambda: self.sort_by("account"))
        self.tree.heading("credit_limit", text="Lim. crédito", command=lambda: self.sort_by("credit_limit"))
        self.tree.heading("last_date", text="Última compra", command=lambda: self.sort_by("last_date"))
        self.tree.heading("status", text="Status", command=lambda: self.sort_by("status"))
        self.tree.column("company", width=200, minwidth=180, anchor="w", stretch=False)
        self.tree.column("code", width=80, minwidth=76, anchor="center", stretch=False)
        self.tree.column("name", width=280, minwidth=240, anchor="w", stretch=True)
        self.tree.column("account", width=170, minwidth=150, anchor="w", stretch=False)
        self.tree.column("credit_limit", width=110, minwidth=100, anchor="e", stretch=False)
        self.tree.column("last_date", width=110, minwidth=100, anchor="center", stretch=False)
        self.tree.column("status", width=95, minwidth=90, anchor="center", stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<Control-a>", self._select_all_rows)
        self.tree.bind("<Control-A>", self._select_all_rows)
        yscroll = ttk.Scrollbar(middle, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll = ttk.Scrollbar(middle, orient="horizontal", command=self.tree.xview)
        xscroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        middle.rowconfigure(0, weight=1)
        middle.columnconfigure(0, weight=1)
        bottom = ttk.Frame(self, padding=(12, 8, 12, 10))
        bottom.pack(fill="x")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left")
        ttk.Label(bottom, text=f"Usuário: {self.current_user}").pack(side="right")
    def _close(self):
        self.destroy()
        if hasattr(self.master_app, "inactive_window"):
            self.master_app.inactive_window = None
        if hasattr(self.master_app, "show_home"):
            self.master_app.show_home()
    def open_config(self):
        ConfigWindow(self, self.config_data, self._apply_new_config)
    def _apply_new_config(self, cfg: Dict[str, Any]):
        self.config_data = cfg
        self.on_config_saved(cfg)
        self.load_data()
    def set_status(self, text: str):
        self.status_var.set(text)
        self.update_idletasks()
    def _normalize_status(self, value: Any) -> str:
        txt = str(value or "").strip()
        if txt in ("A", "Ativo"):
            return "Ativo"
        if txt in ("I", "Inativo"):
            return "Inativo"
        if txt in ("D", "Deletado"):
            return "Deletado"
        return txt
    def _sort_value(self, row: CustomerRow, column: str):
        if column == "company":
            return (row.last_purchase_company or "").lower()
        if column == "code":
            try:
                return (0, int(str(row.customer_code)))
            except Exception:
                return (1, str(row.customer_code or "").lower())
        if column == "name":
            return (row.customer_name or "").lower()
        if column == "account":
            return (row.account_name or "").lower()
        if column == "credit_limit":
            try:
                return float(row.credit_limit or 0)
            except Exception:
                return 0.0
        if column == "last_date":
            value = row.last_purchase_date
            if value is None:
                return datetime.min
            if isinstance(value, datetime):
                return value
            if isinstance(value, date):
                return datetime.combine(value, time.min)
            return str(value)
        if column == "status":
            order = {"Ativo": 0, "Inativo": 1, "Deletado": 2}
            return (order.get(row.customer_status, 99), row.customer_status)
        return ""
    def sort_by(self, column: str):
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False
        self.filtered_rows.sort(key=lambda r: self._sort_value(r, column), reverse=self.sort_reverse)
        self._refresh_tree()
        self._update_heading_titles()
    def _update_heading_titles(self):
        labels = {
            "company": "Empresa da última compra",
            "code": "Código",
            "name": "Cliente",
            "account": "Conta",
            "credit_limit": "Lim. crédito",
            "last_date": "Última compra",
            "status": "Status",
        }
        for col, label in labels.items():
            suffix = ""
            if self.sort_column == col:
                suffix = " ▼" if self.sort_reverse else " ▲"
            self.tree.heading(col, text=label + suffix, command=lambda c=col: self.sort_by(c))
    def load_data(self):
        try:
            self.set_status("Conectando ao banco e carregando clientes.")
            data = Database(self.config_data).list_inactive_customers()
            self.rows = [
                CustomerRow(
                    customer_id=row.get("customer_id"),
                    customer_code=row.get("customer_code"),
                    customer_name=row.get("customer_name") or "",
                    last_purchase_date=row.get("last_purchase_date"),
                    last_purchase_company=row.get("last_purchase_company") or "",
                    account_name=row.get("account_name") or "Sem conta",
                    customer_status=self._normalize_status(row.get("customer_status")),
                    has_account=bool(row.get("has_account")),
                    credit_limit=row.get("credit_limit"),
                    selected=False,
                )
                for row in data
            ]
            self.apply_filter()
            self.set_status(f"{len(self.filtered_rows)} cliente(s) encontrado(s).")
            AuditLogger.write(self.current_user, "carregar_lista", f"tipo=clientes_inativos;quantidade={len(self.filtered_rows)}")
        except Exception as e:
            self.set_status("Falha ao carregar dados.")
            messagebox.showerror(APP_TITLE, f"Erro ao carregar clientes:\n\n{e}", parent=self)
            AuditLogger.write(self.current_user, "erro_carregar_lista", f"tipo=clientes_inativos;erro={e}")
    def apply_filter(self):
        selected_filter = self.FILTER_OPTIONS.get(self.filter_var.get())
        if selected_filter is None:
            self.filtered_rows = list(self.rows)
        elif selected_filter == "__HAS_CREDIT__":
            self.filtered_rows = [r for r in self.rows if float(r.credit_limit or 0) > 0]
        elif selected_filter == "__HAS_ACCOUNT__":
            self.filtered_rows = [r for r in self.rows if bool(r.has_account)]
        else:
            self.filtered_rows = [r for r in self.rows if r.customer_status == selected_filter]
        if self.sort_column:
            self.filtered_rows.sort(key=lambda r: self._sort_value(r, self.sort_column), reverse=self.sort_reverse)
        self._refresh_tree()
        self._update_heading_titles()
        self.set_status(f"{len(self.filtered_rows)} cliente(s) encontrado(s).")
    def _row_values(self, row: CustomerRow):
        return (
            row.last_purchase_company,
            row.customer_code,
            row.customer_name,
            row.account_name,
            row.credit_limit_display(),
            row.last_purchase_date_display(),
            row.customer_status,
        )
    def _refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree_items.clear()
        for row in self.filtered_rows:
            item_id = self.tree.insert("", "end", values=self._row_values(row))
            self.tree_items[item_id] = row


    def mark_all(self):
        items = self.tree.get_children()
        if items:
            self.tree.selection_set(items)

    def unmark_all(self):
        selected_items = self.tree.selection()
        if selected_items:
            self.tree.selection_remove(selected_items)

    def _select_all_rows(self, event=None):
        self.mark_all()
        return "break"
    def selected_rows(self) -> List[CustomerRow]:
        selected_items = self.tree.selection()
        if not selected_items:
            return []
        return [self.tree_items[item_id] for item_id in selected_items if item_id in self.tree_items]
    def run_action(self, query_key: str, action_name: str, new_status: Optional[str]):
        selected = self.selected_rows()
        if not selected:
            messagebox.showwarning(APP_TITLE, "Selecione pelo menos um cliente.", parent=self)
            return
        sql_text = (self.config_data.get("queries", {}).get(query_key) or "").strip()
        if not sql_text:
            messagebox.showwarning(APP_TITLE, "A SQL da ação não está configurada.", parent=self)
            return
        if not messagebox.askyesno(APP_TITLE, f"Deseja executar a ação '{action_name}' para {len(selected)} cliente(s)?", parent=self):
            return
        try:
            customer_ids = [r.customer_id for r in selected]
            affected = Database(self.config_data).execute_action(sql_text, customer_ids)
            if new_status:
                for row in selected:
                    row.customer_status = new_status
            if query_key == "disable_credit_sql":
                for row in selected:
                    row.credit_limit = 0
            AuditLogger.write(self.current_user, "acao_operacional", f"acao={action_name};selecionados={len(selected)};afetados={affected}")
            self.load_data()
            messagebox.showinfo(APP_TITLE, f"Ação '{action_name}' executada com sucesso.", parent=self)
        except Exception as e:
            AuditLogger.write(self.current_user, "erro_acao_operacional", f"acao={action_name};erro={e}")
            messagebox.showerror(APP_TITLE, f"Erro ao executar a ação:\n\n{e}", parent=self)
class OpenInvoicesWindow(tk.Toplevel):
    GROUP_OPTIONS = {
        "Não agrupar": "none",
        "Cliente": "customer",
        "Vencimento": "due_date",
        "Conta": "account_group",
    }
    def __init__(self, master, config_data: Dict[str, Any], current_user: str, on_config_saved):
        super().__init__(master)
        self.master_app = master
        self.config_data = deepcopy(config_data)
        self.current_user = current_user
        self.on_config_saved = on_config_saved
        self.raw_rows: List[InvoiceRow] = []
        self.rows: List[InvoiceRow] = []
        self.tree_items: Dict[str, InvoiceRow] = {}
        self.sort_column: Optional[str] = None
        self.sort_reverse = False
        self.status_var = tk.StringVar(value="Pronto.")
        today_str = datetime.now().strftime("%d/%m/%Y")
        self.period_start_var = tk.StringVar(value=today_str)
        self.period_end_var = tk.StringVar(value=today_str)
        self.group_by_var = tk.StringVar(value="Não agrupar")
        self.customer_var = tk.StringVar(value="Todos")
        self.account_var = tk.StringVar(value="Todas")
        self.customer_options_map: Dict[str, Any] = {"Todos": None}
        self.all_customer_options_map: Dict[str, Any] = {"Todos": None}
        self.account_options_map: Dict[str, Any] = {"Todas": None}
        self.all_account_options_map: Dict[str, Any] = {"Todas": None}
        self._auto_filter_job = None
        self.title(f"{APP_TITLE} - Faturas a receber")
        self.geometry("1480x760")
        self.minsize(1320, 700)
        self.resizable(True, True)
        self.transient(master)
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._setup_style()
        self._build_ui()
        self._center_window()
        self._load_customer_options()
        self._load_account_options()
        self.load_data()
    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("Treeview", rowheight=24, font=("Segoe UI", 10), background="#ffffff", fieldbackground="#ffffff")
        self.configure(background="#f5f7fb")
    def _center_window(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - self.winfo_width()) // 2
        y = (sh - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 20)}+{max(y, 20)}")
    def _build_ui(self):
        top = ttk.Frame(self, padding=(12, 12, 12, 8))
        top.pack(fill="x")
        actions = ttk.Frame(top)
        actions.pack(side="left")
        ttk.Button(actions, text="Limpar filtros", command=self.clear_filters).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Enviar fatura por e-mail", command=self.open_email_window).pack(side="left", padx=(12, 0))
        ttk.Button(actions, text="Voltar ao início", command=self._close).pack(side="left", padx=(12, 0))
        filters = ttk.Frame(self, padding=(12, 0, 12, 8))
        filters.pack(fill="x")
        ttk.Label(filters, text="Vencimento inicial").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        self.period_start_entry = ttk.Entry(filters, textvariable=self.period_start_var, width=14)
        self.period_start_entry.grid(row=0, column=1, sticky="w", pady=4)
        bind_date_entry_shortcuts(self.period_start_entry)
        self.period_start_entry.bind("<FocusOut>", self._schedule_auto_filter, add="+")
        ttk.Label(filters, text="Vencimento final").grid(row=0, column=2, sticky="w", padx=(16, 6), pady=4)
        self.period_end_entry = ttk.Entry(filters, textvariable=self.period_end_var, width=14)
        self.period_end_entry.grid(row=0, column=3, sticky="w", pady=4)
        bind_date_entry_shortcuts(self.period_end_entry)
        self.period_end_entry.bind("<FocusOut>", self._schedule_auto_filter, add="+")
        ttk.Label(filters, text="Cliente").grid(row=0, column=4, sticky="w", padx=(16, 6), pady=4)
        self.customer_entry = ttk.Entry(filters, textvariable=self.customer_var, width=28)
        self.customer_entry.grid(row=0, column=5, sticky="ew", pady=4)
        self.customer_entry.bind("<KeyRelease>", self._on_customer_keyrelease)
        self.customer_entry.bind("<FocusIn>", self._on_customer_focus_in)
        self.customer_entry.bind("<FocusOut>", self._on_customer_focus_out)
        self.customer_entry.bind("<Down>", self._on_customer_arrow_down)
        self.customer_entry.bind("<Up>", self._on_customer_arrow_up)
        self.customer_entry.bind("<Return>", self._on_customer_entry_confirm)
        self.customer_suggestions_frame = ttk.Frame(filters)
        self.customer_suggestions_frame.grid(row=1, column=4, columnspan=2, sticky="nsew", padx=(16, 0), pady=(0, 2))
        self.customer_suggestions_frame.grid_remove()
        self.customer_listbox = tk.Listbox(self.customer_suggestions_frame, height=6, exportselection=False)
        self.customer_listbox.pack(fill="x", expand=True)
        self.customer_listbox.bind("<<ListboxSelect>>", self._on_customer_listbox_select)
        self.customer_listbox.bind("<ButtonRelease-1>", self._on_customer_listbox_confirm)
        self.customer_listbox.bind("<Double-Button-1>", self._on_customer_listbox_confirm)
        self.customer_listbox.bind("<Return>", self._on_customer_listbox_confirm)

        ttk.Label(filters, text="Conta").grid(row=0, column=6, sticky="w", padx=(16, 6), pady=4)
        self.account_entry = ttk.Entry(filters, textvariable=self.account_var, width=28)
        self.account_entry.grid(row=0, column=7, sticky="ew", pady=4)
        self.account_entry.bind("<KeyRelease>", self._on_account_keyrelease)
        self.account_entry.bind("<FocusIn>", self._on_account_focus_in)
        self.account_entry.bind("<FocusOut>", self._on_account_focus_out)
        self.account_entry.bind("<Down>", self._on_account_arrow_down)
        self.account_entry.bind("<Up>", self._on_account_arrow_up)
        self.account_entry.bind("<Return>", self._on_account_entry_confirm)
        self.account_suggestions_frame = ttk.Frame(filters)
        self.account_suggestions_frame.grid(row=1, column=6, columnspan=2, sticky="nsew", padx=(16, 0), pady=(0, 2))
        self.account_suggestions_frame.grid_remove()
        self.account_listbox = tk.Listbox(self.account_suggestions_frame, height=6, exportselection=False)
        self.account_listbox.pack(fill="x", expand=True)
        self.account_listbox.bind("<<ListboxSelect>>", self._on_account_listbox_select)
        self.account_listbox.bind("<ButtonRelease-1>", self._on_account_listbox_confirm)
        self.account_listbox.bind("<Double-Button-1>", self._on_account_listbox_confirm)
        self.account_listbox.bind("<Return>", self._on_account_listbox_confirm)

        ttk.Label(filters, text="Agrupar por").grid(row=0, column=8, sticky="w", padx=(16, 6), pady=4)
        self.group_combo = ttk.Combobox(filters, textvariable=self.group_by_var, state="readonly", width=18, values=list(self.GROUP_OPTIONS.keys()))
        self.group_combo.grid(row=0, column=9, sticky="w", pady=4)
        self.group_combo.bind("<<ComboboxSelected>>", self._schedule_auto_filter)
        ttk.Label(filters, text="Formato: DD/MM/AAAA").grid(row=2, column=0, columnspan=4, sticky="w", pady=(2, 0))
        filters.columnconfigure(5, weight=1)
        filters.columnconfigure(7, weight=1)
        middle = ttk.Frame(self, padding=(12, 0, 12, 0))
        middle.pack(fill="both", expand=True)
        columns = ("company", "account_display", "code", "name", "issue_date", "due_date", "amount", "discount", "open_balance")
        self.tree = ttk.Treeview(middle, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("company", text="Empresa", command=lambda: self.sort_by("company"))
        self.tree.heading("account_display", text="Conta", command=lambda: self.sort_by("account_display"))
        self.tree.heading("code", text="Código", command=lambda: self.sort_by("code"))
        self.tree.heading("name", text="Cliente", command=lambda: self.sort_by("name"))
        self.tree.heading("issue_date", text="Data", command=lambda: self.sort_by("issue_date"))
        self.tree.heading("due_date", text="Vencimento", command=lambda: self.sort_by("due_date"))
        self.tree.heading("amount", text="Valor", command=lambda: self.sort_by("amount"))
        self.tree.heading("discount", text="Desconto", command=lambda: self.sort_by("discount"))
        self.tree.heading("open_balance", text="Saldo em aberto", command=lambda: self.sort_by("open_balance"))
        self.tree.column("company", width=180, minwidth=160, anchor="w", stretch=False)
        self.tree.column("account_display", width=260, minwidth=240, anchor="w", stretch=False)
        self.tree.column("code", width=80, minwidth=70, anchor="center", stretch=False)
        self.tree.column("name", width=250, minwidth=220, anchor="w", stretch=True)
        self.tree.column("issue_date", width=100, minwidth=90, anchor="center", stretch=False)
        self.tree.column("due_date", width=100, minwidth=90, anchor="center", stretch=False)
        self.tree.column("amount", width=110, minwidth=100, anchor="e", stretch=False)
        self.tree.column("discount", width=110, minwidth=100, anchor="e", stretch=False)
        self.tree.column("open_balance", width=130, minwidth=120, anchor="e", stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(middle, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll = ttk.Scrollbar(middle, orient="horizontal", command=self.tree.xview)
        xscroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        middle.rowconfigure(0, weight=1)
        middle.columnconfigure(0, weight=1)
        bottom = ttk.Frame(self, padding=(12, 8, 12, 10))
        bottom.pack(fill="x")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left")
        ttk.Label(bottom, text=f"Usuário: {self.current_user}").pack(side="right")
    def _close(self):
        self.destroy()
        if hasattr(self.master_app, "invoices_window"):
            self.master_app.invoices_window = None
        if hasattr(self.master_app, "show_home"):
            self.master_app.show_home()
    def open_config(self):
        ConfigWindow(self, self.config_data, self._apply_new_config)
    def _apply_new_config(self, cfg: Dict[str, Any]):
        self.config_data = cfg
        self.on_config_saved(cfg)
        self._load_customer_options()
        self._load_account_options()
        self.load_data()
    def _schedule_auto_filter(self, event=None):
        if self._auto_filter_job is not None:
            try:
                self.after_cancel(self._auto_filter_job)
            except Exception:
                pass
        self._auto_filter_job = self.after(350, self._run_auto_filter)
        return None

    def _run_auto_filter(self):
        self._auto_filter_job = None
        self.load_data()

    def clear_filters(self):
        today_str = datetime.now().strftime("%d/%m/%Y")
        self.period_start_var.set(today_str)
        self.period_end_var.set(today_str)
        self.customer_var.set("Todos")
        self.account_var.set("Todas")
        self.group_by_var.set("Não agrupar")
        self._hide_customer_suggestions()
        self._hide_account_suggestions()
        self.load_data()
    def set_status(self, text: str):
        self.status_var.set(text)
        self.update_idletasks()
    def _parse_date(self, value: str):
        value = (value or "").strip()
        if not value:
            return None
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                pass
        raise AppError("Data inválida. Use DD/MM/AAAA.")

    def _selected_customer_id(self):
        typed = (self.customer_var.get() or "").strip()
        if not typed or typed == "Todos":
            return None

        exact = self.all_customer_options_map.get(typed)
        if exact is not None:
            return exact

        typed_lower = typed.lower()
        matches = [
            customer_id
            for label, customer_id in self.all_customer_options_map.items()
            if label != "Todos" and typed_lower in label.lower()
        ]
        if len(matches) == 1:
            return matches[0]
        return None

    def _load_customer_options(self):
        current = (self.customer_var.get() or "").strip() or "Todos"
        options_map = {"Todos": None}
        try:
            rows = Database(self.config_data).list_open_invoice_customers()
            for row in rows:
                label = f"{row.get('codigo_cliente')} - {row.get('cliente')}"
                options_map[label] = row.get("customer_id")
        except Exception:
            pass

        self.all_customer_options_map = options_map
        self.customer_options_map = dict(options_map)
        self.customer_var.set(current if current in options_map else "Todos")
        self._hide_customer_suggestions()

    def _matching_customer_labels(self, typed: str):
        typed = (typed or "").strip().lower()
        if not typed:
            return [label for label in self.all_customer_options_map.keys()]
        return [
            label
            for label in self.all_customer_options_map.keys()
            if label != "Todos" and typed in label.lower()
        ]

    def _show_customer_suggestions(self, labels):
        self.customer_listbox.delete(0, "end")
        for label in labels:
            self.customer_listbox.insert("end", label)

        if labels:
            self.customer_suggestions_frame.grid()
            self.customer_listbox.selection_clear(0, "end")
            self.customer_listbox.selection_set(0)
        else:
            self._hide_customer_suggestions()

    def _hide_customer_suggestions(self):
        if hasattr(self, "customer_suggestions_frame"):
            self.customer_suggestions_frame.grid_remove()

    def _apply_customer_search(self, typed: str):
        labels = self._matching_customer_labels(typed)
        filtered = {label: self.all_customer_options_map.get(label) for label in labels}
        self.customer_options_map = filtered
        if (typed or "").strip():
            self._show_customer_suggestions(labels)
        else:
            self._hide_customer_suggestions()

    def _on_customer_keyrelease(self, event=None):
        keysym = getattr(event, "keysym", "")
        if keysym in ("Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return None
        self._apply_customer_search(self.customer_var.get())
        return None

    def _on_customer_focus_in(self, event=None):
        typed = (self.customer_var.get() or "").strip()
        if typed and typed != "Todos":
            self._apply_customer_search(typed)
        return None

    def _on_customer_focus_out(self, event=None):
        self.after(150, self._hide_customer_suggestions)
        return None

    def _on_customer_arrow_down(self, event=None):
        if not self.customer_suggestions_frame.winfo_ismapped() or self.customer_listbox.size() == 0:
            return "break"
        selection = self.customer_listbox.curselection()
        index = selection[0] + 1 if selection else 0
        if index >= self.customer_listbox.size():
            index = self.customer_listbox.size() - 1
        self.customer_listbox.selection_clear(0, "end")
        self.customer_listbox.selection_set(index)
        self.customer_listbox.activate(index)
        self.customer_listbox.see(index)
        return "break"

    def _on_customer_arrow_up(self, event=None):
        if not self.customer_suggestions_frame.winfo_ismapped() or self.customer_listbox.size() == 0:
            return "break"
        selection = self.customer_listbox.curselection()
        index = selection[0] - 1 if selection else 0
        if index < 0:
            index = 0
        self.customer_listbox.selection_clear(0, "end")
        self.customer_listbox.selection_set(index)
        self.customer_listbox.activate(index)
        self.customer_listbox.see(index)
        return "break"

    def _focus_customer_listbox(self, event=None):
        if self.customer_suggestions_frame.winfo_ismapped() and self.customer_listbox.size() > 0:
            selection = self.customer_listbox.curselection()
            index = selection[0] if selection else 0
            self.customer_listbox.selection_clear(0, "end")
            self.customer_listbox.selection_set(index)
            self.customer_listbox.activate(index)
            self.customer_listbox.see(index)
            return "break"
        return None

    def _on_customer_entry_confirm(self, event=None):
        typed = (self.customer_var.get() or "").strip()
        if self.customer_suggestions_frame.winfo_ismapped() and self.customer_listbox.size() > 0:
            selection = self.customer_listbox.curselection()
            if selection:
                label = self.customer_listbox.get(selection[0])
            else:
                label = self.customer_listbox.get(0)
            self.customer_var.set(label)
        elif typed:
            matches = self._matching_customer_labels(typed)
            if len(matches) == 1:
                self.customer_var.set(matches[0])
        self._hide_customer_suggestions()
        self._schedule_auto_filter()
        return "break"

    def _on_customer_listbox_select(self, event=None):
        # Navegação com seta para cima/baixo não deve aplicar filtro
        # nem alterar o conteúdo do campo. A confirmação ocorre apenas
        # com Enter ou clique do mouse.
        return None

    def _on_customer_listbox_confirm(self, event=None):
        selection = self.customer_listbox.curselection()
        if selection:
            label = self.customer_listbox.get(selection[0])
            self.customer_var.set(label)
            self._hide_customer_suggestions()
            self.customer_entry.focus_set()
            self.customer_entry.icursor("end")
            self._schedule_auto_filter()
        return "break"


    def _selected_account_code(self):
        typed = (self.account_var.get() or "").strip()
        if not typed or typed == "Todas":
            return None

        exact = self.all_account_options_map.get(typed)
        if exact is not None:
            return exact

        typed_lower = typed.lower()
        matches = [
            account_code
            for label, account_code in self.all_account_options_map.items()
            if label != "Todas" and typed_lower in label.lower()
        ]
        if len(matches) == 1:
            return matches[0]
        return None

    def _load_account_options(self):
        current = (self.account_var.get() or "").strip() or "Todas"
        options_map = {"Todas": None}
        try:
            rows = Database(self.config_data).list_open_invoice_accounts()
            for row in rows:
                label = f"{row.get('conta')} - {row.get('conta_nome')}"
                options_map[label] = row.get("conta")
        except Exception:
            pass

        self.all_account_options_map = options_map
        self.account_options_map = dict(options_map)
        self.account_var.set(current if current in options_map else "Todas")
        self._hide_account_suggestions()

    def _matching_account_labels(self, typed: str):
        typed = (typed or "").strip().lower()
        if not typed:
            return [label for label in self.all_account_options_map.keys()]
        return [
            label
            for label in self.all_account_options_map.keys()
            if label != "Todas" and typed in label.lower()
        ]

    def _show_account_suggestions(self, labels):
        self.account_listbox.delete(0, "end")
        for label in labels:
            self.account_listbox.insert("end", label)

        if labels:
            self.account_suggestions_frame.grid()
            self.account_listbox.selection_clear(0, "end")
            self.account_listbox.selection_set(0)
        else:
            self._hide_account_suggestions()

    def _hide_account_suggestions(self):
        if hasattr(self, "account_suggestions_frame"):
            self.account_suggestions_frame.grid_remove()

    def _apply_account_search(self, typed: str):
        labels = self._matching_account_labels(typed)
        filtered = {label: self.all_account_options_map.get(label) for label in labels}
        self.account_options_map = filtered
        if (typed or "").strip():
            self._show_account_suggestions(labels)
        else:
            self._hide_account_suggestions()

    def _on_account_keyrelease(self, event=None):
        keysym = getattr(event, "keysym", "")
        if keysym in ("Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return None
        self._apply_account_search(self.account_var.get())
        return None

    def _on_account_focus_in(self, event=None):
        typed = (self.account_var.get() or "").strip()
        if typed and typed != "Todas":
            self._apply_account_search(typed)
        return None

    def _on_account_focus_out(self, event=None):
        self.after(150, self._hide_account_suggestions)
        return None

    def _on_account_arrow_down(self, event=None):
        if not self.account_suggestions_frame.winfo_ismapped() or self.account_listbox.size() == 0:
            return "break"
        selection = self.account_listbox.curselection()
        index = selection[0] + 1 if selection else 0
        if index >= self.account_listbox.size():
            index = self.account_listbox.size() - 1
        self.account_listbox.selection_clear(0, "end")
        self.account_listbox.selection_set(index)
        self.account_listbox.activate(index)
        self.account_listbox.see(index)
        return "break"

    def _on_account_arrow_up(self, event=None):
        if not self.account_suggestions_frame.winfo_ismapped() or self.account_listbox.size() == 0:
            return "break"
        selection = self.account_listbox.curselection()
        index = selection[0] - 1 if selection else 0
        if index < 0:
            index = 0
        self.account_listbox.selection_clear(0, "end")
        self.account_listbox.selection_set(index)
        self.account_listbox.activate(index)
        self.account_listbox.see(index)
        return "break"

    def _on_account_entry_confirm(self, event=None):
        typed = (self.account_var.get() or "").strip()
        if self.account_suggestions_frame.winfo_ismapped() and self.account_listbox.size() > 0:
            selection = self.account_listbox.curselection()
            if selection:
                label = self.account_listbox.get(selection[0])
            else:
                label = self.account_listbox.get(0)
            self.account_var.set(label)
        elif typed:
            matches = self._matching_account_labels(typed)
            if len(matches) == 1:
                self.account_var.set(matches[0])
        self._hide_account_suggestions()
        self._schedule_auto_filter()
        return "break"

    def _on_account_listbox_select(self, event=None):
        return None

    def _on_account_listbox_confirm(self, event=None):
        selection = self.account_listbox.curselection()
        if selection:
            label = self.account_listbox.get(selection[0])
            self.account_var.set(label)
            self._hide_account_suggestions()
            self.account_entry.focus_set()
            self.account_entry.icursor("end")
            self._schedule_auto_filter()
        return "break"

    def _sort_value(self, row: InvoiceRow, column: str):
        if column == "company":
            return (row.company or "").lower()
        if column == "account_display":
            return (f"{row.account_code or ''} - {row.account_name or ''}").lower().strip(" -")
        if column == "code":
            try:
                return (0, int(str(row.customer_code)))
            except Exception:
                return (1, str(row.customer_code or "").lower())
        if column == "name":
            return (row.customer_name or "").lower()
        if column == "issue_date":
            return row.issue_date or datetime.min
        if column == "due_date":
            return row.due_date or datetime.min
        if column == "amount":
            return float(row.amount or 0)
        if column == "discount":
            return float(row.discount_amount or 0)
        if column == "open_balance":
            return float(row.open_balance or 0)
        return ""
    def sort_by(self, column: str):
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False
        self.rows.sort(key=lambda r: self._sort_value(r, column), reverse=self.sort_reverse)
        self._refresh_tree()
        self._update_heading_titles()
    def _update_heading_titles(self):
        labels = {
            "company": "Empresa",
            "account_display": "Conta",
            "code": "Código",
            "name": "Cliente",
            "issue_date": "Data",
            "due_date": "Vencimento",
            "amount": "Valor",
            "discount": "Desconto",
            "open_balance": "Saldo em aberto",
        }
        for col, label in labels.items():
            suffix = ""
            if self.sort_column == col:
                suffix = " ▼" if self.sort_reverse else " ▲"
            self.tree.heading(col, text=label + suffix, command=lambda c=col: self.sort_by(c))
    def _group_rows(self, rows: List[InvoiceRow]) -> List[InvoiceRow]:
        mode = self.GROUP_OPTIONS.get(self.group_by_var.get(), "none")
        if mode == "none":
            return list(rows)
        grouped = {}
        for row in rows:
            if mode == "customer":
                key = (row.company, row.customer_id, row.customer_code, row.customer_name, row.account_code, row.account_name)
            elif mode == "due_date":
                key = (row.company, row.due_date)
            elif mode == "account_group":
                key = (row.company, row.account_code, row.account_name)
            else:
                key = (row.invoice_id,)
            if key not in grouped:
                if mode == "customer":
                    grouped[key] = InvoiceRow(
                        invoice_id=f"grp_customer_{row.company}_{row.customer_id}_{row.account_code}",
                        company=row.company,
                        customer_id=row.customer_id,
                        customer_code=row.customer_code,
                        customer_name=row.customer_name,
                        motive_code="",
                        motive_name="",
                        account_code=row.account_code,
                        account_name=row.account_name,
                        issue_date=None,
                        due_date=None,
                        amount=0,
                        discount_amount=0,
                        paid_amount=0,
                        open_balance=0,
                    )
                elif mode == "due_date":
                    grouped[key] = InvoiceRow(
                        invoice_id=f"grp_due_{row.company}_{row.due_date}",
                        company=row.company,
                        customer_id="",
                        customer_code="",
                        customer_name="Agrupado por vencimento",
                        motive_code="",
                        motive_name="",
                        account_code=row.account_code,
                        account_name=row.account_name,
                        issue_date=None,
                        due_date=row.due_date,
                        amount=0,
                        discount_amount=0,
                        paid_amount=0,
                        open_balance=0,
                    )
                elif mode == "account_group":
                    grouped[key] = InvoiceRow(
                        invoice_id=f"grp_account_{row.company}_{row.account_code}",
                        company=row.company,
                        customer_id="",
                        customer_code="",
                        customer_name="Agrupado por conta",
                        motive_code="",
                        motive_name="",
                        account_code=row.account_code,
                        account_name=row.account_name,
                        issue_date=None,
                        due_date=None,
                        amount=0,
                        discount_amount=0,
                        paid_amount=0,
                        open_balance=0,
                    )
            grouped[key].amount = float(grouped[key].amount or 0) + float(row.amount or 0)
            grouped[key].discount_amount = float(grouped[key].discount_amount or 0) + float(row.discount_amount or 0)
            grouped[key].paid_amount = float(grouped[key].paid_amount or 0) + float(row.paid_amount or 0)
            grouped[key].open_balance = float(grouped[key].open_balance or 0) + float(row.open_balance or 0)
        return list(grouped.values())
    def load_data(self):
        try:
            date_from = self._parse_date(self.period_start_var.get())
            date_to = self._parse_date(self.period_end_var.get())
            if date_from and date_to and date_from > date_to:
                raise AppError("O período inicial não pode ser maior que o período final.")
            self.set_status("Conectando ao banco e carregando faturas a receber.")
            data = Database(self.config_data).list_open_invoices(
                due_date_from=date_from,
                due_date_to=date_to,
                customer_id=self._selected_customer_id(),
                account_code=self._selected_account_code(),
            )
            self.raw_rows = [
                InvoiceRow(
                    invoice_id=row.get("movto_id"),
                    company=row.get("empresa") or "",
                    customer_id=row.get("customer_id"),
                    customer_code=row.get("codigo_cliente"),
                    customer_name=row.get("cliente") or "",
                    motive_code=row.get("motivo"),
                    motive_name=f"Motivo {row.get('motivo')}" if row.get("motivo") not in (None, "") else "",
                    account_code=row.get("conta") or "",
                    account_name=row.get("conta_nome") or "",
                    issue_date=row.get("data"),
                    due_date=row.get("vencto"),
                    amount=row.get("valor"),
                    discount_amount=row.get("valor_desconto"),
                    paid_amount=row.get("valor_baixado"),
                    open_balance=row.get("saldo_em_aberto"),
                    customer_email=row.get("customer_email") or "",
                )
                for row in data
            ]
            self.rows = self._group_rows(self.raw_rows)
            if self.sort_column:
                self.rows.sort(key=lambda r: self._sort_value(r, self.sort_column), reverse=self.sort_reverse)
            self._refresh_tree()
            self._update_heading_titles()
            total_open = sum(float(r.open_balance or 0) for r in self.rows)
            self.set_status(f"{len(self.rows)} registro(s). Total em aberto: {total_open:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            AuditLogger.write(self.current_user, "carregar_lista", f"tipo=faturas_receber;quantidade={len(self.rows)};agrupar={self.GROUP_OPTIONS.get(self.group_by_var.get(), 'none')}")
        except Exception as e:
            self.set_status("Falha ao carregar dados.")
            messagebox.showerror(APP_TITLE, f"Erro ao carregar faturas a receber:\n\n{e}", parent=self)
    def _selected_invoice_row(self) -> Optional[InvoiceRow]:
        selected = self.tree.selection()
        if not selected:
            return None
        return self.tree_items.get(selected[0])

    def open_email_window(self):
        row = self._selected_invoice_row()
        if not row:
            messagebox.showwarning(APP_TITLE, "Selecione uma fatura na grade.", parent=self)
            return
        if self.GROUP_OPTIONS.get(self.group_by_var.get()) != "none" or str(row.invoice_id).startswith("grp_"):
            messagebox.showwarning(APP_TITLE, "Para enviar e-mail, selecione uma linha detalhada sem agrupamento.", parent=self)
            return
        if not row.customer_id:
            messagebox.showwarning(APP_TITLE, "Não foi possível identificar o cliente da linha selecionada.", parent=self)
            return
        try:
            email = row.customer_email or Database(self.config_data).get_customer_email(row.customer_id)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Erro ao buscar o e-mail do cliente:\n\n{e}", parent=self)
            return
        EmailComposeWindow(self, self.config_data, self.current_user, row, email or "")

    def _row_values(self, row: InvoiceRow):
        return (
            row.company,
            (f"{row.account_code or ''} - {row.account_name or ''}").strip(" -"),
            row.customer_code,
            row.customer_name,
            row.issue_date_display(),
            row.due_date_display(),
            row.amount_display(),
            row.discount_amount_display(),
            row.open_balance_display(),
        )
    def _refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree_items.clear()
        for row in self.rows:
            item_id = self.tree.insert("", "end", values=self._row_values(row))
            self.tree_items[item_id] = row
class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("460x300")
        self.minsize(460, 300)
        self.config_data = ConfigManager.load()
        self.first_time = not CONFIG_PATH.exists()
        self.current_user = ""
        self.current_user_is_master = False
        self.license_data = None
        self.user_var = tk.StringVar(value="")
        self.home_message_var = tk.StringVar(value="")
        self.login_message_var = tk.StringVar(value="")
        self.inactive_window: Optional[InactiveCustomersWindow] = None
        self.invoices_window: Optional[OpenInvoicesWindow] = None
        self._setup_style()
        self._build_menu()
        self._build_frames()
        self._start_application_flow()
    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        base_font = ("Segoe UI", 10)
        self.option_add("*Font", base_font)
        style.configure("TFrame", background="#f5f7fb")
        style.configure("TLabel", background="#f5f7fb", foreground="#1f2937", font=base_font)
        style.configure("TButton", font=("Segoe UI", 9, "bold"), padding=(10, 6))
        self.configure(background="#f5f7fb")
    def _build_menu(self):
        menubar = tk.Menu(self)
        self.login_menu = tk.Menu(menubar, tearoff=0)
        self.login_menu.add_command(label="Cadastrar usuário", command=self.open_create_user)
        self.login_menu.add_command(label="Alterar senha", command=self.open_change_own_password)
        self.login_menu.add_command(label="Alterar usuários", command=self.open_master_manage_users)
        menubar.add_cascade(label="Login", menu=self.login_menu)
        self.cadastro_menu = tk.Menu(menubar, tearoff=0)
        self.cadastro_menu.add_command(label="Clientes inativos", command=self.open_inactive_customers_screen)
        menubar.add_cascade(label="Cadastro", menu=self.cadastro_menu)
        self.financeiro_menu = tk.Menu(menubar, tearoff=0)
        self.financeiro_menu.add_command(label="Faturas a receber", command=self.open_invoices_screen)
        menubar.add_cascade(label="Financeiro", menu=self.financeiro_menu)
        self.config_menu = tk.Menu(menubar, tearoff=0)
        self.config_menu.add_command(label="Configuração local", command=self.open_config)
        menubar.add_cascade(label="Configurações", menu=self.config_menu)
        self.config(menu=menubar)
        self._update_menu_visibility()
    def _update_menu_visibility(self):
        logged_state = "normal" if self.current_user else "disabled"
        master_state = "normal" if self.current_user_is_master else "disabled"
        self.login_menu.entryconfig("Cadastrar usuário", state=logged_state)
        self.login_menu.entryconfig("Alterar senha", state=logged_state)
        self.login_menu.entryconfig("Alterar usuários", state=master_state)
        self.cadastro_menu.entryconfig("Clientes inativos", state=logged_state)
        self.config_menu.entryconfig("Configuração local", state=logged_state)
        self.financeiro_menu.entryconfig("Faturas a receber", state=logged_state)
    def _build_frames(self):
        self.login_frame = ttk.Frame(self, padding=24)
        self.setup_user_frame = ttk.Frame(self, padding=24)
        self.home_frame = ttk.Frame(self, padding=0)
        self.license_block_frame = ttk.Frame(self, padding=24)
        self._build_login_frame()
        self._build_setup_user_frame()
        self._build_home_frame()
        self._build_license_block_frame()
    def _clear_frames(self):
        for frame in (self.login_frame, self.setup_user_frame, self.home_frame, self.license_block_frame):
            frame.pack_forget()
    def _start_application_flow(self):
        try:
            self.license_data = LicenseManager.validate_file()
        except Exception as e:
            self._show_license_block(str(e))
            return
        self._show_login_frame()
    def _build_license_block_frame(self):
        card = ttk.Frame(self.license_block_frame, padding=20)
        card.place(relx=0.5, rely=0.5, anchor="center")
        ttk.Label(card, text="Licença inválida", font=("Segoe UI", 12, "bold")).pack(anchor="center", pady=(0, 10))
        self.license_error_var = tk.StringVar(value="")
        ttk.Label(card, textvariable=self.license_error_var, wraplength=380, justify="center").pack(anchor="center", pady=(0, 14))
        ttk.Label(card, text=f"Arquivo esperado: {LICENSE_FILENAME}", justify="center").pack(anchor="center", pady=(0, 10))
        ttk.Button(card, text="Fechar", command=self.destroy).pack()
    def _show_license_block(self, error_text: str):
        self.geometry("460x250")
        self.minsize(460, 250)
        self.license_error_var.set(error_text)
        self._clear_frames()
        self.license_block_frame.pack(fill="both", expand=True)
    def _show_login_frame(self):
        self._clear_frames()
        self.login_frame.pack(fill="both", expand=True)
        self.after(50, lambda: self.login_user_entry.focus_set())
    def _show_setup_user_frame(self):
        self._clear_frames()
        self.setup_user_frame.pack(fill="both", expand=True)
    def show_home(self):
        self._clear_frames()
        self.home_frame.pack(fill="both", expand=True)
        self._update_home_message()
    def _build_login_frame(self):
        card = ttk.Frame(self.login_frame, padding=20)
        card.place(relx=0.5, rely=0.5, anchor="center")
        ttk.Label(card, text="Acesso ao sistema", font=("Segoe UI", 12, "bold")).pack(anchor="center", pady=(2, 6))
        ttk.Label(card, textvariable=self.login_message_var, justify="center", wraplength=320).pack(anchor="center", pady=(0, 14))
        form = ttk.Frame(card)
        form.pack(fill="x")
        ttk.Label(form, text="Usuário").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=6)
        self.login_user_entry = ttk.Entry(form, width=24)
        self.login_user_entry.grid(row=0, column=1, sticky="ew", pady=6)
        ttk.Label(form, text="Senha").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=6)
        self.login_pass_entry = ttk.Entry(form, width=24, show="*")
        self.login_pass_entry.grid(row=1, column=1, sticky="ew", pady=6)
        form.columnconfigure(1, weight=1)
        btns = ttk.Frame(card)
        btns.pack(fill="x", pady=(18, 0))
        ttk.Button(btns, text="Sair", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Entrar", command=self._handle_login).pack(side="right", padx=(0, 8))
        self.login_pass_entry.bind("<Return>", lambda e: self._handle_login())
    def _build_setup_user_frame(self):
        card = ttk.Frame(self.setup_user_frame, padding=28)
        card.place(relx=0.5, rely=0.5, anchor="center")
        ttk.Label(card, text="Criar usuário local", font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 8))
        ttk.Label(card, text="Você pode criar agora um usuário local para acessar o sistema. Se não quiser, clique em Continuar.", wraplength=500, justify="left").pack(anchor="w", pady=(0, 16))
        form = ttk.Frame(card)
        form.pack(fill="x")
        ttk.Label(form, text="Usuário").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.first_user_entry = ttk.Entry(form, width=34)
        self.first_user_entry.grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Label(form, text="Senha").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        self.first_pass_entry = ttk.Entry(form, width=34, show="*")
        self.first_pass_entry.grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Label(form, text="Confirmar senha").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)
        self.first_confirm_entry = ttk.Entry(form, width=34, show="*")
        self.first_confirm_entry.grid(row=2, column=1, sticky="ew", pady=5)
        form.columnconfigure(1, weight=1)
        btns = ttk.Frame(card)
        btns.pack(fill="x", pady=(20, 0))
        ttk.Button(btns, text="Continuar", command=self._continue_first_access).pack(side="right")
        ttk.Button(btns, text="Criar usuário", command=self._create_first_user).pack(side="right", padx=(0, 8))
    def _build_home_frame(self):
        body = ttk.Frame(self.home_frame, padding=24)
        body.pack(fill="both", expand=True)
        center = ttk.Frame(body)
        center.place(relx=0.5, rely=0.42, anchor="center")
        ttk.Label(center, text="DataHub", font=("Segoe UI", 18, "bold")).pack(anchor="center", pady=(0, 10))
        ttk.Label(center, text="Selecione uma opção no menu para abrir uma funcionalidade.", justify="center", wraplength=520).pack(anchor="center", pady=(0, 6))
        ttk.Label(center, textvariable=self.home_message_var, justify="center", wraplength=620).pack(anchor="center")
        ttk.Label(center, textvariable=self.user_var, justify="center").pack(anchor="center", pady=(16, 0))
    def _update_home_message(self):
        name = ""
        if isinstance(self.license_data, dict):
            name = str(self.license_data.get("customer_name", "")).strip()
        self.home_message_var.set(f"Licença ativa para: {name}" if name else "")
    def _handle_login(self):
        username = self.login_user_entry.get().strip()
        password = self.login_pass_entry.get()
        if not username:
            messagebox.showwarning(APP_TITLE, "Informe o usuário.", parent=self)
            return
        if not password:
            messagebox.showwarning(APP_TITLE, "Informe a senha.", parent=self)
            return
        if username == MASTER_USERNAME and password == MASTER_PASSWORD:
            AuditLogger.write(username, "login_sucesso", "tipo=master")
            self._after_login_success(username, True)
            return
        if self.first_time:
            AuditLogger.write(username or "-", "login_falha", "primeiro_acesso_sem_master")
            messagebox.showerror(APP_TITLE, "No primeiro acesso, apenas o login master é permitido.", parent=self)
            return
        if UserManager.validate_login(self.config_data, username, password):
            AuditLogger.write(username, "login_sucesso", "tipo=local")
            self._after_login_success(username, False)
            return
        AuditLogger.write(username or "-", "login_falha", "credenciais_invalidas")
        messagebox.showerror(APP_TITLE, "Usuário ou senha inválidos.", parent=self)
    def _after_login_success(self, username: str, is_master: bool):
        self.current_user = username
        self.current_user_is_master = is_master
        self.user_var.set(f"Usuário: {username}")
        self._update_menu_visibility()
        self.login_user_entry.delete(0, "end")
        self.login_pass_entry.delete(0, "end")
        if self.first_time:
            self._show_setup_user_frame()
            return
        try:
            LicenseManager.validate_against_database(self.config_data, self.license_data)
        except Exception as e:
            self._show_license_block(str(e))
            return
        self.geometry("900x520")
        self.minsize(700, 420)
        self.show_home()
    def _create_first_user(self):
        try:
            username = self.first_user_entry.get().strip()
            password = self.first_pass_entry.get()
            confirm = self.first_confirm_entry.get()
            if not username:
                raise AppError("Informe o usuário.")
            if not password:
                raise AppError("Informe a senha.")
            if password != confirm:
                raise AppError("A confirmação da senha não confere.")
            UserManager.add_user(self.config_data, username, password)
            ConfigManager.save(self.config_data)
            self.config_data = ConfigManager.load()
            AuditLogger.write(MASTER_USERNAME, "cadastrar_usuario", f"alvo={username};primeiro_acesso=sim")
            self.first_user_entry.delete(0, "end")
            self.first_pass_entry.delete(0, "end")
            self.first_confirm_entry.delete(0, "end")
            messagebox.showinfo(APP_TITLE, "Usuário local criado com sucesso.", parent=self)
            self._continue_first_access()
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e), parent=self)
    def _continue_first_access(self):
        ConfigWindow(self, self.config_data, self._after_first_connection_saved)
    def _after_first_connection_saved(self, cfg: Dict[str, Any]):
        try:
            LicenseManager.validate_against_database(cfg, self.license_data)
        except Exception as e:
            self._show_license_block(str(e))
            return
        self.config_data = cfg
        self.first_time = False
        self._update_menu_visibility()
        self.geometry("900x520")
        self.minsize(700, 420)
        self.show_home()
    def open_create_user(self):
        if not self.current_user:
            return
        CreateUserWindow(self, self.config_data, self.current_user, self._apply_security_config)
    def open_change_own_password(self):
        if not self.current_user:
            return
        if self.current_user_is_master:
            messagebox.showwarning(APP_TITLE, "A senha do usuário master é fixa no sistema e não pode ser alterada.", parent=self)
            return
        ChangeOwnPasswordWindow(self, self.config_data, self.current_user, self._apply_security_config)
    def open_master_manage_users(self):
        if not self.current_user_is_master:
            messagebox.showwarning(APP_TITLE, "Somente o usuário master pode acessar esta tela.", parent=self)
            return
        MasterManageUsersWindow(self, self.config_data, self.current_user, self._apply_security_config)

    def open_config(self):
        if not self.current_user:
            return
        ConfigWindow(self, self.config_data, self._apply_new_config)

    def open_inactive_customers_screen(self):
        if not self.current_user:
            return
        if self.inactive_window is not None and self.inactive_window.winfo_exists():
            self.inactive_window.deiconify()
            self.inactive_window.lift()
            self.inactive_window.focus_force()
            return
        self.inactive_window = InactiveCustomersWindow(self, self.config_data, self.current_user, self._apply_new_config)
    def open_invoices_screen(self):
        if not self.current_user:
            return
        if self.invoices_window is not None and self.invoices_window.winfo_exists():
            self.invoices_window.deiconify()
            self.invoices_window.lift()
            self.invoices_window.focus_force()
            return
        self.invoices_window = OpenInvoicesWindow(self, self.config_data, self.current_user, self._apply_new_config)
    def _apply_security_config(self, cfg: Dict[str, Any]):
        self.config_data = ConfigManager.load()
    def _apply_new_config(self, cfg: Dict[str, Any]):
        try:
            LicenseManager.validate_against_database(cfg, self.license_data)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e), parent=self)
            return
        self.config_data = cfg
def main():
    app = MainApp()
    app.mainloop()
if __name__ == "__main__":
    main()
