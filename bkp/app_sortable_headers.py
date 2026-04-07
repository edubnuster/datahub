# -*- coding: utf-8 -*-
import json
import sys
from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import psycopg2
import psycopg2.extras
import tkinter as tk
from tkinter import ttk, messagebox

APP_TITLE = "Clientes sem movimentação"
CONFIG_FILENAME = "config.json"

DEFAULT_LIST_SQL = """
select
    c.grid as customer_id,
    pessoa_nome_f(l.empresa) as last_purchase_company,
    c.codigo as customer_code,
    c.nome as customer_name,
    coalesce(co.nome, 'Sem conta') as account_name,
    coalesce(pc.lim_credito, 0) as credit_limit,
    max(l.data) as last_purchase_date,
    case c.flag
        when 'A' then 'Ativo'
        when 'I' then 'Inativo'
        when 'D' then 'Deletado'
        else coalesce(c.flag, '')
    end as customer_status
from cliente c
join lancto l
    on l.pessoa = c.grid
left join pessoa_conta pc
    on pc.pessoa = c.grid
left join conta co
    on co.codigo = pc.conta
where l.operacao = 'V'
group by
    c.grid,
    l.empresa,
    c.codigo,
    c.nome,
    coalesce(co.nome, 'Sem conta'),
    coalesce(pc.lim_credito, 0),
    c.flag,
    pessoa_nome_f(l.empresa)
having max(l.data) < current_date - interval '3 months'
order by pessoa_nome_f(l.empresa), max(l.data)
"""

DEFAULT_CONFIG = {
    "connection": {
        "host": "127.0.0.1",
        "port": 5432,
        "dbname": "",
        "user": "postgres",
        "password": "",
        "client_encoding": "LATIN1",
    },
    "queries": {
        "list_inactive_customers_sql": DEFAULT_LIST_SQL,
        "delete_customer_sql": """
            update cliente
               set flag = 'D'
             where grid = %(customer_id)s
        """,
        "inactivate_customer_sql": """
            update cliente
               set flag = 'I'
             where grid = %(customer_id)s
        """,
        "disable_credit_sql": """
            update pessoa_conta pc
               set lim_credito = 0
             where pc.pessoa = %(customer_id)s
        """
    }
}


class AppError(Exception):
    pass


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


CONFIG_PATH = app_dir() / CONFIG_FILENAME


class ConfigManager:
    @staticmethod
    def exists() -> bool:
        return CONFIG_PATH.exists()

    @staticmethod
    def load() -> Dict[str, Any]:
        if not CONFIG_PATH.exists():
            return deepcopy(DEFAULT_CONFIG)

        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)

        merged = deepcopy(DEFAULT_CONFIG)
        ConfigManager._deep_update(merged, data)

        current_sql = merged.get("queries", {}).get("list_inactive_customers_sql", "")
        if "join pessoa_conta pc" in current_sql and "left join pessoa_conta pc" not in current_sql:
            merged["queries"]["list_inactive_customers_sql"] = DEFAULT_LIST_SQL
            current_sql = merged["queries"]["list_inactive_customers_sql"]

        if "credit_limit" not in current_sql:
            merged["queries"]["list_inactive_customers_sql"] = DEFAULT_LIST_SQL

        if not merged.get("queries", {}).get("list_inactive_customers_sql", "").strip():
            merged["queries"]["list_inactive_customers_sql"] = DEFAULT_LIST_SQL

        if not merged.get("queries", {}).get("delete_customer_sql", "").strip():
            merged["queries"]["delete_customer_sql"] = DEFAULT_CONFIG["queries"]["delete_customer_sql"]

        if not merged.get("queries", {}).get("inactivate_customer_sql", "").strip():
            merged["queries"]["inactivate_customer_sql"] = DEFAULT_CONFIG["queries"]["inactivate_customer_sql"]

        disable_sql = merged.get("queries", {}).get("disable_credit_sql", "") or ""
        if not disable_sql.strip() or "update conta" in disable_sql.lower():
            merged["queries"]["disable_credit_sql"] = DEFAULT_CONFIG["queries"]["disable_credit_sql"]

        return merged

    @staticmethod
    def save(data: Dict[str, Any]) -> None:
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

    @staticmethod
    def _deep_update(base: Dict[str, Any], incoming: Dict[str, Any]) -> None:
        for key, value in incoming.items():
            if isinstance(value, dict) and isinstance(base.get(key), dict):
                ConfigManager._deep_update(base[key], value)
            else:
                base[key] = value


class Database:
    def __init__(self, config: Dict[str, Any]):
        self.config = config

    def _connect(self):
        conn_cfg = self.config["connection"]
        conn = psycopg2.connect(
            host=conn_cfg["host"],
            port=conn_cfg["port"],
            dbname=conn_cfg["dbname"],
            user=conn_cfg["user"],
            password=conn_cfg["password"],
            connect_timeout=8
        )
        client_encoding = (conn_cfg.get("client_encoding") or "").strip()
        if client_encoding:
            conn.set_client_encoding(client_encoding)
        return conn

    def test_connection(self) -> None:
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute("select 1")
                cur.fetchone()

    def list_inactive_customers(self) -> List[Dict[str, Any]]:
        sql_text = self.config["queries"].get("list_inactive_customers_sql", "").strip()
        if not sql_text:
            raise AppError("A query de listagem ainda não foi configurada no config.json.")

        with self._connect() as conn:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(sql_text)
                rows = cur.fetchall()
                return [dict(r) for r in rows]

    def execute_action(self, sql_text: str, customer_ids: List[Any]) -> int:
        if not sql_text.strip():
            raise AppError("SQL da ação não configurado.")

        total = 0
        with self._connect() as conn:
            try:
                with conn.cursor() as cur:
                    for customer_id in customer_ids:
                        cur.execute(sql_text, {"customer_id": customer_id})
                        if cur.rowcount is not None and cur.rowcount > 0:
                            total += cur.rowcount
                conn.commit()
            except Exception:
                conn.rollback()
                raise
        return total


@dataclass
class CustomerRow:
    customer_id: Any
    customer_code: Any
    customer_name: str
    last_purchase_date: Optional[Any]
    last_purchase_company: Optional[str]
    account_name: Optional[str]
    credit_limit: Any
    customer_status: str
    selected: bool = False

    def last_purchase_date_display(self) -> str:
        value = self.last_purchase_date
        if value in (None, ""):
            return "Sem compra"
        if isinstance(value, datetime):
            return value.strftime("%d/%m/%Y %H:%M")
        try:
            return value.strftime("%d/%m/%Y")
        except Exception:
            return str(value)

    def checkbox(self) -> str:
        return "☑" if self.selected else "☐"

    def credit_limit_display(self) -> str:
        value = self.credit_limit
        if value in (None, ""):
            return "0,00"
        try:
            num = float(value)
            return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return str(value)


class ConfigWindow(tk.Toplevel):
    def __init__(self, master, current_config: Dict[str, Any], on_save):
        super().__init__(master)
        self.title("Configuração de conexão")
        self.resizable(False, False)
        self.geometry("520x390")
        self.current_config = deepcopy(current_config)
        self.on_save = on_save
        self.entries: Dict[str, tk.Widget] = {}
        self._build_ui()
        self.transient(master)
        self.grab_set()

    def _build_ui(self):
        wrapper = ttk.Frame(self, padding=14)
        wrapper.pack(fill="both", expand=True)

        ttk.Label(
            wrapper,
            text="Configuração do banco de dados",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", pady=(0, 6))

        ttk.Label(
            wrapper,
            text="Informe apenas os dados de conexão. As queries serão configuradas depois.",
            wraplength=470,
            justify="left",
        ).pack(anchor="w", pady=(0, 12))

        form = ttk.Frame(wrapper)
        form.pack(fill="x")

        fields = [
            ("Host", "connection.host"),
            ("Porta", "connection.port"),
            ("Banco", "connection.dbname"),
            ("Usuário", "connection.user"),
            ("Senha", "connection.password"),
            ("Encoding", "connection.client_encoding"),
        ]

        for i, (label, key) in enumerate(fields):
            ttk.Label(form, text=label).grid(row=i, column=0, sticky="w", padx=(0, 8), pady=5)
            entry = ttk.Entry(form, width=42, show="*" if key.endswith("password") else None)
            entry.grid(row=i, column=1, sticky="ew", pady=5)
            entry.insert(0, str(self._get_value(key)))
            self.entries[key] = entry

        form.columnconfigure(1, weight=1)

        btns = ttk.Frame(wrapper)
        btns.pack(fill="x", pady=(20, 0))

        ttk.Button(btns, text="Testar conexão", command=self._test_connection).pack(side="left")
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Salvar conexão", command=self._save).pack(side="right", padx=(0, 8))

        ttk.Label(
            wrapper,
            text=f"Arquivo de configuração: {CONFIG_PATH}",
            wraplength=470,
            justify="left"
        ).pack(anchor="w", pady=(18, 0))

    def _get_value(self, dotted_key: str):
        ref = self.current_config
        for part in dotted_key.split("."):
            ref = ref[part]
        return ref

    def _set_value(self, data: Dict[str, Any], dotted_key: str, value: Any):
        parts = dotted_key.split(".")
        ref = data
        for part in parts[:-1]:
            ref = ref[part]
        ref[parts[-1]] = value

    def _collect_config(self) -> Dict[str, Any]:
        data = deepcopy(self.current_config)

        for key, entry in self.entries.items():
            value = entry.get().strip()

            if key == "connection.port":
                if not value:
                    raise AppError("Informe a porta do banco.")
                try:
                    value = int(value)
                except ValueError:
                    raise AppError("A porta deve ser numérica.")

            self._set_value(data, key, value)

        if not data["connection"]["host"]:
            raise AppError("Informe o host.")
        if not data["connection"]["dbname"]:
            raise AppError("Informe o nome do banco.")
        if not data["connection"]["user"]:
            raise AppError("Informe o usuário.")

        return data

    def _test_connection(self):
        try:
            cfg = self._collect_config()
            Database(cfg).test_connection()
            messagebox.showinfo(APP_TITLE, "Conexão realizada com sucesso.")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Falha ao testar conexão:\n\n{e}")

    def _save(self):
        try:
            cfg = self._collect_config()
            ConfigManager.save(cfg)
            self.on_save(cfg)
            messagebox.showinfo(APP_TITLE, f"Configuração salva em:\n{CONFIG_PATH}")
            self.destroy()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Erro ao salvar configuração:\n\n{e}")


class MainApp(tk.Tk):
    FILTER_OPTIONS = {
        "Todos": None,
        "Ativos": "Ativo",
        "Inativos": "Inativo",
        "Deletados": "Deletado",
    }

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1320x720")
        self.minsize(1240, 660)

        self.config_data = ConfigManager.load()
        self.rows: List[CustomerRow] = []
        self.filtered_rows: List[CustomerRow] = []
        self.tree_items: Dict[str, CustomerRow] = {}

        self.filter_var = tk.StringVar(value="Todos")
        self.sort_column = None
        self.sort_reverse = False

        self._setup_style()
        self._build_ui()
        self.after(150, self._first_run_check)

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
        style.configure("TCombobox", padding=4)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("Treeview", rowheight=24, font=("Segoe UI", 10), background="#ffffff", fieldbackground="#ffffff")
        self.configure(background="#f5f7fb")

    def _build_ui(self):
        top = ttk.Frame(self, padding=(12, 12, 12, 8))
        top.pack(fill="x")

        left_actions = ttk.Frame(top)
        left_actions.pack(side="left")

        ttk.Button(left_actions, text="Atualizar lista", command=self.load_data).pack(side="left")
        ttk.Button(left_actions, text="Marcar todos", command=self.mark_all).pack(side="left", padx=(8, 0))
        ttk.Button(left_actions, text="Desmarcar todos", command=self.unmark_all).pack(side="left", padx=(8, 0))
        ttk.Button(left_actions, text="Configuração", command=self.open_config).pack(side="left", padx=(16, 0))

        filter_box = ttk.Frame(top)
        filter_box.pack(side="left", padx=(18, 0))
        ttk.Label(filter_box, text="Mostrar:").pack(side="left", padx=(0, 6))
        filtro = ttk.Combobox(
            filter_box,
            textvariable=self.filter_var,
            values=list(self.FILTER_OPTIONS.keys()),
            state="readonly",
            width=12
        )
        filtro.pack(side="left")
        filtro.bind("<<ComboboxSelected>>", lambda e: self.apply_filter())

        right_actions = ttk.Frame(top)
        right_actions.pack(side="right")

        ttk.Button(
            right_actions,
            text="Excluir cliente",
            command=lambda: self.run_action("delete_customer_sql", "Excluir cliente", "D", "Deletado")
        ).pack(side="right")

        ttk.Button(
            right_actions,
            text="Inativar cliente",
            command=lambda: self.run_action("inactivate_customer_sql", "Inativar cliente", "I", "Inativo")
        ).pack(side="right", padx=(0, 8))

        ttk.Button(
            right_actions,
            text="Inativar vendas a prazo",
            command=lambda: self.run_action("disable_credit_sql", "Inativar vendas a prazo", None, None)
        ).pack(side="right", padx=(0, 8))

        middle = ttk.Frame(self, padding=(12, 0, 12, 0))
        middle.pack(fill="both", expand=True)

        columns = ("selected", "company", "code", "name", "account", "credit_limit", "last_date", "status")
        self.tree = ttk.Treeview(middle, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("selected", text="Selecionado", command=lambda: self.sort_by("selected"))
        self.tree.heading("company", text="Empresa da última compra", command=lambda: self.sort_by("company"))
        self.tree.heading("code", text="Código", command=lambda: self.sort_by("code"))
        self.tree.heading("name", text="Cliente", command=lambda: self.sort_by("name"))
        self.tree.heading("account", text="Conta", command=lambda: self.sort_by("account"))
        self.tree.heading("credit_limit", text="Lim. crédito", command=lambda: self.sort_by("credit_limit"))
        self.tree.heading("last_date", text="Última compra", command=lambda: self.sort_by("last_date"))
        self.tree.heading("status", text="Status", command=lambda: self.sort_by("status"))

        self.tree.column("selected", width=88, minwidth=88, anchor="center", stretch=False)
        self.tree.column("company", width=180, minwidth=165, anchor="w", stretch=False)
        self.tree.column("code", width=76, minwidth=72, anchor="center", stretch=False)
        self.tree.column("name", width=275, minwidth=235, anchor="w", stretch=True)
        self.tree.column("account", width=165, minwidth=145, anchor="w", stretch=False)
        self.tree.column("credit_limit", width=105, minwidth=95, anchor="e", stretch=False)
        self.tree.column("last_date", width=106, minwidth=98, anchor="center", stretch=False)
        self.tree.column("status", width=92, minwidth=88, anchor="center", stretch=False)

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<Double-1>", self._toggle_selected_from_click)
        self.tree.bind("<Button-1>", self._handle_single_click_checkbox)

        yscroll = ttk.Scrollbar(middle, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")

        xscroll = ttk.Scrollbar(middle, orient="horizontal", command=self.tree.xview)
        xscroll.grid(row=1, column=0, sticky="ew")

        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        middle.rowconfigure(0, weight=1)
        middle.columnconfigure(0, weight=1)

        bottom = ttk.Frame(self, padding=(12, 8, 12, 10))
        bottom.pack(fill="x")

        self.status_var = tk.StringVar(value="Pronto.")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left")

        note = (
            "Dica: clique no checkbox da primeira coluna ou dê duplo clique na linha para marcar ou desmarcar. "
            "O arquivo config.json será salvo na mesma pasta do executável."
        )
        ttk.Label(bottom, text=note).pack(side="right")

    def _sort_value(self, row: CustomerRow, column: str):
        if column == "selected":
            return 1 if row.selected else 0
        if column == "company":
            return (row.last_purchase_company or "").lower()
        if column == "code":
            value = row.customer_code
            try:
                return (0, int(str(value)))
            except Exception:
                return (1, str(value or "").lower())
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
            try:
                return datetime.combine(value, datetime.min.time())
            except Exception:
                return str(value or "")
        if column == "status":
            ordem = {"Ativo": 0, "Inativo": 1, "Deletado": 2}
            return ordem.get(row.customer_status, 99), row.customer_status
        return ""

    def sort_by(self, column: str):
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False

        self.filtered_rows.sort(
            key=lambda row: self._sort_value(row, column),
            reverse=self.sort_reverse
        )
        self._refresh_tree()
        self._update_heading_titles()

    def _update_heading_titles(self):
        labels = {
            "selected": "Selecionado",
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

    def _first_run_check(self):
        if not ConfigManager.exists():
            messagebox.showinfo(
                APP_TITLE,
                "Primeiro acesso detectado. Informe apenas os dados de conexão com o banco para continuar."
            )
            self.open_config()
        else:
            self.load_data()

    def open_config(self):
        ConfigWindow(self, self.config_data, self._apply_new_config)

    def _apply_new_config(self, cfg: Dict[str, Any]):
        self.config_data = cfg
        try:
            self.load_data()
        except Exception:
            pass

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

    def load_data(self):
        try:
            self.set_status("Conectando ao banco e carregando clientes.")
            db = Database(self.config_data)
            data = db.list_inactive_customers()

            self.rows = [
                CustomerRow(
                    customer_id=row.get("customer_id"),
                    customer_code=row.get("customer_code"),
                    customer_name=row.get("customer_name") or "",
                    last_purchase_date=row.get("last_purchase_date"),
                    last_purchase_company=row.get("last_purchase_company") or "",
                    account_name=row.get("account_name") or "",
                    credit_limit=row.get("credit_limit"),
                    customer_status=self._normalize_status(row.get("customer_status")),
                    selected=False,
                )
                for row in data
            ]

            self.apply_filter()
            self.set_status(f"{len(self.filtered_rows)} cliente(s) encontrado(s).")
        except Exception as e:
            self.set_status("Falha ao carregar dados.")
            messagebox.showerror(APP_TITLE, f"Erro ao carregar clientes:\n\n{e}")

    def apply_filter(self):
        selected_status = self.FILTER_OPTIONS.get(self.filter_var.get())
        if selected_status is None:
            self.filtered_rows = list(self.rows)
        else:
            self.filtered_rows = [r for r in self.rows if r.customer_status == selected_status]

        if self.sort_column:
            self.filtered_rows.sort(
                key=lambda row: self._sort_value(row, self.sort_column),
                reverse=self.sort_reverse
            )
        self._refresh_tree()
        self._update_heading_titles()
        self.set_status(f"{len(self.filtered_rows)} cliente(s) encontrado(s).")

    def _row_values(self, row: CustomerRow):
        return (
            row.checkbox(),
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

    def _toggle_selected_from_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        row = self.tree_items[item_id]
        row.selected = not row.selected
        self.tree.item(item_id, values=self._row_values(row))

    def _handle_single_click_checkbox(self, event):
        region = self.tree.identify("region", event.x, event.y)
        column = self.tree.identify_column(event.x)
        item_id = self.tree.identify_row(event.y)

        if region == "cell" and column == "#1" and item_id:
            row = self.tree_items[item_id]
            row.selected = not row.selected
            self.tree.item(item_id, values=self._row_values(row))
            return "break"

    def mark_all(self):
        for row in self.filtered_rows:
            row.selected = True
        self._refresh_tree()

    def unmark_all(self):
        for row in self.filtered_rows:
            row.selected = False
        self._refresh_tree()

    def selected_rows(self) -> List[CustomerRow]:
        return [r for r in self.rows if r.selected]

    def run_action(self, query_key: str, action_name: str, new_flag: Optional[str], new_status: Optional[str]):
        selected = self.selected_rows()
        if not selected:
            messagebox.showwarning(APP_TITLE, "Selecione pelo menos um cliente.")
            return

        if query_key == "disable_credit_sql":
            sql_text = (self.config_data.get("queries", {}).get(query_key) or "").strip()
            if not sql_text:
                messagebox.showwarning(APP_TITLE, "A SQL para inativar vendas a prazo ainda não foi configurada.")
                return

        confirm = messagebox.askyesno(
            APP_TITLE,
            f"Deseja executar a ação '{action_name}' para {len(selected)} cliente(s)?"
        )
        if not confirm:
            return

        try:
            sql_text = self.config_data["queries"].get(query_key, "").strip()
            db = Database(self.config_data)
            customer_ids = [r.customer_id for r in selected]
            affected = db.execute_action(sql_text, customer_ids)

            if query_key == "disable_credit_sql":
                for row in self.rows:
                    if row.selected:
                        row.credit_limit = 0
                        row.selected = False
            elif new_flag and new_status:
                for row in self.rows:
                    if row.selected:
                        row.customer_status = new_status
                        row.selected = False
            else:
                for row in self.rows:
                    if row.selected:
                        row.selected = False

            self.apply_filter()
            messagebox.showinfo(APP_TITLE, f"Ação concluída. Registros afetados: {affected}")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Erro ao executar ação:\n\n{e}")


if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
