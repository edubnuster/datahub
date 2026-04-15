# -*- coding: utf-8 -*-
import json
from copy import deepcopy
from typing import Any, Dict
from .constants import DEFAULT_CONFIG, DEFAULT_LIST_SQL, DEFAULT_OPEN_INVOICES_SQL, CONFIG_PATH


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

        if "users" not in merged.get("security", {}):
            merged["security"]["users"] = []

        merged.setdefault("smtp", {})
        merged["smtp"].setdefault("email", "app@databrev.com.br")
        merged["smtp"].setdefault("sender_name", "")
        merged["smtp"].setdefault("host", "smtp.zoho.com")
        merged["smtp"].setdefault("password", "")
        merged["smtp"].setdefault("port", 465)
        merged["smtp"].setdefault("delay_seconds", 5)

        merged.setdefault("financeiro_agendas", [])

        current_sql = merged.get("queries", {}).get("list_inactive_customers_sql", "") or ""
        current_sql_lower = current_sql.lower()
        must_reset_inactive_sql = False
        if (
            "left join pessoa_conta" not in current_sql_lower
            or "credit_limit" not in current_sql
            or "has_account" not in current_sql_lower
        ):
            must_reset_inactive_sql = True
        if (
            "having max(l.data)" in current_sql_lower
            or "order by pessoa_nome_f(l.empresa)" in current_sql_lower
            or ("group by" in current_sql_lower and "pessoa_nome_f(l.empresa)" in current_sql_lower)
        ):
            must_reset_inactive_sql = True
        if must_reset_inactive_sql:
            merged["queries"]["list_inactive_customers_sql"] = DEFAULT_LIST_SQL

        invoices_sql = merged.get("queries", {}).get("list_open_invoices_sql", "") or ""
        invoices_sql_lower = invoices_sql.lower()

        must_reset_invoices_sql = False
        if (
            "saldo_em_aberto" not in invoices_sql_lower
            or "valor_desconto" not in invoices_sql_lower
            or "valor_baixado" not in invoices_sql_lower
            or "customer_id" not in invoices_sql_lower
            or "conta" not in invoices_sql_lower
            or "conta_nome" not in invoices_sql_lower
        ):
            must_reset_invoices_sql = True

        if "m.conta_debitar like '1.3.04%'" not in invoices_sql_lower:
            must_reset_invoices_sql = True
        if "exists (select 1 from boleto" not in invoices_sql_lower:
            must_reset_invoices_sql = True
        if "m.vencto <= current_date" in invoices_sql_lower:
            must_reset_invoices_sql = True

        if must_reset_invoices_sql:
            merged["queries"]["list_open_invoices_sql"] = DEFAULT_OPEN_INVOICES_SQL

        if not (merged.get("queries", {}).get("delete_customer_sql") or "").strip():
            merged["queries"]["delete_customer_sql"] = DEFAULT_CONFIG["queries"]["delete_customer_sql"]
        if not (merged.get("queries", {}).get("inactivate_customer_sql") or "").strip():
            merged["queries"]["inactivate_customer_sql"] = DEFAULT_CONFIG["queries"]["inactivate_customer_sql"]
        disable_sql = (merged.get("queries", {}).get("disable_credit_sql") or "").strip().lower()
        if not disable_sql or "update conta" in disable_sql:
            merged["queries"]["disable_credit_sql"] = DEFAULT_CONFIG["queries"]["disable_credit_sql"]

        normalized_agendas = []
        agendas = merged.get("financeiro_agendas", []) or []
        if isinstance(agendas, list):
            for a in agendas:
                if not isinstance(a, dict):
                    continue
                agenda = dict(a)
                agenda.setdefault("id", str(len(normalized_agendas) + 1))
                agenda.setdefault("name", "Alerta de vencimento")
                agenda["enabled"] = bool(agenda.get("enabled", False))
                agenda["send_time"] = str(agenda.get("send_time") or "06:00").strip() or "06:00"
                try:
                    agenda["days_before_due"] = max(0, min(365, int(agenda.get("days_before_due", 5) or 5)))
                except Exception:
                    agenda["days_before_due"] = 5
                try:
                    agenda["days_after_due"] = max(0, min(365, int(agenda.get("days_after_due", 0) or 0)))
                except Exception:
                    agenda["days_after_due"] = 0
                if agenda.get("days_after_due", 0) > 0:
                    agenda["days_before_due"] = 0
                elif agenda.get("days_before_due", 0) > 0:
                    agenda["days_after_due"] = 0
                agenda["group_id"] = agenda.get("group_id")
                agenda["portador_id"] = agenda.get("portador_id")
                agenda["customer_id"] = agenda.get("customer_id")
                agenda["extra_body"] = str(agenda.get("extra_body") or "")
                agenda["last_run_date"] = str(agenda.get("last_run_date") or "")
                agenda["last_run_at"] = str(agenda.get("last_run_at") or "")
                agenda["last_due_date"] = str(agenda.get("last_due_date") or "")
                agenda["last_late_minutes"] = int(agenda.get("last_late_minutes") or 0)
                agenda["last_out_of_time"] = bool(agenda.get("last_out_of_time", False))
                normalized_agendas.append(agenda)
        merged["financeiro_agendas"] = normalized_agendas

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
