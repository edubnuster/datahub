import unittest
from datetime import date

from app_core.models import InvoiceRow
from ui import (
    _detect_email_attachment_flags,
    _mime_parts_from_filename,
    build_attachments_note_text,
    build_faturas_detalhamento_txt_bytes,
)


class FaturaTxtAttachmentTests(unittest.TestCase):
    def test_mime_txt(self):
        self.assertEqual(_mime_parts_from_filename("fatura_123.txt"), ("text", "plain"))

    def test_detect_flag_fatura_txt(self):
        flags = _detect_email_attachment_flags(["faturas_cliente_20260101_0000.txt"])
        self.assertTrue(flags.get("fatura_txt"))

    def test_note_mentions_fatura_txt(self):
        note = build_attachments_note_text(
            has_boleto=False,
            has_fatura_pdf=False,
            has_fatura_txt=True,
            has_xml=False,
            has_danfe=False,
            has_assinatura=False,
        )
        self.assertIn("fatura", note)

    def test_build_txt_has_title_and_items(self):
        inv = InvoiceRow(
            invoice_id="ADM3404",
            company="AUTO POSTO KANINHA LTDA",
            customer_id=1,
            customer_code="7384",
            customer_name="ALEX SCHOLER LAITARTE",
            issue_date=date(2026, 4, 13),
            due_date=date(2026, 4, 28),
            open_balance=75.62,
            movto_id=100,
        )
        purchase_map = {
            100: {
                "documents": [
                    {
                        "documento": "12345",
                        "dt": date(2026, 4, 13),
                        "total": 75.62,
                        "items": [{"product": "GASOLINA COMUM", "quantity": 11.70, "item_total": 75.62}],
                    }
                ],
                "items_total": 75.62,
                "invoice_amount": 75.62,
            }
        }
        data, name = build_faturas_detalhamento_txt_bytes([inv], purchase_info_map=purchase_map)
        self.assertTrue(name.endswith(".txt"))
        txt = data.decode("utf-8")
        self.assertIn("DETALHAMENTO DE FATURAS", txt)
        self.assertIn("Fatura nr.: ADM3404", txt)
        self.assertIn("GASOLINA COMUM", txt)


if __name__ == "__main__":
    unittest.main()
