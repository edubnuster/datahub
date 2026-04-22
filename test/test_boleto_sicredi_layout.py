import unittest

from ui import build_boleto_pdf_bytes


class BoletoSicrediLayoutTests(unittest.TestCase):
    def test_sicredi_layout_contains_delivery_receipt(self):
        boleto_data = {
            "banco_codigo": "748",
            "banco_nome": "Sicredi",
            "portador_nome": "Sicredi",
            "nosso_numero": "26/247098-7",
            "documento": "LAN6100",
            "vencto_display": "15/04/2026",
            "valor_display": "235,11",
            "linha_digitavel": "74891.12628 47098.702682 06066.221091 8 14170000023511",
            "codigo_barra": "7489141700000023511126284709870268206066221091",
            "cedente_nome": "AUTO POSTO KANINHA LTDA - MATRIZ",
            "cedente_documento": "88.305.412/0001-80",
            "sacado_nome": "TRR KANINHA COMB LTDA",
            "sacado_inscricao": "10.595.727/0001-12",
            "sacado_endereco": "RUA JULIO DE CASTILHOS, 910 - CENTRO",
            "sacado_cidade_uf": "99.950-000 - TAPEJARA RS",
            "agencia": "0268",
            "agencia_digito": "0",
            "nr_conta": "06622",
            "conta_digito": "2",
            "portador_carteira": "01",
            "mensagem": "TESTE",
        }
        inv = type("Inv", (), {"open_balance": 235.11, "amount": 235.11})()
        pdf = build_boleto_pdf_bytes(boleto_data, inv, include_pix_qrcode=False)
        self.assertTrue(pdf and pdf.startswith(b"%PDF-"))
        self.assertIn(b"Recibo de Entrega", pdf)
        self.assertIn(b"Recibo do Pagador", pdf)
        self.assertIn(b"/Subtype /Image", pdf)


if __name__ == "__main__":
    unittest.main()
