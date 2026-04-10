import unittest

from ui import build_pix_brcode_payload, pix_amount_str


class PixAmountStrTests(unittest.TestCase):
    def test_br_decimal_comma(self):
        self.assertEqual(pix_amount_str("32,41"), "32.41")

    def test_decimal_dot(self):
        self.assertEqual(pix_amount_str("32.41"), "32.41")

    def test_float(self):
        self.assertEqual(pix_amount_str(32.41), "32.41")

    def test_br_thousands(self):
        self.assertEqual(pix_amount_str("3.241,00"), "3241.00")

    def test_us_thousands(self):
        self.assertEqual(pix_amount_str("3,241.00"), "3241.00")

    def test_currency_prefix(self):
        self.assertEqual(pix_amount_str("R$ 32,41"), "32.41")

    def test_zero_and_negative(self):
        self.assertIsNone(pix_amount_str("0"))
        self.assertIsNone(pix_amount_str(0))
        self.assertIsNone(pix_amount_str("-10"))

    def test_invalid(self):
        self.assertIsNone(pix_amount_str("abc"))


class PixBRCodePayloadTests(unittest.TestCase):
    def _crc16_ccitt(self, payload: str) -> str:
        crc = 0xFFFF
        for b in payload.encode("utf-8"):
            crc ^= b << 8
            for _ in range(8):
                if crc & 0x8000:
                    crc = (crc << 1) ^ 0x1021
                else:
                    crc <<= 1
                crc &= 0xFFFF
        return f"{crc:04X}"

    def test_payload_has_crc(self):
        full = build_pix_brcode_payload("12345678901", "RECEBEDOR TESTE", "CIDADE", "32,41", txid="***")
        self.assertTrue(full.endswith(full[-4:]))
        self.assertTrue(full[-4:].isalnum())
        self.assertIn("6304", full)
        self.assertEqual(full[-8:-4], "6304")

        prefix = full[:-4]
        expected_crc = self._crc16_ccitt(prefix)
        self.assertEqual(full[-4:], expected_crc)

    def test_payload_has_amount_field(self):
        full = build_pix_brcode_payload("12345678901", "RECEBEDOR TESTE", "CIDADE", "32,41", txid="***")
        self.assertIn("54", full)


if __name__ == "__main__":
    unittest.main()

