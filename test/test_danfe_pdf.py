import base64
import struct
import unittest
import zlib

from app_core.danfe import danfe_pdf_from_nfe_xml


class DanfePdfTests(unittest.TestCase):
    def test_generate_pdf_from_xml_text(self):
        xml = '<NFe xmlns=" `http://www.portalfiscal.inf.br/nfe` "><infNFe Id="NFe4326"><ide><natOp>VENDA</natOp><serie>1</serie><nNF>6725</nNF><dhEmi>2026-03-26T10:10:18-03:00</dhEmi></ide><emit><CNPJ>111</CNPJ><xNome>EMITENTE</xNome><IE>ISENTO</IE><enderEmit><xLgr>RUA A</xLgr><nro>1</nro><xBairro>CENTRO</xBairro><xMun>TAPEJARA</xMun><UF>RS</UF><CEP>99950000</CEP></enderEmit></emit><dest><CNPJ>222</CNPJ><xNome>DESTINATARIO</xNome><enderDest><xLgr>RUA B</xLgr><nro>2</nro><xBairro>CENTRO</xBairro><xMun>TAPEJARA</xMun><UF>RS</UF><CEP>99950000</CEP></enderDest></dest><det nItem="1"><prod><cProd>7</cProd><xProd>DIESEL</xProd><CFOP>5929</CFOP><uCom>L</uCom><qCom>200.0060</qCom><vUnCom>7.810000</vUnCom><vProd>1562.04</vProd></prod></det><total><ICMSTot><vProd>1562.04</vProd><vDesc>234.40</vDesc><vFrete>0.00</vFrete><vNF>1327.64</vNF><vTotTrib>570.15</vTotTrib></ICMSTot></total><cobr><fat><nFat>10616</nFat><vOrig>1327.64</vOrig><vLiq>1327.64</vLiq></fat><dup><nDup>001</nDup><dVenc>2026-04-10</dVenc><vDup>1327.64</vDup></dup></cobr><infAdic><infCpl>TESTE</infCpl></infAdic></infNFe></NFe>'
        pdf_bytes, filename = danfe_pdf_from_nfe_xml(xml, fallback_suffix="x")
        self.assertTrue(pdf_bytes and pdf_bytes.startswith(b"%PDF-"))
        self.assertTrue(filename.startswith("danfe_"))
        self.assertTrue(filename.endswith(".pdf"))
        self.assertIn(b"DANFE", pdf_bytes)

    def test_html_entities_are_unescaped(self):
        xml = '<NFe xmlns="http://www.portalfiscal.inf.br/nfe"><infNFe Id="NFe4326"><ide><natOp>VENDA</natOp><serie>1</serie><nNF>6725</nNF><dhEmi>2026-03-26T10:10:18-03:00</dhEmi></ide><emit><CNPJ>111</CNPJ><xNome>EMITENTE</xNome><IE>ISENTO</IE><enderEmit><xLgr>RUA A</xLgr><nro>1</nro><xBairro>CENTRO</xBairro><xMun>TAPEJARA</xMun><UF>RS</UF><CEP>99950000</CEP></enderEmit></emit><dest><CNPJ>222</CNPJ><xNome>DESTINATARIO</xNome><enderDest><xLgr>RUA B</xLgr><nro>2</nro><xBairro>CENTRO</xBairro><xMun>TAPEJARA</xMun><UF>RS</UF><CEP>99950000</CEP></enderDest></dest><total><ICMSTot><vProd>1.00</vProd><vNF>1.00</vNF></ICMSTot></total><infAdic><infCpl>REFERENTE NFC-e S&amp;Eacute;RIE: 001, N&amp;Uacute;MERO: 821394</infCpl></infAdic></infNFe></NFe>'
        pdf_bytes, _ = danfe_pdf_from_nfe_xml(xml, fallback_suffix="x")
        self.assertTrue(pdf_bytes and pdf_bytes.startswith(b"%PDF-"))
        self.assertIn(b"S\\311RIE", pdf_bytes)
        self.assertIn(b"N\\332MERO", pdf_bytes)

    def test_embed_logo_png_as_xobject(self):
        def _chunk(ctype: bytes, data: bytes) -> bytes:
            return struct.pack(">I", len(data)) + ctype + data + struct.pack(">I", zlib.crc32(ctype + data) & 0xFFFFFFFF)

        raw = b"\x00" + bytes([255, 0, 0, 255])
        idat = zlib.compress(raw)
        ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 6, 0, 0, 0)
        logo_png = b"\x89PNG\r\n\x1a\n" + _chunk(b"IHDR", ihdr) + _chunk(b"IDAT", idat) + _chunk(b"IEND", b"")
        xml = '<NFe xmlns="http://www.portalfiscal.inf.br/nfe"><infNFe Id="NFe4326"><ide><natOp>VENDA</natOp><serie>1</serie><nNF>6725</nNF><dhEmi>2026-03-26T10:10:18-03:00</dhEmi></ide><emit><CNPJ>111</CNPJ><xNome>EMITENTE</xNome><IE>ISENTO</IE><enderEmit><xLgr>RUA A</xLgr><nro>1</nro><xBairro>CENTRO</xBairro><xMun>TAPEJARA</xMun><UF>RS</UF><CEP>99950000</CEP></enderEmit></emit><dest><CNPJ>222</CNPJ><xNome>DESTINATARIO</xNome><enderDest><xLgr>RUA B</xLgr><nro>2</nro><xBairro>CENTRO</xBairro><xMun>TAPEJARA</xMun><UF>RS</UF><CEP>99950000</CEP></enderDest></dest><total><ICMSTot><vProd>1.00</vProd><vNF>1.00</vNF></ICMSTot></total></infNFe></NFe>'
        pdf_bytes, _ = danfe_pdf_from_nfe_xml(xml, fallback_suffix="x", emit_logo_png_bytes=logo_png)
        self.assertTrue(pdf_bytes and pdf_bytes.startswith(b"%PDF-"))
        self.assertIn(b"/Subtype /Image", pdf_bytes)
        self.assertIn(b"/XObject", pdf_bytes)
        self.assertIn(b"/LogoEmit Do", pdf_bytes)


if __name__ == "__main__":
    unittest.main()
