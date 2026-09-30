"""Caso de uso de importação gravando numa pasta temporária (nunca em dados/)."""
import datetime as dt
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from invest.aplicacao.importacao import importar
from invest.infra import repositorio

HOJE = dt.date(2026, 9, 29)
ATIVOS = "ticker;tipo;segmento;preco;provento_anual_cota;peso_ideal;atualizado_em\nAAAA3;AÇÃO;X;10,5;1;1;29/09/2026\n"


class TestImportacao(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        pasta = Path(self.tmp.name)
        (pasta / "ativos.csv").write_text(ATIVOS, encoding="utf-8-sig")
        (pasta / "lancamentos.csv").write_text("data;ticker;operacao;quantidade;preco;custos\n"
                                               "01/01/2026;AAAA3;COMPRA;10;9,5;0\n", encoding="utf-8-sig")
        (pasta / "proventos.csv").write_text("data;ticker;tipo;valor_cota;quantidade\n", encoding="utf-8-sig")
        (pasta / "config.json").write_text("{}", encoding="utf-8")
        self.pasta = pasta
        self.patches = [mock.patch.object(repositorio, "DADOS", pasta), mock.patch("invest.config.DADOS", pasta)]
        for p in self.patches:
            p.start()

    def tearDown(self):
        for p in self.patches:
            p.stop()
        self.tmp.cleanup()

    def ler(self, nome):
        return (self.pasta / nome).read_text(encoding="utf-8-sig").splitlines()

    def test_grava_validos_e_devolve_pendentes(self):
        compras = [{"Ticker": "aaaa3", "Quantidade de cotas": 5},                       # preço = cotação
                   {"Ticker": "AAAA3", "Operação": "VENDA", "Quantidade de cotas": 99},  # venda maior
                   {"Ticker": "NOVO11", "Quantidade de cotas": 2, "Preço por cota (R$)": "100,00",
                    "Tipo do ativo": "FII", "Segmento": "Logística"}]
        proventos = [{"Ticker": "AAAA3", "Valor por cota (R$)": 0.25, "Data pagamento": "15/09/2026"}]
        resumo, pend_c, pend_p = importar(compras, proventos, HOJE)

        self.assertIn("2 lançamento(s)", resumo)
        self.assertEqual(len(pend_c), 1)
        self.assertIn("venda maior", pend_c[0]["Observação"])
        self.assertEqual(pend_p, [])
        self.assertEqual(self.ler("lancamentos.csv")[-2:], ["29/09/2026;AAAA3;COMPRA;5;10,50;0,00",
                                                           "29/09/2026;NOVO11;COMPRA;2;100,00;0,00"])
        self.assertEqual(self.ler("proventos.csv")[-1], "15/09/2026;AAAA3;DIVIDENDO;0,2500;15")
        self.assertEqual(self.ler("ativos.csv")[-1], "NOVO11;FII;Logística;0;0;0;manual")


if __name__ == "__main__":
    unittest.main()
