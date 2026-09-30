import unittest

import pandas as pd

from invest.dominio.carteira import fluxos_internos, posicoes, tir
from tests import fabrica


class TestCarteira(unittest.TestCase):
    def test_posicoes_somam_compras_e_subtraem_vendas(self):
        lanc = fabrica.lancamentos([("01/01/2026", "AAAA3", "COMPRA", 10, 9, 0),
                                    ("01/02/2026", "AAAA3", "VENDA", 4, 11, 0),
                                    ("01/03/2026", "BBBB11", "COMPRA", 2, 90, 1)])
        q = posicoes(fabrica.ativos(), lanc)
        self.assertEqual(q.to_dict(), {"AAAA3": 6.0, "BBBB11": 2.0, "CCCC4": 0.0})

    def test_posicoes_sem_lancamentos(self):
        self.assertTrue((posicoes(fabrica.ativos(), fabrica.lancamentos()) == 0).all())

    def test_tir_de_10_por_cento_ao_ano(self):
        fluxos = [(pd.Timestamp("2025-01-01"), -1000), (pd.Timestamp("2026-01-01"), 1100)]
        self.assertAlmostEqual(tir(fluxos), 0.10 * 365.25 / 365, places=3)

    def test_tir_sem_solucao(self):
        self.assertIsNone(tir([(pd.Timestamp("2025-01-01"), 100), (pd.Timestamp("2026-01-01"), 50)]))
        self.assertIsNone(tir([]))

    def test_fluxos_incluem_posicao_na_data_de_referencia(self):
        lanc = fabrica.lancamentos([("01/01/2026", "AAAA3", "COMPRA", 10, 8, 2)])
        prov = fabrica.proventos([("01/06/2026", "AAAA3", "DIVIDENDO", 0.5, 10)])
        ativos = fabrica.ativos()
        f = fluxos_internos(ativos, lanc, prov, posicoes(ativos, lanc), fabrica.HOJE)
        self.assertEqual([v for _, v in f], [-82.0, 5.0, 100.0])
        self.assertEqual(f[-1][0], fabrica.HOJE)


if __name__ == "__main__":
    unittest.main()
