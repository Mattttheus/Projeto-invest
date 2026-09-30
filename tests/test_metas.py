import unittest

import pandas as pd

from invest.dominio import metas
from tests import fabrica

CFG = {"meta_mensal_por_ativo": 1099, "prazo_meta_meses": 60}


class TestMetas(unittest.TestCase):
    def test_padrao_para_todos_sem_personalizacao(self):
        m = metas.resolver(fabrica.ativos(), metas.vazias(), CFG)
        self.assertEqual(m["meta_mensal"].tolist(), [1099.0] * 3)
        self.assertEqual(m["prazo_meses"].tolist(), [60] * 3)
        self.assertFalse(m["personalizada"].any())

    def test_personalizacao_parcial_completa_com_padrao(self):
        pers = pd.DataFrame({"ticker": ["BBBB11"], "meta_mensal": [500.0], "prazo_meses": [float("nan")]})
        m = metas.resolver(fabrica.ativos(), pers, CFG)
        self.assertEqual((m.loc["BBBB11", "meta_mensal"], m.loc["BBBB11", "prazo_meses"]), (500.0, 60))
        self.assertTrue(m.loc["BBBB11", "personalizada"])
        self.assertEqual(m.loc["AAAA3", "meta_mensal"], 1099.0)

    def test_entrada_digitada(self):
        self.assertEqual(metas.entrada("X", "=META_MENSAL", None), {"ticker": "X", "meta_mensal": None, "prazo_meses": None})
        self.assertEqual(metas.entrada("X", "R$ 1.500,00", 36.4), {"ticker": "X", "meta_mensal": 1500.0, "prazo_meses": 36})
        with self.assertRaises(ValueError):
            metas.entrada("X", 100, 0)

    def test_validacao(self):
        pers = pd.DataFrame({"ticker": ["ZZZZ3", "AAAA3"], "meta_mensal": [100.0, -5.0], "prazo_meses": [12.0, 12.0]})
        textos = " | ".join(a.mensagem for a in metas.validar(pers, fabrica.ativos()))
        self.assertIn("não está em ativos.csv", textos)
        self.assertIn("negativa", textos)


if __name__ == "__main__":
    unittest.main()
