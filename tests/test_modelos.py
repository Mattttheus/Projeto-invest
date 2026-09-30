import unittest

import numpy as np
import pandas as pd

from invest.dominio import modelos
from invest.dominio.analise import limpar


class TestCalculo(unittest.TestCase):
    def test_derivadas_de_um_polinomio_conhecido(self):
        # log-preço f(t) = 0,1 t + 0,5 t² (t em anos, 0 = hoje): f'(0) = 0,1 e f''(0) = 1
        t = (np.arange(52) - 51) / 52
        d = modelos.derivadas_polinomiais(0.1 * t + 0.5 * t ** 2)
        self.assertAlmostEqual(d["taxa_inst"], np.expm1(0.1), places=6)
        self.assertAlmostEqual(d["curvatura"], 1.0, places=6)
        self.assertAlmostEqual(d["r2_poli"], 1.0, places=9)

    def test_ewma_de_retornos_constantes(self):
        self.assertAlmostEqual(modelos.vol_ewma(np.full(100, 0.01)), 0.01 * np.sqrt(52), places=6)

    def test_fisher(self):
        self.assertAlmostEqual(modelos.real(0.1025, 0.05), 0.05)


class TestLimpeza(unittest.TestCase):
    def test_remove_pico_isolado(self):
        p = pd.Series([109.0, 110.0, 1.06, 109.5, 110.7])
        limpo, n = limpar(p)
        self.assertEqual(n, 1)
        self.assertAlmostEqual(limpo.iloc[2], (110.0 + 109.5) / 2)

    def test_mantem_salto_verdadeiro(self):
        p = pd.Series([10.0, 10.0, 20.0, 20.5, 21.0])       # subiu e ficou: não é erro
        self.assertEqual(limpar(p)[1], 0)


class TestAlgebraLinear(unittest.TestCase):
    def setUp(self):
        self.cov = pd.DataFrame([[0.04, 0.01, 0.0], [0.01, 0.09, 0.02], [0.0, 0.02, 0.16]],
                                index=list("ABC"), columns=list("ABC"))
        self.mu = pd.Series([0.10, 0.14, 0.20], index=list("ABC"))

    def test_autovalores_reconstroem_a_covariancia(self):
        f = modelos.fatores(self.cov)
        self.assertAlmostEqual(f["explicado"].sum(), 1.0)
        self.assertAlmostEqual(f["autovalores"].sum(), np.trace(self.cov.values))

    def test_markowitz_pesos_validos(self):
        r = modelos.markowitz(self.cov, self.mu, 0.05)
        for w in r.values():
            self.assertAlmostEqual(w.sum(), 1.0, places=6)
            self.assertTrue((w >= 0).all())
        mv = modelos.risco_retorno(r["minima_variancia"], self.cov, self.mu, 0.05)
        igual = modelos.risco_retorno(pd.Series(1 / 3, index=list("ABC")), self.cov, self.mu, 0.05)
        ms = modelos.risco_retorno(r["maximo_sharpe"], self.cov, self.mu, 0.05)
        self.assertLessEqual(mv["vol"], igual["vol"])
        self.assertGreaterEqual(ms["sharpe"], igual["sharpe"])


class TestMonteCarlo(unittest.TestCase):
    def test_sem_volatilidade_e_deterministico(self):
        mc = modelos.monte_carlo(0, 100, 0.0, 0.0, 1, 0, 0.06, 0.0, n=10)
        self.assertAlmostEqual(mc["anos"].iloc[0]["p50"], 1200.0)
        self.assertAlmostEqual(mc["anos"].iloc[0]["cdi"], 1200.0)

    def test_percentis_ordenados_e_reproduziveis(self):
        a = modelos.monte_carlo(1000, 100, 0.1, 0.2, 2, 0, 0.06, 0.1, n=500)["anos"]
        b = modelos.monte_carlo(1000, 100, 0.1, 0.2, 2, 0, 0.06, 0.1, n=500)["anos"]
        self.assertTrue(a.equals(b))
        self.assertTrue((a["p5"] <= a["p50"]).all() and (a["p50"] <= a["p95"]).all())


if __name__ == "__main__":
    unittest.main()
