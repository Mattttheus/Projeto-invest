"""Página web: dados.json serializável (sem NaN/numpy) e arquivos certos para cada modo."""
import json
import tempfile
import unittest
from pathlib import Path

from invest.apresentacao import web
from tests import fabrica
from tests.test_planilha import base


class TestWeb(unittest.TestCase):
    def test_dados_json_estrito(self):
        d = web.dados(base(), "demo", fabrica.historico())
        texto = json.dumps(d, allow_nan=False)            # falha se sobrar NaN/infinito
        self.assertIn("AAAA3", texto)
        self.assertEqual(d["lancamentos"][0]["data"], "2026-01-02")
        self.assertFalse(d["analise"]["vazio"])
        self.assertIn("monte_carlo", d["analise"]["modelos"])
        self.assertEqual({m["ticker"] for m in d["metas"]}, {"AAAA3", "BBBB11", "CCCC4"})

    def test_sem_historico(self):
        d = web.dados(base(com_historico=False), "local")
        self.assertTrue(d["analise"]["vazio"])
        json.dumps(d, allow_nan=False)

    def test_demo_nao_leva_api(self):
        with tempfile.TemporaryDirectory() as pasta:
            for modo, tem_api in (("demo", False), ("local", True)):
                destino = Path(pasta) / modo
                web.gerar(base(), destino, modo, fabrica.historico())
                self.assertTrue((destino / "index.html").exists())
                self.assertEqual((destino / "api.php").exists(), tem_api)
                self.assertEqual(json.loads((destino / "dados.json").read_text(encoding="utf-8"))["modo"], modo)


if __name__ == "__main__":
    unittest.main()
