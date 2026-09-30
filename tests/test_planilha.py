"""Ponta a ponta sem internet e sem Excel: domínio + apresentação montando as duas versões da planilha."""
import unittest
import warnings

from invest.apresentacao.excel import construir
from invest.dominio.analise import analisar
from invest.dominio.base import BaseDados
from invest.dominio.validacao import validar
from tests import fabrica

ABAS = ["Painel", "Carteira", "Análises", "Meta", "Quant", "Modelos", "Projeção", "Cadastro", "Ativos", "Lançamentos",
        "Proventos", "Premissas", "Verificar", "Base Painel"]


def base(com_historico=True):
    ativos = fabrica.ativos()
    lanc = fabrica.lancamentos([("02/01/2026", "AAAA3", "COMPRA", 100, 9, 1), ("05/02/2026", "BBBB11", "COMPRA", 10, 95, 0)])
    prov = fabrica.proventos([("15/06/2026", "BBBB11", "RENDIMENTO", 0.8, 10)])
    hist = fabrica.historico() if com_historico else fabrica.historico(tickers=())
    resultado = analisar(ativos, lanc, prov, fabrica.CFG, hist, fabrica.HOJE) if com_historico else None
    avisos = validar(ativos, lanc, prov) + (resultado["avisos"] if resultado else [])
    return BaseDados(ativos, lanc, prov, fabrica.CFG, avisos, resultado)


class TestPlanilha(unittest.TestCase):
    def test_excel_e_google_com_todas_as_abas(self):
        b = base()
        for google in (False, True):
            with warnings.catch_warnings():
                warnings.simplefilter("error", FutureWarning)
                wb = construir(b, google=google)
            self.assertEqual(wb.sheetnames, ABAS)
            self.assertEqual(wb["Quant"]["A5"].value, "=Carteira!A5")

    def test_carteira_real_e_tir_entram_na_analise(self):
        c = base().analise["carteira"]
        self.assertEqual(c["base"], "carteira real (lançamentos)")
        self.assertIsNotNone(c["tir"])
        self.assertAlmostEqual(base().analise["por_ativo"]["contrib_risco"].sum(), 1.0, places=6)

    def test_meta_padrao_e_personalizada(self):
        import pandas as pd
        from invest.dominio import metas
        b = base()
        pers = pd.DataFrame({"ticker": ["BBBB11"], "meta_mensal": [500.0], "prazo_meses": [24.0]})
        b.metas = metas.resolver(b.ativos, pers, {"meta_mensal_por_ativo": 1099, "prazo_meta_meses": 60})
        ws = construir(b)["Meta"]
        self.assertEqual(ws["G5"].value, '=IF(A5="","",META_MENSAL)')        # AAAA3: padrão (fórmula)
        self.assertEqual((ws["G6"].value, ws["H6"].value), (500.0, 24))        # BBBB11: personalizada
        self.assertIn("ROUNDUP(G5/E5,0)", ws["I5"].value)
        self.assertIn("(1+F5)^H5", ws["T5"].value)
        self.assertEqual(construir(b)["Carteira"]["W5"].value, "=IF(N(Meta!G5)>0,MIN(1,V5/Meta!G5),0)")

    def test_sem_historico_gera_mesmo_assim(self):
        wb = construir(base(com_historico=False))
        self.assertIn("Sem histórico", wb["Quant"]["A3"].value)


if __name__ == "__main__":
    unittest.main()
