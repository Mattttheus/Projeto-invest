import datetime as dt
import unittest

from invest.dominio.lancamentos import numero, validar_lancamento, validar_provento

HOJE = dt.date(2026, 9, 29)


def cotacao(_):
    return 10.0


class TestNumero(unittest.TestCase):
    def test_formatos(self):
        self.assertEqual(numero("R$ 1.234,56"), 1234.56)
        self.assertEqual(numero("12.5"), 12.5)          # ponto decimal digitado como texto
        self.assertEqual(numero(7), 7.0)
        self.assertIsNone(numero(""))


class TestLancamento(unittest.TestCase):
    def novo(self, **campos):
        d = {"Ticker": "aaaa3", "Quantidade de cotas": 10, **campos}
        return validar_lancamento(d, {"AAAA3"}, {}, {"AAAA3": 5}, cotacao, HOJE)

    def test_compra_usa_cotacao_e_data_de_hoje_quando_vazias(self):
        l = self.novo()
        self.assertEqual((l.ticker, l.operacao, l.preco, l.data), ("AAAA3", "COMPRA", 10.0, HOJE))

    def test_venda_maior_que_as_cotas(self):
        with self.assertRaisesRegex(ValueError, "venda maior"):
            self.novo(**{"Operação": "venda"})

    def test_ticker_novo_exige_tipo(self):
        with self.assertRaisesRegex(ValueError, "Tipo do ativo"):
            validar_lancamento({"Ticker": "NOVO3", "Quantidade de cotas": 1}, set(), {}, {}, cotacao, HOJE)
        novos = {}
        validar_lancamento({"Ticker": "NOVO3", "Quantidade de cotas": 1, "Tipo do ativo": "fii"},
                           set(), novos, {}, cotacao, HOJE)
        self.assertEqual(novos, {"NOVO3": ("FII", "")})

    def test_provento_usa_cotas_atuais(self):
        p = validar_provento({"Ticker": "AAAA3", "Valor por cota (R$)": "0,50"}, {"AAAA3"}, {}, {"AAAA3": 5}, HOJE)
        self.assertEqual((p.tipo, p.valor_cota, p.quantidade), ("DIVIDENDO", 0.5, 5))


if __name__ == "__main__":
    unittest.main()
