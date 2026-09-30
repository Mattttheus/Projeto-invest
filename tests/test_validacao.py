import unittest

from invest.dominio.avisos import ERRO, tem_erro
from invest.dominio.validacao import validar
from tests import fabrica


def mensagens(avisos):
    return " | ".join(a.mensagem for a in avisos)


class TestValidacao(unittest.TestCase):
    def test_base_limpa_sem_avisos(self):
        self.assertEqual(validar(fabrica.ativos(), fabrica.lancamentos(), fabrica.proventos()), [])

    def test_ticker_desconhecido_e_erro(self):
        lanc = fabrica.lancamentos([("01/01/2026", "ZZZZ3", "COMPRA", 1, 10, 0)])
        self.assertTrue(tem_erro(validar(fabrica.ativos(), lanc, fabrica.proventos())))

    def test_venda_maior_que_a_posicao(self):
        lanc = fabrica.lancamentos([("01/01/2026", "AAAA3", "COMPRA", 5, 10, 0),
                                    ("01/02/2026", "AAAA3", "VENDA", 8, 10, 0)])
        self.assertIn("Venda maior", mensagens(validar(fabrica.ativos(), lanc, fabrica.proventos())))

    def test_numero_invalido_e_erro(self):
        ativos = fabrica.ativos(preco=[10.0, float("nan"), 20.0])
        avisos = validar(ativos, fabrica.lancamentos(), fabrica.proventos())
        self.assertTrue(any(a.gravidade == ERRO and "não numérico" in a.mensagem for a in avisos))

    def test_operacao_vazia_nao_quebra(self):
        lanc = fabrica.lancamentos([("01/01/2026", "AAAA3", None, 5, 10, 0)])
        self.assertIn("Operação inválida", mensagens(validar(fabrica.ativos(), lanc, fabrica.proventos())))


if __name__ == "__main__":
    unittest.main()
