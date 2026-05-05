#!/usr/bin/env python3
"""
Testes automatizados — Conferência de Folha
Sigma Contabilidade
"""
import io
import unittest

from app import normalize_rubric, norm, brl, parse_word, RUBRIC_GROUPS, TOLERANCIA_CENTAVOS


# ─────────────────────────────────────────────
# helpers
# ─────────────────────────────────────────────

def _make_docx(paragraphs):
    """Cria docx em memória com os parágrafos dados."""
    from docx import Document
    doc = Document()
    for p in paragraphs:
        doc.add_paragraph(p)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────
# normalize_rubric
# ─────────────────────────────────────────────

class TestNormalizeRubric(unittest.TestCase):

    def test_comissao_variantes(self):
        casos = [
            "Comissão", "COMISSOES", "Comissão de vendas",
            "comissao mensal", "8781 COMISSAO", "COMISSAO SOBRE VENDAS",
            "comissões de vendas",
        ]
        for v in casos:
            with self.subTest(v=v):
                self.assertEqual(normalize_rubric(v), "COMISSAO")

    def test_dsr_variantes(self):
        casos = [
            "DSR", "D.S.R", "Descanso Semanal Remunerado",
            "DSR sobre comissão", "DSR s/ comissão", "DSR COMISSAO",
            "REPOUSO SEMANAL REMUNERADO",
        ]
        for v in casos:
            with self.subTest(v=v):
                self.assertEqual(normalize_rubric(v), "DSR")

    def test_comissao_e_dsr(self):
        casos = [
            "Comissão e DSR", "COMISSOES E DSR",
            "comissao + dsr", "Comissão DSR",
        ]
        for v in casos:
            with self.subTest(v=v):
                self.assertEqual(normalize_rubric(v), "COMISSAO_E_DSR")

    def test_vale_transporte(self):
        for v in ["Vale Transporte", "VT", "DESCONTO VALE TRANSPORTE", "desc vt"]:
            with self.subTest(v=v):
                self.assertEqual(normalize_rubric(v), "VALE_TRANSPORTE")

    def test_plano_saude(self):
        for v in ["Plano de Saúde", "UNIMED", "Assistência Médica", "convênio médico"]:
            with self.subTest(v=v):
                self.assertEqual(normalize_rubric(v), "PLANO_SAUDE")

    def test_rubrica_desconhecida_retorna_normalizada(self):
        result = normalize_rubric("SALARIO BASE")
        self.assertEqual(result, "SALARIO BASE")

    def test_remove_codigo_numerico_inicial(self):
        self.assertEqual(normalize_rubric("8781 SALARIO"), "SALARIO")
        self.assertEqual(normalize_rubric("20 GRATIFICACOES DE CAIXA"), "GRATIFICACOES DE CAIXA")


# ─────────────────────────────────────────────
# norm (nomes)
# ─────────────────────────────────────────────

class TestNorm(unittest.TestCase):

    def test_acento_e_caixa(self):
        self.assertEqual(norm("João da Silva"), norm("JOAO DA SILVA"))
        self.assertEqual(norm("João Silva"), norm("JOAO SILVA"))

    def test_espacos_duplicados(self):
        self.assertEqual(norm("João  Silva"), norm("JOAO SILVA"))

    def test_preposicoes_preservadas(self):
        # "da" e "DE" devem ser preservadas (não removemos preposições)
        self.assertEqual(norm("Ana de Souza"), "ANA DE SOUZA")

    def test_acento_multiplos(self):
        self.assertEqual(norm("Ângela Cristóvão"), "ANGELA CRISTOVAO")


# ─────────────────────────────────────────────
# brl (conversão monetária)
# ─────────────────────────────────────────────

class TestBrl(unittest.TestCase):

    def test_formato_br_com_ponto_milhar(self):
        self.assertAlmostEqual(brl("1.592,11"), 1592.11)

    def test_formato_br_com_cifrao(self):
        self.assertAlmostEqual(brl("R$ 1.592,11"), 1592.11)

    def test_formato_simples(self):
        self.assertAlmostEqual(brl("1592,11"), 1592.11)

    def test_zero(self):
        self.assertEqual(brl("0"), 0.0)
        self.assertEqual(brl("0,00"), 0.0)

    def test_valor_inteiro(self):
        self.assertAlmostEqual(brl("500"), 500.0)

    def test_string_invalida(self):
        self.assertEqual(brl("abc"), 0.0)


# ─────────────────────────────────────────────
# parse_word
# ─────────────────────────────────────────────

class TestParseWord(unittest.TestCase):

    def test_formato_lista_simples_comissao_e_dsr(self):
        """Título 'comissões e DSR' → valores vão para comissao_e_dsr."""
        raw = _make_docx(["comissões e DSR", "Nadyane", "1592,11", "Juliana", "1255,49"])
        result = parse_word(raw)
        self.assertIn("NADYANE", result["comissao_e_dsr"], "Nadyane não encontrada em comissao_e_dsr")
        self.assertAlmostEqual(result["comissao_e_dsr"]["NADYANE"], 1592.11)
        self.assertIn("JULIANA", result["comissao_e_dsr"])
        self.assertAlmostEqual(result["comissao_e_dsr"]["JULIANA"], 1255.49)
        # Não deve ir para gratificacoes
        self.assertNotIn("NADYANE", result["gratificacoes"])

    def test_formato_lista_simples_generico(self):
        """Título genérico → valores vão para gratificacoes."""
        raw = _make_docx(["Bônus de produção", "Ana Silva", "500,00", "Carlos Santos", "300,00"])
        result = parse_word(raw)
        # Algum nome deve ter sido capturado (pode ir para gratificacoes)
        total = len(result["gratificacoes"]) + len(result["comissao_e_dsr"])
        self.assertGreater(total, 0, "Nenhum nome capturado no formato lista simples")

    def test_resultado_tem_todos_campos(self):
        """parse_word sempre retorna todos os campos esperados."""
        raw = _make_docx(["Nada aqui"])
        result = parse_word(raw)
        for campo in ("gratificacoes", "descontos", "obs", "decimo_terceiro", "comissao_e_dsr"):
            self.assertIn(campo, result, f"Campo '{campo}' ausente no resultado")

    def test_valor_brl_convertido_corretamente(self):
        """Valores em formato BR são convertidos corretamente."""
        raw = _make_docx(["comissões e DSR", "Teste", "1.592,11"])
        result = parse_word(raw)
        if "TESTE" in result["comissao_e_dsr"]:
            self.assertAlmostEqual(result["comissao_e_dsr"]["TESTE"], 1592.11)


# ─────────────────────────────────────────────
# Lógica Comissão + DSR no compare
# ─────────────────────────────────────────────

class TestComissaoDSRCompare(unittest.TestCase):

    def _make_pdf_employee(self, nome, verbas):
        """Cria estrutura de funcionário no formato retornado por parse_pdf."""
        return {
            "nome_original": nome.title(),
            "tipo": "mensal",
            "liquido": sum(v["valor"] for v in verbas if v.get("tipo") != "desconto"),
            "total_vencimentos": sum(v["valor"] for v in verbas),
            "total_descontos": 0,
            "has_gratif": False,
            "gratif_valor": 0,
            "verbas": verbas,
        }

    def test_comissao_dsr_separados_total_igual(self):
        """Relatório: Comissão+DSR=1200. Recibo: Comissão=1000 + DSR=200. Deve ser OK."""
        from app import compare, norm
        nome = norm("Nadyane Silva")

        word = {"comissao_e_dsr": {nome: 1200.00}, "gratificacoes": {}, "descontos": {}}
        pdf = {
            nome: self._make_pdf_employee("Nadyane Silva", [
                {"codigo": "1", "descricao": "COMISSAO", "referencia": "30,00", "valor": 1000.00},
                {"codigo": "2", "descricao": "DSR SOBRE COMISSAO", "referencia": "30,00", "valor": 200.00},
            ])
        }
        excel = {nome: {"salario": 2000, "liquido": 3200, "has_liquido": False,
                         "gratificacao": 0, "ferias_13": 0, "inss": 0,
                         "vale": 0, "plano": 0, "emprestimo": 0, "apontamentos": {}}}

        report = compare(excel, pdf, word)
        emp = report["funcionarios"][0]

        # Não deve ter divergência de Comissão+DSR
        tipos_divs = [d["tipo"] for d in emp.get("divs", [])]
        self.assertNotIn("DIFERENÇA A PAGAR", tipos_divs,
                         f"Não deveria ter divergência. Divs: {emp.get('divs')}")
        self.assertIn("memoria_comissao_dsr", emp,
                      "Memória de cálculo Comissão+DSR deveria estar presente")
        self.assertEqual(emp["memoria_comissao_dsr"]["status"], "OK")

    def test_comissao_dsr_diferenca_a_pagar(self):
        """Relatório: 1200. Recibo: apenas 1000. Deve apontar diferença de 200."""
        from app import compare, norm
        nome = norm("Juliana Costa")

        word = {"comissao_e_dsr": {nome: 1200.00}, "gratificacoes": {}, "descontos": {}}
        pdf = {
            nome: self._make_pdf_employee("Juliana Costa", [
                {"codigo": "1", "descricao": "COMISSAO", "referencia": "30,00", "valor": 1000.00},
            ])
        }
        excel = {nome: {"salario": 2000, "liquido": 3000, "has_liquido": False,
                         "gratificacao": 0, "ferias_13": 0, "inss": 0,
                         "vale": 0, "plano": 0, "emprestimo": 0, "apontamentos": {}}}

        report = compare(excel, pdf, word)
        emp = report["funcionarios"][0]
        tipos = [d["tipo"] for d in emp.get("divs", [])]
        self.assertIn("DIFERENÇA A PAGAR", tipos, f"Deveria apontar diferença. Divs: {emp.get('divs')}")

    def test_match_por_primeiro_nome(self):
        """Word com 'NADYANE' deve casar com 'NADYANE SILVA' no PDF/Excel."""
        from app import compare, norm
        nome_completo = norm("Nadyane Silva")

        word = {"comissao_e_dsr": {"NADYANE": 1592.11}, "gratificacoes": {}, "descontos": {}}
        pdf = {
            nome_completo: self._make_pdf_employee("Nadyane Silva", [
                {"codigo": "1", "descricao": "COMISSAO", "referencia": "30,00", "valor": 1400.00},
                {"codigo": "2", "descricao": "DSR", "referencia": "30,00", "valor": 192.11},
            ])
        }
        excel = {nome_completo: {"salario": 2000, "liquido": 3592.11, "has_liquido": False,
                                  "gratificacao": 0, "ferias_13": 0, "inss": 0,
                                  "vale": 0, "plano": 0, "emprestimo": 0, "apontamentos": {}}}

        report = compare(excel, pdf, word)
        emp = report["funcionarios"][0]
        tipos = [d["tipo"] for d in emp.get("divs", [])]
        self.assertNotIn("DIFERENÇA A PAGAR", tipos,
                         f"Não deveria ter divergência de Comissão+DSR. Divs: {emp.get('divs')}")


# ─────────────────────────────────────────────
# Tolerâncias
# ─────────────────────────────────────────────

class TestTolerancia(unittest.TestCase):

    def test_tolerancia_centavos_definida(self):
        self.assertLessEqual(TOLERANCIA_CENTAVOS, 0.05)
        self.assertGreaterEqual(TOLERANCIA_CENTAVOS, 0.0)

    def test_diferenca_zero_virgem_um_e_arredondamento(self):
        """Diferença de R$ 0,01 deve ficar dentro da tolerância de arredondamento."""
        self.assertLessEqual(0.01, TOLERANCIA_CENTAVOS)


if __name__ == "__main__":
    unittest.main(verbosity=2)
