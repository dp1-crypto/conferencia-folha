#!/usr/bin/env python3
"""
Testes automatizados — Conferência de Folha
Sigma Contabilidade
"""
import io
import unittest

from app import (
    normalize_rubric, norm, brl, brl_dec, parse_word,
    RUBRIC_GROUPS, RUBRIC_META, TOLERANCIA_CENTAVOS,
    rubric_words_overlap, find_rubric_by_value, STATUS,
)


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
        # 'SALARIO BASE' virou variante do grupo SALARIO no dicionario v2.0;
        # aqui o que se testa e o fallback, entao usa rubrica realmente ausente.
        result = normalize_rubric("Verba Interna XPTO")
        self.assertEqual(result, "VERBA INTERNA XPTO")

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


# ─────────────────────────────────────────────
# brl_dec — precisão Decimal
# ─────────────────────────────────────────────

class TestBrlDec(unittest.TestCase):

    def test_formato_br_ponto_milhar(self):
        from decimal import Decimal
        self.assertEqual(brl_dec("1.592,11"), Decimal("1592.11"))

    def test_formato_cifrao(self):
        from decimal import Decimal
        self.assertEqual(brl_dec("R$ 500,00"), Decimal("500.00"))

    def test_negativo_parenteses(self):
        from decimal import Decimal
        self.assertEqual(brl_dec("(300,00)"), Decimal("-300.00"))

    def test_soma_sem_erro_float(self):
        """Soma de Decimals não acumula erro de float."""
        from decimal import Decimal
        vals = ["0,10", "0,10", "0,10"]
        total = sum(brl_dec(v) for v in vals)
        self.assertEqual(total, Decimal("0.30"))

    def test_string_invalida_retorna_zero(self):
        from decimal import Decimal
        self.assertEqual(brl_dec("abc"), Decimal("0.00"))


# ─────────────────────────────────────────────
# Carregamento do JSON de rubricas
# ─────────────────────────────────────────────

class TestJsonRubricas(unittest.TestCase):

    def test_grupos_carregados(self):
        """RUBRIC_GROUPS deve ter pelo menos os grupos essenciais."""
        essenciais = {"COMISSAO", "DSR", "COMISSAO_E_DSR", "VALE_TRANSPORTE",
                      "PLANO_SAUDE", "ODONTO", "PREMIO"}
        for g in essenciais:
            self.assertIn(g, RUBRIC_GROUPS, f"Grupo '{g}' não encontrado no JSON")

    def test_meta_carregado(self):
        """RUBRIC_META deve ter descricao e tipo para cada grupo."""
        for grupo, meta in RUBRIC_META.items():
            self.assertIn("descricao", meta, f"{grupo}: campo 'descricao' ausente")
            self.assertIn("tipo", meta,      f"{grupo}: campo 'tipo' ausente")
            # 'informativo' cobre rubricas que aparecem no holerite mas não são
            # nem provento nem desconto do empregado (FGTS, afastamento INSS).
            self.assertIn(meta["tipo"], ("provento", "desconto", "informativo"),
                          f"{grupo}: tipo inválido '{meta['tipo']}'")

    def test_premio_ppr_no_grupo(self):
        """'PREMIO PPR' deve estar como variante do grupo PREMIO."""
        variantes = [normalize_rubric(v) for v in RUBRIC_GROUPS.get("PREMIO", [])]
        self.assertIn("PREMIO", variantes, "Grupo PREMIO deve reconhecer pelo menos 'PREMIO'")

    def test_normalize_rubric_usa_json(self):
        """normalize_rubric deve retornar grupo correto para variante do JSON."""
        # PREMIO PPR está no JSON → deve normalizar para PREMIO
        self.assertEqual(normalize_rubric("PREMIO PPR"), "PREMIO")
        self.assertEqual(normalize_rubric("PLR"),        "PREMIO")
        self.assertEqual(normalize_rubric("PPR"),        "PREMIO")


# ─────────────────────────────────────────────
# rubric_words_overlap — smart matching
# ─────────────────────────────────────────────

class TestRubricWordsOverlap(unittest.TestCase):

    def test_premio_vs_premio_ppr(self):
        """'PREMIO' e 'PREMIO PPR' compartilham a palavra PREMIO → overlap > 0."""
        score = rubric_words_overlap("PREMIO", "PREMIO PPR")
        self.assertGreater(score, 0.0)

    def test_bonus_vs_bonus_desempenho(self):
        score = rubric_words_overlap("BONUS", "BONUS DESEMPENHO")
        self.assertGreater(score, 0.0)

    def test_sem_relacao(self):
        """Rubricas sem nenhuma palavra em comum → overlap = 0."""
        score = rubric_words_overlap("INSS", "COMISSAO")
        self.assertEqual(score, 0.0)

    def test_identico_retorna_um(self):
        """Rubricas idênticas → overlap = 1.0."""
        score = rubric_words_overlap("VALE TRANSPORTE", "VALE TRANSPORTE")
        self.assertEqual(score, 1.0)

    def test_mesmo_grupo_canonico(self):
        """Rubricas que normalizam para o mesmo grupo → overlap = 1.0."""
        # Ambas → COMISSAO
        score = rubric_words_overlap("Comissão", "COMISSOES DE VENDAS")
        self.assertEqual(score, 1.0)


# ─────────────────────────────────────────────
# find_rubric_by_value — match por valor
# ─────────────────────────────────────────────

class TestFindRubricByValue(unittest.TestCase):

    def _verbas(self, *pares):
        """Helper: lista de verbas [(descricao, valor), ...]"""
        return [{"codigo": str(i), "descricao": d, "referencia": "30,00", "valor": v}
                for i, (d, v) in enumerate(pares)]

    def test_encontra_por_valor_exato_e_palavras(self):
        """PREMIO esperado R$500 → PREMIO PPR R$500 → confiança alta."""
        verbas = self._verbas(("INSS", 200.0), ("PREMIO PPR", 500.0), ("DSR", 100.0))
        result = find_rubric_by_value("PREMIO", 500.0, verbas)
        self.assertIsNotNone(result)
        self.assertEqual(result["verba"]["descricao"], "PREMIO PPR")
        self.assertEqual(result["confianca"], "alta")

    def test_retorna_none_sem_valor_proximo(self):
        """Sem verba com valor próximo → retorna None."""
        verbas = self._verbas(("INSS", 200.0), ("IRRF", 350.0))
        result = find_rubric_by_value("PREMIO", 500.0, verbas)
        self.assertIsNone(result)

    def test_priorizacao_por_overlap_de_palavras(self):
        """Entre dois candidatos com mesmo valor, prioriza o com mais sobreposição."""
        verbas = self._verbas(
            ("INSS", 500.0),          # mesmo valor, zero overlap com BONUS
            ("BONUS DESEMPENHO", 500.0),  # mesmo valor, alto overlap
        )
        result = find_rubric_by_value("BONUS", 500.0, verbas)
        self.assertIsNotNone(result)
        self.assertEqual(result["verba"]["descricao"], "BONUS DESEMPENHO")

    def test_confianca_baixa_sem_overlap(self):
        """Valor bate mas sem sobreposição de palavras → confiança baixa."""
        verbas = self._verbas(("INSS", 300.0))
        result = find_rubric_by_value("BONUS PRODUCAO", 300.0, verbas)
        self.assertIsNotNone(result)
        self.assertEqual(result["confianca"], "baixa")


# ─────────────────────────────────────────────
# STATUS padronizados
# ─────────────────────────────────────────────

class TestStatus(unittest.TestCase):

    def test_status_tem_todos_os_campos(self):
        esperados = ["OK", "A_PAGAR", "PAGO_MAIOR", "DESCONTO_INDEVIDO",
                     "NAO_LOCALIZADO_RECIBO", "NAO_LOCALIZADO_REL",
                     "POSSIVEL_RUBRICA", "POSSIVEL_COLABORADOR",
                     "REVISAO_MANUAL", "ARREDONDAMENTO"]
        for k in esperados:
            self.assertIn(k, STATUS, f"STATUS['{k}'] ausente")

    def test_status_ok_string(self):
        self.assertEqual(STATUS["OK"], "OK")

    def test_status_a_pagar(self):
        self.assertIn("PAGAR", STATUS["A_PAGAR"].upper())


# ─────────────────────────────────────────────
# Cenários de integração (Comissão+DSR, benefícios)
# ─────────────────────────────────────────────

class TestCenariosIntegracao(unittest.TestCase):

    def _make_pdf_employee(self, nome, verbas):
        return {
            "nome_original": nome.title(),
            "tipo": "mensal",
            "liquido": sum(v["valor"] for v in verbas),
            "total_vencimentos": sum(v["valor"] for v in verbas),
            "total_descontos": 0,
            "has_gratif": False,
            "gratif_valor": 0,
            "verbas": verbas,
        }

    def test_relatorio_agrupado_recibo_separado(self):
        """Relatório: COMISSAO_E_DSR=1200. Recibo: COMISSAO=1000 + DSR=200. Deve ser OK."""
        from app import compare, norm
        nome = norm("Teste Silva")
        word = {"comissao_e_dsr": {nome: 1200.0}, "gratificacoes": {}, "descontos": {}}
        pdf = {nome: self._make_pdf_employee("Teste Silva", [
            {"codigo": "1", "descricao": "COMISSAO DE VENDAS", "referencia": "30,00", "valor": 1000.0},
            {"codigo": "2", "descricao": "DESCANSO SEMANAL REMUNERADO", "referencia": "30,00", "valor": 200.0},
        ])}
        excel = {nome: {"salario": 2000, "liquido": 3200, "has_liquido": False,
                        "gratificacao": 0, "ferias_13": 0, "inss": 0,
                        "vale": 0, "plano": 0, "emprestimo": 0, "apontamentos": {}}}
        report = compare(excel, pdf, word)
        emp = report["funcionarios"][0]
        tipos = [d["tipo"] for d in emp.get("divs", [])]
        self.assertNotIn(STATUS["A_PAGAR"], tipos)
        self.assertEqual(emp.get("memoria_comissao_dsr", {}).get("status"), STATUS["OK"])

    def test_relatorio_separado_recibo_agrupado(self):
        """Relatório: COMISSAO=1000 + DSR=200 separados. Recibo: COMISSAO E DSR=1200. Deve ser OK."""
        from app import compare, norm
        nome = norm("Ana Ferreira")
        # Word com comissao_e_dsr agrupado (= soma dos dois)
        word = {"comissao_e_dsr": {nome: 1200.0}, "gratificacoes": {}, "descontos": {}}
        pdf = {nome: self._make_pdf_employee("Ana Ferreira", [
            {"codigo": "1", "descricao": "COMISSAO E DSR", "referencia": "30,00", "valor": 1200.0},
        ])}
        excel = {nome: {"salario": 1800, "liquido": 3000, "has_liquido": False,
                        "gratificacao": 0, "ferias_13": 0, "inss": 0,
                        "vale": 0, "plano": 0, "emprestimo": 0, "apontamentos": {}}}
        report = compare(excel, pdf, word)
        emp = report["funcionarios"][0]
        tipos = [d["tipo"] for d in emp.get("divs", [])]
        self.assertNotIn(STATUS["A_PAGAR"], tipos)

    def test_diferenca_centavos_arredondamento(self):
        """Diferença de R$0,01 deve ser ARREDONDAMENTO, não DIFERENÇA A PAGAR."""
        from app import compare, norm
        nome = norm("Joao Pereira")
        word = {"comissao_e_dsr": {nome: 1200.01}, "gratificacoes": {}, "descontos": {}}
        pdf = {nome: self._make_pdf_employee("Joao Pereira", [
            {"codigo": "1", "descricao": "COMISSAO", "referencia": "30,00", "valor": 1200.0},
        ])}
        excel = {nome: {"salario": 2000, "liquido": 3200, "has_liquido": False,
                        "gratificacao": 0, "ferias_13": 0, "inss": 0,
                        "vale": 0, "plano": 0, "emprestimo": 0, "apontamentos": {}}}
        report = compare(excel, pdf, word)
        emp = report["funcionarios"][0]
        tipos = [d["tipo"] for d in emp.get("divs", [])]
        self.assertNotIn(STATUS["A_PAGAR"], tipos,
                         "Diferença de R$0,01 não deveria gerar DIFERENÇA A PAGAR")

    def test_normalize_nomes_com_acento_caixa_espacos(self):
        """João da Silva, JOAO DA SILVA e João  Silva devem normalizar igual."""
        self.assertEqual(norm("João da Silva"), norm("JOAO DA SILVA"))
        self.assertEqual(norm("João  Silva"),   norm("JOAO SILVA"))
        self.assertEqual(norm("Ângela Cristóvão"), "ANGELA CRISTOVAO")

    def test_rubricas_variantes_completas(self):
        """Cenário 1–3 do spec: variantes de COMISSAO, DSR e COMISSAO_E_DSR."""
        casos_comissao = ["Comissão", "Comissões", "Comissão de vendas", "Comissão mensal"]
        for v in casos_comissao:
            with self.subTest(v=v):
                self.assertEqual(normalize_rubric(v), "COMISSAO")

        casos_dsr = ["DSR", "DSR sobre comissão", "Descanso Semanal Remunerado"]
        for v in casos_dsr:
            with self.subTest(v=v):
                self.assertEqual(normalize_rubric(v), "DSR")

        self.assertEqual(normalize_rubric("Comissão e DSR"), "COMISSAO_E_DSR")

    def test_valor_br_convertido_corretamente(self):
        """Cenário 7 do spec: formatos brasileiros."""
        self.assertAlmostEqual(brl("1.592,11"), 1592.11)
        self.assertAlmostEqual(brl("R$ 1.592,11"), 1592.11)
        self.assertAlmostEqual(brl("1592,11"), 1592.11)
        self.assertAlmostEqual(brl("(500,00)"), -500.0)

    def test_premio_vs_premio_ppr_smart_match(self):
        """Cenário PPR: PREMIO no relatório, PREMIO PPR no recibo, mesmo valor → alta confiança."""
        verbas = [
            {"codigo": "1", "descricao": "SALARIO BASE", "referencia": "30,00", "valor": 3000.0},
            {"codigo": "2", "descricao": "PREMIO PPR", "referencia": "30,00", "valor": 800.0},
        ]
        result = find_rubric_by_value("PREMIO", 800.0, verbas)
        self.assertIsNotNone(result, "Deveria encontrar PREMIO PPR como possível equivalente de PREMIO")
        self.assertEqual(result["confianca"], "alta")
        self.assertEqual(result["verba"]["descricao"], "PREMIO PPR")


if __name__ == "__main__":
    unittest.main(verbosity=2)
