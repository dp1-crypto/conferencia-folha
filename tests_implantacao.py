"""
Testes do motor de implantacao — leitura generica de holerite e analise
Sigma Contabilidade

Roda com: python3 tests_implantacao.py
"""
import os
import unittest

from app.core.textutils import (
    parse_competencia, competencia_from_filename, competencia_range,
    competencia_label, valid_cpf, valid_cnpj, find_cpf, is_money, is_ref, money,
)
from app.services.holerite_parser import (
    parse_holerites, resolver_grupo, colapsa_siglas, valida_integridade,
    _corrige_classificacao,
)
from app.services.implantacao import analisar, _classifica_serie, chave_funcionario
from app.services.implantacao_export import (
    mapa_rubricas, planilha_importacao, fichas_funcionarios, nome_arquivo,
)
from tests_fixtures import (
    holerite_duas_colunas, holerite_coluna_unica, gerar_amostras,
    FUNC_A, FUNC_B, VERBAS_A, VERBAS_B,
)


# ─────────────────────────────────────────────
# COMPETENCIA
# ─────────────────────────────────────────────

class TestCompetencia(unittest.TestCase):

    def test_rotulo_com_barra(self):
        self.assertEqual(parse_competencia("Competencia: 03/2026"), "2026-03")

    def test_rotulo_mes_extenso(self):
        self.assertEqual(parse_competencia("Referencia MARCO/2026"), "2026-03")
        self.assertEqual(parse_competencia("Folha Mensal - Referencia 03/2026"), "2026-03")

    def test_periodo_usa_mes_inicial(self):
        self.assertEqual(parse_competencia("01/03/2026 a 31/03/2026"), "2026-03")

    def test_data_de_admissao_nao_vira_competencia(self):
        """Com duas datas diferentes e sem rotulo, prefere nao adivinhar."""
        txt = "Admissao 12/03/2019   Demissao 05/08/2022"
        self.assertIsNone(parse_competencia(txt))

    def test_mes_invalido_rejeitado(self):
        self.assertIsNone(parse_competencia("Competencia: 13/2026"))

    def test_nome_de_arquivo(self):
        self.assertEqual(competencia_from_filename("folha_2026-03.pdf"), "2026-03")
        self.assertEqual(competencia_from_filename("holerite 03-2026.pdf"), "2026-03")
        self.assertEqual(competencia_from_filename("mariana_MARCO_2026.pdf"), "2026-03")
        self.assertEqual(competencia_from_filename("holerite-marco-2026.pdf"), "2026-03")
        # Nome sem mes nenhum nao pode ser adivinhado
        self.assertIsNone(competencia_from_filename("mariana.pdf"))

    def test_range_preenche_buracos(self):
        r = competencia_range({"2025-11", "2026-02"})
        self.assertEqual(r, ["2025-11", "2025-12", "2026-01", "2026-02"])

    def test_label(self):
        self.assertEqual(competencia_label("2026-03"), "Marco/2026")


# ─────────────────────────────────────────────
# DOCUMENTOS E VALORES
# ─────────────────────────────────────────────

class TestDocumentos(unittest.TestCase):

    def test_cpf_valido(self):
        self.assertTrue(valid_cpf("529.982.247-25"))
        self.assertFalse(valid_cpf("111.111.111-11"))
        self.assertFalse(valid_cpf("529.982.247-26"))

    def test_cnpj_valido(self):
        self.assertTrue(valid_cnpj("11.222.333/0001-81"))
        self.assertFalse(valid_cnpj("11.222.333/0001-82"))

    def test_find_cpf_ignora_numero_qualquer(self):
        self.assertIsNone(find_cpf("Matricula 123.456.789-00 invalida"))
        self.assertEqual(find_cpf("CPF 529.982.247-25"), "52998224725")


class TestValores(unittest.TestCase):

    def test_is_money(self):
        for v in ["1.234,56", "0,00", "(1.000,00)", "12,00"]:
            self.assertTrue(is_money(v), v)
        for v in ["1234", "3:41", "abc", "12.34"]:
            self.assertFalse(is_money(v), v)

    def test_negativo_por_sinal_a_direita(self):
        """Relatorio de folha costuma imprimir negativo como '100,00-'."""
        self.assertEqual(money("100,00-"), -100.0)
        self.assertEqual(money("(100,00)"), -100.0)

    def test_is_ref_aceita_horas(self):
        self.assertTrue(is_ref("3:41"))
        self.assertTrue(is_ref("8,33"))


# ─────────────────────────────────────────────
# RUBRICAS
# ─────────────────────────────────────────────

class TestResolverGrupo(unittest.TestCase):

    def test_sigla_com_ponto(self):
        self.assertEqual(colapsa_siglas("I N S S"), "INSS")
        self.assertEqual(resolver_grupo("I.N.S.S."), "INSS")
        self.assertEqual(resolver_grupo("D.S.R. SOBRE COMISSAO"), "DSR")

    def test_sufixo_percentual(self):
        self.assertEqual(resolver_grupo("HORAS EXTRAS 50%"), "HORA_EXTRA")
        self.assertEqual(resolver_grupo("HORAS EXTRAS 100"), "HORA_EXTRA")

    def test_contencao_de_palavras(self):
        self.assertEqual(resolver_grupo("PLANO DE SAUDE UNIMED"), "PLANO_SAUDE")
        self.assertEqual(resolver_grupo("ADICIONAL DE INSALUBRIDADE"), "INSALUBRIDADE")

    def test_nao_confunde_adicionais_diferentes(self):
        """Metade das palavras em comum nao pode virar equivalencia."""
        self.assertNotEqual(resolver_grupo("ADICIONAL NOTURNO"),
                            resolver_grupo("ADICIONAL DE TRANSFERENCIA"))

    def test_variante_mais_especifica_vence(self):
        self.assertEqual(resolver_grupo("SALARIO FAMILIA"), "SALARIO_FAMILIA")
        self.assertEqual(resolver_grupo("SALARIO BASE"), "SALARIO")

    def test_desconhecida_volta_normalizada(self):
        self.assertEqual(resolver_grupo("VERBA XPTO INTERNA"), "VERBA XPTO INTERNA")


# ─────────────────────────────────────────────
# INTEGRIDADE
# ─────────────────────────────────────────────

class TestIntegridade(unittest.TestCase):

    def _verbas(self, prov, desc):
        return ([{"tipo": "provento", "valor": v} for v in prov]
                + [{"tipo": "desconto", "valor": v} for v in desc])

    def test_fecha_certo(self):
        r = valida_integridade(
            self._verbas([1000.0, 200.0], [150.0]),
            {"total_proventos": 1200.0, "total_descontos": 150.0, "liquido": 1050.0},
        )
        self.assertTrue(r["ok"])
        self.assertEqual(r["nivel"], "ok")

    def test_aponta_erro_quando_nao_fecha(self):
        r = valida_integridade(
            self._verbas([1000.0], [150.0]),
            {"total_proventos": 1200.0, "total_descontos": 150.0, "liquido": 1050.0},
        )
        self.assertFalse(r["ok"])
        self.assertEqual(r["nivel"], "erro")
        self.assertTrue(r["mensagens"])

    def test_sem_totais_impressos_nao_finge_que_conferiu(self):
        r = valida_integridade(self._verbas([1000.0], [150.0]), {})
        self.assertEqual(r["nivel"], "sem_referencia")
        self.assertFalse(r["ok"])

    def test_centavos_dentro_da_tolerancia(self):
        r = valida_integridade(
            self._verbas([1000.01], [150.0]),
            {"total_proventos": 1000.0, "total_descontos": 150.0, "liquido": 850.0},
        )
        self.assertTrue(r["ok"])


class TestCorrecaoClassificacao(unittest.TestCase):

    def test_inverte_rubrica_para_fechar_total(self):
        verbas = [
            {"tipo": "provento", "tipo_origem": "nome", "valor": 1000.0, "descricao": "SALARIO"},
            {"tipo": "provento", "tipo_origem": "nome", "valor": 200.0, "descricao": "INSS"},
        ]
        _corrige_classificacao(verbas, 1000.0, 200.0)
        self.assertEqual(verbas[1]["tipo"], "desconto")
        self.assertEqual(verbas[1]["tipo_origem"], "corrigido")

    def test_indefinido_vira_provento(self):
        verbas = [{"tipo": "indefinido", "tipo_origem": "nome", "valor": 500.0, "descricao": "XPTO"}]
        _corrige_classificacao(verbas, 500.0, 0.0)
        self.assertEqual(verbas[0]["tipo"], "provento")


# ─────────────────────────────────────────────
# CLASSIFICACAO DE SERIE
# ─────────────────────────────────────────────

class TestClassificaSerie(unittest.TestCase):

    MESES = ["2026-01", "2026-02", "2026-03", "2026-04"]

    def test_fixa(self):
        vals = {m: 450.0 for m in self.MESES}
        self.assertEqual(_classifica_serie(vals, self.MESES)["classe"], "fixa")

    def test_fixa_reajustada(self):
        vals = {"2026-01": 1800.0, "2026-02": 1800.0, "2026-03": 1980.0, "2026-04": 1980.0}
        self.assertEqual(_classifica_serie(vals, self.MESES)["classe"], "fixa_reajustada")

    def test_variavel(self):
        vals = {"2026-01": 100.0, "2026-02": 250.0, "2026-03": 80.0, "2026-04": 990.0}
        self.assertEqual(_classifica_serie(vals, self.MESES)["classe"], "variavel")

    def test_eventual(self):
        vals = {"2026-02": 300.0}
        self.assertEqual(_classifica_serie(vals, self.MESES)["classe"], "eventual")


class TestChaveFuncionario(unittest.TestCase):

    def test_cpf_tem_prioridade(self):
        r = {"cpf": "52998224725", "matricula": "137", "nome_norm": "MARIANA"}
        self.assertEqual(chave_funcionario(r), "cpf:52998224725")

    def test_homonimos_nao_colidem_com_cpf(self):
        a = {"cpf": "52998224725", "matricula": "", "nome_norm": "JOSE SILVA"}
        b = {"cpf": "11144477735", "matricula": "", "nome_norm": "JOSE SILVA"}
        self.assertNotEqual(chave_funcionario(a), chave_funcionario(b))


# ─────────────────────────────────────────────
# EXTRACAO PONTA A PONTA
# ─────────────────────────────────────────────

class TestExtracaoPDF(unittest.TestCase):

    def test_todos_os_layouts_fecham_a_conta(self):
        """O criterio de aceite do motor: integridade OK em todo layout."""
        for nome, raw in gerar_amostras().items():
            with self.subTest(layout=nome):
                r = parse_holerites(raw, nome)
                self.assertTrue(r["recibos"], f"{nome}: nenhum recibo lido")
                for rec in r["recibos"]:
                    self.assertTrue(
                        rec["integridade"]["ok"],
                        f"{nome} / {rec['nome']}: " + "; ".join(rec["integridade"]["mensagens"]),
                    )

    def test_identifica_funcionario(self):
        raw = holerite_duas_colunas(FUNC_A, VERBAS_A)
        rec = parse_holerites(raw, "t.pdf")["recibos"][0]
        self.assertEqual(rec["nome"], "Mariana Ribeiro Alves")
        self.assertEqual(rec["cpf"], "52998224725")
        self.assertEqual(rec["matricula"], "137")
        self.assertGreaterEqual(rec["confianca_nome"], 0.7)

    def test_separa_provento_de_desconto_por_coluna(self):
        raw = holerite_duas_colunas(FUNC_A, VERBAS_A)
        rec = parse_holerites(raw, "t.pdf")["recibos"][0]
        tipos = {v["descricao"]: v["tipo"] for v in rec["verbas"]}
        self.assertEqual(tipos["SALARIO BASE"], "provento")
        self.assertEqual(tipos["INSS"], "desconto")
        self.assertEqual(tipos["VALE TRANSPORTE"], "desconto")

    def test_coluna_unica_usa_o_nome_da_rubrica(self):
        raw = holerite_coluna_unica(FUNC_B, VERBAS_B)
        rec = parse_holerites(raw, "t.pdf")["recibos"][0]
        tipos = {v["descricao"]: v["tipo"] for v in rec["verbas"]}
        self.assertEqual(tipos["I.N.S.S"], "desconto")
        self.assertEqual(tipos["COMISSAO SOBRE VENDAS"], "provento")

    def test_rubrica_abreviada_nao_e_descartada(self):
        """'I.N.S.S.' vira 'I N S S' ao normalizar — nao pode sumir."""
        raw = holerite_duas_colunas(FUNC_B, VERBAS_B)
        rec = parse_holerites(raw, "t.pdf")["recibos"][0]
        grupos = [resolver_grupo(v["descricao"]) for v in rec["verbas"]]
        self.assertIn("INSS", grupos)

    def test_extrai_competencia_do_documento(self):
        raw = holerite_duas_colunas(FUNC_A, VERBAS_A, comp="07/2025")
        r = parse_holerites(raw, "sem_data_no_nome.pdf")
        self.assertEqual(r["competencia_documento"], "2025-07")

    def test_detecta_layout_pelo_fingerprint(self):
        r = parse_holerites(holerite_coluna_unica(FUNC_A, VERBAS_A), "t.pdf")
        self.assertEqual(r["layout"]["id"], "dominio")
        r2 = parse_holerites(holerite_duas_colunas(FUNC_A, VERBAS_A), "t.pdf")
        self.assertEqual(r2["layout"]["id"], "questor")

    def test_tipo_recibo_nao_confunde_com_rubrica(self):
        """Rubrica 'ADIANTAMENTO SALARIAL' no corpo nao muda o tipo do recibo."""
        rec = parse_holerites(holerite_duas_colunas(FUNC_B, VERBAS_B), "t.pdf")["recibos"][0]
        self.assertEqual(rec["tipo"], "mensal")


# ─────────────────────────────────────────────
# ANALISE E EXPORTS
# ─────────────────────────────────────────────

class TestAnalise(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        COMPS = ["10/2025", "11/2025", "12/2025", "01/2026"]
        docs = []
        for i, comp in enumerate(COMPS):
            verbas = list(VERBAS_A)
            if i >= 2:   # rubrica nova a partir do 3o mes
                verbas = verbas + [("905", "ADICIONAL DE INSALUBRIDADE", "20,00", 303.60, "P")]
            docs.append(parse_holerites(
                holerite_duas_colunas(FUNC_A, verbas, comp=comp), f"a_{comp}.pdf"))
            if i < 2:    # segundo funcionario sai no meio do periodo
                docs.append(parse_holerites(
                    holerite_duas_colunas(FUNC_B, VERBAS_B, comp=comp), f"b_{comp}.pdf"))
        cls.analise = analisar(docs)

    def test_periodo_e_funcionarios(self):
        r = self.analise["resumo"]
        self.assertEqual(r["meses"], 4)
        self.assertEqual(r["funcionarios"], 2)

    def test_todos_os_recibos_conferem(self):
        q = self.analise["qualidade"]
        self.assertEqual(q["recibos_ok"], q["recibos"])

    def test_classifica_fixas_e_variaveis(self):
        # A rubrica e identificada pelo CODIGO de origem (unico dentro de uma
        # mesma empresa), nao pela descricao — que sofre quebra de texto no PDF.
        por_codigo = {}
        for c in self.analise["catalogo_rubricas"]:
            for cod in c["codigos_origem"]:
                por_codigo[cod] = c
        self.assertEqual(por_codigo["1"]["classe"], "fixa")        # salario
        self.assertEqual(por_codigo["9201"]["classe"], "calculada")  # INSS
        self.assertEqual(por_codigo["9301"]["classe"], "fixa")     # vale transporte

    def test_detecta_saida_e_rubrica_nova(self):
        tipos = {e["tipo"] for e in self.analise["eventos"]}
        self.assertIn("saida", tipos)
        self.assertIn("rubrica_nova", tipos)

    def test_inss_nao_entra_como_rubrica_a_lancar(self):
        r = self.analise["resumo"]
        calculadas = [c for c in self.analise["catalogo_rubricas"] if c["classe"] == "calculada"]
        self.assertTrue(calculadas)
        self.assertEqual(r["rubricas_calculadas"], len(calculadas))

    def test_exports_geram_arquivo_valido(self):
        m = mapa_rubricas(self.analise)
        p = planilha_importacao(self.analise)
        f = fichas_funcionarios(self.analise)
        self.assertTrue(m.startswith(b"PK"), "mapa nao e xlsx")
        self.assertTrue(p.startswith(b"PK"), "planilha nao e xlsx")
        self.assertTrue(f.startswith(b"%PDF"), "ficha nao e pdf")
        self.assertGreater(len(m), 3000)
        self.assertGreater(len(f), 2000)

    def test_planilha_tem_as_quatro_abas(self):
        import openpyxl
        import io
        wb = openpyxl.load_workbook(io.BytesIO(planilha_importacao(self.analise)))
        self.assertEqual(wb.sheetnames,
                         ["Cadastro", "Eventos fixos", "Eventos variaveis", "Resumo mensal"])

    def test_nome_arquivo_previsivel(self):
        n = nome_arquivo(self.analise, "mapa-rubricas", "xlsx")
        self.assertTrue(n.startswith("mapa-rubricas_"))
        self.assertTrue(n.endswith(".xlsx"))


# ─────────────────────────────────────────────
# REGRESSAO COM FOLHA REAL
# ─────────────────────────────────────────────

class TestAmostrasReais(unittest.TestCase):
    """
    Roda contra os PDFs reais em amostras/ quando existirem.

    Esses arquivos NAO vao para o repositorio (dado de folha de cliente, LGPD)
    — a pasta esta no .gitignore. Sem eles o teste e pulado, entao a suite
    continua verde em qualquer maquina.

    O criterio de aceite e o mais duro possivel: a soma do que foi extraido
    tem que bater com o RESUMO GERAL impresso pelo proprio relatorio, tanto em
    valor quanto em quantidade de funcionarios.
    """

    @classmethod
    def setUpClass(cls):
        import glob
        import os
        base = os.path.join(os.path.dirname(__file__), "amostras")
        cls.arquivos = sorted(glob.glob(os.path.join(base, "*.pdf")))
        if not cls.arquivos:
            raise unittest.SkipTest("amostras/ sem PDFs reais — teste pulado")

    def test_todo_recibo_fecha_com_os_totais_impressos(self):
        for caminho in self.arquivos:
            with self.subTest(arquivo=os.path.basename(caminho)):
                r = parse_holerites(open(caminho, "rb").read(), os.path.basename(caminho))
                self.assertTrue(r["recibos"], "nenhum recibo lido")
                ruins = [x for x in r["recibos"] if not x["integridade"]["ok"]]
                self.assertEqual(
                    ruins, [],
                    "\n".join(f"{x['nome']}: " + "; ".join(x["integridade"]["mensagens"])
                              for x in ruins[:5]),
                )

    def test_soma_bate_com_o_resumo_do_relatorio(self):
        for caminho in self.arquivos:
            with self.subTest(arquivo=os.path.basename(caminho)):
                r = parse_holerites(open(caminho, "rb").read(), os.path.basename(caminho))
                resumos = r.get("resumos_empresa") or []
                if not resumos:
                    self.skipTest("relatorio sem resumo geral")
                res = resumos[0]
                soma_p = round(sum(x["integridade"]["soma_proventos"] for x in r["recibos"]), 2)
                soma_d = round(sum(x["integridade"]["soma_descontos"] for x in r["recibos"]), 2)
                self.assertAlmostEqual(soma_p, res["proventos"], places=2,
                                       msg="proventos extraidos != resumo do relatorio")
                self.assertAlmostEqual(soma_d, res["descontos"], places=2,
                                       msg="descontos extraidos != resumo do relatorio")
                self.assertEqual(len(r["recibos"]), res["funcionarios"],
                                 "quantidade de funcionarios != resumo do relatorio")

    def test_competencia_nao_inventa_meses(self):
        """Data de admissao dentro do bloco nao pode virar competencia."""
        for caminho in self.arquivos:
            with self.subTest(arquivo=os.path.basename(caminho)):
                r = parse_holerites(open(caminho, "rb").read(), os.path.basename(caminho))
                comps = {x["competencia"] for x in r["recibos"]}
                self.assertEqual(
                    comps, {r["competencia_documento"]},
                    f"recibos com competencia diferente da do documento: {comps}",
                )


if __name__ == "__main__":
    unittest.main(verbosity=2)
