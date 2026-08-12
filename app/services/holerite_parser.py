"""
Extrator generico de holerite / recibo de pagamento
Sigma Contabilidade

Le contracheque de QUALQUER sistema de folha usando a posicao das palavras,
nao um regex fixo. Todo recibo extraido passa por validacao de integridade
(proventos - descontos = liquido impresso). Quando a conta nao fecha, o
recibo e marcado como leitura duvidosa em vez de entregar numero errado
com cara de certo.
"""
import re
from functools import lru_cache

from app.core.textutils import (
    flat, deaccent, money, is_money, is_ref, find_cpf, find_cnpj,
    parse_competencia, competencia_from_filename, only_digits,
)
from app.core.config import TOLERANCIA_CENTAVOS
from app.services import rubrics
from app.services import layout_profiles
from app.services.pdf_grid import (
    read_pdf, group_rows, row_text, money_columns, col_of,
    split_row, value_column_labels, words_in, detect_secoes,
)

# ─────────────────────────────────────────────
# DICIONARIOS DE APOIO
# ─────────────────────────────────────────────

# Rubricas que sao desconto em qualquer sistema (usado quando o layout tem
# coluna unica de valor e nao da para separar por posicao).
DESCONTO_KW = [
    "INSS", "PREVIDENCIA", "IRRF", "IRPF", "IMPOSTO DE RENDA", "I R R F",
    "VALE TRANSPORTE", "VT ", "TRANSPORTE", "VALE REFEICAO", "VALE ALIMENTACAO",
    "ADIANTAMENTO", "ADTO", "PLANO DE SAUDE", "UNIMED", "ODONTO", "AMIL", "HAPVIDA",
    "FARMACIA", "EMPRESTIMO", "CONSIGNADO", "PENSAO", "ALIMENTOS",
    "CONTRIBUICAO", "MENSALIDADE", "SINDICAL", "ASSISTENCIAL", "CONFEDERATIVA",
    "SEGURO", "FALTA", "ATRASO", "DESCONTO", "DEVOLUCAO", "CO PARTICIPACAO",
    "COPARTICIPACAO", "MULTA", "ARREDONDAMENTO", "REEMBOLSO", "CESTA",
    "CONVENIO", "UNIODONTO", "PREVIDENCIA PRIVADA", "PGBL", "ADIANT",
]

# Linhas que aparecem na area da tabela mas nao sao verba.
# Casa por INICIO da descricao, nunca por trecho solto: 'BASE' como substring
# eliminaria 'SALARIO BASE', e 'FUNCAO' eliminaria 'GRATIFICACAO DE FUNCAO'.
RX_NAO_VERBA = re.compile(
    r"^("
    r"BASE(S)?\b|SAL(ARIO)? CONTR|SALARIO CONTRIBUICAO|FGTS\b|F G T S|"
    r"CNPJ\b|CPF\b|TOTAL\b|LIQUIDO\b|VALOR LIQUIDO|DEPOSITO\b|ASSINATURA|"
    r"DECLARO|RECEBI\b|MENSAGEM|OBSERVAC|COMPETENCIA|ADMISSAO|FILIAL\b|"
    r"EMPRESA\b|RECIBO\b|DEMONSTRATIVO|CODIGO\b|COD\b|DESCRICAO\b|REFERENCIA\b|"
    r"IMPRESSO EM|PAGINA\b|CARGO\b|FUNCAO\b|DEPARTAMENTO|CBO\b|MATRICULA|"
    r"NOME\b|FUNCIONARIO|COLABORADOR|EMPREGADO\b|BANCO\b|AGENCIA|CONTA\b|"
    r"VENCIMENTOS?$|PROVENTOS?$|DESCONTOS?$|"
    # Linhas de rodape do espelho de folha (SCI e similares)
    r"FOLHA\b|RAIS\b|ADMITIDO|DEMITIDO|AFASTADO|SALARIO BASE ->"
    r")"
)

# Chaves de rotulo que carregam VALOR monetario. As demais (funcao, matricula,
# admissao) sao texto e nao podem entrar na varredura de valores — 'FUNCAO'
# capturaria o valor da rubrica 'GRATIFICACAO DE FUNCAO'.
ROTULOS_MONETARIOS = {
    "total_proventos", "total_descontos", "liquido",
    "base_inss", "base_fgts", "fgts_mes", "base_irrf",
}

TIPO_RECIBO = [
    ("rescisao",        r"RESCIS|TRCT|TERMO DE (QUITACAO|RESCIS)|HOMOLOG"),
    ("13_adiantamento", r"13.{0,12}(ADIANT|1.?\s*PARCELA|PRIMEIRA)|ADIANT.{0,12}13|GRATIFICACAO NATALINA.{0,20}ADIANT"),
    ("13_segunda",      r"13.{0,12}(2.?\s*PARCELA|SEGUNDA|INTEGRAL|COMPLEMENT)|GRATIFICACAO NATALINA"),
    ("ferias",          r"\bFERIAS\b|AVISO DE FERIAS|RECIBO DE FERIAS"),
    ("adiantamento",    r"ADIANTAMENTO (SALARIAL|DE SALARIO|QUINZENAL)|VALE SALARIO"),
    ("plr",             r"\bPLR\b|PARTICIPACAO NOS? (LUCROS|RESULTADOS)"),
]


def _mapa_rotulos() -> dict:
    """Rotulos MONETARIOS do layouts-folha.json, normalizados e em tokens."""
    r = layout_profiles.rotulos()
    mapa = {}
    for chave, lista in r.items():
        if chave not in ROTULOS_MONETARIOS:
            continue
        mapa[chave] = [flat(x).split() for x in lista if x]
    return mapa


# ─────────────────────────────────────────────
# LEITURA DE ROTULOS COM VALOR
# ─────────────────────────────────────────────

def _pos_label(words_flat: list, tokens: list) -> int:
    """Indice da ULTIMA palavra do rotulo dentro da linha, ou -1."""
    n = len(tokens)
    if n == 0 or n > len(words_flat):
        return -1
    for i in range(len(words_flat) - n + 1):
        if words_flat[i:i + n] == tokens:
            return i + n - 1
    return -1


def extrai_rotulados(rows: list, mapa: dict) -> dict:
    """
    Varre as linhas procurando 'ROTULO ... valor' e devolve:
      {chave: {"valor": float, "x1": float, "row": int, "rotulo": str}}

    Rotulos mais longos tem prioridade (TOTAL DE DESCONTOS vence DESCONTOS),
    e cada chave so e preenchida uma vez — a primeira ocorrencia confiavel.
    """
    achados = {}

    # (chave, tokens) ordenados do rotulo mais especifico para o mais generico
    pares = [(c, t) for c, lst in mapa.items() for t in lst]
    pares.sort(key=lambda p: -len(p[1]))

    for idx, row in enumerate(rows):
        wf = [flat(w["text"]) for w in row]
        tem_dinheiro = any(is_money(w["text"]) for w in row)
        # Cabecalho da tabela ("Cod Descricao Referencia Vencimentos Descontos"):
        # os rotulos ali nomeiam COLUNAS, nao totais. Sem esse corte, 'VENCIMENTOS'
        # do cabecalho puxaria o valor da primeira verba como se fosse o total.
        eh_cabecalho = bool(
            re.search(r"\b(DESCRICAO|REFERENCIA|CODIGO|COD)\b", " ".join(wf))
            and not tem_dinheiro
        )
        usados = set()
        for chave, tokens in pares:
            if chave in achados:
                continue
            pos = _pos_label(wf, tokens)
            if pos < 0 or any(p in usados for p in range(pos - len(tokens) + 1, pos + 1)):
                continue
            if eh_cabecalho:
                continue

            # Valor: primeiro token monetario a direita do rotulo, na mesma linha
            alvo = None
            for j in range(pos + 1, len(row)):
                if is_money(row[j]["text"]):
                    alvo = row[j]
                    break

            # Valor abaixo do rotulo (layout em caixas).
            # So para rotulo especifico (2+ palavras) e em linha sem valor —
            # rotulo generico de uma palavra so daria falso positivo.
            if alvo is None and len(tokens) >= 2 and not tem_dinheiro:
                lx0 = row[pos]["x0"] - 15
                lx1 = row[pos]["x1"] + 90
                for prox in rows[idx + 1: idx + 3]:
                    cand = [w for w in prox
                            if is_money(w["text"]) and lx0 <= w["x1"] <= lx1]
                    if cand:
                        alvo = cand[0]
                        break

            if alvo is None:
                continue

            achados[chave] = {
                "valor": money(alvo["text"]),
                "x1": alvo["x1"],
                "row": idx,
                "rotulo": " ".join(tokens),
            }
            usados.update(range(pos - len(tokens) + 1, pos + 1))

    return achados


# ─────────────────────────────────────────────
# SEPARACAO EM RECIBOS
# ─────────────────────────────────────────────

def compacto(texto: str) -> str:
    """
    Texto sem acento, sem espaco e em maiuscula.

    A extracao de PDF quebra palavra no meio sem aviso ('Total de pro v entos',
    'Admiti d o', 'Seg.de Vida e m Grupo'). Toda deteccao ESTRUTURAL — fim de
    bloco, rodape, resumo — compara nesta forma, senao uma quebra invisivel
    funde dois funcionarios num recibo so.
    """
    return re.sub(r"\s+", "", deaccent(texto)).upper()


_RX_LIQUIDO = re.compile(
    r"VALORLIQUIDO|LIQUIDOARECEBER|TOTALLIQUIDO|LIQUIDODAFOLHA|VALORARECEBER"
)

_RX_TOTAIS = re.compile(r"TOTAL(?:DE)?(PROVENTOS|VENCIMENTOS)")


def split_blocos(rows: list) -> list:
    """
    Quebra a pagina em blocos, um por recibo/funcionario.

    Holerite termina no 'valor liquido'. Ja o espelho de folha (relatorio
    analitico) nao imprime esse rotulo — fecha cada funcionario na linha
    'Total de proventos -> X  Total de descontos -> Y'. Por isso ha dois
    criterios de fecho, e a quebra e feita pelo fim do bloco, nao pelo inicio:
    a mesma pagina pode trazer duas vias do recibo ou dezenas de funcionarios.
    """
    fechos = [i for i, r in enumerate(rows)
              if _RX_LIQUIDO.search(compacto(row_text(r)))]
    sobra = 4          # holerite: linhas de assinatura/bases apos o liquido

    if not fechos:
        fechos = [i for i, r in enumerate(rows)
                  if _RX_TOTAIS.search(compacto(row_text(r)))]
        sobra = 2      # espelho: so a linha de bases/liquido logo abaixo

    if not fechos:
        return [rows] if rows else []

    blocos, ini = [], 0
    for f in fechos:
        fim = min(len(rows), f + sobra)
        bloco = rows[ini:fim]
        if len(bloco) >= 3:
            blocos.append(bloco)
        ini = fim
    if ini < len(rows) and len(rows) - ini >= 5:
        blocos.append(rows[ini:])
    return blocos


# ─────────────────────────────────────────────
# IDENTIFICACAO DO FUNCIONARIO
# ─────────────────────────────────────────────

_RX_NOME_VALIDO = re.compile(r"^[A-ZÀ-Ü][A-ZÀ-Üa-zà-ü' ]{4,60}$")
_STOP_NOME = {
    "TOTAL", "VALOR", "LIQUIDO", "DESCONTOS", "VENCIMENTOS", "PROVENTOS",
    "EMPRESA", "RECIBO", "FOLHA", "SALARIO", "BASE", "CODIGO", "DESCRICAO",
    "REFERENCIA", "COMPETENCIA", "PAGAMENTO", "DEMONSTRATIVO", "MENSAL",
    "CNPJ", "CPF", "ENDERECO", "MATRICULA", "FUNCIONARIO", "COLABORADOR",
    "DEPARTAMENTO", "FUNCAO", "CARGO", "ADMISSAO", "PAGINA", "LTDA", "EIRELI",
    "ME", "EPP", "MEI", "SA", "S A", "MUNICIPIO", "BANCO", "AGENCIA", "CONTA",
}


def _parece_nome(txt: str) -> bool:
    t = deaccent(txt or "").strip()
    if not (8 <= len(t) <= 60):
        return False
    palavras = t.split()
    if len(palavras) < 2:
        return False
    if any(p.upper() in _STOP_NOME for p in palavras):
        return False
    if re.search(r"\d", t):
        return False
    return bool(re.fullmatch(r"[A-Za-z' ]+", t))


def identifica_funcionario(rows: list, perfil: dict) -> dict:
    """
    Descobre quem e o funcionario do bloco, tentando varias estrategias e
    informando o nivel de confianca em vez de fingir certeza.

    Retorna: {nome, matricula, cpf, cbo, funcao, admissao, confianca, metodo}
    """
    texto = "\n".join(row_text(r) for r in rows)
    tflat = deaccent(texto)
    out = {
        "nome": "", "matricula": "", "cpf": "", "cbo": "",
        "funcao": "", "admissao": "", "confianca": 0.0, "metodo": "",
    }

    # CPF do funcionario (ignora o primeiro se for o da empresa — CNPJ ja filtrado)
    cpf = find_cpf(texto)
    if cpf:
        out["cpf"] = cpf

    # Admissao
    m = re.search(r"ADMISS[AÃ]O[^\d]{0,15}(\d{2}[/.-]\d{2}[/.-]\d{2,4})", tflat, re.I)
    if m:
        out["admissao"] = m.group(1).replace(".", "/").replace("-", "/")

    # CBO (6 digitos, as vezes com hifen)
    m = re.search(r"\bCBO[^\d]{0,10}(\d{4,6}[-]?\d?)\b", tflat, re.I)
    if m:
        out["cbo"] = only_digits(m.group(1))

    # ── Estrategia 1: regex do perfil conhecido ──────────────────────────
    rx_perfil = perfil.get("nome_regex")
    if rx_perfil:
        m = re.search(rx_perfil, texto)
        if m and _parece_nome(m.group(1)):
            out.update({"nome": m.group(1).strip(), "confianca": 0.95, "metodo": "perfil"})

    # ── Estrategia 2: rotulo explicito na mesma linha ────────────────────
    if not out["nome"]:
        RX = re.compile(
            r"(?:NOME DO (?:FUNCION[AÁ]RIO|EMPREGADO|COLABORADOR)|"
            r"FUNCION[AÁ]RIO|COLABORADOR|EMPREGADO|SERVIDOR|NOME)\s*[:\-]?\s+"
            r"([A-ZÀ-Ü][A-ZÀ-Üa-zà-ü' ]{6,60})", re.I
        )
        for row in rows[:18]:
            m = RX.search(deaccent(row_text(row)))
            if m:
                cand = re.split(r"\s{2,}", m.group(1).strip())[0].strip()
                cand = re.sub(r"\s+(CBO|FUNCAO|CARGO|ADMISSAO|MATRICULA).*$", "", cand, flags=re.I)
                if _parece_nome(cand):
                    out.update({"nome": cand, "confianca": 0.9, "metodo": "rotulo"})
                    break

    # ── Estrategia 2b: matricula + NOME na linha de cadastro ─────────────
    # Espelho de folha: "92 FULANO DE TAL DA SILVA 0 0 Admitido em 18/03/2026".
    # Exige o marcador de admissao na propria linha — sem ele, a verba
    # '120 HORAS EXTRAS 50% 8:30' seria lida como matricula + nome.
    if not out["nome"]:
        for row in rows[:12]:
            linha = deaccent(row_text(row)).upper()
            if not re.search(r"ADMITID|ADMISS|DEMITID|PRO-?LABORE", compacto(linha)):
                continue
            m = re.match(r"^(\d{1,6})\s+([A-ZÀ-Ü][A-ZÀ-Ü' ]{5,50}?)\s+(?=\d|ADMITID)", linha)
            if m and _parece_nome(m.group(2)):
                out.update({"nome": m.group(2).strip(), "matricula": m.group(1),
                            "confianca": 0.85, "metodo": "matricula_nome"})
                break

    # ── Estrategia 3: matricula + NOME + codigo (Dominio e similares) ────
    if not out["nome"]:
        m = re.search(r"\b(\d{1,5})\s+([A-ZÀ-Ü][A-ZÀ-Ü ]{6,50}?)\s+(\d{4,6})\b", texto)
        if m and _parece_nome(m.group(2)):
            out.update({
                "nome": m.group(2).strip(), "matricula": m.group(1),
                "cbo": out["cbo"] or m.group(3), "confianca": 0.8, "metodo": "matricula_cbo",
            })

    # ── Estrategia 4: linha do CPF — o trecho alfabetico e o nome ────────
    if not out["nome"] and cpf:
        for row in rows[:25]:
            txt = row_text(row)
            if only_digits(txt).find(cpf) >= 0 or find_cpf(txt) == cpf:
                for pedaco in re.split(r"[\d/.:\-]{3,}", txt):
                    pedaco = pedaco.strip()
                    if _parece_nome(pedaco):
                        out.update({"nome": pedaco, "confianca": 0.7, "metodo": "linha_cpf"})
                        break
            if out["nome"]:
                break

    # ── Estrategia 5: primeira linha que parece nome proprio ─────────────
    if not out["nome"]:
        for row in rows[:20]:
            for pedaco in re.split(r"\s{3,}|[|]", row_text(row)):
                pedaco = re.sub(r"^\d+\s*", "", pedaco.strip())
                if _parece_nome(pedaco) and len(pedaco.split()) >= 2:
                    out.update({"nome": pedaco, "confianca": 0.45, "metodo": "heuristica"})
                    break
            if out["nome"]:
                break

    # Matricula por rotulo
    if not out["matricula"]:
        m = re.search(r"(?:MATR[IÍ]CULA|REGISTRO|CHAPA|C[OÓ]D(?:IGO)?)\s*[:\-]?\s*(\d{1,8})",
                      tflat, re.I)
        if m:
            out["matricula"] = m.group(1)

    # Funcao / cargo
    m = re.search(r"(?:FUN[CÇ][AÃ]O|CARGO)\s*[:\-]?\s+([A-ZÀ-Ü][A-ZÀ-Üa-zà-ü /]{3,40})",
                  deaccent(texto), re.I)
    if m:
        out["funcao"] = m.group(1).strip()

    out["nome"] = " ".join(out["nome"].split()).title() if out["nome"] else ""
    return out


# ─────────────────────────────────────────────
# VERBAS
# ─────────────────────────────────────────────

_RX_SIGLA = re.compile(r"\b(?:[A-Z]\s+){1,}[A-Z]\b")
_RX_SUFIXO_NUM = re.compile(r"\s+\d+([.,]\d+)?\s*%?$")
_STOP = {"DE", "DA", "DO", "DOS", "DAS", "E", "A", "O", "S", "EM", "NO", "NA",
         "NOS", "NAS", "SOBRE", "COM", "POR", "PARA", "AO", "AOS"}


def _palavras(t: str) -> set:
    """
    Palavras significativas de uma rubrica ja normalizada.

    Numero entra sempre, mesmo com 1 digito: em folha o numero E o que
    distingue ('13 SALARIO' x 'SALARIO', '1/3 FERIAS' x '1/12 FERIAS').
    Sigla de 2 letras tambem entra — descartar o 'VT' de 'DESC VT' deixava a
    variante valendo so 'DESC', e ela passava a casar com qualquer desconto.
    Fica de fora apenas conectivo e letra solta (lixo de quebra de texto).
    """
    return {w for w in t.split()
            if w not in _STOP and (w.isdigit() or len(w) >= 2)}


def colapsa_siglas(texto: str) -> str:
    """'I.N.S.S.' normaliza para 'I N S S'; aqui volta a virar 'INSS'."""
    return _RX_SIGLA.sub(lambda m: m.group(0).replace(" ", ""), texto)


@lru_cache(maxsize=4096)
def resolver_grupo(descricao: str) -> str:
    """
    Grupo canonico da rubrica, com tres tentativas antes de desistir:

      1. match exato no dicionario ('COMISSAO DE VENDAS' -> COMISSAO);
      2. match exato apos colapsar sigla e tirar sufixo numerico
         ('HORAS EXTRAS 50%' -> 'HORAS EXTRAS');
      3. maior sobreposicao de palavras com as variantes cadastradas
         ('ADICIONAL DE INSALUBRIDADE' -> INSALUBRIDADE).

    Sem isso, cada variacao de escrita de cada escritorio viraria uma rubrica
    "sem equivalencia" e o mapa de importacao ficaria cheio de pendencia falsa.
    """
    if not descricao:
        return ""

    base = colapsa_siglas(flat(descricao))
    sem_sufixo = _RX_SUFIXO_NUM.sub("", base).strip()

    for cand in (descricao, base, sem_sufixo):
        g = rubrics.normalize_rubric(cand)
        if g in rubrics.RUBRIC_META:
            return g

    # Contencao de palavras: a variante vale se TODAS as palavras significativas
    # dela aparecem na descricao. Vence a variante mais especifica (mais palavras).
    # Similaridade percentual seria perigosa aqui — 'ADICIONAL NOTURNO' e
    # 'ADICIONAL DE TRANSFERENCIA' dividem metade das palavras e nao sao a
    # mesma rubrica; contencao nunca casa as duas.
    alvo = _palavras(sem_sufixo or base)
    if not alvo:
        return sem_sufixo or base

    melhor, melhor_peso = "", 0
    for grupo, variantes in rubrics.RUBRIC_GROUPS.items():
        for v in variantes:
            pv = _palavras(colapsa_siglas(flat(v)))
            if pv and pv <= alvo and len(pv) > melhor_peso:
                melhor, melhor_peso = grupo, len(pv)
    return melhor or (sem_sufixo or base)


def _tipo_por_nome(descricao: str) -> str:
    """Classifica a rubrica em provento/desconto pelo nome. 'indefinido' quando nao souber."""
    meta = rubrics.RUBRIC_META.get(resolver_grupo(descricao))
    if meta and meta.get("tipo") in ("provento", "desconto"):
        return meta["tipo"]
    d = colapsa_siglas(flat(descricao))
    for kw in DESCONTO_KW:
        if kw.strip() in d:
            return "desconto"
    return "indefinido"


def _eh_linha_verba(sr: dict, papeis: list) -> bool:
    """Uma linha e verba se tem descricao com letra e ao menos um valor de verba."""
    desc = flat(sr["descricao"])
    # Conta letras no total, nao letras seguidas: rubrica abreviada com ponto
    # ('I.N.S.S.', 'D.S.R.') vira 'I N S S' ao normalizar e seria descartada.
    if not desc or len(re.sub(r"[^A-Z]", "", desc)) < 2:
        return False
    if RX_NAO_VERBA.match(desc):
        return False
    tem_valor = any(
        papeis[ci] in ("provento", "desconto", "valor")
        for ci in sr["valores"] if ci < len(papeis)
    )
    return tem_valor


def extrai_verbas(rows: list, faixas: list, papeis: list) -> list:
    """
    Le as linhas de verba do bloco. A coluna em que o valor foi impresso
    define provento x desconto quando o layout tem duas colunas.
    """
    verbas = []
    for row in rows:
        sr = split_row(row, faixas)
        if not _eh_linha_verba(sr, papeis):
            continue

        referencia = sr["referencia"]
        for ci, val in sorted(sr["valores"].items()):
            if ci >= len(papeis):
                continue
            papel = papeis[ci]
            if papel == "referencia":
                if not referencia:
                    referencia = f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                continue
            if papel not in ("provento", "desconto", "valor"):
                continue
            if abs(val) < 0.01:
                continue

            desc = sr["descricao"]
            tipo = papel if papel in ("provento", "desconto") else _tipo_por_nome(desc)
            verbas.append({
                "codigo": sr["codigo"],
                "descricao": desc,
                "grupo": resolver_grupo(desc),
                "referencia": referencia,
                "valor": abs(val),
                "tipo": tipo,
                "tipo_origem": "coluna" if papel in ("provento", "desconto") else "nome",
            })
    return verbas


def extrai_verbas_secoes(bloco: list, secoes: list, totais: dict = None) -> list:
    """
    Extrai verbas de layout com tabelas lado a lado (espelho / analitico).

    Cada secao e uma tabela independente com suas proprias colunas, e o papel
    da secao (proventos ou descontos) ja define o lado da verba — nao depende
    do nome da rubrica nem de adivinhacao.
    """
    # Linhas que atravessam as duas secoes e nao sao verba: cabecalho do
    # funcionario ('92 FULANO ... Admitido em 18/03/2026 Salario base -> 13,63')
    # e rodape de bases ('Folha INSS -> ... Liquido -> ...'). Precisam ser
    # descartadas pela linha INTEIRA — recortadas por secao elas perdem o
    # inicio e escapam do filtro de descricao.
    uteis = [row for row in bloco if not _linha_estrutural(row)]

    verbas = []
    for s in secoes:
        if s.get("papel") not in ("provento", "desconto"):
            continue
        # As colunas sao detectadas com o bloco INTEIRO, inclusive a linha de
        # totais: funcionario com um unico desconto tem so um valor na secao, e
        # sem a linha de total nao haveria ocorrencias suficientes para a coluna
        # existir — o desconto sumia calado.
        todas = [l for l in (words_in(row, s["x0"], s["x1"]) for row in bloco) if l]
        faixas = money_columns(todas)
        linhas = [l for l in (words_in(row, s["x0"], s["x1"]) for row in uteis) if l]
        if not faixas:
            continue

        # Qual coluna e a de valor: a que contem o TOTAL impresso da secao.
        # Nao da para assumir "a mais a direita" — o espelho imprime a coluna
        # de FGTS depois dos descontos, e ela roubava o papel de valor,
        # zerando os descontos do funcionario.
        chave_total = "total_proventos" if s["papel"] == "provento" else "total_descontos"
        info = (totais or {}).get(chave_total)
        ci_valor = col_of({"x1": info["x1"]}, faixas, folga=14.0) if info else None
        if ci_valor is None:
            ci_valor = len(faixas) - 1
        papeis = ["valor" if i == ci_valor else "referencia" for i in range(len(faixas))]

        for linha in linhas:
            # A referencia vem da coluna propria, entao nao se adivinha pelo
            # texto — senao '50%' de 'INT INTRAJORNADA 50%' sairia do nome.
            sr = split_row(linha, faixas, usar_ref_texto=False)
            if not _eh_linha_verba(sr, papeis):
                continue
            ref = sr["referencia"]
            for ci in sorted(sr["valores"]):
                if ci < len(papeis) and papeis[ci] == "referencia" and not ref:
                    ref = _fmt_ref(sr["valores"][ci])
            desc, ref_fim = _separa_ref_do_fim(sr["descricao"])
            ref = ref or ref_fim
            for ci, val in sorted(sr["valores"].items()):
                if ci >= len(papeis) or papeis[ci] != "valor" or abs(val) < 0.01:
                    continue
                verbas.append({
                    "codigo": sr["codigo"],
                    "descricao": desc,
                    "grupo": resolver_grupo(desc),
                    "referencia": ref,
                    "valor": abs(val),
                    "tipo": s["papel"],
                    "tipo_origem": "secao",
                })
    return verbas


# Data completa so aparece em linha de cadastro (admissao/demissao), nunca em verba
_RX_DATA = re.compile(r"\d{2}/\d{2}/\d{4}")
# Ancorado no inicio ou em '->' de proposito: 'FOLHA' solto eliminaria a
# rubrica legitima 'Arred. Prov. Folha'.
_RX_RODAPE = re.compile(
    r"^(FOLHAINSS|TOTALDE(PROVENTOS|DESCONTOS|VENCIMENTOS))|RAIS->|LIQUIDO->"
)
# Referencia colada no fim do nome da rubrica ('I.N.S.S. 9,7583').
# Precisa sair: ela muda todo mes e fragmentaria a rubrica no historico.
_RX_REF_FIM = re.compile(r"\s+(\d{1,4}(?:[.,]\d{1,4})?\s*%?)$")


def _linha_estrutural(row: list) -> bool:
    txt = compacto(row_text(row))
    return bool(_RX_DATA.search(txt) or _RX_RODAPE.search(txt))


def _separa_ref_do_fim(desc: str):
    """'I.N.S.S. 9,7583' -> ('I.N.S.S.', '9,7583')."""
    m = _RX_REF_FIM.search(desc or "")
    if not m:
        return desc, ""
    limpo = desc[:m.start()].strip(" .:-")
    # So corta se sobrar nome de rubrica utilizavel
    return (limpo, m.group(1).strip()) if len(re.sub(r"[^A-Za-z]", "", limpo)) >= 2 else (desc, "")


def _fmt_ref(v: float) -> str:
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _corrige_classificacao(verbas: list, tot_prov: float, tot_desc: float) -> list:
    """
    Fecha a conta quando a classificacao por nome errou uma rubrica.

    Se a soma dos descontos nao bate com o total impresso, procura UMA verba
    cujo valor explique exatamente a diferenca e inverte o lado dela. E o erro
    mais comum em layout de coluna unica (ex: 'ADIANTAMENTO' de 13o que na
    verdade e provento). Se nao achar, deixa como esta — quem aponta o
    problema e o validador de integridade.
    """
    avisos = []

    # Resolve os indefinidos ANTES de conferir a soma — rubrica sem grupo
    # conhecido e provento na esmagadora maioria dos casos (desconto tem
    # vocabulario pequeno e bem coberto pelo dicionario).
    for v in verbas:
        if v["tipo"] == "indefinido":
            v["tipo"] = "provento"
            v["tipo_origem"] = "presumido"

    if tot_desc <= 0 and tot_prov <= 0:
        return avisos

    soma_d = sum(v["valor"] for v in verbas if v["tipo"] == "desconto")
    dif = round(soma_d - tot_desc, 2)
    if abs(dif) <= TOLERANCIA_CENTAVOS:
        return avisos

    if dif > 0:   # desconto a mais -> alguma virou desconto sem ser
        cands = [v for v in verbas
                 if v["tipo"] == "desconto" and v["tipo_origem"] == "nome"
                 and abs(v["valor"] - dif) <= TOLERANCIA_CENTAVOS]
        if cands:
            cands[0]["tipo"] = "provento"
            cands[0]["tipo_origem"] = "corrigido"
            avisos.append(f"Rubrica '{cands[0]['descricao']}' reclassificada como provento para fechar o total.")
    else:         # desconto a menos -> alguma deveria ser desconto
        alvo = -dif
        cands = [v for v in verbas
                 if v["tipo"] == "provento" and v["tipo_origem"] in ("nome", "presumido")
                 and abs(v["valor"] - alvo) <= TOLERANCIA_CENTAVOS]
        if cands:
            cands[0]["tipo"] = "desconto"
            cands[0]["tipo_origem"] = "corrigido"
            avisos.append(f"Rubrica '{cands[0]['descricao']}' reclassificada como desconto para fechar o total.")

    return avisos


# ─────────────────────────────────────────────
# INTEGRIDADE
# ─────────────────────────────────────────────

def valida_integridade(verbas: list, totais: dict) -> dict:
    """
    Confere a extracao contra os totais impressos no proprio documento.

    Tres checagens:
      1. soma dos proventos  == total de vencimentos impresso
      2. soma dos descontos  == total de descontos impresso
      3. proventos - descontos == liquido impresso

    nivel: "ok" | "alerta" (centavos) | "erro" (nao fecha) | "sem_referencia"
    """
    sp = round(sum(v["valor"] for v in verbas if v["tipo"] == "provento"), 2)
    sd = round(sum(v["valor"] for v in verbas if v["tipo"] == "desconto"), 2)
    calc_liq = round(sp - sd, 2)

    tp = totais.get("total_proventos")
    td = totais.get("total_descontos")
    tl = totais.get("liquido")

    msgs, pior = [], "ok"

    def _cmp(nome, calc, impresso):
        nonlocal pior
        if impresso is None:
            return None
        d = round(calc - impresso, 2)
        if abs(d) <= TOLERANCIA_CENTAVOS:
            return d
        msgs.append(
            f"{nome}: extraido R$ {calc:,.2f} x impresso R$ {impresso:,.2f} "
            f"(diferenca R$ {d:,.2f})".replace(",", "X").replace(".", ",").replace("X", ".")
        )
        pior = "erro"
        return d

    d_prov = _cmp("Proventos", sp, tp)
    d_desc = _cmp("Descontos", sd, td)
    d_liq = _cmp("Liquido", calc_liq, tl)

    if tp is None and td is None and tl is None:
        pior = "sem_referencia"
        msgs.append("Documento sem totais impressos — nao foi possivel validar a extracao.")

    return {
        "nivel": pior,
        "ok": pior == "ok",
        "soma_proventos": sp,
        "soma_descontos": sd,
        "liquido_calculado": calc_liq,
        "total_proventos_impresso": tp,
        "total_descontos_impresso": td,
        "liquido_impresso": tl,
        "diff_proventos": d_prov,
        "diff_descontos": d_desc,
        "diff_liquido": d_liq,
        "mensagens": msgs,
    }


# ─────────────────────────────────────────────
# ORQUESTRACAO
# ─────────────────────────────────────────────

def _papeis_das_colunas(rows: list, faixas: list, totais: dict, perfil: dict) -> list:
    """
    Decide o papel de cada coluna de valor: provento | desconto | referencia | valor.

    Ordem de confianca:
      1. posicao do TOTAL DE VENCIMENTOS / TOTAL DE DESCONTOS impresso
         (independe de cabecalho e de sistema — e o sinal mais forte);
      2. rotulo do cabecalho da tabela;
      3. convencao brasileira: penultima coluna = vencimentos, ultima = descontos;
      4. coluna unica = 'valor' (o lado sai pelo nome da rubrica).
    """
    papeis = [None] * len(faixas)
    if not faixas:
        return papeis

    # 1) Pela posicao dos totais impressos
    col_tot = {}
    for chave, papel in (("total_proventos", "provento"), ("total_descontos", "desconto")):
        info = totais.get(chave)
        if not info:
            continue
        ci = col_of({"x1": info["x1"]}, faixas, folga=14.0)
        if ci is not None:
            col_tot[papel] = ci

    # Total de vencimentos e total de descontos impressos na MESMA coluna =
    # layout de coluna unica (Dominio e afins). A coluna nao diz o lado da
    # rubrica; quem decide e o nome dela.
    if col_tot.get("provento") is not None and col_tot.get("provento") == col_tot.get("desconto"):
        ci = col_tot["provento"]
        papeis[ci] = "valor"
        for i in range(len(papeis)):
            if papeis[i] is None:
                papeis[i] = "referencia" if i < ci else "valor"
        return papeis

    for papel, ci in col_tot.items():
        if papeis[ci] is None:
            papeis[ci] = papel

    # 2) Pelo cabecalho
    if not all(papeis):
        for i, r in enumerate(value_column_labels(rows, faixas)):
            if r in ("provento", "desconto") and papeis[i] is None:
                papeis[i] = r

    tem_prov = "provento" in papeis
    tem_desc = "desconto" in papeis

    # 3) Convencao: duas ultimas colunas sao vencimentos e descontos
    if not tem_prov and not tem_desc and len(faixas) >= 2 and perfil.get("colunas") != "unica":
        papeis[-2] = "provento"
        papeis[-1] = "desconto"
        tem_prov = tem_desc = True

    # 4) Coluna unica
    if not tem_prov and not tem_desc:
        papeis[-1] = "valor"

    # Colunas restantes a esquerda das de valor sao referencia
    idx_valor = [i for i, p in enumerate(papeis) if p in ("provento", "desconto", "valor")]
    limite = min(idx_valor) if idx_valor else len(papeis)
    for i in range(len(papeis)):
        if papeis[i] is None:
            papeis[i] = "referencia" if i < limite else "valor"
    return papeis


_RX_RESUMO = re.compile(
    r"RESUMOGERAL|RESUMODAFOLHA|RESUMODOMES|TOTA(L|IS)GERA(L|IS)|"
    r"TOTAISDAEMPRESA|TOTALIZACAODAFOLHA"
)


def _eh_resumo(texto: str) -> bool:
    """Bloco de fechamento da empresa, nao um funcionario."""
    return bool(_RX_RESUMO.search(compacto(texto)))


def _resumo_empresa(bloco: list, totais: dict) -> dict:
    """
    Totais consolidados da empresa no relatorio.

    Servem de prova de que nenhum funcionario ficou de fora da leitura: a soma
    dos recibos extraidos tem que bater com este total.
    """
    texto = deaccent("\n".join(row_text(r) for r in bloco))
    qtd = None
    m = re.search(r"QUANTIDADE\s+(\d{1,5})", texto, re.I)
    if m:
        qtd = int(m.group(1))
    m = re.search(r"ATIVOS\s*=\s*(\d{1,5})", texto, re.I)
    if qtd is None and m:
        qtd = int(m.group(1))
    return {
        "funcionarios": qtd,
        "proventos": totais.get("total_proventos"),
        "descontos": totais.get("total_descontos"),
        "liquido": totais.get("liquido"),
    }


def _tipo_recibo(bloco: list) -> str:
    """
    Tipo do recibo pelo CABECALHO do bloco, nunca pelo corpo: a rubrica
    'ADIANTAMENTO SALARIAL' dentro da tabela nao transforma um holerite
    mensal em recibo de adiantamento.
    """
    t = deaccent("\n".join(row_text(r) for r in bloco[:8])).upper()
    for tipo, rx in TIPO_RECIBO:
        if re.search(rx, t):
            return tipo
    return "mensal"


def _salario_base(verbas: list, totais: dict):
    """Salario base pela rubrica, nao por rotulo solto (que pega a referencia)."""
    if totais.get("salario_base"):
        return totais["salario_base"]
    for v in verbas:
        if v["tipo"] != "provento":
            continue
        d = flat(v["descricao"])
        if re.match(r"^SAL(ARIO)?\b", d) and "FAMILIA" not in d and "MATERNIDADE" not in d:
            return v["valor"]
    return None


def parse_holerites(raw: bytes, filename: str = "", force_ocr: bool = False) -> dict:
    """
    Ponto de entrada: PDF -> recibos normalizados.

    Retorna:
      {
        "arquivo", "layout": {...}, "ocr": bool, "competencia_documento",
        "empresa": {"nome", "cnpj"},
        "recibos": [ {...} ],
        "avisos": [str],
      }
    """
    doc = read_pdf(raw, force_ocr=force_ocr)
    avisos = list(doc["avisos"])
    paginas = doc["paginas"]

    texto_total = "\n".join(p["texto"] for p in paginas)
    perfil = layout_profiles.detect(texto_total)
    mapa = _mapa_rotulos()

    comp_doc = parse_competencia(texto_total) or competencia_from_filename(filename)
    if not comp_doc:
        avisos.append("Competencia nao identificada no documento nem no nome do arquivo.")

    empresa = {"nome": "", "cnpj": find_cnpj(texto_total) or ""}
    for p in paginas[:1]:
        for row in p["rows"][:6]:
            # Pega so o primeiro bloco da linha: o nome do sistema de folha
            # costuma ser impresso na mesma altura, la na direita.
            bloco_esq = []
            for w in row:
                if bloco_esq and w["x0"] - bloco_esq[-1]["x1"] > 20:
                    break
                bloco_esq.append(w)
            t = " ".join(w["text"] for w in bloco_esq).strip()
            if len(t) > 8 and re.search(r"LTDA|EIRELI|S/?A\b|\bME\b|\bEPP\b|\bMEI\b|SOCIEDADE", t, re.I):
                empresa["nome"] = " ".join(t.split())[:80]
                break
        if empresa["nome"]:
            break

    recibos, resumos = [], []
    for pag in paginas:
        rows = pag["rows"]
        if not rows:
            continue
        faixas = money_columns(rows)

        for bloco in split_blocos(rows):
            texto_bloco = "\n".join(row_text(r) for r in bloco)
            if len(texto_bloco.strip()) < 40:
                continue

            totais_raw = extrai_rotulados(bloco, mapa)

            # Bloco de fechamento da empresa: guarda como conferencia cruzada
            # e nao gera recibo — senao vira um "funcionario" fantasma com o
            # valor da folha inteira.
            if _eh_resumo(texto_bloco):
                tv = {k: (v["valor"] if v else None) for k, v in totais_raw.items()}
                resumo_emp = _resumo_empresa(bloco, tv)
                if resumo_emp.get("proventos"):
                    resumos.append(dict(resumo_emp, competencia=comp_doc))
                continue

            secoes = pag.get("secoes") or []
            if len([s for s in secoes if s.get("papel")]) >= 2:
                verbas = extrai_verbas_secoes(bloco, secoes, totais_raw)
            else:
                papeis = _papeis_das_colunas(bloco, faixas, totais_raw, perfil)
                verbas = extrai_verbas(bloco, faixas, papeis)
            if not verbas:
                continue

            totais = {k: (v["valor"] if v else None) for k, v in totais_raw.items()}
            av_class = _corrige_classificacao(
                verbas,
                totais.get("total_proventos") or 0.0,
                totais.get("total_descontos") or 0.0,
            )

            func = identifica_funcionario(bloco, perfil)
            integridade = valida_integridade(verbas, totais)

            # Bloco que nao da para nomear NEM conferir nao e recibo — e a
            # tabela de rubricas consolidadas que fecha o relatorio. Emiti-la
            # como funcionario dobraria a folha inteira.
            if integridade["nivel"] == "sem_referencia" and func["confianca"] < 0.5:
                continue
            comp = parse_competencia(texto_bloco, estrito=True) or comp_doc

            chave = (func["cpf"] or func["matricula"]
                     or flat(func["nome"]) or f"pag{pag['numero']}")

            recibos.append({
                "chave": chave,
                "nome": func["nome"] or "(nao identificado)",
                "nome_norm": flat(func["nome"]),
                "matricula": func["matricula"],
                "cpf": func["cpf"],
                "cbo": func["cbo"],
                "funcao": func["funcao"],
                "admissao": func["admissao"],
                "confianca_nome": func["confianca"],
                "metodo_nome": func["metodo"],
                "competencia": comp,
                "tipo": _tipo_recibo(bloco),
                "pagina": pag["numero"],
                "ocr": pag.get("ocr", False),
                "verbas": verbas,
                "totais": {
                    "proventos": totais.get("total_proventos"),
                    "descontos": totais.get("total_descontos"),
                    "liquido": totais.get("liquido"),
                    "base_inss": totais.get("base_inss"),
                    "base_fgts": totais.get("base_fgts"),
                    "fgts": totais.get("fgts_mes"),
                    "base_irrf": totais.get("base_irrf"),
                    "salario_base": _salario_base(verbas, totais),
                },
                "integridade": integridade,
                "avisos": av_class,
            })

    recibos = _dedup(recibos)

    if not recibos:
        avisos.append(
            "Nenhum recibo reconhecido neste arquivo. "
            "Use a conferencia assistida para ensinar o layout."
        )

    return {
        "arquivo": filename,
        "layout": perfil,
        "ocr": doc["ocr"],
        "competencia_documento": comp_doc,
        "empresa": empresa,
        "resumos_empresa": resumos,
        "recibos": recibos,
        "avisos": avisos,
    }


def _dedup(recibos: list) -> list:
    """
    Remove a 2a via do mesmo recibo (mesma pessoa, competencia, tipo e liquido)
    — mas PRESERVA recibos distintos da mesma pessoa no mes (mensal + ferias +
    13o), que o parser antigo descartava.
    """
    vistos, out = {}, []
    for r in recibos:
        liq = r["totais"].get("liquido")
        liq = round(liq, 2) if liq is not None else round(r["integridade"]["liquido_calculado"], 2)
        k = (r["chave"], r["competencia"], r["tipo"], liq, len(r["verbas"]))
        if k in vistos:
            # Fica com a leitura de melhor qualidade
            ant = vistos[k]
            if (r["integridade"]["ok"], r["confianca_nome"]) > (ant["integridade"]["ok"], ant["confianca_nome"]):
                out[out.index(ant)] = r
                vistos[k] = r
            continue
        vistos[k] = r
        out.append(r)
    return out
