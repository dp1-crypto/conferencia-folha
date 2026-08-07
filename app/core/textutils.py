"""
Utilitarios de texto para extracao de documentos de folha
Sigma Contabilidade
"""
import re
import unicodedata
from datetime import date

# ─────────────────────────────────────────────
# NORMALIZACAO
# ─────────────────────────────────────────────

def deaccent(s: str) -> str:
    """Remove acentos preservando o restante."""
    return "".join(
        c for c in unicodedata.normalize("NFD", str(s))
        if unicodedata.category(c) != "Mn"
    )


def flat(s: str) -> str:
    """Maiusculo, sem acento, sem pontuacao, espacos colapsados. Para comparacao."""
    t = deaccent(str(s)).upper()
    t = re.sub(r"[^\w\s]", " ", t)
    return " ".join(t.split())


def only_digits(s) -> str:
    return re.sub(r"\D", "", str(s or ""))


# ─────────────────────────────────────────────
# DOCUMENTOS
# ─────────────────────────────────────────────

def valid_cpf(raw) -> bool:
    """Valida CPF pelos digitos verificadores."""
    c = only_digits(raw)
    if len(c) != 11 or c == c[0] * 11:
        return False
    for pos in (9, 10):
        soma = sum(int(c[i]) * ((pos + 1) - i) for i in range(pos))
        dv = (soma * 10) % 11
        if dv == 10:
            dv = 0
        if dv != int(c[pos]):
            return False
    return True


def valid_cnpj(raw) -> bool:
    """Valida CNPJ pelos digitos verificadores."""
    c = only_digits(raw)
    if len(c) != 14 or c == c[0] * 14:
        return False
    pesos1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    pesos2 = [6] + pesos1
    for pesos, pos in ((pesos1, 12), (pesos2, 13)):
        soma = sum(int(c[i]) * pesos[i] for i in range(pos))
        dv = soma % 11
        dv = 0 if dv < 2 else 11 - dv
        if dv != int(c[pos]):
            return False
    return True


def find_cpf(text: str):
    """Primeiro CPF valido encontrado no texto (formatado ou nao)."""
    for m in re.finditer(r"\b\d{3}[.\s]?\d{3}[.\s]?\d{3}[-\s]?\d{2}\b", text or ""):
        if valid_cpf(m.group(0)):
            return only_digits(m.group(0))
    return None


def find_cnpj(text: str):
    for m in re.finditer(r"\b\d{2}[.\s]?\d{3}[.\s]?\d{3}[/\s]?\d{4}[-\s]?\d{2}\b", text or ""):
        if valid_cnpj(m.group(0)):
            return only_digits(m.group(0))
    return None


def fmt_cpf(c) -> str:
    c = only_digits(c)
    return f"{c[:3]}.{c[3:6]}.{c[6:9]}-{c[9:]}" if len(c) == 11 else str(c or "")


def fmt_cnpj(c) -> str:
    c = only_digits(c)
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}" if len(c) == 14 else str(c or "")


# ─────────────────────────────────────────────
# COMPETENCIA
# ─────────────────────────────────────────────

MESES = {
    "JANEIRO": 1, "JAN": 1, "FEVEREIRO": 2, "FEV": 2, "MARCO": 3, "MAR": 3,
    "ABRIL": 4, "ABR": 4, "MAIO": 5, "MAI": 5, "JUNHO": 6, "JUN": 6,
    "JULHO": 7, "JUL": 7, "AGOSTO": 8, "AGO": 8, "SETEMBRO": 9, "SET": 9,
    "OUTUBRO": 10, "OUT": 10, "NOVEMBRO": 11, "NOV": 11, "DEZEMBRO": 12, "DEZ": 12,
}
MESES_NOME = ["", "Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho",
              "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

# Rotulos que costumam preceder a competencia no holerite
_ROTULO_COMP = r"(?:COMPET[EÊ]NCIA|REFER[EÊ]NCIA|M[EÊ]S[/\s]*ANO|PER[IÍ]ODO|FOLHA|M[EÊ]S)"


def _valid_ym(ano: int, mes: int) -> bool:
    return 1 <= mes <= 12 and 1990 <= ano <= date.today().year + 1


def _ym(ano: int, mes: int) -> str:
    return f"{ano:04d}-{mes:02d}"


def parse_competencia(text: str, estrito: bool = False):
    """
    Extrai a competencia (AAAA-MM) de um texto de holerite / relatorio de folha.

    Cobre os formatos usuais:
      "Competencia: 03/2026"      "Referencia MARCO/2026"
      "Mes/Ano: 03 2026"          "MARCO DE 2026"
      "01/03/2026 a 31/03/2026"   "03-2026"

    Retorna None quando nao encontra nada confiavel — o chamador decide
    se pergunta ao usuario ou usa o nome do arquivo.
    """
    if not text:
        return None
    t = deaccent(text).upper()

    # 0) "REFERENTE AO MES DE JUNHO DE 2026" — cabecalho de relatorio de folha.
    #    Vem antes de tudo: e a afirmacao mais explicita que o documento faz.
    m = re.search(
        r"(?:REFERENTE\s+AO\s+M[EÊ]S\s+DE|M[EÊ]S\s+DE|COMPETENCIA\s+DE|"
        r"REFERENTE\s+A)\s+([A-Z]{3,9})\s*(?:DE|/|-)?\s*(\d{4})", t)
    if m and m.group(1) in MESES and _valid_ym(int(m.group(2)), MESES[m.group(1)]):
        return _ym(int(m.group(2)), MESES[m.group(1)])

    # 1) Rotulo explicito + MM/AAAA  (mais confiavel)
    m = re.search(_ROTULO_COMP + r"[^\d\n]{0,20}(\d{1,2})\s*[/\-.\s]\s*(\d{4})", t)
    if m and _valid_ym(int(m.group(2)), int(m.group(1))):
        return _ym(int(m.group(2)), int(m.group(1)))

    # 2) Rotulo explicito + MES POR EXTENSO + ano
    m = re.search(_ROTULO_COMP + r"[^A-Z\n]{0,20}([A-Z]{3,9})\s*(?:DE\s*)?[/\-\s]?\s*(\d{4})", t)
    if m and m.group(1) in MESES and _valid_ym(int(m.group(2)), MESES[m.group(1)]):
        return _ym(int(m.group(2)), MESES[m.group(1)])

    # 3) Periodo "01/03/2026 a 31/03/2026" -> usa o mes inicial
    m = re.search(r"\b01[/\-.](\d{1,2})[/\-.](\d{4})\s*(?:A|ATE|-)\s*\d{2}[/\-.]\1[/\-.]\2\b", t)
    if m and _valid_ym(int(m.group(2)), int(m.group(1))):
        return _ym(int(m.group(2)), int(m.group(1)))

    # Modo estrito: dai para baixo tudo e inferencia sem rotulo. Num bloco de
    # UM funcionario isso e perigoso — a data de admissao ('23/02/2026') viraria
    # a competencia do recibo e criaria meses que nunca existiram na folha.
    if estrito:
        return None

    # 4) Mes por extenso + ano, sem rotulo (ex: cabecalho "MARCO/2026")
    m = re.search(r"\b([A-Z]{3,9})\s*[/\-]\s*(\d{4})\b", t)
    if m and m.group(1) in MESES and _valid_ym(int(m.group(2)), MESES[m.group(1)]):
        return _ym(int(m.group(2)), MESES[m.group(1)])

    # 5) MM/AAAA solto — so aceita se houver UMA unica ocorrencia no documento,
    #    para nao confundir com data de admissao, vencimento etc.
    achados = set()
    for mm in re.finditer(r"(?<![\d/])(\d{1,2})\s*[/\-]\s*(\d{4})(?![\d/])", t):
        ano, mes = int(mm.group(2)), int(mm.group(1))
        if _valid_ym(ano, mes):
            achados.add(_ym(ano, mes))
    if len(achados) == 1:
        return achados.pop()

    return None


def competencia_from_filename(filename: str):
    """Ultimo recurso: tenta ler a competencia do nome do arquivo."""
    if not filename:
        return None
    base = deaccent(filename).upper()
    base = re.sub(r"\.(PDF|XLSX?|DOCX?|CSV|TXT|JPE?G|PNG)$", "", base)

    # (?<!\d) em vez de \b: em 'folha_2026-03' o underscore e caractere de
    # palavra, entao \b nao casa antes do ano.
    m = re.search(r"(?<!\d)(\d{4})\s*[-_.]?\s*(\d{2})(?!\d)", base)   # 2026-03 / 2026_03
    if m and _valid_ym(int(m.group(1)), int(m.group(2))):
        return _ym(int(m.group(1)), int(m.group(2)))

    m = re.search(r"(?<!\d)(\d{1,2})\s*[-_.]\s*(\d{4})(?!\d)", base)  # 03-2026
    if m and _valid_ym(int(m.group(2)), int(m.group(1))):
        return _ym(int(m.group(2)), int(m.group(1)))

    # Underscore e caractere de palavra, entao \b nao casa em 'folha_MARCO_2026'.
    # Trocar separadores por espaco resolve sem afrouxar a fronteira — afrouxar
    # faria 'mariana2026.pdf' virar marco, porque MARIANA contem MAR.
    base_sep = re.sub(r"[_\-.]+", " ", base)
    for nome, num in MESES.items():
        if len(nome) < 3:
            continue
        m = re.search(r"\b" + nome + r"\b[^\d]{0,6}(\d{4})\b", base_sep)
        if m and _valid_ym(int(m.group(1)), num):
            return _ym(int(m.group(1)), num)
    return None


def competencia_label(comp) -> str:
    """'2026-03' -> 'Marco/2026'."""
    if not comp or "-" not in str(comp):
        return str(comp or "-")
    ano, mes = str(comp).split("-")[:2]
    try:
        return f"{MESES_NOME[int(mes)]}/{ano}"
    except Exception:
        return str(comp)


def competencia_range(comps):
    """Lista ordenada de competencias contiguas entre a menor e a maior informada."""
    vals = sorted({c for c in comps if c and "-" in str(c)})
    if not vals:
        return []
    ini_a, ini_m = (int(x) for x in vals[0].split("-")[:2])
    fim_a, fim_m = (int(x) for x in vals[-1].split("-")[:2])
    out, a, m = [], ini_a, ini_m
    while (a, m) <= (fim_a, fim_m) and len(out) < 240:
        out.append(_ym(a, m))
        m += 1
        if m > 12:
            m, a = 1, a + 1
    return out


# ─────────────────────────────────────────────
# VALORES MONETARIOS
# ─────────────────────────────────────────────

# Numero no padrao BR: 1.234,56 | 1234,56 | 0,00 | (1.234,56)
RE_MONEY = re.compile(r"^\(?-?\s*\d{1,3}(?:\.\d{3})*,\d{2}\)?-?$|^\(?-?\s*\d+,\d{2}\)?-?$")
# Referencia: 30,00 | 220:00 | 8,33% | 3:41 | 30
RE_REF = re.compile(r"^\d{1,4}(?:[.,:]\d{1,4})?\s*%?$")


def is_money(tok: str) -> bool:
    """True se o token e um valor monetario no padrao brasileiro."""
    return bool(RE_MONEY.match(str(tok).strip()))


def is_ref(tok: str) -> bool:
    """True se o token parece uma referencia (dias, horas, percentual)."""
    return bool(RE_REF.match(str(tok).strip()))


def money(tok):
    """
    Converte token monetario BR para float, tratando negativo por
    parenteses '(100,00)' e por sinal a direita '100,00-' (comum em relatorios).
    """
    s = re.sub(r"[R$\s]", "", str(tok))
    neg = (s.startswith("(") and s.endswith(")")) or s.endswith("-") or s.startswith("-")
    s = s.strip("()-")
    s = s.replace(".", "").replace(",", ".")
    try:
        v = float(s)
    except Exception:
        return 0.0
    return -v if neg else v
