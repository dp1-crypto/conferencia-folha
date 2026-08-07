"""
Extracao de PDF por coordenadas (words -> linhas -> colunas)
Sigma Contabilidade

Motivo de existir: regex sobre o texto corrido so funciona quando o layout e
conhecido. Holerite de escritorio concorrente muda de sistema para sistema.
Trabalhando com a posicao (x, y) de cada palavra da para descobrir a estrutura
da tabela sem saber de antemao qual sistema gerou o PDF — inclusive separar
PROVENTO de DESCONTO pela coluna em que o valor esta impresso.
"""
import io
import re

from app.core.textutils import is_money, is_ref, money

# Tolerancia vertical para considerar que duas palavras estao na mesma linha.
# Holerite usa fonte 6-9pt; 2.5pt separa linhas sem quebrar sobrescritos.
Y_TOL = 2.5


# ─────────────────────────────────────────────
# LEITURA
# ─────────────────────────────────────────────

def has_text_layer(raw: bytes, min_chars: int = 40) -> bool:
    """True se o PDF tem camada de texto utilizavel (nao e escaneado)."""
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(raw)) as pdf:
            total = 0
            for page in pdf.pages[:5]:
                total += len((page.extract_text() or "").strip())
                if total >= min_chars:
                    return True
        return False
    except Exception:
        return False


def ocr_available() -> bool:
    """True se pytesseract + binario tesseract estao instalados."""
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def _ocr_words(raw: bytes, dpi: int = 300) -> list:
    """
    Rasteriza o PDF e roda OCR devolvendo palavras com posicao — mesmo
    formato do pdfplumber, para o resto do pipeline nao saber a diferenca.
    Requer: pytesseract + tesseract + PyMuPDF.
    """
    import pytesseract
    import fitz  # PyMuPDF
    from PIL import Image

    paginas = []
    doc = fitz.open(stream=raw, filetype="pdf")
    try:
        for pno in range(doc.page_count):
            pix = doc.load_page(pno).get_pixmap(dpi=dpi)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            data = pytesseract.image_to_data(
                img, lang="por", output_type=pytesseract.Output.DICT,
                config="--psm 6",
            )
            # Converte pixels de volta para pontos PDF (72 dpi)
            esc = 72.0 / dpi
            words = []
            for i, txt in enumerate(data["text"]):
                txt = (txt or "").strip()
                if not txt:
                    continue
                try:
                    conf = float(data["conf"][i])
                except (TypeError, ValueError):
                    conf = -1.0
                if conf < 30:          # descarta lixo de OCR
                    continue
                x, y = data["left"][i] * esc, data["top"][i] * esc
                w, h = data["width"][i] * esc, data["height"][i] * esc
                words.append({
                    "text": txt, "x0": x, "x1": x + w,
                    "top": y, "bottom": y + h, "conf": conf,
                })
            paginas.append({
                "numero": pno + 1,
                "largura": pix.width * esc,
                "altura": pix.height * esc,
                "words": words,
                "ocr": True,
            })
    finally:
        doc.close()
    return paginas


def read_pdf(raw: bytes, force_ocr: bool = False) -> dict:
    """
    Le um PDF e devolve estrutura uniforme:
      {"paginas": [{numero, largura, altura, words, rows, texto, ocr}],
       "ocr": bool, "avisos": [str]}

    Faz OCR automaticamente quando o PDF nao tem camada de texto.
    Se o OCR nao estiver disponivel, devolve paginas vazias com aviso claro
    em vez de silenciosamente entregar nada.
    """
    avisos = []
    usou_ocr = False
    paginas = []

    precisa_ocr = force_ocr or not has_text_layer(raw)

    if precisa_ocr:
        if ocr_available():
            try:
                paginas = _ocr_words(raw)
                usou_ocr = True
                avisos.append(
                    "PDF sem camada de texto — extraido por OCR. "
                    "Confira os valores antes de importar."
                )
            except Exception as e:
                avisos.append(f"Falha no OCR: {e}")
        else:
            avisos.append(
                "PDF escaneado (sem texto) e OCR nao instalado. "
                "Instale com: brew install tesseract tesseract-lang && pip3 install pytesseract"
            )

    if not paginas:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(raw)) as pdf:
            for i, page in enumerate(pdf.pages):
                words = [
                    {"text": w["text"], "x0": w["x0"], "x1": w["x1"],
                     "top": w["top"], "bottom": w["bottom"], "conf": 100.0}
                    for w in page.extract_words(
                        use_text_flow=False, keep_blank_chars=False,
                        extra_attrs=[],
                    )
                ]
                paginas.append({
                    "numero": i + 1,
                    "largura": page.width,
                    "altura": page.height,
                    "words": words,
                    "ocr": False,
                })

    for p in paginas:
        p["rows"] = [repair_row(r) for r in group_rows(p["words"])]
        p["texto"] = "\n".join(row_text(r) for r in p["rows"])
        p["secoes"] = detect_secoes(p["rows"])

    return {"paginas": paginas, "ocr": usou_ocr, "avisos": avisos}


# ─────────────────────────────────────────────
# LINHAS
# ─────────────────────────────────────────────

def group_rows(words: list, ytol: float = Y_TOL) -> list:
    """
    Agrupa palavras em linhas visuais pela coordenada vertical.
    Cada linha e uma lista de words ordenada da esquerda para a direita.
    """
    if not words:
        return []
    ordenadas = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))
    linhas, atual, ref = [], [], None
    for w in ordenadas:
        centro = (w["top"] + w["bottom"]) / 2.0
        if ref is None or abs(centro - ref) <= ytol:
            atual.append(w)
            ref = centro if ref is None else (ref + centro) / 2.0
        else:
            linhas.append(sorted(atual, key=lambda x: x["x0"]))
            atual, ref = [w], centro
    if atual:
        linhas.append(sorted(atual, key=lambda x: x["x0"]))
    return linhas


def row_text(row: list) -> str:
    return " ".join(w["text"] for w in row)


# ─────────────────────────────────────────────
# REPARO DE TEXTO QUEBRADO
# ─────────────────────────────────────────────

def _split_run(run: list) -> list:
    """
    Recompoe um trecho com espacamento entre letras ('F o l h a  I N S S').

    Dentro do trecho existem dois tipos de espaco: o que separa letras da
    mesma palavra (bem pequeno) e o que separa palavras (maior). Corta pelo
    proprio ritmo do trecho, sem numero magico: qualquer folga acima da
    mediana + 1pt e separador de palavra.
    """
    gaps = [run[i + 1]["x0"] - run[i]["x1"] for i in range(len(run) - 1)]
    if not gaps:
        return run
    ordenados = sorted(gaps)
    mediana = ordenados[len(ordenados) // 2]
    corte = mediana + 1.0

    saida, atual = [], dict(run[0])
    for i, w in enumerate(run[1:]):
        if gaps[i] > corte:
            saida.append(atual)
            atual = dict(w)
        else:
            atual["text"] += w["text"]
            atual["x1"] = w["x1"]
    saida.append(atual)
    return saida


def _repara_letras_soltas(row: list, gap_max: float = 3.5, minimo: int = 4) -> list:
    """Junta sequencias de caracteres isolados de volta em palavras."""
    saida, i = [], 0
    while i < len(row):
        j = i
        while (j + 1 < len(row) and len(row[j]["text"]) == 1
               and len(row[j + 1]["text"]) == 1
               and row[j + 1]["x0"] - row[j]["x1"] <= gap_max):
            j += 1
        if j - i + 1 >= minimo:
            saida.extend(_split_run(row[i:j + 1]))
            i = j + 1
        else:
            saida.append(dict(row[i]))
            i += 1
    return saida


def _repara_numeros(row: list, gap_max: float = 3.0) -> list:
    """
    Remonta valor partido pela extracao ('599 , 72' -> '599,72').
    So junta quando o resultado e um valor monetario valido — assim nunca
    cola dois valores de colunas diferentes.
    """
    saida, i = [], 0
    while i < len(row):
        melhor = None
        texto, x1 = row[i]["text"], row[i]["x1"]
        for j in range(i + 1, min(i + 5, len(row))):
            if row[j]["x0"] - x1 > gap_max:
                break
            texto += row[j]["text"]
            x1 = row[j]["x1"]
            if is_money(texto):
                melhor = (j, texto, x1)
        if melhor and not is_money(row[i]["text"]):
            j, texto, x1 = melhor
            novo = dict(row[i])
            novo["text"], novo["x1"] = texto, x1
            saida.append(novo)
            i = j + 1
        else:
            saida.append(dict(row[i]))
            i += 1
    return saida


def repair_row(row: list) -> list:
    """Aplica os dois reparos na ordem: letras soltas primeiro, depois numeros."""
    if not row:
        return row
    return _repara_numeros(_repara_letras_soltas(row))


# ─────────────────────────────────────────────
# SECOES VERTICAIS (tabelas lado a lado)
# ─────────────────────────────────────────────

_RX_COL_DESC = re.compile(r"^(COD\.?)?DESCRI[CÇ]", re.I)
_RX_PROV = re.compile(r"^(VENCIMENT|PROVENT|CREDIT|RENDIMENT)", re.I)
_RX_DESC = re.compile(r"^(DESCONT|DEBIT|DEDUC)", re.I)


def detect_secoes(rows: list) -> list:
    """
    Descobre tabelas lado a lado na mesma pagina.

    Relatorio de folha (espelho / analitico) costuma imprimir PROVENTOS a
    esquerda e DESCONTOS a direita, cada um com suas colunas Cod/Descricao/
    Referencia/Valor. Sem separar as duas, a linha
    '60 Gratificacoes 1.000,00 9.101 I.N.S.S. 517,08' viraria uma verba so,
    com a descricao de uma e o valor da outra.

    Retorna [] quando a pagina tem uma tabela unica (holerite comum).
    """
    from app.core.textutils import deaccent

    for row in rows[:40]:
        cabecas = [w for w in row if _RX_COL_DESC.match(deaccent(w["text"]))]
        if len(cabecas) < 2:
            continue
        cortes = [c["x0"] - 8 for c in cabecas[1:]]
        limites = [0.0] + cortes + [1e6]
        secoes = [{"x0": limites[i], "x1": limites[i + 1], "papel": None}
                  for i in range(len(limites) - 1)]
        _rotula_secoes(rows, secoes)
        return secoes
    return []


def _rotula_secoes(rows: list, secoes: list):
    """Marca cada secao como provento ou desconto pelo titulo impresso acima."""
    from app.core.textutils import deaccent
    for row in rows[:40]:
        for w in row:
            t = deaccent(w["text"])
            papel = "provento" if _RX_PROV.match(t) else ("desconto" if _RX_DESC.match(t) else None)
            if not papel:
                continue
            centro = (w["x0"] + w["x1"]) / 2.0
            for s in secoes:
                if s["x0"] <= centro < s["x1"] and s["papel"] is None:
                    s["papel"] = papel
        if all(s["papel"] for s in secoes):
            return


def words_in(row: list, x0: float, x1: float) -> list:
    """Palavras da linha dentro da faixa horizontal da secao."""
    return [w for w in row if x0 <= (w["x0"] + w["x1"]) / 2.0 < x1]


def row_top(row: list) -> float:
    return min(w["top"] for w in row) if row else 0.0


def find_row(rows: list, pattern, inicio: int = 0):
    """Indice da primeira linha cujo texto casa com o padrao (regex, sem acento)."""
    from app.core.textutils import deaccent
    rx = re.compile(pattern, re.I) if isinstance(pattern, str) else pattern
    for i in range(inicio, len(rows)):
        if rx.search(deaccent(row_text(rows[i]))):
            return i
    return -1


# ─────────────────────────────────────────────
# COLUNAS
# ─────────────────────────────────────────────

def money_columns(rows: list, min_ocorrencias: int = 2, gap: float = 12.0) -> list:
    """
    Descobre as colunas de VALOR agrupando, pelo x da borda direita, todos os
    tokens monetarios da pagina. Numero costuma ser alinhado a direita, entao
    a borda direita e muito mais estavel que a esquerda.

    min_ocorrencias=2 de proposito: em folha pequena (1 funcionario, 3 rubricas)
    a coluna de vencimentos aparece poucas vezes, e exigir 3 fazia a coluna
    inteira ser ignorada — os proventos somavam zero.

    Devolve as faixas ordenadas da esquerda para a direita:
      [{"x_min", "x_max", "centro", "n"}]
    """
    marcas = []
    for row in rows:
        for w in row:
            if is_money(w["text"]):
                marcas.append(w["x1"])
    if not marcas:
        return []

    marcas.sort()
    grupos, atual = [], [marcas[0]]
    for x in marcas[1:]:
        if x - atual[-1] <= gap:
            atual.append(x)
        else:
            grupos.append(atual)
            atual = [x]
    grupos.append(atual)

    faixas = [
        {"x_min": min(g), "x_max": max(g), "centro": sum(g) / len(g), "n": len(g)}
        for g in grupos if len(g) >= min_ocorrencias
    ]
    return sorted(faixas, key=lambda f: f["centro"])


def col_of(word: dict, faixas: list, folga: float = 10.0):
    """Indice da faixa de coluna a que o valor pertence, ou None."""
    for i, f in enumerate(faixas):
        if f["x_min"] - folga <= word["x1"] <= f["x_max"] + folga:
            return i
    return None


def split_row(row: list, faixas: list, usar_ref_texto: bool = True) -> dict:
    """
    Quebra uma linha da tabela de verbas em:
      descricao (texto a esquerda da 1a coluna de valor),
      codigo    (numero inteiro no inicio da linha, quando houver),
      valores   {indice_coluna: float},
      referencia (token de referencia antes do 1o valor).

    E aqui que PROVENTO se separa de DESCONTO: pelo indice da coluna
    em que o valor foi impresso.
    """
    if not row or not faixas:
        return {"codigo": "", "descricao": row_text(row), "referencia": "", "valores": {}}

    limite_esq = faixas[0]["x_min"] - 10.0
    valores, ref, desc_tokens = {}, "", []

    for w in row:
        t = w["text"].strip()
        if w["x1"] <= limite_esq:
            desc_tokens.append(t)
            continue
        if is_money(t):
            ci = col_of(w, faixas)
            if ci is not None:
                # Linha com dois numeros na mesma coluna: mantem o ultimo
                valores[ci] = money(t)
                continue
        if usar_ref_texto and not ref and is_ref(t) and not is_money(t):
            ref = t
            continue
        desc_tokens.append(t)

    codigo = ""
    if desc_tokens and re.fullmatch(r"\d{1,3}(?:\.\d{3})*|\d{1,6}", desc_tokens[0]):
        codigo = desc_tokens.pop(0)

    # Referencia pode estar colada ao fim da descricao
    if usar_ref_texto and not ref and desc_tokens and is_ref(desc_tokens[-1]) and not desc_tokens[-1].isalpha():
        ref = desc_tokens.pop()

    return {
        "codigo": codigo,
        "descricao": " ".join(desc_tokens).strip(" .:-"),
        "referencia": ref,
        "valores": valores,
    }


def value_column_labels(rows: list, faixas: list) -> list:
    """
    Tenta rotular cada coluna de valor lendo o cabecalho da tabela
    (VENCIMENTOS / PROVENTOS / DESCONTOS / VALOR / REFERENCIA).

    Retorna lista paralela a `faixas`: "provento" | "desconto" | "valor" | None
    """
    from app.core.textutils import deaccent

    PROV = re.compile(r"VENCIMENT|PROVENT|CREDIT|RENDIMENT|GANHO", re.I)
    DESC = re.compile(r"DESCONT|DEBIT|DEDUC", re.I)
    VAL = re.compile(r"\bVALOR\b|\bTOTAL\b", re.I)

    # Cabecalho de tabela: tem rotulo de coluna e NAO tem valor monetario.
    # Linha de totais ("Total de Vencimentos 3.964,75") tem valor e por isso
    # e descartada aqui — senao o rotulo do total seria lido como cabecalho da
    # coluna errada, ja que fica deslocado a esquerda do numero.
    rotulos = [None] * len(faixas)
    for row in rows[:40]:
        txt = deaccent(row_text(row))
        if not (PROV.search(txt) or DESC.search(txt)):
            continue
        if re.search(r"\bTOTAL\b|\bL[IÍ]QUIDO\b", txt, re.I):
            continue
        if any(is_money(w["text"]) for w in row):
            continue
        for w in row:
            t = deaccent(w["text"])
            alvo = None
            if PROV.search(t):
                alvo = "provento"
            elif DESC.search(t):
                alvo = "desconto"
            elif VAL.search(t):
                alvo = "valor"
            if not alvo:
                continue
            # Cabecalho costuma ficar sobre a coluna: casa pelo centro
            centro_w = (w["x0"] + w["x1"]) / 2.0
            melhor, dist = None, 1e9
            for i, f in enumerate(faixas):
                d = abs(f["centro"] - centro_w)
                if d < dist:
                    melhor, dist = i, d
            if melhor is not None and dist < 60 and rotulos[melhor] is None:
                rotulos[melhor] = alvo
        if any(rotulos):
            break
    return rotulos
