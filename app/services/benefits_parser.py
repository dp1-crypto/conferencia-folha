"""
Parsers de beneficios — fatura plano de saude, extrato de folha, referencia simples
"""
import io
import re

from app.core.utils import norm, brl


def fix_spaced(text: str) -> str:
    """Remove espacos entre digitos (formato Unimed: '6 2 9 , 3 5' -> '629,35')."""
    for _ in range(15):
        text = re.sub(r"(\d) (\d)", r"\1\2", text)
    text = re.sub(r"(\d)\s*,\s*(\d)", r"\1,\2", text)
    text = re.sub(r"(\d)\s+\.\s*(\d)", r"\1.\2", text)
    text = re.sub(r"(\d)\s*\.\s+(\d)", r"\1.\2", text)
    return text


def _page_lines_smart(page) -> list:
    """
    Extrai linhas de uma pagina usando posicoes de caracteres para reconstruir
    limites de palavras reais. Funciona com PDFs onde cada letra e espacada
    individualmente (ex: Unimed analitico), detectando que word-gaps tem ~2x
    o espacamento de letter-gaps.
    """
    from collections import defaultdict
    try:
        from statistics import median
    except ImportError:
        def median(lst):
            s = sorted(lst)
            n = len(s)
            return s[n // 2] if n % 2 else (s[n // 2 - 1] + s[n // 2]) / 2

    chars = page.chars
    if not chars:
        return []

    rows = defaultdict(list)
    for c in chars:
        y_key = round(c["doctop"] / 2) * 2
        rows[y_key].append(c)

    lines = []
    for y in sorted(rows.keys()):
        row = [c for c in rows[y] if c["text"].strip()]
        if not row:
            continue
        row.sort(key=lambda c: c["x0"])
        if len(row) == 1:
            lines.append(row[0]["text"])
            continue

        gaps = [row[i + 1]["x0"] - row[i]["x0"] for i in range(len(row) - 1)]
        med = median(gaps)
        threshold = max(med * 1.6, 1.5)

        text = row[0]["text"]
        for i, gap in enumerate(gaps):
            if gap > threshold:
                text += " "
            text += row[i + 1]["text"]
        lines.append(text)

    return lines


def _merge_fatura(base: dict, new: dict) -> dict:
    """Acumula valores da fatura — soma ao inves de sobrescrever (multi-arquivo/mes)."""
    for key, val in new.items():
        if key in base:
            base[key]["mensalidade"] = round(base[key]["mensalidade"] + val["mensalidade"], 2)
            base[key]["mensalidade_dependentes"] = round(
                base[key].get("mensalidade_dependentes", 0) + val.get("mensalidade_dependentes", 0), 2
            )
            base[key]["sos_tam"] = round(base[key].get("sos_tam", 0) + val.get("sos_tam", 0), 2)
            base[key]["total"] = round(base[key]["total"] + val["total"], 2)
            dep = base[key].get("dependentes")
            new_dep = val.get("dependentes", [])
            if isinstance(dep, list):
                base[key]["dependentes"] = dep + new_dep
            else:
                base[key]["dependentes"] = (dep or 0) + len(new_dep)
        else:
            base[key] = dict(val)
    return base


def _merge_extrato(base: dict, new: dict) -> dict:
    """Acumula descontos do extrato — soma ao inves de sobrescrever (multi-arquivo/mes)."""
    for key, val in new.items():
        if key in base:
            base[key]["plano_descontado"] = round(
                base[key]["plano_descontado"] + val["plano_descontado"], 2
            )
            if val.get("salario", 0) > base[key].get("salario", 0):
                base[key]["salario"] = val["salario"]
        else:
            base[key] = dict(val)
    return base


def parse_plano_fatura(raw: bytes, filtro_linha: str = "MENSALIDADE") -> dict:
    """
    Le fatura/relatorio de beneficio e extrai valor por titular.
    Suporta formato Unimed analitico (letras espacadas individualmente) e
    documentos genericos com codigo-nome-valor.
    """
    import pdfplumber

    filtro = filtro_linha.upper().strip()
    titulares = {}
    titular_atual = None

    with pdfplumber.open(io.BytesIO(raw)) as pdf:
        for page in pdf.pages:
            # Usa extracao inteligente por posicao de chars para reconstruir
            # word boundaries reais (resolve formato Unimed tudo-espacado)
            lines = _page_lines_smart(page)
            for line in lines:
                line = line.strip()
                if not line:
                    continue

                filtro_ok  = filtro in line.upper()
                sos_tam_ok = bool(re.search(r"\bSOS\b|\bTAM\b", line.upper()))
                if not filtro_ok and not sos_tam_ok:
                    continue

                # Normaliza espacos em digitos que ainda estejam separados
                line_fixed = fix_spaced(line)

                valores = re.findall(r"(?<!\d)[\d.]+,\d{2}(?!\d)", line_fixed)
                valores_validos = [v for v in valores if brl(v) >= 1.0]
                total_val = brl(valores_validos[-1]) if valores_validos else 0.0

                if not filtro_ok and sos_tam_ok:
                    if titular_atual and titular_atual in titulares and total_val > 0:
                        titulares[titular_atual]["sos_tam"] += total_val
                    continue

                if total_val == 0:
                    continue

                cod_m = re.search(r"(\d{4}\.\d{4}\.\d{6})-(\d{2,3})", line_fixed)

                if cod_m:
                    codigo_base = cod_m.group(1)
                    sufixo      = cod_m.group(2)

                    pos_cod  = line_fixed.find(cod_m.group(0))
                    pos_filt = line_fixed.upper().find(filtro.split()[0])
                    if pos_filt < 0:
                        pos_filt = len(line_fixed)
                    trecho = line_fixed[pos_cod + len(cod_m.group(0)):pos_filt].strip()

                    nome_raw = re.sub(r"^\s*[AIER]\s+", "", trecho).strip()
                    nome_raw = re.sub(r"\s+[AIER]\s*$", "", nome_raw).strip()
                    nome_raw = re.sub(r"\b\d+\b", "", nome_raw).strip()
                    nome_raw = re.sub(r"\s{2,}", " ", nome_raw).strip()

                    if sufixo == "00":
                        titular_atual = codigo_base
                        titulares[codigo_base] = {
                            "nome_original": nome_raw.title() if nome_raw else codigo_base,
                            "nome_norm":     norm(nome_raw) if nome_raw else codigo_base,
                            "mensalidade":   total_val,
                            "sos_tam":       0.0,
                            "dependentes":   [],
                        }
                    else:
                        if titular_atual and titular_atual in titulares:
                            titulares[titular_atual]["dependentes"].append({
                                "nome": nome_raw.title(),
                                "valor": total_val,
                            })
                else:
                    pos_filt = line_fixed.upper().find(filtro.split()[0])
                    nome_raw = line_fixed[:pos_filt].strip() if pos_filt > 0 else ""
                    nome_raw = re.sub(r"\b\d+\b", "", nome_raw).strip()
                    nome_raw = re.sub(r"^\s*[AIER]\s+", "", nome_raw).strip()
                    nome_raw = re.sub(r"\s+[AIER]\s*$", "", nome_raw).strip()
                    nome_raw = re.sub(r"\s{2,}", " ", nome_raw).strip()

                    if nome_raw and len(nome_raw.split()) >= 2:
                        nn = norm(nome_raw)
                        titular_atual = nn
                        titulares[nn] = {
                            "nome_original": nome_raw.title(),
                            "nome_norm":     nn,
                            "mensalidade":   total_val,
                            "sos_tam":       0.0,
                            "dependentes":   [],
                        }
                    elif titular_atual and titular_atual in titulares:
                        titulares[titular_atual]["dependentes"].append({
                            "nome": "",
                            "valor": total_val,
                        })

    result = {}
    for dados in titulares.values():
        nn = dados["nome_norm"]
        if not nn:
            continue
        total_dep = sum(d["valor"] for d in dados["dependentes"])
        result[nn] = {
            "nome_original":          dados["nome_original"],
            "mensalidade":            dados["mensalidade"],
            "sos_tam":                dados["sos_tam"],
            "total":                  round(dados["mensalidade"] + total_dep + dados["sos_tam"], 2),
            "mensalidade_dependentes": total_dep,
            "dependentes":            dados["dependentes"],
        }

    return result


def parse_extrato_plano(raw: bytes, codigo: str = "8111") -> dict:
    """
    Le extrato de folha e extrai o evento pelo codigo informado.
    Retorna: {NOME_NORM: {nome_original, plano_descontado, salario}}
    """
    import pdfplumber

    result = {}
    cod = str(codigo).strip()

    with pdfplumber.open(io.BytesIO(raw)) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += (page.extract_text() or "") + "\n"

    blocos = re.split(r"(?=Empr\.:?\s*\d+[A-Z])", full_text)

    for bloco in blocos:
        nm = re.search(r"Empr\.:?\s*\d+([A-Z][A-Z\s]+?)(?:\s+Situa[çc][aã]o\s*:|\s+CPF\s*:)", bloco)
        if not nm:
            continue

        nome = nm.group(1).strip()
        nome = re.sub(r"\s{2,}", " ", nome)

        sal_m = re.search(r"Sal[aá]rio:\s*([\d.,]+)", bloco, re.I)
        salario = brl(sal_m.group(1)) if sal_m else 0.0

        # Busca TODAS as ocorrencias do codigo e soma (desconto pode aparecer mais de uma vez)
        matches = re.findall(
            rf"{re.escape(cod)}\s+[A-Z][A-Z0-9\s./ºÇÃÕÁÉÍÓÚ]*?\s+([\d.,]+)\s+([\d.,]+)\s*D?",
            bloco, re.I
        )
        valor_descontado = round(sum(brl(m[1]) for m in matches), 2)

        if not matches and salario == 0.0:
            continue

        chave = norm(nome)
        if chave in result:
            # Mesmo funcionario em outro mes do mesmo arquivo — acumula
            result[chave]["plano_descontado"] = round(
                result[chave]["plano_descontado"] + valor_descontado, 2
            )
            if salario > result[chave].get("salario", 0):
                result[chave]["salario"] = salario
        else:
            result[chave] = {
                "nome_original": nome.title(),
                "plano_descontado": valor_descontado,
                "salario": salario,
            }

    return result


def parse_referencia_simples(raw: bytes, filename: str) -> dict:
    """
    Le documento de referencia em formato Excel simples (nome | valor).
    Retorna: {NOME_NORM: {nome_original, mensalidade, total}}
    """
    rows = []
    if filename.lower().endswith(".xls"):
        import xlrd
        wb = xlrd.open_workbook(file_contents=raw)
        ws = wb.sheets()[0]
        for r in range(ws.nrows):
            rows.append([ws.cell_value(r, c) for c in range(ws.ncols)])
    else:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            rows.append(list(row))

    result = {}
    for row in rows:
        if not row or not row[0]:
            continue
        nome_raw = str(row[0]).strip()
        if len(nome_raw) < 3 or not re.search(r"[A-Za-zÀ-ÿ]{2}", nome_raw):
            continue
        # Procura o primeiro valor numerico na linha
        valor = 0.0
        for cell in row[1:]:
            if isinstance(cell, (int, float)) and cell > 0:
                valor = float(cell)
                break
            elif isinstance(cell, str):
                v = brl(cell)
                if v > 0:
                    valor = v
                    break
        if valor > 0:
            nn = norm(nome_raw)
            result[nn] = {
                "nome_original": nome_raw.title(),
                "mensalidade": valor,
                "mensalidade_dependentes": 0.0,
                "sos_tam": 0.0,
                "total": valor,
                "dependentes": 0,
            }
    return result
