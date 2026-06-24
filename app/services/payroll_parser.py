"""
Parsers de folha de pagamento — Excel, PDF (recibos) e Word (instrucoes)
"""
import io
import re

from app.core.utils import norm, brl
from app.services.rubrics import normalize_rubric


# ─────────────────────────────────────────────
# PARSER EXCEL
# ─────────────────────────────────────────────

def parse_excel(raw: bytes, filename: str) -> dict:
    """
    Le planilha de folha de pagamento.
    Retorna: {NOME_NORM: {salario, gratificacao, ferias_13, inss, vale, plano, emprestimo, liquido}}
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

    # Detecta colunas pelos cabecalhos
    col = {k: None for k in ["salario", "gratif", "ferias", "inss", "vale", "plano", "emprestimo", "liquido"]}
    # Colunas de apontamentos (planilha de eventos variaveis)
    apc = {}  # chave -> indice de coluna
    name_col = 0  # padrao: nomes na 1a coluna
    for row in rows[:10]:
        cells = [str(v or "").upper().strip() for v in row]
        for i, c in enumerate(cells):
            # detecta coluna de nomes
            if c in ("COLABORADORES", "NOME", "FUNCIONARIO", "FUNCIONÁRIO", "NOME DO FUNCIONARIO", "NOME DO FUNCIONÁRIO"):
                name_col = i
            if c in ("SALARIO", "SALÁRIO") and col["salario"] is None and i > 0:
                col["salario"] = i
            if ("GRATIF" in c or "ADICION" in c) and col["gratif"] is None and i > 0:
                col["gratif"] = i
            if ("FERIAS" in c or "FÉRIAS" in c or "13" in c) and col["ferias"] is None and i > 0:
                col["ferias"] = i
            if "INSS" in c and col["inss"] is None and i > 0:
                col["inss"] = i
            if "VALE" in c and col["vale"] is None and i > 0:
                col["vale"] = i
            if "PLANO" in c and col["plano"] is None and i > 0:
                col["plano"] = i
            if "EMPRESTIMO" in c and col["emprestimo"] is None and i > 0:
                col["emprestimo"] = i
            if ("LIQUIDO" in c or "LÍQUIDO" in c) and col["liquido"] is None and i > 0:
                col["liquido"] = i
            # Apontamentos
            if "ASSIDUIDADE" in c and "assiduidade" not in apc:
                apc["assiduidade"] = i
            if "PONTUALIDADE" in c and "pontualidade" not in apc:
                apc["pontualidade"] = i
            if "GRATIF" in c and "TEMPO" in c and "gratif_tempo" not in apc:
                apc["gratif_tempo"] = i
            if ("PREMIO" in c or "PRÊMIO" in c) and "premio" not in apc:
                apc["premio"] = i
            if "VALE ALIMENT" in c and "va_desconto" not in apc:
                apc["va_desconto"] = i
            if "DESCONTO FALTA" in c and "HORA" not in c and "falta" not in apc:
                apc["falta"] = i
            if "HORAS FALTA" in c and "horas_faltas" not in apc:
                apc["horas_faltas"] = i
            if "HORA EXTRA" in c and "hora_extra" not in apc:
                apc["hora_extra"] = i
            if "NOTURNO" in c and "noturno" not in apc:
                apc["noturno"] = i
            if "ADIANTAMENTO" in c and "adiantamento" not in apc:
                apc["adiantamento"] = i
            if ("FARMACIA" in c or "FARMÁCIA" in c) and "farmacia" not in apc:
                apc["farmacia"] = i

    SKIP = {"TOTAL", "NOME", "SALARIO", "SALÁRIO", "FUNCIONARIO", "FUNCIONÁRIO", "COLABORADORES", ""}
    SKIP_KW = ["LTDA", "EPP", "S/A", "CNPJ", "LISTA DE", "CENTRO MEDICO", "PAGAMENTO", "PLANILHA"]
    employees = {}

    import datetime as _dt

    for row in rows:
        first = str(row[name_col] if name_col < len(row) else "").strip()
        if not first or first.upper() in SKIP:
            continue
        if any(kw in first.upper() for kw in SKIP_KW):
            continue
        if re.match(r"^[\d\s.,\-/]+$", first):
            continue
        if len(first.split()) < 2:
            continue
        if not re.search(r"[A-Za-zÀ-ÿ]{3}", first):
            continue

        def g(k):
            ci = col[k]
            if ci is None or ci >= len(row):
                return 0.0
            v = row[ci]
            return float(v) if isinstance(v, (int, float)) else brl(v)

        def ga(k):
            """Le coluna de apontamento; retorna None se vazio."""
            ci = apc.get(k)
            if ci is None or ci >= len(row):
                return None
            v = row[ci]
            if v is None:
                return None
            if isinstance(v, _dt.datetime):
                return v.strftime("%d/%m/%Y")
            if isinstance(v, _dt.timedelta):
                ts = int(v.total_seconds())
                return f"{ts//3600}:{(ts%3600)//60:02d}"
            try:
                f = float(v)
                return f if f != 0 else None
            except Exception:
                s = str(v).strip()
                return s if s else None

        # fallback posicional — inclui zeros para manter posicoes corretas
        nums = [float(v) for v in row[name_col+1:] if isinstance(v, (int, float))]

        aponts = {k: ga(k) for k in apc if ga(k) is not None}

        employees[norm(first)] = {
            "salario":      g("salario")     or (nums[0] if nums else 0),
            "gratificacao": g("gratif")      or (nums[1] if len(nums) > 1 else 0),
            "ferias_13":    g("ferias")      or (nums[2] if len(nums) > 2 else 0),
            "inss":         g("inss"),
            "vale":         g("vale"),
            "plano":        g("plano"),
            "emprestimo":   g("emprestimo"),
            "liquido":      g("liquido")     or (nums[-1] if nums else 0),
            "has_liquido":  col["liquido"] is not None,
            "apontamentos": aponts,
        }

    return employees


# ─────────────────────────────────────────────
# PARSER PDF (Recibos)
# ─────────────────────────────────────────────

def parse_pdf(raw: bytes) -> dict:
    """
    Le PDF de recibos de pagamento.
    Retorna: {NOME_NORM: {liquido, total_vencimentos, total_descontos, verbas, tipo, has_gratif}}
    """
    import pdfplumber

    employees = {}
    seen = set()

    with pdfplumber.open(io.BytesIO(raw)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            # ── Extrai nome do funcionario ──────────────────────────────────
            # Padrao: codigo numerico + NOME EM CAPS + CBO (6 digitos)
            # Exemplo: "58 ANDREIA PEREIRA BARBOSA 514320 1 1"
            nm = re.search(
                r"\b\d{1,3}\s+([A-Z][A-Z ]{4,50}?)\s+\d{6}\b",
                text,
            )

            if not nm:
                continue

            emp_name = nm.group(1).strip()
            emp_norm = norm(emp_name)

            if emp_norm in seen or len(emp_norm.split()) < 2:
                continue
            seen.add(emp_norm)

            # ── Tipo de recibo ──────────────────────────────────────────────
            tipo = "mensal"
            if re.search(r"13[oOº°].*adiantamento|adiantamento.*13", text[:600], re.I):
                tipo = "13_adiantamento"
            elif re.search(r"f[eé]rias", text[:600], re.I):
                tipo = "ferias"

            # ── Totais ──────────────────────────────────────────────────────
            # "Valor Liquido 1.143,17"  (o \uf0f0 e a seta Wingdings do sistema)
            liq_m = re.search(
                r"Valor\s+L[íi]quido[^0-9]*([\d.,]+)",
                text, re.I,
            )
            # "Total de Descontos\n883,08" — aparece ANTES do liquido no texto
            td_m = re.search(
                r"Total\s+de\s+Descontos[^0-9]*([\d.,]+)",
                text, re.I,
            )

            liquido    = brl(liq_m.group(1)) if liq_m else 0.0
            total_desc = brl(td_m.group(1)) if td_m else 0.0
            total_venc = round(liquido + total_desc, 2)   # sempre verdadeiro contabilmente

            # ── Verbas ──────────────────────────────────────────────────────
            verbas = []
            # Padrao: "8781 DIAS NORMAIS 30,00 1.621,00"
            # Referencia aceita ponto (ex: "1.200,00" para Unimed/emprestimos)
            for m in re.finditer(
                r"(\d{2,4})\s+([A-ZÀ-Ü][A-ZÀ-Ü .%º/°()]+?)\s+([\d:.,]+)\s+([\d.,]+)",
                text,
            ):
                desc = m.group(2).strip()
                ref  = m.group(3)
                # Referencia legitima: tem virgula (valor monetario) OU dois pontos (horas, ex: "3:41")
                if "," not in ref and ":" not in ref:
                    continue
                # Ignora palavras-chave que nao sao verbas
                if any(kw in desc for kw in ["CNPJ", "BASE CALC", "SAL. CONTR", "F.G.T.S"]):
                    continue
                valor = brl(m.group(4))
                # Valores < 1 sao ruidos (ex: filial = "1")
                if valor < 1:
                    continue
                verbas.append({
                    "codigo": m.group(1),
                    "descricao": desc,
                    "referencia": ref,
                    "valor": valor,
                })

            # Captura verbas sem codigo (ex: "DESCONTO VALE ALIMENTACAO 18,00 18,00")
            _SKIP_CODELESS = {"CODIGO", "DESCRIÇÃO", "DESCRICAO", "TOTAL", "REFERENCIA", "REFERÊNCIA",
                              "VALOR LIQUIDO", "VALOR LÍQUIDO", "BASE CALC", "SAL.", "SALARIO", "SALÁRIO"}
            for m in re.finditer(
                r"^([A-ZÀ-Ü][A-ZÀ-Ü ]{3,50}?)\s+([\d:.,]+)\s+([\d.,]+)\s*$",
                text, re.MULTILINE
            ):
                desc_c = m.group(1).strip()
                ref_c  = m.group(2)
                if "," not in ref_c and ":" not in ref_c:
                    continue
                if len(desc_c.split()) < 2:
                    continue
                if any(kw in desc_c.upper() for kw in _SKIP_CODELESS):
                    continue
                if any(kw in desc_c.upper() for kw in ["CNPJ", "F.G.T.S", "BASE"]):
                    continue
                # Evita duplicar verbas ja capturadas pelo padrao com codigo
                if any(v["descricao"].upper() == desc_c.upper() for v in verbas):
                    continue
                valor_c = brl(m.group(3))
                if valor_c < 1:
                    continue
                verbas.append({
                    "codigo": "",
                    "descricao": desc_c,
                    "referencia": ref_c,
                    "valor": valor_c,
                })

            has_gratif = any(
                re.search(r"GRATIF|PREMIACAO|PREMIO", v["descricao"].upper())
                for v in verbas
            )
            gratif_valor = sum(
                v["valor"] for v in verbas
                if re.search(r"GRATIF|PREMIACAO|PREMIO", v["descricao"].upper())
            )

            employees[emp_norm] = {
                "nome_original":    emp_name.title(),
                "tipo":             tipo,
                "liquido":          liquido,
                "total_vencimentos": total_venc,
                "total_descontos":  total_desc,
                "has_gratif":       has_gratif,
                "gratif_valor":     gratif_valor,
                "verbas":           verbas,
            }

    return employees


# ─────────────────────────────────────────────
# PARSER WORD
# ─────────────────────────────────────────────

def parse_word(raw: bytes) -> dict:
    """
    Le documento Word com instrucoes da folha.
    Retorna: {gratificacoes, descontos, obs, decimo_terceiro}
    """
    from docx import Document

    doc = Document(io.BytesIO(raw))
    text = "\n".join(p.text for p in doc.paragraphs)

    result = {"gratificacoes": {}, "descontos": {}, "obs": [], "decimo_terceiro": [], "comissao_e_dsr": {}}

    # Gratificacoes
    m = re.search(
        r"Gratifica[çc][õo]es?:?(.*?)(?:Funcionar|Férias|Descontos|Goiânia|$)",
        text, re.I | re.S,
    )
    if m:
        for entry in re.finditer(
            r"([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\s]{3,40}?)\s*[\-–\(]?\s*\(?([\d.,]{3,})\)?",
            m.group(1),
        ):
            name, val = entry.group(1).strip(), brl(entry.group(2))
            if val > 0 and len(name.split()) >= 2:
                result["gratificacoes"][norm(name)] = val

    # Descontos (Unimed, etc.)
    m2 = re.search(r"Descontos?:?(.*?)(?:Goiânia|Ass|$)", text, re.I | re.S)
    if m2:
        tipo_atual = "Desconto"
        for line in m2.group(1).split("\n"):
            lt = line.strip()
            if not lt:
                continue
            if re.match(r"^([A-Za-zÀ-ÿ]+):?\s*$", lt):
                tipo_atual = lt.rstrip(":")
                continue
            entry = re.search(
                r"([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\s]{4,40}?)\s*[\-–]\s*\(?([\d.,]{3,})\)?", lt
            )
            if entry:
                name, val = entry.group(1).strip(), brl(entry.group(2))
                if val > 0 and len(name.split()) >= 2:
                    nn = norm(name)
                    result["descontos"].setdefault(nn, {})[tipo_atual] = val

    # Adiantamento 13o
    result["decimo_terceiro"] = [
        norm(x)
        for x in re.findall(
            r"(?:13[oOº°]|décimo\s+terceiro)[^\w]*([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\s]{5,40})",
            text, re.I,
        )
        if len(x.split()) >= 2
    ]

    # Observacoes
    result["obs"] = [
        l.strip() for l in text.split("\n") if re.match(r"OBS|Obs", l.strip())
    ]

    # Fallback: formato lista simples (titulo na 1a linha, depois nome\nvalor alternados)
    # Ex: "comissoes e DSR\nNadyane\n1592,11\nJuliana\n1255,49"
    if not result["gratificacoes"] and not result["descontos"]:
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        if len(lines) >= 3:
            # Detecta o tipo da secao pelo titulo (linha 0)
            titulo_norm = normalize_rubric(lines[0])
            is_comissao_dsr = titulo_norm in ("COMISSAO_E_DSR", "COMISSAO E DSR")

            # Detecta se e padrao nome/valor: linhas alternadas texto/numero
            pares = []
            i = 1  # pula titulo (linha 0)
            while i < len(lines) - 1:
                nome_linha = lines[i]
                val_linha = lines[i + 1]
                val_clean = val_linha.replace(".", "").replace(",", ".").replace(" ", "")
                if re.match(r"^\d+(\.\d+)?$", val_clean):
                    pares.append((nome_linha, val_linha))
                    i += 2
                else:
                    i += 1
            if pares:
                for nome, val_str in pares:
                    val = brl(val_str)
                    name = norm(nome)
                    if val > 0 and len(name) >= 2:  # aceita nome com 1 palavra
                        if is_comissao_dsr:
                            result["comissao_e_dsr"][name] = val
                        else:
                            result["gratificacoes"][name] = val

    return result
