"""
Motor de comparacao — Excel vs PDF vs Word
"""
from app.core.utils import fmt_brl
from app.core.config import TOLERANCIA_CENTAVOS, TOLERANCIA_DIVERGENCIA, STATUS
from app.services.rubrics import normalize_rubric, find_rubric_by_value


def match_names(excel: dict, pdf: dict) -> dict:
    """
    Casa nomes do Excel (abreviados) com nomes do PDF (completos).
    Retorna: {excel_name: pdf_name}
    """
    mapping = {}
    used_pdf = set()

    for en in excel:
        en_words = en.split()
        best = None
        # Tenta prefixo exato: "ANDREIA PEREIRA" -> "ANDREIA PEREIRA BARBOSA"
        for pn in pdf:
            if pn in used_pdf:
                continue
            pn_words = pn.split()
            if pn_words[:len(en_words)] == en_words:
                best = pn
                break
        # Fallback: verifica se as 2 primeiras palavras batem
        if not best:
            for pn in pdf:
                if pn in used_pdf:
                    continue
                pn_w = pn.split()
                en_w = en.split()
                if pn_w[:2] == en_w[:2]:
                    best = pn
                    break
        if best:
            mapping[en] = best
            used_pdf.add(best)

    return mapping


def compare(excel: dict, pdf: dict, word: dict) -> dict:
    # Casa nomes Excel (abreviados) <-> PDF (completos)
    name_map = match_names(excel, pdf)          # excel_name -> pdf_name
    rev_map  = {v: k for k, v in name_map.items()}  # pdf_name -> excel_name

    # Todos os funcionarios (usando nomes do Excel como chave canonica quando possivel)
    canonical = {}  # canonical_name -> {excel_key, pdf_key}
    for en in excel:
        pn = name_map.get(en)
        canonical[en] = {"excel_key": en, "pdf_key": pn}
    for pn in pdf:
        en = rev_map.get(pn)
        if not en:
            canonical[pn] = {"excel_key": None, "pdf_key": pn}

    all_names = sorted(canonical)

    report = {
        "resumo": {"total": len(all_names), "divergencias": 0, "ok": 0},
        "funcionarios": [],
        "observacoes": (word or {}).get("obs", []),
        "word_gratificacoes": {k: fmt_brl(v) for k, v in (word or {}).get("gratificacoes", {}).items()},
        "word_descontos": (word or {}).get("descontos", {}),
    }

    for name in all_names:
        keys    = canonical[name]
        exc_key = keys["excel_key"]
        pdf_key = keys["pdf_key"]
        exc = excel.get(exc_key) if exc_key else None
        rec = pdf.get(pdf_key)   if pdf_key else None

        nome_exibir = (rec or {}).get("nome_original") or (exc_key or name).title()

        emp = {
            "nome":        name,
            "nome_exibir": nome_exibir,
            "status":      "OK",
            "divs":        [],
            "dados_excel": exc,
            "dados_recibo": rec,
        }

        # ── Presenca ────────────────────────────────────────────────────────
        # So aponta ausencia se o respectivo tipo de arquivo foi enviado
        if not exc and excel:
            emp["divs"].append({
                "g": "alta",
                "tipo": "Ausente na planilha",
                "desc": "Funcionario tem recibo mas nao esta na planilha Excel.",
            })
        if not rec and pdf:
            emp["divs"].append({
                "g": "alta",
                "tipo": "Sem recibo PDF",
                "desc": "Funcionario esta na planilha mas nao ha recibo PDF.",
            })

        # ── Valor liquido ───────────────────────────────────────────────────
        if exc and rec and exc.get("has_liquido", True):
            el = exc.get("liquido", 0)
            rl = rec.get("liquido", 0)
            if el > 0 and abs(el - rl) > TOLERANCIA_DIVERGENCIA:
                emp["divs"].append({
                    "g": "alta",
                    "tipo": "Liquido divergente",
                    "desc": (
                        f"Planilha: {fmt_brl(el)} | Recibo: {fmt_brl(rl)} "
                        f"| Diferenca: {fmt_brl(abs(el - rl))}"
                    ),
                })

        # ── Apontamentos vs Verbas do recibo ────────────────────────────────
        if exc and rec:
            aponts = exc.get("apontamentos", {})
            verbas = rec.get("verbas", [])

            def _find_verba(keywords, codigo=None):
                # Busca por codigo + keyword (evita falsos positivos de parsing)
                if codigo:
                    for v in verbas:
                        if (v.get("codigo") == str(codigo) and
                                any(kw.upper() in v.get("descricao", "").upper() for kw in keywords)):
                            return v
                # Fallback: keyword apenas
                for v in verbas:
                    if any(kw.upper() in v.get("descricao", "").upper() for kw in keywords):
                        return v
                return None

            def _check_bonus(chave, label, keywords, codigo=None):
                val = aponts.get(chave)
                if not val:
                    return
                try:
                    val = float(val)
                except Exception:
                    return
                v = _find_verba(keywords, codigo)
                if not v:
                    # ── Smart matching: busca por valor quando nome nao reconhecido ──
                    match = find_rubric_by_value(label, val, verbas)
                    if match and match["confianca"] in ("alta", "media"):
                        vb = match["verba"]
                        diff = round(abs(vb["valor"] - val), 2)
                        emp.setdefault("sugestoes_equivalencia", []).append({
                            "esperado":   label,
                            "encontrado": vb["descricao"],
                            "valor":      val,
                            "confianca":  match["confianca"],
                            "dica":       f"Adicione '{vb['descricao']}' ao grupo correspondente em rubricas-equivalentes.json",
                        })
                        if diff <= TOLERANCIA_CENTAVOS:
                            emp["divs"].append({
                                "g":    "manual",
                                "tipo": STATUS["POSSIVEL_RUBRICA"],
                                "desc": (
                                    f"'{label}' nao encontrado pelo nome, mas "
                                    f"'{vb['descricao']}' tem o mesmo valor {fmt_brl(val)}. "
                                    f"Verifique se sao equivalentes e adicione ao config."
                                ),
                            })
                        else:
                            emp["divs"].append({
                                "g":    "media",
                                "tipo": STATUS["POSSIVEL_RUBRICA"],
                                "desc": (
                                    f"'{label}' nao encontrado pelo nome. Possivel equivalente: "
                                    f"'{vb['descricao']}' ({fmt_brl(vb['valor'])} vs esperado {fmt_brl(val)}). "
                                    f"Diferenca: {fmt_brl(diff)}."
                                ),
                            })
                    else:
                        emp["divs"].append({
                            "g":    "alta",
                            "tipo": STATUS["NAO_LOCALIZADO_RECIBO"],
                            "desc": f"Planilha indica {label} de {fmt_brl(val)}, mas nao encontrado no recibo.",
                        })
                elif abs(v.get("valor", 0) - val) <= TOLERANCIA_CENTAVOS:
                    pass  # arredondamento — nao reporta
                elif abs(v.get("valor", 0) - val) > TOLERANCIA_DIVERGENCIA:
                    diff_v = v.get("valor", 0)
                    emp["divs"].append({
                        "g":    "media",
                        "tipo": STATUS["A_PAGAR"] if diff_v < val else STATUS["PAGO_MAIOR"],
                        "desc": f"Planilha: {fmt_brl(val)} | Recibo: {fmt_brl(diff_v)} | Diferenca: {fmt_brl(abs(diff_v - val))}",
                    })

            _check_bonus("pontualidade", "Pontualidade", ["PONTUALIDADE"], "221")
            _check_bonus("assiduidade",  "Assiduidade",  ["ASSIDUIDADE"],  "222")
            _check_bonus("gratif_tempo", "Gratificacao Tempo de Servico", ["GRATIF", "TEMPO"], "228")
            _check_bonus("premio",       "Premio",       ["PREMIO", "PREMIA"], None)
            _check_bonus("va_desconto",  "Vale Alimentacao", ["VALE ALIMENT"], "204")
            _check_bonus("adiantamento", "Adiantamento Salarial", ["ADIANT"], None)
            _check_bonus("farmacia",     "Farmacia", ["FARMACIA", "FARMÁCIA"], None)

            # Falta: so verifica presenca da verba, nao o valor (que e data na planilha)
            if aponts.get("falta"):
                v = _find_verba(["FALTA", "DIAS FALTA", "DESCONTO FALTA"], "8792")
                if not v:
                    emp["divs"].append({
                        "g": "media",
                        "tipo": "Desconto de falta ausente no recibo",
                        "desc": f"Planilha indica falta em {aponts['falta']}, mas nao ha desconto de falta no recibo.",
                    })

            # Horas faltas: verifica presenca
            if aponts.get("horas_faltas"):
                v = _find_verba(["HORA FALTA", "HORAS FALTA", "FALTAS PARC"], "8069")
                if not v:
                    emp["divs"].append({
                        "g": "media",
                        "tipo": "Horas de falta parcial ausente no recibo",
                        "desc": f"Planilha indica {aponts['horas_faltas']}h de falta parcial, mas nao encontrado no recibo.",
                    })

            # Hora extra: verifica presenca
            if aponts.get("hora_extra"):
                v = _find_verba(["HORA EXTRA", "H.EXTRA", "HORAS EXTRA"])
                if not v:
                    emp["divs"].append({
                        "g": "media",
                        "tipo": "Hora extra ausente no recibo",
                        "desc": f"Planilha indica {aponts['hora_extra']}h extra, mas nao encontrado no recibo.",
                    })

            # Adicional noturno: verifica presenca
            if aponts.get("noturno"):
                v = _find_verba(["NOTURNO", "ADICIONAL NOT"])
                if not v:
                    emp["divs"].append({
                        "g": "media",
                        "tipo": "Adicional noturno ausente no recibo",
                        "desc": f"Planilha indica {aponts['noturno']}h noturno, mas nao encontrado no recibo.",
                    })

        # ── Gratificacoes (Word -> Recibo) ────────────────────────────────────
        if word:
            # Word pode usar nome abreviado, completo ou apenas primeiro nome
            wg = word.get("gratificacoes", {})
            gratif = wg.get(exc_key or name) or wg.get(pdf_key or name) or 0
            # Fallback: match por primeiro nome (quando Word usa so "NADYANE")
            if not gratif:
                first_name = (exc_key or pdf_key or name or "").split()[0] if (exc_key or pdf_key or name) else ""
                if first_name:
                    for wk, wv in wg.items():
                        if wk.split()[0] == first_name:
                            gratif = wv
                            break
            if gratif > 0:
                if rec and not rec.get("has_gratif", False):
                    emp["divs"].append({
                        "g": "alta",
                        "tipo": "Gratificacao ausente no recibo",
                        "desc": (
                            f"Gratificacao de {fmt_brl(gratif)} consta no Word "
                            f"mas NAO aparece no recibo como verba separada."
                        ),
                    })
                exc_g = (exc or {}).get("gratificacao", 0)
                if exc_g > 0 and abs(exc_g - gratif) > TOLERANCIA_DIVERGENCIA:
                    emp["divs"].append({
                        "g": "media",
                        "tipo": "Valor de gratificacao divergente",
                        "desc": f"Word: {fmt_brl(gratif)} | Planilha: {fmt_brl(exc_g)}",
                    })

            # ── Descontos especiais (Word -> Recibo) ─────────────────────────
            wd = word.get("descontos", {})
            word_descs = wd.get(exc_key or name) or wd.get(pdf_key or name) or {}
            for tipo_desc, val in word_descs.items():
                verbas = (rec or {}).get("verbas", [])
                found = any(
                    tipo_desc.upper() in v.get("descricao", "").upper()
                    or abs(v.get("valor", 0) - val) < TOLERANCIA_DIVERGENCIA
                    for v in verbas
                )
                if not found and rec:
                    emp["divs"].append({
                        "g": "media",
                        "tipo": f"Desconto '{tipo_desc}' nao localizado",
                        "desc": (
                            f"Desconto {tipo_desc} de {fmt_brl(val)} mencionado no Word "
                            f"nao foi encontrado no recibo."
                        ),
                    })

        # ── Comissao + DSR (Word -> Recibo) ──────────────────────────────────
        if word:
            cdsr = word.get("comissao_e_dsr", {})
            val_word = cdsr.get(exc_key or name) or cdsr.get(pdf_key or name) or 0
            if not val_word:
                first = (exc_key or pdf_key or name or "").split()[0] if (exc_key or pdf_key or name) else ""
                if first:
                    for wk, wv in cdsr.items():
                        if wk.split()[0] == first:
                            val_word = wv
                            break
            if val_word and val_word > 0 and rec:
                val_recibo = sum(
                    v["valor"] for v in rec.get("verbas", [])
                    if normalize_rubric(v["descricao"]) in ("COMISSAO", "DSR", "COMISSAO_E_DSR")
                )
                diff = round(abs(val_word - val_recibo), 2)
                mem = {
                    "word_valor": val_word,
                    "recibo_verbas": [
                        {"desc": v["descricao"], "valor": v["valor"]}
                        for v in rec.get("verbas", [])
                        if normalize_rubric(v["descricao"]) in ("COMISSAO", "DSR", "COMISSAO_E_DSR")
                    ],
                    "recibo_total": val_recibo,
                    "diferenca": diff,
                }
                if diff <= TOLERANCIA_CENTAVOS:
                    emp["memoria_comissao_dsr"] = {**mem, "status": STATUS["OK"]}
                elif diff <= TOLERANCIA_DIVERGENCIA:
                    emp["memoria_comissao_dsr"] = {**mem, "status": STATUS["ARREDONDAMENTO"]}
                elif val_word > val_recibo:
                    emp["divs"].append({
                        "g": "alta",
                        "tipo": STATUS["A_PAGAR"],
                        "desc": (
                            f"Comissao+DSR -- Relatorio: {fmt_brl(val_word)} | "
                            f"Recibo: {fmt_brl(val_recibo)} | Falta: {fmt_brl(diff)}"
                        ),
                        "memoria": mem,
                    })
                else:
                    emp["divs"].append({
                        "g": "media",
                        "tipo": STATUS["PAGO_MAIOR"],
                        "desc": (
                            f"Comissao+DSR -- Relatorio: {fmt_brl(val_word)} | "
                            f"Recibo: {fmt_brl(val_recibo)} | Excesso: {fmt_brl(diff)}"
                        ),
                        "memoria": mem,
                    })

        if emp["divs"]:
            emp["status"] = "DIVERGENTE"
            report["resumo"]["divergencias"] += 1
        else:
            report["resumo"]["ok"] += 1

        report["funcionarios"].append(emp)

    # ── Possiveis homonimos ──────────────────────────────────────────────
    possiveis = []
    nomes = [c for c in all_names if canonical[c]["pdf_key"]]
    for i in range(len(nomes)):
        for j in range(i + 1, len(nomes)):
            a_words = nomes[i].split()
            b_words = nomes[j].split()
            if (len(a_words) >= 2 and len(b_words) >= 2
                    and a_words[:2] == b_words[:2] and nomes[i] != nomes[j]):
                possiveis.append({
                    "nomes": [nomes[i], nomes[j]],
                    "aviso": "Possivel homonimo: mesmas 2 primeiras palavras",
                })
    report["possiveis_homonimos"] = possiveis

    return report
