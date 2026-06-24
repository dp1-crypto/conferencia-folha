"""
Comparador de beneficios — fatura vs extrato de folha
"""
from app.core.utils import fmt_brl


def _abbrev_match(fw_list: list, ew_list: list) -> bool:
    """
    Verifica se fw_list (nomes da fatura, possivelmente abreviados) representa
    a mesma pessoa que ew_list (nomes completos do extrato).
    Regras:
    - Primeira palavra deve coincidir exatamente
    - Palavras de 1 letra sao tratadas como inicial — "O" bate com "OLIVEIRA"
    - Palavras do extrato podem ser puladas (nomes do meio ausentes na fatura)
    """
    if not fw_list or not ew_list or fw_list[0] != ew_list[0]:
        return False
    fi, ei = 1, 1
    while fi < len(fw_list) and ei < len(ew_list):
        fw, ew = fw_list[fi], ew_list[ei]
        if fw == ew:
            fi += 1; ei += 1
        elif len(fw) == 1 and ew.startswith(fw):
            fi += 1; ei += 1
        else:
            ei += 1  # pula palavra do extrato (nome do meio nao abreviado)
    return fi == len(fw_list)


def match_names_beneficio(fatura: dict, extrato: dict) -> dict:
    """
    Casa nomes da fatura (truncados/abreviados) com nomes completos do extrato.
    Retorna: {fatura_key -> extrato_key}
    """
    mapping = {}
    used = set()

    for fk in fatura:
        fk_words = fk.split()
        best = None

        # 1) Prefixo exato (fatura e prefixo do extrato)
        for ek in extrato:
            if ek in used: continue
            ek_words = ek.split()
            if ek_words[:len(fk_words)] == fk_words:
                best = ek; break

        # 2) Match com abreviacoes (ex: "christiane o santos" = "christiane oliveira dos santos")
        if not best:
            for ek in extrato:
                if ek in used: continue
                if _abbrev_match(fk_words, ek.split()):
                    best = ek; break

        # 3) Primeiras 2 palavras exatas (fallback)
        if not best:
            for ek in extrato:
                if ek in used: continue
                if ek.split()[:2] == fk_words[:2]:
                    best = ek; break

        if best:
            mapping[fk] = best
            used.add(best)

    return mapping


def compare_plano_saude(fatura: dict, extrato: dict, regra: dict = None) -> dict:
    """
    Compara valor esperado (fatura ou regra) com o que foi descontado no extrato.
    regra: {"tipo": "fatura"|"pct_fatura"|"pct_salario"|"fixo", "valor": float}
    """
    regra = regra or {"tipo": "fatura", "valor": 0.0}
    name_map = match_names_beneficio(fatura, extrato)
    rev_map  = {v: k for k, v in name_map.items()}

    todos = set(fatura.keys()) | {rev_map.get(ek, ek) for ek in extrato}

    resultados = []
    total_esperado = 0.0
    total_extrato  = 0.0
    divergentes    = 0

    def calc_esperado(fat, ext):
        """Calcula valor esperado baseado na regra."""
        tipo = regra["tipo"]
        pct  = regra["valor"] / 100.0
        fixo = regra["valor"]
        if tipo == "fatura":
            return (fat or {}).get("total", 0.0)
        elif tipo == "pct_fatura":
            return round((fat or {}).get("total", 0.0) * pct, 2)
        elif tipo == "pct_salario":
            return round((ext or {}).get("salario", 0.0) * pct, 2)
        elif tipo == "fixo":
            return fixo if (ext or fat) else 0.0
        return 0.0

    processed_ek = set()  # rastreia extrato keys ja processadas

    for fk in sorted(todos):
        ek  = name_map.get(fk)
        fat = fatura.get(fk)
        ext = extrato.get(ek) if ek else None
        if not ext and fk in extrato:
            ext = extrato[fk]
            ek  = fk

        if ek:
            processed_ek.add(ek)

        sem_fatura  = fat is None
        sem_extrato = ext is None
        nome_exibir = (fat or {}).get("nome_original") or (ext or {}).get("nome_original") or fk.title()

        val_esperado   = calc_esperado(fat, ext)
        val_descontado = (ext or {}).get("plano_descontado", 0.0)
        diferenca      = round(val_esperado - val_descontado, 2)

        if abs(diferenca) <= 0.05:
            status = "OK"
        elif diferenca > 0:
            status = "MAIOR"
        else:
            status = "MENOR"

        if status != "OK" or sem_fatura or sem_extrato:
            divergentes += 1

        total_esperado += val_esperado
        total_extrato  += val_descontado

        resultados.append({
            "nome":                   nome_exibir,
            "mensalidade_titular":    (fat or {}).get("mensalidade", 0.0),
            "mensalidade_dependentes":(fat or {}).get("mensalidade_dependentes", 0.0),
            "sos_tam":                (fat or {}).get("sos_tam", 0.0),
            "total_fatura":           (fat or {}).get("total", 0.0),
            "salario":                (ext or {}).get("salario", 0.0),
            "valor_esperado":         val_esperado,
            "dependentes":            (fat or {}).get("dependentes", 0),
            "valor_descontado":       val_descontado,
            "diferenca":              diferenca,
            "status":                 status,
            "sem_fatura":             sem_fatura,
            "sem_extrato":            sem_extrato,
        })

    # Funcionarios so no extrato, nao casados com fatura (evita duplicatas)
    for ek in extrato:
        if ek in processed_ek:
            continue  # ja processado no loop principal
        if rev_map.get(ek) is None and ek not in fatura:
            ext = extrato[ek]
            val_descontado = ext.get("plano_descontado", 0.0)
            val_esperado   = calc_esperado(None, ext)
            diferenca      = round(val_esperado - val_descontado, 2)
            total_esperado += val_esperado
            total_extrato  += val_descontado
            divergentes    += 1
            resultados.append({
                "nome": ext.get("nome_original", ek.title()),
                "mensalidade_titular": 0.0, "mensalidade_dependentes": 0.0,
                "sos_tam": 0.0, "total_fatura": 0.0,
                "salario": ext.get("salario", 0.0),
                "valor_esperado": val_esperado,
                "dependentes": 0,
                "valor_descontado": val_descontado,
                "diferenca": diferenca,
                "status": "MENOR" if val_descontado > 0 else "OK",
                "sem_fatura": True, "sem_extrato": False,
            })

    resultados.sort(key=lambda x: (x["status"] == "OK", x["nome"]))

    return {
        "resultados":      resultados,
        "total_esperado":  round(total_esperado, 2),
        "total_extrato":   round(total_extrato, 2),
        "total_diferenca": round(total_esperado - total_extrato, 2),
        "divergentes":     divergentes,
        "total":           len(resultados),
        "regra":           regra,
    }
