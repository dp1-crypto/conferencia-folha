"""
Analise de folha para implantacao de cliente novo
Sigma Contabilidade

Recebe N meses de recibos ja extraidos e responde as perguntas que importam
para migrar a folha sem erro:
  - quem sao os funcionarios e o que cada um recebe;
  - quais rubricas sao FIXAS (cadastrar como evento fixo no sistema) e quais
    sao VARIAVEIS (lancar mes a mes);
  - o que MUDOU no periodo (admissao, desligamento, reajuste, rubrica que
    entrou ou parou);
  - o que precisa de decisao humana antes de importar.
"""
import re
import statistics
from collections import defaultdict

from app.core.textutils import flat, competencia_range, competencia_label
from app.core.config import TOLERANCIA_CENTAVOS
from app.services import rubrics
from app.services.holerite_parser import colapsa_siglas, resolver_grupo

# Percentual de meses ativos a partir do qual a rubrica e considerada recorrente
LIMIAR_RECORRENTE = 0.6
# Variacao acima disso em rubrica variavel vira alerta de salto
LIMIAR_SALTO = 0.5

# Rubricas que o sistema de destino RECALCULA sozinho a partir da base — nao
# entram na planilha de importacao de eventos. Separa-las evita que a lista de
# "variaveis a lancar" venha poluida com INSS e IRRF, que variam todo mes por
# definicao mas nunca sao digitados.
RX_CALCULADA = re.compile(
    r"^(INSS|I N S S|IRRF|I R R F|IRPF|IMPOSTO DE RENDA|IR RETIDO|"
    r"FGTS|F G T S|CONTRIBUICAO PREVIDENCIARIA|PREVIDENCIA SOCIAL|"
    r"ARREDONDAMENTO|BASE )"
)


# ─────────────────────────────────────────────
# CHAVES
# ─────────────────────────────────────────────

def chave_funcionario(rec: dict) -> str:
    """CPF > matricula > nome. CPF e a unica chave que sobrevive a homonimo."""
    if rec.get("cpf"):
        return "cpf:" + rec["cpf"]
    if rec.get("matricula"):
        return "mat:" + str(rec["matricula"])
    return "nome:" + (rec.get("nome_norm") or flat(rec.get("nome", "")))


def chave_rubrica(v: dict) -> str:
    """
    Identidade da rubrica ao longo dos meses.

    O CODIGO de origem vem primeiro por dois motivos: e o que o sistema
    anterior usa de fato para identificar a verba, e nao sofre com a quebra de
    texto do PDF — 'Seg.de Vida em Grupo' e 'S eg.de Vida em Grupo' sao a mesma
    rubrica 351, mas pela descricao virariam duas, e o de-para sairia dobrado.
    """
    cod = str(v.get("codigo") or "").strip()
    if cod:
        return "cod:" + cod
    grupo = v.get("grupo") or ""
    if grupo and grupo in rubrics.RUBRIC_META:
        return grupo
    return colapsa_siglas(flat(v.get("descricao", ""))) or "(sem descricao)"


def melhor_grafia(nomes) -> str:
    """
    Escolhe a grafia mais confiavel entre as vistas para a mesma pessoa.

    A extracao quebra palavras ('Benedito De Deus Andrad E'), entao vence o
    nome com menos fragmentos soltos e, em empate, o mais completo.
    """
    validos = [n for n in nomes if n and n != "(nao identificado)"]
    if not validos:
        return "(nao identificado)"
    def penalidade(n):
        pedacos = n.split()
        soltos = sum(1 for p in pedacos if len(p) <= 2 and p.upper() not in
                     ("DE", "DA", "DO", "DAS", "DOS", "E"))
        # Menos pedacos = menos quebra: 'Seg.de Vida em Grupo' ganha de
        # 'Seg .de Vida em Grupo', que tem o mesmo texto com um corte a mais.
        return (soltos, len(pedacos), -len(n))
    return sorted(validos, key=penalidade)[0]


# ─────────────────────────────────────────────
# CONSOLIDACAO
# ─────────────────────────────────────────────

def consolidar(documentos: list) -> dict:
    """
    Junta todos os documentos lidos numa visao por funcionario x competencia.

    documentos = lista de retornos de holerite_parser.parse_holerites
    """
    funcionarios = {}
    comps = set()
    empresa = {"nome": "", "cnpj": ""}
    qualidade = {
        "arquivos": len(documentos), "recibos": 0, "recibos_ok": 0,
        "recibos_erro": 0, "recibos_sem_referencia": 0, "recibos_ocr": 0,
        "nomes_baixa_confianca": 0, "layouts": {},
    }
    problemas = []

    resumos = []
    for doc in documentos:
        for res in doc.get("resumos_empresa", []):
            if res.get("proventos"):
                resumos.append(res)

    for doc in documentos:
        if doc.get("empresa", {}).get("nome") and not empresa["nome"]:
            empresa["nome"] = doc["empresa"]["nome"]
        if doc.get("empresa", {}).get("cnpj") and not empresa["cnpj"]:
            empresa["cnpj"] = doc["empresa"]["cnpj"]

        lay = doc.get("layout", {}).get("nome", "?")
        qualidade["layouts"][lay] = qualidade["layouts"].get(lay, 0) + 1

        for aviso in doc.get("avisos", []):
            problemas.append({"arquivo": doc.get("arquivo", ""), "mensagem": aviso, "gravidade": "media"})

        for rec in doc.get("recibos", []):
            qualidade["recibos"] += 1
            nivel = rec["integridade"]["nivel"]
            if nivel == "ok":
                qualidade["recibos_ok"] += 1
            elif nivel == "sem_referencia":
                qualidade["recibos_sem_referencia"] += 1
            else:
                qualidade["recibos_erro"] += 1
                problemas.append({
                    "arquivo": doc.get("arquivo", ""),
                    "funcionario": rec["nome"],
                    "competencia": rec["competencia"],
                    "mensagem": "Leitura nao fecha com os totais impressos: "
                                + "; ".join(rec["integridade"]["mensagens"]),
                    "gravidade": "alta",
                })
            if rec.get("ocr"):
                qualidade["recibos_ocr"] += 1
            if rec.get("confianca_nome", 0) < 0.7:
                qualidade["nomes_baixa_confianca"] += 1
                problemas.append({
                    "arquivo": doc.get("arquivo", ""),
                    "funcionario": rec["nome"],
                    "competencia": rec["competencia"],
                    "mensagem": f"Nome do funcionario identificado com baixa confianca "
                                f"({int(rec.get('confianca_nome', 0) * 100)}%) — confira antes de importar.",
                    "gravidade": "alta",
                })

            comp = rec.get("competencia")
            if comp:
                comps.add(comp)

            k = chave_funcionario(rec)
            f = funcionarios.setdefault(k, {
                "chave": k, "nome": rec["nome"], "matricula": rec.get("matricula", ""),
                "cpf": rec.get("cpf", ""), "cbo": rec.get("cbo", ""),
                "funcao": rec.get("funcao", ""), "admissao": rec.get("admissao", ""),
                "meses": {}, "nomes_vistos": set(),
            })
            f["nomes_vistos"].add(rec["nome"])
            for campo in ("matricula", "cpf", "cbo", "funcao", "admissao"):
                if not f[campo] and rec.get(campo):
                    f[campo] = rec[campo]

            m = f["meses"].setdefault(comp or "?", {
                "competencia": comp, "proventos": 0.0, "descontos": 0.0,
                "liquido": 0.0, "tipos": [], "verbas": [], "salario_base": None,
                "integridade": [],
            })
            i = rec["integridade"]
            m["proventos"] += i["soma_proventos"]
            m["descontos"] += i["soma_descontos"]
            m["liquido"] += i["liquido_calculado"]
            m["tipos"].append(rec["tipo"])
            m["verbas"].extend(rec["verbas"])
            m["integridade"].append(nivel)
            sb = rec["totais"].get("salario_base")
            if sb and rec["tipo"] == "mensal":
                m["salario_base"] = sb

    for f in funcionarios.values():
        f["nomes_vistos"] = sorted(f["nomes_vistos"])
        f["nome"] = melhor_grafia(f["nomes_vistos"])
        if len(f["nomes_vistos"]) > 1:
            problemas.append({
                "funcionario": f["nome"],
                "mensagem": "Mesmo funcionario aparece com grafias diferentes: "
                            + " | ".join(f["nomes_vistos"]),
                "gravidade": "baixa",
            })

    return {
        "empresa": empresa,
        "competencias": competencia_range(comps),
        "funcionarios": funcionarios,
        "qualidade": qualidade,
        "problemas": problemas,
        "resumos_empresa": resumos,
    }


def conferir_com_resumo(consolidado: dict) -> dict:
    """
    Confere a leitura contra o total impresso pelo proprio relatorio.

    E a prova de que nenhum funcionario ficou de fora: a soma do que foi
    extraido tem que bater com o resumo geral da folha, competencia a
    competencia. Sem isso, uma pagina mal lida passaria despercebida — os
    recibos lidos fechariam certo individualmente e a folha viria incompleta.
    """
    por_comp = {}
    for f in consolidado["funcionarios"].values():
        for comp, m in f["meses"].items():
            a = por_comp.setdefault(comp, {"proventos": 0.0, "descontos": 0.0, "funcionarios": 0})
            a["proventos"] += m["proventos"]
            a["descontos"] += m["descontos"]
            a["funcionarios"] += 1

    conferencias = []
    for res in consolidado.get("resumos_empresa", []):
        comp = res.get("competencia")
        extraido = por_comp.get(comp)
        if not extraido or not res.get("proventos"):
            continue
        dp = round(extraido["proventos"] - res["proventos"], 2)
        dq = (extraido["funcionarios"] - res["funcionarios"]) if res.get("funcionarios") else None
        conferencias.append({
            "competencia": comp,
            "proventos_extraido": round(extraido["proventos"], 2),
            "proventos_relatorio": res["proventos"],
            "diferenca": dp,
            "funcionarios_extraido": extraido["funcionarios"],
            "funcionarios_relatorio": res.get("funcionarios"),
            "diferenca_funcionarios": dq,
            "ok": abs(dp) <= TOLERANCIA_CENTAVOS and (dq in (0, None)),
        })

    consolidado["conferencia_resumo"] = conferencias
    return consolidado


# ─────────────────────────────────────────────
# CLASSIFICACAO DE RUBRICAS
# ─────────────────────────────────────────────

def _classifica_serie(valores_por_comp: dict, meses_ativos: list) -> dict:
    """
    Classifica uma rubrica a partir da sua serie historica.

    classe:
      fixa             -> mesmo valor em todos os meses ativos (evento fixo)
      fixa_reajustada  -> valor constante que mudou de patamar (dissidio/promocao)
      variavel         -> presente todo mes com valor oscilando (lancar mes a mes)
      eventual         -> aparece em parte dos meses
    """
    presentes = [c for c in meses_ativos if c in valores_por_comp]
    vals = [valores_por_comp[c] for c in presentes]
    n_ativos = len(meses_ativos) or 1
    cobertura = len(presentes) / n_ativos

    distintos = sorted({round(v, 2) for v in vals})

    # Rubrica que comecou no meio do periodo e seguiu ate o fim (ou que existia
    # e parou) e RECORRENTE, nao eventual — precisa ser cadastrada, nao "lancada
    # quando ocorrer". So a cobertura nao distingue esses dois casos.
    idx = [meses_ativos.index(c) for c in presentes] if presentes else []
    contiguo = bool(idx) and idx == list(range(idx[0], idx[-1] + 1))
    ponta = contiguo and len(idx) >= 2 and (idx[-1] == n_ativos - 1 or idx[0] == 0)
    recorrente = cobertura >= LIMIAR_RECORRENTE or ponta

    if not vals:
        classe = "eventual"
    elif recorrente:
        if len(distintos) == 1:
            classe = "fixa"
        elif len(distintos) <= 3 and _muda_de_patamar(presentes, valores_por_comp):
            classe = "fixa_reajustada"
        else:
            classe = "variavel"
    else:
        classe = "eventual"

    return {
        "classe": classe,
        "meses_presente": len(presentes),
        "meses_ativos": n_ativos,
        "cobertura": round(cobertura, 3),
        "valores": {c: round(valores_por_comp[c], 2) for c in presentes},
        "min": round(min(vals), 2) if vals else 0.0,
        "max": round(max(vals), 2) if vals else 0.0,
        "media": round(statistics.fmean(vals), 2) if vals else 0.0,
        "ultimo": round(vals[-1], 2) if vals else 0.0,
        "distintos": distintos,
    }


def _muda_de_patamar(presentes: list, valores: dict) -> bool:
    """
    True quando a serie e constante por trechos (ex: 1.800 nos 3 primeiros
    meses e 1.980 nos 3 ultimos) — reajuste, nao variacao mensal.
    """
    seq = [round(valores[c], 2) for c in presentes]
    trocas = sum(1 for a, b in zip(seq, seq[1:]) if abs(a - b) > TOLERANCIA_CENTAVOS)
    return trocas <= 2 and len(seq) >= 3


def analisar_rubricas(consolidado: dict) -> dict:
    """
    Monta a serie historica de cada rubrica por funcionario e o catalogo
    consolidado da empresa.
    """
    comps = consolidado["competencias"]
    catalogo = {}

    for f in consolidado["funcionarios"].values():
        meses_ativos = [c for c in comps if c in f["meses"]]
        f["meses_ativos"] = meses_ativos

        series = defaultdict(dict)   # chave_rubrica -> {comp: valor}
        meta = {}
        for comp in meses_ativos:
            por_rubrica = defaultdict(float)
            for v in f["meses"][comp]["verbas"]:
                kr = chave_rubrica(v)
                por_rubrica[kr] += v["valor"]
                if kr not in meta:
                    meta[kr] = {
                        "descricao": v["descricao"],
                        "codigo_origem": v.get("codigo", ""),
                        "tipo": v["tipo"],
                        "grupo": v.get("grupo", ""),
                    }
                elif v.get("codigo") and not meta[kr]["codigo_origem"]:
                    meta[kr]["codigo_origem"] = v["codigo"]
            for kr, val in por_rubrica.items():
                series[kr][comp] = val

        f["rubricas"] = {}
        for kr, vals in series.items():
            info = _classifica_serie(vals, meses_ativos)
            info.update(meta.get(kr, {}))
            info["chave"] = kr
            alvo = colapsa_siglas(flat(info.get("descricao", kr)))
            if RX_CALCULADA.match(alvo) or RX_CALCULADA.match(kr):
                info["classe"] = "calculada"
            f["rubricas"][kr] = info

            c = catalogo.setdefault(kr, {
                "chave": kr, "descricao": info.get("descricao", kr),
                "descricoes_vistas": set(), "codigos_origem": set(),
                "tipo": info.get("tipo", "provento"), "grupo": info.get("grupo", ""),
                "funcionarios": 0, "ocorrencias": 0, "total": 0.0,
                "classes": defaultdict(int), "min": None, "max": None,
            })
            c["funcionarios"] += 1
            c["ocorrencias"] += info["meses_presente"]
            c["total"] += sum(info["valores"].values())
            c["classes"][info["classe"]] += 1
            c["descricoes_vistas"].add(info.get("descricao", ""))
            if info.get("codigo_origem"):
                c["codigos_origem"].add(info["codigo_origem"])
            c["min"] = info["min"] if c["min"] is None else min(c["min"], info["min"])
            c["max"] = info["max"] if c["max"] is None else max(c["max"], info["max"])

    # Classe predominante da rubrica na empresa
    for c in catalogo.values():
        c["mapeado"] = bool(c.get("grupo")) and c["grupo"] in rubrics.RUBRIC_META
        c["descricoes_vistas"] = sorted(x for x in c["descricoes_vistas"] if x)
        if c["descricoes_vistas"]:
            c["descricao"] = melhor_grafia(c["descricoes_vistas"])
        c["codigos_origem"] = sorted(c["codigos_origem"])
        c["classe"] = max(c["classes"].items(), key=lambda kv: kv[1])[0] if c["classes"] else "eventual"
        c["classes"] = dict(c["classes"])
        c["total"] = round(c["total"], 2)

    consolidado["catalogo_rubricas"] = sorted(
        catalogo.values(), key=lambda x: (-x["ocorrencias"], -x["total"])
    )
    return consolidado


# ─────────────────────────────────────────────
# VARIACOES DO PERIODO
# ─────────────────────────────────────────────

def detectar_variacoes(consolidado: dict) -> dict:
    """
    O que mudou no periodo — por funcionario e consolidado para a empresa.
    E esta a lista que evita implantar a folha com a foto de um mes so.
    """
    comps = consolidado["competencias"]
    if not comps:
        consolidado["eventos"] = []
        return consolidado

    primeiro, ultimo = comps[0], comps[-1]
    eventos = []

    for f in consolidado["funcionarios"].values():
        ativos = f.get("meses_ativos", [])
        if not ativos:
            continue
        f_eventos = []

        # Entrada no periodo
        if ativos[0] != primeiro:
            f_eventos.append({
                "tipo": "entrada", "competencia": ativos[0],
                "titulo": "Entrou na folha durante o periodo",
                "detalhe": f"Primeiro recibo em {competencia_label(ativos[0])}"
                           + (f" — admissao {f['admissao']}" if f.get("admissao") else ""),
                "gravidade": "media",
            })

        # Saida do periodo
        rescisao = any("rescisao" in f["meses"][c]["tipos"] for c in ativos)
        if ativos[-1] != ultimo or rescisao:
            f_eventos.append({
                "tipo": "saida", "competencia": ativos[-1],
                "titulo": "Rescisao no periodo" if rescisao else "Saiu da folha durante o periodo",
                "detalhe": f"Ultimo recibo em {competencia_label(ativos[-1])}",
                "gravidade": "alta" if rescisao else "media",
            })

        # Meses faltando no meio (buraco na documentacao)
        faltando = [c for c in comps
                    if ativos[0] < c < ativos[-1] and c not in f["meses"]]
        if faltando:
            f_eventos.append({
                "tipo": "lacuna", "competencia": faltando[0],
                "titulo": "Meses sem recibo no meio do periodo",
                "detalhe": "Faltam: " + ", ".join(competencia_label(c) for c in faltando),
                "gravidade": "alta",
            })

        # Reajuste salarial
        sal = [(c, f["meses"][c]["salario_base"]) for c in ativos
               if f["meses"][c].get("salario_base")]
        for (c1, v1), (c2, v2) in zip(sal, sal[1:]):
            if abs(v2 - v1) > TOLERANCIA_CENTAVOS:
                pct = ((v2 / v1) - 1) * 100 if v1 else 0
                f_eventos.append({
                    "tipo": "reajuste", "competencia": c2,
                    "titulo": "Mudanca de salario base",
                    "detalhe": f"R$ {v1:,.2f} -> R$ {v2:,.2f} ({pct:+.1f}%) em {competencia_label(c2)}"
                               .replace(",", "X").replace(".", ",").replace("X", "."),
                    "gravidade": "media",
                })

        # Rubricas que entraram / cessaram / deram salto
        for kr, info in f.get("rubricas", {}).items():
            presentes = sorted(info["valores"].keys())
            if not presentes:
                continue
            desc = info.get("descricao", kr)

            if info["classe"] != "eventual":
                if presentes[0] != ativos[0]:
                    f_eventos.append({
                        "tipo": "rubrica_nova", "competencia": presentes[0],
                        "titulo": f"Rubrica passou a ser paga: {desc}",
                        "detalhe": f"Comeca em {competencia_label(presentes[0])} e segue ate "
                                   f"{competencia_label(presentes[-1])}",
                        "gravidade": "media",
                    })
                if presentes[-1] != ativos[-1]:
                    f_eventos.append({
                        "tipo": "rubrica_cessou", "competencia": presentes[-1],
                        "titulo": f"Rubrica deixou de ser paga: {desc}",
                        "detalhe": f"Ultimo pagamento em {competencia_label(presentes[-1])}",
                        "gravidade": "media",
                    })

            if info["classe"] == "variavel" and info["media"] > 0:
                for c, v in info["valores"].items():
                    desvio = (v - info["media"]) / info["media"]
                    if abs(desvio) >= LIMIAR_SALTO:
                        f_eventos.append({
                            "tipo": "salto", "competencia": c,
                            "titulo": f"Valor fora do padrao: {desc}",
                            "detalhe": f"R$ {v:,.2f} em {competencia_label(c)} contra media de "
                                       f"R$ {info['media']:,.2f} ({desvio*100:+.0f}%)"
                                       .replace(",", "X").replace(".", ",").replace("X", "."),
                            "gravidade": "baixa",
                        })

        f_eventos.sort(key=lambda e: (e["competencia"] or "", e["tipo"]))
        f["eventos"] = f_eventos
        for e in f_eventos:
            eventos.append(dict(e, funcionario=f["nome"], chave=f["chave"],
                                competencia_label=competencia_label(e["competencia"])))

    ordem = {"alta": 0, "media": 1, "baixa": 2}
    eventos.sort(key=lambda e: (ordem.get(e["gravidade"], 3), e["competencia"] or ""))
    consolidado["eventos"] = eventos
    return consolidado


# ─────────────────────────────────────────────
# RESUMO E ALERTAS
# ─────────────────────────────────────────────

def montar_resumo(consolidado: dict) -> dict:
    comps = consolidado["competencias"]
    funcs = list(consolidado["funcionarios"].values())

    por_comp = []
    for c in comps:
        ativos = [f for f in funcs if c in f["meses"]]
        prov = sum(f["meses"][c]["proventos"] for f in ativos)
        desc = sum(f["meses"][c]["descontos"] for f in ativos)
        por_comp.append({
            "competencia": c, "label": competencia_label(c),
            "funcionarios": len(ativos),
            "proventos": round(prov, 2), "descontos": round(desc, 2),
            "liquido": round(prov - desc, 2),
        })

    cat = consolidado.get("catalogo_rubricas", [])
    consolidado["resumo"] = {
        "empresa": consolidado["empresa"],
        "competencias": comps,
        "periodo": (f"{competencia_label(comps[0])} a {competencia_label(comps[-1])}"
                    if comps else "-"),
        "meses": len(comps),
        "funcionarios": len(funcs),
        "rubricas_distintas": len(cat),
        "rubricas_fixas": sum(1 for c in cat if c["classe"] in ("fixa", "fixa_reajustada")),
        "rubricas_variaveis": sum(1 for c in cat if c["classe"] == "variavel"),
        "rubricas_eventuais": sum(1 for c in cat if c["classe"] == "eventual"),
        "rubricas_calculadas": sum(1 for c in cat if c["classe"] == "calculada"),
        "por_competencia": por_comp,
        "massa_media": round(statistics.fmean([p["proventos"] for p in por_comp]), 2) if por_comp else 0.0,
    }
    return consolidado


def montar_alertas(consolidado: dict) -> dict:
    """Pendencias que precisam de decisao humana ANTES da importacao."""
    alertas = list(consolidado.get("problemas", []))
    q = consolidado["qualidade"]

    if q["recibos_erro"]:
        alertas.insert(0, {
            "gravidade": "alta",
            "mensagem": f"{q['recibos_erro']} recibo(s) com leitura que nao fecha com os "
                        f"totais impressos. Nao importe esses meses sem conferir.",
        })
    for cf in consolidado.get("conferencia_resumo", []):
        if cf["ok"]:
            alertas.append({
                "gravidade": "baixa",
                "mensagem": f"{competencia_label(cf['competencia'])}: leitura conferida contra o "
                            f"resumo do relatorio — {cf['funcionarios_extraido']} funcionarios e "
                            f"R$ {cf['proventos_extraido']:,.2f} em proventos, exatamente como impresso."
                            .replace(",", "X").replace(".", ",").replace("X", "."),
            })
        else:
            alertas.insert(0, {
                "gravidade": "alta",
                "mensagem": f"{competencia_label(cf['competencia'])}: a leitura NAO bate com o resumo "
                            f"do relatorio. Extraido {cf['funcionarios_extraido']} funcionarios / "
                            f"R$ {cf['proventos_extraido']:,.2f}; o relatorio diz "
                            f"{cf['funcionarios_relatorio']} / R$ {cf['proventos_relatorio']:,.2f} "
                            f"(diferenca R$ {cf['diferenca']:,.2f}). Ha folha faltando."
                            .replace(",", "X").replace(".", ",").replace("X", "."),
            })

    if q["recibos_ocr"]:
        alertas.append({
            "gravidade": "media",
            "mensagem": f"{q['recibos_ocr']} recibo(s) vieram de PDF escaneado (OCR). "
                        f"Confira os valores manualmente por amostragem.",
        })

    # Rubricas sem grupo canonico = sem equivalente mapeado no destino
    sem_grupo = [c for c in consolidado.get("catalogo_rubricas", [])
                 if c["classe"] != "calculada"
                 and (not c.get("grupo") or c["grupo"] not in rubrics.RUBRIC_META)]
    if sem_grupo:
        alertas.append({
            "gravidade": "alta",
            "mensagem": f"{len(sem_grupo)} rubrica(s) sem equivalencia cadastrada: "
                        + ", ".join(c["descricao"] for c in sem_grupo[:8])
                        + (" ..." if len(sem_grupo) > 8 else "")
                        + ". Defina o codigo de destino no mapa de rubricas.",
        })

    # Funcionario sem CPF — chave fraca para importacao
    sem_cpf = [f for f in consolidado["funcionarios"].values() if not f.get("cpf")]
    if sem_cpf:
        alertas.append({
            "gravidade": "media",
            "mensagem": f"{len(sem_cpf)} funcionario(s) sem CPF identificado no documento: "
                        + ", ".join(f["nome"] for f in sem_cpf[:8])
                        + (" ..." if len(sem_cpf) > 8 else ""),
        })

    ordem = {"alta": 0, "media": 1, "baixa": 2}
    alertas.sort(key=lambda a: ordem.get(a.get("gravidade"), 3))
    consolidado["alertas"] = alertas
    return consolidado


# ─────────────────────────────────────────────
# PIPELINE
# ─────────────────────────────────────────────

def analisar(documentos: list) -> dict:
    """Pipeline completo: documentos extraidos -> analise pronta para a tela."""
    c = consolidar(documentos)
    c = conferir_com_resumo(c)
    c = analisar_rubricas(c)
    c = detectar_variacoes(c)
    c = montar_resumo(c)
    c = montar_alertas(c)

    # Serializa funcionarios como lista ordenada por nome
    c["funcionarios"] = sorted(
        c["funcionarios"].values(),
        key=lambda f: (f["nome"] or "zzz").upper(),
    )
    return c
