#!/usr/bin/env python3
"""
Importa o catalogo de rubricas do Dominio e monta o de-para
Sigma Contabilidade

Le um Extrato Mensal (ou Recibo de Pagamento) do Dominio, extrai codigo +
descricao de cada rubrica, resolve o grupo canonico e preenche o campo
'codigo_destino' de cada grupo no rubricas-equivalentes.json.

Uso:
    python3 importar_dominio.py amostras/dominio/*.pdf            # so relatorio
    python3 importar_dominio.py --gravar amostras/dominio/*.pdf   # grava o de-para

Por que existe: metade do de-para (rubrica de origem -> grupo Sigma) sai do
dicionario. A outra metade (grupo Sigma -> codigo no Dominio) depende de saber
os codigos DESTA instalacao do Dominio — que ninguem pode adivinhar, mas que
estao impressos em qualquer folha ja processada pela Sigma.
"""
import json
import os
import re
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.core.textutils import is_money                        # noqa: E402
from app.services.pdf_grid import read_pdf, row_text          # noqa: E402
from app.services.holerite_parser import (                     # noqa: E402
    resolver_grupo, colapsa_siglas, _palavras,
)
from app.core.textutils import flat                            # noqa: E402
from app.services import rubrics                              # noqa: E402

JSON_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "rubricas-equivalentes.json")

# Extrato Mensal do Dominio: "8781DIAS NORMAIS 30,00 1.754,50P"
# O codigo vem colado na descricao e o marcador P/D no fim diz o lado.
RX_EXTRATO = re.compile(
    r"(\d{2,5})\s?([A-ZÀ-Ü][A-Za-zÀ-ÿ0-9 .%/º°()\-]*?)\s+([\d.,]+)\s+([\d.,]+)([PD])\b"
)
# Recibo de Pagamento: "8781 DIAS NORMAIS 30,00 1.754,50" (sem marcador de lado)
RX_RECIBO = re.compile(
    r"^(\d{2,5})\s+([A-ZÀ-Ü][A-Za-zÀ-ÿ0-9 .%/º°()\-]{3,40}?)\s+([\d.,]+)\s+([\d.,]+)\s*$"
)

# Rubricas do Dominio que nao sao verba (totalizador do proprio relatorio)
RX_IGNORAR = re.compile(r"^(LIQUIDO|TOTAL|BASE|PROVENTOS|DESCONTOS)")


def extrair(caminhos: list) -> dict:
    """{codigo: {descricoes, lados, ocorrencias}} lido dos PDFs do Dominio."""
    cat = defaultdict(lambda: {"desc": set(), "lado": set(), "n": 0})
    for caminho in caminhos:
        with open(caminho, "rb") as f:
            doc = read_pdf(f.read())
        for pag in doc["paginas"]:
            for row in pag["rows"]:
                texto = row_text(row)
                achou = False
                for m in RX_EXTRATO.finditer(texto):
                    achou = True
                    _guarda(cat, m.group(1), m.group(2), m.group(5))
                if not achou:
                    m = RX_RECIBO.match(texto.strip())
                    # O valor tem que ser monetario de verdade. Sem isso a linha
                    # de cadastro do funcionario ("425 ADAIANE MARIA SOARES
                    # 521140 1 1") entra no catalogo como se fosse rubrica.
                    if m and is_money(m.group(4)):
                        _guarda(cat, m.group(1), m.group(2), "")
    return cat


def _guarda(cat, codigo, descricao, lado):
    desc = " ".join(descricao.split()).strip(" .:-")
    if not desc or RX_IGNORAR.match(desc.upper()):
        return
    cat[codigo]["desc"].add(desc)
    if lado:
        cat[codigo]["lado"].add(lado)
    cat[codigo]["n"] += 1


def propor(cat: dict) -> dict:
    """
    grupo canonico -> lista de candidatos {codigo, descricao, ocorrencias}

    Mais de um codigo do Dominio pode cair no mesmo grupo (ex: DIAS FALTAS,
    DIAS FALTAS DSR e HORAS FALTAS PARCIAL sao todos FALTAS). Nesse caso a
    escolha e humana — o script mostra os candidatos em vez de decidir.
    """
    por_grupo = defaultdict(list)
    sem_grupo = []
    for codigo, d in cat.items():
        desc = sorted(d["desc"], key=len)[-1]
        grupo = resolver_grupo(desc)
        item = {"codigo": codigo, "descricao": desc, "n": d["n"],
                "lado": "/".join(sorted(d["lado"]))}
        if grupo in rubrics.RUBRIC_META:
            por_grupo[grupo].append(item)
        else:
            sem_grupo.append(item)
    for g in por_grupo:
        por_grupo[g].sort(key=lambda x: -x["n"])
    return por_grupo, sem_grupo


def _casa_exato(descricao: str, grupo: str) -> bool:
    """
    True quando a descricao do Dominio e a rubrica GENERICA do grupo, nao uma
    variante especializada dele.

    Um candidato unico nao significa candidato certo: nesta folha o unico
    codigo de hora extra e 'HORAS EXTRAS 60%', e gravar 60% como destino de
    toda hora extra jogaria a HE de 100% no codigo errado. Mesma coisa com
    'DIFERENCA DE 1/3 DE FERIAS', que e a diferenca e nao o terco.

    A comparacao e contra o NOME DO GRUPO, nao contra suas variantes: as
    variantes existem para reconhecer a rubrica na leitura, e varias delas sao
    justamente as formas especializadas. Compara-las aqui seria circular —
    'ADIANTAMENTO DE FERIAS' e variante de ADIANTAMENTO e passaria como se
    fosse o adiantamento generico.
    """
    alvo = _palavras(colapsa_siglas(flat(descricao)))
    canonico = _palavras(grupo.replace("_", " "))
    return bool(alvo) and alvo == canonico


def gravar(por_grupo: dict) -> tuple:
    """
    Grava codigo_destino so quando ha UM candidato E ele e a rubrica generica
    do grupo. Nos outros casos a escolha e da Sigma — o script mostra os
    candidatos em vez de decidir. Nunca sobrescreve valor preenchido a mao.
    """
    with open(JSON_PATH, encoding="utf-8") as f:
        dados = json.load(f)

    preenchidos, ambiguos, mantidos = [], [], []
    for grupo, candidatos in por_grupo.items():
        alvo = dados["grupos"].get(grupo)
        if alvo is None:
            continue
        if alvo.get("codigo_destino"):
            mantidos.append((grupo, alvo["codigo_destino"]))
            continue
        exatos = [c for c in candidatos if _casa_exato(c["descricao"], grupo)]
        if len(exatos) == 1:
            alvo["codigo_destino"] = exatos[0]["codigo"]
            alvo["obs_destino"] = f"Dominio: {exatos[0]['descricao']}"
            preenchidos.append((grupo, exatos[0]))
        else:
            ambiguos.append((grupo, candidatos))

    tmp = JSON_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    os.replace(tmp, JSON_PATH)
    return preenchidos, ambiguos, mantidos


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    escrever = "--gravar" in sys.argv
    if not args:
        print(__doc__)
        return 1

    cat = extrair(args)
    por_grupo, sem_grupo = propor(cat)
    print(f"{len(cat)} rubricas do Dominio lidas de {len(args)} arquivo(s)\n")

    if escrever:
        preenchidos, ambiguos, mantidos = gravar(por_grupo)
        print(f"PREENCHIDOS ({len(preenchidos)}):")
        for g, c in sorted(preenchidos):
            print(f"   {g:<26} -> {c['codigo']:<6} {c['descricao'][:44]}")
        if mantidos:
            print(f"\nJA TINHAM CODIGO, mantidos ({len(mantidos)}):")
            for g, c in sorted(mantidos):
                print(f"   {g:<26} -> {c}")
        if ambiguos:
            print(f"\nPRECISAM DE DECISAO — mais de um codigo no mesmo grupo ({len(ambiguos)}):")
            for g, cands in sorted(ambiguos):
                print(f"   {g}")
                for c in cands:
                    print(f"      {c['codigo']:<6} n={c['n']:<3} {c['descricao'][:48]}")
    else:
        for g, cands in sorted(por_grupo.items()):
            marca = " " if len(cands) == 1 else "!"
            print(f"{marca} {g}")
            for c in cands:
                print(f"      {c['codigo']:<6} {c['lado']:<3} n={c['n']:<3} {c['descricao'][:48]}")

    if sem_grupo:
        print(f"\nSEM GRUPO CANONICO ({len(sem_grupo)}) — cadastre em rubricas-equivalentes.json:")
        for c in sorted(sem_grupo, key=lambda x: -x["n"]):
            print(f"   {c['codigo']:<6} n={c['n']:<3} {c['descricao'][:52]}")

    faltando = [g for g in rubrics.RUBRIC_META if g not in por_grupo]
    if faltando:
        print(f"\nGRUPOS SEM CODIGO DO DOMINIO ({len(faltando)}) — nao apareceram nesta folha:")
        print("   " + ", ".join(sorted(faltando)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
