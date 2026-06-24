"""
Comparacao entre folhas de meses diferentes
"""
from app.core.utils import norm, fmt_brl
from app.core.config import TOLERANCIA_DIVERGENCIA


def compare_months(folha_atual: dict, folha_anterior: dict) -> dict:
    """
    Compara folha atual com mes anterior.
    Ambas sao resultado de parse_pdf() -> {NOME_NORM: {liquido, total_vencimentos, total_descontos, verbas, ...}}

    Retorna dict com:
    - colaboradores_novos: [{nome, liquido}]
    - colaboradores_desligados: [{nome, liquido}]
    - alteracoes: [{nome, campo, valor_anterior, valor_atual, diferenca, pct_variacao, criticidade, badge_text}]
    - rubricas_novas: [{nome_colaborador, rubrica, valor}]  (verbas que apareceram na atual mas nao na anterior)
    - rubricas_removidas: [{nome_colaborador, rubrica, valor}]
    - resumo: {total_atual, total_anterior, novos, desligados, alterados, sem_alteracao, total_criticos}
    """
    novos = []
    desligados = []
    alteracoes = []
    rubricas_novas = []
    rubricas_removidas = []
    sem_alteracao = 0

    # Match de nomes (mesmo algoritmo do payroll_comparator)
    def _match(a_keys, b_keys):
        mapping = {}
        used = set()
        for ak in a_keys:
            aw = ak.split()
            best = None
            for bk in b_keys:
                if bk in used: continue
                bw = bk.split()
                if bw[:len(aw)] == aw or aw[:len(bw)] == bw:
                    best = bk; break
            if not best:
                for bk in b_keys:
                    if bk in used: continue
                    if bk.split()[:2] == ak.split()[:2]:
                        best = bk; break
            if best:
                mapping[ak] = best
                used.add(best)
        return mapping

    atual_keys = set(folha_atual.keys())
    anterior_keys = set(folha_anterior.keys())

    # Casa nomes
    map_atual_para_anterior = _match(list(atual_keys), list(anterior_keys))
    map_anterior_para_atual = {v: k for k, v in map_atual_para_anterior.items()}

    # Novos (so na atual)
    for nome in atual_keys:
        if nome not in map_atual_para_anterior:
            emp = folha_atual[nome]
            novos.append({
                'nome': emp.get('nome_original', nome.title()),
                'liquido': emp.get('liquido', 0),
                'total_vencimentos': emp.get('total_vencimentos', 0),
            })

    # Desligados (so na anterior)
    for nome in anterior_keys:
        if nome not in map_anterior_para_atual:
            emp = folha_anterior[nome]
            desligados.append({
                'nome': emp.get('nome_original', nome.title()),
                'liquido': emp.get('liquido', 0),
            })

    # Alteracoes nos que estao em ambas — agrupa por colaborador
    comparativo = []  # um registro por colaborador, todos os campos lado a lado

    for nome_atual, nome_ant in map_atual_para_anterior.items():
        atual    = folha_atual[nome_atual]
        anterior = folha_anterior[nome_ant]
        nome_exibir = atual.get('nome_original', nome_atual.title())

        def _campo(chave):
            v_ant = anterior.get(chave, 0) or 0
            v_atu = atual.get(chave, 0) or 0
            diff  = round(v_atu - v_ant, 2)
            pct   = round((diff / v_ant * 100), 1) if v_ant else 0
            return v_ant, v_atu, diff, pct

        liq_ant, liq_atu, liq_diff, liq_pct       = _campo('liquido')
        venc_ant, venc_atu, venc_diff, venc_pct    = _campo('total_vencimentos')
        desc_ant, desc_atu, desc_diff, desc_pct    = _campo('total_descontos')

        # Criticidade baseada na variação do líquido (campo mais relevante)
        tem_diff = (abs(liq_diff) > TOLERANCIA_DIVERGENCIA or
                    abs(venc_diff) > TOLERANCIA_DIVERGENCIA or
                    abs(desc_diff) > TOLERANCIA_DIVERGENCIA)

        if not tem_diff:
            criticidade = 'ok'
        elif abs(liq_pct) > 20 or abs(liq_diff) > 500:
            criticidade = 'alta'
        elif abs(liq_pct) > 5 or abs(liq_diff) > 100:
            criticidade = 'media'
        else:
            criticidade = 'baixa'

        if tem_diff:
            alteracoes.append({'nome': nome_exibir, 'criticidade': criticidade})

        # Compara rubricas
        verbas_atu = {v['descricao'].upper(): v['valor'] for v in atual.get('verbas', [])}
        verbas_ant = {v['descricao'].upper(): v['valor'] for v in anterior.get('verbas', [])}
        rb_novas     = []
        rb_removidas = []

        for desc, val in verbas_atu.items():
            if desc not in verbas_ant:
                rb_novas.append({'rubrica': desc.title(), 'valor': val})
                rubricas_novas.append({'nome': nome_exibir, 'rubrica': desc.title(), 'valor': val})

        for desc, val in verbas_ant.items():
            if desc not in verbas_atu:
                rb_removidas.append({'rubrica': desc.title(), 'valor': val})
                rubricas_removidas.append({'nome': nome_exibir, 'rubrica': desc.title(), 'valor': val})

        if not tem_diff and not rb_novas and not rb_removidas:
            sem_alteracao += 1

        comparativo.append({
            'nome':          nome_exibir,
            'criticidade':   criticidade,
            # Líquido
            'liq_ant':  liq_ant,
            'liq_atu':  liq_atu,
            'liq_diff': liq_diff,
            'liq_pct':  liq_pct,
            # Vencimentos
            'venc_ant':  venc_ant,
            'venc_atu':  venc_atu,
            'venc_diff': venc_diff,
            'venc_pct':  venc_pct,
            # Descontos
            'desc_ant':  desc_ant,
            'desc_atu':  desc_atu,
            'desc_diff': desc_diff,
            'desc_pct':  desc_pct,
            # Rubricas
            'rubricas_novas':     rb_novas,
            'rubricas_removidas': rb_removidas,
        })

    # Ordena: críticos primeiro, depois por nome
    _ordem = {'alta': 0, 'media': 1, 'baixa': 2, 'ok': 3}
    comparativo.sort(key=lambda x: (_ordem.get(x['criticidade'], 9), x['nome']))

    total_criticos = sum(1 for c in comparativo if c['criticidade'] == 'alta')
    total_criticos += len(novos) + len(desligados)
    alterados = sum(1 for c in comparativo if c['criticidade'] != 'ok')

    return {
        'colaboradores_novos':      sorted(novos,      key=lambda x: x['nome']),
        'colaboradores_desligados': sorted(desligados, key=lambda x: x['nome']),
        'comparativo':  comparativo,
        # mantém alteracoes para compatibilidade com exportar_excel
        'alteracoes': [
            {'nome': c['nome'], 'campo': 'Liquido a Receber',
             'valor_anterior': c['liq_ant'], 'valor_atual': c['liq_atu'],
             'diferenca': c['liq_diff'], 'pct_variacao': c['liq_pct'],
             'criticidade': c['criticidade'], 'badge_text': c['criticidade'].upper()}
            for c in comparativo if abs(c['liq_diff']) > TOLERANCIA_DIVERGENCIA
        ],
        'rubricas_novas':     rubricas_novas,
        'rubricas_removidas': rubricas_removidas,
        'resumo': {
            'total_atual':    len(folha_atual),
            'total_anterior': len(folha_anterior),
            'novos':          len(novos),
            'desligados':     len(desligados),
            'alterados':      alterados,
            'sem_alteracao':  sem_alteracao,
            'total_criticos': total_criticos,
        }
    }
