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

    # Alteracoes nos que estao em ambas
    for nome_atual, nome_ant in map_atual_para_anterior.items():
        atual = folha_atual[nome_atual]
        anterior = folha_anterior[nome_ant]
        nome_exibir = atual.get('nome_original', nome_atual.title())
        tem_alteracao = False

        for campo, label in [
            ('liquido', 'Liquido a Receber'),
            ('total_vencimentos', 'Total Vencimentos'),
            ('total_descontos', 'Total Descontos'),
        ]:
            v_ant = anterior.get(campo, 0) or 0
            v_atu = atual.get(campo, 0) or 0
            diff = round(v_atu - v_ant, 2)

            if abs(diff) <= TOLERANCIA_DIVERGENCIA:
                continue

            pct = round((diff / v_ant * 100), 1) if v_ant else 0

            if abs(pct) > 20 or abs(diff) > 500:
                criticidade = 'alta'
                badge_text = 'CRITICO'
            elif abs(pct) > 5 or abs(diff) > 100:
                criticidade = 'media'
                badge_text = 'ATENCAO'
            else:
                criticidade = 'baixa'
                badge_text = 'BAIXO'

            alteracoes.append({
                'nome': nome_exibir,
                'campo': label,
                'valor_anterior': v_ant,
                'valor_atual': v_atu,
                'diferenca': diff,
                'pct_variacao': pct,
                'criticidade': criticidade,
                'badge_text': badge_text,
            })
            tem_alteracao = True

        # Compara rubricas
        verbas_atu = {v['descricao'].upper(): v['valor'] for v in atual.get('verbas', [])}
        verbas_ant = {v['descricao'].upper(): v['valor'] for v in anterior.get('verbas', [])}

        for desc, val in verbas_atu.items():
            if desc not in verbas_ant:
                rubricas_novas.append({'nome': nome_exibir, 'rubrica': desc.title(), 'valor': val})
                tem_alteracao = True

        for desc, val in verbas_ant.items():
            if desc not in verbas_atu:
                rubricas_removidas.append({'nome': nome_exibir, 'rubrica': desc.title(), 'valor': val})
                tem_alteracao = True

        if not tem_alteracao:
            sem_alteracao += 1

    total_criticos = sum(1 for a in alteracoes if a['criticidade'] == 'alta')
    total_criticos += len(novos) + len(desligados)

    return {
        'colaboradores_novos': sorted(novos, key=lambda x: x['nome']),
        'colaboradores_desligados': sorted(desligados, key=lambda x: x['nome']),
        'alteracoes': sorted(alteracoes, key=lambda x: (x['criticidade'] != 'alta', x['criticidade'] != 'media', x['nome'])),
        'rubricas_novas': rubricas_novas,
        'rubricas_removidas': rubricas_removidas,
        'resumo': {
            'total_atual': len(folha_atual),
            'total_anterior': len(folha_anterior),
            'novos': len(novos),
            'desligados': len(desligados),
            'alterados': len(set(a['nome'] for a in alteracoes)),
            'sem_alteracao': sem_alteracao,
            'total_criticos': total_criticos,
        }
    }
