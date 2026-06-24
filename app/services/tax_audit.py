"""
Auditoria de INSS e IRRF
"""
from app.core.utils import norm, fmt_brl
from app.core.config import TOLERANCIA_DIVERGENCIA

# Tabela INSS progressiva 2024/2025
TABELA_INSS = [
    (1412.00,  0.075),
    (2666.68,  0.09),
    (4000.03,  0.12),
    (7786.02,  0.14),
]
TETO_INSS = 7786.02
INSS_MAXIMO = 908.86

# Tabela IRRF 2024 (base, aliquota, parcela_deducao)
TABELA_IRRF = [
    (2259.20,  0.000,  0.00),
    (2826.65,  0.075, 169.44),
    (3751.05,  0.150, 381.44),
    (4664.68,  0.225, 662.77),
    (float('inf'), 0.275, 896.00),
]
DEDUCAO_DEPENDENTE = 189.59


def calcular_inss(salario_bruto: float) -> float:
    """Calcula INSS pela tabela progressiva 2024/2025."""
    base = min(salario_bruto, TETO_INSS)
    total = 0.0
    faixa_anterior = 0.0
    for teto, aliquota in TABELA_INSS:
        if base <= faixa_anterior:
            break
        valor_na_faixa = min(base, teto) - faixa_anterior
        total += valor_na_faixa * aliquota
        faixa_anterior = teto
        if base <= teto:
            break
    return round(min(total, INSS_MAXIMO), 2)


def calcular_irrf(base_calculo: float) -> float:
    """Calcula IRRF sobre a base (ja deduzidos INSS e dependentes)."""
    if base_calculo <= 0:
        return 0.0
    for limite, aliquota, parcela in TABELA_IRRF:
        if base_calculo <= limite:
            return round(max(0, base_calculo * aliquota - parcela), 2)
    return 0.0


_KW_INSS = ['INSS', 'PREVIDENCIA', 'PREV SOCIAL', 'CONTRIB PREV', 'PREVIDENCI']
_KW_IRRF = ['IRRF', 'IRPF', 'IMP RENDA', 'IMPOSTO RENDA', 'I.R.', 'IR ']
_KW_SALARIO = ['SALARIO', 'SAL NORMAL', 'DIAS NORMAIS', 'SAL BASE', 'ORDENADO']


def _find_value_in_verbas(verbas: list, keywords: list) -> float:
    for v in verbas:
        desc = v.get('descricao', '').upper()
        if any(kw in desc for kw in keywords):
            return v.get('valor', 0)
    return 0.0


def _find_verbas_by_kw(verbas: list, keywords: list) -> list:
    """Retorna todas as verbas que batem com alguma keyword."""
    return [
        {'descricao': v.get('descricao', ''), 'valor': v.get('valor', 0), 'codigo': v.get('codigo', '')}
        for v in verbas
        if any(kw in v.get('descricao', '').upper() for kw in keywords)
    ]


def _classificar_verba(verbas: list) -> str:
    """Classifica uma verba como provento ou desconto a partir das keywords."""
    _KW_DESCONTO = ['INSS', 'IRRF', 'IRPF', 'IMP RENDA', 'FGTS', 'DESCONTO', 'VALE', 'PLANO',
                    'EMPRESTIMO', 'ADIANT', 'FALTAS', 'FALTA']
    for v in verbas:
        desc = v.get('descricao', '').upper()
        if any(kw in desc for kw in _KW_DESCONTO):
            return 'desconto'
    return 'provento'


def auditar_colaborador(nome: str, verbas: list) -> dict:
    """Audita INSS e IRRF de um colaborador baseado nas verbas do recibo."""
    salario_bruto = _find_value_in_verbas(verbas, _KW_SALARIO)
    inss_encontrado = _find_value_in_verbas(verbas, _KW_INSS)
    irrf_encontrado = _find_value_in_verbas(verbas, _KW_IRRF)

    # Verbas detalhadas por tipo (para exibição no frontend)
    verbas_inss = _find_verbas_by_kw(verbas, _KW_INSS)
    verbas_irrf = _find_verbas_by_kw(verbas, _KW_IRRF)
    verbas_salario = _find_verbas_by_kw(verbas, _KW_SALARIO)

    # Tenta encontrar base de calculo pelo total de vencimentos
    # Considera o salario bruto como os vencimentos tributaveis
    inss_calculado = calcular_inss(salario_bruto) if salario_bruto > 0 else 0
    base_irrf = max(0, salario_bruto - inss_calculado)
    irrf_calculado = calcular_irrf(base_irrf) if base_irrf > 0 else 0

    divergencias = []

    def _check_inss(calculado, encontrado):
        diff = round(abs(calculado - encontrado), 2)
        if calculado == 0 and encontrado == 0:
            return 'SEM_DADOS'
        if encontrado == 0 and calculado > 0:
            return 'AUSENTE'
        if diff <= TOLERANCIA_DIVERGENCIA:
            return 'OK'
        if diff <= 2.00:
            return 'ARREDONDAMENTO'
        return 'DIVERGENTE'

    def _check_irrf(calculado, encontrado):
        diff = round(abs(calculado - encontrado), 2)
        if calculado == 0 and encontrado == 0:
            return 'SEM_DADOS'
        if encontrado == 0 and calculado > 0:
            return 'AUSENTE'
        if diff <= 10.00:   # IRRF: tolerância de R$ 10,00 (diferenças pequenas são normais)
            return 'OK'
        return 'DIVERGENTE'

    inss_status = _check_inss(inss_calculado, inss_encontrado)
    irrf_status = _check_irrf(irrf_calculado, irrf_encontrado)

    if inss_status in ('DIVERGENTE', 'AUSENTE'):
        divergencias.append({
            'campo': 'INSS',
            'esperado': inss_calculado,
            'encontrado': inss_encontrado,
            'diferenca': round(inss_calculado - inss_encontrado, 2),
            'criticidade': 'alta' if inss_status == 'AUSENTE' else 'media',
            'status': inss_status,
        })

    if irrf_status in ('DIVERGENTE', 'AUSENTE'):
        divergencias.append({
            'campo': 'IRRF',
            'esperado': irrf_calculado,
            'encontrado': irrf_encontrado,
            'diferenca': round(irrf_calculado - irrf_encontrado, 2),
            'criticidade': 'alta' if irrf_status == 'AUSENTE' else 'media',
            'status': irrf_status,
        })

    # Todas as rubricas classificadas como provento/desconto para exibição completa
    rubricas_detalhadas = []
    for v in verbas:
        desc_upper = v.get('descricao', '').upper()
        if any(kw in desc_upper for kw in _KW_INSS):
            tipo_rb = 'inss'
        elif any(kw in desc_upper for kw in _KW_IRRF):
            tipo_rb = 'irrf'
        elif any(kw in desc_upper for kw in _KW_SALARIO):
            tipo_rb = 'salario'
        else:
            tipo_rb = 'outro'
        rubricas_detalhadas.append({
            'codigo': v.get('codigo', ''),
            'descricao': v.get('descricao', ''),
            'valor': v.get('valor', 0),
            'tipo': tipo_rb,
        })

    return {
        'nome': nome,
        'salario_bruto': salario_bruto,
        'base_inss': salario_bruto,
        'inss_calculado': inss_calculado,
        'inss_encontrado': inss_encontrado,
        'inss_status': inss_status,
        'verbas_inss': verbas_inss,
        'verbas_irrf': verbas_irrf,
        'verbas_salario': verbas_salario,
        'rubricas_detalhadas': rubricas_detalhadas,
        'base_irrf': base_irrf,
        'irrf_calculado': irrf_calculado,
        'irrf_encontrado': irrf_encontrado,
        'irrf_status': irrf_status,
        'divergencias': divergencias,
    }


def auditar_folha(pdf_data: dict) -> dict:
    """Audita toda a folha. pdf_data = resultado de parse_pdf()."""
    colaboradores = []
    total_divergencias = 0
    total_criticos = 0
    total_sem_dados = 0

    for nome_norm, emp in pdf_data.items():
        nome_exibir = emp.get('nome_original', nome_norm.title())
        verbas = emp.get('verbas', [])
        resultado = auditar_colaborador(nome_exibir, verbas)
        colaboradores.append(resultado)

        if resultado['divergencias']:
            total_divergencias += len(resultado['divergencias'])
            total_criticos += sum(1 for d in resultado['divergencias'] if d['criticidade'] == 'alta')

        if resultado['inss_status'] == 'SEM_DADOS' and resultado['irrf_status'] == 'SEM_DADOS':
            total_sem_dados += 1

    colaboradores.sort(key=lambda x: (
        len(x['divergencias']) == 0,
        -len(x['divergencias'])
    ))

    return {
        'colaboradores': colaboradores,
        'resumo': {
            'total': len(colaboradores),
            'com_divergencia': sum(1 for c in colaboradores if c['divergencias']),
            'ok': sum(1 for c in colaboradores if not c['divergencias'] and c['inss_status'] not in ('SEM_DADOS',)),
            'sem_dados': total_sem_dados,
            'total_divergencias': total_divergencias,
            'total_criticos': total_criticos,
        },
        'tabela_inss': [{'faixa': f'Ate R$ {t:,.2f}', 'aliquota': f'{a*100:.1f}%'} for t, a in TABELA_INSS],
    }
