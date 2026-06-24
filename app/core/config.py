"""
Configuracoes globais — Conferencia de Folha
Sigma Contabilidade
"""

TOLERANCIA_CENTAVOS = 0.02      # diferenca ate R$ 0,02 = arredondamento
TOLERANCIA_DIVERGENCIA = 0.05   # acima disso = divergencia real
MAX_CONTENT_LENGTH = 100 * 1024 * 1024  # 100 MB

STATUS = {
    "OK":                    "OK",
    "A_PAGAR":               "DIFERENÇA A PAGAR",
    "PAGO_MAIOR":            "PAGO A MAIOR",
    "DESCONTO_INDEVIDO":     "DESCONTO INDEVIDO",
    "NAO_LOCALIZADO_RECIBO": "NÃO LOCALIZADO NO RECIBO",
    "NAO_LOCALIZADO_REL":    "NÃO LOCALIZADO NO RELATÓRIO",
    "POSSIVEL_RUBRICA":      "POSSÍVEL RUBRICA EQUIVALENTE",
    "POSSIVEL_COLABORADOR":  "POSSÍVEL COLABORADOR EQUIVALENTE",
    "REVISAO_MANUAL":        "NECESSITA REVISÃO MANUAL",
    "ARREDONDAMENTO":        "DIFERENÇA DE ARREDONDAMENTO",
}
