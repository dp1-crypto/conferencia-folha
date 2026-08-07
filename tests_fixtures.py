"""
Gerador de holerites sinteticos para teste do motor de extracao
Sigma Contabilidade

Produz PDFs em layouts diferentes (duas colunas, coluna unica, com/sem codigo)
para validar que o extrator por coordenadas nao depende de um sistema especifico.
Nao contem dado real de nenhum cliente.
"""
import io

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

W, H = A4


def _brl(v):
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def holerite_duas_colunas(func, verbas, comp="03/2026", sistema="QUESTOR SISTEMAS"):
    """
    Layout classico: REFERENCIA | VENCIMENTOS | DESCONTOS em colunas separadas.
    verbas = [(codigo, descricao, referencia, valor, 'P'|'D')]
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    y = H - 50

    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, y, "COMERCIO EXEMPLO LTDA")
    c.setFont("Helvetica", 7)
    c.drawString(400, y, sistema)
    y -= 12
    c.drawString(40, y, "CNPJ: 11.222.333/0001-81")
    c.drawString(300, y, f"Competencia: {comp}")
    y -= 12
    c.drawString(40, y, "Recibo de Pagamento de Salario")
    y -= 20

    c.setFont("Helvetica", 8)
    c.drawString(40, y, f"Matricula: {func['matricula']}")
    c.drawString(120, y, f"Nome do Funcionario: {func['nome']}")
    y -= 11
    c.drawString(40, y, f"CPF: {func['cpf']}")
    c.drawString(180, y, f"Admissao: {func['admissao']}")
    c.drawString(300, y, f"CBO: {func['cbo']}")
    c.drawString(400, y, f"Funcao: {func['funcao']}")
    y -= 22

    # Cabecalho da tabela
    c.setFont("Helvetica-Bold", 8)
    c.drawString(40, y, "Cod.")
    c.drawString(75, y, "Descricao")
    c.drawRightString(360, y, "Referencia")
    c.drawRightString(450, y, "Vencimentos")
    c.drawRightString(540, y, "Descontos")
    y -= 4
    c.line(40, y, 540, y)
    y -= 12

    c.setFont("Helvetica", 8)
    tp = td = 0.0
    for cod, desc, ref, val, lado in verbas:
        c.drawString(40, y, cod)
        c.drawString(75, y, desc)
        if ref:
            c.drawRightString(360, y, ref)
        if lado == "P":
            c.drawRightString(450, y, _brl(val))
            tp += val
        else:
            c.drawRightString(540, y, _brl(val))
            td += val
        y -= 12

    y -= 6
    c.line(40, y, 540, y)
    y -= 14
    c.setFont("Helvetica-Bold", 8)
    c.drawString(300, y, "Total de Vencimentos")
    c.drawRightString(450, y, _brl(tp))
    c.drawString(300, y - 12, "Total de Descontos")
    c.drawRightString(540, y - 12, _brl(td))
    y -= 30
    c.drawString(300, y, "Valor Liquido")
    c.drawRightString(540, y, _brl(tp - td))
    y -= 20
    c.setFont("Helvetica", 7)
    c.drawString(40, y, f"Base INSS {_brl(tp)}   Base FGTS {_brl(tp)}   FGTS do Mes {_brl(round(tp * 0.08, 2))}   Base IRRF {_brl(tp)}")

    c.showPage()
    c.save()
    return buf.getvalue()


def holerite_coluna_unica(func, verbas, comp="03/2026", sistema="DOMINIO SISTEMAS"):
    """
    Layout de coluna unica (padrao Dominio): o lado da rubrica sai pelo nome.
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    y = H - 50

    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, y, "INDUSTRIA MODELO LTDA")
    c.setFont("Helvetica", 7)
    c.drawString(430, y, sistema)
    y -= 12
    c.drawString(40, y, "CNPJ: 11.222.333/0001-81")
    c.drawString(300, y, f"Folha Mensal - Referencia {comp}")
    y -= 24

    c.setFont("Helvetica", 8)
    c.drawString(40, y, func["matricula"])
    c.drawString(75, y, func["nome"].upper())
    c.drawString(300, y, func["cbo"])
    c.drawString(360, y, func["funcao"])
    y -= 11
    c.drawString(40, y, f"CPF {func['cpf']}   Admissao {func['admissao']}")
    y -= 22

    c.setFont("Helvetica-Bold", 8)
    c.drawString(40, y, "Codigo")
    c.drawString(90, y, "Descricao")
    c.drawRightString(400, y, "Referencia")
    c.drawRightString(500, y, "Valor")
    y -= 4
    c.line(40, y, 500, y)
    y -= 12

    c.setFont("Helvetica", 8)
    tp = td = 0.0
    for cod, desc, ref, val, lado in verbas:
        c.drawString(40, y, cod)
        c.drawString(90, y, desc)
        if ref:
            c.drawRightString(400, y, ref)
        c.drawRightString(500, y, _brl(val))
        if lado == "P":
            tp += val
        else:
            td += val
        y -= 12

    y -= 10
    c.setFont("Helvetica-Bold", 8)
    c.drawString(300, y, "Total de Vencimentos")
    c.drawRightString(500, y, _brl(tp))
    y -= 12
    c.drawString(300, y, "Total de Descontos")
    c.drawRightString(500, y, _brl(td))
    y -= 16
    c.drawString(300, y, "Valor Liquido")
    c.drawRightString(500, y, _brl(tp - td))

    c.showPage()
    c.save()
    return buf.getvalue()


# ─────────────────────────────────────────────
# DADOS DE EXEMPLO (ficticios)
# ─────────────────────────────────────────────

FUNC_A = {
    "matricula": "137", "nome": "Mariana Ribeiro Alves",
    "cpf": "529.982.247-25", "admissao": "12/03/2019",
    "cbo": "142105", "funcao": "Analista Fiscal",
}
FUNC_B = {
    "matricula": "1042", "nome": "Carlos Eduardo Monteiro",
    "cpf": "111.444.777-35", "admissao": "05/08/2022",
    "cbo": "521110", "funcao": "Vendedor",
}

VERBAS_A = [
    ("1",    "SALARIO BASE",            "30,00",  3200.00, "P"),
    ("120",  "HORAS EXTRAS 50%",        "8:30",    218.75, "P"),
    ("135",  "ADICIONAL NOTURNO",      "12,00",    96.00, "P"),
    ("902",  "GRATIFICACAO DE FUNCAO",      "",   450.00, "P"),
    ("9201", "INSS",                    "9,00",   357.31, "D"),
    ("9203", "IRRF",                    "7,50",   112.44, "D"),
    ("9301", "VALE TRANSPORTE",         "6,00",   192.00, "D"),
    ("9402", "PLANO DE SAUDE UNIMED",       "",   180.00, "D"),
]

VERBAS_B = [
    ("1",    "SALARIO",                "30,00",  1620.00, "P"),
    ("310",  "COMISSAO SOBRE VENDAS",       "",  1245.80, "P"),
    ("315",  "D.S.R. SOBRE COMISSAO",       "",   207.63, "P"),
    ("9201", "I.N.S.S.",                "9,00",   276.61, "D"),
    ("9301", "VALE TRANSPORTE",         "6,00",    97.20, "D"),
    ("9501", "ADIANTAMENTO SALARIAL",       "",   648.00, "D"),
]


def gerar_amostras() -> dict:
    """Devolve {nome_arquivo: bytes} com as amostras sinteticas."""
    return {
        "duas_colunas_marciana_03-2026.pdf": holerite_duas_colunas(FUNC_A, VERBAS_A),
        "duas_colunas_carlos_03-2026.pdf":  holerite_duas_colunas(FUNC_B, VERBAS_B),
        "coluna_unica_mariana_03-2026.pdf": holerite_coluna_unica(FUNC_A, VERBAS_A),
        "coluna_unica_carlos_03-2026.pdf":  holerite_coluna_unica(FUNC_B, VERBAS_B),
    }


if __name__ == "__main__":
    import os
    destino = os.path.join(os.path.dirname(__file__), "amostras", "sinteticas")
    os.makedirs(destino, exist_ok=True)
    for nome, raw in gerar_amostras().items():
        with open(os.path.join(destino, nome), "wb") as f:
            f.write(raw)
        print("gerado:", nome)
