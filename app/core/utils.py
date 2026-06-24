"""
Utilitarios base — normalizacao e conversao monetaria
"""
import re
import unicodedata
from decimal import Decimal, ROUND_HALF_UP


def norm(s: str) -> str:
    """Normaliza nome: maiusculo, sem acento, espacos colapsados."""
    s = s.upper().strip()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return " ".join(s.split())


def brl(s) -> float:
    """Converte string monetaria BR para float."""
    s = re.sub(r"[R$\s]", "", str(s))
    # Suporte a negativos entre parenteses: (1.000,00)
    negativo = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    s = s.replace(".", "").replace(",", ".")
    try:
        v = float(s)
        return -v if negativo else v
    except Exception:
        return 0.0


def brl_dec(s) -> Decimal:
    """
    Converte string monetaria BR para Decimal com precisao centesimal.
    Use para calculos financeiros criticos onde float introduz erro de arredondamento.
    Exemplos: '1.592,11' -> Decimal('1592.11'), '(300,00)' -> Decimal('-300.00')
    """
    s = re.sub(r"[R$\s]", "", str(s))
    negativo = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    s = s.replace(".", "").replace(",", ".")
    try:
        v = Decimal(s).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        return -v if negativo else v
    except Exception:
        return Decimal("0.00")


def fmt_brl(v) -> str:
    if not v:
        return "-"
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)
