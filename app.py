#!/usr/bin/env python3
"""
app.py -- shim de compatibilidade.
Reexporta simbolos do package app/ para manter testes legados funcionando.
"""
# Core
from app.core.utils import norm, brl, brl_dec, fmt_brl
from app.core.config import TOLERANCIA_CENTAVOS, TOLERANCIA_DIVERGENCIA, STATUS

# Rubricas
from app.services.rubrics import (
    RUBRIC_GROUPS, RUBRIC_META, normalize_rubric,
    rubric_words_overlap, find_rubric_by_value, reload_rubric_config,
)

# Parsers
from app.services.payroll_parser import parse_excel, parse_pdf, parse_word
from app.services.payroll_comparator import match_names, compare
from app.services.benefits_parser import (
    fix_spaced, parse_plano_fatura,
    parse_extrato_plano, parse_referencia_simples,
)
from app.services.benefits_comparator import match_names_beneficio, compare_plano_saude

# App Flask
from app import create_app
app = create_app()

if __name__ == '__main__':
    app.run(debug=True, port=5096)
