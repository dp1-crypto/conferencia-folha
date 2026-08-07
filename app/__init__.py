from flask import Flask


def create_app():
    app = Flask(__name__, template_folder='templates', static_folder='static')
    app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

    from app.routes.main import main_bp
    from app.routes.payroll import payroll_bp
    from app.routes.benefits import benefits_bp
    from app.routes.audit import audit_bp
    from app.routes.export_routes import export_bp
    from app.routes.implantacao_routes import implantacao_bp

    app.register_blueprint(main_bp)
    app.register_blueprint(payroll_bp)
    app.register_blueprint(benefits_bp)
    app.register_blueprint(audit_bp)
    app.register_blueprint(export_bp)
    app.register_blueprint(implantacao_bp)

    return app


# ── Re-exports para compatibilidade com tests.py e app.py shim ──
from app.core.utils import norm, brl, brl_dec, fmt_brl
from app.core.config import TOLERANCIA_CENTAVOS, TOLERANCIA_DIVERGENCIA, STATUS

from app.services.rubrics import (
    RUBRIC_GROUPS, RUBRIC_META, normalize_rubric,
    rubric_words_overlap, find_rubric_by_value, reload_rubric_config,
)

from app.services.payroll_parser import parse_excel, parse_pdf, parse_word
from app.services.payroll_comparator import match_names, compare
from app.services.benefits_parser import (
    fix_spaced, parse_plano_fatura,
    parse_extrato_plano, parse_referencia_simples,
)
from app.services.benefits_comparator import match_names_beneficio, compare_plano_saude
