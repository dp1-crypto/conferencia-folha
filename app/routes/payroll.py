from flask import Blueprint, request, jsonify

from app.core.utils import fmt_brl
from app.services.payroll_parser import parse_excel, parse_pdf, parse_word
from app.services.payroll_comparator import compare
from app.services.rubrics import reload_rubric_config, RUBRIC_GROUPS

payroll_bp = Blueprint('payroll', __name__)


@payroll_bp.route("/analisar", methods=["POST"])
def analisar():
    excel_data, pdf_data, word_data = {}, {}, {}
    errors = []

    for key, f in request.files.items():
        if not f.filename:
            continue
        raw = f.read()
        fn  = f.filename.lower()
        try:
            if fn.endswith((".xls", ".xlsx")):
                excel_data.update(parse_excel(raw, f.filename))
            elif fn.endswith(".pdf"):
                pdf_data.update(parse_pdf(raw))
            elif fn.endswith((".docx", ".doc")):
                parsed = parse_word(raw)
                word_data.setdefault("gratificacoes", {}).update(parsed.get("gratificacoes", {}))
                word_data.setdefault("comissao_e_dsr", {}).update(parsed.get("comissao_e_dsr", {}))
                word_data.setdefault("descontos", {}).update(parsed.get("descontos", {}))
                word_data.setdefault("obs", []).extend(parsed.get("obs", []))
                word_data.setdefault("decimo_terceiro", []).extend(parsed.get("decimo_terceiro", []))
        except Exception as e:
            errors.append(f"{f.filename}: {e}")

    if not excel_data and not pdf_data:
        return jsonify({"error": "Nenhum arquivo processado. " + " | ".join(errors)})

    report = compare(excel_data, pdf_data, word_data)
    report["erros"] = errors

    # Coleta todas as sugestoes de equivalencia para exibir no relatorio
    sugestoes = []
    for emp in report.get("funcionarios", []):
        for s in emp.get("sugestoes_equivalencia", []):
            s["colaborador"] = emp.get("nome_exibir", emp.get("nome", ""))
            sugestoes.append(s)
    report["sugestoes_equivalencia"] = sugestoes

    return jsonify(report)


@payroll_bp.route("/recarregar-rubricas", methods=["POST"])
def recarregar_rubricas():
    """
    Recarrega o arquivo rubricas-equivalentes.json sem reiniciar o servidor.
    Chamar apos editar o arquivo manualmente.
    """
    try:
        reload_rubric_config()
        return jsonify({
            "ok": True,
            "grupos": list(RUBRIC_GROUPS.keys()),
            "total_variantes": sum(len(v) for v in RUBRIC_GROUPS.values()),
        })
    except Exception as e:
        return jsonify({"ok": False, "erro": str(e)})
