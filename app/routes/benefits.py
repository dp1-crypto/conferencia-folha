from flask import Blueprint, request, jsonify

from app.services.benefits_parser import (
    parse_plano_fatura, parse_extrato_plano, parse_referencia_simples,
    _merge_fatura, _merge_extrato,
)
from app.services.benefits_comparator import compare_plano_saude

benefits_bp = Blueprint('benefits', __name__)


@benefits_bp.route("/comparar-beneficio", methods=["POST"])
def comparar_beneficio():
    errors = []
    fatura_data = {}
    extrato_data = {}

    # Multiplos arquivos de referencia
    fatura_files = request.files.getlist("fatura")
    extrato_files = request.files.getlist("extrato")

    # Regra de desconto esperado
    regra_tipo  = request.form.get("regra_tipo", "fatura")
    regra_valor = float(request.form.get("regra_valor", 0) or 0)
    regra = {"tipo": regra_tipo, "valor": regra_valor}

    # Codigo do evento a analisar no extrato
    evento_codigo = request.form.get("evento_codigo", "8111").strip() or "8111"

    # Palavra-chave do documento de referencia (default: MENSALIDADE)
    filtro_linha = request.form.get("filtro_linha", "MENSALIDADE").strip() or "MENSALIDADE"

    if not extrato_files or not any(f.filename for f in extrato_files):
        return jsonify({"error": "Envie ao menos o Extrato de Folha PDF."})

    for f in fatura_files:
        if not f.filename:
            continue
        try:
            fn = f.filename.lower()
            raw = f.read()
            if fn.endswith((".xls", ".xlsx")):
                fatura_data = _merge_fatura(fatura_data, parse_referencia_simples(raw, f.filename))
            else:
                fatura_data = _merge_fatura(fatura_data, parse_plano_fatura(raw, filtro_linha))
        except Exception as e:
            errors.append(f"{f.filename}: {e}")

    for f in extrato_files:
        if not f.filename:
            continue
        try:
            extrato_data = _merge_extrato(extrato_data, parse_extrato_plano(f.read(), evento_codigo))
        except Exception as e:
            errors.append(f"{f.filename}: {e}")

    if not extrato_data:
        return jsonify({"error": "Nenhum dado extraido do extrato. " + " | ".join(errors)})

    result = compare_plano_saude(fatura_data, extrato_data, regra)
    result["erros"] = errors
    # Debug: valores extraidos de cada arquivo
    result["debug_fatura"] = {
        k: {"nome": v.get("nome_original"), "valor": v.get("total", 0)}
        for k, v in fatura_data.items()
    }
    result["debug_extrato"] = {
        k: {"nome": v.get("nome_original"), "valor_descontado": v.get("plano_descontado", 0), "salario": v.get("salario", 0)}
        for k, v in extrato_data.items()
    }
    result["evento_codigo"] = evento_codigo
    return jsonify(result)
