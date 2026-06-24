from flask import Blueprint, request, jsonify
from app.services.payroll_parser import parse_pdf
from app.services.previous_month import compare_months
from app.services.tax_audit import auditar_folha

audit_bp = Blueprint('audit', __name__)


@audit_bp.route('/conferir-mes-anterior', methods=['POST'])
def conferir_mes_anterior():
    errors = []
    folha_atual = {}
    folha_anterior = {}

    for f in request.files.getlist('folha_atual'):
        if f.filename:
            try:
                folha_atual.update(parse_pdf(f.read()))
            except Exception as e:
                errors.append(f'{f.filename}: {e}')

    for f in request.files.getlist('folha_anterior'):
        if f.filename:
            try:
                folha_anterior.update(parse_pdf(f.read()))
            except Exception as e:
                errors.append(f'{f.filename}: {e}')

    if not folha_atual or not folha_anterior:
        return jsonify({'error': 'Envie a folha atual e a folha do mes anterior (PDF). ' + ' | '.join(errors)})

    result = compare_months(folha_atual, folha_anterior)
    result['erros'] = errors
    return jsonify(result)


@audit_bp.route('/auditoria-impostos', methods=['POST'])
def auditoria_impostos():
    errors = []
    pdf_data = {}

    for f in request.files.getlist('pdf'):
        if f.filename and f.filename.lower().endswith('.pdf'):
            try:
                pdf_data.update(parse_pdf(f.read()))
            except Exception as e:
                errors.append(f'{f.filename}: {e}')

    if not pdf_data:
        return jsonify({'error': 'Envie os recibos em PDF. ' + ' | '.join(errors)})

    result = auditar_folha(pdf_data)
    result['erros'] = errors
    return jsonify(result)
