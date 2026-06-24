from flask import Blueprint, request, jsonify, Response
from app.services.report_export import exportar_excel, exportar_csv

export_bp = Blueprint('export', __name__)


@export_bp.route('/exportar', methods=['POST'])
def exportar():
    body = request.get_json()
    tipo = body.get('tipo', 'folha')
    formato = body.get('formato', 'excel')
    dados = body.get('dados', {})

    if formato == 'excel':
        xlsx = exportar_excel(dados, tipo)
        return Response(
            xlsx,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={'Content-Disposition': f'attachment; filename=conferencia_{tipo}.xlsx'}
        )
    elif formato == 'csv':
        csv_text = exportar_csv(dados, tipo)
        return Response(
            csv_text,
            mimetype='text/csv',
            headers={'Content-Disposition': f'attachment; filename=conferencia_{tipo}.csv'}
        )

    return jsonify({'error': 'Formato invalido. Use excel ou csv.'})
