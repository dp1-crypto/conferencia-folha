"""
Rotas da aba Implantacao — analise de folha de cliente novo
Sigma Contabilidade
"""
import io
import time
import threading

from flask import Blueprint, request, jsonify, send_file

from app.services.holerite_parser import parse_holerites
from app.services.implantacao import analisar
from app.services.implantacao_export import (
    mapa_rubricas, planilha_importacao, fichas_funcionarios, nome_arquivo,
)
from app.services import layout_profiles

implantacao_bp = Blueprint("implantacao", __name__)

# Ultima analise por sessao — os exports reaproveitam sem reprocessar os PDFs.
# Guarda em memoria com validade curta; nada de dado de folha em disco.
_CACHE = {}
_CACHE_LOCK = threading.Lock()
_CACHE_TTL = 60 * 60  # 1 hora


def _guarda(sid: str, analise: dict):
    with _CACHE_LOCK:
        agora = time.time()
        for k in [k for k, v in _CACHE.items() if agora - v["ts"] > _CACHE_TTL]:
            del _CACHE[k]
        _CACHE[sid] = {"ts": agora, "analise": analise}


def _recupera(sid: str):
    with _CACHE_LOCK:
        item = _CACHE.get(sid)
    if not item or time.time() - item["ts"] > _CACHE_TTL:
        return None
    return item["analise"]


def _sid() -> str:
    return (request.form.get("sid") or request.args.get("sid") or "default").strip()[:64]


# ─────────────────────────────────────────────
# ANALISE
# ─────────────────────────────────────────────

@implantacao_bp.route("/implantacao/analisar", methods=["POST"])
def implantacao_analisar():
    """Recebe N arquivos de folha e devolve a analise completa."""
    documentos, erros = [], []
    force_ocr = request.form.get("ocr") == "1"

    arquivos = [f for f in request.files.values() if f.filename]
    if not arquivos:
        return jsonify({"error": "Nenhum arquivo enviado."})

    for f in arquivos:
        try:
            raw = f.read()
            fn = f.filename.lower()
            if fn.endswith(".pdf"):
                documentos.append(parse_holerites(raw, f.filename, force_ocr=force_ocr))
            elif fn.endswith((".xls", ".xlsx", ".csv")):
                erros.append(
                    f"{f.filename}: planilha ainda nao suportada nesta aba — "
                    f"envie os contracheques em PDF."
                )
            else:
                erros.append(f"{f.filename}: formato nao reconhecido.")
        except Exception as e:
            erros.append(f"{f.filename}: falha ao ler ({e})")

    if not documentos:
        return jsonify({"error": "Nenhum documento pôde ser lido. " + " | ".join(erros)})

    try:
        analise = analisar(documentos)
    except Exception as e:
        return jsonify({"error": f"Falha na analise: {e}"})

    analise["erros"] = erros
    _guarda(_sid(), analise)
    return jsonify(_serializavel(analise))


def _serializavel(analise: dict) -> dict:
    """Remove estruturas nao serializaveis e enxuga o payload para a tela."""
    out = {
        "resumo": analise.get("resumo", {}),
        "competencias": analise.get("competencias", []),
        "catalogo_rubricas": analise.get("catalogo_rubricas", []),
        "eventos": analise.get("eventos", []),
        "alertas": analise.get("alertas", []),
        "qualidade": analise.get("qualidade", {}),
        "erros": analise.get("erros", []),
        "funcionarios": [],
    }
    for f in analise.get("funcionarios", []):
        out["funcionarios"].append({
            "chave": f["chave"], "nome": f["nome"], "cpf": f.get("cpf", ""),
            "matricula": f.get("matricula", ""), "cbo": f.get("cbo", ""),
            "funcao": f.get("funcao", ""), "admissao": f.get("admissao", ""),
            "meses_ativos": f.get("meses_ativos", []),
            "eventos": f.get("eventos", []),
            "rubricas": [
                {
                    "chave": k, "descricao": i.get("descricao", k),
                    "grupo": i.get("grupo", ""), "tipo": i.get("tipo", ""),
                    "classe": i["classe"], "valores": i["valores"],
                    "media": i["media"], "min": i["min"], "max": i["max"],
                    "ultimo": i["ultimo"],
                    "cobertura": f"{i['meses_presente']}/{i['meses_ativos']}",
                }
                for k, i in sorted(f.get("rubricas", {}).items())
            ],
            "totais_mes": {
                c: {
                    "proventos": round(m["proventos"], 2),
                    "descontos": round(m["descontos"], 2),
                    "liquido": round(m["liquido"], 2),
                    "salario_base": m.get("salario_base"),
                    "integridade": ("erro" if "erro" in m["integridade"] else "ok"),
                }
                for c, m in f.get("meses", {}).items()
            },
        })
    return out


# ─────────────────────────────────────────────
# EXPORTS
# ─────────────────────────────────────────────

def _envia(raw: bytes, nome: str, mime: str):
    return send_file(io.BytesIO(raw), mimetype=mime,
                     as_attachment=True, download_name=nome)


@implantacao_bp.route("/implantacao/exportar/<tipo>", methods=["GET"])
def implantacao_exportar(tipo):
    analise = _recupera(_sid())
    if not analise:
        return jsonify({"error": "Nenhuma analise em memoria. Rode a analise novamente."}), 400

    XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    try:
        if tipo == "mapa":
            return _envia(mapa_rubricas(analise), nome_arquivo(analise, "mapa-rubricas", "xlsx"), XLSX)
        if tipo == "importacao":
            return _envia(planilha_importacao(analise), nome_arquivo(analise, "importacao", "xlsx"), XLSX)
        if tipo == "fichas":
            return _envia(fichas_funcionarios(analise), nome_arquivo(analise, "fichas", "pdf"), "application/pdf")
    except Exception as e:
        return jsonify({"error": f"Falha ao gerar arquivo: {e}"}), 500

    return jsonify({"error": f"Tipo de export desconhecido: {tipo}"}), 400


# ─────────────────────────────────────────────
# LAYOUTS APRENDIDOS
# ─────────────────────────────────────────────

@implantacao_bp.route("/implantacao/layouts", methods=["GET"])
def implantacao_layouts():
    """Lista perfis conhecidos e aprendidos, para a aba Configuracoes."""
    layout_profiles.reload_layouts()
    todos = layout_profiles.perfis()
    return jsonify({
        "ok": True,
        "perfis": [
            {"id": k, "nome": v.get("nome", k), "colunas": v.get("colunas", "auto"),
             "aprendido": len(k) == 16 and all(c in "0123456789abcdef" for c in k)}
            for k, v in sorted(todos.items(), key=lambda kv: kv[1].get("nome", ""))
        ],
    })


@implantacao_bp.route("/implantacao/layouts/aprender", methods=["POST"])
def implantacao_aprender():
    """Grava o layout que o usuario ensinou na conferencia assistida."""
    dados = request.get_json(silent=True) or {}
    fp = (dados.get("fingerprint") or "").strip()
    if not fp:
        return jsonify({"ok": False, "erro": "fingerprint ausente"}), 400
    ok = layout_profiles.salvar_aprendido(fp, {
        "nome": dados.get("nome") or "Layout aprendido",
        "colunas": dados.get("colunas") or "auto",
        "nome_regex": dados.get("nome_regex") or "",
        "obs": dados.get("obs") or "",
    })
    return jsonify({"ok": ok})


@implantacao_bp.route("/implantacao/layouts/esquecer", methods=["POST"])
def implantacao_esquecer():
    dados = request.get_json(silent=True) or {}
    fp = (dados.get("fingerprint") or "").strip()
    return jsonify({"ok": layout_profiles.esquecer_aprendido(fp)})
