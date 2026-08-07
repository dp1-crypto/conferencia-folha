"""
Perfis de layout de folha — deteccao por impressao digital e aprendizado
Sigma Contabilidade

O motor de extracao funciona SEM perfil (modo generico). O perfil serve para:
  1. elevar a confianca da leitura;
  2. guardar o que o usuario ensinou na conferencia assistida, para o
     proximo arquivo daquele mesmo escritorio de origem sair certo sozinho.
"""
import os
import json
import hashlib
import threading

from app.core.textutils import flat

_JSON_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(__file__))),
    "layouts-folha.json",
)
_LOCK = threading.Lock()

_CACHE = {"perfis": {}, "aprendidos": {}, "rotulos": {}, "mtime": None}


def _load(force: bool = False) -> dict:
    """Carrega o JSON de layouts, recarregando quando o arquivo muda no disco."""
    try:
        mtime = os.path.getmtime(_JSON_PATH)
    except OSError:
        return {"perfis": {}, "aprendidos": {}, "rotulos": {}}

    if not force and _CACHE["mtime"] == mtime:
        return _CACHE

    try:
        with open(_JSON_PATH, encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"[AVISO] layouts-folha.json ilegivel: {e}")
        return _CACHE

    _CACHE["perfis"] = data.get("perfis", {})
    _CACHE["aprendidos"] = data.get("aprendidos", {})
    _CACHE["rotulos"] = data.get("rotulos_globais", {})
    _CACHE["mtime"] = mtime
    return _CACHE


def rotulos() -> dict:
    """Dicionario global de rotulos (total de vencimentos, liquido, etc)."""
    return _load().get("rotulos", {})


def perfis() -> dict:
    d = _load()
    todos = dict(d.get("perfis", {}))
    todos.update(d.get("aprendidos", {}))
    return todos


def reload_layouts():
    _load(force=True)


# ─────────────────────────────────────────────
# IMPRESSAO DIGITAL
# ─────────────────────────────────────────────

def fingerprint(texto: str) -> str:
    """
    Assinatura estavel do layout: rotulos fixos do documento, sem os dados.
    Serve de chave para o perfil aprendido — dois holerites do mesmo sistema
    e da mesma empresa produzem a mesma assinatura mesmo mudando funcionario.
    """
    t = flat(texto)
    # Remove tudo que varia entre funcionarios: numeros e valores
    palavras = [p for p in t.split() if p.isalpha() and len(p) > 3]
    # Mantem apenas os rotulos mais frequentes/estruturais do inicio do doc
    assinatura = " ".join(sorted(set(palavras[:120])))
    return hashlib.sha1(assinatura.encode("utf-8")).hexdigest()[:16]


def detect(texto: str) -> dict:
    """
    Identifica o sistema de origem do documento.

    Retorna:
      {"id", "nome", "colunas", "confianca" (0-1), "origem": "fingerprint|aprendido|generico", ...}
    """
    t = flat(texto)
    fp = fingerprint(texto)

    d = _load()

    # 1) Perfil aprendido para esta assinatura exata — maior prioridade
    aprendidos = d.get("aprendidos", {})
    if fp in aprendidos:
        p = dict(aprendidos[fp])
        p.update({"id": fp, "confianca": 1.0, "origem": "aprendido", "fingerprint": fp})
        return p

    # 2) Fingerprint textual dos sistemas conhecidos
    melhor, melhor_score = None, 0
    for pid, p in d.get("perfis", {}).items():
        acertos = sum(1 for f in p.get("fingerprints", []) if f in t)
        if acertos > melhor_score:
            melhor, melhor_score = pid, acertos

    if melhor:
        p = dict(d["perfis"][melhor])
        p.update({
            "id": melhor,
            "confianca": min(1.0, 0.6 + 0.2 * melhor_score),
            "origem": "fingerprint",
            "fingerprint": fp,
        })
        return p

    # 3) Generico — o motor por coordenadas se vira sozinho
    return {
        "id": "generico",
        "nome": "Layout não identificado",
        "colunas": "auto",
        "confianca": 0.0,
        "origem": "generico",
        "fingerprint": fp,
    }


# ─────────────────────────────────────────────
# APRENDIZADO
# ─────────────────────────────────────────────

def salvar_aprendido(fp: str, dados: dict) -> bool:
    """
    Grava o que o usuario ensinou na conferencia assistida.
    Escrita atomica com lock — o app roda multi-thread no gunicorn.
    """
    if not fp:
        return False
    with _LOCK:
        try:
            with open(_JSON_PATH, encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            return False

        data.setdefault("aprendidos", {})
        atual = data["aprendidos"].get(fp, {})
        atual.update(dados)
        atual.setdefault("nome", dados.get("nome", "Layout aprendido"))
        atual.setdefault("colunas", dados.get("colunas", "auto"))
        data["aprendidos"][fp] = atual

        tmp = _JSON_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, _JSON_PATH)

    _load(force=True)
    return True


def esquecer_aprendido(fp: str) -> bool:
    with _LOCK:
        try:
            with open(_JSON_PATH, encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            return False
        if fp not in data.get("aprendidos", {}):
            return False
        del data["aprendidos"][fp]
        tmp = _JSON_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, _JSON_PATH)
    _load(force=True)
    return True
