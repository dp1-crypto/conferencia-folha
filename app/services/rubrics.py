"""
Gerenciamento de rubricas — carregamento, normalizacao e matching
"""
import os
import re
import json
import unicodedata

from app.core.utils import norm
from app.core.config import TOLERANCIA_CENTAVOS, TOLERANCIA_DIVERGENCIA

# ─────────────────────────────────────────────
# CARREGAMENTO DE RUBRICAS DO JSON
# ─────────────────────────────────────────────

_JSON_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "rubricas-equivalentes.json")


def _load_rubric_groups() -> dict:
    """
    Carrega grupos de rubricas do arquivo JSON editavel.
    Fallback para dict vazio se arquivo nao existir.
    """
    if not os.path.exists(_JSON_PATH):
        return {}
    try:
        with open(_JSON_PATH, encoding="utf-8") as f:
            data = json.load(f)
        grupos = data.get("grupos", {})
        # Retorna apenas as listas de variantes, compativel com o resto do codigo
        return {k: v.get("variantes", []) for k, v in grupos.items()}
    except Exception as e:
        print(f"[AVISO] Nao foi possivel carregar rubricas-equivalentes.json: {e}")
        return {}


def _load_rubric_meta() -> dict:
    """Carrega metadados dos grupos (descricao, tipo) para uso nas sugestoes."""
    if not os.path.exists(_JSON_PATH):
        return {}
    try:
        with open(_JSON_PATH, encoding="utf-8") as f:
            data = json.load(f)
        return {k: {"descricao": v.get("descricao", k), "tipo": v.get("tipo", "provento")}
                for k, v in data.get("grupos", {}).items()}
    except Exception:
        return {}


# ─── Indice invertido: texto normalizado -> chave do grupo ───────────────────
_RUBRIC_INDEX: dict = {}


def _norm_rubric_key(t: str) -> str:
    """Normalizacao interna usada para construir o indice de rubricas."""
    t = t.upper().strip()
    t = "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn")
    t = re.sub(r"[^\w\s]", "", t)
    return " ".join(t.split())


def _build_rubric_index():
    _RUBRIC_INDEX.clear()
    for group, variants in RUBRIC_GROUPS.items():
        for v in variants:
            _RUBRIC_INDEX[_norm_rubric_key(v)] = group


# Carrega grupos do JSON editavel (rubricas-equivalentes.json)
RUBRIC_GROUPS: dict = _load_rubric_groups()
RUBRIC_META:   dict = _load_rubric_meta()

_build_rubric_index()


def reload_rubric_config():
    """Recarrega rubricas do JSON sem reiniciar o app (util apos edicao manual)."""
    global RUBRIC_GROUPS, RUBRIC_META
    RUBRIC_GROUPS = _load_rubric_groups()
    RUBRIC_META   = _load_rubric_meta()
    _build_rubric_index()


def normalize_rubric(text: str) -> str:
    """
    Normaliza rubrica e retorna o grupo canonico (ex: 'COMISSAO_E_DSR')
    ou a rubrica normalizada se nao houver grupo correspondente.
    Remove codigo numerico inicial (ex: '8781 SALARIO' -> 'SALARIO').
    """
    t = str(text).upper().strip()
    # Remove acento
    t = "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn")
    # Remove codigo numerico inicial
    t = re.sub(r"^\d+\s+", "", t)
    # Remove pontuacao
    t = re.sub(r"[^\w\s]", "", t)
    # Colapsa espacos
    t = " ".join(t.split())
    return _RUBRIC_INDEX.get(t, t)


# ─────────────────────────────────────────────
# SMART MATCHING DE RUBRICAS POR VALOR
# ─────────────────────────────────────────────

_STOP_WORDS = {"DE", "DA", "DO", "DOS", "DAS", "E", "A", "O", "S", "EM",
               "NO", "NA", "NOS", "NAS", "SOBRE", "COM", "POR", "PARA"}


def _palavras_sig(rubrica: str) -> set:
    """Retorna palavras significativas de uma rubrica ja normalizada."""
    return {w for w in rubrica.split() if len(w) > 2 and w not in _STOP_WORDS}


def rubric_words_overlap(a: str, b: str) -> float:
    """
    Score de sobreposicao de palavras significativas entre duas rubricas (0.0-1.0).
    Usa as formas normalizadas (sem acento, sem codigo). 0.5+ indica provavel equivalencia.

    Exemplos:
      "PREMIO" / "PREMIO PPR"     -> 1.0  (PREMIO pertence a ambas)
      "BONUS"  / "BONUS DESEMPENHO" -> 1.0
      "DSR"    / "COMISSAO"       -> 0.0  (sem palavras em comum)
    """
    na = normalize_rubric(a)
    nb = normalize_rubric(b)
    # Se ambas ja resolvem para o mesmo grupo canonico, overlap = 1.0
    if na == nb:
        return 1.0
    wa = _palavras_sig(na)
    wb = _palavras_sig(nb)
    if not wa or not wb:
        return 0.0
    comuns = wa & wb
    if not comuns:
        return 0.0
    # Jaccard ponderado: favorece quando uma e subconjunto da outra
    return len(comuns) / max(len(wa), len(wb))


def find_rubric_by_value(
    descricao_esperada: str,
    valor_esperado: float,
    verbas: list,
    tolerance: float = None,
):
    """
    Busca uma verba no recibo por valor quando o grupo nao foi reconhecido pelo nome.

    Retorna o melhor candidato com campos:
      verba        -- dict da verba encontrada
      diff         -- diferenca de valor
      word_overlap -- score de sobreposicao de palavras (0-1)
      confianca    -- "alta" | "media" | "baixa"

    Regra de confianca:
      alta  -> diff <= TOLERANCIA_CENTAVOS  E  overlap > 0.3  (mesmo valor + palavras em comum)
      media -> diff <= tolerance            E  overlap > 0     (mesmo valor, alguma palavra em comum)
      baixa -> diff <= tolerance            (mesmo valor, sem sobreposicao de palavras)

    Exemplo: esperado "PREMIO" R$ 500,00 -> encontrado "PREMIO PPR" R$ 500,00 -> alta confianca.
    """
    if tolerance is None:
        tolerance = TOLERANCIA_DIVERGENCIA

    candidatos = []
    for v in verbas:
        diff = abs(v["valor"] - valor_esperado)
        if diff > tolerance:
            continue
        overlap = rubric_words_overlap(descricao_esperada, v["descricao"])
        if diff <= TOLERANCIA_CENTAVOS and overlap > 0.3:
            confianca = "alta"
        elif diff <= tolerance and overlap > 0:
            confianca = "media"
        else:
            confianca = "baixa"
        candidatos.append({
            "verba":        v,
            "diff":         diff,
            "word_overlap": overlap,
            "confianca":    confianca,
        })

    if not candidatos:
        return None
    # Prioriza: maior overlap -> menor diff
    candidatos.sort(key=lambda x: (-x["word_overlap"], x["diff"]))
    return candidatos[0]
