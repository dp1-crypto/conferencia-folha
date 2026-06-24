# Estrutura — app.py (Conferência de Folha)

> Arquivo único com 2.672 linhas. Abaixo o mapa completo para planejar refatoração.

---

## Bloco 1 — Config e Utilitários (L1–253)

| Linha | O que é |
|-------|---------|
| 1–18 | Imports + inicialização Flask |
| 19–108 | Carrega `rubricas-equivalentes.json` → `RUBRIC_GROUPS` e `RUBRIC_META` |
| 52–108 | Funções utilitárias: `norm`, `brl`, `brl_dec`, `fmt_brl` |
| 109–253 | Motor de rubricas: `normalize_rubric`, `rubric_words_overlap`, `find_rubric_by_value` |

---

## Bloco 2 — Módulo Folha (L254–993)

| Linha | O que é |
|-------|---------|
| 254–393 | `parse_excel` — lê planilha de comissões |
| 394–531 | `parse_pdf` — extrai recibos de salário |
| 532–629 | `parse_word` — lê Word de gratificações |
| 630–665 | `match_names` — cruza nomes Excel × PDF |
| 666–993 | `compare` — lógica principal de comparação (~330 linhas, mais complexa) |

---

## Bloco 3 — Módulo Benefício (L994–1503)

| Linha | O que é |
|-------|---------|
| 994–1003 | `fix_spaced` — normaliza texto PDF espaçado |
| 1004–1052 | `_page_lines_smart` — parser inteligente de linhas de PDF |
| 1053–1087 | `_merge_fatura` / `_merge_extrato` — merge de múltiplos arquivos |
| 1088–1207 | `parse_plano_fatura` — lê fatura Unimed analítico |
| 1208–1263 | `parse_extrato_plano` — lê extrato de folha (rubrica 8111) |
| 1264–1317 | `parse_referencia_simples` — fallback para Excel de benefício |
| 1318–1380 | `_abbrev_match` + `match_names_beneficio` — cruza nomes fatura × folha |
| 1381–1503 | `compare_plano_saude` — lógica de comparação de benefício |

---

## Bloco 4 — Rotas Flask (L1504–1629)

| Linha | Rota |
|-------|------|
| 1504 | `GET /` — serve o HTML |
| 1508 | `POST /comparar-beneficio` — endpoint benefício |
| 1570 | `POST /analisar` — endpoint folha |
| 1612 | `POST /recarregar-rubricas` — recarga JSON sem restart |

---

## Bloco 5 — Template HTML/CSS/JS (L1633–2672)

| Linha | O que é |
|-------|---------|
| 1633–1639 | Variável `HTML = r"""...` |
| 1640–1781 | `<style>` — todo o CSS inline |
| 1785–1951 | HTML estrutural (header, tabs, formulários, divs de resultado) |
| 1952–2663 | `<script>` — todo o JS inline (fetch, render de tabelas, lógica de UI) |

---

## Arquivos do projeto

| Arquivo | O que é |
|---------|---------|
| `app.py` | Tudo — backend, rotas, template HTML/CSS/JS |
| `rubricas-equivalentes.json` | 15 grupos de rubricas, editável em runtime via `/recarregar-rubricas` |
| `tests.py` | 54 testes unitários (546 linhas), unittest puro |
| `requirements.txt` | Dependências Python |
| `Dockerfile` | Build para deploy |

---

## Infraestrutura

- **Produção:** https://conferencia.gsigma.com.br
- **VPS:** `ssh -i ~/.ssh/id_sigma_vps -p 22022 root@129.121.54.101`
- **Código na VPS:** `/root/conferencia-folha/`
- **Serviço:** `systemd` → `conferencia-folha` (porta 5096 → Traefik → HTTPS)
- **Deploy:** `cd /root/conferencia-folha && git pull && systemctl restart conferencia-folha`
- **GitHub:** https://github.com/dp1-crypto/conferencia-folha

---

## Possíveis direções de melhoria

- Separar `app.py` em módulos (`rubrics.py`, `folha.py`, `beneficio.py`, `routes.py`)
- Mover template HTML para arquivo separado (`templates/index.html`)
- Adicionar nova aba / funcionalidade
- Melhorar parser de PDF / Excel
- Novos tipos de rubrica no JSON
- Filtros na interface (por colaborador, status, rubrica)
- Extração automática de competência/mês dos documentos
