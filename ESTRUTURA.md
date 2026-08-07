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

## Módulo Implantação — Cliente Novo (ago/2026)

Lê contracheques de **qualquer** sistema de folha (escritório anterior) e monta
a análise para migrar a folha sem erro.

### Serviços

| Arquivo | O que faz |
|---------|-----------|
| `app/core/textutils.py` | Competência, CPF/CNPJ com dígito verificador, valores BR |
| `app/services/pdf_grid.py` | Extração por **coordenadas** (words → linhas → colunas) + OCR |
| `app/services/layout_profiles.py` | Fingerprint do sistema de origem + layouts aprendidos |
| `app/services/holerite_parser.py` | Holerite → recibos normalizados + validação de integridade |
| `app/services/implantacao.py` | Rubricas fixas × variáveis, variações do período, alertas |
| `app/services/implantacao_export.py` | Mapa de rubricas, planilha de importação, fichas PDF |
| `app/routes/implantacao_routes.py` | `/implantacao/analisar`, `/exportar/<tipo>`, `/layouts` |
| `layouts-folha.json` | 13 perfis de sistemas + rótulos globais + aprendidos |

### Como ele não erra calado

1. **Coordenadas, não regex fixo** — a coluna em que o valor foi impresso separa
   provento de desconto, funcionando em layout desconhecido.
2. **Posição dos totais define as colunas** — inclusive dentro de cada seção.
   Não dá para assumir "a coluna mais à direita é o valor": o espelho do SCI
   imprime FGTS depois dos descontos e ela roubaria o papel.
3. **Validação de integridade obrigatória** — `proventos − descontos = líquido
   impresso`. Não fechou, o recibo vira "leitura duvidosa" e aparece em alerta.
4. **Conferência contra o resumo do relatório** — a soma extraída tem que bater
   com o RESUMO GERAL impresso, em valor e em quantidade de funcionários. É o
   que prova que nenhuma página ficou de fora.
5. **Autocorreção de classificação** — se a soma dos descontos não bate por
   exatamente o valor de uma rubrica, ela é reclassificada e o ajuste é logado.
6. **Chave por CPF** (fallback matrícula, depois nome) — homônimo não colide.
   Rubrica é chaveada pelo **código de origem**, que não sofre quebra de texto.
7. **Texto quebrado é reparado antes de tudo** — `599 , 72` → `599,72`,
   `F o l h a` → `Folha`. Detecção estrutural compara sem espaços, senão
   `Total de pro v entos` deixa de ser fim de bloco e funde dois funcionários.

### Formatos suportados

| Formato | Exemplo | Como é lido |
|---------|---------|-------------|
| Holerite, duas colunas | Questor, Alterdata | Vencimentos e Descontos em colunas |
| Holerite, coluna única | Domínio | Lado sai pelo nome da rubrica |
| Espelho da folha | SCI Visual Practice | Duas tabelas lado a lado, N func/página |

### Validado com folha real (ago/2026)

3 empresas × 2 competências, 438 recibos: **100% com integridade OK** e soma
idêntica ao resumo do relatório (funcionários, proventos e descontos).
1 recibo ficou sem nome identificado — sinalizado em alerta, valores corretos.

### Classificação de rubricas

| Classe | Significado | O que fazer |
|--------|-------------|-------------|
| `fixa` | Mesmo valor todos os meses | Evento fixo no cadastro |
| `fixa_reajustada` | Constante que mudou de patamar | Evento fixo com o último valor |
| `variavel` | Muda todo mês | Lançar mês a mês |
| `eventual` | Só em alguns meses, sem continuidade | Lançar quando ocorrer |
| `calculada` | INSS / IRRF / FGTS | Não importar — o destino recalcula |

### Pendente

- Preencher `codigo_destino` de cada grupo em `rubricas-equivalentes.json` com
  os códigos do Domínio da Sigma. Sem isso o mapa sai com a coluna em branco.
- Cadastrar as rubricas do SCI ainda sem equivalência (Ajuda de Custos,
  Seg. de Vida em Grupo, FERIADO HE, INT INTRAJORNADA, Gratificações).
- OCR exige `brew install tesseract tesseract-lang && pip3 install pytesseract`.
- Aba de conferência assistida (ensinar layout pela tela) — backend pronto
  (`/implantacao/layouts/aprender`), falta a interface.
- Deploy na VPS não foi feito.

### Premissa importante

Uma análise = **uma empresa**. A rubrica é chaveada pelo código de origem, que
só é único dentro do mesmo sistema/empresa. Misturar arquivos de clientes
diferentes na mesma análise faz códigos colidirem.

---

## Arquivos do projeto

| Arquivo | O que é |
|---------|---------|
| `app.py` | Shim de compatibilidade; o app real está em `app/` |
| `rubricas-equivalentes.json` | 43 grupos de rubricas (v2.0), com `codigo_destino` |
| `layouts-folha.json` | Perfis de layout de folha por sistema de origem |
| `tests.py` | 54 testes das abas originais |
| `tests_implantacao.py` | 48 testes do motor de implantação |
| `tests_fixtures.py` | Gerador de holerites sintéticos para teste |
| `requirements.txt` | Dependências Python |
| `Dockerfile` | Build para deploy |

---

## Infraestrutura

- **Produção:** https://conferencia.gsigma.com.br
- **VPS:** `ssh -i ~/.ssh/id_sigma_vps -p 22022 root@129.121.54.101`
- **Código na VPS:** `/root/conferencia-folha/`
- **Serviço:** `systemd` → `conferencia-folha` (porta 5096 → Traefik → HTTPS)
- **Deploy:** `cd /root/conferencia-folha && git pull && pip3 install -r requirements.txt && systemctl restart conferencia-folha`

> ⚠️ **O gunicorn precisa rodar com `--workers 1 --threads 8`.**
> A aba Implantação guarda a análise em memória para os botões de exportação
> reaproveitarem sem reprocessar os PDFs (assim nada de folha vai para o disco).
> Com 2 workers, o export cai no processo que não tem a análise e devolve 400.
> O `--timeout 300` também é necessário: analisar 12 meses de uma folha grande
> passa fácil dos 30s padrão e o worker seria morto no meio.
>
> `ExecStart=/usr/local/bin/gunicorn -b 0.0.0.0:5096 run:app --workers 1 --threads 8 --timeout 300`
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
