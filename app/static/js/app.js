/* ── Sigma Contabilidade — Conferencia de Folha — JS ── */

const FILES = {excel:[], pdf:[], word:[], fatura:[], extrato:[], atual:[], anterior:[], 'pdf-impostos':[]};

// Estado global para auditoria consolidada
const AUDIT_STATE = { folha: null, beneficio: null, mesAnterior: null, impostos: null };

function switchTab(tab) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
  var btn = document.querySelector('.tab[onclick="switchTab(\''+tab+'\')"]');
  if(btn) btn.classList.add('active');
  var sec = document.getElementById('section-'+tab);
  if(sec) sec.classList.add('active');
  // Atualiza aba auditoria quando selecionada
  if(tab === 'auditoria') renderAuditoriaConsolidada();
}

function sel(type, inp) {
  FILES[type] = Array.from(inp.files);
  var names = FILES[type].map(function(f){return f.name}).join(' \u2022 ');
  var el = document.getElementById('fname-'+type);
  if(el) el.textContent = names;
  var zone = document.getElementById('zone-'+type);
  if(zone) zone.classList.toggle('done', FILES[type].length>0);
}

function toggleEventoCodigo(){
  var s = document.getElementById('evento-select');
  var wrap = document.getElementById('evento-codigo-wrap');
  var info = document.getElementById('evento-info');
  var isCustom = s.value === 'custom';
  wrap.style.display = isCustom ? 'block' : 'none';
  info.style.display = isCustom ? 'none' : 'block';
}

function getEventoCodigo(){
  var s = document.getElementById('evento-select');
  if(s.value === 'custom'){
    return document.getElementById('evento-codigo-input').value.trim() || '8111';
  }
  return s.value;
}

function toggleRegraValor(){
  var tipo = document.getElementById('regra-tipo').value;
  var wrap = document.getElementById('regra-valor-wrap');
  var label = document.getElementById('regra-valor-label');
  var hint  = document.getElementById('regra-hint');
  var inp   = document.getElementById('regra-valor');
  if(tipo === 'fatura'){
    wrap.style.opacity='0.4'; inp.disabled=true;
    hint.textContent='O valor esperado sera lido diretamente do documento de referencia.';
  } else {
    wrap.style.opacity='1'; inp.disabled=false;
    if(tipo==='pct_fatura'){label.textContent='Percentual (%)';inp.placeholder='Ex: 15';hint.textContent='Calcula X% do valor do documento de referencia por funcionario.';}
    else if(tipo==='pct_salario'){label.textContent='Percentual (%)';inp.placeholder='Ex: 15';hint.textContent='Calcula X% do salario de cada funcionario (requer campo Salario no extrato).';}
    else if(tipo==='fixo'){label.textContent='Valor fixo (R$)';inp.placeholder='Ex: 150.00';hint.textContent='Aplica o mesmo valor esperado para todos os funcionarios.';}
  }
}

function drag(e,t){e.preventDefault();document.getElementById('zone-'+t).classList.add('over')}
function undrag(e,t){document.getElementById('zone-'+t).classList.remove('over')}
function drop(e,t){
  e.preventDefault();
  document.getElementById('zone-'+t).classList.remove('over');
  var inp=document.getElementById('file-'+t) || document.getElementById('inp-'+t);
  var dt=new DataTransfer();
  Array.from(e.dataTransfer.files).forEach(function(f){dt.items.add(f)});
  inp.files=dt.files;
  sel(t,inp);
}

function brl(v){
  if(!v&&v!==0)return'-';
  return'R$ '+Number(v).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
}

function show(html, targetId){
  targetId = targetId || 'results';
  var el=document.getElementById(targetId);
  el.style.display='block';
  el.innerHTML=html;
  el.scrollIntoView({behavior:'smooth'});
}

/* ═══════════════════════════════════════════
   ABA 1: CONFERENCIA DE FOLHA
   ═══════════════════════════════════════════ */

async function analyze(){
  if(!FILES.excel.length && !FILES.pdf.length){alert('Envie pelo menos a planilha Excel ou os recibos PDF.');return}
  var btn=document.getElementById('btn');
  btn.disabled=true;
  document.getElementById('loading').style.display='block';
  document.getElementById('results').style.display='none';

  var fd=new FormData();
  FILES.excel.forEach(function(f,i){fd.append('excel_'+i,f)});
  FILES.pdf.forEach(function(f,i){fd.append('pdf_'+i,f)});
  FILES.word.forEach(function(f,i){fd.append('word_'+i,f)});

  try{
    var res=await fetch('/analisar',{method:'POST',body:fd});
    var data=await res.json();
    AUDIT_STATE.folha = data;
    render(data);
  } catch(e){
    show('<div class="err-box"><h4>Erro de comunicacao</h4><p>'+e.message+'</p></div>');
  } finally{
    btn.disabled=false;
    document.getElementById('loading').style.display='none';
  }
}

function render(data){
  if(data.error){show('<div class="err-box"><h4>Erro</h4><p>'+data.error+'</p></div>');return}

  var r=data.resumo;
  var html='<div class="stats">'
    +'<div class="stat t"><div class="n">'+r.total+'</div><div class="l">Funcionarios analisados</div></div>'
    +'<div class="stat g"><div class="n">'+r.ok+'</div><div class="l">Sem divergencias</div></div>'
    +'<div class="stat d"><div class="n">'+r.divergencias+'</div><div class="l">Com divergencias</div></div>'
    +'</div>';

  // Gratificacoes do Word
  var wg=data.word_gratificacoes||{};
  if(Object.keys(wg).length){
    html+='<div class="card"><div class="sec-title">Gratificacoes informadas no Word</div>'
      +'<table class="gratif-table"><tr><th>Funcionario</th><th>Valor</th><th>No recibo?</th></tr>';
    Object.entries(wg).forEach(function(kv){
      var n=kv[0], v=kv[1];
      var rec=data.funcionarios.find(function(f){return f.nome===n});
      var ok=rec&&rec.dados_recibo&&rec.dados_recibo.has_gratif;
      html+='<tr><td>'+(rec?rec.nome_exibir:n)+'</td><td>'+v+'</td>'
        +'<td>'+(ok?'<span style="color:#059669;font-weight:600">Sim</span>':'<span style="color:#A72C31;font-weight:600">Nao</span>')+'</td></tr>';
    });
    html+='</table></div>';
  }

  // Observacoes
  if(data.observacoes&&data.observacoes.length){
    html+='<div class="card"><div class="sec-title">Observacoes do documento Word</div>';
    data.observacoes.forEach(function(o){html+='<div class="obs-box">'+o+'</div>'});
    html+='</div>';
  }

  // Funcionarios
  html+='<div class="card"><div class="sec-title">Resultado por funcionario</div>';

  var sorted=data.funcionarios.slice().sort(function(a,b){
    if(a.status!==b.status) return a.status==='DIVERGENTE'?-1:1;
    return a.nome.localeCompare(b.nome);
  });

  sorted.forEach(function(emp){
    var divg=emp.status==='DIVERGENTE';
    var badge=divg
      ?'<span class="badge div">'+emp.divs.length+' divergencia'+(emp.divs.length>1?'s':'')+'</span>'
      :'<span class="badge ok">OK</span>';

    var recTipo={mensal:'Mensal','13_adiantamento':'13o Adiant.',ferias:'Ferias'}[emp.dados_recibo?emp.dados_recibo.tipo:'']||'';

    html+='<div class="emp'+(divg?' open':'')+'" onclick="tog(this)">'
      +'<div class="emp-hdr"><div style="flex:1;min-width:0">'
      +'<div class="emp-name">'+(emp.nome_exibir||emp.nome)+'</div>'
      +'<div class="emp-sub">'
      +(emp.dados_excel?'Planilha: '+brl(emp.dados_excel.liquido):'Sem planilha')
      +(emp.dados_recibo?' | Recibo: '+brl(emp.dados_recibo.liquido)+(recTipo?' ('+recTipo+')':''):'| Sem recibo')
      +'</div></div>'+badge+'</div>';

    if(divg||emp.dados_excel||emp.dados_recibo){
      html+='<div class="emp-body">';
      emp.divs.forEach(function(d){
        html+='<div class="div-item '+d.g+'"><div><div class="div-tipo">'+d.tipo+'</div><div class="div-desc">'+d.desc+'</div></div></div>';
      });

      if(emp.dados_excel || emp.dados_recibo){
        html+='<details style="margin-top:.6rem"><summary style="font-size:.75rem;color:#6b7280;cursor:pointer;user-select:none;padding:.25rem 0;list-style:none">'
          +'<span style="text-decoration:underline;text-decoration-style:dotted">Ver dados completos (planilha / recibo)</span></summary>'
          +'<div class="data-grid" style="margin-top:.6rem">';
        if(emp.dados_excel){
          var e=emp.dados_excel;
          html+='<div class="dtbl"><h4>Planilha Excel</h4><table>'
            +'<tr><td>Salario</td><td>'+brl(e.salario)+'</td></tr>'
            +(e.gratificacao?'<tr><td>Gratificacao</td><td>'+brl(e.gratificacao)+'</td></tr>':'')
            +(e.ferias_13?'<tr><td>Ferias / 13o</td><td>'+brl(e.ferias_13)+'</td></tr>':'')
            +(e.inss?'<tr><td>INSS</td><td>- '+brl(e.inss)+'</td></tr>':'')
            +(e.vale?'<tr><td>Vale Transp.</td><td>- '+brl(e.vale)+'</td></tr>':'')
            +(e.plano?'<tr><td>Plano/Unimed</td><td>- '+brl(e.plano)+'</td></tr>':'')
            +(e.emprestimo?'<tr><td>Emprestimo</td><td>- '+brl(e.emprestimo)+'</td></tr>':'')
            +'<tr><td>Liquido</td><td>'+brl(e.liquido)+'</td></tr></table></div>';
        }
        if(emp.dados_recibo){
          var rc=emp.dados_recibo;
          html+='<div class="dtbl"><h4>Recibo PDF</h4><table>'
            +'<tr><td>Total Vencimentos</td><td>'+brl(rc.total_vencimentos)+'</td></tr>'
            +'<tr><td>Total Descontos</td><td>- '+brl(rc.total_descontos)+'</td></tr>'
            +'<tr><td>Liquido</td><td>'+brl(rc.liquido)+'</td></tr></table>';
          if(rc.verbas&&rc.verbas.length){
            html+='<details><summary style="font-size:.72rem;cursor:pointer;color:#6b7280">Ver '+rc.verbas.length+' verbas</summary>'
              +'<table style="margin-top:.4rem;font-size:.73rem">';
            rc.verbas.forEach(function(v){
              html+='<tr><td>'+v.codigo+' -- '+v.descricao+'</td><td style="text-align:right;padding-left:.5rem">'+brl(v.valor)+'</td></tr>';
            });
            html+='</table></details>';
          }
          html+='</div>';
        }
        html+='</div></details>';
      }
      html+='</div>';
    }
    html+='</div>';
  });

  html+='</div>';

  // Sugestoes de equivalencia
  var sug=data.sugestoes_equivalencia||[];
  if(sug.length){
    html+='<div class="card" style="border-left:4px solid #3b82f6">'
      +'<div class="sec-title" style="color:#1d4ed8">Sugestoes de equivalencia de rubricas</div>'
      +'<p style="font-size:.8rem;color:#374151;margin-bottom:.9rem">O sistema encontrou rubricas com valores iguais mas nomes diferentes. Adicione ao arquivo <strong>rubricas-equivalentes.json</strong>.</p>'
      +'<table style="width:100%;border-collapse:collapse;font-size:.8rem"><thead><tr style="background:#eff6ff">'
      +'<th style="text-align:left;padding:.4rem .6rem;color:#1e40af">Colaborador</th>'
      +'<th style="text-align:left;padding:.4rem .6rem;color:#1e40af">Esperado</th>'
      +'<th style="text-align:left;padding:.4rem .6rem;color:#1e40af">Encontrado</th>'
      +'<th style="text-align:right;padding:.4rem .6rem;color:#1e40af">Valor</th>'
      +'<th style="text-align:center;padding:.4rem .6rem;color:#1e40af">Confianca</th>'
      +'</tr></thead><tbody>';
    sug.forEach(function(s){
      var conf={alta:'Alta',media:'Media',baixa:'Baixa'}[s.confianca]||s.confianca;
      html+='<tr style="border-bottom:1px solid #f0f0f0">'
        +'<td style="padding:.35rem .6rem">'+s.colaborador+'</td>'
        +'<td style="padding:.35rem .6rem;font-weight:600">'+s.esperado+'</td>'
        +'<td style="padding:.35rem .6rem;color:#059669">'+s.encontrado+'</td>'
        +'<td style="padding:.35rem .6rem;text-align:right">'+brl(s.valor)+'</td>'
        +'<td style="padding:.35rem .6rem;text-align:center">'+conf+'</td></tr>';
    });
    html+='</tbody></table>'
      +'<div style="margin-top:.8rem;font-size:.73rem;color:#6b7280;background:#f8faff;border-radius:8px;padding:.6rem .8rem">'
      +'<strong>Como aplicar:</strong> Edite o <code>rubricas-equivalentes.json</code> e clique em '
      +'<button onclick="recarregarRubricas()" style="background:#3b82f6;color:#fff;border:none;border-radius:5px;padding:.2rem .6rem;font-size:.72rem;cursor:pointer;font-weight:600">Recarregar config</button></div></div>';
  }

  // Erros
  if(data.erros&&data.erros.length){
    html+='<div class="err-box"><h4>Avisos de processamento</h4>';
    data.erros.forEach(function(e){html+='<p>'+e+'</p>'});
    html+='</div>';
  }

  // Botoes exportar + imprimir
  html+='<div class="export-bar">'
    +'<button class="btn-export" onclick="exportar(\'folha\',\'excel\')">Exportar Excel</button>'
    +'<button class="btn-export" onclick="exportar(\'folha\',\'csv\')">Exportar CSV</button>'
    +'<button class="btn-export" onclick="window.print()">Imprimir / PDF</button>'
    +'</div>';

  show(html);
}

/* ═══════════════════════════════════════════
   ABA 2: MES ANTERIOR
   ═══════════════════════════════════════════ */

async function analisarMesAnterior(){
  if(!FILES.atual.length || !FILES.anterior.length){alert('Envie a folha atual e a folha do mes anterior (PDF).');return}
  var btn=document.getElementById('btn-anterior');
  btn.disabled=true;
  document.getElementById('loading-anterior').style.display='block';
  document.getElementById('results-anterior').innerHTML='';

  var fd=new FormData();
  FILES.atual.forEach(function(f){fd.append('folha_atual',f)});
  FILES.anterior.forEach(function(f){fd.append('folha_anterior',f)});

  try{
    var res=await fetch('/conferir-mes-anterior',{method:'POST',body:fd});
    var data=await res.json();
    AUDIT_STATE.mesAnterior = data;
    renderMesAnterior(data);
  } catch(e){
    document.getElementById('results-anterior').innerHTML='<div class="err-box"><h4>Erro</h4><p>'+e.message+'</p></div>';
  } finally{
    btn.disabled=false;
    document.getElementById('loading-anterior').style.display='none';
  }
}

function renderMesAnterior(data){
  var el=document.getElementById('results-anterior');
  if(data.error){el.innerHTML='<div class="err-box"><h4>Erro</h4><p>'+data.error+'</p></div>';return}

  var r=data.resumo;

  // ── Cards de resumo ──────────────────────────────────────────────────────
  var html='<div class="stats" style="grid-template-columns:repeat(5,1fr)">'
    +'<div class="stat t"><div class="n">'+r.total_atual+'</div><div class="l">Total na folha</div></div>'
    +'<div class="stat" style="--c:#10b981"><div class="n" style="color:#10b981">'+r.novos+'</div><div class="l">Novos</div></div>'
    +'<div class="stat" style="--c:#A72C31"><div class="n" style="color:#A72C31">'+r.desligados+'</div><div class="l">Desligados</div></div>'
    +'<div class="stat" style="--c:#f59e0b"><div class="n" style="color:#f59e0b">'+r.alterados+'</div><div class="l">Com alteração</div></div>'
    +'<div class="stat" style="--c:#6b7280"><div class="n" style="color:#6b7280">'+r.sem_alteracao+'</div><div class="l">Sem alteração</div></div>'
    +'</div>';

  // ── Novos e Desligados lado a lado ───────────────────────────────────────
  var temNovos = data.colaboradores_novos&&data.colaboradores_novos.length;
  var temDeslig = data.colaboradores_desligados&&data.colaboradores_desligados.length;
  if(temNovos||temDeslig){
    html+='<div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-bottom:1.25rem">';
    if(temNovos){
      html+='<div class="card" style="margin-bottom:0"><div class="sec-title" style="color:#059669">&#10010; Novos Colaboradores</div>';
      data.colaboradores_novos.forEach(function(n){
        html+='<div style="display:flex;justify-content:space-between;align-items:center;padding:.5rem .4rem;border-bottom:1px solid #f3f4f6">'
          +'<span style="font-weight:600;font-size:.83rem">'+n.nome+'</span>'
          +'<span class="badge-novo" style="font-size:.7rem">'+brl(n.liquido)+'</span></div>';
      });
      html+='</div>';
    } else { html+='<div></div>'; }
    if(temDeslig){
      html+='<div class="card" style="margin-bottom:0"><div class="sec-title" style="color:#A72C31">&#10006; Desligados</div>';
      data.colaboradores_desligados.forEach(function(d){
        html+='<div style="display:flex;justify-content:space-between;align-items:center;padding:.5rem .4rem;border-bottom:1px solid #f3f4f6">'
          +'<span style="font-weight:600;font-size:.83rem">'+d.nome+'</span>'
          +'<span class="badge-desligado" style="font-size:.7rem">'+brl(d.liquido)+'</span></div>';
      });
      html+='</div>';
    }
    html+='</div>';
  }

  // ── Tabela comparativa — um colaborador por linha ────────────────────────
  if(data.comparativo&&data.comparativo.length){
    html+='<div class="card"><div class="sec-title">Comparativo por Colaborador</div>'
      +'<div style="overflow-x:auto">'
      +'<table class="tbl-comp"><thead><tr>'
      +'<th rowspan="2" style="min-width:160px">Colaborador</th>'
      +'<th rowspan="2" style="text-align:center;width:80px">Status</th>'
      +'<th colspan="3" style="text-align:center;border-left:2px solid #e5e7eb">Líquido a Receber</th>'
      +'<th colspan="3" style="text-align:center;border-left:2px solid #e5e7eb">Total Vencimentos</th>'
      +'<th colspan="3" style="text-align:center;border-left:2px solid #e5e7eb">Total Descontos</th>'
      +'<th rowspan="2" style="text-align:center;width:60px">Rubricas</th>'
      +'</tr><tr>'
      +'<th style="text-align:right;border-left:2px solid #e5e7eb;font-weight:500;color:#6b7280">Anterior</th>'
      +'<th style="text-align:right;font-weight:500;color:#6b7280">Atual</th>'
      +'<th style="text-align:right;font-weight:600">Δ</th>'
      +'<th style="text-align:right;border-left:2px solid #e5e7eb;font-weight:500;color:#6b7280">Anterior</th>'
      +'<th style="text-align:right;font-weight:500;color:#6b7280">Atual</th>'
      +'<th style="text-align:right;font-weight:600">Δ</th>'
      +'<th style="text-align:right;border-left:2px solid #e5e7eb;font-weight:500;color:#6b7280">Anterior</th>'
      +'<th style="text-align:right;font-weight:500;color:#6b7280">Atual</th>'
      +'<th style="text-align:right;font-weight:600">Δ</th>'
      +'</tr></thead><tbody>';

    data.comparativo.forEach(function(c){
      var rowCls = c.criticidade==='alta'?'comp-row-alta'
                 : c.criticidade==='media'?'comp-row-media'
                 : c.criticidade==='baixa'?'comp-row-baixa'
                 : 'comp-row-ok';
      var badgeCls = c.criticidade==='alta'?'badge-alta'
                   : c.criticidade==='media'?'badge-media'
                   : c.criticidade==='baixa'?'badge-baixa'
                   : 'badge ok';
      var badgeTxt = c.criticidade==='alta'?'CRÍTICO'
                   : c.criticidade==='media'?'ATENÇÃO'
                   : c.criticidade==='baixa'?'BAIXO'
                   : 'OK';

      function deltaCell(diff, pct, borderLeft){
        if(Math.abs(diff)<0.06) return '<td style="text-align:right'+(borderLeft?';border-left:2px solid #e5e7eb':'')+'"></td><td style="text-align:right"></td><td style="text-align:right;color:#9ca3af">—</td>';
        var cls = diff>0?'pct-up':'pct-down';
        var arrow = diff>0?'▲':'▼';
        var pctStr = pct!==0?' <small style="font-size:.68rem">('+pct+'%)</small>':'';
        return '<td style="text-align:right'+(borderLeft?';border-left:2px solid #e5e7eb':'')+'">'
          +brl(diff>0?c.liq_ant:c.venc_ant||c.desc_ant||0)+'</td>'  // placeholder — veja abaixo
          +'<td style="text-align:right"></td>'
          +'<td style="text-align:right" class="'+cls+'">'+arrow+' '+brl(Math.abs(diff))+pctStr+'</td>';
      }

      // Células individuais por campo
      function fieldCells(ant, atu, diff, pct, borderLeft){
        var diffCls = diff>0?'pct-up':diff<0?'pct-down':'';
        var arrow = diff>0?'▲':diff<0?'▼':'';
        var pctStr = Math.abs(pct)>=1?' <small style="font-size:.68rem">('+pct+'%)</small>':'';
        var diffFmt = Math.abs(diff)<0.06
          ? '<span style="color:#9ca3af">—</span>'
          : '<span class="'+diffCls+'">'+arrow+' '+brl(Math.abs(diff))+pctStr+'</span>';
        return '<td style="text-align:right'+(borderLeft?';border-left:2px solid #e5e7eb':'')+'">'+brl(ant)+'</td>'
          +'<td style="text-align:right">'+brl(atu)+'</td>'
          +'<td style="text-align:right">'+diffFmt+'</td>';
      }

      // Rubricas novas/removidas — tooltip compacto
      var rbCount = (c.rubricas_novas||[]).length+(c.rubricas_removidas||[]).length;
      var rbCell = '';
      if(rbCount){
        var rbTip = '';
        (c.rubricas_novas||[]).forEach(function(rb){rbTip+='+ '+rb.rubrica+' ('+brl(rb.valor)+')\n'});
        (c.rubricas_removidas||[]).forEach(function(rb){rbTip+='- '+rb.rubrica+' ('+brl(rb.valor)+')\n'});
        rbCell='<td style="text-align:center"><span title="'+rbTip.trim()+'" style="cursor:help;background:#eff6ff;color:#3b82f6;border-radius:20px;padding:.15rem .55rem;font-size:.72rem;font-weight:700">'+rbCount+'</span></td>';
      } else {
        rbCell='<td style="text-align:center;color:#9ca3af;font-size:.75rem">—</td>';
      }

      html+='<tr class="'+rowCls+'">'
        +'<td style="font-weight:600;font-size:.83rem">'+c.nome+'</td>'
        +'<td style="text-align:center"><span class="'+badgeCls+'" style="font-size:.68rem">'+badgeTxt+'</span></td>'
        +fieldCells(c.liq_ant,  c.liq_atu,  c.liq_diff,  c.liq_pct,  true)
        +fieldCells(c.venc_ant, c.venc_atu, c.venc_diff, c.venc_pct, true)
        +fieldCells(c.desc_ant, c.desc_atu, c.desc_diff, c.desc_pct, true)
        +rbCell
        +'</tr>';
    });
    html+='</tbody></table></div>'
      +'<p style="font-size:.72rem;color:#9ca3af;margin-top:.6rem">▲ aumento &nbsp;▼ redução &nbsp;·&nbsp; Rubricas: número de verbas novas/removidas (passe o mouse para ver detalhes)</p>'
      +'</div>';
  }

  if(data.erros&&data.erros.length){
    html+='<div class="err-box"><h4>Avisos</h4>';
    data.erros.forEach(function(e){html+='<p>'+e+'</p>'});
    html+='</div>';
  }

  html+='<div class="export-bar">'
    +'<button class="btn-export" onclick="exportar(\'mes_anterior\',\'excel\')">Exportar Excel</button>'
    +'<button class="btn-export" onclick="exportar(\'mes_anterior\',\'csv\')">Exportar CSV</button>'
    +'</div>';

  el.innerHTML=html;
  el.scrollIntoView({behavior:'smooth'});
}

/* ═══════════════════════════════════════════
   ABA 3: INSS / IRRF
   ═══════════════════════════════════════════ */

async function analisarImpostos(){
  if(!FILES['pdf-impostos'].length){alert('Envie os recibos em PDF.');return}
  var btn=document.getElementById('btn-impostos');
  btn.disabled=true;
  document.getElementById('loading-impostos').style.display='block';
  document.getElementById('results-impostos').innerHTML='';

  var fd=new FormData();
  FILES['pdf-impostos'].forEach(function(f){fd.append('pdf',f)});

  try{
    var res=await fetch('/auditoria-impostos',{method:'POST',body:fd});
    var data=await res.json();
    AUDIT_STATE.impostos = data;
    renderImpostos(data);
  } catch(e){
    document.getElementById('results-impostos').innerHTML='<div class="err-box"><h4>Erro</h4><p>'+e.message+'</p></div>';
  } finally{
    btn.disabled=false;
    document.getElementById('loading-impostos').style.display='none';
  }
}

function renderImpostos(data){
  var el=document.getElementById('results-impostos');
  if(data.error){el.innerHTML='<div class="err-box"><h4>Erro</h4><p>'+data.error+'</p></div>';return}

  var r=data.resumo;

  // ── Cards de resumo ──────────────────────────────────────────────────────
  var html='<div class="stats" style="grid-template-columns:repeat(4,1fr)">'
    +'<div class="stat t"><div class="n">'+r.total+'</div><div class="l">Colaboradores</div></div>'
    +'<div class="stat"><div class="n" style="color:#10b981">'+r.ok+'</div><div class="l">OK</div></div>'
    +'<div class="stat"><div class="n" style="color:#f59e0b">'+r.com_divergencia+'</div><div class="l">Com divergência</div></div>'
    +'<div class="stat"><div class="n" style="color:#A72C31">'+r.total_criticos+'</div><div class="l">Críticos</div></div>'
    +'</div>';

  // ── Tabela principal ─────────────────────────────────────────────────────
  html+='<div class="card"><div class="sec-title">Auditoria INSS / IRRF — por Colaborador</div>'
    +'<div style="overflow-x:auto">'
    +'<table class="tbl-comp"><thead>'
    +'<tr>'
    +'<th rowspan="2" style="min-width:160px">Colaborador</th>'
    +'<th rowspan="2" style="text-align:right;width:100px">Sal. Bruto</th>'
    +'<th colspan="4" style="text-align:center;border-left:2px solid #e5e7eb;background:#fef9f9">INSS</th>'
    +'<th colspan="4" style="text-align:center;border-left:2px solid #e5e7eb;background:#f9f9ff">IRRF</th>'
    +'</tr>'
    +'<tr>'
    +'<th style="text-align:right;border-left:2px solid #e5e7eb;font-weight:500;color:#6b7280;background:#fef9f9">Calculado</th>'
    +'<th style="text-align:right;font-weight:500;color:#6b7280;background:#fef9f9">Encontrado</th>'
    +'<th style="text-align:right;font-weight:600;background:#fef9f9">Δ</th>'
    +'<th style="text-align:center;background:#fef9f9">Status</th>'
    +'<th style="text-align:right;border-left:2px solid #e5e7eb;font-weight:500;color:#6b7280;background:#f9f9ff">Calculado</th>'
    +'<th style="text-align:right;font-weight:500;color:#6b7280;background:#f9f9ff">Encontrado</th>'
    +'<th style="text-align:right;font-weight:600;background:#f9f9ff">Δ</th>'
    +'<th style="text-align:center;background:#f9f9ff">Status</th>'
    +'</tr>'
    +'</thead><tbody>';

  function statusBadge(st){
    if(st==='OK')        return '<span class="imp-badge imp-ok">OK</span>';
    if(st==='AUSENTE')   return '<span class="imp-badge imp-ausente">AUSENTE</span>';
    if(st==='DIVERGENTE')return '<span class="imp-badge imp-div">DIVERGENTE</span>';
    if(st==='ARREDONDAMENTO') return '<span class="imp-badge imp-arr">ARRED.</span>';
    return '<span class="imp-badge imp-nd">SEM DADOS</span>';
  }

  function deltaImposto(calc, enc, status){
    if(status==='SEM_DADOS'||status==='OK'||status==='ARREDONDAMENTO'){
      return status==='OK'?'<span style="color:#10b981;font-weight:600">—</span>':'<span style="color:#9ca3af">—</span>';
    }
    var diff = Math.abs(calc - enc);
    var cls  = calc > enc ? 'pct-up' : 'pct-down';
    var lbl  = calc > enc ? '▲ falta' : '▼ excesso';
    return '<span class="'+cls+'" style="font-size:.78rem;font-weight:700">'+lbl+' '+brl(diff)+'</span>';
  }

  data.colaboradores.forEach(function(c){
    var temDiv = c.divergencias && c.divergencias.length > 0;
    var rowCls = temDiv
      ? (c.divergencias.some(function(d){return d.criticidade==='alta'}) ? 'comp-row-alta' : 'comp-row-media')
      : (c.inss_status==='SEM_DADOS'&&c.irrf_status==='SEM_DADOS' ? '' : 'comp-row-ok');

    html+='<tr class="'+rowCls+'">'
      +'<td style="font-weight:600;font-size:.83rem">'+c.nome+'</td>'
      +'<td style="text-align:right">'+brl(c.salario_bruto)+'</td>'
      // INSS
      +'<td style="text-align:right;border-left:2px solid #e5e7eb">'+brl(c.inss_calculado)+'</td>'
      +'<td style="text-align:right">'+brl(c.inss_encontrado)+'</td>'
      +'<td style="text-align:right">'+deltaImposto(c.inss_calculado, c.inss_encontrado, c.inss_status)+'</td>'
      +'<td style="text-align:center">'+statusBadge(c.inss_status)+'</td>'
      // IRRF
      +'<td style="text-align:right;border-left:2px solid #e5e7eb">'+brl(c.irrf_calculado)+'</td>'
      +'<td style="text-align:right">'+brl(c.irrf_encontrado)+'</td>'
      +'<td style="text-align:right">'+deltaImposto(c.irrf_calculado, c.irrf_encontrado, c.irrf_status)+'</td>'
      +'<td style="text-align:center">'+statusBadge(c.irrf_status)+'</td>'
      +'</tr>';
  });

  html+='</tbody></table></div>'
    +'<p style="font-size:.72rem;color:#9ca3af;margin-top:.6rem">'
    +'▲ falta = calculado maior que encontrado &nbsp;·&nbsp; ▼ excesso = encontrado maior que calculado &nbsp;·&nbsp; '
    +'IRRF: tolerância de R$ 10,00 (dedução de dependentes/pensão não computados)'
    +'</p></div>';

  // ── Tabela INSS de referência ─────────────────────────────────────────────
  if(data.tabela_inss){
    html+='<div class="card"><details><summary style="font-size:.82rem;font-weight:600;color:#374151;cursor:pointer;padding:.2rem 0">Tabela INSS 2024/2025 — referência</summary>'
      +'<table class="tbl-sigma" style="margin-top:.75rem;max-width:320px"><thead><tr><th>Faixa salarial</th><th>Alíquota</th></tr></thead><tbody>';
    data.tabela_inss.forEach(function(f){
      html+='<tr><td>'+f.faixa+'</td><td style="font-weight:600;color:#A72C31">'+f.aliquota+'</td></tr>';
    });
    html+='</tbody></table></details></div>';
  }

  if(data.erros&&data.erros.length){
    html+='<div class="err-box"><h4>Avisos</h4>';
    data.erros.forEach(function(e){html+='<p>'+e+'</p>'});
    html+='</div>';
  }

  html+='<div class="export-bar">'
    +'<button class="btn-export" onclick="exportar(\'impostos\',\'excel\')">Exportar Excel</button>'
    +'<button class="btn-export" onclick="exportar(\'impostos\',\'csv\')">Exportar CSV</button>'
    +'</div>';

  el.innerHTML=html;
  el.scrollIntoView({behavior:'smooth'});
}

/* ═══════════════════════════════════════════
   ABA 4: BENEFICIOS
   ═══════════════════════════════════════════ */

async function analyzeBeneficio(){
  if(!FILES.extrato.length){alert('Envie o Extrato de Folha PDF.');return}
  var btn=document.getElementById('btn-beneficio');
  btn.disabled=true;
  document.getElementById('loading-beneficio').style.display='block';
  document.getElementById('results-beneficio').innerHTML='';

  var fd=new FormData();
  FILES.fatura.forEach(function(f){fd.append('fatura',f)});
  FILES.extrato.forEach(function(f){fd.append('extrato',f)});
  fd.append('regra_tipo', document.getElementById('regra-tipo').value);
  fd.append('regra_valor', document.getElementById('regra-valor').value||'0');
  fd.append('evento_codigo', getEventoCodigo());
  fd.append('filtro_linha', (document.getElementById('filtro-linha').value||'MENSALIDADE').toUpperCase());

  try{
    var res=await fetch('/comparar-beneficio',{method:'POST',body:fd});
    var data=await res.json();
    AUDIT_STATE.beneficio = data;
    renderBeneficio(data);
  } catch(e){
    document.getElementById('results-beneficio').innerHTML='<div class="err-box"><h4>Erro</h4><p>'+e.message+'</p></div>';
  } finally{
    btn.disabled=false;
    document.getElementById('loading-beneficio').style.display='none';
  }
}

function renderBeneficio(data){
  var el=document.getElementById('results-beneficio');
  if(data.error){el.innerHTML='<div class="err-box"><h4>Erro</h4><p>'+data.error+'</p></div>';return}

  var difClass = function(v){return v > 0.05 ? 'diff-pos' : (v < -0.05 ? 'diff-neg' : 'diff-ok')};
  var difSign  = function(v){return v > 0.05 ? '+' : ''};
  var fmtDif   = function(v){return '<span class="'+difClass(v)+'">'+difSign(v)+brl(v)+'</span>'};

  var regra = data.regra || {};
  var regraDesc = {
    fatura: 'Valor do documento de referencia',
    pct_fatura: regra.valor+'% do valor do documento',
    pct_salario: regra.valor+'% do salario',
    fixo: 'Valor fixo '+brl(regra.valor)
  }[regra.tipo] || '';

  var difAbs = Math.abs(data.total_diferenca);
  var difLabel = data.total_diferenca > 0.05 ? 'A Descontar' : (data.total_diferenca < -0.05 ? 'A Devolver' : 'Tudo OK');
  var difColor = data.total_diferenca > 0.05 ? '#A72C31' : (data.total_diferenca < -0.05 ? '#f59e0b' : '#10b981');
  var okCount  = data.total - data.divergentes;

  var html='<div class="stats"><div class="stat t"><div class="n">'+data.total+'</div><div class="l">Funcionarios</div></div>'
    +'<div class="stat" style="border-top:3px solid '+difColor+'"><div class="n" style="color:'+difColor+'">'+difLabel+'</div><div class="l">'+(difAbs>0.05?brl(difAbs):'--')+'</div></div>'
    +'<div class="stat d"><div class="n">'+data.divergentes+'</div><div class="l">Divergentes</div></div></div>';

  if(regraDesc) html+='<div style="background:#fff7ed;border-radius:8px;padding:.5rem 1rem;font-size:.78rem;color:#92400e;margin-bottom:.8rem;border-left:3px solid #A72C31">Regra: <strong>'+regraDesc+'</strong> | Evento: <strong>'+(data.evento_codigo||'8111')+'</strong></div>';

  // Tabela
  html+='<div class="card"><div class="sec-title">Resultado por funcionario</div>'
    +'<div style="overflow-x:auto"><table class="ben-table"><thead><tr>'
    +'<th>Funcionario</th><th style="text-align:right">'+(regraDesc||'Esperado')+'</th>'
    +'<th style="text-align:right">Descontado</th><th style="text-align:right">Diferenca</th>'
    +'<th style="text-align:center">Status</th></tr></thead><tbody>';

  data.resultados.forEach(function(r){
    var rowCls = r.sem_extrato?'sem-doc':(r.status==='MAIOR'?'maior':(r.status==='MENOR'?'menor':''));
    var badge = r.sem_extrato ? '<span class="ben-badge nd">S/ DOC</span>'
      : r.status==='OK'&&!r.sem_extrato ? '<span class="ben-badge ok">OK</span>'
      : r.status==='MAIOR' ? '<span class="ben-badge maior">A DESCONTAR</span>'
      : '<span class="ben-badge menor">A DEVOLVER</span>';
    html+='<tr class="'+rowCls+'"><td style="font-weight:500">'+r.nome+'</td>'
      +'<td class="valor">'+(r.valor_esperado?brl(r.valor_esperado):'--')+'</td>'
      +'<td class="valor">'+(r.valor_descontado?brl(r.valor_descontado):'--')+'</td>'
      +'<td class="valor">'+fmtDif(r.diferenca)+'</td>'
      +'<td style="text-align:center">'+badge+'</td></tr>';
  });

  html+='</tbody><tfoot><tr style="background:#f1f3f6;font-weight:700;font-size:.85rem">'
    +'<td style="padding:.55rem .7rem;border-top:2px solid #e5e7eb">TOTAL GERAL</td>'
    +'<td style="text-align:right;padding:.55rem .7rem;border-top:2px solid #e5e7eb">'+brl(data.total_esperado)+'</td>'
    +'<td style="text-align:right;padding:.55rem .7rem;border-top:2px solid #e5e7eb">'+brl(data.total_extrato)+'</td>'
    +'<td style="text-align:right;padding:.55rem .7rem;border-top:2px solid #e5e7eb;color:'+(data.total_diferenca>0.05?'#A72C31':data.total_diferenca<-0.05?'#f59e0b':'#10b981')+'">'
    +(Math.abs(data.total_diferenca)>0.05?brl(data.total_diferenca):'--')+'</td>'
    +'<td></td></tr></tfoot></table></div></div>';

  if(data.erros&&data.erros.length){
    html+='<div class="err-box"><h4>Avisos</h4>';
    data.erros.forEach(function(e){html+='<p>'+e+'</p>'});
    html+='</div>';
  }

  html+='<div class="export-bar">'
    +'<button class="btn-export" onclick="exportar(\'beneficio\',\'excel\')">Exportar Excel</button>'
    +'<button class="btn-export" onclick="exportar(\'beneficio\',\'csv\')">Exportar CSV</button>'
    +'<button class="btn-export" onclick="window.print()">Imprimir</button>'
    +'</div>';

  el.innerHTML=html;
  el.scrollIntoView({behavior:'smooth'});
}

/* ═══════════════════════════════════════════
   ABA 5: AUDITORIA CONSOLIDADA
   ═══════════════════════════════════════════ */

function renderAuditoriaConsolidada(){
  var el=document.getElementById('results-auditoria');
  var hasData = AUDIT_STATE.folha || AUDIT_STATE.beneficio || AUDIT_STATE.mesAnterior || AUDIT_STATE.impostos;
  if(!hasData){
    el.innerHTML='<div class="card"><p style="text-align:center;color:#9ca3af;padding:2rem">Nenhuma analise realizada ainda. Use as outras abas para processar arquivos e o resumo aparecera aqui.</p></div>';
    return;
  }

  var html='<div class="stats-4">';
  var totalDivs=0, totalOk=0, totalFunc=0, totalCrit=0;

  if(AUDIT_STATE.folha){
    var f=AUDIT_STATE.folha.resumo;
    totalDivs+=f.divergencias; totalOk+=f.ok; totalFunc+=f.total;
  }
  if(AUDIT_STATE.beneficio){
    totalDivs+=AUDIT_STATE.beneficio.divergentes;
    totalOk+=(AUDIT_STATE.beneficio.total-AUDIT_STATE.beneficio.divergentes);
    totalFunc+=AUDIT_STATE.beneficio.total;
  }
  if(AUDIT_STATE.mesAnterior){
    var m=AUDIT_STATE.mesAnterior.resumo;
    totalCrit+=m.total_criticos;
  }
  if(AUDIT_STATE.impostos){
    var im=AUDIT_STATE.impostos.resumo;
    totalDivs+=im.com_divergencia; totalOk+=im.ok; totalCrit+=im.total_criticos;
  }

  html+='<div class="stat"><div class="n" style="color:#333">'+totalFunc+'</div><div class="l">Total Analisados</div></div>'
    +'<div class="stat"><div class="n" style="color:#10b981">'+totalOk+'</div><div class="l">OK</div></div>'
    +'<div class="stat"><div class="n" style="color:#f59e0b">'+totalDivs+'</div><div class="l">Divergencias</div></div>'
    +'<div class="stat"><div class="n" style="color:#A72C31">'+totalCrit+'</div><div class="l">Criticos</div></div>'
    +'</div>';

  // Detalhamento por modulo
  if(AUDIT_STATE.folha){
    var rf=AUDIT_STATE.folha.resumo;
    html+='<div class="card"><div class="sec-title">Folha x Lancamentos</div>'
      +'<p style="font-size:.82rem;color:#6b7280">'+rf.total+' funcionarios | '+rf.ok+' OK | '+rf.divergencias+' divergencias</p></div>';
  }
  if(AUDIT_STATE.mesAnterior){
    var rm=AUDIT_STATE.mesAnterior.resumo;
    html+='<div class="card"><div class="sec-title">Mes Anterior</div>'
      +'<p style="font-size:.82rem;color:#6b7280">'+rm.novos+' novos | '+rm.desligados+' desligados | '+rm.alterados+' alterados | '+rm.total_criticos+' criticos</p></div>';
  }
  if(AUDIT_STATE.impostos){
    var ri=AUDIT_STATE.impostos.resumo;
    html+='<div class="card"><div class="sec-title">INSS / IRRF</div>'
      +'<p style="font-size:.82rem;color:#6b7280">'+ri.total+' colaboradores | '+ri.ok+' OK | '+ri.com_divergencia+' divergencias</p></div>';
  }
  if(AUDIT_STATE.beneficio){
    html+='<div class="card"><div class="sec-title">Beneficios</div>'
      +'<p style="font-size:.82rem;color:#6b7280">'+AUDIT_STATE.beneficio.total+' funcionarios | '+(AUDIT_STATE.beneficio.total-AUDIT_STATE.beneficio.divergentes)+' OK | '+AUDIT_STATE.beneficio.divergentes+' divergentes'
      +' | Diferenca total: '+brl(AUDIT_STATE.beneficio.total_diferenca)+'</p></div>';
  }

  el.innerHTML=html;
}

/* ═══════════════════════════════════════════
   ABA 6: CONFIGURACOES
   ═══════════════════════════════════════════ */

async function recarregarRubricas(){
  try{
    var res=await fetch('/recarregar-rubricas',{method:'POST'});
    var data=await res.json();
    if(data.ok){
      alert('Config recarregada! '+data.grupos.length+' grupos, '+data.total_variantes+' variantes carregadas.');
    } else {
      alert('Erro ao recarregar: '+data.erro);
    }
  } catch(e){
    alert('Erro: '+e.message);
  }
}

/* ═══════════════════════════════════════════
   EXPORTACAO
   ═══════════════════════════════════════════ */

async function exportar(tipo, formato){
  var dados = null;
  if(tipo==='folha') dados = AUDIT_STATE.folha;
  else if(tipo==='beneficio') dados = AUDIT_STATE.beneficio;
  else if(tipo==='mes_anterior') dados = AUDIT_STATE.mesAnterior;
  else if(tipo==='impostos') dados = AUDIT_STATE.impostos;

  if(!dados){alert('Nenhum dado disponivel para exportar. Execute a analise primeiro.');return}

  try{
    var res = await fetch('/exportar', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({tipo: tipo, formato: formato, dados: dados})
    });
    if(!res.ok) throw new Error('Erro ao exportar');
    var blob = await res.blob();
    var ext = formato==='excel'?'.xlsx':'.csv';
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'conferencia_'+tipo+ext;
    a.click();
    URL.revokeObjectURL(url);
  } catch(e){
    alert('Erro na exportacao: '+e.message);
  }
}

/* ═══════════════════════════════════════════
   UTILIDADES
   ═══════════════════════════════════════════ */

function tog(el){el.classList.toggle('open')}

function toggleForm(e, btn){
  e.stopPropagation();
  var form = btn.nextElementSibling;
  var isOpen = form.classList.contains('open');
  form.classList.toggle('open', !isOpen);
  btn.textContent = isOpen ? '+ Apontar divergencia manual' : '- Cancelar';
}

function calcDiff(inp){
  var form = inp.closest('.div-form');
  var esp = parseFloat(form.querySelector('.f-esp').value) || 0;
  var enc = parseFloat(form.querySelector('.f-enc').value) || 0;
  var diff = form.querySelector('.f-diff');
  if(esp || enc){
    var d = Math.abs(esp - enc);
    diff.textContent = 'R$ ' + d.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
    diff.style.color = d > 0 ? '#dc2626' : '#059669';
  } else {
    diff.textContent = '--';
    diff.style.color = '';
  }
}

function saveDiv(btn, nomeFunc){
  var form = btn.closest('.div-form');
  var tipo = form.querySelector('.f-tipo').value.trim();
  var desc = form.querySelector('.f-desc').value.trim();
  var esp  = parseFloat(form.querySelector('.f-esp').value) || 0;
  var enc  = parseFloat(form.querySelector('.f-enc').value) || 0;

  if(!tipo && !desc){ alert('Informe ao menos o tipo ou a descricao da divergencia.'); return; }

  var descFull = desc || '';
  if(esp || enc){
    var fmtBrl = function(v){return 'R$ ' + v.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})};
    var parts = [];
    if(esp) parts.push('Esperado: '+fmtBrl(esp));
    if(enc) parts.push('Encontrado: '+fmtBrl(enc));
    if(esp && enc) parts.push('Diferenca: '+fmtBrl(Math.abs(esp-enc)));
    descFull += (descFull ? ' | ' : '') + parts.join(' | ');
  }

  var divItem = document.createElement('div');
  divItem.className = 'div-item manual';
  divItem.innerHTML = '<div><div class="div-tipo">'+(tipo || 'Divergencia manual')+' <span class="manual-tag">manual</span></div>'
    +(descFull ? '<div class="div-desc">'+descFull+'</div>' : '')+'</div>';

  var addBtn = form.previousElementSibling;
  addBtn.parentNode.insertBefore(divItem, addBtn);

  var empCard = form.closest('.emp');
  empCard.classList.add('open');
  var badge = empCard.querySelector('.badge');
  var currentDivs = empCard.querySelectorAll('.div-item').length;
  badge.className = 'badge div';
  badge.innerHTML = currentDivs+' divergencia'+(currentDivs>1?'s':'');

  form.querySelector('.f-tipo').value = '';
  form.querySelector('.f-desc').value = '';
  form.querySelector('.f-esp').value = '';
  form.querySelector('.f-enc').value = '';
  form.querySelector('.f-diff').textContent = '--';
  form.classList.remove('open');
  addBtn.textContent = '+ Apontar divergencia manual';
}
