/* ── Sigma Contabilidade — Conferencia de Folha — JS ── */

const FILES = {excel:[], pdf:[], word:[], fatura:[], extrato:[], anterior:[]};

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
  updateAllFileStatuses();
  updateDocSummary();
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
  var inp=document.getElementById('file-'+t);
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

  // Auditoria INSS/IRRF inline (quando PDF foi enviado)
  if(data.auditoria_impostos && !data.auditoria_impostos.erro){
    html += renderAuditoriaImpostosInline(data.auditoria_impostos);
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
  if(!FILES.pdf.length || !FILES.anterior.length){alert('Envie os Recibos PDF (folha atual) e os Recibos PDF do Mes Anterior na aba Documentos.');return}
  var btn=document.getElementById('btn-anterior');
  btn.disabled=true;
  document.getElementById('loading-anterior').style.display='block';
  document.getElementById('results-anterior').innerHTML='';

  var fd=new FormData();
  FILES.pdf.forEach(function(f){fd.append('folha_atual',f)});
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

/* ── helpers de célula reutilizáveis ── */
function _fieldCells(ant, atu, diff, pct, borderLeft){
  var diffCls = diff>0?'pct-up':diff<0?'pct-down':'';
  var arrow   = diff>0?'▲':diff<0?'▼':'';
  var pctStr  = Math.abs(pct)>=1?' <small style="font-size:.68rem">('+pct+'%)</small>':'';
  var diffFmt = Math.abs(diff)<0.06
    ? '<span style="color:#9ca3af">—</span>'
    : '<span class="'+diffCls+'">'+arrow+' '+brl(Math.abs(diff))+pctStr+'</span>';
  var bl = borderLeft?';border-left:2px solid #e5e7eb':'';
  return '<td style="text-align:right'+bl+'">'+brl(ant)+'</td>'
    +'<td style="text-align:right">'+brl(atu)+'</td>'
    +'<td style="text-align:right">'+diffFmt+'</td>';
}

function _badgeComp(criticidade){
  var map = {alta:['badge-alta','CRÍTICO'],media:['badge-media','ATENÇÃO'],baixa:['badge-baixa','BAIXO'],ok:['badge ok','OK']};
  var b = map[criticidade]||map.ok;
  return '<span class="'+b[0]+'" style="font-size:.68rem">'+b[1]+'</span>';
}

/* ── sub-tabela de rubricas do colaborador ── */
function _rubricas_subtable(rubricas){
  if(!rubricas||!rubricas.length) return '<p style="font-size:.78rem;color:#9ca3af;padding:.5rem">Nenhuma rubrica disponível.</p>';
  var divergentes = rubricas.filter(function(r){return r.status!=='ok'});
  var iguais      = rubricas.filter(function(r){return r.status==='ok'});

  var colStyle = 'border-collapse:collapse;width:100%;font-size:.78rem';
  var thStyle  = 'text-align:{align};padding:.3rem .5rem;font-size:.75rem;color:#6b7280;font-weight:500';

  function rbRow(rb){
    var st = rb.status;
    var bg = st==='novo'?'background:#f0fdf4':st==='removido'?'background:#fff1f2':st==='alterado'?'background:#fffbeb':'';
    var badge = st==='novo'
      ? '<span class="rb-badge rb-novo">NOVA</span>'
      : st==='removido'
      ? '<span class="rb-badge rb-removido">REMOVIDA</span>'
      : st==='alterado'
      ? '<span class="rb-badge rb-alterado">ALTERADA</span>'
      : '<span class="rb-badge rb-ok">OK</span>';
    var diff = rb.diferenca||0;
    var diffFmt = Math.abs(diff)<0.06
      ? ''
      : '<span class="'+(diff>0?'pct-up':'pct-down')+'">'+(diff>0?'▲':'▼')+' '+brl(Math.abs(diff))+'</span>';
    return '<div style="display:grid;grid-template-columns:1fr auto auto auto;gap:.25rem .6rem;align-items:center;padding:.3rem .4rem;border-bottom:1px solid #f0f0f0;'+bg+'">'
      +'<span>'+rb.rubrica+'</span>'
      +'<span style="text-align:right;color:#6b7280">'+brl(rb.valor_anterior)+'</span>'
      +'<span style="text-align:right">'+brl(rb.valor_atual)+'</span>'
      +'<span>'+badge+'</span>'
      +'</div>';
  }

  var hdr = '<div style="display:grid;grid-template-columns:1fr auto auto auto;gap:.25rem .6rem;padding:.3rem .4rem;background:#f8f9fa;border-radius:4px 4px 0 0;margin-top:.5rem">'
    +'<span style="font-size:.73rem;color:#9ca3af;font-weight:500">Rubrica</span>'
    +'<span style="font-size:.73rem;color:#9ca3af;font-weight:500;text-align:right">Anterior</span>'
    +'<span style="font-size:.73rem;color:#9ca3af;font-weight:500;text-align:right">Atual</span>'
    +'<span style="font-size:.73rem;color:#9ca3af;font-weight:500">Status</span>'
    +'</div>'
    +'<div style="border:1px solid #f0f0f0;border-radius:0 0 4px 4px;overflow:hidden">';

  var body = '';
  divergentes.forEach(function(rb){body+=rbRow(rb)});

  if(iguais.length){
    body+='<details style="border-top:'+(divergentes.length?'1px solid #e5e7eb':'none')+'">'
      +'<summary style="font-size:.72rem;color:#9ca3af;cursor:pointer;padding:.35rem .5rem;list-style:none;background:#fafafa">'
      +'&#9656; '+iguais.length+' rubrica(s) sem alteração</summary>'
      +'<div>';
    iguais.forEach(function(rb){body+=rbRow(rb)});
    body+='</div></details>';
  }

  if(!divergentes.length && !iguais.length){
    body='<p style="font-size:.78rem;color:#9ca3af;padding:.5rem">Nenhuma rubrica encontrada.</p>';
  }

  return hdr+body+'</div>';
}

function renderMesAnterior(data){
  var el=document.getElementById('results-anterior');
  if(data.error){el.innerHTML='<div class="err-box"><h4>Erro</h4><p>'+data.error+'</p></div>';return}

  var r=data.resumo;

  // ── Cards de resumo ──────────────────────────────────────────────────────
  var html='<div class="stats" style="grid-template-columns:repeat(5,1fr)">'
    +'<div class="stat t"><div class="n">'+r.total_atual+'</div><div class="l">Total na folha</div></div>'
    +'<div class="stat"><div class="n" style="color:#10b981">'+r.novos+'</div><div class="l">Novos</div></div>'
    +'<div class="stat"><div class="n" style="color:#A72C31">'+r.desligados+'</div><div class="l">Desligados</div></div>'
    +'<div class="stat"><div class="n" style="color:#f59e0b">'+r.alterados+'</div><div class="l">Com alteração</div></div>'
    +'<div class="stat"><div class="n" style="color:#6b7280">'+r.sem_alteracao+'</div><div class="l">Sem alteração</div></div>'
    +'</div>';

  // ── Novos e Desligados ───────────────────────────────────────────────────
  var temNovos  = data.colaboradores_novos&&data.colaboradores_novos.length;
  var temDeslig = data.colaboradores_desligados&&data.colaboradores_desligados.length;
  if(temNovos||temDeslig){
    html+='<div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-bottom:1.25rem">';
    if(temNovos){
      html+='<div class="card" style="margin-bottom:0"><div class="sec-title" style="color:#059669">✚ Novos Colaboradores</div>';
      data.colaboradores_novos.forEach(function(n){
        html+='<div style="display:flex;justify-content:space-between;align-items:center;padding:.45rem .4rem;border-bottom:1px solid #f3f4f6">'
          +'<span style="font-weight:600;font-size:.83rem">'+n.nome+'</span>'
          +'<span class="badge-novo" style="font-size:.7rem">'+brl(n.liquido)+'</span></div>';
      });
      html+='</div>';
    } else { html+='<div></div>'; }
    if(temDeslig){
      html+='<div class="card" style="margin-bottom:0"><div class="sec-title" style="color:#A72C31">✖ Desligados</div>';
      data.colaboradores_desligados.forEach(function(d){
        html+='<div style="display:flex;justify-content:space-between;align-items:center;padding:.45rem .4rem;border-bottom:1px solid #f3f4f6">'
          +'<span style="font-weight:600;font-size:.83rem">'+d.nome+'</span>'
          +'<span class="badge-desligado" style="font-size:.7rem">'+brl(d.liquido)+'</span></div>';
      });
      html+='</div>';
    }
    html+='</div>';
  }

  // ── Tabela comparativa expansível ────────────────────────────────────────
  if(data.comparativo&&data.comparativo.length){
    html+='<div class="card"><div class="sec-title">Comparativo por Colaborador'
      +'<span style="font-size:.72rem;font-weight:400;color:#9ca3af;margin-left:.5rem">— clique em ▶ para ver rubricas detalhadas</span></div>'
      +'<div style="overflow-x:auto">'
      +'<table class="tbl-comp" id="tbl-ant"><thead><tr>'
      +'<th style="width:28px"></th>'
      +'<th style="min-width:150px">Colaborador</th>'
      +'<th style="text-align:center;width:76px">Status</th>'
      +'<th colspan="3" style="text-align:center;border-left:2px solid #e5e7eb">Líquido a Receber</th>'
      +'<th colspan="3" style="text-align:center;border-left:2px solid #e5e7eb">Total Vencimentos</th>'
      +'<th colspan="3" style="text-align:center;border-left:2px solid #e5e7eb">Total Descontos</th>'
      +'</tr><tr>'
      +'<th></th><th></th><th></th>'
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

    data.comparativo.forEach(function(c, idx){
      var rowCls = c.criticidade==='alta'?'comp-row-alta'
                 : c.criticidade==='media'?'comp-row-media'
                 : c.criticidade==='baixa'?'comp-row-baixa':'comp-row-ok';
      var ndivRb = (c.rubricas_comparadas||[]).filter(function(r){return r.status!=='ok'}).length;
      var detailId = 'det-ant-'+idx;

      html+='<tr class="'+rowCls+'" style="cursor:pointer" onclick="togDetail(\''+detailId+'\')">'
        +'<td style="text-align:center;color:#9ca3af;font-size:.8rem" id="arr-'+detailId+'">▶</td>'
        +'<td style="font-weight:600;font-size:.83rem">'+c.nome
          +(ndivRb>0?' <span style="background:#fee2e2;color:#991b1b;border-radius:10px;padding:.1rem .45rem;font-size:.65rem;font-weight:700">'+ndivRb+' dif.</span>':'')
        +'</td>'
        +'<td style="text-align:center">'+_badgeComp(c.criticidade)+'</td>'
        +_fieldCells(c.liq_ant,  c.liq_atu,  c.liq_diff,  c.liq_pct,  true)
        +_fieldCells(c.venc_ant, c.venc_atu, c.venc_diff, c.venc_pct, true)
        +_fieldCells(c.desc_ant, c.desc_atu, c.desc_diff, c.desc_pct, true)
        +'</tr>'
        +'<tr id="'+detailId+'" style="display:none"><td colspan="12" style="padding:.75rem 1rem 1rem;background:#fafbff;border-bottom:2px solid #e5e7eb">'
        +_rubricas_subtable(c.rubricas_comparadas)
        +'</td></tr>';
    });

    html+='</tbody></table></div>'
      +'<p style="font-size:.72rem;color:#9ca3af;margin-top:.6rem">▲ aumento &nbsp;▼ redução &nbsp;·&nbsp; Clique na linha para expandir o detalhamento de rubricas</p>'
      +'</div>';
  }

  // Auditoria INSS/IRRF inline da folha atual
  if(data.auditoria_impostos && !data.auditoria_impostos.erro){
    html += renderAuditoriaImpostosInline(data.auditoria_impostos);
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

function togDetail(id){
  var tr=document.getElementById(id);
  var arr=document.getElementById('arr-'+id);
  if(!tr) return;
  var open=tr.style.display!=='none';
  tr.style.display=open?'none':'table-row';
  if(arr) arr.textContent=open?'▶':'▼';
}

/* ═══════════════════════════════════════════
   ABA 3: INSS / IRRF
   ═══════════════════════════════════════════ */

async function analisarImpostos(){
  if(!FILES.pdf.length){alert('Envie os Recibos PDF na aba Documentos.');return}
  var btn=document.getElementById('btn-impostos');
  btn.disabled=true;
  document.getElementById('loading-impostos').style.display='block';
  document.getElementById('results-impostos').innerHTML='';

  var fd=new FormData();
  FILES.pdf.forEach(function(f){fd.append('pdf',f)});

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

/* ─────────────────────────────────────────────────────────────
   AUDITORIA INSS/IRRF INLINE — reutilizável em qualquer aba
   ───────────────────────────────────────────────────────────── */
function renderAuditoriaImpostosInline(audit){
  if(!audit||!audit.colaboradores) return '';

  var r = audit.resumo || {};
  var divs = r.com_divergencia || 0;
  var total = r.total || 0;
  var criticos = r.total_criticos || 0;

  // Cor do cabeçalho do details
  var cor = divs===0 ? '#166534' : (criticos>0 ? '#991b1b' : '#92400e');
  var icone = divs===0 ? '✔' : '⚠';
  var resumoTxt = divs===0
    ? total + ' colaboradores — todos OK'
    : divs + ' divergência(s) em ' + total + ' colaboradores' + (criticos?' · '+criticos+' crítica(s)':'');

  function statusBadge(st){
    if(st==='OK')         return '<span class="imp-badge imp-ok">OK</span>';
    if(st==='AUSENTE')    return '<span class="imp-badge imp-ausente">AUSENTE</span>';
    if(st==='DIVERGENTE') return '<span class="imp-badge imp-div">DIVERGENTE</span>';
    if(st==='ARREDONDAMENTO') return '<span class="imp-badge imp-arr">ARRED.</span>';
    return '<span class="imp-badge imp-nd">—</span>';
  }

  function deltaImp(calc, enc, status){
    if(status==='SEM_DADOS'||status==='OK'||status==='ARREDONDAMENTO')
      return '<span style="color:#9ca3af">—</span>';
    var diff = Math.abs(calc - enc);
    var cls  = calc > enc ? 'pct-up' : 'pct-down';
    var lbl  = calc > enc ? '▲' : '▼';
    return '<span class="'+cls+'" style="font-weight:700">'+lbl+' '+brl(diff)+'</span>';
  }

  var rows = '';
  audit.colaboradores.forEach(function(c){
    var temDiv = c.divergencias && c.divergencias.length > 0;
    var rowCls = temDiv
      ? (c.divergencias.some(function(d){return d.criticidade==='alta'}) ? 'comp-row-alta' : 'comp-row-media')
      : 'comp-row-ok';
    rows += '<tr class="'+rowCls+'">'
      +'<td style="font-weight:600;font-size:.8rem">'+c.nome+'</td>'
      +'<td style="text-align:right">'+brl(c.salario_bruto)+'</td>'
      +'<td style="text-align:right;border-left:2px solid #e5e7eb">'+brl(c.inss_calculado)+'</td>'
      +'<td style="text-align:right">'+brl(c.inss_encontrado)+'</td>'
      +'<td style="text-align:right">'+deltaImp(c.inss_calculado,c.inss_encontrado,c.inss_status)+'</td>'
      +'<td style="text-align:center">'+statusBadge(c.inss_status)+'</td>'
      +'<td style="text-align:right;border-left:2px solid #e5e7eb">'+brl(c.irrf_calculado)+'</td>'
      +'<td style="text-align:right">'+brl(c.irrf_encontrado)+'</td>'
      +'<td style="text-align:right">'+deltaImp(c.irrf_calculado,c.irrf_encontrado,c.irrf_status)+'</td>'
      +'<td style="text-align:center">'+statusBadge(c.irrf_status)+'</td>'
      +'</tr>';
  });

  return '<div class="card" style="margin-top:1.25rem">'
    +'<details'+(divs>0?' open':'')+'>'
    +'<summary style="cursor:pointer;font-size:.9rem;font-weight:700;color:'+cor+';padding:.2rem 0;list-style:none;display:flex;align-items:center;gap:.5rem">'
    +'<span style="font-size:1rem">'+icone+'</span>'
    +'<span>Auditoria INSS / IRRF</span>'
    +'<span style="font-size:.78rem;font-weight:400;color:#6b7280;margin-left:.4rem">— '+resumoTxt+'</span>'
    +'</summary>'
    +'<div style="overflow-x:auto;margin-top:.9rem">'
    +'<table class="tbl-comp"><thead>'
    +'<tr>'
    +'<th rowspan="2" style="min-width:140px">Colaborador</th>'
    +'<th rowspan="2" style="text-align:right;width:90px">Sal. Bruto</th>'
    +'<th colspan="4" style="text-align:center;border-left:2px solid #e5e7eb;background:#fef9f9">INSS</th>'
    +'<th colspan="4" style="text-align:center;border-left:2px solid #e5e7eb;background:#f9f9ff">IRRF</th>'
    +'</tr><tr>'
    +'<th style="text-align:right;border-left:2px solid #e5e7eb;font-weight:500;color:#6b7280;background:#fef9f9">Calc.</th>'
    +'<th style="text-align:right;font-weight:500;color:#6b7280;background:#fef9f9">Enc.</th>'
    +'<th style="text-align:right;font-weight:600;background:#fef9f9">Δ</th>'
    +'<th style="text-align:center;background:#fef9f9">Status</th>'
    +'<th style="text-align:right;border-left:2px solid #e5e7eb;font-weight:500;color:#6b7280;background:#f9f9ff">Calc.</th>'
    +'<th style="text-align:right;font-weight:500;color:#6b7280;background:#f9f9ff">Enc.</th>'
    +'<th style="text-align:right;font-weight:600;background:#f9f9ff">Δ</th>'
    +'<th style="text-align:center;background:#f9f9ff">Status</th>'
    +'</tr></thead><tbody>'+rows+'</tbody></table></div>'
    +'<p style="font-size:.7rem;color:#9ca3af;margin-top:.5rem">Tabela progressiva INSS 2024 · IRRF: tolerância R$ 10,00</p>'
    +'</details></div>';
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

  // ── helpers locais ───────────────────────────────────────────────────────
  function statusBadge(st){
    if(st==='OK')             return '<span class="imp-badge imp-ok">OK</span>';
    if(st==='AUSENTE')        return '<span class="imp-badge imp-ausente">AUSENTE</span>';
    if(st==='DIVERGENTE')     return '<span class="imp-badge imp-div">DIVERGENTE</span>';
    if(st==='ARREDONDAMENTO') return '<span class="imp-badge imp-arr">ARRED.</span>';
    return '<span class="imp-badge imp-nd">SEM DADOS</span>';
  }

  function deltaImposto(calc, enc, status){
    if(status==='SEM_DADOS'||status==='OK'||status==='ARREDONDAMENTO')
      return status==='OK'?'<span style="color:#10b981;font-weight:600">—</span>':'<span style="color:#9ca3af">—</span>';
    var diff=Math.abs(calc-enc);
    var cls=calc>enc?'pct-up':'pct-down';
    var lbl=calc>enc?'▲ falta':'▼ excesso';
    return '<span class="'+cls+'" style="font-size:.78rem;font-weight:700">'+lbl+' '+brl(diff)+'</span>';
  }

  function verbas_detail_imp(c){
    var rubricas = c.rubricas_detalhadas||[];

    var tipoCor   = {inss:'background:#fff0f0',irrf:'background:#eff6ff',salario:'background:#f0fdf4',outro:''};
    var tipoBadge = {
      inss:   '<span class="rb-badge" style="background:#fee2e2;color:#991b1b">INSS</span>',
      irrf:   '<span class="rb-badge" style="background:#eff6ff;color:#1d4ed8">IRRF</span>',
      salario:'<span class="rb-badge" style="background:#f0fdf4;color:#166534">SALÁRIO</span>',
      outro:  '<span style="color:#d1d5db;font-size:.7rem">—</span>',
    };

    var divINSS = c.inss_status!=='OK'&&c.inss_status!=='SEM_DADOS'&&c.inss_status!=='ARREDONDAMENTO';
    var divIRRF = c.irrf_status!=='OK'&&c.irrf_status!=='SEM_DADOS'&&c.irrf_status!=='ARREDONDAMENTO';

    // Cabeçalho via grid (sem tabela aninhada)
    var html='<div style="margin-top:.5rem">'
      +'<div style="display:grid;grid-template-columns:1fr 90px 70px;gap:.25rem .5rem;padding:.3rem .5rem;background:#f8f9fa;border-radius:4px 4px 0 0">'
      +'<span style="font-size:.72rem;color:#9ca3af;font-weight:500">Rubrica do recibo</span>'
      +'<span style="font-size:.72rem;color:#9ca3af;font-weight:500;text-align:right">Valor</span>'
      +'<span style="font-size:.72rem;color:#9ca3af;font-weight:500">Tipo</span>'
      +'</div>'
      +'<div style="border:1px solid #f0f0f0;border-top:none;border-radius:0 0 4px 4px">';

    // Linhas de referência (INSS e IRRF esperados)
    html+='<div style="display:grid;grid-template-columns:1fr 90px 70px;gap:.25rem .5rem;padding:.35rem .5rem;background:#fef9c3;border-bottom:1px solid #f0f0f0">'
      +'<span style="font-size:.75rem;color:#92400e;font-weight:600">⚖ INSS esperado (tabela progressiva)</span>'
      +'<span style="font-size:.75rem;color:#92400e;font-weight:700;text-align:right">'+brl(c.inss_calculado)+'</span>'
      +'<span style="font-size:.7rem;color:#9ca3af">Base: '+brl(c.base_inss)+'</span>'
      +'</div>';
    html+='<div style="display:grid;grid-template-columns:1fr 90px 70px;gap:.25rem .5rem;padding:.35rem .5rem;background:#fef9c3;border-bottom:1px solid #e5e7eb">'
      +'<span style="font-size:.75rem;color:#92400e;font-weight:600">⚖ IRRF esperado (tolerância R$ 10)</span>'
      +'<span style="font-size:.75rem;color:#92400e;font-weight:700;text-align:right">'+brl(c.irrf_calculado)+'</span>'
      +'<span style="font-size:.7rem;color:#9ca3af">Base: '+brl(c.base_irrf)+'</span>'
      +'</div>';

    if(!rubricas.length){
      html+='<p style="font-size:.78rem;color:#9ca3af;padding:.5rem">Nenhuma rubrica encontrada no recibo.</p>';
    } else {
      rubricas.forEach(function(v){
        var bg = tipoCor[v.tipo]||'';
        var marcador = (v.tipo==='inss'&&divINSS)||(v.tipo==='irrf'&&divIRRF)
          ? ' <span style="color:#A72C31;font-weight:700">⚠</span>' : '';
        html+='<div style="display:grid;grid-template-columns:1fr 90px 70px;gap:.25rem .5rem;padding:.3rem .5rem;border-bottom:1px solid #f5f5f5;'+bg+'">'
          +'<span style="font-size:.77rem">'+(v.codigo?'<span style="color:#9ca3af;margin-right:.3rem">'+v.codigo+'</span>':'')+v.descricao+marcador+'</span>'
          +'<span style="font-size:.77rem;font-weight:600;text-align:right">'+brl(v.valor)+'</span>'
          +'<span>'+tipoBadge[v.tipo||'outro']+'</span>'
          +'</div>';
      });
    }

    html+='</div></div>';
    return html;
  }

  // ── Tabela principal expansível ──────────────────────────────────────────
  html+='<div class="card"><div class="sec-title">Auditoria INSS / IRRF — por Colaborador'
    +'<span style="font-size:.72rem;font-weight:400;color:#9ca3af;margin-left:.5rem">— clique em ▶ para ver rubricas do recibo</span></div>'
    +'<div style="overflow-x:auto">'
    +'<table class="tbl-comp"><thead>'
    +'<tr>'
    +'<th style="width:28px"></th>'
    +'<th style="min-width:150px">Colaborador</th>'
    +'<th style="text-align:right;width:90px">Sal. Bruto</th>'
    +'<th colspan="4" style="text-align:center;border-left:2px solid #e5e7eb;background:#fef9f9">INSS</th>'
    +'<th colspan="4" style="text-align:center;border-left:2px solid #e5e7eb;background:#f9f9ff">IRRF</th>'
    +'</tr><tr>'
    +'<th></th><th></th><th></th>'
    +'<th style="text-align:right;border-left:2px solid #e5e7eb;font-weight:500;color:#6b7280;background:#fef9f9">Calc.</th>'
    +'<th style="text-align:right;font-weight:500;color:#6b7280;background:#fef9f9">Enc.</th>'
    +'<th style="text-align:right;font-weight:600;background:#fef9f9">Δ</th>'
    +'<th style="text-align:center;background:#fef9f9">Status</th>'
    +'<th style="text-align:right;border-left:2px solid #e5e7eb;font-weight:500;color:#6b7280;background:#f9f9ff">Calc.</th>'
    +'<th style="text-align:right;font-weight:500;color:#6b7280;background:#f9f9ff">Enc.</th>'
    +'<th style="text-align:right;font-weight:600;background:#f9f9ff">Δ</th>'
    +'<th style="text-align:center;background:#f9f9ff">Status</th>'
    +'</tr></thead><tbody>';

  data.colaboradores.forEach(function(c, idx){
    var temDiv=c.divergencias&&c.divergencias.length>0;
    var rowCls=temDiv
      ?(c.divergencias.some(function(d){return d.criticidade==='alta'})?'comp-row-alta':'comp-row-media')
      :(c.inss_status==='SEM_DADOS'&&c.irrf_status==='SEM_DADOS'?'':'comp-row-ok');
    var detailId='det-imp-'+idx;

    html+='<tr class="'+rowCls+'" style="cursor:pointer" onclick="togDetail(\''+detailId+'\')">'
      +'<td style="text-align:center;color:#9ca3af;font-size:.8rem" id="arr-'+detailId+'">▶</td>'
      +'<td style="font-weight:600;font-size:.83rem">'+c.nome+'</td>'
      +'<td style="text-align:right">'+brl(c.salario_bruto)+'</td>'
      +'<td style="text-align:right;border-left:2px solid #e5e7eb">'+brl(c.inss_calculado)+'</td>'
      +'<td style="text-align:right">'+brl(c.inss_encontrado)+'</td>'
      +'<td style="text-align:right">'+deltaImposto(c.inss_calculado,c.inss_encontrado,c.inss_status)+'</td>'
      +'<td style="text-align:center">'+statusBadge(c.inss_status)+'</td>'
      +'<td style="text-align:right;border-left:2px solid #e5e7eb">'+brl(c.irrf_calculado)+'</td>'
      +'<td style="text-align:right">'+brl(c.irrf_encontrado)+'</td>'
      +'<td style="text-align:right">'+deltaImposto(c.irrf_calculado,c.irrf_encontrado,c.irrf_status)+'</td>'
      +'<td style="text-align:center">'+statusBadge(c.irrf_status)+'</td>'
      +'</tr>'
      +'<tr id="'+detailId+'" style="display:none"><td colspan="11" style="padding:.75rem 1rem 1rem;background:#fafbff;border-bottom:2px solid #e5e7eb">'
      +verbas_detail_imp(c)
      +'</td></tr>';
  });

  html+='</tbody></table></div>'
    +'<p style="font-size:.72rem;color:#9ca3af;margin-top:.6rem">'
    +'▲ falta · ▼ excesso &nbsp;·&nbsp; ⚠ rubrica divergente &nbsp;·&nbsp; IRRF: tolerância R$ 10,00'
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
   STATUS DE ARQUIVOS NAS ABAS DE ANALISE
   ═══════════════════════════════════════════ */

function _fstatusItem(key, label, optional){
  var count = FILES[key] ? FILES[key].length : 0;
  var has = count > 0;
  if(has){
    return '<span class="fsi fsi-ok">&#10003; '+label+' <em>('+count+')</em></span>';
  } else if(optional){
    return '<span class="fsi fsi-opt">&#9675; '+label+' <em>(opcional)</em></span>';
  } else {
    return '<span class="fsi fsi-miss">&#9888; '+label+'</span>';
  }
}

function updateAllFileStatuses(){
  var link = '<button class="fsi-link" onclick="switchTab(\'documentos\')">Ir para Documentos &#x2197;</button>';

  // Folha x Lancamentos: excel OU pdf (pelo menos um)
  var el1 = document.getElementById('file-status-folha');
  if(el1){
    el1.innerHTML = '<div class="fsi-row">'
      +_fstatusItem('pdf','Recibos PDF',false)
      +_fstatusItem('excel','Planilha Excel',false)
      +_fstatusItem('word','Documento Word',true)
      +link+'</div>'
      +(FILES.pdf.length===0&&FILES.excel.length===0
        ?'<p class="fsi-hint">Envie pelo menos os Recibos PDF ou a Planilha Excel.</p>':'');
  }

  // Mes Anterior: pdf (atual) E anterior
  var el2 = document.getElementById('file-status-anterior');
  if(el2){
    var faltaAnt = FILES.pdf.length===0||FILES.anterior.length===0;
    el2.innerHTML = '<div class="fsi-row">'
      +_fstatusItem('pdf','Recibos PDF (folha atual)',false)
      +_fstatusItem('anterior','Recibos PDF (mes anterior)',false)
      +link+'</div>'
      +(faltaAnt?'<p class="fsi-hint">Envie os dois PDFs para comparar os meses.</p>':'');
  }

  // INSS / IRRF: pdf
  var el3 = document.getElementById('file-status-impostos');
  if(el3){
    el3.innerHTML = '<div class="fsi-row">'
      +_fstatusItem('pdf','Recibos PDF',false)
      +link+'</div>'
      +(FILES.pdf.length===0?'<p class="fsi-hint">Envie os Recibos PDF para auditar INSS e IRRF.</p>':'');
  }

  // Beneficios: extrato (fatura opcional)
  var el4 = document.getElementById('file-status-beneficio');
  if(el4){
    el4.innerHTML = '<div class="fsi-row">'
      +_fstatusItem('extrato','Extrato de Folha PDF',false)
      +_fstatusItem('fatura','Fatura / Referencia',true)
      +link+'</div>'
      +(FILES.extrato.length===0?'<p class="fsi-hint">Envie o Extrato de Folha PDF.</p>':'');
  }
}

function updateDocSummary(){
  var el = document.getElementById('doc-summary');
  if(!el) return;
  var total = Object.values(FILES).reduce(function(s,a){return s+a.length},0);
  if(total===0){
    el.innerHTML='<span style="color:#9ca3af;font-size:.82rem">Nenhum arquivo carregado ainda. Arraste ou clique nas areas acima.</span>';
    return;
  }
  var parts = [];
  if(FILES.pdf.length)       parts.push('<strong>'+FILES.pdf.length+'</strong> Recibo(s) PDF');
  if(FILES.excel.length)     parts.push('<strong>'+FILES.excel.length+'</strong> Excel');
  if(FILES.word.length)      parts.push('<strong>'+FILES.word.length+'</strong> Word');
  if(FILES.anterior.length)  parts.push('<strong>'+FILES.anterior.length+'</strong> PDF (mes anterior)');
  if(FILES.fatura.length)    parts.push('<strong>'+FILES.fatura.length+'</strong> Fatura');
  if(FILES.extrato.length)   parts.push('<strong>'+FILES.extrato.length+'</strong> Extrato');
  el.innerHTML='<span class="doc-summary-ok">&#10003; Carregado: '+parts.join(' &middot; ')+'</span>'
    +'<span style="font-size:.75rem;color:#6b7280;margin-left:.8rem">Use as abas de analise para processar</span>';
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

// Inicializa status ao carregar a pagina
document.addEventListener('DOMContentLoaded', function(){
  updateAllFileStatuses();
  updateDocSummary();
});
