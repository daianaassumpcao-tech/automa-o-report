/* Garante que XLSX está disponível no navegador (CDN) e no Node (require). */
if (typeof XLSX === 'undefined' && typeof require !== 'undefined') { XLSX = require('xlsx'); }

/* ═══════════════════════════════════════════════════════════════════════
   relatorios-core.js — Hype Coworking · Facilities

   Lógica de negócio PURA: parse de planilha → dados → insights → HTML de PDF.
   Sem referências a document/window/FileReader/DOM.

   Carregamento:
     Navegador → <script src="relatorios-core.js"></script>   (vira global)
     Node       → const Core = require('./relatorios-core.js')
   ═══════════════════════════════════════════════════════════════════════ */

/* CSS dos relatórios PDF, embutido como string pra injetar no iframe de impressão
   (o iframe é um documento separado — não herda o <style> da página admin). */
const PDF_CSS = `
.pdf-root{font-family:Arial,Helvetica,sans-serif;color:#e8e9ed;font-size:13px;line-height:1.4;background:#111417;}
.pdf-page{width:794px;min-height:1123px;padding:34px 56px 28px 56px;background:#111417;display:flex;flex-direction:column;box-sizing:border-box;page-break-after:always;break-after:page;}
.pdf-page:last-child{page-break-after:auto;break-after:auto;}
.pdf-page *{box-sizing:border-box;}
.pdf-header{display:flex;justify-content:space-between;align-items:flex-start;padding-bottom:8px;margin-bottom:11px;}
.pdf-header h1{font-size:13px;font-weight:600;color:#8a8f99;text-transform:uppercase;letter-spacing:.05em;}
.pdf-header h1 .pagetag{font-weight:400;color:#5e636d;text-transform:none;letter-spacing:0;}
.pdf-header .meta{text-align:right;font-size:11px;color:#5e636d;}
.pdf-kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:9px;margin-bottom:9px;}
.pdf-kpi-grid.cols3{grid-template-columns:repeat(3,1fr);}
.pdf-kpi-grid.cols2{grid-template-columns:repeat(2,1fr);}
.pdf-kpi{border-radius:7px;padding:10px 12px;background:#1a1e23;}
.pdf-kpi .lbl{font-size:9px;text-transform:uppercase;letter-spacing:.04em;color:#7c818c;font-weight:600;margin-bottom:6px;}
.pdf-kpi .val{font-size:23px;font-weight:700;color:#eceef2;line-height:1.05;}
.pdf-kpi .val .unit2{font-size:15px;font-weight:400;color:#8a8f99;}
.pdf-kpi .val.unit{font-size:20px;}
.pdf-kpi .val.green{color:#00ffa3;}
.pdf-kpi .val.red{color:#e45252;}
.pdf-kpi .val.amber{color:#ffb24d;}
.pdf-kpi .sub{font-size:10px;color:#8a8f99;margin-top:6px;}
.pdf-kpi .sub.dot{padding-left:15px;position:relative;}
.pdf-kpi .sub.dot::before{content:'';width:7px;height:7px;border-radius:50%;position:absolute;left:0;top:3px;}
.pdf-kpi .sub.dot.good{color:#00e08f;}
.pdf-kpi .sub.dot.good::before{background:#00e08f;}
.pdf-kpi .sub.dot.warn{color:#e45252;}
.pdf-kpi .sub.dot.warn::before{background:#e45252;}
.pdf-kpi .sub.dot.amber{color:#ffb24d;}
.pdf-kpi .sub.dot.amber::before{background:#ffb24d;}
.pdf-section-title{font-size:10px;font-weight:600;color:#7c818c;text-transform:uppercase;letter-spacing:.04em;margin:9px 0 6px;}
.pdf-table{width:100%;border-collapse:collapse;font-size:10px;margin-bottom:8px;background:#1a1e23;border-radius:7px;overflow:hidden;}
.pdf-table th{background:#22262c;color:#8a8f99;padding:5px 10px;text-align:left;font-weight:600;font-size:9px;text-transform:uppercase;letter-spacing:.03em;}
.pdf-table td{padding:5px 10px;border-bottom:1px solid #22262c;color:#d5d7dc;}
.pdf-table tr:last-child td{font-weight:700;background:#22262c;color:#eceef2;}
.pdf-table tr:last-child td.pdf-status-ok{color:#00e08f;}
.pdf-table tr:last-child td.pdf-status-warn{color:#e45252;}
.pdf-status-ok{color:#00e08f!important;}
.pdf-status-warn{color:#e45252!important;}
.pdf-progress{background:#1a1e23;border-radius:7px;padding:11px 16px;margin-bottom:8px;}
.pdf-prow{display:flex;align-items:center;gap:12px;padding:6px 0;font-size:10px;}
.pdf-prow .plabel{width:175px;color:#aeb2bb;flex-shrink:0;}
.pdf-ptrack{flex:1;height:11px;background:#22262c;border-radius:5px;overflow:hidden;position:relative;}
.pdf-pfill{height:100%;border-radius:5px;}
.pdf-pval{width:120px;text-align:right;color:#eceef2;font-weight:600;flex-shrink:0;}
.pdf-status-areas{margin-top:24px;}
.pdf-highlight{border:1px solid #1f4d36;background:#0f2318;border-radius:8px;padding:10px 15px;margin-bottom:12px;font-size:11px;font-weight:600;color:#5fe3ab;}
.pdf-status-block{display:flex;gap:12px;margin-bottom:10px;}
.pdf-status-block .sdot{width:8px;height:8px;border-radius:50%;background:#5e636d;margin-top:6px;flex-shrink:0;}
.pdf-status-block .st-title{font-size:11px;font-weight:700;color:#eceef2;margin-bottom:3px;}
.pdf-status-block .st-text{font-size:10.5px;color:#aeb2bb;line-height:1.45;}
.pdf-footnote{background:#2d1515;border:1px solid #5a2424;border-radius:8px;padding:10px 15px;font-size:10px;color:#f08a8a;margin-top:11px;line-height:1.4;}
.pdf-footer-bar{display:flex;justify-content:space-between;font-size:9.5px;color:#4a4e57;margin-top:auto;padding-top:14px;}
.pdf-list{background:#1a1e23;border-radius:7px;padding:4px 16px;margin-bottom:8px;}
.pdf-list .lrow{display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid #22262c;font-size:10px;color:#aeb2bb;}
.pdf-list .lrow:last-child{border-bottom:none;}
.pdf-list .lrow span:last-child{color:#eceef2;}
.pdf-list .lrow .amber{color:#ffb24d!important;font-weight:600;}
@media print {
  @page { size: 794px 1123px; margin: 0; }
  html, body { margin:0; padding:0; background:#111417; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
}
`;

function norm(s){
  return String(s==null?'':s).normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase().replace(/[\s_-]+/g,'');
}
// acha o nome real da aba mesmo que esteja com acento/caixa/espaço diferente
function resolveSheet(wb, wanted){
  if(wb.Sheets[wanted]) return wanted;
  const alvo = norm(wanted);
  return wb.SheetNames.find(n => norm(n) === alvo) || null;
}
// pega o valor de uma coluna pelo nome, ignorando acento/caixa/espaço
function pick(row, wanted){
  const alvo = norm(wanted);
  const k = Object.keys(row).find(kk => norm(kk) === alvo);
  return k!==undefined ? row[k] : undefined;
}

function sheetToObjects(wb, name){
  const real = resolveSheet(wb, name);
  if(!real) throw new Error(`Aba "${name}" não encontrada. Abas que o sistema enxergou neste arquivo: [ ${wb.SheetNames.join('  |  ')} ]`);
  return XLSX.utils.sheet_to_json(wb.Sheets[real],{defval:null,raw:false});
}
// versões OPCIONAIS: se a aba não existir, devolve [] / {} em vez de travar.
// Usadas para abas novas (Desfalque, Orcamento, Status_Areas) para que
// planilhas antigas — sem essas abas — continuem funcionando normalmente.
function sheetToObjectsOpt(wb, name){
  const real = resolveSheet(wb, name);
  if(!real) return [];
  return XLSX.utils.sheet_to_json(wb.Sheets[real],{defval:null,raw:false});
}
function kvSheetToObjOpt(wb, name){
  const real = resolveSheet(wb, name);
  if(!real) return {};
  return kvSheetToObj(wb, real);
}

function kvSheetToObj(wb, name){
  const rows = sheetToObjects(wb, name);
  const o = {};
  rows.forEach(r=>{
    const kv = pick(r,'campo');
    const k = (kv==null?'':String(kv)).trim();
    if(k) o[k] = pick(r,'valor');
  });
  return o;
}

// Detecta o separador decimal pela POSIÇÃO, não pelo símbolo:
// o separador mais à direita é o decimal (nunca tem mais de 2 dígitos depois dele
// num valor monetário/percentual); o(s) outro(s) é(são) separador de milhar.
// Isso cobre os dois formatos que o SheetJS pode produzir:
//   - número Excel com formatação #,##0.00 -> "5,566.67" (vírgula=milhar, ponto=decimal)
//   - texto digitado em formato BR -> "1.234,56" (ponto=milhar, vírgula=decimal)
//   - número Excel simples, sem formatação -> "781.44" ou "66700" (ponto=decimal ou nenhum separador)
function num(v){
  if(v==null||v==='') return 0;
  if(typeof v === 'number') return v;
  let s = String(v).trim();
  const lastComma = s.lastIndexOf(',');
  const lastDot = s.lastIndexOf('.');
  if(lastComma===-1 && lastDot===-1) return parseFloat(s)||0;
  if(lastComma > lastDot){
    // vírgula é o separador mais à direita -> decimal BR (ex: "1.234,56")
    s = s.replace(/\./g,'').replace(',','.');
  } else {
    // ponto é o separador mais à direita -> decimal US/SheetJS (ex: "5,566.67" ou "781.44")
    s = s.replace(/,/g,'');
  }
  return parseFloat(s)||0;
}

/* ═══════════════════════════════════════════════════════════════
   MOTOR DE TEXTO — gera rascunhos de "Status das Áreas" (Fase 2)
   ═══════════════════════════════════════════════════════════════ */
/* ═══════════════════════════════════════════════════════════════
   MOTOR DE TEXTO — gera rascunhos de "Status das Áreas"
   a partir dos números já calculados (dados.operacional/comercial/
   desfalque/orcamento). Espelha o tom dos relatórios PDF mensais:
   título curto + parágrafo direto com números e comparação à meta.
   ═══════════════════════════════════════════════════════════════ */

function fmtBRL(v){
  return 'R$' + v.toLocaleString('pt-BR', {minimumFractionDigits:0, maximumFractionDigits:0});
}
function fmtBRLk(v){
  // formato compacto: R$25,6k (como nos PDFs)
  if(Math.abs(v) >= 1000){
    return 'R$' + (v/1000).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1}) + 'k';
  }
  return fmtBRL(v);
}
function pct(v, dec=0){
  return v.toLocaleString('pt-BR',{minimumFractionDigits:dec,maximumFractionDigits:dec}) + '%';
}
function gapTxt(atual, meta){
  const gap = atual - meta;
  const sign = gap >= 0 ? '+' : '';
  return `${sign}${gap.toFixed(0)}pp`;
}

/* ── BLOCO: Conversão (Comercial) ── */
function gerarTextoConversao(dados){
  const c = dados.comercial;
  const totalLeads = c.leads_mensal.reduce((s,x)=>s+x.atendidos,0);
  const totalConv = c.leads_mensal.reduce((s,x)=>s+x.convertidos,0);
  const taxaGeral = totalLeads>0 ? (totalConv/totalLeads*100) : 0;
  const meta = c.config.meta_conversao;
  const mesesAcimaMeta = c.leads_mensal.filter(m=>{
    const t = m.atendidos>0 ? (m.convertidos/m.atendidos*100) : 0;
    return t >= meta;
  }).length;
  const totalMeses = c.leads_mensal.length;
  const ultimoMes = c.leads_mensal[c.leads_mensal.length-1];
  const taxaUltimo = ultimoMes.atendidos>0 ? (ultimoMes.convertidos/ultimoMes.atendidos*100) : 0;
  const status = taxaGeral >= meta ? 'Meta atingida' : 'Abaixo da meta';
  const titulo = taxaGeral >= meta
    ? `Conversão — Meta atingida ${mesesAcimaMeta===totalMeses ? `nos ${totalMeses} meses consecutivos` : `em ${mesesAcimaMeta} de ${totalMeses} meses`}`
    : `Conversão — Abaixo da meta mínima`;
  const texto = `${pct(taxaGeral)} acumulado no período. ${ultimoMes.mes}: ${pct(taxaUltimo)} com ${ultimoMes.atendidos} leads. `+
    `Meta mínima de ${pct(meta)} ${taxaGeral>=meta?'superada':'não atingida'} ${mesesAcimaMeta===totalMeses?'em todos os meses':`em ${mesesAcimaMeta} de ${totalMeses} meses`}.`;
  return {titulo, texto};
}

/* ── BLOCO: Faturamento (Comercial) ── */
function gerarTextoFaturamento(dados){
  const c = dados.comercial.config;
  const fat = c.faturamento_atual;
  const meta = c.meta_faturamento_mensal;
  const pctMeta = meta>0 ? (fat/meta*100) : 0;
  const gap = meta - fat;
  const evolucao = dados.comercial.leads_mensal.map(m=>m.mes).join(' → ');
  const melhorMes = pctMeta >= 90;
  const titulo = pctMeta >= 100 ? 'Faturamento — Meta superada'
    : melhorMes ? 'Faturamento — Melhor resultado do ano'
    : 'Faturamento — Abaixo da meta';
  let texto = `${fmtBRLk(fat)} no mês = ${pct(pctMeta,1)} da meta de ${fmtBRLk(meta)}.`;
  if(gap > 0){
    texto += ` Gap de ${fmtBRLk(gap)} para atingir a meta.`;
  }
  if(dados.desfalque && dados.desfalque.total_mensal > 0){
    const fatComRenegoc = fat + dados.desfalque.total_mensal;
    if(fatComRenegoc >= meta){
      texto += ` Renegociação dos contratos defasados (${fmtBRLk(dados.desfalque.total_mensal)}/mês) permitiria superar a meta sem novos contratos.`;
    }
  }
  return {titulo, texto};
}

/* ── BLOCO: Leads (Comercial) ── */
function gerarTextoLeads(dados){
  const c = dados.comercial;
  const totalLeads = c.leads_mensal.reduce((s,x)=>s+x.atendidos,0);
  const nMeses = c.leads_mensal.length;
  const ritmoAtual = totalLeads / nMeses;
  const metaAnual = c.config.meta_leads_anuais;
  const mesesRestantes = 12 - nMeses;
  const faltam = Math.max(0, metaAnual - totalLeads);
  const ritmoNecessario = mesesRestantes>0 ? faltam/mesesRestantes : 0;
  const noRitmo = ritmoAtual >= ritmoNecessario;
  const titulo = noRitmo ? 'Leads — Ritmo dentro do necessário' : 'Leads — Ritmo abaixo do necessário';
  let texto = `${totalLeads} atendidos em ${nMeses} meses (${ritmoAtual.toFixed(0)}/mês). Meta anual: ${metaAnual} leads. `;
  if(faltam > 0 && mesesRestantes > 0){
    texto += `Faltam ${faltam} leads em ${mesesRestantes} meses — necessário ${ritmoNecessario.toFixed(1).replace('.',',')}/mês.`;
    if(!noRitmo){
      const aumentoPct = ritmoAtual>0 ? ((ritmoNecessario/ritmoAtual - 1)*100) : 0;
      texto += ` Necessário +${aumentoPct.toFixed(0)}% de aumento no ritmo atual.`;
    }
  } else if(faltam === 0){
    texto += `Meta anual já atingida.`;
  }
  return {titulo, texto};
}

/* ── BLOCO: Ocupação (Comercial) ── */
function gerarTextoOcupacao(dados){
  const c = dados.comercial.config;
  const atual = c.ocupacao_atual;
  const meta = c.meta_ocupacao;
  const prevDez = c.ocupacao_previsao_dez;
  const gap = meta - atual;
  const titulo = atual >= meta ? 'Ocupação — Meta atingida' : 'Ocupação — Gap em relação à meta';
  let texto = `${pct(atual)} hoje. Meta de ${pct(meta)}`;
  texto += gap>0 ? ` — gap de ${gap.toFixed(0)}pp.` : ` — superada.`;
  texto += ` Previsão dezembro: ${pct(prevDez)}.`;
  return {titulo, texto};
}

/* ── BLOCO: Vendedores (Comercial) ── */
function gerarTextoVendedores(dados){
  const v = dados.comercial.vendedores;
  if(v.length===0) return {titulo:'Vendedores', texto:'Sem dados de vendedores no período.'};
  const totalContratos = v.reduce((s,x)=>s+x.contratos,0);
  const totalReceita = v.reduce((s,x)=>s+x.receita_anualizada,0);
  const ordenado = [...v].sort((a,b)=>b.contratos-a.contratos);
  const top = ordenado[0];
  const outros = ordenado.slice(1);
  const pctReceitaTop = totalReceita>0 ? (top.receita_anualizada/totalReceita*100) : 0;
  const concentracao = pctReceitaTop >= 80;
  const titulo = concentracao ? `Vendedores — concentração total em ${top.nome}` : 'Vendedores — Desempenho da equipe';
  const taxaTop = top.leads_atendidos>0 ? (top.contratos/top.leads_atendidos*100) : 0;
  let texto = `${top.nome}: ${top.contratos} contratos, ${pct(taxaTop)} conversão, ${pct(pctReceitaTop,0)} da receita (${fmtBRL(top.receita_anualizada)} anualizado).`;
  outros.forEach(o=>{
    const taxaO = o.leads_atendidos>0 ? (o.contratos/o.leads_atendidos*100) : 0;
    if(o.contratos === 0){
      texto += ` ${o.nome}: 0 vendas em ${o.leads_atendidos} leads no período.`;
    } else {
      const palavraContrato = o.contratos === 1 ? 'contrato' : 'contratos';
      texto += ` ${o.nome}: ${o.contratos} ${palavraContrato}, ${pct(taxaO)} conversão.`;
    }
  });
  if(concentracao){
    texto += ` Dependência de único vendedor é risco operacional.`;
  }
  const semVendas = outros.filter(o=>o.contratos===0);
  if(semVendas.length>0){
    texto += ` Plano de ação necessário para ${semVendas.map(o=>o.nome).join(', ')}.`;
  }
  return {titulo, texto};
}

/* ── BLOCO: Desfalque (Comercial/Operacional) ── */
function gerarTextoDesfalque(dados){
  const d = dados.desfalque;
  if(!d || (d.contratos.length===0 && d.posicoes_individuais.length===0)){
    return {titulo:'Desfalque', texto:'Sem dados de desfalque cadastrados para o período.'};
  }
  const fat = dados.comercial.config.faturamento_atual;
  const meta = dados.comercial.config.meta_faturamento_mensal;
  const fatComRenegoc = fat + d.total_mensal;
  const superariaMeta = fatComRenegoc >= meta;
  const titulo = 'Desfalque — ação prioritária no próximo trimestre';
  const nomesContratos = d.contratos.map(c=>`${c.nome} (${fmtBRL(c.desfalque_mensal)}/mês)`);
  const listaContratos = nomesContratos.length<=1
    ? (nomesContratos[0]||'')
    : nomesContratos.slice(0,-1).join(', ') + ' e ' + nomesContratos[nomesContratos.length-1];
  let texto = `Renegociação de ${listaContratos} permitiria alcançar ${fmtBRLk(fatComRenegoc)}/mês`;
  texto += superariaMeta ? ` — acima da meta de ${fmtBRLk(meta)} sem novos contratos.` : `, ainda abaixo da meta de ${fmtBRLk(meta)}.`;
  if(d.posicoes_individuais.length>0){
    const listaPos = d.posicoes_individuais.map(p=>`${p.nome} ${fmtBRL(p.desfalque_anual)}/ano`).join(' · ');
    texto += ` Posições individuais ativas: ${listaPos}.`;
  }
  return {titulo, texto};
}

/* ── BLOCO: Custo Operacional (Operacional) ── */
function gerarTextoCustoOperacional(dados){
  const o = dados.orcamento;
  if(!o || o.realizado_mes===0){
    return {titulo:'Custo operacional', texto:'Sem dados de orçamento cadastrados para o período.'};
  }
  const saldo = o.orcado_mes - o.realizado_mes;
  const saldoPct = o.orcado_mes>0 ? (saldo/o.orcado_mes*100) : 0;
  const positivo = saldo >= 0;
  const consumoEsperado = o.consumo_acumulado_esperado_pct;
  const consumoReal = o.anual>0 ? (o.consumo_acumulado_valor/o.anual*100) : 0;
  const diffConsumo = consumoReal - consumoEsperado;
  const consumoComparativo = Math.abs(diffConsumo) < 0.5 ? 'em linha com'
    : diffConsumo < 0 ? 'abaixo dos' : 'acima dos';
  const titulo = 'Custo operacional' + (positivo ? '' : ' — saldo negativo no mês');
  let texto = `${fmtBRLk(o.realizado_mes)} realizado vs ${fmtBRLk(o.orcado_mes)} orçado — `;
  texto += positivo ? `saldo positivo de ${fmtBRLk(saldo)} (${pct(Math.abs(saldoPct),1)}).` : `saldo negativo de ${fmtBRLk(Math.abs(saldo))} (${pct(Math.abs(saldoPct),1)}).`;
  if(o.cards_custo_concluidos>0){
    const custoMedio = o.custo_total_mes / o.cards_custo_concluidos;
    texto += ` ${o.cards_custo_concluidos} cards concluídos com custo total de ${fmtBRL(o.custo_total_mes)} (média ${fmtBRL(custoMedio)}/card).`;
  }
  if(o.mao_obra_mes>0 || o.material_mes>0){
    texto += ` Mão de obra: ${fmtBRL(o.mao_obra_mes)} · Material: ${fmtBRL(o.material_mes)}.`;
  }
  texto += ` Acumulado em ${pct(consumoReal,0)} do orçamento anual (${consumoComparativo} ${pct(consumoEsperado,0)} esperados).`;
  return {titulo, texto};
}

/* ── BLOCO: Manutenção (Operacional) ── */
function gerarTextoManutencao(dados){
  const m = dados.operacional.manutencao;
  const meta = dados.operacional.config.meta_sla_manutencao;
  const encTotal = m.encerrados_no_prazo + m.encerrados_atrasados;
  const sla = encTotal>0 ? (m.encerrados_no_prazo/encTotal*100) : 0;
  const totalCards = encTotal + m.abertos_no_prazo + m.abertos_atrasados;
  const atingiuMeta = sla >= meta;
  const titulo = atingiuMeta ? `Manutenção — SLA em ${pct(sla,0)}` : `Manutenção — SLA em ${pct(sla,0)}, abaixo da meta`;
  let texto = `${totalCards} cards no período: ${m.encerrados_no_prazo} encerrados no prazo, ${m.encerrados_atrasados} encerrados em atraso`;
  if(m.abertos_no_prazo>0 || m.abertos_atrasados>0){
    texto += `, ${m.abertos_no_prazo} em aberto dentro do prazo`;
    if(m.abertos_atrasados>0) texto += `, ${m.abertos_atrasados} em aberto atrasados`;
  }
  texto += `. SLA ${pct(sla,0)} — ${atingiuMeta?'dentro da':'abaixo da'} meta de ${pct(meta,0)}.`;
  // nível crítico
  const niveisComSla = m.niveis.map(n=>({...n, sla: n.encerrados_total>0 ? (n.encerrados_no_prazo/n.encerrados_total*100) : null})).filter(n=>n.sla!==null);
  if(niveisComSla.length>0){
    const pior = niveisComSla.reduce((a,b)=> a.sla < b.sla ? a : b);
    if(pior.sla < meta){
      texto += ` ${pior.nivel} crítico: apenas ${pct(pior.sla,0)} no prazo.`;
    }
  }
  return {titulo, texto};
}

/* ── BLOCO: Viagens (Operacional) ── */
function gerarTextoViagens(dados){
  const v = dados.operacional.viagens;
  const meta = dados.operacional.config.meta_sla_viagens;
  const sla = v.concluidos_avaliados>0 ? (v.concluidos_no_prazo/v.concluidos_avaliados*100) : 0;
  const atingiuMeta = sla >= meta;
  const gap = sla - meta;
  const titulo = `Viagens — SLA em ${pct(sla,0)}`;
  let texto = `${v.total} viagens no período, ${v.concluidos} concluídas`;
  if(v.cancelados>0) texto += `, ${v.cancelados} canceladas`;
  texto += `. `;
  texto += `${v.concluidos_no_prazo}/${v.concluidos_avaliados} avaliadas no prazo`;
  texto += atingiuMeta ? `. Meta de ${pct(meta,0)} ${Math.abs(gap)<0.5?'atingida':'superada'}.` : `, abaixo da meta de ${pct(meta,0)}.`;
  // acumulado do ano
  const evo = v.evolucao;
  if(evo && evo.length>0){
    const totalAno = evo.reduce((s,m)=>s+m.total,0);
    const noPrazoAno = evo.reduce((s,m)=>s+m.no_prazo,0);
    const slaAno = totalAno>0 ? (noPrazoAno/totalAno*100) : 0;
    texto += ` Acumulado do ano: ${pct(slaAno,1)} (${noPrazoAno}/${totalAno}).`;
  }
  return {titulo, texto};
}

/* ── ORQUESTRADOR: gera todos os blocos, respeitando overrides manuais (Status_Areas) ── */
function gerarRascunhosStatus(dados){
  const geradores = {
    conversao: gerarTextoConversao,
    faturamento: gerarTextoFaturamento,
    leads: gerarTextoLeads,
    ocupacao: gerarTextoOcupacao,
    vendedores: gerarTextoVendedores,
    desfalque: gerarTextoDesfalque,
    custo_operacional: gerarTextoCustoOperacional,
    manutencao: gerarTextoManutencao,
    viagens: gerarTextoViagens
  };
  const manual = dados.statusAreasManual || {};
  const resultado = {};
  for(const [chave, fn] of Object.entries(geradores)){
    if(manual[chave] && manual[chave].texto){
      resultado[chave] = {titulo: manual[chave].titulo || '', texto: manual[chave].texto, origem:'manual'};
    } else {
      const r = fn(dados);
      resultado[chave] = {...r, origem:'automatico'};
    }
  }
  return resultado;
}



/* ── montarDadosDeWorkbook(wb)
   Recebe um workbook SheetJS já lido e devolve
   {operacional, comercial, desfalque, orcamento, statusAreasManual}.
   Usada pelo admin.html (via FileReader) e pelo script Node de automação —
   mesma lógica, sem duplicar. ── */
function montarDadosDeWorkbook(wb){
  const cfg = kvSheetToObj(wb,'Configuracao');
        const mKpi = kvSheetToObj(wb,'Manutencao_KPIs');
        const mNiv = sheetToObjects(wb,'Manutencao_Niveis').map(r=>({
          nivel:String(pick(r,'nivel')),
          prazo_du:num(pick(r,'prazo_du')),
          encerrados_total:num(pick(r,'encerrados_total')),
          encerrados_no_prazo:num(pick(r,'encerrados_no_prazo'))
        }));
        const vKpi = kvSheetToObj(wb,'Viagens_KPIs');
        const vEvo = sheetToObjects(wb,'Viagens_Evolucao').map(r=>({mes:String(pick(r,'mes')),total:num(pick(r,'total')),no_prazo:num(pick(r,'no_prazo'))}));
        const lMen = sheetToObjects(wb,'Leads_Mensal').map(r=>({mes:String(pick(r,'mes')),atendidos:num(pick(r,'atendidos')),convertidos:num(pick(r,'convertidos'))}));
        const vend = sheetToObjects(wb,'Vendedores').map(r=>({
          nome:String(pick(r,'nome')),
          cor:String(pick(r,'cor')||'olive'),
          contratos:num(pick(r,'contratos')),
          receita_anualizada:num(pick(r,'receita_anualizada')),
          leads_atendidos:num(pick(r,'leads_atendidos')),
          convertidos:num(pick(r,'convertidos'))
        }));
        const fun = kvSheetToObj(wb,'Funil');

        // ── NOVO: abas opcionais (Desfalque, Orcamento, Status_Areas) ──
        // Se a planilha não tiver essas abas (formato antigo), os arrays/objetos
        // ficam vazios e o restante do fluxo (zip, dashboards) funciona igual antes.
        const desfRows = sheetToObjectsOpt(wb,'Desfalque').map(r=>({
          tipo: String(pick(r,'tipo')||'contrato'),
          nome: String(pick(r,'nome')||''),
          desfalque_mensal: num(pick(r,'desfalque_mensal')),
          desfalque_anual: num(pick(r,'desfalque_anual')),
          observacao: pick(r,'observacao')||''
        })).filter(r=>r.nome);
        const orc = kvSheetToObjOpt(wb,'Orcamento');
        const statusRows = sheetToObjectsOpt(wb,'Status_Areas').map(r=>({
          bloco: String(pick(r,'bloco')||''),
          titulo: String(pick(r,'titulo')||''),
          texto: String(pick(r,'texto')||'')
        })).filter(r=>r.bloco && r.bloco!=='(opcional)');

        const operacional = {
          config:{
            periodo: cfg.periodo_operacional || cfg.periodo || '—',
            ultima_atualizacao: cfg.ultima_atualizacao || '—',
            meta_sla_manutencao: num(cfg.meta_sla_manutencao||90),
            meta_sla_viagens: num(cfg.meta_sla_viagens||90)
          },
          manutencao:{
            encerrados_no_prazo: num(mKpi.encerrados_no_prazo),
            encerrados_atrasados: num(mKpi.encerrados_atrasados),
            abertos_no_prazo: num(mKpi.abertos_no_prazo),
            abertos_atrasados: num(mKpi.abertos_atrasados),
            niveis: mNiv
          },
          viagens:{
            total: num(vKpi.total),
            concluidos: num(vKpi.concluidos),
            cancelados: num(vKpi.cancelados),
            atrasados_encerrados: num(vKpi.atrasados_encerrados),
            atrasados_abertos: num(vKpi.atrasados_abertos),
            concluidos_no_prazo: num(vKpi.concluidos_no_prazo),
            concluidos_avaliados: num(vKpi.concluidos_avaliados),
            sem_data_encerramento: num(vKpi.sem_data_encerramento),
            evolucao: vEvo
          }
        };

        const comercial = {
          config:{
            periodo_titulo: cfg.periodo_comercial || cfg.periodo || '—',
            mes_referencia: cfg.mes_referencia || '—',
            ultima_atualizacao: cfg.ultima_atualizacao || '—',
            meta_conversao: num(cfg.meta_conversao||15),
            meta_ocupacao: num(cfg.meta_ocupacao||80),
            meta_leads_anuais: num(cfg.meta_leads_anuais||200),
            meta_faturamento_mensal: num(cfg.meta_faturamento_mensal||28000),
            ocupacao_atual: num(cfg.ocupacao_atual),
            ocupacao_previsao_dez: num(cfg.ocupacao_previsao_dez),
            faturamento_atual: num(cfg.faturamento_atual),
            novos_contratos: num(cfg.novos_contratos),
            destaque_contrato: cfg.destaque_contrato || '',
            subtitulo_contratos: cfg.subtitulo_contratos || '',
            callout: cfg.callout || ''
          },
          leads_mensal: lMen,
          vendedores: vend,
          funil:{
            leads_atendidos: num(fun.leads_atendidos),
            em_aberto: num(fun.em_aberto),
            convertidos: num(fun.convertidos),
            churned: num(fun.churned)
          }
        };

        const desfalque = {
          contratos: desfRows.filter(r=>r.tipo==='contrato'),
          posicoes_individuais: desfRows.filter(r=>r.tipo==='posicao_individual'),
          total_mensal: desfRows.reduce((s,r)=>s+r.desfalque_mensal,0),
          total_anual: desfRows.reduce((s,r)=>s+r.desfalque_anual,0)
        };
        const orcamento = {
          realizado_mes: num(orc.orcamento_realizado_mes),
          orcado_mes: num(orc.orcamento_orcado_mes),
          inicial_mes: num(orc.orcamento_inicial_mes),
          anual: num(orc.orcamento_anual),
          consumo_acumulado_valor: num(orc.consumo_acumulado_valor),
          consumo_acumulado_esperado_pct: num(orc.consumo_acumulado_esperado_pct),
          custo_total_mes: num(orc.custo_total_mes),
          mao_obra_mes: num(orc.mao_obra_mes),
          material_mes: num(orc.material_mes),
          cards_custo_concluidos: num(orc.cards_custo_concluidos)
        };
        const statusAreasManual = {};
        statusRows.forEach(r=>{ statusAreasManual[r.bloco] = {titulo:r.titulo, texto:r.texto}; });

  return {operacional,comercial,desfalque,orcamento,statusAreasManual};
}

function pdfBrl(v, dec=0){
  v = Number(v)||0;
  const s = v.toLocaleString('pt-BR', {minimumFractionDigits:dec, maximumFractionDigits:dec});
  return 'R$' + s;
}
function pdfBrlK(v){
  v = Number(v)||0;
  return 'R$' + (v/1000).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1}) + 'k';
}
function pdfPct(v, dec=0){
  v = Number(v)||0;
  return v.toLocaleString('pt-BR',{minimumFractionDigits:dec,maximumFractionDigits:dec}) + '%';
}

/* Guarda contra um problema recorrente: quando a célula de período/data no Excel
   continua como DATA (em vez de texto), o SheetJS devolve a string em inglês/formato
   americano (ex: "May-26", "5/18/26") mesmo a planilha sendo em PT-BR. Esta função
   detecta esses padrões e normaliza para português, evitando que isso vaze pro PDF. */
const MESES_EN_PT = {jan:'Jan',feb:'Fev',mar:'Mar',apr:'Abr',may:'Mai',jun:'Jun',jul:'Jul',aug:'Ago',sep:'Set',oct:'Out',nov:'Nov',dec:'Dez'};
function normalizaPeriodoPt(s){
  if(!s) return s;
  let out = String(s).trim();
  // "May-26" -> "Mai-26"
  const m1 = out.match(/^([A-Za-z]{3})-(\d{2,4})$/);
  if(m1 && MESES_EN_PT[m1[1].toLowerCase()]){
    return `${MESES_EN_PT[m1[1].toLowerCase()]}-${m1[2]}`;
  }
  // "5/18/26" (M/D/YY, formato americano do SheetJS) -> "18/05/2026"
  const m2 = out.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if(m2){
    let [_, mo, da, yr] = m2;
    if(yr.length===2) yr = '20'+yr;
    return `${da.padStart(2,'0')}/${mo.padStart(2,'0')}/${yr}`;
  }
  return out;
}

/* Calcula todos os indicadores derivados a partir de state.dados,
   espelhando exatamente a lógica já usada nos dashboards/motor de texto. */
function calcPdfData(dados){
  const c = dados.comercial, o = dados.operacional;
  const m = o.manutencao, v = o.viagens;
  const desf = dados.desfalque || {contratos:[], posicoes_individuais:[], total_mensal:0, total_anual:0};
  const orc = dados.orcamento || {};

  const totalLeads = c.leads_mensal.reduce((s,x)=>s+x.atendidos,0);
  const totalConv = c.leads_mensal.reduce((s,x)=>s+x.convertidos,0);
  const taxaGeral = totalLeads>0 ? (totalConv/totalLeads*100) : 0;

  const encTotal = m.encerrados_no_prazo + m.encerrados_atrasados;
  const slaManut = encTotal>0 ? (m.encerrados_no_prazo/encTotal*100) : 0;
  const totalCardsManut = encTotal + m.abertos_no_prazo + m.abertos_atrasados;

  const slaViagens = v.concluidos_avaliados>0 ? (v.concluidos_no_prazo/v.concluidos_avaliados*100) : 0;
  const totalEvo = (v.evolucao||[]).reduce((s,x)=>s+x.total,0);
  const noPrazoEvo = (v.evolucao||[]).reduce((s,x)=>s+x.no_prazo,0);
  const slaViagensAcum = totalEvo>0 ? (noPrazoEvo/totalEvo*100) : 0;

  const fat = c.config.faturamento_atual, metaFat = c.config.meta_faturamento_mensal;
  const fatComRenegoc = fat + desf.total_mensal;
  const pctFat = metaFat>0 ? (fat/metaFat*100) : 0;

  const saldoOrc = (orc.orcado_mes||0) - (orc.realizado_mes||0);
  const pctSaldo = orc.orcado_mes ? (saldoOrc/orc.orcado_mes*100) : 0;
  const custoMedioCard = orc.cards_custo_concluidos ? (orc.custo_total_mes/orc.cards_custo_concluidos) : 0;
  const consumoRealPct = orc.anual ? (orc.consumo_acumulado_valor/orc.anual*100) : 0;

  const vendOrdenado = [...c.vendedores].sort((a,b)=>b.contratos-a.contratos);
  const totalContratos = c.vendedores.reduce((s,x)=>s+x.contratos,0);
  const totalReceita = c.vendedores.reduce((s,x)=>s+x.receita_anualizada,0);

  const niveisSla = m.niveis.map(n=>({...n, sla: n.encerrados_total>0 ? (n.encerrados_no_prazo/n.encerrados_total*100) : 0}));

  return {
    c, o, m, v, desf, orc,
    totalLeads, totalConv, taxaGeral,
    encTotal, slaManut, totalCardsManut,
    slaViagens, slaViagensAcum, totalEvo, noPrazoEvo,
    fat, metaFat, fatComRenegoc, pctFat,
    saldoOrc, pctSaldo, custoMedioCard, consumoRealPct,
    vendOrdenado, totalContratos, totalReceita, niveisSla
  };
}

function pdfHeader(periodoLabel, atualizadoLabel, extra){
  const tag = extra ? `<span class="pagetag"> ${extra}</span>` : '';
  return `<div class="pdf-header">
    <h1>NÚMEROS DO MÊS — ${periodoLabel.toUpperCase()}${tag}</h1>
    <div class="meta">${periodoLabel} | Atualizado: ${atualizadoLabel}</div>
  </div>`;
}
function pdfFooterBar(area, periodoLabel, atualizadoLabel){
  return `<div class="pdf-footer-bar"><span>Seja Hype Coworking | Facilities ${area} | Uso interno — confidencial</span><span>${periodoLabel} | Atualizado: ${atualizadoLabel}</span></div>`;
}

/* ─────────────────────────────────────────────────────────────
   COMERCIAL — 3 páginas
   ───────────────────────────────────────────────────────────── */
function buildComercialPdfHtml(dados, textos){
  const D = calcPdfData(dados);
  const T = textos;
  const periodoLabel = normalizaPeriodoPt(D.c.config.mes_referencia) || D.c.config.periodo_titulo || '—';
  const atualizado = normalizaPeriodoPt(D.c.config.ultima_atualizacao) || '—';
  const vend = D.vendOrdenado, lmen = D.c.leads_mensal;

  const vendCardsHtml = vend.map(v=>{
    if(v.contratos>0){
      return `<div class="pdf-kpi"><div class="lbl">${v.nome.toUpperCase()} — VENDAS</div><div class="val green">${v.contratos}<span class="unit2">${v.contratos===1?'sale':'sales'}</span></div>
        <div class="sub">${v.leads_atendidos} leads · ${v.leads_atendidos? (v.convertidos/v.leads_atendidos*100).toFixed(0):0}% conversão</div>
        <div class="sub dot good">${pdfBrl(v.receita_anualizada)} anualizado</div></div>
      <div class="pdf-kpi"><div class="lbl">${v.nome.toUpperCase()} — RECEITA</div><div class="val unit">${pdfBrlK(v.receita_anualizada)}</div>
        <div class="sub">${D.totalReceita? (v.receita_anualizada/D.totalReceita*100).toFixed(0):0}% da receita ativa</div>
        <div class="sub dot good">base ${periodoLabel.split(' ')[0]||''}</div></div>`;
    } else {
      const convPct = v.leads_atendidos ? (v.convertidos/v.leads_atendidos*100) : 0;
      return `<div class="pdf-kpi"><div class="lbl">${v.nome.toUpperCase()} — VENDAS</div><div class="val red">0<span class="unit2">sales</span></div>
        <div class="sub">${v.leads_atendidos} leads atendidos</div>
        <div class="sub dot warn">${convPct.toFixed(0)}% conversão</div></div>
      <div class="pdf-kpi"><div class="lbl">${v.nome.toUpperCase()} — CONVERSÃO</div><div class="val red">${convPct.toFixed(0)}%</div>
        <div class="sub">${v.convertidos} de ${v.leads_atendidos} leads convertidos</div>
        <div class="sub dot warn">plano de ação necessário</div></div>`;
    }
  }).join('');

  return `
<div class="pdf-page">
  ${pdfHeader(periodoLabel, atualizado)}
  <div class="pdf-kpi-grid">
    <div class="pdf-kpi"><div class="lbl">Leads Atendidos</div><div class="val">${D.totalLeads}</div>
      <div class="sub">${lmen.map(r=>r.mes+':'+r.atendidos).join(' · ')}</div>
      <div class="sub dot amber">ritmo: ${(D.totalLeads/lmen.length).toFixed(0)}/mês — meta ${D.c.config.meta_leads_anuais}/ano</div></div>
    <div class="pdf-kpi"><div class="lbl">Leads Convertidos</div><div class="val green">${D.totalConv}</div>
      <div class="sub">acumulado no período</div>
      <div class="sub dot good">taxa geral: ${pdfPct(D.taxaGeral)}</div></div>
    <div class="pdf-kpi"><div class="lbl">Taxa de Conversão</div><div class="val green">${pdfPct(D.taxaGeral)}</div>
      <div class="sub">meta mínima: ${D.c.config.meta_conversao}%</div>
      <div class="sub dot good">acima da meta</div></div>
    <div class="pdf-kpi"><div class="lbl">Contratos Firmados</div><div class="val unit">${D.totalContratos}<span class="unit2">sales</span></div>
      <div class="sub">${vend.map(v=>v.nome+': '+v.contratos).join(', ')}</div>
      <div class="sub dot good">${D.c.config.destaque_contrato||''}</div></div>
  </div>
  <div class="pdf-kpi-grid">
    <div class="pdf-kpi"><div class="lbl">Coworking — Faturamento</div><div class="val unit">${pdfBrlK(D.fat)}</div>
      <div class="sub">Meta: ${pdfBrlK(D.metaFat)}</div>
      <div class="sub dot good">${pdfPct(D.pctFat,1)} da meta</div></div>
    <div class="pdf-kpi"><div class="lbl">Coworking — Ocupação</div><div class="val green">${D.c.config.ocupacao_atual}%</div>
      <div class="sub">Meta: ${D.c.config.meta_ocupacao}%</div>
      <div class="sub dot good">previsão dez: ${D.c.config.ocupacao_previsao_dez}%</div></div>
    <div class="pdf-kpi"><div class="lbl">Coworking — Conversão</div><div class="val green">${pdfPct(D.taxaGeral)}</div>
      <div class="sub">Meta mínima: ${D.c.config.meta_conversao}%</div>
      <div class="sub dot good">acima da meta</div></div>
    <div class="pdf-kpi"><div class="lbl">Novos Contratos</div><div class="val unit">${D.c.config.novos_contratos}</div>
      <div class="sub">${D.c.config.subtitulo_contratos||''}</div>
      <div class="sub dot good">${D.c.config.destaque_contrato||''}</div></div>
  </div>
  <div class="pdf-status-areas">
    <div class="pdf-section-title">Status das Áreas</div>
    ${T.conversao?`<div class="pdf-status-block"><div class="sdot"></div><div><div class="st-title">${T.conversao.titulo}</div><div class="st-text">${T.conversao.texto}</div></div></div>`:''}
    ${T.faturamento?`<div class="pdf-status-block"><div class="sdot"></div><div><div class="st-title">${T.faturamento.titulo}</div><div class="st-text">${T.faturamento.texto}</div></div></div>`:''}
    ${T.leads?`<div class="pdf-status-block"><div class="sdot"></div><div><div class="st-title">${T.leads.titulo}</div><div class="st-text">${T.leads.texto}</div></div></div>`:''}
    ${T.ocupacao?`<div class="pdf-status-block"><div class="sdot"></div><div><div class="st-title">${T.ocupacao.titulo}</div><div class="st-text">${T.ocupacao.texto}</div></div></div>`:''}
  </div>
  <div class="pdf-footnote">Renegociação de contratos existentes é condição necessária para atingir meta de faturamento. Desfalque ativo: ${pdfBrlK(D.desf.total_mensal)}/mês — principal alavanca disponível</div>
  ${pdfFooterBar('Comercial', periodoLabel, atualizado)}
</div>

<div class="pdf-page">
  ${pdfHeader(periodoLabel, atualizado, 'Metas &amp; Vendedores')}
  <div class="pdf-section-title">Metas Anuais — Progresso</div>
  <div class="pdf-kpi-grid">
    <div class="pdf-kpi"><div class="lbl">Leads Anuais</div><div class="val amber">${(D.totalLeads/D.c.config.meta_leads_anuais*100).toFixed(0)}%</div>
      <div class="sub">${D.totalLeads}/${D.c.config.meta_leads_anuais} atendidos</div>
      <div class="sub dot warn">faltam ${Math.max(0,D.c.config.meta_leads_anuais-D.totalLeads)}</div></div>
    <div class="pdf-kpi"><div class="lbl">Ocupação Atual</div><div class="val green">${D.c.config.ocupacao_atual}%</div>
      <div class="sub">Meta: ${D.c.config.meta_ocupacao}%</div>
      <div class="sub dot good">previsão dez: ${D.c.config.ocupacao_previsao_dez}%</div></div>
    <div class="pdf-kpi"><div class="lbl">Faturamento</div><div class="val green">${pdfPct(D.pctFat,1)}</div>
      <div class="sub">${pdfBrlK(D.fat)} de ${pdfBrlK(D.metaFat)}</div>
      <div class="sub dot warn">gap de ${pdfBrlK(Math.max(0,D.metaFat-D.fat))}/mês</div></div>
    <div class="pdf-kpi"><div class="lbl">Contratos (Total)</div><div class="val unit">${D.totalContratos}<span class="unit2">sales</span></div>
      <div class="sub">${vend.map(v=>v.nome+': '+v.contratos).join(', ')}</div></div>
  </div>
  <div class="pdf-section-title">Funil Comercial</div>
  <div class="pdf-progress">
    <div class="pdf-prow"><div class="plabel">Leads atendidos</div><div class="pdf-ptrack"><div class="pdf-pfill" style="width:100%;background:#3a3f47"></div></div><div class="pdf-pval">${D.c.funil.leads_atendidos} (100%)</div></div>
    <div class="pdf-prow"><div class="plabel">Em aberto (neg.)</div><div class="pdf-ptrack"><div class="pdf-pfill" style="width:${(D.c.funil.em_aberto/D.c.funil.leads_atendidos*100)||0}%;background:#ffb24d"></div></div><div class="pdf-pval">${D.c.funil.em_aberto} (${((D.c.funil.em_aberto/D.c.funil.leads_atendidos*100)||0).toFixed(0)}%)</div></div>
    <div class="pdf-prow"><div class="plabel">Convertidos (Sale)</div><div class="pdf-ptrack"><div class="pdf-pfill" style="width:${(D.c.funil.convertidos/D.c.funil.leads_atendidos*100)||0}%;background:#00ffa3"></div></div><div class="pdf-pval">${D.c.funil.convertidos} (${((D.c.funil.convertidos/D.c.funil.leads_atendidos*100)||0).toFixed(0)}%)</div></div>
    <div class="pdf-prow"><div class="plabel">Churned</div><div class="pdf-ptrack"><div class="pdf-pfill" style="width:${(D.c.funil.churned/D.c.funil.leads_atendidos*100)||0}%;background:#e45252"></div></div><div class="pdf-pval">${D.c.funil.churned} (${((D.c.funil.churned/D.c.funil.leads_atendidos*100)||0).toFixed(0)}%)</div></div>
  </div>
  <div class="pdf-section-title">Taxa de Conversão Mensal</div>
  <table class="pdf-table">
    <tr><th>Mês</th><th>Leads</th><th>Sales</th><th>Taxa Real.</th><th>Meta Mín.</th><th>Status</th></tr>
    ${lmen.map(r=>`<tr><td>${r.mes}</td><td>${r.atendidos}</td><td>${r.convertidos}</td><td>${r.atendidos?(r.convertidos/r.atendidos*100).toFixed(0):0}%</td><td>${D.c.config.meta_conversao}%</td><td class="pdf-status-ok">Acima da meta</td></tr>`).join('')}
    <tr><td>TOTAL/MÉDIA</td><td>${D.totalLeads}</td><td>${D.totalConv}</td><td>${pdfPct(D.taxaGeral)}</td><td>${D.c.config.meta_conversao}%</td><td class="pdf-status-ok">Acima da meta</td></tr>
  </table>
  <div class="pdf-section-title">Análise por Vendedor</div>
  <div class="pdf-kpi-grid">${vendCardsHtml}</div>
  <div class="pdf-status-areas">
    <div class="pdf-section-title">Status das Áreas</div>
    ${T.vendedores?`<div class="pdf-status-block"><div class="sdot"></div><div><div class="st-title">${T.vendedores.titulo}</div><div class="st-text">${T.vendedores.texto}</div></div></div>`:''}
  </div>
  ${pdfFooterBar('Comercial', periodoLabel, atualizado)}
</div>

<div class="pdf-page">
  ${pdfHeader(periodoLabel, atualizado, 'Desfalque Financeiro')}
  <div class="pdf-section-title">Desfalque Financeiro — Contratos Ativos</div>
  <div class="pdf-kpi-grid cols3">
    ${D.desf.contratos.map(ct=>`<div class="pdf-kpi"><div class="lbl">DESFALQUE ${ct.nome.toUpperCase()}</div><div class="val red">${pdfBrlK(ct.desfalque_mensal)}<span class="unit2">/mês</span></div>
      <div class="sub">${pdfBrl(ct.desfalque_anual)} ao ano</div>
      <div class="sub dot warn">${ct.observacao||''}</div></div>`).join('')}
  </div>
  <div class="pdf-kpi-grid cols2">
    <div class="pdf-kpi"><div class="lbl">Total Desfalque</div><div class="val red">${pdfBrlK(D.desf.total_mensal)}<span class="unit2">/mês</span></div>
      <div class="sub">${pdfBrl(D.desf.total_anual)} ao ano</div>
      <div class="sub dot warn">impacto direto na meta</div></div>
    <div class="pdf-kpi"><div class="lbl">Receita c/ Renegociação</div><div class="val green">${pdfBrlK(D.fatComRenegoc)}</div>
      <div class="sub">Atual ${pdfBrlK(D.fat)} + ${pdfBrlK(D.desf.total_mensal)}</div>
      <div class="sub dot good">${D.fatComRenegoc>=D.metaFat?'superaria':'ainda abaixo de'} meta de ${pdfBrlK(D.metaFat)}</div></div>
  </div>
  <div class="pdf-section-title">Detalhamento por Contrato</div>
  <table class="pdf-table">
    <tr><th>Contrato / Responsável</th><th>Desfalque Mensal</th><th>Desfalque Anual</th><th>Observação</th></tr>
    ${D.desf.contratos.map(ct=>`<tr><td>${ct.nome}</td><td>${pdfBrl(ct.desfalque_mensal,2)}</td><td>${pdfBrl(ct.desfalque_anual)}</td><td>${ct.observacao||''}</td></tr>`).join('')}
    <tr><td>TOTAL DESFALQUE</td><td>${pdfBrl(D.desf.total_mensal,2)}</td><td>${pdfBrl(D.desf.total_anual)}</td><td>Impacto direto na receita potencial</td></tr>
  </table>
  <div class="pdf-status-areas">
    <div class="pdf-section-title">Status das Áreas</div>
    ${T.desfalque?`<div class="pdf-status-block"><div class="sdot"></div><div><div class="st-title">${T.desfalque.titulo}</div><div class="st-text">${T.desfalque.texto}</div></div></div>`:''}
  </div>
  <div class="pdf-footnote">Renegociação de contratos de desfalque (${pdfBrlK(D.desf.total_mensal)}/mês) é a principal alavanca disponível para superar a meta de ${pdfBrlK(D.metaFat)} sem depender exclusivamente de novos contratos.</div>
  ${pdfFooterBar('Comercial', periodoLabel, atualizado)}
</div>`;
}

/* ─────────────────────────────────────────────────────────────
   OPERACIONAL — 3 páginas
   ───────────────────────────────────────────────────────────── */
function buildOperacionalPdfHtml(dados, textos){
  const D = calcPdfData(dados);
  const T = textos;
  const periodoLabel = normalizaPeriodoPt(D.o.config.periodo) || '—';
  const atualizado = normalizaPeriodoPt(D.o.config.ultima_atualizacao) || '—';
  const desf = D.desf, orc = D.orc, mniv = D.niveisSla;

  return `
<div class="pdf-page">
  ${pdfHeader(periodoLabel, atualizado)}
  <div class="pdf-kpi-grid">
    <div class="pdf-kpi"><div class="lbl">Orçamento Realizado</div><div class="val unit">${pdfBrlK(orc.realizado_mes||0)}</div>
      <div class="sub">Orçado: ${pdfBrlK(orc.orcado_mes||0)}</div>
      <div class="sub dot ${D.saldoOrc>=0?'good':'warn'}">${pdfBrlK(Math.abs(D.saldoOrc))} de ${D.saldoOrc>=0?'saldo':'excedente'}</div></div>
    <div class="pdf-kpi"><div class="lbl">Consumo Acumulado</div><div class="val">${pdfPct(D.consumoRealPct,0)}</div>
      <div class="sub">Orç. anual: ${pdfBrlK(orc.anual||0)}</div>
      <div class="sub dot good">acompanhando o esperado</div></div>
    <div class="pdf-kpi"><div class="lbl">SLA Manutenção</div><div class="val ${D.slaManut>=D.o.config.meta_sla_manutencao?'green':'red'}">${pdfPct(D.slaManut,0)}</div>
      <div class="sub">Meta: ≥ ${D.o.config.meta_sla_manutencao}%</div>
      <div class="sub dot ${D.slaManut>=D.o.config.meta_sla_manutencao?'good':'warn'}">${D.slaManut>=D.o.config.meta_sla_manutencao?'+':''}${(D.slaManut-D.o.config.meta_sla_manutencao).toFixed(0)}pp vs meta</div></div>
    <div class="pdf-kpi"><div class="lbl">SLA Viagens</div><div class="val green">${pdfPct(D.slaViagens,0)}</div>
      <div class="sub">Meta: ≥ ${D.o.config.meta_sla_viagens}%</div>
      <div class="sub dot good">+${(D.slaViagens-D.o.config.meta_sla_viagens).toFixed(0)}pp acima da meta ✓</div></div>
  </div>
  <div class="pdf-kpi-grid">
    <div class="pdf-kpi"><div class="lbl">Custo Operacional</div><div class="val unit">${pdfBrl(orc.custo_total_mes||0,2)}</div>
      <div class="sub">${orc.cards_custo_concluidos||0} cards concluídos</div>
      <div class="sub dot amber">${pdfBrl(D.custoMedioCard,2)} por card (média)</div></div>
    <div class="pdf-kpi"><div class="lbl">Mão de Obra + Material</div><div class="val unit">${pdfBrl(orc.mao_obra_mes||0,2)}</div>
      <div class="sub">M.obra: ${pdfBrl(orc.mao_obra_mes||0,2)} | Mat.: ${pdfBrl(orc.material_mes||0,2)}</div>
      <div class="sub">pós-obra incluída</div></div>
    <div class="pdf-kpi"><div class="lbl">Saldo do Mês</div><div class="val ${D.saldoOrc>=0?'green':'red'}">${pdfBrlK(Math.abs(D.saldoOrc))}</div>
      <div class="sub">Realizado vs. orçado</div>
      <div class="sub dot ${D.saldoOrc>=0?'good':'warn'}">${D.saldoOrc>=0?'saldo positivo':'saldo negativo'}</div></div>
    <div class="pdf-kpi"><div class="lbl">Economia vs. Inicial</div><div class="val unit">${pdfBrlK((orc.inicial_mes||0)-(orc.orcado_mes||0))}</div>
      <div class="sub">Orçado inicial: ${pdfBrlK(orc.inicial_mes||0)}</div>
      <div class="sub dot good">após redução aprovada</div></div>
  </div>
  <div class="pdf-status-areas">
    <div class="pdf-section-title">Status das Áreas</div>
    ${T.custo_operacional?`<div class="pdf-status-block"><div class="sdot"></div><div><div class="st-title">${T.custo_operacional.titulo}</div><div class="st-text">${T.custo_operacional.texto}</div></div></div>`:''}
    ${T.manutencao?`<div class="pdf-status-block"><div class="sdot"></div><div><div class="st-title">${T.manutencao.titulo}</div><div class="st-text">${T.manutencao.texto}</div></div></div>`:''}
    ${T.viagens?`<div class="pdf-status-block"><div class="sdot"></div><div><div class="st-title">${T.viagens.titulo}</div><div class="st-text">${T.viagens.texto}</div></div></div>`:''}
  </div>
  ${pdfFooterBar('Operacional', periodoLabel, atualizado)}
</div>

<div class="pdf-page">
  ${pdfHeader(periodoLabel, atualizado, 'Detalhe Operacional')}
  <div class="pdf-section-title">Manutenção — Detalhamento</div>
  <div class="pdf-kpi-grid">
    <div class="pdf-kpi"><div class="lbl">Total Cards</div><div class="val">${D.totalCardsManut}</div>
      <div class="sub">${D.m.encerrados_no_prazo} conc. · ${D.m.abertos_no_prazo} ab. prazo</div></div>
    <div class="pdf-kpi"><div class="lbl">Concluídos no Prazo</div><div class="val green">${D.m.encerrados_no_prazo}<span class="unit2">/${D.encTotal}</span></div>
      <div class="sub">${pdfPct(D.slaManut,0)} dos encerrados</div>
      <div class="sub dot good">SLA encerrados: ${pdfPct(D.slaManut,0)}</div></div>
    <div class="pdf-kpi"><div class="lbl">Custo Total</div><div class="val amber">${pdfBrl(orc.custo_total_mes||0,2)}</div>
      <div class="sub">Média: ${pdfBrl(D.custoMedioCard,2)}/card</div></div>
    <div class="pdf-kpi"><div class="lbl">SLA Geral (Todos)</div><div class="val ${D.slaManut>=D.o.config.meta_sla_manutencao?'green':'red'}">${pdfPct(D.slaManut,0)}</div>
      <div class="sub">Meta: ≥ ${D.o.config.meta_sla_manutencao}%</div>
      <div class="sub dot ${D.m.abertos_no_prazo+D.m.abertos_atrasados>0?'amber':'good'}">${D.m.abertos_no_prazo+D.m.abertos_atrasados} cards em aberto</div></div>
  </div>
  <div class="pdf-section-title">Desempenho por Nível</div>
  <table class="pdf-table">
    <tr><th>Nível</th><th>Prazo</th><th>Total</th><th>Enc.Prazo</th><th>Enc.Atras</th><th>% SLA</th></tr>
    ${mniv.map(n=>`<tr><td>${n.nivel}</td><td>${n.prazo_du} du</td><td>${n.encerrados_total}</td><td>${n.encerrados_no_prazo}</td><td>${n.encerrados_total-n.encerrados_no_prazo}</td><td>${n.sla.toFixed(0)}%</td></tr>`).join('')}
    <tr><td>TOTAL</td><td>—</td><td>${mniv.reduce((s,n)=>s+n.encerrados_total,0)}</td><td>${mniv.reduce((s,n)=>s+n.encerrados_no_prazo,0)}</td><td>${mniv.reduce((s,n)=>s+(n.encerrados_total-n.encerrados_no_prazo),0)}</td><td>${(mniv.reduce((s,n)=>s+n.encerrados_no_prazo,0)/mniv.reduce((s,n)=>s+n.encerrados_total,0)*100).toFixed(1)}%</td></tr>
  </table>
  <div class="pdf-section-title">Custo Operacional</div>
  <div class="pdf-list">
    <div class="lrow"><span>Mão de obra (pós-obra)</span><span>${pdfBrl(orc.mao_obra_mes||0,2)}</span></div>
    <div class="lrow"><span>Material</span><span>${pdfBrl(orc.material_mes||0,2)}</span></div>
    <div class="lrow"><span>Total do mês</span><span class="amber">${pdfBrl(orc.custo_total_mes||0,2)}</span></div>
    <div class="lrow"><span>Cards concluídos</span><span>${orc.cards_custo_concluidos||0} cards</span></div>
    <div class="lrow"><span>Custo médio por card</span><span class="amber">${pdfBrl(D.custoMedioCard,2)}</span></div>
  </div>
  <div class="pdf-section-title">Viagens — Detalhamento</div>
  <div class="pdf-kpi-grid">
    <div class="pdf-kpi"><div class="lbl">Total Viagens</div><div class="val">${D.v.total}</div>
      <div class="sub">${D.v.concluidos} concluídas · ${D.v.cancelados} canceladas</div></div>
    <div class="pdf-kpi"><div class="lbl">No Prazo</div><div class="val green">${D.v.concluidos_no_prazo}<span class="unit2">/${D.v.concluidos_avaliados}</span></div>
      <div class="sub">${pdfPct(D.slaViagens,0)} das avaliadas</div></div>
    <div class="pdf-kpi"><div class="lbl">SLA Viagens</div><div class="val green">${pdfPct(D.slaViagens,0)}</div>
      <div class="sub dot good">Meta: ≥ ${D.o.config.meta_sla_viagens}% ✓</div></div>
    <div class="pdf-kpi"><div class="lbl">Acumulado</div><div class="val green">${pdfPct(D.slaViagensAcum,1)}</div>
      <div class="sub">${D.noPrazoEvo}/${D.totalEvo} no prazo</div></div>
  </div>
  <div class="pdf-section-title">Orçamento &amp; Acumulado</div>
  <div class="pdf-kpi-grid cols3">
    <div class="pdf-kpi"><div class="lbl">Realizado</div><div class="val unit">${pdfBrlK(orc.realizado_mes||0)}</div>
      <div class="sub">Orçado: ${pdfBrlK(orc.orcado_mes||0)}</div>
      <div class="sub dot ${D.saldoOrc>=0?'good':'warn'}">${D.saldoOrc>=0?'Saldo':'Excedente'}: ${pdfBrlK(Math.abs(D.saldoOrc))}</div></div>
    <div class="pdf-kpi"><div class="lbl">Orçado Inicial</div><div class="val unit">${pdfBrlK(orc.inicial_mes||0)}</div>
      <div class="sub">Orç. c/ redução: ${pdfBrlK(orc.orcado_mes||0)}</div></div>
    <div class="pdf-kpi"><div class="lbl">Consumo Acum.</div><div class="val green">${pdfPct(D.consumoRealPct,0)}</div>
      <div class="sub">${pdfBrlK(orc.consumo_acumulado_valor||0)} de ${pdfBrlK(orc.anual||0)}</div></div>
  </div>
  ${pdfFooterBar('Operacional', periodoLabel, atualizado)}
</div>`;
}

// Monta os textos revisados (editados ou rascunho automático) no formato
// {chave: {titulo, texto}} esperado pelos builders, a partir de state.statusTextos

function slugPeriodo(s){
  return String(s||'periodo')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[^a-zA-Z0-9]+/g,'_').replace(/^_+|_+$/g,'') || 'periodo';
}

// Renderiza um HTML de relatório num iframe oculto e aciona a impressão nativa
// do navegador (window.print()). O motor de print do Chrome/Edge é o mesmo
// Chromium que renderiza a tela — ele entende CSS Grid corretamente, ao
// contrário do html2canvas (que "fotografa" a página e quebra em layouts

/* ── Exporta tudo no Node; no navegador module é undefined, não roda. ── */
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    PDF_CSS,
    norm, resolveSheet, pick, sheetToObjects, sheetToObjectsOpt,
    kvSheetToObjOpt, kvSheetToObj, num, fmtBRL, fmtBRLk, pct, gapTxt,
    gerarTextoConversao, gerarTextoFaturamento, gerarTextoLeads,
    gerarTextoOcupacao, gerarTextoVendedores, gerarTextoDesfalque,
    gerarTextoCustoOperacional, gerarTextoManutencao, gerarTextoViagens,
    gerarRascunhosStatus,
    montarDadosDeWorkbook,
    pdfBrl, pdfBrlK, pdfPct, normalizaPeriodoPt, calcPdfData,
    pdfHeader, pdfFooterBar, buildComercialPdfHtml, buildOperacionalPdfHtml,
    slugPeriodo
  };
}
