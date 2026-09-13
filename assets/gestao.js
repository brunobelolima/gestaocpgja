(() => {
  'use strict';

  const EXPECTED = ['Pacientes','Leads','Atendimentos','Programas','Financeiro','Custos','Capacidade','Qualidade'];
  const state = { sheets:{}, selectedSheet:'Pacientes', selectedRow:null, charts:{}, filters:{start:'',end:'',program:''} };
  const $ = s => document.querySelector(s);
  const money = new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'});
  const num = new Intl.NumberFormat('pt-BR',{maximumFractionDigits:1});

  function status(msg, error=false){ const el=$('#status'); el.textContent=msg; el.style.color=error?'#A33A31':'#627276'; }
  function normalize(v){ return (v ?? '').toString().trim(); }
  function isPatientNameHeader(h){ const x=normalize(h).toLowerCase().replace(/[_-]+/g,' ').replace(/\s+/g,' '); return ['nome','nome do paciente','nome paciente','paciente'].includes(x); }
  function truthy(v){ return ['sim','s','yes','true','1'].includes(normalize(v).toLowerCase()); }
  function parseMoney(v){ if(typeof v==='number') return v; const s=normalize(v).replace(/R\$\s?/g,'').replace(/\./g,'').replace(',','.'); const n=Number(s); return Number.isFinite(n)?n:0; }
  function dateValue(v){ if(v instanceof Date) return v; if(typeof v==='number' && window.XLSX){ const d=XLSX.SSF.parse_date_code(v); return d?new Date(d.y,d.m-1,d.d):null; } const s=normalize(v); if(!s) return null; const br=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); if(br) return new Date(+br[3],+br[2]-1,+br[1]); const d=new Date(s); return Number.isNaN(d.getTime())?null:d; }
  function monthKey(v){ const d=dateValue(v); return d?`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`:normalize(v).slice(0,7); }
  function inPeriod(v){ const d=dateValue(v); if(!state.filters.start && !state.filters.end) return true; if(!d) return true; if(state.filters.start && d < new Date(state.filters.start+'T00:00:00')) return false; if(state.filters.end && d > new Date(state.filters.end+'T23:59:59')) return false; return true; }
  function getRows(name){ return state.sheets[name]?.rows || []; }
  function getHeaders(name){ return state.sheets[name]?.headers || []; }
  function rowObj(name,row){ const h=getHeaders(name); const o={}; h.forEach((x,i)=>o[x]=row[i]); return o; }
  function objects(name){ return getRows(name).map(r=>rowObj(name,r)); }

  function findHeaderRow(rows){
    const keywords=['ID Paciente','ID Lead','ID Atendimento','Código','ID Lançamento','ID Custo','Competência','ID Registro'];
    for(let i=0;i<Math.min(rows.length,12);i++) if(rows[i].some(v=>keywords.includes(normalize(v)))) return i;
    return 0;
  }

  function workbookToState(wb){
    const sheets={};
    wb.SheetNames.forEach(name=>{
      const aoa=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:'',raw:true});
      const hi=findHeaderRow(aoa);
      let headers=(aoa[hi]||[]).map(x=>normalize(x));
      if(!headers.some(Boolean)) return;
      const originalWidth=headers.length;
      let rows=aoa.slice(hi+1).map(r=>Array.from({length:originalWidth},(_,i)=>r[i]??'')).filter(r=>r.some(v=>normalize(v)!==''));
      // Privacidade: a área de gestão não utiliza nem mantém a coluna com nomes de pacientes.
      // Na aba Pacientes, qualquer coluna explicitamente identificada como nome é removida da memória do navegador.
      if(name==='Pacientes'){
        const keep=headers.map((h,i)=>({h,i})).filter(x=>!isPatientNameHeader(x.h));
        headers=keep.map(x=>x.h);
        rows=rows.map(r=>keep.map(x=>r[x.i]));
      }
      sheets[name]={headers,rows};
    });
    state.sheets=sheets;
    state.selectedSheet=EXPECTED.find(x=>sheets[x]) || Object.keys(sheets)[0] || '';
    state.selectedRow=null;
    refreshAll();
  }

  async function loadArrayBuffer(buffer,label){
    try{ const wb=XLSX.read(buffer,{type:'array',cellDates:true}); workbookToState(wb); status(`${label} carregado com sucesso.`); }
    catch(e){ console.error(e); status('Não foi possível ler o arquivo Excel.',true); }
  }

  async function loadTemplate(){
    status('Carregando modelo…');
    try{ const res=await fetch('assets/Painel_Gestao_Clinica_Atualizado.xlsx',{cache:'no-store'}); if(!res.ok) throw new Error('HTTP '+res.status); await loadArrayBuffer(await res.arrayBuffer(),'Modelo'); }
    catch(e){ console.error(e); status('Modelo não carregado. Use “Carregar Excel”.',true); }
  }

  function exportWorkbook(){
    const wb=XLSX.utils.book_new();
    Object.entries(state.sheets).forEach(([name,s])=>{ const aoa=[s.headers,...s.rows]; XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(aoa),name.slice(0,31)); });
    XLSX.writeFile(wb,'Painel_Gestao_Clinica_Atualizado.xlsx');
  }

  function refreshSelectors(){
    const ss=$('#sheetSelect'); ss.innerHTML=''; Object.keys(state.sheets).forEach(name=>{ const o=document.createElement('option');o.value=name;o.textContent=name;ss.appendChild(o); }); if(state.selectedSheet) ss.value=state.selectedSheet;
    const programs=objects('Programas').map(x=>normalize(x['Código'])).filter(Boolean);
    const pf=$('#programFilter'); const current=state.filters.program; pf.innerHTML='<option value="">Todos</option>'+programs.map(p=>`<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`).join(''); pf.value=current;
  }

  function escapeHtml(s){ return normalize(s).replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c])); }

  function renderTable(){
    const name=state.selectedSheet; const t=$('#dataTable'); t.innerHTML=''; if(!name||!state.sheets[name]) return;
    const {headers,rows}=state.sheets[name]; const thead=document.createElement('thead'), hr=document.createElement('tr');
    if(name==='Pacientes'){ const th=document.createElement('th'); th.textContent='Nº'; hr.appendChild(th); }
    headers.forEach(h=>{ const th=document.createElement('th');th.textContent=h||'Coluna';hr.appendChild(th); }); thead.appendChild(hr);t.appendChild(thead);
    const tb=document.createElement('tbody');
    rows.forEach((row,ri)=>{ const tr=document.createElement('tr'); if(state.selectedRow===ri) tr.classList.add('selected'); tr.addEventListener('click',()=>{state.selectedRow=ri;renderTable();}); if(name==='Pacientes'){ const n=document.createElement('td'); n.textContent=String(ri+1); n.className='row-number'; tr.appendChild(n); } headers.forEach((_,ci)=>{ const td=document.createElement('td');td.contentEditable='true';td.textContent=row[ci]??'';td.addEventListener('click',e=>e.stopPropagation());td.addEventListener('focus',()=>{state.selectedRow=ri;});td.addEventListener('blur',()=>{row[ci]=td.textContent.trim(); refreshDashboard();});tr.appendChild(td); });tb.appendChild(tr); }); t.appendChild(tb);
  }

  function filteredPatients(){ return objects('Pacientes').filter(r=>{ const active=normalize(r['Status']).toLowerCase()==='ativo'; const p=!state.filters.program || normalize(r['Programa atual'])===state.filters.program; return active&&p&&inPeriod(r['Data cadastro']); }); }
  function filteredLeads(){ return objects('Leads').filter(r=>inPeriod(r['Data contato'])); }
  function filteredFinance(){ return objects('Financeiro').filter(r=>inPeriod(r['Competência']||r['Data pagamento'])); }
  function filteredCosts(){ return objects('Custos').filter(r=>inPeriod(r['Competência']||r['Data'])); }

  function setText(id,v){ $(id).textContent=v; }
  function groupCount(items,key){ const m={};items.forEach(x=>{const k=normalize(x[key])||'Não informado';m[k]=(m[k]||0)+1;});return m; }
  function chart(id,type,labels,data,label){ if(state.charts[id]) state.charts[id].destroy(); const ctx=$(id); state.charts[id]=new Chart(ctx,{type,data:{labels,datasets:[{label,data,borderWidth:1}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:type==='doughnut',position:'bottom'}},scales:type==='doughnut'?{}:{y:{beginAtZero:true}}}}); }

  function refreshDashboard(){
    if(!state.sheets.Pacientes) return;
    const patients=filteredPatients(), leads=filteredLeads(), finance=filteredFinance(), costs=filteredCosts();
    const converted=leads.filter(r=>truthy(r['Convertido?'])).length;
    const received=finance.reduce((a,r)=>a+parseMoney(r['Valor recebido']),0);
    const expectedOpen=finance.filter(r=>normalize(r['Status']).toLowerCase()!=='pago').reduce((a,r)=>a+Math.max(0,parseMoney(r['Valor previsto'])-parseMoney(r['Valor recebido'])),0);
    const costTotal=costs.reduce((a,r)=>a+parseMoney(r['Valor']),0);
    const priceMap={}; objects('Programas').forEach(r=>priceMap[normalize(r['Código'])]=parseMoney(r['Valor']));
    const mrr=patients.reduce((a,r)=>a+(priceMap[normalize(r['Programa atual'])]||0),0);
    setText('#kpiPatients',patients.length); setText('#kpiLeads',leads.length); setText('#kpiConversion',leads.length?`${num.format(converted/leads.length*100)}%`:'0%'); setText('#kpiMrr',money.format(mrr)); setText('#kpiRevenue',money.format(received)); setText('#kpiCosts',money.format(costTotal)); setText('#kpiMargin',money.format(received-costTotal)); setText('#kpiDefault',money.format(expectedOpen));

    const pg=groupCount(patients,'Programa atual'); chart('#programChart','doughnut',Object.keys(pg),Object.values(pg),'Pacientes');
    const stages=['Contato','Agendado','Avaliação realizada','Proposta','Convertido']; const sv=stages.map(s=>leads.filter(r=>{const e=normalize(r['Etapa']).toLowerCase(); if(s==='Contato') return true; if(s==='Agendado') return ['agendado','avaliação realizada','proposta','convertido'].some(x=>e.includes(x)); if(s==='Avaliação realizada') return ['avaliação realizada','proposta','convertido'].some(x=>e.includes(x)); if(s==='Proposta') return ['proposta','convertido'].some(x=>e.includes(x)); return truthy(r['Convertido?']);}).length); chart('#leadChart','bar',stages,sv,'Leads');
    const months=[...new Set([...finance.map(r=>monthKey(r['Competência']||r['Data pagamento'])),...costs.map(r=>monthKey(r['Competência']||r['Data']))].filter(Boolean))].sort(); if(state.charts.finance) state.charts.finance.destroy(); state.charts.finance=new Chart($('#financeChart'),{type:'bar',data:{labels:months,datasets:[{label:'Receita',data:months.map(m=>finance.filter(r=>monthKey(r['Competência']||r['Data pagamento'])===m).reduce((a,r)=>a+parseMoney(r['Valor recebido']),0))},{label:'Custos',data:months.map(m=>costs.filter(r=>monthKey(r['Competência']||r['Data'])===m).reduce((a,r)=>a+parseMoney(r['Valor']),0))}]},options:{responsive:true,maintainAspectRatio:false,scales:{y:{beginAtZero:true}}}});
    const og=groupCount(leads,'Origem'); chart('#originChart','doughnut',Object.keys(og),Object.values(og),'Leads');

    const cap=objects('Capacidade').filter(r=>inPeriod(r['Competência'])); $('#capacitySummary').innerHTML=cap.length?cap.slice(0,8).map(r=>{const occ=parseMoney(r['Ocupação']);const label=[normalize(r['Recurso']),normalize(r['Profissional'])].filter(Boolean).join(' — ')||'Capacidade';return `<div class="list-row"><span>${escapeHtml(label)}</span><strong>${Number.isFinite(occ)?num.format(occ*100)+'%':'—'}</strong></div>`;}).join(''):'<div class="list-row"><span>Sem registros no período</span><strong>—</strong></div>';
    const qual=objects('Qualidade').filter(r=>inPeriod(r['Data'])); const sats=qual.map(r=>Number(r['Satisfação 0-10'])).filter(Number.isFinite); const nps=qual.map(r=>Number(r['NPS 0-10'])).filter(Number.isFinite); const ps=qual.filter(r=>truthy(r['PS não planejado'])).length; const ints=qual.filter(r=>truthy(r['Internação não planejada'])).length; $('#qualitySummary').innerHTML=[['Registros',qual.length],['Satisfação média',sats.length?num.format(sats.reduce((a,b)=>a+b,0)/sats.length):'—'],['NPS médio',nps.length?num.format(nps.reduce((a,b)=>a+b,0)/nps.length):'—'],['PS não planejado',ps],['Internações não planejadas',ints]].map(x=>`<div class="list-row"><span>${x[0]}</span><strong>${x[1]}</strong></div>`).join('');
  }

  function refreshAll(){ refreshSelectors(); renderTable(); refreshDashboard(); }

  document.querySelectorAll('.tab').forEach(btn=>btn.addEventListener('click',()=>{document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));btn.classList.add('active');document.querySelectorAll('.tab-panel').forEach(x=>x.classList.add('hidden'));$('#'+btn.dataset.tab).classList.remove('hidden'); if(btn.dataset.tab==='planilha') renderTable();}));
  $('#fileInput').addEventListener('change',async e=>{const f=e.target.files[0];if(f) await loadArrayBuffer(await f.arrayBuffer(),f.name);e.target.value='';});
  $('#loadTemplate').addEventListener('click',loadTemplate); $('#exportXlsx').addEventListener('click',exportWorkbook);
  $('#clearData').addEventListener('click',()=>{if(confirm('Limpar todos os dados carregados desta sessão?')){state.sheets={};state.selectedSheet='';document.querySelectorAll('canvas').forEach(c=>{const ch=Chart.getChart(c);if(ch)ch.destroy();});refreshSelectors();renderTable();['#kpiPatients','#kpiLeads'].forEach(id=>setText(id,'0')); status('Sessão limpa. Nenhum dado foi salvo no servidor.');}});
  $('#sheetSelect').addEventListener('change',e=>{state.selectedSheet=e.target.value;state.selectedRow=null;renderTable();});
  $('#addRow').addEventListener('click',()=>{const s=state.sheets[state.selectedSheet];if(!s)return;s.rows.push(Array(s.headers.length).fill(''));state.selectedRow=s.rows.length-1;renderTable();});
  $('#deleteRow').addEventListener('click',()=>{const s=state.sheets[state.selectedSheet];if(!s||state.selectedRow===null)return;if(confirm('Excluir a linha selecionada?')){s.rows.splice(state.selectedRow,1);state.selectedRow=null;renderTable();refreshDashboard();}});
  $('#applyFilters').addEventListener('click',()=>{state.filters={start:$('#dateStart').value,end:$('#dateEnd').value,program:$('#programFilter').value};refreshDashboard();});
  $('#resetFilters').addEventListener('click',()=>{$('#dateStart').value='';$('#dateEnd').value='';$('#programFilter').value='';state.filters={start:'',end:'',program:''};refreshDashboard();});

  window.addEventListener('load',()=>{ if(window.XLSX&&window.Chart) loadTemplate(); else status('Bibliotecas externas não carregaram. Verifique a conexão com a internet.',true); });
})();
