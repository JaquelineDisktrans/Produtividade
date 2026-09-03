(() => {
  'use strict';
  const $ = (selector) => document.querySelector(selector);
  const esc = (value) => String(value ?? '').replace(/[&<>"']/g, (char) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
  const state = { data: null, view: 'overview', search: '', status: 'todos' };
  const ptDate = (iso) => iso ? new Intl.DateTimeFormat('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(new Date(iso)) : '—';
  const n = (value) => new Intl.NumberFormat('pt-BR').format(Number(value || 0));
  const percent = (value) => `${Number(value || 0).toLocaleString('pt-BR', { maximumFractionDigits: 1 })}%`;
  const mins = (value) => {
    if (value === '' || value === null || value === undefined || Number.isNaN(Number(value))) return '—';
    const total = Math.round(Number(value));
    return `${Math.floor(total / 60)}h ${String(total % 60).padStart(2, '0')}min`;
  };
  const showError = (message) => { const node = $('#error'); node.textContent = message; node.hidden = false; };
  const clearError = () => { $('#error').hidden = true; };
  const maxValue = (rows) => Math.max(1, ...(rows || []).map((row) => Number(row.value || 0)));

  function renderBars(target, rows, labelKey, limit = 8) {
    const node = $(target); node.textContent = '';
    const values = (rows || []).slice(-limit); const max = maxValue(values);
    values.forEach((row) => {
      const col = document.createElement('div'); col.className = 'bar-col';
      const value = document.createElement('span'); value.className = 'bar-value'; value.textContent = n(row.value);
      const bar = document.createElement('div'); bar.className = 'bar'; bar.style.height = `${Math.max(2, Number(row.value || 0) / max * 100)}%`; bar.setAttribute('data-tooltip', `${row[labelKey]}: ${n(row.value)} casos`);
      const label = document.createElement('span'); label.className = 'bar-label'; label.textContent = row[labelKey];
      col.append(value, bar, label); node.append(col);
    });
  }

  function renderSla() {
    const node = $('#sla-chart'); node.textContent = '';
    const k = state.data.kpis; const answered = Number(k.respondidos || 0); const responseValues = state.data.casos.map((c) => Number(c.tempo_resposta || 0)).filter((v) => v > 0);
    const groups = [[30, 'Até 30 min'], [60, 'Até 1 hora'], [120, 'Até 2 horas'], [240, 'Até 4 horas'], [1440, 'Até 24h / D+1']];
    groups.forEach(([limit, label]) => {
      const count = responseValues.filter((value) => value <= limit).length; const row = document.createElement('div'); row.className = 'sla-row';
      row.innerHTML = `<span class="sla-label">${label}</span><div class="sla-track"><div class="sla-fill" style="width:${answered ? count / answered * 100 : 0}%"></div></div><span class="sla-value">${answered ? (count / answered * 100).toFixed(1).replace('.', ',') : '0,0'}%</span>`;
      node.append(row);
    });
  }

  function renderWeekdays() {
    const node = $('#weekday-chart'); node.textContent = ''; const rows = state.data.volume.dia_semana || []; const max = maxValue(rows);
    const order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']; const names = { Monday: 'Seg', Tuesday: 'Ter', Wednesday: 'Qua', Thursday: 'Qui', Friday: 'Sex', Saturday: 'Sáb', Sunday: 'Dom' };
    order.forEach((day) => { const row = rows.find((item) => item.day === day) || { value: 0 }; const col = document.createElement('div'); col.className = 'weekday-col'; col.innerHTML = `<span class="weekday-value">${n(row.value)}</span><div class="weekday-bar" style="height:${Math.max(2, Number(row.value) / max * 100)}%"></div><span class="weekday-label">${names[day]}</span>`; node.append(col); });
  }

  function renderRisks() {
    const node = $('#risk-list'); node.textContent = ''; const k = state.data.kpis;
    const risks = [{ level: 'red', title: `${n(k.sem_resposta)} casos sem resposta aparente`, note: `${percent(k.sem_resposta / Math.max(1, k.casos) * 100)} dos casos consolidados` }, { level: 'red', title: `${percent(k.cobranca / Math.max(1, k.casos) * 100)} com possível cobrança`, note: 'sinal heurístico de falta de retorno/status' }, { level: 'orange', title: `${n(k.evitaveis)} contatos potencialmente evitáveis`, note: 'status, prazo, Portal, cadastro ou cobrança' }];
    risks.forEach((risk) => { const item = document.createElement('div'); item.className = 'risk-item'; item.innerHTML = `<span class="risk-level ${risk.level}"></span><div><strong>${risk.title}</strong><span>${risk.note}</span></div>`; node.append(item); });
  }

  function renderHorizontal(target, rows, color = 'blue', limit = 10) {
    const node = $(target); node.textContent = ''; const values = (rows || []).slice(0, limit); const max = maxValue(values);
    values.forEach((row) => { const item = document.createElement('div'); item.className = 'hbar-row'; item.innerHTML = `<span class="hbar-label" title="${esc(row.label)}">${esc(row.label)}</span><div class="hbar-track"><div class="hbar-fill" style="width:${Number(row.value || 0) / max * 100}%;background:${color === 'orange' ? 'var(--orange)' : 'var(--blue)'}"></div></div><span class="hbar-value">${n(row.value)}</span>`; node.append(item); });
  }

  function renderMini(target, rows) { const node = $(target); node.textContent = ''; (rows || []).slice(0, 8).forEach((row) => { const item = document.createElement('div'); item.className = 'mini-row'; item.innerHTML = `<span title="${esc(row.label)}">${esc(row.label)}</span><span>${n(row.value)}</span>`; node.append(item); }); if (!(rows || []).length) node.textContent = 'Não identificado'; }

  function renderCases() {
    const query = state.search.toLocaleLowerCase('pt-BR'); const rows = state.data.casos.filter((c) => (state.status === 'todos' || c.status === state.status) && (!query || [c.assunto, c.cliente, c.unidade, c.motivo].join(' ').toLocaleLowerCase('pt-BR').includes(query))).slice(0, 200);
    const body = $('#cases-body'); body.textContent = '';
    rows.forEach((c) => { const tr = document.createElement('tr'); const statusClass = c.status === 'Respondido' ? 'ok' : 'wait'; tr.innerHTML = `<td>${ptDate(c.data)}</td><td><strong>${esc(c.cliente)}</strong><span>${esc(c.unidade)}</span></td><td><strong title="${esc(c.assunto)}">${esc(c.assunto)}</strong></td><td>${esc(c.motivo)}</td><td><span class="status ${statusClass}">${esc(c.status)}</span></td><td>${c.tempo_resposta === '' ? '—' : mins(c.tempo_resposta)}</td><td>${n(c.interacoes)}</td>`; body.append(tr); });
    $('#case-count').textContent = `${n(rows.length)} atendimentos exibidos${state.data.casos.length > rows.length ? ` · limite de 200 para navegação rápida` : ''}`;
  }

  function render() {
    const d = state.data; const k = d.kpis; $('#subtitle').textContent = `${d.meta.caixa} · ${d.meta.periodo}`; $('#generated').textContent = `Atualizado ${ptDate(d.meta.gerado_em)}`;
    $('#k-cases').textContent = n(k.casos); $('#k-received').textContent = n(k.recebidos); $('#k-response').textContent = percent(k.taxa_resposta); $('#k-response-note').textContent = `${n(k.respondidos)} de ${n(k.casos)} casos`; $('#k-unanswered').textContent = n(k.sem_resposta); $('#k-avg').textContent = k.tempo_medio_legivel || '—'; $('#k-median').textContent = `mediana ${k.tempo_mediano_legivel || '—'}`;
    const riskRate = Number(k.sem_resposta || 0) / Math.max(1, Number(k.casos || 0)) * 100; $('#insight-title').textContent = riskRate >= 50 ? 'A fila pede atenção' : 'Operação sob acompanhamento'; $('#insight-text').textContent = `${n(k.sem_resposta)} casos aparecem sem resposta aparente e a primeira resposta média está em ${k.tempo_medio_legivel}. Use a aba de atendimentos para revisar os casos prioritários.`;
    renderBars('#monthly-chart', d.volume.mensal, 'month', 8); renderSla(); renderWeekdays(); renderRisks(); renderHorizontal('#reasons-chart', d.motivos, 'blue', 10); renderHorizontal('#bottleneck-chart', d.gargalos, 'orange', 10); renderMini('#clients-table', d.clientes); renderMini('#units-table', d.unidades); renderCases();
  }

  async function load() { clearError(); try { const response = await fetch('saida/dashboard_data.json', { cache: 'no-store' }); if (!response.ok) throw new Error('Arquivo de dados não encontrado'); state.data = await response.json(); render(); } catch (error) { showError('Não foi possível carregar os dados. Execute gerar_dashboard_data.py e abra o painel pelo abrir_dashboard.bat.'); $('#subtitle').textContent = 'Aguardando a base local'; $('#generated').textContent = 'Sem dados'; } }
  document.querySelectorAll('.side-link').forEach((button) => button.addEventListener('click', () => { state.view = button.dataset.view; document.querySelectorAll('.side-link').forEach((item) => item.classList.toggle('active', item === button)); document.querySelectorAll('.view').forEach((view) => { view.hidden = view.id !== `view-${state.view}`; }); }));
  $('#refresh').addEventListener('click', load); $('#case-search').addEventListener('input', (event) => { state.search = event.target.value; if (state.data) renderCases(); }); $('#case-status').addEventListener('change', (event) => { state.status = event.target.value; if (state.data) renderCases(); });
  load();
})();
