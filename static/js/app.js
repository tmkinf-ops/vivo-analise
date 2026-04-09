'use strict';
/* ============================================================
   AUDITORIA TELECOM — Single Page Application
   ============================================================ */

// -------------------------------------------------------
// CONSTANTS
// -------------------------------------------------------
const PAGE_TITLES = {
  dashboard:      'Dashboard',
  cadastro:       'Cadastro de Linhas',
  contratos:      'Base de Contratos',
  'upload-contas':'Upload de Contas',
  'fatura-xls':   'Fatura — Importar & Comparar',
  coopernac:      'Coopernac',
  comparacoes:    'Conferência',
  historico:      'Histórico de Importações',
  relatorios:     'Relatórios e Exportação',
  configuracoes:  'Configurações',
};

const STATUS_LABELS = {
  ok:           'Conforme',
  aproximado:   'Aproximado',
  divergente:   'Divergente',
  sem_contrato: 'Sem Contrato',
  sem_fatura:   'Sem Fatura',
  ambiguo:      'Ambíguo',
};

const STATUS_ICONS = {
  ok:           'fa-circle-check',
  aproximado:   'fa-circle-half-stroke',
  divergente:   'fa-circle-xmark',
  sem_contrato: 'fa-circle-question',
  sem_fatura:   'fa-circle-minus',
  ambiguo:      'fa-triangle-exclamation',
};

// -------------------------------------------------------
// APPLICATION STATE
// -------------------------------------------------------
const S = {
  charts: {},
  contratos: { page: 1, search: '', operadora: '' },
  comparacoes: { page: 1, status: '', competencia: '', operadora: '', search: '' },
  uploadContrato: { step: 1, file: null, rows: [], filename: '' },
  uploadConta:    { step: 1, file: null, rows: [], filename: '', result: null },
  histDetailChart: null,
};

// -------------------------------------------------------
// API HELPERS
// -------------------------------------------------------
const API = {
  async get(url) {
    const r = await fetch(url);
    if (!r.ok) { const d = await r.json().catch(() => ({})); throw new Error(d.error || `Erro HTTP ${r.status}`); }
    return r.json();
  },
  async post(url, body) {
    const r = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    const d = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(d.error || `Erro HTTP ${r.status}`);
    return d;
  },
  async put(url, body) {
    const r = await fetch(url, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    const d = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(d.error || `Erro HTTP ${r.status}`);
    return d;
  },
  async del(url) {
    const r = await fetch(url, { method: 'DELETE' });
    if (!r.ok) throw new Error(`Erro HTTP ${r.status}`);
    return r.json().catch(() => ({}));
  },
  async upload(url, file) {
    const fd = new FormData();
    fd.append('file', file);
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), 120000); // 2 min timeout para PDFs grandes
    try {
      const r = await fetch(url, { method: 'POST', body: fd, signal: ctrl.signal });
      clearTimeout(timer);
      const d = await r.json().catch(() => ({}));
      if (!r.ok) throw new Error(d.error || `Erro HTTP ${r.status}`);
      return d;
    } catch (e) {
      clearTimeout(timer);
      if (e.name === 'AbortError') throw new Error('Tempo limite excedido ao processar o PDF (> 2 min). Tente um arquivo menor.');
      throw e;
    }
  },
};

// -------------------------------------------------------
// UTILITIES
// -------------------------------------------------------
const U = {
  money(v) {
    if (v === null || v === undefined || v === '') return '—';
    return Number(v).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  },
  num2(v) {
    if (v === null || v === undefined) return '—';
    return Number(v).toFixed(2).replace('.', ',');
  },
  pct(v) {
    if (v === null || v === undefined) return '—';
    return `${Number(v).toFixed(1)}%`;
  },
  phone(v) {
    if (!v) return '—';
    const d = String(v).replace(/\D/g, '');
    if (d.length === 13) return `+${d.slice(0,2)} (${d.slice(2,4)}) ${d.slice(4,9)}-${d.slice(9)}`;
    if (d.length === 11) return `(${d.slice(0,2)}) ${d.slice(2,7)}-${d.slice(7)}`;
    if (d.length === 10) return `(${d.slice(0,2)}) ${d.slice(2,6)}-${d.slice(6)}`;
    return v;
  },
  esc(s) {
    return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  },
  diffClass(v) {
    if (v === null || v === undefined) return '';
    if (v > 0.005) return 'val-pos';
    if (v < -0.005) return 'val-neg';
    return 'val-zero';
  },
  statusBadge(s) {
    const label = STATUS_LABELS[s] || s || '—';
    const icon  = STATUS_ICONS[s]  || 'fa-circle';
    return `<span class="badge badge-${s || 'sem_fatura'}"><i class="fas ${icon}"></i> ${label}</span>`;
  },
  // Debounce
  debounce(fn, ms) {
    let t;
    return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
  },
  // Toast
  toast(msg, type = 'success') {
    const el = document.createElement('div');
    el.className = `toast toast-${type}`;
    const iconMap = { success: 'fa-check-circle', error: 'fa-circle-xmark', info: 'fa-circle-info', warning: 'fa-triangle-exclamation' };
    el.innerHTML = `<i class="fas ${iconMap[type] || 'fa-circle-info'}"></i><span>${msg}</span>`;
    document.getElementById('toasts').appendChild(el);
    requestAnimationFrame(() => el.classList.add('show'));
    setTimeout(() => { el.classList.remove('show'); setTimeout(() => el.remove(), 400); }, 3800);
  },
  // Loading overlay
  loading(show, txt = 'Processando...') {
    document.getElementById('loading-msg').textContent = txt;
    document.getElementById('loading-overlay').classList.toggle('hidden', !show);
  },
  // Modal
  modal(html, size = '') {
    document.getElementById('modal-body').innerHTML = html;
    const box = document.getElementById('modal-box');
    box.className = 'modal-box' + (size ? ` modal-${size}` : '');
    document.getElementById('modal-backdrop').classList.remove('hidden');
    document.body.style.overflow = 'hidden';
  },
  closeModal() {
    document.getElementById('modal-backdrop').classList.add('hidden');
    document.body.style.overflow = '';
  },
  // Empty state
  emptyState(icon, title, subtitle, actions = '') {
    return `<div class="empty-state">
      <div class="empty-icon"><i class="fas ${icon}"></i></div>
      <h3>${title}</h3><p>${subtitle}</p>${actions}
    </div>`;
  },
};

// -------------------------------------------------------
// NAVIGATION
// -------------------------------------------------------
function navigate(page) {
  document.querySelectorAll('.page-section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));

  const section = document.getElementById(`page-${page}`);
  if (!section) return;
  section.classList.add('active');

  const nav = document.querySelector(`.nav-item[data-page="${page}"]`);
  if (nav) nav.classList.add('active');

  document.getElementById('page-title').textContent = PAGE_TITLES[page] || page;

  switch (page) {
    case 'dashboard':      Dashboard.load(); break;
    case 'cadastro':       Cadastro.load(); break;
    case 'contratos':      Contratos.load(); break;
    case 'upload-contas':  UploadContas.render(); break;
    case 'fatura-xls':     FaturaXls.render(); break;
    case 'coopernac':      Coopernac.load(); break;
    case 'comparacoes':    Comparacoes.load(); break;
    case 'historico':      Historico.load(); break;
    case 'relatorios':     Relatorios.render(); break;
    case 'configuracoes':  Configuracoes.load(); break;
  }
}

function modalBackdropClick(e) {
  if (e.target === document.getElementById('modal-backdrop')) U.closeModal();
}

// -------------------------------------------------------
// DASHBOARD
// -------------------------------------------------------
const Dashboard = {
  async load() {
    const sec = document.getElementById('page-dashboard');
    sec.innerHTML = `<div class="page-header"><h1>Dashboard</h1></div><div class="metrics-grid">${'<div class="metric-card skeleton"></div>'.repeat(6)}</div>`;
    try {
      const [d, alertas] = await Promise.all([
        API.get('/api/dashboard'),
        API.get('/api/alertas/vencimento').catch(() => []),
      ]);
      sec.innerHTML = Dashboard.html(d, alertas);
      Dashboard.charts(d);
    } catch(e) {
      sec.innerHTML = U.emptyState('fa-triangle-exclamation', 'Erro ao carregar', e.message);
    }
  },

  html(d, alertas = []) {
    const conf = d.conformidade;
    const confCls = conf >= 90 ? 'success' : conf >= 70 ? 'warning' : 'danger';
    const finDiff = d.total_faturado - d.total_contratado;
    return `
      <div class="page-header">
        <h1>Dashboard</h1>
        <div class="page-header-actions">
          <button class="btn btn-outline" onclick="Dashboard.load()"><i class="fas fa-rotate-right"></i> Atualizar</button>
        </div>
      </div>

      <div class="metrics-grid">
        <div class="metric-card">
          <div class="metric-icon metric-blue"><i class="fas fa-file-signature"></i></div>
          <div class="metric-info"><span class="metric-value">${d.total_contratos}</span><span class="metric-label">Contratos Ativos</span></div>
        </div>
        <div class="metric-card">
          <div class="metric-icon metric-purple"><i class="fas fa-sim-card"></i></div>
          <div class="metric-info"><span class="metric-value">${d.total_linhas_ativas}</span><span class="metric-label">Linhas Ativas</span></div>
        </div>
        <div class="metric-card">
          <div class="metric-icon metric-indigo"><i class="fas fa-file-invoice"></i></div>
          <div class="metric-info"><span class="metric-value">${d.total_contas}</span><span class="metric-label">Contas Importadas</span></div>
        </div>
        <div class="metric-card">
          <div class="metric-icon metric-green"><i class="fas fa-circle-check"></i></div>
          <div class="metric-info"><span class="metric-value">${d.ok_count}</span><span class="metric-label">Conferências OK</span></div>
        </div>
        <div class="metric-card">
          <div class="metric-icon metric-red"><i class="fas fa-circle-xmark"></i></div>
          <div class="metric-info"><span class="metric-value">${d.div_count}</span><span class="metric-label">Divergências</span></div>
        </div>
        <div class="metric-card metric-conformidade">
          <div class="metric-icon metric-${confCls}"><i class="fas fa-gauge-high"></i></div>
          <div class="metric-info">
            <span class="metric-value ${confCls === 'success' ? 'text-success' : confCls === 'warning' ? 'text-warning' : 'text-danger'}">${conf}%</span>
            <span class="metric-label">Conformidade Geral</span>
          </div>
          <div class="conformidade-bar"><div class="conformidade-fill bg-${confCls}" style="width:${conf}%"></div></div>
        </div>
      </div>

      <div class="cards-row">
        <div class="summary-card"><span class="summary-label">Total Contratado</span><span class="summary-value">${U.money(d.total_contratado)}</span></div>
        <div class="summary-card"><span class="summary-label">Total Faturado</span><span class="summary-value ${finDiff > 0.005 ? 'text-danger' : ''}">${U.money(d.total_faturado)}</span></div>
        <div class="summary-card"><span class="summary-label">Diferença Total</span><span class="summary-value ${finDiff > 0.005 ? 'val-pos' : finDiff < -0.005 ? 'val-neg' : 'val-zero'}">${U.money(finDiff)}</span></div>
        <div class="summary-card"><span class="summary-label">Aproximados</span><span class="summary-value text-warning">${d.aprox_count}</span></div>
        <div class="summary-card"><span class="summary-label">Sem Contrato</span><span class="summary-value text-info">${d.sem_contrato_count}</span></div>
        <div class="summary-card"><span class="summary-label">Sem Fatura</span><span class="summary-value text-muted">${d.sem_fatura_count}</span></div>
      </div>

      ${d.fl_total > 0 ? `
      <div class="section-title" style="margin-top:28px;margin-bottom:12px">
        <h3 style="font-size:15px;font-weight:600;color:var(--text-primary)">
          <i class="fas fa-file-invoice-dollar" style="color:#660099;margin-right:6px"></i>Auditoria de Faturas
        </h3>
      </div>
      <div class="cards-row">
        <div class="summary-card"><span class="summary-label">Linhas Importadas</span><span class="summary-value">${d.fl_total}</span></div>
        <div class="summary-card"><span class="summary-label">Total Faturado</span><span class="summary-value">${U.money(d.fl_total_faturado)}</span></div>
        <div class="summary-card"><span class="summary-label">Total Contrato</span><span class="summary-value">${U.money(d.fl_total_contrato)}</span></div>
        <div class="summary-card"><span class="summary-label">Diferença Total</span><span class="summary-value ${d.fl_diferenca > 0.005 ? 'val-pos' : d.fl_diferenca < -0.005 ? 'val-neg' : 'val-zero'}">${U.money(d.fl_diferenca)}</span></div>
        <div class="summary-card"><span class="summary-label">Divergentes</span><span class="summary-value text-danger">${d.fl_divergentes}</span></div>
        <div class="summary-card"><span class="summary-label">Conformes</span><span class="summary-value text-success">${d.fl_ok}</span></div>
      </div>` : `
      <div class="section-title" style="margin-top:28px;margin-bottom:12px">
        <h3 style="font-size:15px;font-weight:600;color:var(--text-primary)">
          <i class="fas fa-file-invoice-dollar" style="color:#660099;margin-right:6px"></i>Auditoria de Faturas
        </h3>
      </div>
      <div class="cards-row">
        <div class="summary-card" style="grid-column:1/-1;text-align:center;color:var(--text-secondary);padding:24px">
          <i class="fas fa-inbox" style="font-size:20px;margin-bottom:6px;display:block"></i>
          Nenhuma fatura importada.
          <button class="btn-link" style="margin-left:4px" onclick="navigate('fatura-xls')">Importar agora</button>
        </div>
      </div>`}

      <div class="charts-grid">
        <div class="chart-card">
          <h3 class="chart-title"><i class="fas fa-chart-pie" style="color:var(--primary);margin-right:6px"></i>Status das Comparações</h3>
          <div class="chart-wrap"><canvas id="chart-status"></canvas></div>
        </div>
        <div class="chart-card chart-wide">
          <h3 class="chart-title"><i class="fas fa-chart-bar" style="color:var(--primary);margin-right:6px"></i>Evolução Mensal</h3>
          <div class="chart-wrap"><canvas id="chart-monthly"></canvas></div>
        </div>
        <div class="chart-card">
          <h3 class="chart-title"><i class="fas fa-building" style="color:var(--primary);margin-right:6px"></i>Por Operadora</h3>
          <div class="chart-wrap"><canvas id="chart-op"></canvas></div>
        </div>
      </div>

      <div class="quick-actions">
        <h3>Ações Rápidas</h3>
        <div class="quick-actions-grid">
          <button class="quick-btn" onclick="navigate('cadastro')"><i class="fas fa-address-book"></i><span>Cadastro de Linhas</span></button>
          <button class="quick-btn" onclick="navigate('contratos')"><i class="fas fa-file-signature"></i><span>Gerenciar Contratos</span></button>
          <button class="quick-btn" onclick="navigate('upload-contas')"><i class="fas fa-cloud-upload-alt"></i><span>Importar Conta</span></button>
          <button class="quick-btn" onclick="navigate('fatura-xls')"><i class="fas fa-file-excel"></i><span>Importar Faturas</span></button>
          <button class="quick-btn" onclick="navigate('comparacoes')"><i class="fas fa-scale-balanced"></i><span>Ver Conferências</span></button>
          <button class="quick-btn" onclick="navigate('relatorios')"><i class="fas fa-file-export"></i><span>Exportar Relatório</span></button>
        </div>
      </div>

      ${alertas && alertas.length > 0 ? `
      <div class="config-card" style="border-left:4px solid var(--danger);margin-top:20px">
        <h3 style="color:var(--danger);font-size:15px;margin-bottom:12px">
          <i class="fas fa-bell" style="margin-right:6px"></i>Alertas de Vencimento
        </h3>
        <table class="data-table" style="font-size:13px">
          <thead><tr>
            <th>Número</th><th>Funcionário</th><th>Empresa</th><th>Dia Venc.</th><th>Dias Restantes</th><th>Valor</th><th>Urgência</th>
          </tr></thead>
          <tbody>${alertas.map(a => {
            const urgCls = a.urgencia === 'vencido' ? 'badge-divergente' : a.urgencia === 'critico' ? 'badge-divergente' : a.urgencia === 'alerta' ? 'badge-aproximado' : 'badge-ok';
            const urgLabel = a.urgencia === 'vencido' ? 'Vencido' : a.urgencia === 'critico' ? 'Crítico' : a.urgencia === 'alerta' ? 'Atenção' : 'Normal';
            return `<tr>
              <td>${U.phone(a.numero_telefone)}</td>
              <td>${U.esc(a.nome_funcionario || '—')}</td>
              <td>${U.esc(a.empresa || '—')}</td>
              <td style="text-align:center">${a.vencimento_dia}</td>
              <td style="text-align:center;font-weight:600">${a.dias_faltam <= 0 ? 'Vencido' : a.dias_faltam + ' dias'}</td>
              <td>${U.money(a.valor_plano)}</td>
              <td><span class="badge ${urgCls}">${urgLabel}</span></td>
            </tr>`;
          }).join('')}</tbody>
        </table>
      </div>` : ''}`;
  },

  charts(d) {
    Object.values(S.charts).forEach(c => { try { c.destroy(); } catch(e){} });
    S.charts = {};

    const total = d.ok_count + d.aprox_count + d.div_count + d.sem_contrato_count + d.sem_fatura_count + d.ambiguo_count;
    if (total > 0) {
      S.charts.status = new Chart(document.getElementById('chart-status'), {
        type: 'doughnut',
        data: {
          labels: ['Conforme', 'Aproximado', 'Divergente', 'Sem Contrato', 'Sem Fatura', 'Ambíguo'],
          datasets: [{ data: [d.ok_count, d.aprox_count, d.div_count, d.sem_contrato_count, d.sem_fatura_count, d.ambiguo_count],
            backgroundColor: ['#22c55e', '#f59e0b', '#ef4444', '#3b82f6', '#9ca3af', '#f97316'], borderWidth: 2, borderColor: '#fff' }],
        },
        options: {
          responsive: true, maintainAspectRatio: false, cutout: '65%',
          plugins: { legend: { position: 'bottom', labels: { font: { family: 'Inter', size: 11 }, padding: 10 } } },
        },
      });
    } else {
      const cv = document.getElementById('chart-status');
      if (cv) { const ctx = cv.getContext('2d'); ctx.fillStyle = '#f1f5f9'; ctx.fillRect(0,0,cv.width,cv.height); }
    }

    if (d.monthly_data && d.monthly_data.length > 0) {
      S.charts.monthly = new Chart(document.getElementById('chart-monthly'), {
        type: 'bar',
        data: {
          labels: d.monthly_data.map(m => m.competencia),
          datasets: [
            { label: 'Conforme',   data: d.monthly_data.map(m => m.ok),        backgroundColor: '#22c55e88', borderColor: '#22c55e', borderWidth: 1 },
            { label: 'Aproximado', data: d.monthly_data.map(m => m.aproximado), backgroundColor: '#f59e0b88', borderColor: '#f59e0b', borderWidth: 1 },
            { label: 'Divergente', data: d.monthly_data.map(m => m.divergente), backgroundColor: '#ef444488', borderColor: '#ef4444', borderWidth: 1 },
          ],
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: { legend: { position: 'bottom', labels: { font: { family: 'Inter', size: 11 } } } },
          scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true } },
        },
      });
    }

    if (d.operadoras && d.operadoras.length > 0) {
      S.charts.op = new Chart(document.getElementById('chart-op'), {
        type: 'doughnut',
        data: {
          labels: d.operadoras.map(o => o.operadora),
          datasets: [{ data: d.operadoras.map(o => o.count),
            backgroundColor: ['#6366f1','#22c55e','#f59e0b','#ef4444','#3b82f6','#8b5cf6','#14b8a6','#f97316'],
            borderWidth: 2, borderColor: '#fff' }],
        },
        options: {
          responsive: true, maintainAspectRatio: false, cutout: '65%',
          plugins: { legend: { position: 'bottom', labels: { font: { family: 'Inter', size: 11 }, padding: 10 } } },
        },
      });
    }
  },
};

// -------------------------------------------------------
// CONTRATOS
// -------------------------------------------------------
const Contratos = {
  async load() {
    const sec = document.getElementById('page-contratos');
    sec.innerHTML = this.shell();
    document.getElementById('ct-search').addEventListener('input', U.debounce(() => {
      S.contratos.search = document.getElementById('ct-search').value;
      S.contratos.page = 1;
      Contratos.fetch();
    }, 380));
    document.getElementById('ct-op').addEventListener('change', () => {
      S.contratos.operadora = document.getElementById('ct-op').value;
      S.contratos.page = 1;
      Contratos.fetch();
    });
    await this.fetch();
  },

  shell() {
    return `
      <div class="page-header">
        <h1>Base de Contratos</h1>
        <div class="page-header-actions">
          <button class="btn btn-outline" onclick="Contratos.openUpload()"><i class="fas fa-file-pdf"></i> Upload PDF</button>
          <button class="btn btn-primary" onclick="Contratos.openForm()"><i class="fas fa-plus"></i> Novo Contrato</button>
        </div>
      </div>
      <div class="filter-bar">
        <div class="search-field">
          <i class="fas fa-search"></i>
          <input id="ct-search" type="text" placeholder="Buscar por contrato, linha ou cliente..." value="${U.esc(S.contratos.search)}">
        </div>
        <select id="ct-op" class="select-filter">
          <option value="">Todas as operadoras</option>
          <option value="Vivo" ${S.contratos.operadora==='Vivo'?'selected':''}>Vivo</option>
          <option value="Claro" ${S.contratos.operadora==='Claro'?'selected':''}>Claro</option>
          <option value="Tim" ${S.contratos.operadora==='Tim'?'selected':''}>Tim</option>
          <option value="Oi" ${S.contratos.operadora==='Oi'?'selected':''}>Oi</option>
        </select>
        <button class="btn btn-ghost" onclick="Contratos.clearFilters()"><i class="fas fa-filter-circle-xmark"></i> Limpar</button>
      </div>
      <div id="ct-body"><div class="loading-inline"><div class="spinner-sm"></div> Carregando...</div></div>`;
  },

  clearFilters() {
    S.contratos.search = ''; S.contratos.operadora = ''; S.contratos.page = 1;
    const si = document.getElementById('ct-search'); if(si) si.value = '';
    const op = document.getElementById('ct-op'); if(op) op.value = '';
    this.fetch();
  },

  async fetch() {
    const body = document.getElementById('ct-body');
    if (!body) return;
    body.innerHTML = '<div class="loading-inline"><div class="spinner-sm"></div> Carregando...</div>';
    try {
      const url = `/api/contratos?page=${S.contratos.page}&per_page=20&search=${encodeURIComponent(S.contratos.search)}&operadora=${encodeURIComponent(S.contratos.operadora)}`;
      const res = await API.get(url);
      body.innerHTML = this.table(res);
    } catch(e) {
      body.innerHTML = `<div class="alert alert-error"><i class="fas fa-triangle-exclamation"></i>${e.message}</div>`;
    }
  },

  table(res) {
    if (!res.data.length) {
      const hint = S.contratos.search ? 'Tente outros termos.' : 'Clique em "Novo Contrato" ou faça upload de um PDF de contrato.';
      return U.emptyState('fa-file-signature', 'Nenhum contrato encontrado', hint,
        `<button class="btn btn-primary" onclick="Contratos.openForm()"><i class="fas fa-plus"></i> Novo Contrato</button>
         <button class="btn btn-outline" style="margin-left:8px" onclick="Contratos.openUpload()"><i class="fas fa-file-pdf"></i> Upload PDF</button>`);
    }
    const rows = res.data.map(c => `
      <tr>
        <td><span class="mono">${U.esc(c.numero_contrato||'—')}</span></td>
        <td><span class="phone-number">${U.phone(c.linha_telefone)}</span></td>
        <td><strong>${U.money(c.valor_contratado)}</strong></td>
        <td>${U.esc(c.cliente||'—')}</td>
        <td>${c.operadora?`<span class="chip">${U.esc(c.operadora)}</span>`:'—'}</td>
        <td class="text-muted text-sm">${U.esc(c.vigencia_inicio||'—')}</td>
        <td class="text-muted text-sm">${U.esc(c.data_importacao||'—')}</td>
        <td>
          <div class="table-actions">
            <button class="btn-icon-sm" title="Editar" onclick="Contratos.openForm(${c.id})"><i class="fas fa-pencil"></i></button>
            <button class="btn-icon-sm btn-danger" title="Excluir" onclick="Contratos.confirmDelete(${c.id})"><i class="fas fa-trash"></i></button>
          </div>
        </td>
      </tr>`).join('');

    const pages = Math.ceil(res.total / 20);
    const pag = pages > 1 ? `
      <div class="pagination">
        <button class="btn-page" onclick="Contratos.goPage(${S.contratos.page-1})" ${S.contratos.page<=1?'disabled':''}><i class="fas fa-chevron-left"></i></button>
        <span>Página ${S.contratos.page} de ${pages} &nbsp;·&nbsp; ${res.total} registros</span>
        <button class="btn-page" onclick="Contratos.goPage(${S.contratos.page+1})" ${S.contratos.page>=pages?'disabled':''}><i class="fas fa-chevron-right"></i></button>
      </div>` : `<div class="table-footer-info">${res.total} registro(s)</div>`;

    return `
      <div class="table-wrapper">
        <table class="data-table">
          <thead><tr>
            <th>Nº Contrato</th><th>Linha</th><th>Valor Contratado</th>
            <th>Cliente</th><th>Operadora</th><th>Vigência Início</th>
            <th>Importado em</th><th>Ações</th>
          </tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>${pag}`;
  },

  goPage(p) { if (p<1) return; S.contratos.page=p; this.fetch(); },

  openForm(id) {
    if (id) {
      API.get(`/api/contratos/${id}`).then(c => {
        U.modal(this.formHtml('Editar Contrato', c));
        document.getElementById('ctform').dataset.id = id;
      }).catch(e => U.toast(e.message,'error'));
    } else {
      U.modal(this.formHtml('Novo Contrato', {}));
    }
  },

  formHtml(title, c) {
    return `
      <h2 class="modal-title"><i class="fas fa-file-signature" style="color:var(--primary)"></i>${title}</h2>
      <form id="ctform" onsubmit="Contratos.saveForm(event)" class="modal-form">
        <div class="form-grid">
          <div class="form-group"><label>Nº do Contrato</label><input name="numero_contrato" value="${U.esc(c.numero_contrato||'')}" placeholder="Ex: CTR-2024-001"></div>
          <div class="form-group required"><label>Linha / Telefone *</label><input name="linha_telefone" value="${U.esc(c.linha_telefone||'')}" placeholder="(11) 99999-9999" required></div>
          <div class="form-group required"><label>Valor Contratado (R$) *</label><input type="number" name="valor_contratado" value="${c.valor_contratado??''}" step="0.01" min="0" placeholder="0.00" required></div>
          <div class="form-group"><label>Cliente / Empresa</label><input name="cliente" value="${U.esc(c.cliente||'')}" placeholder="Nome do cliente"></div>
          <div class="form-group"><label>Operadora</label><input name="operadora" value="${U.esc(c.operadora||'')}" placeholder="Vivo, Claro, Tim..."></div>
          <div class="form-group"><label>Observações</label><textarea name="observacoes" rows="2">${U.esc(c.observacoes||'')}</textarea></div>
        </div>
        <div class="form-actions">
          <button type="button" class="btn btn-ghost" onclick="U.closeModal()">Cancelar</button>
          <button type="submit" class="btn btn-primary"><i class="fas fa-check"></i> Salvar</button>
        </div>
      </form>`;
  },

  async saveForm(e) {
    e.preventDefault();
    const f = e.target;
    const id = f.dataset.id;
    const data = {
      numero_contrato: f.numero_contrato.value.trim(),
      linha_telefone:  f.linha_telefone.value.trim(),
      valor_contratado: parseFloat(f.valor_contratado.value),
      cliente:   f.cliente.value.trim(),
      operadora: f.operadora.value.trim(),
      observacoes: f.observacoes.value.trim(),
    };
    try {
      U.loading(true, id ? 'Atualizando contrato...' : 'Salvando contrato...');
      if (id) { await API.put(`/api/contratos/${id}`, data); U.toast('Contrato atualizado!'); }
      else     { await API.post('/api/contratos', data);      U.toast('Contrato cadastrado!'); }
      U.closeModal(); U.loading(false); Contratos.fetch();
    } catch(err) { U.loading(false); U.toast(err.message,'error'); }
  },

  confirmDelete(id) {
    U.modal(`
      <div class="confirm-dialog">
        <div class="confirm-icon"><i class="fas fa-triangle-exclamation text-danger"></i></div>
        <h3>Confirmar exclusão</h3>
        <p>Deseja desativar este contrato? O registro será preservado no histórico e não aparecerá mais nas conferências.</p>
        <div class="form-actions">
          <button class="btn btn-ghost" onclick="U.closeModal()">Cancelar</button>
          <button class="btn btn-danger" onclick="Contratos.doDelete(${id})"><i class="fas fa-trash"></i> Excluir</button>
        </div>
      </div>`);
  },

  async doDelete(id) {
    try {
      U.loading(true,'Excluindo...');
      await API.del(`/api/contratos/${id}`);
      U.toast('Contrato desativado.','info');
      U.closeModal(); U.loading(false); Contratos.fetch();
    } catch(e) { U.loading(false); U.toast(e.message,'error'); }
  },

  openUpload() {
    S.uploadContrato = { step: 1, file: null, rows: [], filename: '' };
    U.modal(this.uploadModalHtml(), 'lg');
    this.bindDropzone();
  },

  uploadModalHtml() {
    return `
      <h2 class="modal-title"><i class="fas fa-file-pdf" style="color:#ef4444"></i> Importar Contratos via PDF</h2>
      <div class="wizard-steps" id="ct-wsteps">
        <div class="wizard-step active" id="ct-ws1"><span>1</span> Upload</div>
        <div class="wizard-step" id="ct-ws2"><span>2</span> Revisar Dados</div>
        <div class="wizard-step" id="ct-ws3"><span>3</span> Concluído</div>
      </div>
      <div id="ct-wbody">${this.uploadStep1()}</div>`;
  },

  uploadStep1() {
    return `
      <div class="dropzone" id="ct-dz">
        <input type="file" id="ct-file" accept=".pdf" style="display:none">
        <i class="fas fa-file-pdf dropzone-icon"></i>
        <h3>Arraste o PDF de contrato aqui</h3>
        <p>ou <button class="btn-link" onclick="document.getElementById('ct-file').click()">selecione o arquivo</button></p>
        <p class="text-muted text-sm">PDF com texto selecionável · até 50 MB</p>
      </div>`;
  },

  bindDropzone() {
    const dz = document.getElementById('ct-dz');
    const fi = document.getElementById('ct-file');
    if (!dz || !fi) return;

    dz.addEventListener('click', () => fi.click());
    fi.addEventListener('change', () => { if (fi.files[0]) Contratos.processUpload(fi.files[0]); });

    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
    dz.addEventListener('drop', e => {
      e.preventDefault(); dz.classList.remove('drag-over');
      const f = e.dataTransfer.files[0];
      if (f && f.name.endsWith('.pdf')) Contratos.processUpload(f);
      else U.toast('Apenas arquivos PDF são aceitos.','error');
    });
  },

  async processUpload(file) {
    S.uploadContrato.file = file;
    S.uploadContrato.filename = file.name;
    const wbody = document.getElementById('ct-wbody');
    wbody.innerHTML = `<div class="loading-inline"><div class="spinner-sm"></div> Extraindo dados do PDF <b>${U.esc(file.name)}</b>...</div>`;
    try {
      const res = await API.upload('/api/contratos/upload-pdf', file);
      S.uploadContrato.rows     = res.extracted || [];
      S.uploadContrato.rawText  = res.raw_text  || '';
      this.setStep(2);
      wbody.innerHTML = this.uploadStep2();
      this.bindStep2();
    } catch(e) {
      wbody.innerHTML = `<div class="alert alert-error"><i class="fas fa-triangle-exclamation"></i> ${e.message}</div>${this.uploadStep1()}`;
      this.bindDropzone();
    }
  },

  setStep(n) {
    [1,2,3].forEach(i => {
      const el = document.getElementById(`ct-ws${i}`);
      if (!el) return;
      el.className = 'wizard-step' + (i < n ? ' done' : i === n ? ' active' : '');
    });
  },

  uploadStep2() {
    const rows = S.uploadContrato.rows;
    const rawText = S.uploadContrato.rawText || '';

    const nothingExtracted = rows.length === 0;
    const hint = nothingExtracted
      ? `<div class="alert alert-warning">
           <i class="fas fa-triangle-exclamation"></i>
           <div>
             <strong>Nenhum dado extraído automaticamente.</strong><br>
             O PDF pode usar fonte não selecionável (imagem/OCR) ou layout não reconhecível.
             Preencha os dados manualmente abaixo ou
             <button class="btn-link" onclick="Contratos.toggleRawText()">ver texto bruto extraído</button>
             para identificar o formato.
           </div>
         </div>
         <div id="ct-raw-wrap" style="display:none;margin-bottom:12px">
           <div style="font-size:12px;font-weight:600;color:var(--text-secondary);margin-bottom:4px">Texto bruto extraído do PDF (primeiros 3000 caracteres):</div>
           <textarea readonly style="width:100%;height:180px;font-family:monospace;font-size:11px;border:1px solid var(--border);border-radius:6px;padding:8px;background:#f8fafc;resize:vertical">${U.esc(rawText || '(nenhum texto encontrado — PDF pode ser escaneado/imagem)')}</textarea>
         </div>` : '';

    const tableRows = rows.map((r, i) => `
      <tr id="ct-row-${i}">
        <td><input name="numero_contrato" value="${U.esc(r.numero_contrato||'')}" placeholder="Nº contrato"></td>
        <td><input name="linha_telefone" value="${U.esc(r.linha_telefone||'')}" placeholder="(11)99999-9999" required></td>
        <td><input type="number" name="valor_contratado" value="${r.valor_contratado??''}" step="0.01" placeholder="0.00" required></td>
        <td><input name="cliente" value="${U.esc(r.cliente||'')}" placeholder="Cliente"></td>
        <td><input name="operadora" value="${U.esc(r.operadora||'')}" placeholder="Operadora"></td>
        <td><button type="button" class="btn-remove-row" title="Remover linha" onclick="Contratos.removePreviewRow(${i})"><i class="fas fa-trash"></i></button></td>
      </tr>`).join('');

    return `
      <div class="alert alert-info" style="margin-bottom:12px">
        <i class="fas fa-circle-info"></i>
        <span>Arquivo: <strong>${U.esc(S.uploadContrato.filename)}</strong> · <strong>${rows.length}</strong> linha(s) extraída(s).
        Revise os dados antes de salvar.</span>
      </div>
      ${hint}
      <form id="ct-preview-form">
        <div class="preview-table-wrap">
          <table class="preview-table" id="ct-preview-table">
            <thead><tr><th>Nº Contrato</th><th>Linha *</th><th>Valor Contratado *</th><th>Cliente</th><th>Operadora</th><th></th></tr></thead>
            <tbody>${tableRows}</tbody>
          </table>
        </div>
        <button type="button" class="btn-add-row" onclick="Contratos.addPreviewRow()"><i class="fas fa-plus"></i> Adicionar linha manualmente</button>
      </form>
      <div class="form-actions" style="margin-top:16px">
        <button class="btn btn-ghost" onclick="Contratos.goStep1()"><i class="fas fa-arrow-left"></i> Voltar</button>
        <button class="btn btn-primary" onclick="Contratos.savePreview()"><i class="fas fa-check"></i> Salvar Contratos</button>
      </div>`;
  },

  bindStep2() {},

  toggleRawText() {
    const el = document.getElementById('ct-raw-wrap');
    if (el) el.style.display = el.style.display === 'none' ? 'block' : 'none';
  },

  goStep1() {
    S.uploadContrato.rows = [];
    this.setStep(1);
    document.getElementById('ct-wbody').innerHTML = this.uploadStep1();
    this.bindDropzone();
  },

  addPreviewRow() {
    S.uploadContrato.rows.push({ numero_contrato:'', linha_telefone:'', valor_contratado:'', cliente:'', operadora:'' });
    const wb = document.getElementById('ct-wbody');
    if (wb) { this.setStep(2); wb.innerHTML = this.uploadStep2(); }
  },

  removePreviewRow(i) {
    S.uploadContrato.rows.splice(i, 1);
    const wb = document.getElementById('ct-wbody');
    if (wb) { wb.innerHTML = this.uploadStep2(); }
  },

  readPreviewForm() {
    const tbody = document.querySelector('#ct-preview-table tbody');
    if (!tbody) return [];
    return Array.from(tbody.rows).map(row => ({
      numero_contrato:  row.querySelector('[name="numero_contrato"]')?.value.trim() || '',
      linha_telefone:   row.querySelector('[name="linha_telefone"]')?.value.trim()  || '',
      valor_contratado: row.querySelector('[name="valor_contratado"]')?.value       || '',
      cliente:          row.querySelector('[name="cliente"]')?.value.trim()          || '',
      operadora:        row.querySelector('[name="operadora"]')?.value.trim()        || '',
      arquivo_pdf_origem: S.uploadContrato.filename,
    })).filter(r => r.linha_telefone);
  },

  async savePreview() {
    const contratos = this.readPreviewForm();
    if (!contratos.length) { U.toast('Nenhuma linha válida para salvar.','warning'); return; }
    try {
      U.loading(true, `Salvando ${contratos.length} contrato(s)...`);
      const res = await API.post('/api/contratos/salvar-lote', { contratos, arquivo_nome: S.uploadContrato.filename });
      U.loading(false);
      this.setStep(3);
      document.getElementById('ct-wbody').innerHTML = `
        <div class="result-panel">
          <span class="result-icon">🎉</span>
          <h3>${res.saved} contrato(s) salvo(s) com sucesso!</h3>
          ${res.errors?.length ? `<p style="color:var(--danger-text)">${res.errors.length} erro(s): ${res.errors.slice(0,3).join(', ')}</p>` : ''}
          <div class="form-actions" style="justify-content:center;margin-top:20px">
            <button class="btn btn-ghost" onclick="U.closeModal()">Fechar</button>
            <button class="btn btn-primary" onclick="U.closeModal();Contratos.fetch()"><i class="fas fa-table-list"></i> Ver Contratos</button>
          </div>
        </div>`;
      U.toast(`${res.saved} contrato(s) importado(s)!`);
    } catch(e) { U.loading(false); U.toast(e.message,'error'); }
  },
};

// -------------------------------------------------------
// UPLOAD CONTAS
// -------------------------------------------------------
const UploadContas = {
  render() {
    S.uploadConta = { step: 1, file: null, rows: [], filename: '', result: null };
    const sec = document.getElementById('page-upload-contas');
    sec.innerHTML = this.shell();
    this.bindDropzone();
  },

  shell() {
    return `
      <div class="page-header">
        <h1>Upload de Contas Telefônicas</h1>
      </div>
      <div class="wizard-steps" id="uc-wsteps">
        <div class="wizard-step active" id="uc-ws1"><span>1</span> Selecionar PDF</div>
        <div class="wizard-step" id="uc-ws2"><span>2</span> Revisar Dados</div>
        <div class="wizard-step" id="uc-ws3"><span>3</span> Resultado</div>
      </div>
      <div id="uc-body">${this.step1()}</div>`;
  },

  step1() {
    return `
      <div style="max-width:600px;margin:0 auto">
        <div class="dropzone" id="uc-dz">
          <input type="file" id="uc-file" accept=".pdf" style="display:none">
          <i class="fas fa-file-invoice dropzone-icon" style="color:var(--primary)"></i>
          <h3>Arraste a conta/fatura telefônica aqui</h3>
          <p>ou <button class="btn-link" onclick="document.getElementById('uc-file').click()">selecione o arquivo PDF</button></p>
          <p class="text-muted text-sm">O sistema irá extrair automaticamente as linhas e valores · PDF até 100 MB</p>
        </div>
        <div class="alert alert-info" style="margin-top:14px">
          <i class="fas fa-circle-info"></i>
          <span>Após o upload, você poderá revisar e corrigir os dados extraídos antes de confirmar.
          A conferência com os contratos será feita automaticamente.</span>
        </div>
      </div>`;
  },

  bindDropzone() {
    const dz = document.getElementById('uc-dz');
    const fi = document.getElementById('uc-file');
    if (!dz || !fi) return;

    dz.addEventListener('click', () => fi.click());
    fi.addEventListener('change', () => { if (fi.files[0]) UploadContas.processUpload(fi.files[0]); });
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
    dz.addEventListener('drop', e => {
      e.preventDefault(); dz.classList.remove('drag-over');
      const f = e.dataTransfer.files[0];
      if (f && f.name.endsWith('.pdf')) UploadContas.processUpload(f);
      else U.toast('Apenas arquivos PDF são aceitos.','error');
    });
  },

  setStep(n) {
    [1,2,3].forEach(i => {
      const el = document.getElementById(`uc-ws${i}`);
      if (!el) return;
      el.className = 'wizard-step' + (i < n ? ' done' : i === n ? ' active' : '');
    });
  },

  async processUpload(file) {
    S.uploadConta.file = file;
    S.uploadConta.filename = file.name;
    const MAX_CHUNK = 3.5 * 1024 * 1024;

    if (file.size <= MAX_CHUNK) {
      // Upload direto
      document.getElementById('uc-body').innerHTML = `<div class="loading-inline"><div class="spinner-sm"></div> Extraindo dados da conta <b>${U.esc(file.name)}</b>...</div>`;
      try {
        const res = await API.upload('/api/contas/upload-pdf', file);
        S.uploadConta.rows     = res.extracted || [];
        S.uploadConta.filename = res.filename  || file.name;
        S.uploadConta.rawText  = res.raw_text  || '';
        this.setStep(2);
        document.getElementById('uc-body').innerHTML = this.step2();
      } catch(e) {
        document.getElementById('uc-body').innerHTML = `<div class="alert alert-error" style="max-width:600px;margin:0 auto"><i class="fas fa-triangle-exclamation"></i> ${e.message}</div>` + `<div style="max-width:600px;margin:12px auto">${this.step1()}</div>`;
        this.bindDropzone();
      }
    } else {
      // Dividir PDF grande em partes
      document.getElementById('uc-body').innerHTML = `<div class="loading-inline"><div class="spinner-sm"></div> Dividindo <b>${U.esc(file.name)}</b> (${(file.size/1024/1024).toFixed(1)} MB)...</div>`;
      try {
        const arrayBuf = await file.arrayBuffer();
        const pdfDoc = await PDFLib.PDFDocument.load(arrayBuf);
        const totalPages = pdfDoc.getPageCount();
        const avgPageSize = file.size / totalPages;
        const pagesPerChunk = Math.max(1, Math.floor(3.5 * 1024 * 1024 / avgPageSize));

        const chunks = [];
        for (let start = 0; start < totalPages; start += pagesPerChunk) {
          chunks.push({ start, end: Math.min(start + pagesPerChunk, totalPages) });
        }

        let allExtracted = [];
        for (let i = 0; i < chunks.length; i++) {
          const { start, end } = chunks[i];
          document.getElementById('uc-body').innerHTML = `<div class="loading-inline"><div class="spinner-sm"></div> Parte ${i+1}/${chunks.length} (pág. ${start+1}-${end})...</div>
            <div style="background:var(--border);border-radius:6px;height:6px;margin:12px auto;max-width:400px;overflow:hidden">
              <div style="background:var(--primary);height:100%;width:${Math.round(((i+1)/chunks.length)*100)}%;transition:width .3s"></div>
            </div>`;

          const chunkPdf = await PDFLib.PDFDocument.create();
          const pages = await chunkPdf.copyPages(pdfDoc, Array.from({length: end-start}, (_,k) => start+k));
          pages.forEach(p => chunkPdf.addPage(p));
          const chunkBytes = await chunkPdf.save();
          const blob = new Blob([chunkBytes], {type:'application/pdf'});
          const chunkFile = new File([blob], `${file.name}_p${i+1}.pdf`, {type:'application/pdf'});

          const res = await API.upload('/api/contas/upload-pdf', chunkFile);
          if (res.extracted) allExtracted = allExtracted.concat(res.extracted);
        }

        S.uploadConta.rows = allExtracted;
        S.uploadConta.rawText = '';
        this.setStep(2);
        document.getElementById('uc-body').innerHTML = this.step2();
        U.toast(`PDF dividido em ${chunks.length} partes. ${allExtracted.length} linhas extraídas.`);
      } catch(e) {
        document.getElementById('uc-body').innerHTML = `<div class="alert alert-error" style="max-width:600px;margin:0 auto"><i class="fas fa-triangle-exclamation"></i> ${e.message}</div>` + `<div style="max-width:600px;margin:12px auto">${this.step1()}</div>`;
        this.bindDropzone();
      }
    }
  },

  step2() {
    const rows = S.uploadConta.rows;
    const rawText = S.uploadConta.rawText || '';

    const nothingExtracted = rows.length === 0;
    const hint = nothingExtracted
      ? `<div class="alert alert-warning">
           <i class="fas fa-triangle-exclamation"></i>
           <div>
             <strong>Nenhum dado extraído automaticamente.</strong><br>
             O PDF pode ser escaneado (imagem). Preencha manualmente ou
             <button class="btn-link" onclick="UploadContas.toggleRawText()">ver texto bruto extraído</button>.
           </div>
         </div>
         <div id="uc-raw-wrap" style="display:none;margin-bottom:12px">
           <div style="font-size:12px;font-weight:600;color:var(--text-secondary);margin-bottom:4px">Texto bruto extraído do PDF:</div>
           <textarea readonly style="width:100%;height:180px;font-family:monospace;font-size:11px;border:1px solid var(--border);border-radius:6px;padding:8px;background:#f8fafc;resize:vertical">${U.esc(rawText || '(nenhum texto encontrado)')}</textarea>
         </div>` : '';

    const tableRows = rows.map((r, i) => `
      <tr id="uc-row-${i}">
        <td><input name="linha_telefone" value="${U.esc(r.linha_telefone||'')}" placeholder="(11)99999-9999" required></td>
        <td><input type="number" name="valor_fatura" value="${r.valor_fatura??''}" step="0.01" placeholder="0.00" required></td>
        <td><input name="competencia" value="${U.esc(r.competencia||'')}" placeholder="2025-01"></td>
        <td><input name="operadora" value="${U.esc(r.operadora||'')}" placeholder="Vivo"></td>
        <td><input name="numero_fatura" value="${U.esc(r.numero_fatura||'')}" placeholder="Nº da fatura"></td>
        <td><button type="button" class="btn-remove-row" title="Remover" onclick="UploadContas.removeRow(${i})"><i class="fas fa-trash"></i></button></td>
      </tr>`).join('');

    return `
      <div class="alert alert-info">
        <i class="fas fa-circle-info"></i>
        <span>Arquivo: <strong>${U.esc(S.uploadConta.filename)}</strong> · <strong>${rows.length}</strong> linha(s) extraída(s).
        Revise os dados antes de confirmar.</span>
      </div>
      ${hint}
      <div class="preview-table-wrap">
        <table class="preview-table" id="uc-preview-table">
          <thead><tr><th>Linha *</th><th>Valor Fatura *</th><th>Competência</th><th>Operadora</th><th>Nº Fatura</th><th></th></tr></thead>
          <tbody>${tableRows}</tbody>
        </table>
      </div>
      <button type="button" class="btn-add-row" onclick="UploadContas.addRow()"><i class="fas fa-plus"></i> Adicionar linha manualmente</button>
      <div class="form-actions" style="margin-top:16px">
        <button class="btn btn-ghost" onclick="UploadContas.render()"><i class="fas fa-arrow-left"></i> Voltar</button>
        <button class="btn btn-primary" onclick="UploadContas.save()"><i class="fas fa-scale-balanced"></i> Confirmar e Processar Conferência</button>
      </div>`;
  },

  toggleRawText() {
    const el = document.getElementById('uc-raw-wrap');
    if (el) el.style.display = el.style.display === 'none' ? 'block' : 'none';
  },

  addRow() {
    S.uploadConta.rows.push({ linha_telefone:'', valor_fatura:'', competencia:'', operadora:'', numero_fatura:'' });
    document.getElementById('uc-body').innerHTML = this.step2();
  },

  removeRow(i) {
    S.uploadConta.rows.splice(i, 1);
    document.getElementById('uc-body').innerHTML = this.step2();
  },

  readForm() {
    const tbody = document.querySelector('#uc-preview-table tbody');
    if (!tbody) return [];
    return Array.from(tbody.rows).map(row => ({
      linha_telefone: row.querySelector('[name="linha_telefone"]')?.value.trim() || '',
      valor_fatura:   row.querySelector('[name="valor_fatura"]')?.value   || '',
      competencia:    row.querySelector('[name="competencia"]')?.value.trim() || '',
      operadora:      row.querySelector('[name="operadora"]')?.value.trim()   || '',
      numero_fatura:  row.querySelector('[name="numero_fatura"]')?.value.trim() || '',
      arquivo_pdf_origem: S.uploadConta.filename,
    })).filter(r => r.linha_telefone);
  },

  async save() {
    const contas = this.readForm();
    if (!contas.length) { U.toast('Nenhuma linha válida para salvar.','warning'); return; }
    try {
      U.loading(true, `Salvando ${contas.length} conta(s) e processando conferência...`);
      const res = await API.post('/api/contas/salvar-lote', { contas, arquivo_nome: S.uploadConta.filename });
      U.loading(false);
      S.uploadConta.result = res;
      this.setStep(3);
      document.getElementById('uc-body').innerHTML = this.step3(res, contas.length);
    } catch(e) { U.loading(false); U.toast(e.message,'error'); }
  },

  step3(res, total) {
    const comp = res.comparacao || {};
    return `
      <div class="result-panel">
        <span class="result-icon">✅</span>
        <h3>${res.saved} conta(s) importada(s) com sucesso!</h3>
        <p>A conferência automática foi realizada contra a base de contratos.</p>
        ${res.errors?.length ? `<div class="alert alert-warning"><i class="fas fa-triangle-exclamation"></i>${res.errors.length} linha(s) com erro: ${res.errors.slice(0,3).join(', ')}</div>` : ''}
        <div class="result-stats">
          <div class="result-stat" style="border-color:var(--success-border);background:var(--success-bg)">
            <strong style="color:var(--success-text)">${comp.ok||0}</strong><span>Conformes</span>
          </div>
          <div class="result-stat" style="border-color:var(--warning-border);background:var(--warning-bg)">
            <strong style="color:var(--warning-text)">${comp.aproximado||0}</strong><span>Aproximados</span>
          </div>
          <div class="result-stat" style="border-color:var(--danger-border);background:var(--danger-bg)">
            <strong style="color:var(--danger-text)">${comp.divergente||0}</strong><span>Divergentes</span>
          </div>
          <div class="result-stat" style="border-color:var(--info-border);background:var(--info-bg)">
            <strong style="color:var(--info-text)">${comp.sem_contrato||0}</strong><span>Sem Contrato</span>
          </div>
        </div>
        <div class="form-actions" style="justify-content:center">
          <button class="btn btn-ghost" onclick="UploadContas.render()"><i class="fas fa-upload"></i> Nova Importação</button>
          <button class="btn btn-primary" onclick="navigate('comparacoes')"><i class="fas fa-scale-balanced"></i> Ver Conferência</button>
        </div>
      </div>`;
  },
};

// -------------------------------------------------------
// COMPARAÇÕES
// -------------------------------------------------------
const Comparacoes = {
  async load() {
    const sec = document.getElementById('page-comparacoes');
    sec.innerHTML = this.shell();

    // Bind search
    const si = document.getElementById('cmp-search');
    if (si) si.addEventListener('input', U.debounce(() => {
      S.comparacoes.search = si.value;
      S.comparacoes.page = 1;
      Comparacoes.fetch();
    }, 380));

    // Load filter options
    try {
      const [comps, ops] = await Promise.all([API.get('/api/competencias'), API.get('/api/operadoras')]);
      const selComp = document.getElementById('cmp-comp');
      if (selComp) comps.forEach(c => { const o = document.createElement('option'); o.value=c; o.textContent=c; if(c===S.comparacoes.competencia)o.selected=true; selComp.appendChild(o); });
      const selOp = document.getElementById('cmp-op');
      if (selOp) ops.forEach(o => { const el = document.createElement('option'); el.value=o; el.textContent=o; if(o===S.comparacoes.operadora)el.selected=true; selOp.appendChild(el); });
    } catch(e) { /* ignore */ }

    await this.fetch();
  },

  shell() {
    return `
      <div class="page-header">
        <h1>Conferência de Contas</h1>
        <div class="page-header-actions">
          <button class="btn btn-outline" onclick="Comparacoes.fetch()"><i class="fas fa-rotate-right"></i> Atualizar</button>
          <button class="btn btn-success" onclick="Comparacoes.runManual()"><i class="fas fa-play"></i> Reprocessar</button>
        </div>
      </div>

      <div class="filter-bar">
        <div class="search-field">
          <i class="fas fa-search"></i>
          <input id="cmp-search" type="text" placeholder="Buscar por linha ou contrato..." value="${U.esc(S.comparacoes.search)}">
        </div>
        <select id="cmp-comp" class="select-filter" onchange="Comparacoes.filterChange('competencia',this.value)">
          <option value="">Todas as competências</option>
        </select>
        <select id="cmp-op" class="select-filter" onchange="Comparacoes.filterChange('operadora',this.value)">
          <option value="">Todas as operadoras</option>
        </select>
        <button class="btn btn-ghost" onclick="Comparacoes.clearFilters()"><i class="fas fa-filter-circle-xmark"></i> Limpar</button>
      </div>

      <div id="cmp-stats" class="stats-row"></div>
      <div id="cmp-body"><div class="loading-inline"><div class="spinner-sm"></div> Carregando...</div></div>`;
  },

  filterChange(field, val) {
    S.comparacoes[field] = val;
    S.comparacoes.page = 1;
    this.fetch();
  },

  clearFilters() {
    S.comparacoes = { page: 1, status: '', competencia: '', operadora: '', search: '' };
    const si = document.getElementById('cmp-search'); if(si) si.value = '';
    const sc = document.getElementById('cmp-comp');  if(sc) sc.value = '';
    const so = document.getElementById('cmp-op');    if(so) so.value = '';
    this.fetch();
  },

  setStatusFilter(s) {
    S.comparacoes.status = s;
    S.comparacoes.page = 1;
    // Update pills
    document.querySelectorAll('.stat-pill').forEach(p => {
      p.classList.toggle('active', p.dataset.status === s || (s === '' && p.dataset.status === 'all'));
    });
    this.fetch();
  },

  async fetch() {
    const body = document.getElementById('cmp-body');
    if (!body) return;
    body.innerHTML = '<div class="loading-inline"><div class="spinner-sm"></div> Carregando conferências...</div>';
    try {
      const url = `/api/comparacoes?page=${S.comparacoes.page}&per_page=100`
        + `&status=${encodeURIComponent(S.comparacoes.status)}`
        + `&competencia=${encodeURIComponent(S.comparacoes.competencia)}`
        + `&operadora=${encodeURIComponent(S.comparacoes.operadora)}`
        + `&search=${encodeURIComponent(S.comparacoes.search)}`;
      const res = await API.get(url);
      this.renderStats(res.data);
      body.innerHTML = this.table(res);
    } catch(e) {
      body.innerHTML = `<div class="alert alert-error"><i class="fas fa-triangle-exclamation"></i> ${e.message}</div>`;
    }
  },

  renderStats(data) {
    const counts = { all: data.length, ok: 0, aproximado: 0, divergente: 0, sem_contrato: 0, sem_fatura: 0, ambiguo: 0 };
    data.forEach(r => { if (counts[r.status_comparacao] !== undefined) counts[r.status_comparacao]++; });

    const statsEl = document.getElementById('cmp-stats');
    if (!statsEl) return;
    statsEl.innerHTML = `
      <div class="stat-pill stat-all  ${S.comparacoes.status===''?'active':''}" data-status="all"  onclick="Comparacoes.setStatusFilter('')"><strong>${counts.all}</strong><span>Todos</span></div>
      <div class="stat-pill stat-ok   ${S.comparacoes.status==='ok'?'active':''}" data-status="ok"            onclick="Comparacoes.setStatusFilter('ok')"><strong>${counts.ok}</strong><span>Conformes</span></div>
      <div class="stat-pill stat-aprx ${S.comparacoes.status==='aproximado'?'active':''}" data-status="aproximado"  onclick="Comparacoes.setStatusFilter('aproximado')"><strong>${counts.aproximado}</strong><span>Aproximados</span></div>
      <div class="stat-pill stat-div  ${S.comparacoes.status==='divergente'?'active':''}" data-status="divergente"  onclick="Comparacoes.setStatusFilter('divergente')"><strong>${counts.divergente}</strong><span>Divergentes</span></div>
      <div class="stat-pill stat-sc   ${S.comparacoes.status==='sem_contrato'?'active':''}" data-status="sem_contrato" onclick="Comparacoes.setStatusFilter('sem_contrato')"><strong>${counts.sem_contrato}</strong><span>Sem Contrato</span></div>
      <div class="stat-pill stat-sf   ${S.comparacoes.status==='sem_fatura'?'active':''}" data-status="sem_fatura"   onclick="Comparacoes.setStatusFilter('sem_fatura')"><strong>${counts.sem_fatura}</strong><span>Sem Fatura</span></div>
    `;
  },

  table(res) {
    if (!res.data.length) {
      if (S.comparacoes.status || S.comparacoes.competencia || S.comparacoes.search) {
        return U.emptyState('fa-filter', 'Nenhum resultado para esse filtro', 'Tente remover os filtros para ver todos os registros.',
          `<button class="btn btn-ghost" onclick="Comparacoes.clearFilters()"><i class="fas fa-filter-circle-xmark"></i> Limpar Filtros</button>`);
      }
      return U.emptyState('fa-scale-balanced', 'Nenhuma conferência realizada ainda',
        'Importe contratos e faça upload de contas para iniciar a auditoria.',
        `<button class="btn btn-outline" onclick="navigate('contratos')"><i class="fas fa-file-signature"></i> Contratos</button>
         <button class="btn btn-primary" style="margin-left:8px" onclick="navigate('upload-contas')"><i class="fas fa-upload"></i> Importar Conta</button>`);
    }

    const rows = res.data.map(c => `
      <tr class="row-${c.status_comparacao} row-clickable" onclick="Comparacoes.openDetail(${c.id})">
        <td>${U.statusBadge(c.status_comparacao)}</td>
        <td><span class="phone-number">${U.phone(c.linha_telefone)}</span></td>
        <td><span class="mono">${U.esc(c.numero_contrato||'—')}</span></td>
        <td class="text-right">${c.valor_contratado !== null && c.valor_contratado !== undefined ? U.money(c.valor_contratado) : '—'}</td>
        <td class="text-right">${c.valor_fatura !== null && c.valor_fatura !== undefined ? U.money(c.valor_fatura) : '—'}</td>
        <td class="text-right ${U.diffClass(c.diferenca_valor)}">${c.diferenca_valor !== null && c.diferenca_valor !== undefined ? (c.diferenca_valor>0?'+':'')+U.money(c.diferenca_valor) : '—'}</td>
        <td class="text-right ${U.diffClass(c.diferenca_percentual)}">${c.diferenca_percentual !== null && c.diferenca_percentual !== undefined ? (c.diferenca_percentual>0?'+':'')+U.pct(c.diferenca_percentual) : '—'}</td>
        <td>${c.competencia ? `<span class="comp-tag">${U.esc(c.competencia)}</span>` : '—'}</td>
        <td>${c.operadora ? `<span class="chip">${U.esc(c.operadora)}</span>` : '—'}</td>
        <td class="text-muted text-sm" title="${U.esc(c.observacao||'')}">${U.esc((c.observacao||'').slice(0,40)) + ((c.observacao||'').length>40?'…':'')}</td>
      </tr>`).join('');

    const pages = Math.ceil(res.total / 100);
    const pag = pages > 1 ? `
      <div class="pagination">
        <button class="btn-page" onclick="Comparacoes.goPage(${S.comparacoes.page-1})" ${S.comparacoes.page<=1?'disabled':''}><i class="fas fa-chevron-left"></i></button>
        <span>Página ${S.comparacoes.page} de ${pages} &nbsp;·&nbsp; ${res.total} registros</span>
        <button class="btn-page" onclick="Comparacoes.goPage(${S.comparacoes.page+1})" ${S.comparacoes.page>=pages?'disabled':''}><i class="fas fa-chevron-right"></i></button>
      </div>` : `<div class="table-footer-info">${res.total} comparação(ões)</div>`;

    return `
      <div class="table-wrapper">
        <table class="data-table">
          <thead><tr>
            <th>Status</th><th>Linha</th><th>Contrato</th>
            <th class="text-right">V. Contratado</th><th class="text-right">V. Fatura</th>
            <th class="text-right">Dif. R$</th><th class="text-right">Dif. %</th>
            <th>Competência</th><th>Operadora</th><th>Observação</th>
          </tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>${pag}`;
  },

  goPage(p) { if(p<1) return; S.comparacoes.page=p; this.fetch(); },

  async openDetail(id) {
    try {
      const [comp, hist] = await Promise.all([
        API.get(`/api/comparacoes?page=1&per_page=1&search=`).then(() =>
          // fetch single through list filtered  — ou abrir simples via historico por id
          API.get(`/api/comparacoes/${id}/historico`).then(h => ({ hist: h }))
        ).catch(() => ({})),
        API.get(`/api/comparacoes/${id}/historico`),
      ]);
      U.modal(this.detailHtml(id, hist), 'lg');
      // Render mini history chart
      setTimeout(() => this.renderHistChart(hist), 100);
    } catch(e) {
      U.toast(e.message, 'error');
    }
  },

  detailHtml(id, hist) {
    const latest = hist[0] || {};
    const diffCls = latest.diferenca_valor > 0.005 ? 'diff-div-pos' : latest.diferenca_valor < -0.005 ? 'diff-div-neg' :
                    (latest.status_comparacao === 'aproximado' ? 'diff-aprox' : 'diff-ok');
    return `
      <h2 class="modal-title"><i class="fas fa-scale-balanced" style="color:var(--primary)"></i>Detalhe da Conferência</h2>
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:16px">
        ${U.statusBadge(latest.status_comparacao)}
        <span class="phone-number" style="font-size:16px">${U.phone(latest.linha_telefone)}</span>
        ${latest.competencia ? `<span class="comp-tag">${U.esc(latest.competencia)}</span>` : ''}
      </div>

      <div class="detail-grid">
        <div class="detail-section">
          <h4><i class="fas fa-file-signature"></i> Contrato</h4>
          <div class="detail-field"><label>Nº Contrato</label><span>${U.esc(latest.numero_contrato||'—')}</span></div>
          <div class="detail-field"><label>Valor Contratado</label><span>${U.money(latest.valor_contratado)}</span></div>
          <div class="detail-field"><label>Operadora</label><span>${U.esc(latest.operadora||'—')}</span></div>
        </div>
        <div class="detail-section">
          <h4><i class="fas fa-file-invoice"></i> Fatura</h4>
          <div class="detail-field"><label>Valor Fatura</label><span>${U.money(latest.valor_fatura)}</span></div>
          <div class="detail-field"><label>Competência</label><span>${U.esc(latest.competencia||'—')}</span></div>
        </div>
      </div>

      <div class="detail-section" style="margin-bottom:16px;text-align:center">
        <h4>Diferença Calculada</h4>
        <div class="diff-highlight ${diffCls}" style="margin-top:8px">
          ${latest.diferenca_valor !== null && latest.diferenca_valor !== undefined
            ? `${latest.diferenca_valor > 0 ? '+' : ''}${U.money(latest.diferenca_valor)} (${latest.diferenca_valor > 0 ? '+' : ''}${U.pct(latest.diferenca_percentual)})`
            : 'Não calculado'}
        </div>
        ${latest.observacao ? `<p style="font-size:13px;color:var(--text-secondary);margin-top:10px">${U.esc(latest.observacao)}</p>` : ''}
      </div>

      ${hist.length > 1 ? `
        <div class="detail-section">
          <h4><i class="fas fa-clock-rotate-left"></i> Histórico desta Linha (últimos ${hist.length} meses)</h4>
          <div class="history-chart-wrap"><canvas id="hist-chart"></canvas></div>
        </div>` : ''}

      <div class="form-actions">
        <button class="btn btn-ghost" onclick="U.closeModal()">Fechar</button>
      </div>`;
  },

  renderHistChart(hist) {
    const cv = document.getElementById('hist-chart');
    if (!cv || hist.length < 2) return;
    if (S.histDetailChart) { try { S.histDetailChart.destroy(); } catch(e){} }
    const items = [...hist].reverse();
    S.histDetailChart = new Chart(cv, {
      type: 'line',
      data: {
        labels: items.map(h => h.competencia || h.data_processamento?.slice(0,7) || ''),
        datasets: [
          { label: 'Valor Contratado', data: items.map(h => h.valor_contratado), borderColor: '#6366f1', backgroundColor: '#6366f122', fill: true, tension: 0.3, borderWidth: 2 },
          { label: 'Valor Fatura',     data: items.map(h => h.valor_fatura),     borderColor: '#ef4444', backgroundColor: '#ef444422', fill: true, tension: 0.3, borderWidth: 2 },
        ],
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { position: 'bottom', labels: { font: { family: 'Inter', size: 11 } } } },
        scales: { y: { beginAtZero: false } },
      },
    });
  },

  async runManual() {
    try {
      U.loading(true, 'Reprocessando conferências...');
      const res = await API.post('/api/comparar', { competencia: S.comparacoes.competencia || null });
      U.loading(false);
      U.toast(`Conferência concluída: ${res.totais?.ok||0} ok, ${res.totais?.divergente||0} divergentes.`);
      this.fetch();
    } catch(e) { U.loading(false); U.toast(e.message,'error'); }
  },
};

// -------------------------------------------------------
// HISTÓRICO
// -------------------------------------------------------
const Historico = {
  async load() {
    const sec = document.getElementById('page-historico');
    sec.innerHTML = `
      <div class="page-header">
        <h1>Histórico de Importações</h1>
        <div class="page-header-actions">
          <button class="btn btn-outline" onclick="Historico.load()"><i class="fas fa-rotate-right"></i> Atualizar</button>
        </div>
      </div>
      <div id="hist-body"><div class="loading-inline"><div class="spinner-sm"></div> Carregando...</div></div>`;
    try {
      const data = await API.get('/api/historico');
      document.getElementById('hist-body').innerHTML = this.table(data);
    } catch(e) {
      document.getElementById('hist-body').innerHTML = `<div class="alert alert-error">${e.message}</div>`;
    }
  },

  table(data) {
    if (!data.length) return U.emptyState('fa-clock-rotate-left', 'Nenhuma importação registrada', 'As importações de contratos e contas aparecerão aqui.');
    const rows = data.map(i => `
      <tr>
        <td>${i.tipo === 'contrato'
          ? '<span class="badge badge-ok"><i class="fas fa-file-signature"></i> Contrato</span>'
          : '<span class="badge badge-sem_contrato"><i class="fas fa-file-invoice"></i> Conta</span>'}</td>
        <td title="${U.esc(i.arquivo_nome||'')}">${U.esc((i.arquivo_nome||'Importação manual').slice(0,50))}</td>
        <td class="text-muted">${U.esc(i.data_importacao||'—')}</td>
        <td><strong>${i.total_registros}</strong></td>
        <td><span class="badge ${i.status==='concluido'?'badge-ok':'badge-divergente'}">${U.esc(i.status)}</span></td>
        <td class="text-muted text-sm">${U.esc((i.observacoes||'').slice(0,60))}</td>
        <td>
          <button class="btn-icon-sm btn-danger" title="Excluir importação" onclick="Historico.confirmDelete(${i.id})"><i class="fas fa-trash"></i></button>
        </td>
      </tr>`).join('');

    return `
      <div class="historico-table-wrap">
        <table class="data-table">
          <thead><tr><th>Tipo</th><th>Arquivo</th><th>Data</th><th>Registros</th><th>Status</th><th>Observações</th><th>Ação</th></tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
      <div class="table-footer-info">${data.length} importação(ões)</div>`;
  },

  confirmDelete(id) {
    U.modal(`
      <div class="confirm-dialog">
        <div class="confirm-icon"><i class="fas fa-triangle-exclamation text-danger"></i></div>
        <h3>Excluir importação?</h3>
        <p>As contas e comparações relacionadas a esta importação também serão excluídas permanentemente. Esta ação não pode ser desfeita.</p>
        <div class="form-actions">
          <button class="btn btn-ghost" onclick="U.closeModal()">Cancelar</button>
          <button class="btn btn-danger" onclick="Historico.doDelete(${id})"><i class="fas fa-trash"></i> Excluir</button>
        </div>
      </div>`);
  },

  async doDelete(id) {
    try {
      U.loading(true,'Excluindo...');
      await API.del(`/api/historico/${id}`);
      U.toast('Importação excluída.','info');
      U.closeModal(); U.loading(false); Historico.load();
    } catch(e) { U.loading(false); U.toast(e.message,'error'); }
  },
};

// -------------------------------------------------------
// RELATÓRIOS
// -------------------------------------------------------
const Relatorios = {
  async render() {
    const sec = document.getElementById('page-relatorios');
    sec.innerHTML = `
      <div class="page-header"><h1>Relatórios e Exportação</h1></div>
      <div id="rel-body"><div class="loading-inline"><div class="spinner-sm"></div></div></div>`;
    try {
      const [comps, ops] = await Promise.all([API.get('/api/competencias'), API.get('/api/operadoras')]);
      document.getElementById('rel-body').innerHTML = this.html(comps, ops);
    } catch(e) {
      document.getElementById('rel-body').innerHTML = this.html([], []);
    }
  },

  html(comps, ops) {
    const compOpts = comps.map(c => `<option value="${U.esc(c)}">${U.esc(c)}</option>`).join('');
    return `
      <div class="config-card" style="max-width:860px">
        <h3><i class="fas fa-filter" style="color:var(--primary)"></i> Filtros de Exportação (Comparações)</h3>
        <div class="form-row">
          <div class="form-group">
            <label>Competência</label>
            <select id="rel-comp" class="select-filter" style="width:100%">
              <option value="">Todas as competências</option>${compOpts}
            </select>
          </div>
          <div class="form-group">
            <label>Status</label>
            <select id="rel-status" class="select-filter" style="width:100%">
              <option value="">Todos os status</option>
              <option value="ok">Conformes</option>
              <option value="aproximado">Aproximados</option>
              <option value="divergente">Divergentes</option>
              <option value="sem_contrato">Sem Contrato</option>
              <option value="sem_fatura">Sem Fatura</option>
            </select>
          </div>
        </div>
        <div style="display:flex;gap:12px;flex-wrap:wrap;margin-top:12px">
          <button class="btn btn-success" onclick="Relatorios.exportar('excel')">
            <i class="fas fa-file-excel"></i> Exportar Excel (.xlsx)
          </button>
          <button class="btn btn-outline" onclick="Relatorios.exportar('csv')">
            <i class="fas fa-file-csv"></i> Exportar CSV
          </button>
        </div>
      </div>

      <div class="config-card" style="max-width:860px">
        <h3><i class="fas fa-calendar-days" style="color:#660099"></i> Filtrar Faturas por Data / Número</h3>
        <p style="font-size:13px;color:var(--text-secondary);margin-bottom:12px">
          Filtre as faturas importadas por período ou número da linha e exporte para Excel.
        </p>
        <div class="form-row">
          <div class="form-group">
            <label>Data Início</label>
            <input type="date" id="rel-fat-inicio" style="width:100%;padding:7px 10px;font-size:13px;border:1px solid var(--border);border-radius:6px">
          </div>
          <div class="form-group">
            <label>Data Fim</label>
            <input type="date" id="rel-fat-fim" style="width:100%;padding:7px 10px;font-size:13px;border:1px solid var(--border);border-radius:6px">
          </div>
          <div class="form-group">
            <label>Número da Linha</label>
            <input type="text" id="rel-fat-numero" placeholder="Ex: 11999887766" style="width:100%;padding:7px 10px;font-size:13px;border:1px solid var(--border);border-radius:6px">
          </div>
        </div>
        <div style="display:flex;gap:12px;flex-wrap:wrap;margin-top:12px">
          <button class="btn btn-primary" onclick="Relatorios.filtrarFaturas()">
            <i class="fas fa-search"></i> Filtrar
          </button>
          <button class="btn btn-success" onclick="Relatorios.exportarFaturas()">
            <i class="fas fa-file-excel"></i> Exportar Faturas Excel
          </button>
        </div>
        <div id="rel-fat-result" style="margin-top:16px"></div>
      </div>

      <div class="config-card" style="max-width:860px">
        <h3><i class="fas fa-chart-column" style="color:var(--primary)"></i> Relatórios Rápidos</h3>
        <div class="quick-actions-grid" style="margin-top:0">
          <button class="quick-btn" onclick="Relatorios.filtrarE('divergente')">
            <i class="fas fa-circle-xmark" style="color:var(--danger)"></i>
            <span>Linhas Divergentes</span>
          </button>
          <button class="quick-btn" onclick="Relatorios.filtrarE('ok')">
            <i class="fas fa-circle-check" style="color:var(--success)"></i>
            <span>Conformes</span>
          </button>
          <button class="quick-btn" onclick="Relatorios.filtrarE('sem_contrato')">
            <i class="fas fa-circle-question" style="color:var(--info)"></i>
            <span>Sem Contrato</span>
          </button>
          <button class="quick-btn" onclick="Relatorios.filtrarE('sem_fatura')">
            <i class="fas fa-circle-minus" style="color:var(--gray)"></i>
            <span>Sem Fatura</span>
          </button>
        </div>
      </div>`;
  },

  filtrarE(status) {
    S.comparacoes.status = status;
    S.comparacoes.page = 1;
    navigate('comparacoes');
  },

  exportar(tipo) {
    const comp   = document.getElementById('rel-comp')?.value   || '';
    const status = document.getElementById('rel-status')?.value || '';
    const url = `/api/exportar/${tipo}?competencia=${encodeURIComponent(comp)}&status=${encodeURIComponent(status)}`;
    const a = document.createElement('a');
    a.href = url; a.target = '_blank';
    document.body.appendChild(a); a.click(); a.remove();
    U.toast(`Download do relatório ${tipo.toUpperCase()} iniciado.`);
  },

  async filtrarFaturas() {
    const inicio = document.getElementById('rel-fat-inicio')?.value || '';
    const fim = document.getElementById('rel-fat-fim')?.value || '';
    const numero = document.getElementById('rel-fat-numero')?.value || '';
    const resultEl = document.getElementById('rel-fat-result');
    resultEl.innerHTML = '<div class="loading-inline"><div class="spinner-sm"></div> Buscando...</div>';

    try {
      const url = `/api/relatorios/fatura-linhas?data_inicio=${encodeURIComponent(inicio)}&data_fim=${encodeURIComponent(fim)}&numero=${encodeURIComponent(numero)}`;
      const items = await API.get(url);
      if (!items.length) {
        resultEl.innerHTML = '<p style="color:var(--text-secondary);font-size:13px">Nenhuma fatura encontrada com esses filtros.</p>';
        return;
      }
      const rows = items.map(fl => `<tr>
        <td>${U.phone(fl.numero_vivo)}</td>
        <td>${U.esc(fl.plano || '—')}</td>
        <td class="text-right">${U.money(fl.valor_fatura)}</td>
        <td class="text-right">${U.money(fl.valor_contrato)}</td>
        <td class="text-right ${fl.diferenca > 0.005 ? 'val-pos' : fl.diferenca < -0.005 ? 'val-neg' : 'val-zero'}">${U.money(fl.diferenca)}</td>
        <td>${fl.status === 'ok' ? '<span class="badge badge-ok">OK</span>' : fl.status === 'divergente' ? '<span class="badge badge-divergente">Divergente</span>' : '—'}</td>
        <td>${U.esc(fl.competencia || '—')}</td>
        <td style="font-size:11px">${U.esc(fl.data_importacao || '—')}</td>
      </tr>`).join('');

      resultEl.innerHTML = `
        <p style="font-size:13px;margin-bottom:8px"><strong>${items.length}</strong> faturas encontradas</p>
        <div class="table-wrapper">
          <table class="data-table" style="font-size:12px">
            <thead><tr>
              <th>Número</th><th>Plano</th><th class="text-right">Fatura</th><th class="text-right">Contrato</th>
              <th class="text-right">Diferença</th><th>Status</th><th>Competência</th><th>Data</th>
            </tr></thead>
            <tbody>${rows}</tbody>
          </table>
        </div>`;
    } catch(e) {
      resultEl.innerHTML = `<div class="alert alert-error">${e.message}</div>`;
    }
  },

  exportarFaturas() {
    const inicio = document.getElementById('rel-fat-inicio')?.value || '';
    const fim = document.getElementById('rel-fat-fim')?.value || '';
    const numero = document.getElementById('rel-fat-numero')?.value || '';
    const url = `/api/exportar/fatura-excel?data_inicio=${encodeURIComponent(inicio)}&data_fim=${encodeURIComponent(fim)}&numero=${encodeURIComponent(numero)}`;
    const a = document.createElement('a');
    a.href = url; a.target = '_blank';
    document.body.appendChild(a); a.click(); a.remove();
    U.toast('Download do Excel de faturas iniciado.');
  },
};

// -------------------------------------------------------
// CONFIGURAÇÕES
// -------------------------------------------------------
const Configuracoes = {
  async load() {
    const sec = document.getElementById('page-configuracoes');
    sec.innerHTML = `<div class="page-header"><h1>Configurações</h1></div><div id="cfg-body"><div class="loading-inline"><div class="spinner-sm"></div></div></div>`;
    try {
      const cfg = await API.get('/api/configuracoes');
      document.getElementById('cfg-body').innerHTML = this.html(cfg);
    } catch(e) {
      document.getElementById('cfg-body').innerHTML = `<div class="alert alert-error">${e.message}</div>`;
    }
  },

  html(cfg) {
    const tipo  = cfg.tolerancia_tipo  || 'percentual';
    const valor = cfg.tolerancia_valor || '5.0';
    const empresa = cfg.empresa_nome   || 'Auditoria Telecom';
    return `
      <div class="config-card">
        <h3><i class="fas fa-sliders" style="color:var(--primary)"></i> Regras de Tolerância</h3>
        <p style="font-size:13px;color:var(--text-secondary);margin-bottom:16px">
          Define quando um valor é considerado "aproximado" ao invés de "divergente".
          Valores dentro da tolerância aparecem em amarelo; fora, em vermelho.
        </p>
        <form id="cfg-form" onsubmit="Configuracoes.save(event)">
          <div class="form-row">
            <div class="form-group">
              <label>Tipo de Tolerância</label>
              <div class="radio-group">
                <label class="radio-label">
                  <input type="radio" name="tolerancia_tipo" value="percentual" ${tipo==='percentual'?'checked':''}>
                  <span>Percentual (%)</span>
                </label>
                <label class="radio-label">
                  <input type="radio" name="tolerancia_tipo" value="fixo" ${tipo==='fixo'?'checked':''}>
                  <span>Valor fixo (R$)</span>
                </label>
              </div>
            </div>
            <div class="form-group">
              <label id="tol-val-label">Valor da Tolerância (${tipo==='fixo'?'R$':'%'})</label>
              <input type="number" id="tolerancia_valor" name="tolerancia_valor" value="${U.esc(valor)}" step="0.1" min="0" placeholder="5.0">
              <small class="text-muted">Ex: 5 = ±5${tipo==='fixo'?'%':'%'}  ou  1 = ±R$ 1,00</small>
            </div>
          </div>

          <div style="border-top:1px solid var(--border);padding-top:16px;margin-top:8px">
            <h3 style="font-size:14px;font-weight:600;margin-bottom:12px"><i class="fas fa-building" style="color:var(--primary)"></i> Identificação da Empresa</h3>
            <div class="form-group" style="max-width:360px">
              <label>Nome da Empresa</label>
              <input type="text" name="empresa_nome" value="${U.esc(empresa)}" placeholder="Nome da empresa">
            </div>
          </div>

          <div class="form-actions" style="border-top:1px solid var(--border);margin-top:16px">
            <button type="button" class="btn btn-ghost" onclick="Configuracoes.load()">Cancelar</button>
            <button type="submit" class="btn btn-primary"><i class="fas fa-check"></i> Salvar Configurações</button>
          </div>
        </form>
      </div>

      <div class="config-card">
        <h3><i class="fas fa-info-circle" style="color:var(--primary)"></i> Sobre o Sistema</h3>
        <div class="detail-field" style="padding:8px 0;border-bottom:1px solid var(--border)"><label>Sistema</label><span>AuditoriaTel v1.0</span></div>
        <div class="detail-field" style="padding:8px 0;border-bottom:1px solid var(--border)"><label>Banco de dados</label><span>SQLite local</span></div>
        <div class="detail-field" style="padding:8px 0"><label>Extração de PDF</label><span>pdfminer.six + heurísticas</span></div>
      </div>

      <div class="config-card" style="border:2px solid var(--danger)">
        <h3 style="color:var(--danger)"><i class="fas fa-triangle-exclamation" style="margin-right:6px"></i>Zona de Perigo</h3>
        <p style="font-size:13px;color:var(--text-secondary);margin-bottom:16px">
          Esta ação irá <strong>apagar todos os dados</strong> do sistema: contratos, contas, comparações,
          faturas importadas, planos, cadastros e dados Coopernac. As configurações serão restauradas
          para os valores padrão. <strong>Esta ação é irreversível.</strong>
        </p>
        <button class="btn" style="background:var(--danger);color:#fff;border:none" onclick="Configuracoes.resetDB()">
          <i class="fas fa-trash-can"></i> Zerar Banco de Dados
        </button>
      </div>`;
  },

  async save(e) {
    e.preventDefault();
    const f = e.target;
    const tipo = f.querySelector('[name="tolerancia_tipo"]:checked')?.value || 'percentual';
    const data = {
      tolerancia_tipo:  tipo,
      tolerancia_valor: f.tolerancia_valor.value,
      empresa_nome:     f.empresa_nome.value.trim(),
    };
    try {
      U.loading(true, 'Salvando configurações...');
      await API.put('/api/configuracoes', data);
      U.loading(false);
      U.toast('Configurações salvas com sucesso!');
    } catch(err) { U.loading(false); U.toast(err.message,'error'); }
  },

  async resetDB() {
    // Dupla confirmação
    if (!confirm('ATENÇÃO: Todos os dados serão apagados permanentemente.\n\nDeseja continuar?')) return;
    if (!confirm('TEM CERTEZA? Esta ação NÃO pode ser desfeita.\n\nClique OK para confirmar a exclusão de TODOS os dados.')) return;

    try {
      U.loading(true, 'Zerando banco de dados...');
      await API.post('/api/reset-database', { confirmacao: 'CONFIRMAR' });
      U.loading(false);
      U.toast('Banco de dados zerado com sucesso!');
      navigate('dashboard');
    } catch(err) {
      U.loading(false);
      U.toast(err.message, 'error');
    }
  },
};

// -------------------------------------------------------
// INIT
// -------------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {
  // Sidebar navigation
  document.querySelectorAll('.nav-item[data-page]').forEach(el => {
    el.addEventListener('click', e => {
      e.preventDefault();
      navigate(el.dataset.page);
    });
  });

  // Sidebar toggle
  document.getElementById('sidebar-toggle')?.addEventListener('click', () => {
    const sidebar = document.getElementById('sidebar');
    const main    = document.getElementById('main-wrapper');
    if (window.innerWidth <= 900) {
      sidebar.classList.toggle('mobile-open');
    } else {
      sidebar.classList.toggle('collapsed');
      main.classList.toggle('expanded');
    }
  });

  // Tolerância tipo change
  document.addEventListener('change', e => {
    if (e.target.name === 'tolerancia_tipo') {
      const lbl = document.getElementById('tol-val-label');
      if (lbl) lbl.textContent = `Valor da Tolerância (${e.target.value === 'fixo' ? 'R$' : '%'})`;
    }
  });

  // Close modal on Escape
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') U.closeModal();
  });

  // Bootstrap the app
  navigate('dashboard');
});

// -------------------------------------------------------
// CADASTRO COMPLETO DE LINHAS
// -------------------------------------------------------
const Cadastro = {
  items: [],
  empresas: [],

  async load() {
    const sec = document.getElementById('page-cadastro');
    sec.innerHTML = `<div class="page-header">
      <h1>Cadastro de Linhas</h1>
      <div class="page-header-actions">
        <button class="btn btn-outline" onclick="Cadastro.comparar()"><i class="fas fa-scale-balanced"></i> Comparar com Faturas</button>
        <button class="btn btn-primary" onclick="Cadastro.openForm()"><i class="fas fa-plus"></i> Nova Linha</button>
      </div>
    </div>
    <div class="filter-bar">
      <div class="search-field"><i class="fas fa-search"></i>
        <input id="cad-search" type="text" placeholder="Buscar por número, nome, matrícula ou plano...">
      </div>
      <select id="cad-empresa" class="select-filter" onchange="Cadastro.fetch()">
        <option value="">Todas as empresas</option>
      </select>
      <select id="cad-status" class="select-filter" onchange="Cadastro.fetch()">
        <option value="">Todos os status</option>
        <option value="em_uso">Em Uso</option>
        <option value="bloqueado">Bloqueado</option>
        <option value="suspenso_120">Suspenso 120 Dias</option>
        <option value="cancelado">Cancelado</option>
      </select>
    </div>
    <div id="cad-body"><div class="loading-inline"><div class="spinner-sm"></div></div></div>`;

    document.getElementById('cad-search').addEventListener('input', U.debounce(() => Cadastro.fetch(), 380));

    // Carregar empresas
    try {
      this.empresas = await API.get('/api/cadastro/empresas');
      const sel = document.getElementById('cad-empresa');
      this.empresas.forEach(e => { const o = document.createElement('option'); o.value = e; o.textContent = e; sel.appendChild(o); });
    } catch(e) {}

    await this.fetch();
  },

  async fetch() {
    const search = document.getElementById('cad-search')?.value || '';
    const empresa = document.getElementById('cad-empresa')?.value || '';
    const status_linha = document.getElementById('cad-status')?.value || '';
    const body = document.getElementById('cad-body');

    try {
      const url = `/api/cadastro?search=${encodeURIComponent(search)}&empresa=${encodeURIComponent(empresa)}&status_linha=${encodeURIComponent(status_linha)}`;
      this.items = await API.get(url);
      body.innerHTML = this.table();
    } catch(e) {
      body.innerHTML = `<div class="alert alert-error">${e.message}</div>`;
    }
  },

  table() {
    if (!this.items.length) {
      return U.emptyState('fa-address-book', 'Nenhuma linha cadastrada',
        'Cadastre as linhas telefônicas com dados do funcionário e centro de custo.',
        `<button class="btn btn-primary" onclick="Cadastro.openForm()"><i class="fas fa-plus"></i> Nova Linha</button>`);
    }

    const rows = this.items.map(c => {
      const confCls = c.conferencia === 'ok' ? 'badge-ok' : c.conferencia === 'divergente' ? 'badge-divergente' : 'badge-aproximado';
      const confLabel = c.conferencia === 'ok' ? 'OK' : c.conferencia === 'divergente' ? 'Divergente' : 'Pendente';
      const statusMap = { em_uso: ['Em Uso','badge-ok'], bloqueado: ['Bloqueado','badge-aproximado'], suspenso_120: ['Suspenso 120d','badge-divergente'], cancelado: ['Cancelado','badge-sem-contrato'] };
      const [statusLabel, statusCls] = statusMap[c.status_linha] || statusMap.em_uso;
      return `<tr>
        <td><span class="phone-number">${U.phone(c.numero_telefone)}</span></td>
        <td>${U.esc(c.operadora || '—')}</td>
        <td style="text-align:center">${c.vencimento || '—'}</td>
        <td>${U.esc(c.nome_funcionario || '—')}</td>
        <td>${U.esc(c.matricula_funcionario || '—')}</td>
        <td>${U.esc(c.centro_custo || '—')}</td>
        <td>${U.esc(c.plano || '—')}</td>
        <td>${U.esc(c.empresa || '—')}</td>
        <td class="text-right">${U.money(c.valor_contrato)}</td>
        <td class="text-right">${U.money(c.valor_fatura)}</td>
        <td class="text-right ${c.diferenca > 0.005 ? 'val-pos' : c.diferenca < -0.005 ? 'val-neg' : 'val-zero'}">${U.money(c.diferenca)}</td>
        <td><span class="badge ${statusCls}">${statusLabel}</span></td>
        <td><span class="badge ${confCls}">${confLabel}</span></td>
        <td style="white-space:nowrap">
          <button onclick="Cadastro.openForm(${c.id})" style="background:none;border:none;cursor:pointer;color:var(--primary)" title="Editar">
            <i class="fas fa-pen-to-square"></i>
          </button>
          <button onclick="Cadastro.remove(${c.id})" style="background:none;border:none;cursor:pointer;color:#e53e3e" title="Excluir">
            <i class="fas fa-trash"></i>
          </button>
        </td>
      </tr>`;
    }).join('');

    return `
      <div class="table-wrapper">
        <table class="data-table" style="font-size:12px">
          <thead><tr>
            <th>Número</th><th>Operadora</th><th>Venc.</th><th>Funcionário</th><th>Matrícula</th>
            <th>C. Custo</th><th>Plano</th><th>Empresa</th>
            <th class="text-right">V. Contrato</th><th class="text-right">V. Fatura</th>
            <th class="text-right">Diferença</th><th>Status</th><th>Conferência</th><th></th>
          </tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
      <div class="table-footer-info">${this.items.length} linha(s) cadastrada(s)</div>`;
  },

  openForm(id) {
    const c = id ? this.items.find(i => i.id === id) : {};
    const isEdit = !!id;

    U.modal(`
      <h2 class="modal-title">${isEdit ? 'Editar' : 'Nova'} Linha</h2>
      <form id="cad-form" onsubmit="Cadastro.save(event, ${id || 'null'})">
        <div class="form-row">
          <div class="form-group"><label>Número do Telefone *</label>
            <input name="numero_telefone" type="text" required value="${U.esc(c.numero_telefone || '')}" placeholder="11999887766"></div>
          <div class="form-group"><label>Operadora</label>
            <select name="operadora" class="select-filter" style="width:100%">
              <option value="">Selecione</option>
              <option value="Vivo" ${c.operadora==='Vivo'?'selected':''}>Vivo</option>
              <option value="Claro" ${c.operadora==='Claro'?'selected':''}>Claro</option>
              <option value="Tim" ${c.operadora==='Tim'?'selected':''}>Tim</option>
              <option value="Oi" ${c.operadora==='Oi'?'selected':''}>Oi</option>
            </select></div>
          <div class="form-group"><label>Dia Vencimento</label>
            <input name="vencimento" type="number" min="1" max="31" value="${c.vencimento || ''}" placeholder="15"></div>
        </div>
        <div class="form-row">
          <div class="form-group"><label>Nº da Conta</label>
            <input name="numero_conta" type="text" value="${U.esc(c.numero_conta || '')}"></div>
          <div class="form-group"><label>Matrícula Funcionário</label>
            <input name="matricula_funcionario" type="text" value="${U.esc(c.matricula_funcionario || '')}"></div>
          <div class="form-group"><label>Nome Funcionário</label>
            <input name="nome_funcionario" type="text" value="${U.esc(c.nome_funcionario || '')}"></div>
        </div>
        <div class="form-row">
          <div class="form-group"><label>Centro de Custo</label>
            <input name="centro_custo" type="text" value="${U.esc(c.centro_custo || '')}"></div>
          <div class="form-group"><label>Nome do Centro de Custo</label>
            <input name="nome_centro_custo" type="text" value="${U.esc(c.nome_centro_custo || '')}"></div>
        </div>
        <div class="form-row">
          <div class="form-group"><label>Plano</label>
            <input name="plano" type="text" value="${U.esc(c.plano || '')}" placeholder="Ex: SMART EMPRESAS 8GB TE"></div>
          <div class="form-group"><label>Valor do Plano R$</label>
            <input name="valor_plano" type="number" step="0.01" value="${c.valor_plano || ''}"></div>
          <div class="form-group"><label>Empresa / Unidade</label>
            <input name="empresa" type="text" value="${U.esc(c.empresa || '')}" placeholder="Nome da unidade"></div>
        </div>
        <div class="form-row">
          <div class="form-group"><label>Valor Contrato R$</label>
            <input name="valor_contrato" type="number" step="0.01" value="${c.valor_contrato || ''}"></div>
          <div class="form-group"><label>Valor Fatura R$</label>
            <input name="valor_fatura" type="number" step="0.01" value="${c.valor_fatura || ''}"></div>
          <div class="form-group"><label>Conferência</label>
            <select name="conferencia" class="select-filter" style="width:100%">
              <option value="pendente" ${(c.conferencia||'pendente')==='pendente'?'selected':''}>Pendente</option>
              <option value="ok" ${c.conferencia==='ok'?'selected':''}>OK (Verde)</option>
              <option value="divergente" ${c.conferencia==='divergente'?'selected':''}>Divergente (Vermelho)</option>
            </select></div>
        </div>
        <div class="form-row">
          <div class="form-group"><label>Status da Linha</label>
            <select name="status_linha" class="select-filter" style="width:100%">
              <option value="em_uso" ${(c.status_linha||'em_uso')==='em_uso'?'selected':''}>Em Uso</option>
              <option value="bloqueado" ${c.status_linha==='bloqueado'?'selected':''}>Bloqueado</option>
              <option value="suspenso_120" ${c.status_linha==='suspenso_120'?'selected':''}>Suspenso 120 Dias</option>
              <option value="cancelado" ${c.status_linha==='cancelado'?'selected':''}>Cancelado</option>
            </select></div>
        </div>
        <div class="form-actions">
          <button type="button" class="btn btn-ghost" onclick="U.closeModal()">Cancelar</button>
          <button type="submit" class="btn btn-primary"><i class="fas fa-check"></i> ${isEdit ? 'Salvar' : 'Cadastrar'}</button>
        </div>
      </form>
    `, 'lg');
  },

  async save(e, id) {
    e.preventDefault();
    const f = e.target;
    const data = {
      numero_telefone: f.numero_telefone.value,
      operadora: f.operadora.value,
      vencimento: f.vencimento.value || null,
      numero_conta: f.numero_conta.value,
      matricula_funcionario: f.matricula_funcionario.value,
      nome_funcionario: f.nome_funcionario.value,
      centro_custo: f.centro_custo.value,
      nome_centro_custo: f.nome_centro_custo.value,
      plano: f.plano.value,
      valor_plano: f.valor_plano.value || null,
      empresa: f.empresa.value,
      valor_contrato: f.valor_contrato.value || null,
      valor_fatura: f.valor_fatura.value || null,
      conferencia: f.conferencia.value,
      status_linha: f.status_linha.value,
    };
    try {
      if (id) {
        await API.put(`/api/cadastro/${id}`, data);
        U.toast('Linha atualizada!');
      } else {
        await API.post('/api/cadastro', data);
        U.toast('Linha cadastrada!');
      }
      U.closeModal();
      this.fetch();
    } catch(err) { U.toast(err.message, 'error'); }
  },

  async remove(id) {
    if (!confirm('Excluir esta linha?')) return;
    try {
      await API.del(`/api/cadastro/${id}`);
      U.toast('Linha excluída.');
      this.fetch();
    } catch(e) { U.toast(e.message, 'error'); }
  },

  async comparar() {
    if (!confirm('Comparar todas as linhas do cadastro com as faturas importadas?\nIsso atualizará os valores de fatura, diferença e conferência.')) return;
    try {
      U.loading(true, 'Comparando com faturas...');
      const res = await API.post('/api/cadastro/comparar');
      U.loading(false);
      U.toast(`${res.atualizados} linhas atualizadas, ${res.divergentes} divergentes.`);
      this.fetch();
    } catch(e) { U.loading(false); U.toast(e.message, 'error'); }
  },
};

// -------------------------------------------------------
// COOPERNAC
// -------------------------------------------------------
const Coopernac = {
  voz: [], dados: [], resumo: [],

  async load() {
    const sec = document.getElementById('page-coopernac');
    sec.innerHTML = `<div class="page-header"><h1>Coopernac</h1></div>
      <div id="coop-body"><div class="loading-inline"><div class="spinner-sm"></div></div></div>`;

    try {
      [this.voz, this.dados, this.resumo] = await Promise.all([
        API.get('/api/coopernac/voz'),
        API.get('/api/coopernac/dados'),
        API.get('/api/coopernac/resumo'),
      ]);
      document.getElementById('coop-body').innerHTML = this.html();
    } catch(e) {
      document.getElementById('coop-body').innerHTML = `<div class="alert alert-error">${e.message}</div>`;
    }
  },

  html() {
    const vozTotal = this.voz.reduce((s, v) => s + (v.total || 0), 0);
    const dadosTotal = this.dados.reduce((s, v) => s + (v.total || 0), 0);
    const resumoTotal = this.resumo.reduce((s, v) => s + (v.valores || 0), 0);

    return `
      <!-- VOZ -->
      <div class="config-card">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
          <h3 style="margin:0"><i class="fas fa-phone" style="color:#660099;margin-right:6px"></i>Voz</h3>
          <button class="btn btn-primary btn-sm" onclick="Coopernac.addItem('voz')"><i class="fas fa-plus"></i> Adicionar</button>
        </div>
        <table class="fatura-table">
          <thead><tr><th>Número</th><th>Descrição</th><th style="text-align:right">Total R$</th><th style="width:60px"></th></tr></thead>
          <tbody>${this.voz.map(v => `<tr>
            <td>${U.esc(v.numero)}</td><td>${U.esc(v.descricao)}</td>
            <td style="text-align:right">${U.money(v.total)}</td>
            <td style="text-align:center">
              <button onclick="Coopernac.editItem('voz',${v.id})" style="background:none;border:none;cursor:pointer;color:var(--primary)"><i class="fas fa-pen-to-square"></i></button>
              <button onclick="Coopernac.deleteItem('voz',${v.id})" style="background:none;border:none;cursor:pointer;color:#e53e3e"><i class="fas fa-trash"></i></button>
            </td>
          </tr>`).join('') || '<tr><td colspan="4" style="text-align:center;color:var(--text-secondary)">Nenhum registro</td></tr>'}</tbody>
          <tfoot><tr><td colspan="2" style="text-align:right;font-weight:600">TOTAL VOZ</td>
            <td style="text-align:right;font-weight:600">${U.money(vozTotal)}</td><td></td></tr></tfoot>
        </table>
      </div>

      <!-- DADOS -->
      <div class="config-card">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
          <h3 style="margin:0"><i class="fas fa-wifi" style="color:#660099;margin-right:6px"></i>Dados</h3>
          <button class="btn btn-primary btn-sm" onclick="Coopernac.addItem('dados')"><i class="fas fa-plus"></i> Adicionar</button>
        </div>
        <table class="fatura-table">
          <thead><tr><th>Número</th><th>Descrição</th><th style="text-align:right">Total R$</th><th style="width:60px"></th></tr></thead>
          <tbody>${this.dados.map(v => `<tr>
            <td>${U.esc(v.numero)}</td><td>${U.esc(v.descricao)}</td>
            <td style="text-align:right">${U.money(v.total)}</td>
            <td style="text-align:center">
              <button onclick="Coopernac.editItem('dados',${v.id})" style="background:none;border:none;cursor:pointer;color:var(--primary)"><i class="fas fa-pen-to-square"></i></button>
              <button onclick="Coopernac.deleteItem('dados',${v.id})" style="background:none;border:none;cursor:pointer;color:#e53e3e"><i class="fas fa-trash"></i></button>
            </td>
          </tr>`).join('') || '<tr><td colspan="4" style="text-align:center;color:var(--text-secondary)">Nenhum registro</td></tr>'}</tbody>
          <tfoot><tr><td colspan="2" style="text-align:right;font-weight:600">TOTAL DADOS</td>
            <td style="text-align:right;font-weight:600">${U.money(dadosTotal)}</td><td></td></tr></tfoot>
        </table>
      </div>

      <!-- RESUMO -->
      <div class="config-card">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
          <h3 style="margin:0"><i class="fas fa-calculator" style="color:#660099;margin-right:6px"></i>Resumo (Voz + Dados)</h3>
          <button class="btn btn-primary btn-sm" onclick="Coopernac.addResumo()"><i class="fas fa-plus"></i> Adicionar</button>
        </div>
        <table class="fatura-table">
          <thead><tr><th style="text-align:right">Valores R$</th><th>Descrição</th><th>Observação</th><th style="width:60px"></th></tr></thead>
          <tbody>${this.resumo.map(r => `<tr>
            <td style="text-align:right">${U.money(r.valores)}</td><td>${U.esc(r.descricao)}</td><td>${U.esc(r.observacao || '')}</td>
            <td style="text-align:center">
              <button onclick="Coopernac.editResumo(${r.id})" style="background:none;border:none;cursor:pointer;color:var(--primary)"><i class="fas fa-pen-to-square"></i></button>
              <button onclick="Coopernac.deleteResumo(${r.id})" style="background:none;border:none;cursor:pointer;color:#e53e3e"><i class="fas fa-trash"></i></button>
            </td>
          </tr>`).join('') || '<tr><td colspan="4" style="text-align:center;color:var(--text-secondary)">Nenhum registro</td></tr>'}</tbody>
          <tfoot><tr><td style="text-align:right;font-weight:600">${U.money(resumoTotal)}</td>
            <td colspan="2" style="font-weight:600">TOTAL GERAL</td><td></td></tr></tfoot>
        </table>
      </div>

      <!-- Totalizador -->
      <div class="cards-row" style="max-width:600px">
        <div class="summary-card"><span class="summary-label">Total Voz</span><span class="summary-value">${U.money(vozTotal)}</span></div>
        <div class="summary-card"><span class="summary-label">Total Dados</span><span class="summary-value">${U.money(dadosTotal)}</span></div>
        <div class="summary-card"><span class="summary-label">Soma Geral</span><span class="summary-value" style="color:#660099;font-weight:700">${U.money(vozTotal + dadosTotal)}</span></div>
      </div>`;
  },

  addItem(tipo) {
    U.modal(`
      <h3>Novo registro — ${tipo === 'voz' ? 'Voz' : 'Dados'}</h3>
      <form id="coop-form" onsubmit="Coopernac.saveItem(event,'${tipo}')">
        <div class="form-group"><label>Número</label><input name="numero" type="text" required></div>
        <div class="form-group"><label>Descrição</label><input name="descricao" type="text"></div>
        <div class="form-group"><label>Total R$</label><input name="total" type="number" step="0.01" value="0"></div>
        <div class="form-actions">
          <button type="button" class="btn btn-ghost" onclick="U.closeModal()">Cancelar</button>
          <button type="submit" class="btn btn-primary">Salvar</button>
        </div>
      </form>`);
  },

  editItem(tipo, id) {
    const list = tipo === 'voz' ? this.voz : this.dados;
    const item = list.find(v => v.id === id);
    if (!item) return;
    U.modal(`
      <h3>Editar — ${tipo === 'voz' ? 'Voz' : 'Dados'}</h3>
      <form id="coop-form" onsubmit="Coopernac.saveItem(event,'${tipo}',${id})">
        <div class="form-group"><label>Número</label><input name="numero" type="text" value="${U.esc(item.numero)}" required></div>
        <div class="form-group"><label>Descrição</label><input name="descricao" type="text" value="${U.esc(item.descricao)}"></div>
        <div class="form-group"><label>Total R$</label><input name="total" type="number" step="0.01" value="${item.total || 0}"></div>
        <div class="form-actions">
          <button type="button" class="btn btn-ghost" onclick="U.closeModal()">Cancelar</button>
          <button type="submit" class="btn btn-primary">Salvar</button>
        </div>
      </form>`);
  },

  async saveItem(e, tipo, id) {
    e.preventDefault();
    const f = e.target;
    const data = { numero: f.numero.value, descricao: f.descricao.value, total: parseFloat(f.total.value) || 0 };
    try {
      if (id) { await API.put(`/api/coopernac/${tipo}/${id}`, data); }
      else    { await API.post(`/api/coopernac/${tipo}`, data); }
      U.closeModal();
      U.toast('Salvo!');
      this.load();
    } catch(err) { U.toast(err.message, 'error'); }
  },

  async deleteItem(tipo, id) {
    if (!confirm('Excluir este registro?')) return;
    try { await API.del(`/api/coopernac/${tipo}/${id}`); this.load(); } catch(e) { U.toast(e.message,'error'); }
  },

  addResumo() {
    U.modal(`
      <h3>Novo Resumo</h3>
      <form id="coop-res-form" onsubmit="Coopernac.saveResumo(event)">
        <div class="form-group"><label>Valores R$</label><input name="valores" type="number" step="0.01" value="0"></div>
        <div class="form-group"><label>Descrição</label><input name="descricao" type="text"></div>
        <div class="form-group"><label>Observação</label><input name="observacao" type="text"></div>
        <div class="form-actions">
          <button type="button" class="btn btn-ghost" onclick="U.closeModal()">Cancelar</button>
          <button type="submit" class="btn btn-primary">Salvar</button>
        </div>
      </form>`);
  },

  editResumo(id) {
    const item = this.resumo.find(r => r.id === id);
    if (!item) return;
    U.modal(`
      <h3>Editar Resumo</h3>
      <form id="coop-res-form" onsubmit="Coopernac.saveResumo(event,${id})">
        <div class="form-group"><label>Valores R$</label><input name="valores" type="number" step="0.01" value="${item.valores || 0}"></div>
        <div class="form-group"><label>Descrição</label><input name="descricao" type="text" value="${U.esc(item.descricao)}"></div>
        <div class="form-group"><label>Observação</label><input name="observacao" type="text" value="${U.esc(item.observacao || '')}"></div>
        <div class="form-actions">
          <button type="button" class="btn btn-ghost" onclick="U.closeModal()">Cancelar</button>
          <button type="submit" class="btn btn-primary">Salvar</button>
        </div>
      </form>`);
  },

  async saveResumo(e, id) {
    e.preventDefault();
    const f = e.target;
    const data = { valores: parseFloat(f.valores.value) || 0, descricao: f.descricao.value, observacao: f.observacao.value };
    try {
      if (id) { await API.put(`/api/coopernac/resumo/${id}`, data); }
      else    { await API.post('/api/coopernac/resumo', data); }
      U.closeModal();
      U.toast('Salvo!');
      this.load();
    } catch(err) { U.toast(err.message, 'error'); }
  },

  async deleteResumo(id) {
    if (!confirm('Excluir?')) return;
    try { await API.del(`/api/coopernac/resumo/${id}`); this.load(); } catch(e) { U.toast(e.message,'error'); }
  },
};

// -------------------------------------------------------
// FATURA PDF → XLS
// -------------------------------------------------------
const FaturaXls = {
  state: { rows: [], meta: {}, filename: '', rawText: '' },
  planos: [],

  async render() {
    await this.loadPlanos();
    const sec = document.getElementById('page-fatura-xls');
    sec.innerHTML = `
      <div class="page-header">
        <h1>Fatura PDF <i class="fas fa-arrow-right" style="font-size:14px"></i> Importar &amp; Comparar</h1>
        <p style="color:var(--text-secondary);font-size:13px;margin:4px 0 0">
          Importe faturas Vivo, cadastre valores de contrato e audite divergências.
        </p>
      </div>

      <!-- Gestão de Planos -->
      <div class="fatura-xls-card" style="margin-bottom:20px">
        <h3 style="margin:0 0 12px;font-size:15px;font-weight:600">
          <i class="fas fa-tags" style="color:#660099;margin-right:6px"></i>Valores dos Planos (Contrato)
        </h3>
        <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:flex-end;margin-bottom:12px">
          <div style="flex:1;min-width:200px">
            <label style="font-size:12px;font-weight:500;color:var(--text-secondary)">Nome do Plano</label>
            <input id="plano-nome" type="text" placeholder="Ex: SMART EMPRESAS 8GB TE"
                   style="width:100%;padding:7px 10px;font-size:13px;border:1px solid var(--border);border-radius:6px">
          </div>
          <div style="min-width:120px">
            <label style="font-size:12px;font-weight:500;color:var(--text-secondary)">Valor Contrato R$</label>
            <input id="plano-valor" type="number" step="0.01" placeholder="37,49"
                   style="width:100%;padding:7px 10px;font-size:13px;border:1px solid var(--border);border-radius:6px">
          </div>
          <button class="btn btn-primary btn-sm" onclick="FaturaXls.addPlano()" style="height:36px">
            <i class="fas fa-plus"></i> Adicionar
          </button>
        </div>
        <div id="planos-table-wrap">${this._planosTable()}</div>
      </div>

      <!-- Upload PDF -->
      <div class="fatura-xls-card">
        <h3 style="margin:0 0 12px;font-size:15px;font-weight:600">
          <i class="fas fa-file-pdf" style="color:#660099;margin-right:6px"></i>Importar Fatura PDF
        </h3>
        <div class="drop-zone" id="fxls-drop"
             onclick="document.getElementById('fxls-input').click()"
             ondragover="FaturaXls.onDragOver(event)"
             ondragleave="FaturaXls.onDragLeave(event)"
             ondrop="FaturaXls.onDrop(event)">
          <div class="dz-icon"><i class="fas fa-file-pdf"></i></div>
          <strong>Arraste a fatura PDF aqui</strong>
          <p>ou <u style="cursor:pointer">selecione o arquivo</u></p>
          <p>PDF com texto selecionável &middot; até 100 MB (dividido automaticamente)</p>
          <input type="file" id="fxls-input" accept=".pdf" style="display:none"
                 onchange="FaturaXls.onFileSelected(this.files[0])">
        </div>
        <div id="fxls-result"></div>
      </div>`;
  },

  async loadPlanos() {
    try { this.planos = await API.get('/api/planos'); } catch(e) { this.planos = []; }
  },

  _planosTable() {
    if (!this.planos.length) return '<p style="color:var(--text-secondary);font-size:13px;margin:0">Nenhum plano cadastrado ainda.</p>';
    const rows = this.planos.map(p => `<tr>
      <td style="font-size:13px;font-weight:500">${U.esc(p.nome_plano)}</td>
      <td style="font-size:13px;text-align:right">${U.money(p.valor_contrato)}</td>
      <td style="text-align:center;width:60px">
        <button onclick="FaturaXls.editPlano(${p.id},'${U.esc(p.nome_plano)}',${p.valor_contrato})"
                style="background:none;border:none;cursor:pointer;color:var(--primary)" title="Editar">
          <i class="fas fa-pen-to-square"></i>
        </button>
        <button onclick="FaturaXls.deletePlano(${p.id})"
                style="background:none;border:none;cursor:pointer;color:#e53e3e" title="Excluir">
          <i class="fas fa-trash"></i>
        </button>
      </td>
    </tr>`).join('');
    return `<table class="fatura-table" style="margin:0">
      <thead><tr><th>Plano</th><th style="text-align:right">Valor Contrato</th><th></th></tr></thead>
      <tbody>${rows}</tbody>
    </table>`;
  },

  async addPlano() {
    const nomeEl  = document.getElementById('plano-nome');
    const valorEl = document.getElementById('plano-valor');
    const nome  = (nomeEl.value || '').trim();
    const valor = parseFloat(valorEl.value);
    if (!nome) { U.toast('Informe o nome do plano.', 'error'); return; }
    if (isNaN(valor) || valor < 0) { U.toast('Informe um valor válido.', 'error'); return; }
    try {
      await API.post('/api/planos', { nome_plano: nome, valor_contrato: valor });
      nomeEl.value = ''; valorEl.value = '';
      await this.loadPlanos();
      document.getElementById('planos-table-wrap').innerHTML = this._planosTable();
      U.toast('Plano salvo com sucesso!');
    } catch(e) { U.toast(e.message, 'error'); }
  },

  editPlano(id, nome, valor) {
    U.modal(`
      <h3 style="margin:0 0 16px">Editar Plano</h3>
      <div style="margin-bottom:12px">
        <label style="font-size:12px;font-weight:500">Nome do Plano</label>
        <input id="edit-plano-nome" type="text" value="${U.esc(nome)}"
               style="width:100%;padding:7px 10px;font-size:13px;border:1px solid var(--border);border-radius:6px">
      </div>
      <div style="margin-bottom:16px">
        <label style="font-size:12px;font-weight:500">Valor Contrato (R$)</label>
        <input id="edit-plano-valor" type="number" step="0.01" value="${valor}"
               style="width:100%;padding:7px 10px;font-size:13px;border:1px solid var(--border);border-radius:6px">
      </div>
      <div style="display:flex;gap:8px;justify-content:flex-end">
        <button class="btn btn-ghost" onclick="U.closeModal()">Cancelar</button>
        <button class="btn btn-primary" onclick="FaturaXls.savePlano(${id})">Salvar</button>
      </div>
    `);
  },

  async savePlano(id) {
    const nome  = document.getElementById('edit-plano-nome').value.trim();
    const valor = parseFloat(document.getElementById('edit-plano-valor').value);
    if (!nome || isNaN(valor)) { U.toast('Preencha todos os campos.', 'error'); return; }
    try {
      await API.put(`/api/planos/${id}`, { nome_plano: nome, valor_contrato: valor });
      U.closeModal();
      await this.loadPlanos();
      document.getElementById('planos-table-wrap').innerHTML = this._planosTable();
      U.toast('Plano atualizado!');
    } catch(e) { U.toast(e.message, 'error'); }
  },

  async deletePlano(id) {
    if (!confirm('Excluir este plano?')) return;
    try {
      await API.del(`/api/planos/${id}`);
      await this.loadPlanos();
      document.getElementById('planos-table-wrap').innerHTML = this._planosTable();
      U.toast('Plano excluído.');
    } catch(e) { U.toast(e.message, 'error'); }
  },

  onDragOver(e) {
    e.preventDefault();
    document.getElementById('fxls-drop').classList.add('drag-over');
  },

  onDragLeave(e) {
    document.getElementById('fxls-drop').classList.remove('drag-over');
  },

  onDrop(e) {
    e.preventDefault();
    document.getElementById('fxls-drop').classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) this.onFileSelected(file);
  },

  async onFileSelected(file) {
    if (!file || !file.name.toLowerCase().endsWith('.pdf')) {
      U.toast('Selecione um arquivo PDF.', 'error'); return;
    }
    const resultEl = document.getElementById('fxls-result');
    const MAX_CHUNK = 3.5 * 1024 * 1024; // 3.5MB por chunk (margem para o limite de 4.5MB)

    if (file.size <= MAX_CHUNK) {
      // Arquivo pequeno — upload direto
      await this._uploadSinglePdf(file, resultEl);
    } else {
      // Arquivo grande — dividir em partes por páginas
      await this._uploadChunkedPdf(file, resultEl);
    }
  },

  async _uploadSinglePdf(file, resultEl) {
    resultEl.innerHTML = `
      <div style="display:flex;align-items:center;gap:10px;padding:20px 0;color:var(--text-secondary)">
        <div class="spinner-sm"></div>
        <span>Lendo <strong>${U.esc(file.name)}</strong>…</span>
      </div>`;

    try {
      const fd = new FormData();
      fd.append('file', file);
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), 120000);
      let res;
      try {
        const r = await fetch('/api/faturas/pdf-preview', { method: 'POST', body: fd, signal: ctrl.signal });
        clearTimeout(timer);
        res = await r.json().catch(() => ({}));
        if (!r.ok) throw new Error(res.error || `Erro HTTP ${r.status}`);
      } catch (ex) {
        clearTimeout(timer);
        throw ex.name === 'AbortError'
          ? new Error('Tempo limite excedido ao processar o PDF.')
          : ex;
      }

      this.state = {
        rows:     res.linhas || [],
        meta:     { competencia: res.competencia, operadora: res.operadora, numero_fatura: res.numero_fatura },
        filename: file.name,
        rawText:  res.raw_text || '',
      };
      this.renderPreview(res);
    } catch (err) {
      resultEl.innerHTML = `<div class="alert alert-danger" style="margin-top:16px">
        <i class="fas fa-circle-exclamation"></i> ${U.esc(err.message)}
      </div>`;
    }
  },

  async _uploadChunkedPdf(file, resultEl) {
    resultEl.innerHTML = `
      <div style="display:flex;align-items:center;gap:10px;padding:20px 0;color:var(--text-secondary)">
        <div class="spinner-sm"></div>
        <span>Dividindo <strong>${U.esc(file.name)}</strong> (${(file.size / 1024 / 1024).toFixed(1)} MB) em partes…</span>
      </div>`;

    try {
      const arrayBuf = await file.arrayBuffer();
      const pdfDoc = await PDFLib.PDFDocument.load(arrayBuf);
      const totalPages = pdfDoc.getPageCount();

      // Descobrir quantas páginas por chunk (estimar ~tamanho/paginas)
      const avgPageSize = file.size / totalPages;
      const pagesPerChunk = Math.max(1, Math.floor(3.5 * 1024 * 1024 / avgPageSize));

      const chunks = [];
      for (let start = 0; start < totalPages; start += pagesPerChunk) {
        const end = Math.min(start + pagesPerChunk, totalPages);
        chunks.push({ start, end });
      }

      let allLinhas = [];
      let lastMeta = {};

      for (let i = 0; i < chunks.length; i++) {
        const { start, end } = chunks[i];
        resultEl.innerHTML = `
          <div style="display:flex;align-items:center;gap:10px;padding:20px 0;color:var(--text-secondary)">
            <div class="spinner-sm"></div>
            <span>Processando parte ${i + 1}/${chunks.length} (páginas ${start + 1}-${end} de ${totalPages})…</span>
          </div>
          <div style="background:var(--border);border-radius:6px;height:6px;margin-top:8px;overflow:hidden">
            <div style="background:var(--primary);height:100%;width:${Math.round(((i + 1) / chunks.length) * 100)}%;transition:width .3s"></div>
          </div>`;

        // Criar sub-PDF com as páginas do chunk
        const chunkPdf = await PDFLib.PDFDocument.create();
        const pages = await chunkPdf.copyPages(pdfDoc, Array.from({ length: end - start }, (_, k) => start + k));
        pages.forEach(p => chunkPdf.addPage(p));
        const chunkBytes = await chunkPdf.save();

        const blob = new Blob([chunkBytes], { type: 'application/pdf' });
        const chunkFile = new File([blob], `${file.name}_parte${i + 1}.pdf`, { type: 'application/pdf' });

        const fd = new FormData();
        fd.append('file', chunkFile);
        const ctrl = new AbortController();
        const timer = setTimeout(() => ctrl.abort(), 120000);
        try {
          const r = await fetch('/api/faturas/pdf-preview', { method: 'POST', body: fd, signal: ctrl.signal });
          clearTimeout(timer);
          const res = await r.json().catch(() => ({}));
          if (!r.ok) throw new Error(res.error || `Erro HTTP ${r.status}`);
          if (res.linhas) allLinhas = allLinhas.concat(res.linhas);
          if (res.competencia) lastMeta.competencia = res.competencia;
          if (res.operadora) lastMeta.operadora = res.operadora;
          if (res.numero_fatura) lastMeta.numero_fatura = res.numero_fatura;
        } catch (ex) {
          clearTimeout(timer);
          if (ex.name === 'AbortError') throw new Error(`Timeout na parte ${i + 1}`);
          throw ex;
        }
      }

      // Remover duplicatas por numero_vivo (caso páginas se sobreponham)
      const seen = new Set();
      const unique = [];
      allLinhas.forEach(l => {
        const key = (l.numero_vivo || '') + '|' + (l.plano || '');
        if (!seen.has(key)) { seen.add(key); unique.push(l); }
      });

      const combined = {
        linhas: unique,
        competencia: lastMeta.competencia,
        operadora: lastMeta.operadora,
        numero_fatura: lastMeta.numero_fatura,
        filename: file.name,
      };

      this.state = {
        rows:     unique,
        meta:     lastMeta,
        filename: file.name,
        rawText:  '',
      };
      this.renderPreview(combined);
      U.toast(`PDF dividido em ${chunks.length} partes. ${unique.length} linhas extraídas.`);
    } catch (err) {
      resultEl.innerHTML = `<div class="alert alert-danger" style="margin-top:16px">
        <i class="fas fa-circle-exclamation"></i> ${U.esc(err.message)}
      </div>`;
    }
  },

  renderPreview(res) {
    const resultEl = document.getElementById('fxls-result');
    const rows = res.linhas || [];
    const nothingFound = rows.length === 0;

    const badges = [
      res.competencia   && `<span class="fatura-meta-badge"><i class="fas fa-calendar-days"></i> ${U.esc(res.competencia)}</span>`,
      res.operadora     && `<span class="fatura-meta-badge"><i class="fas fa-building"></i> ${U.esc(res.operadora)}</span>`,
      res.numero_fatura && `<span class="fatura-meta-badge"><i class="fas fa-hashtag"></i> ${U.esc(res.numero_fatura)}</span>`,
    ].filter(Boolean).join('');

    // Build plano lookup for preview comparison
    const planoMap = {};
    this.planos.forEach(p => { planoMap[p.nome_plano.toUpperCase()] = p.valor_contrato; });

    let total = 0;
    const trs = rows.map((r, i) => {
      const v = r.valor_total != null ? r.valor_total : null;
      if (v != null) total += v;
      let num = r.numero_vivo || '';
      if (/^\d{11}$/.test(num)) num = `${num.slice(0,2)}-${num.slice(2,7)}-${num.slice(7)}`;
      else if (/^\d{10}$/.test(num)) num = `${num.slice(0,2)}-${num.slice(2,6)}-${num.slice(6)}`;

      // Comparison preview
      const planoUpper = (r.plano || '').toUpperCase();
      const vContrato = planoMap[planoUpper];
      const diff = (v != null && vContrato != null) ? (v - vContrato) : null;
      const diffHtml = diff != null
        ? `<span class="${Math.abs(diff) < 0.01 ? 'val-zero' : diff > 0 ? 'val-pos' : 'val-neg'}" style="font-size:12px;font-weight:600">${U.money(diff)}</span>`
        : '<span style="font-size:11px;color:var(--text-secondary)">—</span>';
      const statusHtml = diff != null
        ? (Math.abs(diff) < 0.01
          ? '<span class="badge badge-ok" style="font-size:10px"><i class="fas fa-circle-check"></i> OK</span>'
          : '<span class="badge badge-divergente" style="font-size:10px"><i class="fas fa-circle-xmark"></i> Divergente</span>')
        : '';

      return `<tr>
        <td><input style="width:130px;font-size:12px;border:1px solid var(--border);border-radius:4px;padding:3px 6px"
                   value="${U.esc(num)}" onchange="FaturaXls.updateRow(${i},'numero_vivo',this.value)"></td>
        <td><input style="width:220px;font-size:12px;border:1px solid var(--border);border-radius:4px;padding:3px 6px"
                   value="${U.esc(r.plano||'')}" onchange="FaturaXls.updateRow(${i},'plano',this.value)"></td>
        <td class="right"><input type="number" step="0.01"
                   style="width:90px;font-size:12px;border:1px solid var(--border);border-radius:4px;padding:3px 6px;text-align:right"
                   value="${v != null ? v : ''}" onchange="FaturaXls.updateRow(${i},'valor_total',parseFloat(this.value)||null)"></td>
        <td class="right" style="font-size:12px">${vContrato != null ? U.money(vContrato) : '—'}</td>
        <td class="right">${diffHtml}</td>
        <td style="text-align:center">${statusHtml}</td>
        <td><button onclick="FaturaXls.removeRow(${i})" style="background:none;border:none;cursor:pointer;color:#e53e3e" title="Remover">
          <i class="fas fa-trash"></i></button></td>
      </tr>`;
    }).join('');

    const warning = nothingFound ? `
      <div class="alert alert-warning" style="margin-bottom:10px">
        <i class="fas fa-triangle-exclamation"></i>
        <div><strong>Nenhum dado extraído automaticamente.</strong><br>
        O PDF pode ser uma imagem escaneada.
        <button class="btn-link" onclick="FaturaXls.toggleRaw()">Ver texto bruto extraído</button></div>
      </div>
      <div id="fxls-raw-wrap" style="display:none;margin-bottom:12px">
        <textarea readonly style="width:100%;height:160px;font-family:monospace;font-size:11px;border:1px solid var(--border);border-radius:6px;padding:8px;background:#f8fafc;resize:vertical">${U.esc(this.state.rawText || '(nenhum texto encontrado)')}</textarea>
      </div>` : '';

    resultEl.innerHTML = `
      ${warning}
      <div class="fatura-preview-header">
        <h3><i class="fas fa-table" style="color:#660099"></i> Preview — ${U.esc(res.filename || this.state.filename)}</h3>
        <div class="fatura-meta-badges">${badges}</div>
      </div>
      <div class="fatura-table-wrap">
        <table class="fatura-table" id="fxls-table">
          <thead><tr>
            <th>Número Vivo</th><th>Plano</th><th>Valor Fatura R$</th>
            <th style="text-align:right">Valor Contrato</th><th style="text-align:right">Diferença</th><th>Status</th><th></th>
          </tr></thead>
          <tbody id="fxls-tbody">${trs}</tbody>
          <tfoot><tr>
            <td colspan="2" style="text-align:right;padding-right:12px">TOTAL</td>
            <td class="right" id="fxls-total">${U.money(total)}</td>
            <td colspan="4"></td>
          </tr></tfoot>
        </table>
      </div>
      <div style="margin-top:10px">
        <button class="btn btn-ghost btn-sm" onclick="FaturaXls.addRow()">
          <i class="fas fa-plus"></i> Adicionar linha
        </button>
      </div>
      <div class="fatura-actions">
        <button class="btn-vivo" onclick="FaturaXls.importar()" id="btn-importar">
          <i class="fas fa-cloud-arrow-up"></i> Importar ao Sistema
        </button>
        <button class="btn btn-outline" onclick="FaturaXls.downloadXls()">
          <i class="fas fa-file-excel"></i> Baixar Excel (.xlsx)
        </button>
        <button class="btn btn-ghost" onclick="FaturaXls.render()">
          <i class="fas fa-arrow-rotate-left"></i> Novo arquivo
        </button>
      </div>
      <div id="fxls-import-result"></div>`;
  },

  toggleRaw() {
    const el = document.getElementById('fxls-raw-wrap');
    if (el) el.style.display = el.style.display === 'none' ? 'block' : 'none';
  },

  updateRow(i, field, value) {
    if (this.state.rows[i]) {
      this.state.rows[i][field] = value;
      this._recalcTotal();
    }
  },

  removeRow(i) {
    this.state.rows.splice(i, 1);
    this.renderPreview({ ...this.state.meta, linhas: this.state.rows, filename: this.state.filename, raw_text: this.state.rawText });
  },

  addRow() {
    this.state.rows.push({ numero_vivo: '', plano: '', valor_total: null });
    this.renderPreview({ ...this.state.meta, linhas: this.state.rows, filename: this.state.filename, raw_text: this.state.rawText });
  },

  _recalcTotal() {
    const total = this.state.rows.reduce((s, r) => s + (r.valor_total || 0), 0);
    const el = document.getElementById('fxls-total');
    if (el) el.textContent = U.money(total);
  },

  async importar() {
    const { rows, meta, filename } = this.state;
    if (!rows.length) { U.toast('Nenhuma linha para importar.', 'error'); return; }

    const btn = document.getElementById('btn-importar');
    if (btn) { btn.disabled = true; btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Importando…'; }

    try {
      const res = await API.post('/api/faturas/importar', {
        linhas: rows,
        meta: { ...meta, filename, competencia: meta.competencia || '' },
      });

      // Show result
      const importEl = document.getElementById('fxls-import-result');
      if (importEl && res.linhas) {
        const divergentes = res.linhas.filter(l => l.status === 'divergente');
        const oks = res.linhas.filter(l => l.status === 'ok');
        const semPlano = res.linhas.filter(l => !l.status);

        importEl.innerHTML = `
          <div class="alert alert-${divergentes.length > 0 ? 'warning' : 'success'}" style="margin-top:16px">
            <i class="fas fa-${divergentes.length > 0 ? 'triangle-exclamation' : 'circle-check'}"></i>
            <div>
              <strong>${res.importados} linhas importadas.</strong>
              ${oks.length > 0 ? `<span style="color:#22c55e;margin-left:8px">${oks.length} conformes</span>` : ''}
              ${divergentes.length > 0 ? `<span style="color:#ef4444;margin-left:8px">${divergentes.length} divergentes</span>` : ''}
              ${semPlano.length > 0 ? `<span style="color:var(--text-secondary);margin-left:8px">${semPlano.length} sem plano cadastrado</span>` : ''}
            </div>
          </div>
          ${divergentes.length > 0 ? `
          <div style="margin-top:12px">
            <h4 style="font-size:14px;font-weight:600;margin-bottom:8px;color:#ef4444">
              <i class="fas fa-circle-xmark"></i> Linhas Divergentes
            </h4>
            <table class="fatura-table">
              <thead><tr><th>Número</th><th>Plano</th><th style="text-align:right">Fatura</th><th style="text-align:right">Contrato</th><th style="text-align:right">Diferença</th></tr></thead>
              <tbody>${divergentes.map(l => `<tr>
                <td>${U.esc(l.numero_vivo)}</td>
                <td>${U.esc(l.plano)}</td>
                <td class="right">${U.money(l.valor_fatura)}</td>
                <td class="right">${U.money(l.valor_contrato)}</td>
                <td class="right val-pos" style="font-weight:600">${U.money(l.diferenca)}</td>
              </tr>`).join('')}</tbody>
            </table>
          </div>` : ''}`;
      }
      U.toast(`${res.importados} linhas importadas!`, 'success');
    } catch(e) {
      U.toast(e.message, 'error');
    } finally {
      if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fas fa-cloud-arrow-up"></i> Importar ao Sistema'; }
    }
  },

  async downloadXls() {
    const { rows, meta, filename } = this.state;
    if (!rows.length) { U.toast('Nenhuma linha para exportar.', 'error'); return; }

    const btn = document.querySelector('.btn.btn-outline');
    if (btn) { btn.disabled = true; btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Gerando…'; }

    try {
      const r = await fetch('/api/faturas/xls-from-json', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ linhas: rows, meta: { ...meta, filename } }),
      });
      if (!r.ok) {
        const d = await r.json().catch(() => ({}));
        throw new Error(d.error || `Erro HTTP ${r.status}`);
      }
      const blob = await r.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href     = url;
      a.download = (filename.replace(/\.pdf$/i, '') || 'fatura') + '.xlsx';
      a.click();
      URL.revokeObjectURL(url);
      U.toast('Excel gerado com sucesso!', 'success');
    } catch (err) {
      U.toast(err.message, 'error');
    } finally {
      if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fas fa-file-excel"></i> Baixar Excel (.xlsx)'; }
    }
  },
};
