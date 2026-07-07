'use strict';

/* ============================================================
   CST Data
   ============================================================ */
const CST_PIS_COFINS_SAIDAS = [
  ["01", "Operação Tributável com Alíquota Básica"],
  ["02", "Operação Tributável com Alíquota Diferenciada"],
  ["03", "Operação Tributável com Alíquota por Unidade de Medida de Produto"],
  ["04", "Operação Tributável Monofásica — Revenda a Alíquota Zero"],
  ["05", "Operação Tributável por Substituição Tributária"],
  ["06", "Operação Tributável a Alíquota Zero"],
  ["07", "Operação Isenta da Contribuição"],
  ["08", "Operação sem Incidência da Contribuição"],
  ["09", "Operação com Suspensão da Contribuição"],
  ["49", "Outras Operações de Saída"],
];

const CST_PIS_COFINS_ENTRADAS = [
  ["50", "Operação com Direito a Crédito — Vinculada Exclusivamente a Receita Tributada no Mercado Interno"],
  ["51", "Operação com Direito a Crédito — Vinculada Exclusivamente a Receita Não Tributada no MI"],
  ["52", "Operação com Direito a Crédito — Vinculada Exclusivamente a Receita de Exportação"],
  ["53", "Operação com Direito a Crédito — Vinculada a Receitas Tributadas e Não-Tributadas no MI"],
  ["54", "Operação com Direito a Crédito — Vinculada a Receitas Tributadas no MI e de Exportação"],
  ["55", "Operação com Direito a Crédito — Vinculada a Receitas Não-Tributadas no MI e de Exportação"],
  ["56", "Operação com Direito a Crédito — Vinculada a Receitas Tributadas e Não-Tributadas no MI e de Exportação"],
  ["60", "Crédito Presumido — Aquisição Vinculada Exclusivamente a Receita Tributada no MI"],
  ["61", "Crédito Presumido — Aquisição Vinculada Exclusivamente a Receita Não-Tributada no MI"],
  ["62", "Crédito Presumido — Aquisição Vinculada Exclusivamente a Receita de Exportação"],
  ["63", "Crédito Presumido — Aquisição Vinculada a Receitas Tributadas e Não-Tributadas no MI"],
  ["64", "Crédito Presumido — Aquisição Vinculada a Receitas Tributadas no MI e de Exportação"],
  ["65", "Crédito Presumido — Aquisição Vinculada a Receitas Não-Tributadas no MI e de Exportação"],
  ["66", "Crédito Presumido — Aquisição Vinculada a Receitas Tributadas e Não-Tributadas no MI e de Exportação"],
  ["67", "Crédito Presumido — Outras Operações"],
  ["70", "Operação de Aquisição sem Direito a Crédito"],
  ["71", "Operação de Aquisição com Isenção"],
  ["72", "Operação de Aquisição com Suspensão"],
  ["73", "Operação de Aquisição a Alíquota Zero"],
  ["74", "Operação de Aquisição sem Incidência da Contribuição"],
  ["75", "Operação de Aquisição por Substituição Tributária"],
  ["98", "Outras Operações de Entrada"],
  ["99", "Outras Operações"],
];

const CST_ICMS_NORMAL = [
  ["00", "Tributada integralmente"],
  ["10", "Tributada e com cobrança do ICMS por substituição tributária"],
  ["20", "Com redução de base de cálculo"],
  ["30", "Isenta ou não tributada e com cobrança do ICMS por substituição tributária"],
  ["40", "Isenta"],
  ["41", "Não tributada"],
  ["50", "Suspensão"],
  ["51", "Diferimento"],
  ["60", "ICMS cobrado anteriormente por substituição tributária"],
  ["70", "Com redução de base de cálculo e cobrança do ICMS por substituição tributária"],
  ["90", "Outras"],
];

const CST_ICMS_SN = [
  ["101", "Tributada pelo Simples Nacional com permissão de crédito"],
  ["102", "Tributada pelo Simples Nacional sem permissão de crédito"],
  ["103", "Isenção do ICMS no Simples Nacional para faixa de receita bruta"],
  ["201", "Tributada pelo Simples Nacional com permissão de crédito e com cobrança do ICMS por ST"],
  ["202", "Tributada pelo Simples Nacional sem permissão de crédito e com cobrança do ICMS por ST"],
  ["203", "Isenção do ICMS para faixa de receita bruta e com cobrança do ICMS por ST"],
  ["300", "Imune"],
  ["400", "Não tributada pelo Simples Nacional"],
  ["500", "ICMS cobrado anteriormente por substituição tributária ou por antecipação"],
  ["900", "Outros"],
];

const CST_IPI_ENTRADAS = [
  ["00", "Entrada com recuperação de crédito"],
  ["01", "Entrada tributada com alíquota zero"],
  ["02", "Entrada isenta"],
  ["03", "Entrada não tributada"],
  ["04", "Entrada imune"],
  ["05", "Entrada com suspensão"],
  ["49", "Outras entradas"],
];

const CST_IPI_SAIDAS = [
  ["50", "Saída tributada"],
  ["51", "Saída tributada com alíquota zero"],
  ["52", "Saída isenta"],
  ["53", "Saída não tributada"],
  ["54", "Saída imune"],
  ["55", "Saída com suspensão"],
  ["99", "Outras saídas"],
];

/* ============================================================
   State
   ============================================================ */
let currentNcmData = null;
let allMonofasicos = [];
let activeCategory = 'Todos';
let lastPdfReport = null;

/* ============================================================
   Tab switching
   ============================================================ */
function switchTab(tab) {
  document.querySelectorAll('.tab-panel').forEach(p => {
    p.classList.remove('active');
    p.classList.add('hidden');
  });
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));

  document.getElementById('tab-' + tab).classList.remove('hidden');
  document.getElementById('tab-' + tab).classList.add('active');
  document.getElementById('btn-tab-' + tab).classList.add('active');

  if (tab === 'monofasicos' && allMonofasicos.length === 0) loadMonofasicos();
  if (tab === 'classtrib' && allClassTrib.length === 0) loadClassTrib();
}

/* ============================================================
   NCM Input mask
   ============================================================ */
function maskNcm(input) {
  let v = input.value.replace(/\D/g, '').slice(0, 8);
  if (v.length > 6) v = v.slice(0,4) + '.' + v.slice(4,6) + '.' + v.slice(6);
  else if (v.length > 4) v = v.slice(0,4) + '.' + v.slice(4);
  input.value = v;
}

/* ============================================================
   Consulta NCM
   ============================================================ */
async function consultarNcm() {
  const raw = document.getElementById('ncm-input').value.replace(/\D/g, '');
  const errEl = document.getElementById('search-error');

  if (raw.length !== 8) {
    errEl.textContent = 'Por favor, informe um NCM com 8 dígitos.';
    errEl.classList.remove('hidden');
    return;
  }
  errEl.classList.add('hidden');

  document.getElementById('result-area').classList.add('hidden');
  document.getElementById('loading').classList.remove('hidden');

  try {
    const res = await fetch(`/api/ncm/${raw}`);
    const data = await res.json();

    if (!res.ok) {
      throw new Error(data.error || 'Erro ao consultar NCM');
    }

    currentNcmData = data;
    renderResult(data);
    document.getElementById('result-area').classList.remove('hidden');
    document.getElementById('result-area').scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (err) {
    errEl.textContent = err.message;
    errEl.classList.remove('hidden');
  } finally {
    document.getElementById('loading').classList.add('hidden');
  }
}

/* ============================================================
   Render result
   ============================================================ */
function renderResult(d) {
  // Header
  document.getElementById('res-ncm-fmt').textContent = d.ncm_formatado;
  document.getElementById('res-desc').textContent = d.descricao;

  let meta = [];
  if (d.tipo_ato) meta.push(`${d.tipo_ato} nº ${d.numero_ato}/${d.ano_ato}`);
  if (d.data_inicio) meta.push(`Vigência desde: ${formatDate(d.data_inicio)}`);
  if (d.data_fim && !d.data_fim.startsWith('9999')) meta.push(`Até: ${formatDate(d.data_fim)}`);
  else if (d.data_fim && d.data_fim.startsWith('9999')) meta.push('Em vigor');
  document.getElementById('res-meta').textContent = meta.join(' · ');

  const badgeEl = document.getElementById('badge-monofasico');
  if (d.monofasico) {
    badgeEl.innerHTML = `<span class="tag-monofasico">⚡ Monofásico · ${d.monofasico_dados?.categoria || ''}</span>`;
  } else {
    badgeEl.innerHTML = `<span class="tag-normal">📦 Tributação Normal</span>`;
  }

  document.getElementById('res-class-trib').innerHTML =
    `<div style="font-size:.7rem;opacity:.8;margin-bottom:3px">ClassTrib</div>${d.tributacao.class_trib}`;

  // IPI
  document.getElementById('ipi-aliq').textContent =
    d.tributacao.ipi.aliquota.toFixed(1).replace('.', ',') + '%';
  document.getElementById('ipi-cst').textContent = d.tributacao.cst_ipi;
  document.getElementById('ipi-detail').textContent = d.tributacao.ipi.regime;

  // PIS
  const pisLr = d.tributacao.pis.lucro_real;
  const pisLp = d.tributacao.pis.lucro_presumido;
  document.getElementById('pis-lr').textContent = pisLr.aliquota.toFixed(2).replace('.', ',') + '%';
  document.getElementById('pis-lp').textContent = pisLp.aliquota.toFixed(2).replace('.', ',') + '%';
  document.getElementById('pis-cst').textContent = pisLr.cst;
  toggleEl('pis-mono', d.tributacao.pis.monofasico);
  toggleEl('pis-zero', d.tributacao.pis.aliquota_zero);
  const pisZeroEl = document.getElementById('pis-zero');
  if (d.tributacao.pis.aliquota_zero) pisZeroEl.textContent = '✅ ' + d.tributacao.pis.nota_aliquota_zero;

  // COFINS
  const cofLr = d.tributacao.cofins.lucro_real;
  const cofLp = d.tributacao.cofins.lucro_presumido;
  document.getElementById('cofins-lr').textContent = cofLr.aliquota.toFixed(2).replace('.', ',') + '%';
  document.getElementById('cofins-lp').textContent = cofLp.aliquota.toFixed(2).replace('.', ',') + '%';
  document.getElementById('cofins-cst').textContent = cofLr.cst;
  toggleEl('cofins-mono', d.tributacao.cofins.monofasico);
  toggleEl('cofins-zero', d.tributacao.cofins.aliquota_zero);
  const cofZeroEl = document.getElementById('cofins-zero');
  if (d.tributacao.cofins.aliquota_zero) cofZeroEl.textContent = '✅ ' + d.tributacao.cofins.nota_aliquota_zero;

  // CST tables
  fillCstTable('cst-pc-saidas', CST_PIS_COFINS_SAIDAS, pisLr.cst);
  fillCstTable('cst-pc-entradas', CST_PIS_COFINS_ENTRADAS, null);
  fillCstTable('cst-icms-normal', CST_ICMS_NORMAL, null);
  fillCstTable('cst-icms-sn', CST_ICMS_SN, null);
  fillCstTable('cst-ipi-entradas', CST_IPI_ENTRADAS, null);
  fillCstTable('cst-ipi-saidas', CST_IPI_SAIDAS, d.tributacao.cst_ipi);

  // ICMS por estado
  fillIcmsTable(d.icms_estados, d.icms_piaui);
  renderIcmsPI(d.icms_piaui);
  const icmsAvisoEl = document.getElementById('icms-hortifruti-aviso');
  if (d.icms_hortifruti_aviso) {
    icmsAvisoEl.innerHTML = `<span style="font-size:1.1rem;line-height:1.2">⚠</span><span>${d.icms_hortifruti_aviso}</span>`;
    icmsAvisoEl.classList.remove('hidden');
  } else {
    icmsAvisoEl.classList.add('hidden');
  }

  // Seção CST-IBS/CBS e cClassTrib
  renderClassTribSection(d.classtrib_sugestao);

  // Reforma Tributária (cards resumo — mantido abaixo da tabela ICMS)
  renderReforma(d.reforma_tributaria);
}

/* ============================================================
   Seção: Reforma Tributária — CST-IBS/CBS e cClassTrib
   ============================================================ */
const RT_TIPO_CLASS = {
  'Padrão':                     'tipo-Padrão',
  'Fixa':                       'tipo-Fixa',
  'Sem alíquota':               'tipo-Sem-alíquota',
  'Uniforme nacional (referência)': 'tipo-Uniforme-nacional',
  'Uniforme setorial':          'tipo-Uniforme-setorial',
};

let _rtShowAll = false;
let _rtList    = [];

function renderClassTribSection(sugestao) {
  if (!sugestao) return;
  _rtShowAll = false;
  _rtList    = sugestao.classtrib_list || [];

  // Cabeçalho
  document.getElementById('rt-cst-code').textContent = sugestao.cst_sugerido || '—';
  document.getElementById('rt-cst-desc').textContent = sugestao.desc_cst || '';
  document.getElementById('rt-motivo').textContent   = sugestao.motivo || '';

  const baseLegal = document.getElementById('rt-base-legal');
  baseLegal.innerHTML = sugestao.base_legal
    ? `📖 ${sugestao.base_legal}`
    : '';

  // CST block color por tipo
  _setCstBlockColor(sugestao.cst_sugerido);

  // Contagem
  document.getElementById('rt-ct-count').textContent =
    `${_rtList.length} classificaç${_rtList.length === 1 ? 'ão' : 'ões'}`;

  // Botão "Ver todos" só quando > 8
  const toggleBtn = document.getElementById('rt-toggle-btn');
  if (_rtList.length > 8) {
    toggleBtn.style.display = '';
    toggleBtn.textContent = `Ver todos (${_rtList.length})`;
    _rtShowAll = false;
  } else {
    toggleBtn.style.display = 'none';
    _rtShowAll = true;
  }

  _renderRtTable(sugestao.classtrib_destaque);
  _renderRtDestaque(sugestao.classtrib_destaque);
}

function _setCstBlockColor(cst) {
  const block = document.getElementById('rt-cst-block');
  const code  = document.getElementById('rt-cst-code');
  const colors = {
    '000': ['#3730A3', '#818CF8'],
    '200': ['#065F46', '#34D399'],
    '400': ['#92400E', '#F59E0B'],
    '410': ['#92400E', '#F59E0B'],
    '620': ['#991B1B', '#F87171'],
    '510': ['#374151', '#9CA3AF'],
    '550': ['#374151', '#9CA3AF'],
  };
  const [textColor, borderColor] = colors[cst] || ['#3730A3', '#818CF8'];
  block.style.borderColor = borderColor;
  code.style.color = textColor;
}

function _renderRtTable(destaqueCode) {
  const tbody = document.getElementById('rt-ct-tbody');
  if (!tbody) return;

  const rows = _rtList.map((r, idx) => {
    const isDestaque = r.cClassTrib === destaqueCode;
    const hidden     = !_rtShowAll && !isDestaque && idx >= 8 ? ' rt-row-hidden' : '';
    const destClass  = isDestaque ? ' rt-row-destaque' : '';

    const badgeDestaque = isDestaque
      ? `<span class="rt-badge-destaque">★ Mais provável</span>`
      : '';

    const tipoCls = RT_TIPO_CLASS[r.tipo_aliquota] || '';
    const tipoHtml = r.tipo_aliquota
      ? `<span class="tipo-badge ${tipoCls}">${r.tipo_aliquota}</span>`
      : '—';

    const redIbs = r.pRedIBS > 0
      ? `<span class="rt-pct-val">${(r.pRedIBS * 100).toFixed(0)}%</span>`
      : `<span class="rt-pct-zero">—</span>`;
    const redCbs = r.pRedCBS > 0
      ? `<span class="rt-pct-val">${(r.pRedCBS * 100).toFixed(0)}%</span>`
      : `<span class="rt-pct-zero">—</span>`;

    const nfe = r.ind_nfe
      ? `<span style="color:var(--green-lt);font-weight:700">✓</span>`
      : `<span style="color:var(--gray-300)">—</span>`;

    const artigo = r.lc_214 || '—';
    const vigencia = r.d_ini_vig
      ? `<span style="font-size:.75rem">${r.d_ini_vig}</span>`
      : '—';

    return `<tr class="${destClass}${hidden}" data-code="${r.cClassTrib}">
      <td>
        <span class="rt-ct-code" onclick="expandRtRow('${r.cClassTrib}')" title="Ver detalhes">${r.cClassTrib}</span>
        ${badgeDestaque}
      </td>
      <td style="font-size:.83rem;line-height:1.35">${r.nome}</td>
      <td>${tipoHtml}</td>
      <td style="text-align:center">${redIbs}</td>
      <td style="text-align:center">${redCbs}</td>
      <td style="text-align:center">${nfe}</td>
      <td>${vigencia}</td>
      <td style="font-size:.75rem;color:var(--blue-lt)">${artigo}</td>
    </tr>`;
  });

  tbody.innerHTML = rows.join('') ||
    '<tr><td colspan="8" style="text-align:center;padding:20px;color:var(--gray-400)">Nenhuma classificação encontrada para este CST</td></tr>';
}

function _renderRtDestaque(destaqueCode) {
  const container = document.getElementById('rt-destaque-detail');
  if (!container || !destaqueCode) { container.innerHTML = ''; return; }

  const item = _rtList.find(r => r.cClassTrib === destaqueCode);
  if (!item) { container.innerHTML = ''; return; }

  const isMono   = item.ind_mono_padrao || item.ind_mono_reten || item.ind_mono_ret || item.ind_mono_dif;
  const isRedIbs = item.pRedIBS > 0;
  const isRedCbs = item.pRedCBS > 0;

  const flagsHtml = [
    ['Tributação Regular',  item.ind_trib_regular, false],
    ['Mono. Padrão',        item.ind_mono_padrao,  true],
    ['Mono. Retenção',      item.ind_mono_reten,   true],
    ['Mono. Retido',        item.ind_mono_ret,     true],
    ['Mono. Diferido',      item.ind_mono_dif,     true],
    ['Créd. Presumido',     item.ind_cred_pres,    false],
    ['Estorno de Crédito',  item.ind_estorno_cred, false],
    ['NF-e',                item.ind_nfe,          false],
    ['NFC-e',               item.ind_nfce,         false],
    ['CT-e',                item.ind_cte,          false],
  ].filter(([, v]) => v || false)
    .map(([label, v, isMono]) => {
      const cls = !v ? 'off' : isMono ? 'mono-on' : 'on';
      const icon = isMono ? '⚡' : '✓';
      return `<span class="rt-flag-chip ${cls}">${icon} ${label}</span>`;
    }).join('');

  const redBox = (isRedIbs || isRedCbs) ? `
    <div style="display:flex;gap:12px;margin-top:8px;flex-wrap:wrap">
      <div style="background:#F0FDF4;border:1px solid #A7F3D0;border-radius:6px;padding:8px 16px;text-align:center">
        <div style="font-size:.68rem;font-weight:700;color:var(--gray-400);margin-bottom:2px">Red. IBS</div>
        <div style="font-size:1.2rem;font-weight:900;color:var(--green-lt)">${item.pRedIBS > 0 ? (item.pRedIBS*100).toFixed(0)+'%' : '—'}</div>
      </div>
      <div style="background:#EFF6FF;border:1px solid #BFDBFE;border-radius:6px;padding:8px 16px;text-align:center">
        <div style="font-size:.68rem;font-weight:700;color:var(--gray-400);margin-bottom:2px">Red. CBS</div>
        <div style="font-size:1.2rem;font-weight:900;color:var(--blue-lt)">${item.pRedCBS > 0 ? (item.pRedCBS*100).toFixed(0)+'%' : '—'}</div>
      </div>
    </div>` : '';

  const linkHtml = item.link
    ? `<a class="rt-lc-link" href="${item.link}" target="_blank" rel="noopener">🔗 Ver artigo na LC 214/2025</a>`
    : '';

  container.innerHTML = `
    <div class="rt-destaque-box">
      <div class="rt-destaque-box-header">
        <span class="rt-destaque-code-big">${item.cClassTrib}</span>
        <span class="rt-destaque-nome">${item.nome}</span>
        <span class="rt-badge-destaque">★ Mais provável</span>
      </div>
      <div class="rt-destaque-desc">${item.descricao || ''}</div>
      ${redBox}
      ${flagsHtml ? `<div class="rt-destaque-flags" style="margin-top:10px">${flagsHtml}</div>` : ''}
      ${item.lc_214 ? `<div style="font-size:.8rem;color:#4338CA;font-weight:600;margin-top:8px">Artigo: ${item.lc_214}</div>` : ''}
      ${linkHtml ? `<div style="margin-top:10px">${linkHtml}</div>` : ''}
    </div>`;
}

function toggleAllClassTribs() {
  _rtShowAll = !_rtShowAll;
  const btn = document.getElementById('rt-toggle-btn');
  document.querySelectorAll('.rt-row-hidden').forEach(tr => tr.classList.remove('rt-row-hidden'));
  if (!_rtShowAll) {
    // Recolhe: esconde linhas a partir do índice 8 que não são destaque
    document.querySelectorAll('#rt-ct-tbody tr').forEach((tr, idx) => {
      if (idx >= 8 && !tr.classList.contains('rt-row-destaque')) {
        tr.classList.add('rt-row-hidden');
      }
    });
    btn.textContent = `Ver todos (${_rtList.length})`;
  } else {
    btn.textContent = 'Recolher';
  }
}

function expandRtRow(code) {
  const item = _rtList.find(r => r.cClassTrib === code);
  if (!item) return;
  // Reutiliza o painel de destaque mas sem o badge
  const isMono = item.ind_mono_padrao || item.ind_mono_reten || item.ind_mono_ret || item.ind_mono_dif;
  const flagsHtml = [
    ['Tributação Regular', item.ind_trib_regular, false],
    ['Mono. Padrão',       item.ind_mono_padrao,  true],
    ['Mono. Retenção',     item.ind_mono_reten,   true],
    ['Mono. Retido',       item.ind_mono_ret,     true],
    ['Mono. Diferido',     item.ind_mono_dif,     true],
    ['Créd. Presumido',    item.ind_cred_pres,    false],
    ['NF-e',               item.ind_nfe,          false],
    ['NFC-e',              item.ind_nfce,         false],
    ['CT-e',               item.ind_cte,          false],
  ].filter(([, v]) => v)
    .map(([label,, mono]) => `<span class="rt-flag-chip ${mono?'mono-on':'on'}">${mono?'⚡':'✓'} ${label}</span>`)
    .join('');

  const linkHtml = item.link
    ? `<a class="rt-lc-link" href="${item.link}" target="_blank" rel="noopener">🔗 Ver artigo na LC 214/2025</a>`
    : '';

  const container = document.getElementById('rt-destaque-detail');
  container.innerHTML = `
    <div class="rt-destaque-box">
      <div class="rt-destaque-box-header">
        <span class="rt-destaque-code-big">${item.cClassTrib}</span>
        <span class="rt-destaque-nome">${item.nome}</span>
      </div>
      <div class="rt-destaque-desc">${item.descricao || '—'}</div>
      ${item.pRedIBS > 0 || item.pRedCBS > 0 ? `
        <div style="display:flex;gap:10px;margin-top:8px">
          <div style="background:#F0FDF4;border:1px solid #A7F3D0;border-radius:6px;padding:6px 14px;text-align:center">
            <div style="font-size:.65rem;font-weight:700;color:var(--gray-400)">Red. IBS</div>
            <div style="font-size:1.1rem;font-weight:900;color:var(--green-lt)">${item.pRedIBS > 0 ? (item.pRedIBS*100).toFixed(0)+'%' : '—'}</div>
          </div>
          <div style="background:#EFF6FF;border:1px solid #BFDBFE;border-radius:6px;padding:6px 14px;text-align:center">
            <div style="font-size:.65rem;font-weight:700;color:var(--gray-400)">Red. CBS</div>
            <div style="font-size:1.1rem;font-weight:900;color:var(--blue-lt)">${item.pRedCBS > 0 ? (item.pRedCBS*100).toFixed(0)+'%' : '—'}</div>
          </div>
        </div>` : ''}
      ${flagsHtml ? `<div class="rt-destaque-flags" style="margin-top:8px">${flagsHtml}</div>` : ''}
      ${item.lc_214 ? `<div style="font-size:.8rem;color:#4338CA;font-weight:600;margin-top:6px">Artigo: ${item.lc_214}</div>` : ''}
      ${linkHtml ? `<div style="margin-top:8px">${linkHtml}</div>` : ''}
    </div>`;
  container.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function fillCstTable(tbodyId, rows, highlight) {
  const tb = document.getElementById(tbodyId);
  if (!tb) return;
  tb.innerHTML = rows.map(([cst, desc]) => {
    const hl = (highlight && cst === highlight) ? ' cst-highlight' : '';
    return `<tr class="${hl}"><td class="ncm-code">${cst}</td><td>${desc}</td></tr>`;
  }).join('');
}

function fillIcmsTable(states, piauiData) {
  const tb = document.getElementById('icms-tbody');
  if (!tb) return;
  tb.innerHTML = states.map(s => {
    if (s.uf === 'PI' && piauiData) {
      const aliq = piauiData.aliquota;
      const aliqFmt = aliq === 0 ? 'Isento' : `${aliq.toFixed(1).replace('.', ',')}%`;
      const aliqColor = aliq >= 27 ? '#DC2626' : aliq >= 25 ? '#D97706' : aliq <= 0 ? '#059669' : aliq <= 12 ? '#059669' : '#1D4ED8';
      const stBadge = piauiData.tem_st
        ? `<span style="margin-left:5px;background:#FEF3C7;color:#92400E;border:1px solid #FCD34D;font-size:.6rem;font-weight:700;padding:1px 5px;border-radius:3px;vertical-align:middle;display:inline-block">ST</span>`
        : '';
      return `<tr style="background:#EFF6FF;border-left:3px solid #1D4ED8">
        <td><strong style="color:#1D4ED8">${s.uf}</strong></td>
        <td><strong>${s.estado}</strong></td>
        <td class="aliq-cell">
          <span style="font-weight:900;font-size:1rem;color:${aliqColor}">${aliqFmt}</span>${stBadge}
          <div style="font-size:.65rem;color:var(--gray-400);margin-top:2px;font-weight:400">${piauiData.categoria}</div>
        </td>
        <td class="aliq-cell">${s.aliq_interestadual_sul_sudeste.toFixed(0)}%</td>
        <td class="aliq-cell">${s.aliq_interestadual_outros.toFixed(0)}%</td>
        <td class="aliq-cell">${s.aliq_importado.toFixed(0)}%</td>
        <td><span style="font-size:.72rem;color:#1D4ED8;font-weight:600">${piauiData.base_legal}</span></td>
      </tr>`;
    }
    const obs = s.obs ? `<span style="font-size:.75rem;color:var(--orange)">${s.obs}</span>` : '—';
    return `<tr>
      <td><strong>${s.uf}</strong></td>
      <td>${s.estado}</td>
      <td class="aliq-cell">${s.aliq_interna.toFixed(1).replace('.', ',')}%</td>
      <td class="aliq-cell">${s.aliq_interestadual_sul_sudeste.toFixed(0)}%</td>
      <td class="aliq-cell">${s.aliq_interestadual_outros.toFixed(0)}%</td>
      <td class="aliq-cell">${s.aliq_importado.toFixed(0)}%</td>
      <td>${obs}</td>
    </tr>`;
  }).join('');
}

function renderIcmsPI(data) {
  const card = document.getElementById('icms-pi-detail-card');
  const content = document.getElementById('icms-pi-content');
  if (!card || !content || !data) return;

  card.style.display = '';

  const aliq = data.aliquota;
  const aliqFmt = aliq === 0 ? 'Isento' : `${aliq.toFixed(1).replace('.', ',')}%`;
  const aliqColor = aliq >= 27 ? '#DC2626' : aliq >= 25 ? '#D97706' : aliq <= 0 ? '#059669' : aliq <= 12 ? '#059669' : '#1D4ED8';
  const aliqBg   = aliq >= 27 ? '#FEF2F2' : aliq >= 25 ? '#FFFBEB' : aliq <= 12 ? '#F0FDF4' : '#EFF6FF';

  const stHtml = data.tem_st
    ? `<div style="display:flex;align-items:flex-start;gap:8px;background:#FEF3C7;border:1px solid #FCD34D;border-radius:8px;padding:12px 14px;margin-top:12px">
        <span style="font-size:1.1rem;line-height:1.4">⚠</span>
        <div>
          <div style="font-weight:700;color:#92400E;font-size:.85rem">Substituição Tributária (ST) — Aplicável</div>
          ${data.obs_st ? `<div style="font-size:.78rem;color:#92400E;margin-top:3px">${data.obs_st}</div>` : ''}
        </div>
      </div>`
    : `<div style="display:flex;align-items:center;gap:8px;background:#F0FDF4;border:1px solid #A7F3D0;border-radius:8px;padding:10px 14px;margin-top:12px">
        <span style="color:#059669;font-weight:700">✓</span>
        <div>
          <div style="font-weight:700;color:#065F46;font-size:.85rem">Substituição Tributária (ST) — Não identificada para este produto</div>
          ${data.obs_st ? `<div style="font-size:.78rem;color:#065F46;margin-top:2px">${data.obs_st}</div>` : ''}
        </div>
      </div>`;

  const faixasHtml = data.faixas.map(f => {
    const isActive = Math.abs(f.aliquota - data.aliquota) < 0.01;
    const fFmt = f.aliquota === 0 ? 'Isento' : `${f.aliquota.toFixed(1).replace('.', ',')}%`;
    const rowStyle = isActive ? 'background:#EFF6FF;' : '';
    const aliqStyle = isActive ? 'font-weight:900;font-size:1rem;color:#1D4ED8' : 'font-weight:500;color:var(--gray-600)';
    const badge = isActive
      ? `<span style="background:#1D4ED8;color:#fff;border-radius:4px;font-size:.62rem;font-weight:700;padding:2px 7px;white-space:nowrap">ESTE PRODUTO</span>`
      : '';
    return `<tr style="${rowStyle}">
      <td style="${aliqStyle};padding:8px 12px;white-space:nowrap">${fFmt}</td>
      <td style="padding:8px 12px;font-size:.82rem;color:var(--gray-600)">${f.descricao}</td>
      <td style="padding:8px 12px;font-size:.75rem;color:var(--gray-400)">${f.caps}</td>
      <td style="padding:8px 12px;text-align:right">${badge}</td>
    </tr>`;
  }).join('');

  content.innerHTML = `
    <div style="display:flex;flex-wrap:wrap;gap:16px;align-items:flex-start">
      <div style="flex:0 0 auto;text-align:center;background:${aliqBg};border:2px solid ${aliqColor};border-radius:12px;padding:20px 28px;min-width:150px">
        <div style="font-size:.68rem;font-weight:700;color:var(--gray-400);margin-bottom:6px;text-transform:uppercase;letter-spacing:.05em">Alíquota Interna PI</div>
        <div style="font-size:2.6rem;font-weight:900;color:${aliqColor};line-height:1">${aliqFmt}</div>
        <div style="font-size:.7rem;color:var(--gray-500);margin-top:6px;max-width:130px">${data.categoria}</div>
      </div>
      <div style="flex:1;min-width:240px">
        ${stHtml}
        <div style="font-size:.77rem;color:var(--gray-400);margin-top:10px;line-height:1.5">
          <strong>Base legal:</strong> ${data.base_legal}<br>
          <strong>Vigência:</strong> a partir de ${data.vigencia}
        </div>
      </div>
    </div>
    <div style="margin-top:20px">
      <div style="font-size:.75rem;font-weight:700;color:var(--gray-500);text-transform:uppercase;letter-spacing:.04em;margin-bottom:8px">Tabela de Alíquotas ICMS — Piauí (Lei nº 8.558/2024)</div>
      <div class="table-wrapper">
        <table class="data-table" style="font-size:.82rem">
          <thead>
            <tr>
              <th style="width:80px">Alíquota</th>
              <th>Categoria de Produto</th>
              <th style="width:160px">Capítulos NCM</th>
              <th style="width:130px;text-align:right">Situação</th>
            </tr>
          </thead>
          <tbody>${faixasHtml}</tbody>
        </table>
      </div>
    </div>
  `;
}

function renderReforma(rt) {
  const grid = document.getElementById('reforma-grid');
  if (!grid) return;

  const items = [];

  // CBS
  items.push(`
    <div class="reforma-item">
      <div class="reforma-item-header">
        <span class="reforma-tag cbs">CBS</span>
        <span class="reforma-item-title">Contribuição sobre Bens e Serviços</span>
      </div>
      <div class="reforma-aliq">${
        rt.cbs.aliquota_efetiva !== undefined
          ? rt.cbs.aliquota_efetiva.toFixed(2).replace('.', ',') + '%'
          : rt.cbs.aliquota_referencia.toFixed(1).replace('.', ',') + '%'
      }</div>
      <div class="reforma-desc">${rt.cbs.descricao}</div>
      <div class="reforma-vigencia">${rt.cbs.vigencia}</div>
    </div>`);

  // IBS
  items.push(`
    <div class="reforma-item">
      <div class="reforma-item-header">
        <span class="reforma-tag ibs">IBS</span>
        <span class="reforma-item-title">Imposto sobre Bens e Serviços</span>
      </div>
      <div class="reforma-aliq">${
        rt.ibs.aliquota_efetiva !== undefined
          ? rt.ibs.aliquota_efetiva.toFixed(2).replace('.', ',') + '%'
          : rt.ibs.aliquota_referencia.toFixed(1).replace('.', ',') + '%'
      }</div>
      <div class="reforma-desc">${rt.ibs.descricao}</div>
      <div class="reforma-vigencia">${rt.ibs.vigencia}</div>
    </div>`);

  // IS
  if (rt.is.incide) {
    items.push(`
      <div class="reforma-item">
        <div class="reforma-item-header">
          <span class="reforma-tag is">IS</span>
          <span class="reforma-item-title">Imposto Seletivo</span>
        </div>
        <div class="reforma-aliq" style="color:var(--red-lt)">Incide ⚠</div>
        <div class="reforma-desc">${rt.is.descricao}</div>
        <div class="reforma-desc" style="margin-top:6px">
          ${rt.is.produtos_sujeitos.map(p => `• ${p}`).join('<br>')}
        </div>
      </div>`);
  }

  // Reduções / Isenções
  if (rt.reducao_aliquota) {
    items.push(`
      <div class="reforma-item">
        <div class="reforma-item-header">
          <span class="reforma-tag obs">Redução</span>
          <span class="reforma-item-title">Alíquota Reduzida</span>
        </div>
        <div class="reforma-aliq" style="color:var(--green-lt)">-${rt.reducao_aliquota}%</div>
        <div class="reforma-desc">Este produto tem redução de ${rt.reducao_aliquota}% na alíquota de referência do CBS e IBS.</div>
      </div>`);
  }

  if (rt.isencao) {
    items.push(`
      <div class="reforma-item">
        <div class="reforma-item-header">
          <span class="reforma-tag ibs">Isenção</span>
          <span class="reforma-item-title">Alíquota Zero</span>
        </div>
        <div class="reforma-aliq" style="color:var(--green-lt)">0,00%</div>
        <div class="reforma-desc">${rt.isencao}</div>
      </div>`);
  }

  // Observações
  if (rt.observacoes && rt.observacoes.length > 0) {
    items.push(`
      <div class="reforma-item">
        <div class="reforma-item-header">
          <span class="reforma-tag obs">Obs.</span>
          <span class="reforma-item-title">Considerações</span>
        </div>
        <ul class="obs-list">
          ${rt.observacoes.map(o => `<li>${o}</li>`).join('')}
        </ul>
      </div>`);
  }

  grid.innerHTML = items.join('');
}

/* ============================================================
   CST Tab switching
   ============================================================ */
function showCstTab(tab) {
  document.querySelectorAll('.cst-tab').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.cst-content').forEach(c => c.classList.remove('active'));
  document.querySelector(`[onclick="showCstTab('${tab}')"]`).classList.add('active');
  document.getElementById('cst-' + tab).classList.add('active');
}

/* ============================================================
   Monofásicos Table
   ============================================================ */
async function loadMonofasicos() {
  try {
    const res = await fetch('/api/monofasicos');
    allMonofasicos = await res.json();
    buildCategoryFilters();
    renderMonofasicosTable(allMonofasicos);
    document.getElementById('stat-total').textContent = allMonofasicos.length;
  } catch (e) {
    console.error(e);
  }
}

function buildCategoryFilters() {
  const cats = ['Todos', ...new Set(allMonofasicos.map(m => m.categoria))];
  const container = document.getElementById('category-filters');
  container.innerHTML = cats.map(c => `
    <button class="cat-btn ${c === 'Todos' ? 'active' : ''}" onclick="filterByCategory('${c}')">${c}</button>
  `).join('');
}

function filterByCategory(cat) {
  activeCategory = cat;
  document.querySelectorAll('.cat-btn').forEach(b => {
    b.classList.toggle('active', b.textContent === cat);
  });
  applyFilters();
}

function filterMonofasicos() {
  applyFilters();
}

function applyFilters() {
  const q = document.getElementById('mono-search').value.toLowerCase();
  let filtered = allMonofasicos;

  if (activeCategory !== 'Todos') {
    filtered = filtered.filter(m => m.categoria === activeCategory);
  }

  if (q) {
    filtered = filtered.filter(m =>
      m.ncm.includes(q) ||
      m.descricao.toLowerCase().includes(q) ||
      m.categoria.toLowerCase().includes(q) ||
      m.referencia.toLowerCase().includes(q)
    );
  }

  renderMonofasicosTable(filtered);
}

function renderMonofasicosTable(items) {
  const tbody = document.getElementById('mono-tbody');
  const catClass = cat => 'cat-badge cat-' + cat.replace(/\s*e\s*/g, 'e').replace(/\s+/g, '');
  tbody.innerHTML = items.map(m => `
    <tr>
      <td class="ncm-code">${formatNcmDisplay(m.ncm)}</td>
      <td>${m.descricao}</td>
      <td><span class="${catClass(m.categoria)}">${m.categoria}</span></td>
      <td style="font-size:.78rem;color:var(--gray-600)">${m.referencia}</td>
    </tr>
  `).join('');
  document.getElementById('mono-count').textContent = `${items.length} de ${allMonofasicos.length} registros`;
}

/* ============================================================
   PDF Upload
   ============================================================ */
function handleFileSelect(event) {
  const file = event.target.files[0];
  if (file) processPdf(file);
}

function handleDrop(event) {
  event.preventDefault();
  document.getElementById('upload-area').classList.remove('drag-over');
  const file = event.dataTransfer.files[0];
  if (file && file.name.toLowerCase().endsWith('.pdf')) {
    processPdf(file);
  }
}

async function processPdf(file) {
  document.getElementById('pdf-result').classList.add('hidden');
  document.getElementById('pdf-loading').classList.remove('hidden');

  // Atualiza nome do arquivo na área de upload
  const uploadText = document.querySelector('.upload-text');
  if (uploadText) uploadText.textContent = `📄 ${file.name}`;

  const formData = new FormData();
  formData.append('file', file);

  try {
    const res = await fetch('/api/upload-pdf', { method: 'POST', body: formData });
    const data = await res.json();

    if (!res.ok) throw new Error(data.error || 'Erro ao processar PDF');
    if (data.total_ncms === 0) {
      showPdfWarning(data.aviso || 'Nenhum NCM encontrado no arquivo.');
      return;
    }

    lastPdfReport = data;
    renderPdfReport(data);
    document.getElementById('pdf-result').classList.remove('hidden');
    document.getElementById('pdf-result').scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (err) {
    showPdfWarning('Erro: ' + err.message);
  } finally {
    document.getElementById('pdf-loading').classList.add('hidden');
  }
}

function showPdfWarning(msg) {
  const area = document.getElementById('upload-area');
  let warn = document.getElementById('pdf-warn');
  if (!warn) {
    warn = document.createElement('div');
    warn.id = 'pdf-warn';
    warn.className = 'error-msg';
    warn.style.marginTop = '10px';
    area.insertAdjacentElement('afterend', warn);
  }
  warn.textContent = msg;
  warn.classList.remove('hidden');
}

function renderPdfReport(d) {
  // Remove aviso anterior se existir
  const warn = document.getElementById('pdf-warn');
  if (warn) warn.classList.add('hidden');

  document.getElementById('rpt-total').textContent = d.total_ncms;
  document.getElementById('rpt-mono').textContent  = d.total_monofasicos;
  document.getElementById('rpt-nao').textContent   = d.total_nao_monofasicos;
  document.getElementById('rpt-pct').textContent   = d.percentual_monofasico.toFixed(1) + '%';

  // Tabela de monofásicos
  // Exibe código NCM, descrição e categoria da tabela oficial — NÃO usa dados do PDF
  const catClass = cat => 'cat-badge cat-' + cat.replace(/\s*e\s*/g, 'e').replace(/\s+/g, '');
  const monoTb = document.getElementById('rpt-mono-tbody');
  monoTb.innerHTML = d.monofasicos.length === 0
    ? '<tr><td colspan="5" style="text-align:center;color:var(--gray-400);padding:20px">Nenhum NCM monofásico encontrado</td></tr>'
    : d.monofasicos.map(m => `
        <tr>
          <td class="ncm-code">${m.ncm_formatado || formatNcmDisplay(m.ncm)}</td>
          <td style="font-size:.83rem">${m.descricao || ''}</td>
          <td><span class="${catClass(m.categoria)}">${m.categoria}</span></td>
          <td style="font-size:.75rem;color:var(--gray-500)">${m.referencia}</td>
          <td style="font-size:.72rem;color:var(--gray-500);font-style:italic">${m.regra_identificacao || 'Monofásico por NCM específico'}</td>
        </tr>`).join('');

  // Tabela de não-monofásicos
  // nao_monofasicos agora é lista de {ncm, ncm_formatado}
  const naoTb = document.getElementById('rpt-nao-tbody');
  naoTb.innerHTML = d.nao_monofasicos.length === 0
    ? '<tr><td colspan="2" style="text-align:center;color:var(--gray-400);padding:20px">Nenhum</td></tr>'
    : d.nao_monofasicos.map(item => {
        const fmt = (typeof item === 'object') ? item.ncm_formatado : formatNcmDisplay(item);
        const raw = (typeof item === 'object') ? item.ncm : item;
        return `<tr>
          <td class="ncm-code">${fmt}</td>
          <td style="font-size:.82rem;color:var(--gray-500)">Não consta na lista de monofásicos</td>
        </tr>`;
      }).join('');
}

/* ============================================================
   Export Excel
   ============================================================ */
async function exportExcel() {
  if (!lastPdfReport) return;

  try {
    const res = await fetch('/api/export-excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(lastPdfReport),
    });

    if (!res.ok) throw new Error('Erro ao gerar Excel');

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'relatorio_ncm_monofasicos.xlsx';
    a.click();
    URL.revokeObjectURL(url);
  } catch (err) {
    alert('Erro ao exportar: ' + err.message);
  }
}

function resetPdf() {
  lastPdfReport = null;
  document.getElementById('pdf-result').classList.add('hidden');
  document.getElementById('pdf-input').value = '';
}

/* ============================================================
   Helpers
   ============================================================ */
function formatNcmDisplay(ncm) {
  if (ncm.length === 8)
    return `${ncm.slice(0,4)}.${ncm.slice(4,6)}.${ncm.slice(6)}`;
  return ncm;
}

function formatDate(dateStr) {
  if (!dateStr) return '';
  const [y, m, d] = dateStr.split('-');
  return `${d}/${m}/${y}`;
}

function toggleEl(id, show) {
  const el = document.getElementById(id);
  if (!el) return;
  if (show) el.classList.remove('hidden');
  else el.classList.add('hidden');
}

/* ============================================================
   cClassTrib Tab
   ============================================================ */
let allClassTrib = [];
let allCstIbs = [];
let activeCstFilter = '';
let activeTipoFilter = '';
let ctDetailData = null;

async function loadClassTrib() {
  if (allClassTrib.length > 0) return;
  try {
    const [r1, r2] = await Promise.all([
      fetch('/api/cst-ibs-cbs'),
      fetch('/api/classtrib'),
    ]);
    allCstIbs = await r1.json();
    const d = await r2.json();
    allClassTrib = d.data;
    renderCstIbsTable();
    buildCtFilters();
    renderClassTribTable(allClassTrib);
  } catch (e) {
    console.error('Erro ao carregar ClassTrib:', e);
  }
}

function renderCstIbsTable() {
  const tb = document.getElementById('cst-ibs-tbody');
  if (!tb) return;

  const cstColors = {
    '000': '#DBEAFE', '010': '#EDE9FE', '011': '#EDE9FE',
    '200': '#D1FAE5', '220': '#D1FAE5', '221': '#D1FAE5', '222': '#D1FAE5',
    '400': '#FEF3C7', '410': '#FEF3C7',
    '510': '#F3F4F6', '515': '#F3F4F6', '550': '#F3F4F6',
    '620': '#FEE2E2',
    '800': '#F0F4FF', '810': '#F0F4FF', '811': '#F0F4FF',
    '820': '#FDF4FF', '830': '#FDF4FF',
  };

  tb.innerHTML = allCstIbs.map(row => {
    const bg = cstColors[row.cst] || '#fff';
    const flag = v => `<span class="flag-dot ${v ? 'on' : 'off'}" title="${v ? 'Sim' : 'Não'}"></span>`;
    return `<tr style="background:${bg}">
      <td><span class="cst-ibs-badge">${row.cst}</span></td>
      <td style="font-weight:600">${row.descricao}</td>
      <td style="text-align:center">${flag(row.ind_ibs_cbs)}</td>
      <td style="text-align:center">${flag(row.ind_mono)}</td>
      <td style="text-align:center">${flag(row.ind_red)}</td>
      <td style="text-align:center">${flag(row.ind_dif)}</td>
      <td style="text-align:center">${flag(row.ind_transf_cred)}</td>
    </tr>`;
  }).join('');
}

function buildCtFilters() {
  // CST filters
  const csts = [...new Set(allClassTrib.map(r => r.cst_ibs_cbs))].sort();
  const cstContainer = document.getElementById('ct-cst-filters');
  if (cstContainer) {
    cstContainer.innerHTML =
      `<button class="ct-filter-btn active" onclick="setCstFilter('')">Todos CST</button>` +
      csts.map(c => {
        const desc = allCstIbs.find(x => x.cst === c)?.descricao || '';
        const short = desc.split(' ').slice(0,2).join(' ');
        return `<button class="ct-filter-btn" onclick="setCstFilter('${c}')" title="${desc}">CST ${c}</button>`;
      }).join('');
  }

  // Tipo filters
  const tipos = [...new Set(allClassTrib.map(r => r.tipo_aliquota).filter(Boolean))].sort();
  const tipoContainer = document.getElementById('ct-tipo-filters');
  if (tipoContainer) {
    tipoContainer.innerHTML =
      `<button class="ct-filter-btn active" onclick="setTipoFilter('')">Todos Tipos</button>` +
      tipos.map(t => `<button class="ct-filter-btn" onclick="setTipoFilter('${t}')">${t}</button>`).join('');
  }
}

function setCstFilter(val) {
  activeCstFilter = val;
  document.querySelectorAll('#ct-cst-filters .ct-filter-btn').forEach(b => {
    b.classList.toggle('active', (val === '' && b.textContent === 'Todos CST') || b.textContent === `CST ${val}`);
  });
  filterClassTrib();
}

function setTipoFilter(val) {
  activeTipoFilter = val;
  document.querySelectorAll('#ct-tipo-filters .ct-filter-btn').forEach(b => {
    b.classList.toggle('active', (val === '' && b.textContent === 'Todos Tipos') || b.textContent === val);
  });
  filterClassTrib();
}

function filterClassTrib() {
  const q = (document.getElementById('ct-search')?.value || '').toLowerCase();
  const monoOnly = document.getElementById('ct-mono-only')?.checked;
  let filtered = allClassTrib;

  if (activeCstFilter) filtered = filtered.filter(r => r.cst_ibs_cbs === activeCstFilter);
  if (activeTipoFilter) filtered = filtered.filter(r => r.tipo_aliquota === activeTipoFilter);
  if (monoOnly) filtered = filtered.filter(r => r.ind_mono_padrao || r.ind_mono_reten || r.ind_mono_ret || r.ind_mono_dif);
  if (q) filtered = filtered.filter(r =>
    r.cClassTrib.includes(q) ||
    r.nome.toLowerCase().includes(q) ||
    r.descricao.toLowerCase().includes(q) ||
    r.cst_ibs_cbs.includes(q) ||
    r.lc_214.toLowerCase().includes(q)
  );

  renderClassTribTable(filtered);
}

function renderClassTribTable(items) {
  const tbody = document.getElementById('ct-tbody');
  if (!tbody) return;

  document.getElementById('ct-count-badge').textContent = `${items.length} classificações`;
  document.getElementById('ct-table-count').textContent = `${items.length} de ${allClassTrib.length} registros`;

  const tipoClass = t => 'tipo-badge tipo-' + (t || '').replace(/\s+/g, '-').replace(/\(/g,'').replace(/\)/g,'');

  tbody.innerHTML = items.map(r => {
    const flags = buildFlagsHtml(r);
    const redIbs = r.pRedIBS > 0 ? `<span style="color:var(--green-lt);font-weight:700">${(r.pRedIBS*100).toFixed(0)}%</span>` : '<span style="color:var(--gray-300)">—</span>';
    const redCbs = r.pRedCBS > 0 ? `<span style="color:var(--green-lt);font-weight:700">${(r.pRedCBS*100).toFixed(0)}%</span>` : '<span style="color:var(--gray-300)">—</span>';
    const vigencia = r.d_ini_vig ? `A partir de ${r.d_ini_vig}${r.d_fim_vig ? '<br><small>Até '+r.d_fim_vig+'</small>' : ''}` : '—';
    const lcRef = r.lc_214 ? `<span style="font-size:.75rem;color:var(--blue-lt)">${r.lc_214}</span>` : '—';
    return `<tr>
      <td><span class="ct-code" onclick="showDetail('${r.cClassTrib}')">${r.cClassTrib}</span></td>
      <td><span class="cst-ibs-badge">${r.cst_ibs_cbs}</span></td>
      <td style="font-size:.83rem;max-width:280px">${r.nome}</td>
      <td><span class="${tipoClass(r.tipo_aliquota)}">${r.tipo_aliquota || '—'}</span></td>
      <td style="text-align:center">${redIbs}</td>
      <td style="text-align:center">${redCbs}</td>
      <td class="flags-cell">${flags}</td>
      <td style="font-size:.75rem">${vigencia}</td>
      <td>${lcRef}</td>
    </tr>`;
  }).join('') || '<tr><td colspan="9" style="text-align:center;padding:24px;color:var(--gray-400)">Nenhum resultado encontrado</td></tr>';
}

function buildFlagsHtml(r) {
  const flags = [
    [r.ind_trib_regular, 'Trib. Regular', '🟢'],
    [r.ind_mono_padrao,  'Mono. Padrão',  '⚡'],
    [r.ind_mono_reten,   'Mono. Retenção','⚡'],
    [r.ind_mono_ret,     'Mono. Ret.',    '⚡'],
    [r.ind_mono_dif,     'Mono. Dif.',    '⚡'],
    [r.ind_cred_pres,    'Créd. Pres.',   '💰'],
    [r.ind_estorno_cred, 'Estorno Créd.', '↩'],
  ];
  return flags
    .filter(([v]) => v)
    .map(([, label, icon]) => `<span title="${label}" style="font-size:.85rem;margin-right:2px">${icon}</span>`)
    .join('') || '<span style="color:var(--gray-300);font-size:.75rem">—</span>';
}

function showDetail(code) {
  const item = allClassTrib.find(r => r.cClassTrib === code);
  if (!item) return;
  ctDetailData = item;

  document.getElementById('ct-detail-code').textContent = item.cClassTrib;
  document.getElementById('ct-detail-name').textContent = item.nome;

  const redIbs = item.pRedIBS > 0 ? `${(item.pRedIBS * 100).toFixed(0)}%` : '0%';
  const redCbs = item.pRedCBS > 0 ? `${(item.pRedCBS * 100).toFixed(0)}%` : '0%';

  const docFlags = [
    ['NF-e', item.ind_nfe],
    ['NFC-e', item.ind_nfce],
    ['CT-e', item.ind_cte],
  ];
  const monoFlags = [
    ['Padrão', item.ind_mono_padrao],
    ['Retenção', item.ind_mono_reten],
    ['Ret.', item.ind_mono_ret],
    ['Dif.', item.ind_mono_dif],
  ];

  document.getElementById('ct-detail-body').innerHTML = `
    <div class="ct-detail-section">
      <div class="ct-detail-section-title">Identificação</div>
      <table style="font-size:.82rem;width:100%;border-collapse:collapse">
        <tr><td style="padding:3px 0;color:var(--gray-400);width:130px">cClassTrib</td><td style="font-family:monospace;font-weight:700;color:var(--blue)">${item.cClassTrib}</td></tr>
        <tr><td style="padding:3px 0;color:var(--gray-400)">CST-IBS/CBS</td><td><span class="cst-ibs-badge">${item.cst_ibs_cbs}</span> ${item.desc_cst}</td></tr>
        <tr><td style="padding:3px 0;color:var(--gray-400)">Tipo Alíquota</td><td>${item.tipo_aliquota || '—'}</td></tr>
        <tr><td style="padding:3px 0;color:var(--gray-400)">Artigo LC 214</td><td style="color:var(--blue-lt);font-weight:600">${item.lc_214 || '—'}</td></tr>
        <tr><td style="padding:3px 0;color:var(--gray-400)">Vigência</td><td>${item.d_ini_vig || '—'}${item.d_fim_vig ? ' até ' + item.d_fim_vig : ''}</td></tr>
        <tr><td style="padding:3px 0;color:var(--gray-400)">Atualização</td><td>${item.d_atualizacao || '—'}</td></tr>
      </table>
      ${item.link ? `<a class="ct-link" href="${item.link}" target="_blank">🔗 Ver na legislação</a>` : ''}
    </div>

    <div class="ct-detail-section">
      <div class="ct-detail-section-title">Reduções de Alíquota</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
        <div>
          <div style="font-size:.75rem;color:var(--gray-400);font-weight:700;margin-bottom:4px">Red. IBS</div>
          <div style="font-size:1.3rem;font-weight:900;color:var(--green-lt)">${redIbs}</div>
          <div class="pct-bar"><div class="pct-bar-fill" style="width:${item.pRedIBS*100}%"></div></div>
        </div>
        <div>
          <div style="font-size:.75rem;color:var(--gray-400);font-weight:700;margin-bottom:4px">Red. CBS</div>
          <div style="font-size:1.3rem;font-weight:900;color:var(--blue-lt)">${redCbs}</div>
          <div class="pct-bar"><div class="pct-bar-fill" style="width:${item.pRedCBS*100}%;background:var(--blue-lt)"></div></div>
        </div>
      </div>
      <div style="margin-top:14px">
        <div class="ct-detail-section-title">Tributação Monofásica</div>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          ${monoFlags.map(([l, v]) => `<span style="padding:3px 10px;border-radius:10px;font-size:.75rem;font-weight:700;background:${v?'var(--orange-bg)':'var(--gray-100)'};color:${v?'var(--orange)':'var(--gray-400)'}">⚡ ${l}</span>`).join('')}
        </div>
      </div>
      <div style="margin-top:14px">
        <div class="ct-detail-section-title">Documentos Fiscais</div>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          ${docFlags.map(([l, v]) => `<span style="padding:3px 10px;border-radius:10px;font-size:.75rem;font-weight:700;background:${v?'var(--green-bg)':'var(--gray-100)'};color:${v?'var(--green-lt)':'var(--gray-400)'}">${l}</span>`).join('')}
        </div>
      </div>
    </div>

    <div class="ct-detail-section" style="grid-column:1/-1">
      <div class="ct-detail-section-title">Descrição</div>
      <div class="ct-detail-text">${item.descricao || '—'}</div>
    </div>

    ${item.lc_redacao ? `
    <div class="ct-detail-section" style="grid-column:1/-1">
      <div class="ct-detail-section-title">Redação LC 214/2025</div>
      <div class="ct-detail-lc">${item.lc_redacao}</div>
    </div>` : ''}
  `;

  const panel = document.getElementById('ct-detail-panel');
  panel.classList.remove('hidden');
  panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function closeDetail() {
  document.getElementById('ct-detail-panel').classList.add('hidden');
}

/* ============================================================
   Init
   ============================================================ */
document.addEventListener('DOMContentLoaded', () => {
  if (document.getElementById('tab-monofasicos').classList.contains('active')) {
    loadMonofasicos();
  }
});
