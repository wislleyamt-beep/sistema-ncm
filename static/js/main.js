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
  if (tab === 'operacoes') loadCadastroOperacoes();
  if (tab === 'servicos') loadServicosInfo();
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

  const monoAvisoEl = document.getElementById('mono-aviso');
  const monoAviso = d.monofasico_dados?.aviso || d.queijo_aviso;
  if (monoAviso) {
    monoAvisoEl.innerHTML = `<span style="font-size:1.1rem;line-height:1.2">⚠</span><span>${monoAviso}</span>`;
    monoAvisoEl.classList.remove('hidden');
  } else {
    monoAvisoEl.classList.add('hidden');
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

  // Operação Questor (Simples Nacional)
  renderOperacaoQuestor(d.operacao_questor);

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

  const avisoEl = document.getElementById('rt-aviso-confirmacao');
  if (sugestao.aviso_confirmacao) {
    avisoEl.innerHTML = `<span style="font-size:1.1rem;line-height:1.2">⚠</span><span>${sugestao.aviso_confirmacao}</span>`;
    avisoEl.classList.remove('hidden');
  } else {
    avisoEl.classList.add('hidden');
  }

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
      ? `<span class="rt-pct-val">${r.pRedIBS.toFixed(0)}%</span>`
      : `<span class="rt-pct-zero">—</span>`;
    const redCbs = r.pRedCBS > 0
      ? `<span class="rt-pct-val">${r.pRedCBS.toFixed(0)}%</span>`
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
    ? '<tr><td colspan="6" style="text-align:center;color:var(--gray-400);padding:20px">Nenhum NCM monofásico encontrado</td></tr>'
    : d.monofasicos.map(m => `
        <tr>
          <td class="ncm-code">${m.ncm_formatado || formatNcmDisplay(m.ncm)}</td>
          <td style="font-size:.83rem">${m.descricao || ''}</td>
          <td><span class="${catClass(m.categoria)}">${m.categoria}</span></td>
          <td style="font-size:.75rem;color:var(--gray-500)">${m.referencia}</td>
          <td style="font-size:.72rem;color:var(--gray-500);font-style:italic">${m.regra_identificacao || 'Monofásico por NCM específico'}</td>
          <td>${_opqCellHtml(m)}</td>
        </tr>`).join('');

  // Tabela de não-monofásicos
  // nao_monofasicos agora é lista de {ncm, ncm_formatado}
  const naoTb = document.getElementById('rpt-nao-tbody');
  naoTb.innerHTML = d.nao_monofasicos.length === 0
    ? '<tr><td colspan="3" style="text-align:center;color:var(--gray-400);padding:20px">Nenhum</td></tr>'
    : d.nao_monofasicos.map(item => {
        const fmt = (typeof item === 'object') ? item.ncm_formatado : formatNcmDisplay(item);
        const raw = (typeof item === 'object') ? item.ncm : item;
        return `<tr>
          <td class="ncm-code">${fmt}</td>
          <td style="font-size:.82rem;color:var(--gray-500)">Não consta na lista de monofásicos</td>
          <td>${(typeof item === 'object') ? _opqCellHtml(item) : ''}</td>
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

  const redIbs = item.pRedIBS > 0 ? `${item.pRedIBS.toFixed(0)}%` : '0%';
  const redCbs = item.pRedCBS > 0 ? `${item.pRedCBS.toFixed(0)}%` : '0%';

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


/* ============================================================
   Operação Questor — Simples Nacional
   ============================================================ */
let opqOperacoes = [];
let opqCadastro = [];

function _esc(v) {
  return String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function _opqCellHtml(item) {
  const src = item.operacao_origem === 'cadastro' ? 'cadastro manual' : 'regra automática';
  const aviso = item.operacao_aviso ? ` title="${_esc(item.operacao_aviso)}"` : '';
  return `<div class="opq-cell-op"${aviso}><strong>${_esc(item.operacao_codigo || '—')}</strong> · ${_esc(item.operacao_rotulo || '')}
    <span class="opq-src">${src} · ICMS: ${_esc(item.icms_situacao || '')}${item.operacao_aviso ? ' · ⚠ ver aviso' : ''}</span></div>`;
}

function renderOperacaoQuestor(oq) {
  const card = document.getElementById('opq-card');
  if (!oq) { card.style.display = 'none'; return; }
  card.style.display = '';
  const op = oq.operacao;
  const manual = oq.origem === 'cadastro';
  card.classList.toggle('origem-cadastro', manual);

  document.getElementById('opq-op').textContent = op ? `${op.operacao} - ${op.descricao}` : 'Definir manualmente';
  document.getElementById('opq-code').textContent = op ? op.codigo : '—';

  let origem;
  if (manual) {
    const m = oq.manual;
    origem = `✅ Cadastro manual (NCM ${_esc(m.ncm_cadastrado)})` +
      (m.observacao ? ` · ${_esc(m.observacao)}` : '') +
      (m.atualizado_em ? ` · ${_esc(m.atualizado_em)}` : '');
  } else {
    const conf = {alta: 'confiança alta', media: 'confiança média', baixa: 'confiança baixa: confira'}[oq.confianca] || '';
    origem = `🤖 Regra automática (${conf})`;
  }
  document.getElementById('opq-origem').innerHTML = origem;

  document.getElementById('opq-pc').textContent = oq.pis_cofins_label;
  document.getElementById('opq-icms').textContent = oq.icms_label;
  document.getElementById('opq-icms-legal').textContent = oq.icms_base_legal || '';

  const avisos = [...(oq.avisos || [])];
  if (oq.divergencia) avisos.unshift(oq.divergencia);
  document.getElementById('opq-avisos').innerHTML = avisos.map(a =>
    `<div class="icms-aviso"><span style="font-size:1.1rem;line-height:1.2">⚠</span><span>${_esc(a)}</span></div>`
  ).join('');
}

async function abrirCadastroComNcm() {
  const ncm = currentNcmData?.ncm_formatado || '';
  switchTab('operacoes');
  await loadCadastroOperacoes();
  window.scrollTo({ top: 0, behavior: 'smooth' });
  document.getElementById('opq-f-ncm').value = ncm;
  const codigo = currentNcmData?.operacao_questor?.operacao?.codigo;
  if (codigo) document.getElementById('opq-f-op').value = codigo;
  document.getElementById('opq-f-op').focus();
}

function _adminHeaders() {
  let token = '';
  try { token = localStorage.getItem('opq_admin_token') || ''; } catch (e) {}
  return { 'Content-Type': 'application/json', 'X-Admin-Token': token };
}

async function _opqFetchWrite(url, options) {
  let res = await fetch(url, { ...options, headers: _adminHeaders() });
  if (res.status === 401) {
    const t = prompt('Senha de administrador para alterar o cadastro:');
    if (t === null) return res;
    try { localStorage.setItem('opq_admin_token', t); } catch (e) {}
    res = await fetch(url, { ...options, headers: { 'Content-Type': 'application/json', 'X-Admin-Token': t } });
  }
  return res;
}

async function loadCadastroOperacoes() {
  try {
    if (opqOperacoes.length === 0) {
      opqOperacoes = await (await fetch('/api/operacoes-questor')).json();
      const sel = document.getElementById('opq-f-op');
      sel.innerHTML = '<option value="">Selecione a operação…</option>' + opqOperacoes.map(o =>
        `<option value="${o.codigo}">${o.codigo} · ${_esc(o.operacao)} - ${_esc(o.descricao)}</option>`).join('');
    }
    const data = await (await fetch('/api/operacoes-ncm')).json();
    opqCadastro = data.itens || [];
    const aviso = document.getElementById('opq-persist-aviso');
    if (!data.persistente) {
      aviso.innerHTML = '<span style="font-size:1.1rem;line-height:1.2">⚠</span><span>O cadastro está sendo salvo em arquivo local do servidor. ' +
        'No plano gratuito do Render ele é <strong>apagado a cada reinício ou deploy</strong>. Configure a variável <code>DATABASE_URL</code> ' +
        '(Postgres gratuito do Neon ou Supabase) para guardar de forma permanente.</span>';
      aviso.classList.remove('hidden');
    } else {
      aviso.classList.add('hidden');
    }
    renderCadastroOperacoes();
  } catch (err) {
    _opqMsg('Erro ao carregar cadastro: ' + err.message, true);
  }
}

function renderCadastroOperacoes() {
  const q = (document.getElementById('opq-search').value || '').toLowerCase().replace(/\./g, '');
  const itens = opqCadastro.filter(i => !q ||
    i.ncm.includes(q) || (i.descricao + i.observacao + i.operacao + i.codigo).toLowerCase().replace(/\./g, '').includes(q));
  document.getElementById('opq-count-badge').textContent = `${opqCadastro.length} cadastrados`;
  const tb = document.getElementById('opq-tbody');
  if (itens.length === 0) {
    tb.innerHTML = '<tr><td colspan="6" style="text-align:center;color:var(--gray-400);padding:20px">Nenhum NCM cadastrado</td></tr>';
    return;
  }
  tb.innerHTML = itens.map(i => {
    const auto = i.automatica;
    const diverge = auto && auto.codigo !== i.codigo;
    const autoTxt = auto
      ? `<span class="${diverge ? 'opq-divergente' : ''}">${diverge ? '⚠ ' : '✓ '}${_esc(auto.codigo)} · ${_esc(auto.operacao)}</span>`
      : '<span style="color:var(--gray-400)">prefixo</span>';
    return `<tr>
      <td class="ncm-code">${_esc(i.ncm_formatado)}</td>
      <td class="opq-cell-op"><strong>${_esc(i.codigo)}</strong> · ${_esc(i.operacao)} - ${_esc(i.descricao)}</td>
      <td class="opq-cell-op">${autoTxt}</td>
      <td style="font-size:.8rem">${_esc(i.observacao)}</td>
      <td style="font-size:.75rem;color:var(--gray-500)">${_esc(i.atualizado_em)}</td>
      <td><button class="opq-del" title="Excluir" onclick="excluirCadastroOperacao('${_esc(i.ncm)}')">🗑</button></td>
    </tr>`;
  }).join('');
}

function _opqMsg(msg, erro) {
  const el = document.getElementById('opq-f-msg');
  el.className = erro ? 'error-msg' : 'ok-msg';
  el.innerHTML = msg;
  el.classList.remove('hidden');
}

async function _enviarCadastro(itens) {
  const res = await _opqFetchWrite('/api/operacoes-ncm', { method: 'POST', body: JSON.stringify({ itens }) });
  const data = await res.json().catch(() => ({}));
  if (!res.ok && !data.salvos) {
    _opqMsg(_esc(data.error || (data.erros || []).join('<br>') || 'Erro ao salvar'), true);
    return false;
  }
  let msg = `${data.salvos} registro(s) salvo(s).`;
  if (data.erros && data.erros.length) msg += '<br>' + data.erros.map(_esc).join('<br>');
  _opqMsg(msg, !!(data.erros && data.erros.length));
  await loadCadastroOperacoes();
  return true;
}

async function salvarCadastroOperacao() {
  const ncm = document.getElementById('opq-f-ncm').value;
  const codigo = document.getElementById('opq-f-op').value;
  const observacao = document.getElementById('opq-f-obs').value;
  if (!codigo) { _opqMsg('Selecione a operação.', true); return; }
  if (await _enviarCadastro([{ ncm, codigo, observacao }])) {
    document.getElementById('opq-f-ncm').value = '';
    document.getElementById('opq-f-obs').value = '';
  }
}

async function salvarLoteOperacoes() {
  const linhas = document.getElementById('opq-lote-txt').value.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const itens = linhas.map(l => {
    const [ncm, codigo, ...obs] = l.split(/[;\t]/);
    return { ncm: (ncm || '').trim(), codigo: (codigo || '').trim(), observacao: obs.join(';').trim() };
  });
  if (!itens.length) { _opqMsg('Cole ao menos uma linha.', true); return; }
  if (await _enviarCadastro(itens)) document.getElementById('opq-lote-txt').value = '';
}

async function excluirCadastroOperacao(ncm) {
  const res = await _opqFetchWrite('/api/operacoes-ncm/' + encodeURIComponent(ncm), { method: 'DELETE' });
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    _opqMsg(_esc(data.error || 'Erro ao excluir'), true);
    return;
  }
  _opqMsg(`NCM ${_esc(ncm)} removido do cadastro.`, false);
  await loadCadastroOperacoes();
}


/* ============================================================
   Serviços NFS-e Nacional
   ============================================================ */
let svcDetalheAtual = null;
let svcNbsSel = 0;
let svcInfoCarregada = false;

const _fmtBRL = v => (v ?? 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const _fmtPct = v => ((v ?? 0) * 100).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '%';

async function loadServicosInfo() {
  if (svcInfoCarregada) return;
  try {
    const r = await fetch('/api/servicos/info');
    const d = await r.json();
    document.getElementById('svc-fonte').innerHTML =
      `Base: Anexo VIII da NFS-e v${_esc(d.versao_anexo_viii)} · ${d.itens} itens da LC 116 · ${d.nbs} correlações NBS · ${d.cnaes} CNAEs de serviço`;
    svcInfoCarregada = true;
  } catch (e) { /* silencioso */ }
}

function _svcBadgeSimples(simples) {
  if (!simples) return '';
  if (simples.impeditiva && simples.codigos.length === 1)
    return '<span class="svc-badge svc-badge-red">Impeditiva ao Simples</span>';
  const partes = [];
  if (simples.fator_r) partes.push('<span class="svc-badge svc-badge-orange">Fator R · III ou V</span>');
  simples.codigos.filter(c => !c.fator_r).forEach(c => {
    const cls = c.codigo === 'IMP' ? 'svc-badge-red' : (c.anexo === 'IV' ? 'svc-badge-purple' : 'svc-badge-green');
    const txt = c.codigo === 'IMP' ? 'Impeditiva (em parte)' : 'Anexo ' + c.anexo;
    if (!partes.some(p => p.includes(txt))) partes.push(`<span class="svc-badge ${cls}">${_esc(txt)}</span>`);
  });
  return partes.join(' ');
}

async function buscarServico(qForcada) {
  const input = document.getElementById('svc-input');
  const q = (qForcada ?? input.value).trim();
  if (qForcada) input.value = qForcada;
  const err = document.getElementById('svc-erro');
  err.classList.add('hidden');
  if (!q) { err.textContent = 'Digite uma atividade, CNAE, item da LC 116 ou NBS.'; err.classList.remove('hidden'); return; }
  const r = await fetch('/api/servicos/buscar?q=' + encodeURIComponent(q));
  const d = await r.json();
  const box = document.getElementById('svc-resultados');
  const lista = document.getElementById('svc-lista');
  const cnaeInfo = document.getElementById('svc-cnae-info');
  const rel = document.getElementById('svc-cnaes-rel');
  document.getElementById('svc-detalhe').classList.add('hidden');
  box.classList.remove('hidden');
  cnaeInfo.innerHTML = '';
  rel.innerHTML = '';

  if (d.tipo === 'cnae') {
    cnaeInfo.innerHTML = `<div class="svc-cnae-head"><strong>CNAE ${_esc(d.cnae.cnae)}</strong> ${_esc(d.cnae.descricao || '')}</div>` +
      (d.aviso ? `<div class="svc-aviso">${_esc(d.aviso)}</div>` : '') +
      (d.itens.length > 1 ? '<p class="section-desc">Esse CNAE pode corresponder a mais de um item da LC 116. Escolha o serviço que a empresa presta de fato.</p>' : '');
  }
  document.getElementById('svc-res-titulo').textContent =
    d.tipo === 'cnae' ? 'Itens da LC 116 para o CNAE' : d.tipo === 'nbs' ? 'Itens da LC 116 com essa NBS' : 'Serviços encontrados';
  document.getElementById('svc-res-count').textContent = d.itens.length + (d.itens.length === 1 ? ' item' : ' itens');

  if (!d.itens.length && !(d.cnaes || []).length) {
    lista.innerHTML = d.aviso ? '' : '<p class="section-desc">Nada encontrado. Tente outra palavra (ex.: o nome do serviço em vez da profissão) ou o código CNAE.</p>';
    return;
  }
  lista.innerHTML = d.itens.map(it => `
    <button class="svc-item" onclick="abrirItemServico('${_esc(it.item)}')">
      <div class="svc-item-cod">${_esc(it.item)}</div>
      <div class="svc-item-body">
        <div class="svc-item-desc">${_esc(it.descricao)}</div>
        <div class="svc-item-meta">cTribNac ${_esc(it.ctribnac.join(', ') || '—')} · ${it.qtd_nbs} NBS ${_svcBadgeSimples(it.simples)}</div>
      </div>
    </button>`).join('');
  if ((d.cnaes || []).length) {
    rel.innerHTML = '<div class="svc-subtitle">CNAEs com esse termo</div><div class="svc-chips">' +
      d.cnaes.map(c => `<button class="svc-chip" onclick="buscarServico('${_esc(c.cnae)}')" title="${_esc(c.descricao)}">${_esc(c.cnae)} · ${_esc(c.descricao.slice(0, 60))}</button>`).join('') + '</div>';
  }
  if (d.itens.length === 1) abrirItemServico(d.itens[0].item);
}

async function abrirItemServico(item) {
  const r = await fetch('/api/servicos/item/' + encodeURIComponent(item));
  if (!r.ok) return;
  svcDetalheAtual = await r.json();
  svcNbsSel = 0;
  // pré-seleciona a NBS cuja descrição mais se parece com o que foi digitado
  const norm = t => t.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  const termos = norm(document.getElementById('svc-input').value).split(/\s+/).filter(t => t.length > 3 && !/\d/.test(t));
  if (termos.length) {
    let melhor = 0;
    svcDetalheAtual.nbs.forEach((x, i) => {
      const d = norm(x.descricao);
      const pts = termos.filter(t => d.includes(t) || d.includes(t.slice(0, 5))).length;
      if (pts > melhor) { melhor = pts; svcNbsSel = i; }
    });
  }
  renderDetalheServico();
  document.getElementById('svc-detalhe').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function selecionarNbs(i) { svcNbsSel = i; renderDetalheServico(); }

function _copiavel(valor) {
  if (!valor) return '—';
  return `<span class="svc-copy" title="Clique para copiar" onclick="navigator.clipboard.writeText('${_esc(valor)}');this.classList.add('ok');setTimeout(()=>this.classList.remove('ok'),900)">${_esc(valor)}</span>`;
}

function renderDetalheServico() {
  const d = svcDetalheAtual;
  const n = d.nbs[svcNbsSel] || {};
  const el = document.getElementById('svc-detalhe');
  el.classList.remove('hidden');

  const ctrib = d.ctribnac.length > 1
    ? d.ctribnac.map(c => `<div>${_copiavel(c.codigo)} <span class="svc-mini">${_esc(c.descricao)}</span></div>`).join('')
    : _copiavel(d.ctribnac[0]?.codigo) + ` <span class="svc-mini">${_esc(d.ctribnac[0]?.descricao || '')}</span>`;

  const sn = n.ativ_sn_info;
  let snHtml = '—';
  if (sn) {
    snHtml = `<strong>${_esc(sn.rotulo)}</strong>` + (sn.codigo !== 'IMP' ? ` <span class="svc-mini">(cAtivSN ${_esc(sn.codigo)})</span>` : '');
    if (n.ativ_sn_alternativos_info?.length)
      snHtml += '<div class="svc-mini">Alternativa: ' + n.ativ_sn_alternativos_info.map(a => `${_esc(a.rotulo)} (cAtivSN ${_esc(a.codigo)})`).join('; ') + '</div>';
  }
  const alt = (n.cclasstrib_alternativos || []).length
    ? '<div class="svc-mini">Outras possíveis: ' + n.cclasstrib_alternativos.map(a => `<strong>${_esc(a.codigo)}</strong> ${_esc(a.nome)}`).join(' · ') + '</div>' : '';
  const indAlt = (n.indop_alternativos || []).length
    ? '<div class="svc-mini">Conforme o local da prestação: ' + n.indop_alternativos.map(a => `<strong>${_esc(a.codigo)}</strong> ${_esc(a.descricao || a.local_ibs || '')}`).join(' · ') + '</div>' : '';
  const semClass = !n.cclasstrib
    ? '<div class="svc-mini svc-warn">O Anexo VIII deixou o cClassTrib em branco para esta NBS. Confira na aba cClassTrib se o serviço tem redução (ex.: saúde, Anexo III da LC 214/2025, cClassTrib 200029) ou se fica em 000001.</div>' : '';
  const regra = n.indop_regra
    ? `<div class="svc-mini">Adquirente no exterior: <strong>${_esc(n.indop_regra.exterior)}</strong> · operação não onerosa: <strong>${_esc(n.indop_regra.nao_onerosa)}</strong></div>` : '';

  const avisos = [];
  if (d.simples.fator_r) avisos.push('Atividade sujeita ao <strong>fator R</strong>: Anexo III se a folha dos 12 meses for 28% ou mais da receita; abaixo disso, Anexo V. Use a calculadora abaixo.');
  if (d.simples.varia_por_nbs) avisos.push('O enquadramento no Simples <strong>muda conforme a NBS</strong> escolhida. Confira a coluna Simples na tabela.');
  if (d.simples.impeditiva) avisos.push('Há NBS deste item marcada como <strong>atividade impeditiva</strong> ao Simples Nacional.');
  if (d.iss.local_prestacao) avisos.push('ISS devido no <strong>local da prestação</strong> (LC 116/2003, art. 3º), e não no município do prestador.');
  if (d.iss.retencao_tomador) avisos.push('Tomador pessoa jurídica é <strong>responsável pela retenção do ISS</strong> (LC 116/2003, art. 6º, § 2º, II). Confira também a lei do município.');
  if (d.inss) avisos.push(`Sujeito à <strong>retenção de ${_esc(d.inss.aliquota)} de INSS</strong> quando prestado mediante cessão de mão de obra ou empreitada (${_esc(d.inss.base_legal)}). EFD-Reinf: ${_esc(d.inss.codigo_reinf || '—')}.`);
  if ((n.cclasstrib || '').startsWith('2')) avisos.push(`cClassTrib ${_esc(n.cclasstrib)} traz <strong>redução de alíquota</strong> de IBS/CBS: ${_esc(n.nome_cclasstrib)}. Confirme se o serviço e o prestador cumprem as condições da LC 214/2025.`);

  el.innerHTML = `
    <div class="card">
      <div class="card-title-row">
        <h3 class="card-title">Item ${_esc(d.item)} · ${_esc(d.descricao)}</h3>
        <span class="card-badge">${_svcBadgeSimples(d.simples)}</span>
      </div>

      <div class="svc-campos">
        <div class="svc-campo"><div class="svc-campo-lbl">Código de tributação nacional (cTribNac)</div><div class="svc-campo-val">${ctrib}</div></div>
        <div class="svc-campo"><div class="svc-campo-lbl">Item da LC 116</div><div class="svc-campo-val">${_copiavel(d.item)}</div></div>
        <div class="svc-campo"><div class="svc-campo-lbl">NBS</div><div class="svc-campo-val">${_copiavel(n.nbs)} <span class="svc-mini">${_esc(n.descricao || '')}</span></div></div>
        <div class="svc-campo"><div class="svc-campo-lbl">Indicador da operação (cIndOp)</div><div class="svc-campo-val">${_copiavel(n.indop)} <span class="svc-mini">${_esc(n.indop_desc || '')}</span>${indAlt}${regra}</div></div>
        <div class="svc-campo"><div class="svc-campo-lbl">CST IBS/CBS</div><div class="svc-campo-val">${_copiavel(n.cst_ibs_cbs)}</div></div>
        <div class="svc-campo"><div class="svc-campo-lbl">cClassTrib</div><div class="svc-campo-val">${_copiavel(n.cclasstrib)} <span class="svc-mini">${_esc(n.nome_cclasstrib || '')}</span>${alt}${semClass}</div></div>
        <div class="svc-campo"><div class="svc-campo-lbl">Local de incidência do IBS/CBS</div><div class="svc-campo-val">${_esc(n.local_ibs || '—')}</div></div>
        <div class="svc-campo"><div class="svc-campo-lbl">Simples Nacional (cAtivSN)</div><div class="svc-campo-val">${snHtml}</div></div>
      </div>

      ${avisos.length ? '<div class="svc-avisos">' + avisos.map(a => `<div class="svc-aviso">${a}</div>`).join('') + '</div>' : ''}

      <div class="svc-subtitle">NBS deste item (${d.nbs.length}) · clique para preencher os campos acima</div>
      <div class="table-wrapper">
        <table class="data-table svc-nbs-table">
          <thead><tr>${d.nbs.some(x => x.desdobro) ? '<th>cTribNac</th>' : ''}<th>NBS</th><th>Descrição</th><th>indOp</th><th>cClassTrib</th><th>Simples</th></tr></thead>
          <tbody>${d.nbs.map((x, i) => `
            <tr class="${i === svcNbsSel ? 'svc-sel' : ''}" onclick="selecionarNbs(${i})">
              ${d.nbs.some(y => y.desdobro) ? `<td>${_esc(x.desdobro || '')}</td>` : ''}
              <td class="mono">${_esc(x.nbs)}</td>
              <td>${_esc(x.descricao)}</td>
              <td class="mono">${_esc(x.indop || '')}${x.indop_alternativos.length ? ' / ' + x.indop_alternativos.map(a => _esc(a.codigo)).join(' / ') : ''}</td>
              <td class="mono" title="${_esc(x.nome_cclasstrib)}">${_esc(x.cclasstrib || 'em branco')}${x.cclasstrib_alternativos.length ? ' +' + x.cclasstrib_alternativos.length : ''}</td>
              <td>${x.ativ_sn_info ? _esc(x.ativ_sn_info.rotulo) : ''}</td>
            </tr>`).join('')}</tbody>
        </table>
      </div>

      ${d.cnaes.length ? `<div class="svc-subtitle">CNAEs ligados a este item</div><div class="svc-chips">` +
        d.cnaes.slice(0, 24).map(c => `<span class="svc-chip svc-chip-static" title="${_esc(c.descricao)}">${_esc(c.cnae)} · ${_esc(c.descricao.slice(0, 70))}</span>`).join('') +
        (d.cnaes.length > 24 ? `<span class="svc-chip svc-chip-static">+${d.cnaes.length - 24} CNAEs</span>` : '') + '</div>' : ''}

      <p class="svc-fonte">Fonte: Anexo VIII da NFS-e Nacional v${_esc(d.versao_anexo_viii)} (correlação item × NBS × indOp × cClassTrib × atividade do Simples).
        A ligação CNAE × LC 116 e a retenção de INSS vêm de tabela de referência de mercado: confira com a lei do seu município.</p>
    </div>`;
}

async function calcularFatorR() {
  const p = new URLSearchParams({
    folha12: document.getElementById('fr-folha').value,
    rbt12: document.getElementById('fr-rbt12').value,
    receita_mes: document.getElementById('fr-mes').value,
  });
  const r = await fetch('/api/servicos/fator-r?' + p.toString());
  const d = await r.json();
  const el = document.getElementById('fr-resultado');
  if (d.erro) { el.innerHTML = `<div class="error-msg">${_esc(d.erro)}</div>`; return; }
  const ok = d.anexo === 'III';
  const lin = (nome, x) => x.erro ? `<tr><td>${nome}</td><td colspan="3">${_esc(x.erro)}</td></tr>` :
    `<tr class="${(nome.endsWith(d.anexo)) ? 'svc-sel' : ''}"><td>${nome}</td><td>${_fmtPct(x.aliquota_nominal)}</td><td>${_fmtBRL(x.parcela_deduzir)}</td><td><strong>${_fmtPct(x.aliquota_efetiva)}</strong></td></tr>`;
  el.innerHTML = `
    <div class="fr-res ${ok ? 'fr-ok' : 'fr-alerta'}">
      <div class="fr-num">${_fmtPct(d.fator_r)}</div>
      <div>Fator R ${ok ? 'igual ou acima' : 'abaixo'} de 28%: tributação pelo <strong>Anexo ${d.anexo}</strong>.
        ${ok ? '' : `Para ir ao Anexo III, a folha dos 12 meses precisaria chegar a ${_fmtBRL(d.folha_minima_para_iii)} (faltam ${_fmtBRL(d.falta_folha_12m)}).`}
        ${d.das_estimado != null ? `<br>DAS estimado do mês: <strong>${_fmtBRL(d.das_estimado)}</strong> (no outro anexo seria ${_fmtBRL(d.das_outro_anexo)}).` : ''}</div>
    </div>
    <table class="data-table" style="margin-top:10px">
      <thead><tr><th>Anexo</th><th>Alíquota nominal</th><th>Parcela a deduzir</th><th>Alíquota efetiva</th></tr></thead>
      <tbody>${lin('Anexo III', d.anexo_iii)}${lin('Anexo V', d.anexo_v)}</tbody>
    </table>
    <p class="svc-fonte">Estimativa sem considerar segregação de receitas, ISS fixo ou retenções. A partir de 2027 (Res. CGSN 190/2026), nos dois primeiros meses de atividade o fator R é fixado em 0,28.</p>`;
}
