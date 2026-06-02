/* ═══════════════════════════════════════════════════════
   ONE PAGE REPORT v2 — app.js
   ═══════════════════════════════════════════════════════ */

const CELL = {
  titulo:      'B2',
  data:        'M2',
  responsavel: 'M3',
  setor:       'M4',
  faturamento: 'B9',
  nc:          'E9',
  cpi:         'H9',
  spi:         'K9',
  finPct:      'B14',
  procPct:     'E14',
  niPct:       'H14',
  retPct:      'K14',
  finQtd:      'B19',
  procQtd:     'E19',
  niQtd:       'H19',
  retQtd:      'K19',
  proposta:    'B24',
  custo:       'E24',
  eficiencia:  'H24',
  prazo:       'K24',
  desenhoDia:  'B30',
  eficCx:      'E30',
  retCx:       'H30',
  statusEfic:  'E32',
  statusRet:   'H32',
  riscos:      ['B37','B38','B39','B40','B41'],
  acoes:       ['E37','E38','E39','E40','E41'],
  conquistas:  ['H37','H38','H39','H40','H41'],
  reclamacoes: ['K37','K38','K39','K40','K41'],
};

const uploadScreen = document.getElementById('upload-screen');
const reportScreen = document.getElementById('report-screen');
const dropZone     = document.getElementById('drop-zone');
const fileInput    = document.getElementById('file-input');
const errMsg       = document.getElementById('upload-error');

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag'));
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.classList.remove('drag');
  if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', e => { if (e.target.files[0]) handleFile(e.target.files[0]); });

function handleFile(file) {
  errMsg.classList.add('hidden');
  // Aceita por nome — não depende do MIME que o Windows às vezes não detecta
  const nomeOk = /\.(xlsx|xls|xlsm|xlsb|xltx|xltm)$/i.test(file.name);
  const tipoOk = file.type === '' ||
                 file.type.includes('excel') ||
                 file.type.includes('spreadsheet') ||
                 file.type.includes('officedocument') ||
                 file.type.includes('ms-excel');
  if (!nomeOk && !tipoOk) {
    errMsg.classList.remove('hidden');
    return;
  }
  const reader = new FileReader();
  reader.onload = function(ev) {
    try {
      const wb = XLSX.read(ev.target.result, { type: 'array', cellDates: true });
      renderReport(wb.Sheets[wb.SheetNames[0]]);
    } catch(err) {
      console.error('Erro ao ler planilha:', err);
      errMsg.classList.remove('hidden');
    }
  };
  reader.readAsArrayBuffer(file);
}

function c(ws, addr) {
  const cell = ws[addr];
  if (!cell) return null;
  if (cell.t === 'd') return cell.v;
  return cell.v ?? null;
}

function renderReport(ws) {

  // Cabeçalho
  set('h-projeto',     strip(c(ws, CELL.titulo), 'ONE PAGE REPORT -') || 'ONE PAGE REPORT');
  set('h-responsavel', strip(c(ws, CELL.responsavel), 'Responsável:'));
  set('h-setor',       strip(c(ws, CELL.setor), 'Setor:'));
  iniciarRelogio();

  // SPI / CPI
  const spiTxt = String(c(ws, CELL.spi) ?? '—');
  const cpiTxt = String(c(ws, CELL.cpi) ?? '—');
  set('a-spi-big', spiTxt);
  set('a-cpi-big', cpiTxt);

  const heroEl = document.getElementById('a-spi-big');
  if (heroEl) {
    if (/atraso|acima/i.test(spiTxt)) heroEl.style.color = '#ff6b5b';
    else if (/prazo|ok|dentro/i.test(spiTxt)) heroEl.style.color = 'var(--green-lt)';
  }

  // Painel de índices coluna direita
  const eficCx = toNum(c(ws, CELL.eficCx));
  const retCx  = toNum(c(ws, CELL.retCx));

  set('cr-cpi',   cpiTxt);
  set('cr-spi',   spiTxt);
  set('cr-efcx',  fmtPct(eficCx));
  set('cr-retcx', fmtPct(retCx));
  set('cr-nc',    fmtInt(c(ws, CELL.nc)));
  set('cr-ddia',  fmtDecimal(c(ws, CELL.desenhoDia)));

  const crCpi = document.getElementById('cr-cpi');
  const crSpi = document.getElementById('cr-spi');
  if (crCpi && /acima|atraso/i.test(cpiTxt)) crCpi.style.color = '#ff6b5b';
  if (crSpi && /atraso/i.test(spiTxt))       crSpi.style.color = '#ff6b5b';

  // Métricas aside
  const fat = c(ws, CELL.faturamento);
  if (fat !== null && fat !== '') {
    set('a-fat', fmtMoeda(toNum(fat)));
    set('a-fat-note', '');
  } else {
    set('a-fat', '—');
    set('a-fat-note', 'Aguardando dado');
  }

  set('a-nc',    fmtInt(c(ws, CELL.nc)));
  set('a-prazo', c(ws, CELL.prazo) ?? '—');
  set('a-ddia',  fmtDecimal(c(ws, CELL.desenhoDia)));

  // Eficiências aside
  const efic = toNum(c(ws, CELL.eficiencia));
  set('a-efic',    fmtPct(efic));
  set('a-efic-cx', fmtPct(eficCx));
  set('a-ret-cx',  fmtPct(retCx));

  animWidth('bar-efic',    efic   * 100);
  animWidth('bar-efic-cx', eficCx * 100);
  animWidth('bar-ret-cx',  retCx  * 100);

  // Financeiro
  const proposta = toNum(c(ws, CELL.proposta));
  const custo    = toNum(c(ws, CELL.custo));
  const saldo    = proposta - custo;

  set('c-proposta', fmtMoeda(proposta));
  set('c-custo',    fmtMoeda(custo));

  const saldoEl = document.getElementById('c-saldo');
  if (saldoEl) {
    saldoEl.textContent = fmtMoeda(saldo);
    saldoEl.style.color = saldo >= 0 ? 'var(--green)' : 'var(--orange)';
  }

  // Produção
  const finPct  = toNum(c(ws, CELL.finPct));
  const procPct = toNum(c(ws, CELL.procPct));
  const niPct   = toNum(c(ws, CELL.niPct));
  const retPct  = toNum(c(ws, CELL.retPct));

  const finQtd  = toNum(c(ws, CELL.finQtd));
  const procQtd = toNum(c(ws, CELL.procQtd));
  const niQtd   = toNum(c(ws, CELL.niQtd));
  const retQtd  = toNum(c(ws, CELL.retQtd));

  set('pb-total',   `${finQtd + procQtd + niQtd + retQtd} itens no total`);
  set('p-fin-pct',  fmtPct(finPct));
  set('p-proc-pct', fmtPct(procPct));
  set('p-ni-pct',   fmtPct(niPct));
  set('p-ret-pct',  fmtPct(retPct));
  set('p-fin-qtd',  `${finQtd} itens`);
  set('p-proc-qtd', `${procQtd} itens`);
  set('p-ni-qtd',   `${niQtd} itens`);
  set('p-ret-qtd',  `${retQtd} itens`);

  animHeight('tower-fin',  finPct  * 100);
  animHeight('tower-proc', procPct * 100);
  animHeight('tower-ni',   niPct   * 100);
  animHeight('tower-ret',  retPct  * 100);

  set('pbp-pct', fmtPct(finPct));
  animWidth('pbp-fill-fin',  finPct  * 100);
  animWidth('pbp-fill-proc', procPct * 100);
  animWidth('pbp-fill-ni',   niPct   * 100);
  animWidth('pbp-fill-ret',  retPct  * 100);

  // Notas
  renderLista('lista-riscos',      CELL.riscos,      ws);
  renderLista('lista-acoes',       CELL.acoes,       ws);
  renderLista('lista-conquistas',  CELL.conquistas,  ws);
  renderLista('lista-reclamacoes', CELL.reclamacoes, ws);

  set('v-gerado', new Date().toLocaleString('pt-BR'));

  uploadScreen.classList.add('hidden');
  reportScreen.classList.remove('hidden');
}

function renderLista(id, addrs, ws) {
  const ul = document.getElementById(id);
  if (!ul) return;
  ul.innerHTML = '';
  addrs.forEach((addr, i) => {
    const val   = c(ws, addr);
    const isNum = val !== null && !isNaN(val) && Number(val) === (i + 1);
    const li    = document.createElement('li');
    if (val && !isNum && String(val).trim() !== '') {
      li.textContent = String(val).trim();
      li.classList.add('filled');
    } else {
      li.textContent = 'Não informado';
      li.classList.add('empty');
    }
    ul.appendChild(li);
  });
}

function set(id, val) {
  const el = document.getElementById(id);
  if (el) el.textContent = (val !== null && val !== undefined && String(val).trim() !== '') ? val : '—';
}

function animWidth(id, pct) {
  setTimeout(() => {
    const el = document.getElementById(id);
    if (el) el.style.width = clamp(pct) + '%';
  }, 80);
}

function animHeight(id, pct) {
  setTimeout(() => {
    const el = document.getElementById(id);
    if (el) el.style.height = clamp(pct) + '%';
  }, 80);
}

function clamp(v) { return Math.min(Math.max(v || 0, 0), 100); }

function toNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  return parseFloat(String(v).replace(/[^\d.,\-]/g, '').replace(',', '.')) || 0;
}

function fmtMoeda(n) {
  return Number(n).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function fmtPct(n) {
  const v = n <= 1 ? n * 100 : n;
  return v.toFixed(1) + '%';
}

function fmtInt(v) {
  if (v === null || v === undefined || v === '') return '—';
  return String(Math.round(Number(v)));
}

function fmtDecimal(v) {
  if (v === null || v === undefined || v === '') return '—';
  return Number(v).toFixed(2);
}

function strip(str, prefix) {
  if (!str) return '—';
  return String(str).replace(new RegExp('^' + prefix + '\\s*', 'i'), '').trim() || '—';
}

let clockInterval = null;
function iniciarRelogio() {
  if (clockInterval) clearInterval(clockInterval);
  function tick() {
    const now  = new Date();
    const hora = now.toLocaleTimeString('pt-BR');
    const data = now.toLocaleDateString('pt-BR', { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' });
    const clk  = document.getElementById('ts-clock');
    const dt   = document.getElementById('ts-clock-date');
    if (clk) clk.textContent = hora;
    if (dt)  dt.textContent  = data;
  }
  tick();
  clockInterval = setInterval(tick, 1000);
}

function resetApp() {
  fileInput.value = '';
  errMsg.classList.add('hidden');
  reportScreen.classList.add('hidden');
  uploadScreen.classList.remove('hidden');
  if (clockInterval) { clearInterval(clockInterval); clockInterval = null; }
}

