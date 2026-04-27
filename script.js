'use strict';

/* ── State ── */
const state = {
  history: JSON.parse(localStorage.getItem('pws_history') || '[]'),
  queryCount: parseInt(localStorage.getItem('pws_query_count') || '0', 10)
};

/* ── DOM refs ── */
const $ = id => document.getElementById(id);
const input       = $('pws-input');
const searchBtn   = $('search-btn');
const clearBtn    = $('clear-btn');
const resultPanel = $('result-panel');
const totalCount  = $('total-count');
const queryCount  = $('query-count');
const historyCard = $('history-card');
const historyList = $('history-list');

/* ── Init ── */
document.addEventListener('DOMContentLoaded', () => {
  totalCount.textContent = PWS_DATA.length;
  queryCount.textContent = state.queryCount;
  renderHistory();
  startClock();
  input.focus();
});

/* ── Clock ── */
function startClock() {
  const el = $('clock');
  function tick() {
    const now = new Date();
    el.textContent = now.toLocaleTimeString('zh-TW', { hour12: false });
  }
  tick();
  setInterval(tick, 1000);
}

/* ── Search ── */
function normalize(str) {
  return str.trim().toUpperCase().replace(/\s+/g, '');
}

function search(raw) {
  const query = normalize(raw);
  if (!query) return;

  state.queryCount++;
  localStorage.setItem('pws_query_count', state.queryCount);
  queryCount.textContent = state.queryCount;

  // 1. Exact match (case-insensitive)
  const exact = PWS_DATA.find(r => normalize(r.pws) === query);
  if (exact) {
    showResult('ok', exact, query);
    addHistory(raw.trim(), true);
    return;
  }

  // 2. Fuzzy / partial match
  const fuzzy = PWS_DATA.filter(r =>
    normalize(r.pws).includes(query) || query.includes(normalize(r.pws))
  );
  if (fuzzy.length === 1) {
    showResult('ok', fuzzy[0], query);
    addHistory(raw.trim(), true);
    return;
  }
  if (fuzzy.length > 1) {
    showFuzzyList(fuzzy, query);
    addHistory(raw.trim(), false);
    return;
  }

  // 3. No match
  showResult('error', null, query);
  addHistory(raw.trim(), false);
}

/* ── Render: single result ── */
function showResult(type, record, query) {
  resultPanel.innerHTML = '';
  resultPanel.style.display = 'block';

  if (type === 'ok') {
    resultPanel.innerHTML = `
      <div class="result-status ok animate-in">
        <span class="status-icon">✔</span>
        <span class="status-text">查詢成功 — 資料已找到</span>
        <span class="pws-echo">查詢：${escHtml(query)}</span>
      </div>
      <div class="result-body ok animate-in">
        <div class="result-grid">
          <div class="result-cell">
            <div class="cell-label">PWS P/N</div>
            <div class="cell-value">${escHtml(record.pws)}</div>
          </div>
          <div class="result-cell">
            <div class="cell-label">Status</div>
            <div class="cell-value">
              <span class="status-badge ok">✔ OK</span>
            </div>
          </div>
          <div class="result-cell">
            <div class="cell-label">PDB</div>
            <div class="cell-value ok-val">${escHtml(record.pdb)}</div>
          </div>
          <div class="result-cell">
            <div class="cell-label">Test Jig</div>
            <div class="cell-value ok-val">${escHtml(record.jig)}</div>
          </div>
        </div>
      </div>`;
  } else {
    resultPanel.innerHTML = `
      <div class="result-status error animate-in">
        <span class="status-icon">✘</span>
        <span class="status-text">查無資料</span>
        <span class="pws-echo">查詢：${escHtml(query)}</span>
      </div>
      <div class="result-body error animate-in">
        <div class="error-note">
          <strong>No Mapping Found</strong>
          找不到 PWS P/N「${escHtml(query)}」的對應資料，請確認料號是否正確。
        </div>
      </div>`;
  }

  resultPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  triggerPulse(input);
}

/* ── Render: fuzzy list ── */
function showFuzzyList(matches, query) {
  resultPanel.innerHTML = '';
  resultPanel.style.display = 'block';

  const items = matches.map(r => `
    <div class="fuzzy-item" data-pws="${escHtml(r.pws)}">
      <span class="fi-pws">${escHtml(r.pws)}</span>
      <span class="fi-meta">PDB: ${escHtml(r.pdb)} ／ JIG: ${escHtml(r.jig)}</span>
      <span class="fi-arrow">›</span>
    </div>`).join('');

  resultPanel.innerHTML = `
    <div class="result-status ok animate-in">
      <span class="status-icon">⚡</span>
      <span class="status-text">找到 ${matches.length} 筆相近結果，請點選確認</span>
      <span class="pws-echo">查詢：${escHtml(query)}</span>
    </div>
    <div class="result-body ok animate-in">
      <div class="fuzzy-list">
        <div class="fuzzy-list-title">模糊比對結果（點擊展開詳細）</div>
        ${items}
      </div>
    </div>`;

  resultPanel.querySelectorAll('.fuzzy-item').forEach(el => {
    el.addEventListener('click', () => {
      const pws = el.dataset.pws;
      input.value = pws;
      const record = PWS_DATA.find(r => r.pws === pws);
      if (record) showResult('ok', record, pws);
    });
  });

  resultPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

/* ── History ── */
function addHistory(pws, found) {
  const now = new Date().toLocaleTimeString('zh-TW', { hour12: false });
  state.history.unshift({ pws, found, time: now });
  if (state.history.length > 20) state.history.pop();
  localStorage.setItem('pws_history', JSON.stringify(state.history));
  renderHistory();
}

function renderHistory() {
  if (state.history.length === 0) {
    historyCard.style.display = 'none';
    return;
  }
  historyCard.style.display = 'block';
  historyList.innerHTML = state.history.map(h => `
    <div class="history-row" data-pws="${escHtml(h.pws)}">
      <span class="hr-pws">${escHtml(h.pws)}</span>
      <span class="hr-tag ${h.found ? 'ok' : 'error'}">${h.found ? 'OK' : 'MISS'}</span>
      <span class="hr-time">${h.time}</span>
    </div>`).join('');

  historyList.querySelectorAll('.history-row').forEach(el => {
    el.addEventListener('click', () => {
      input.value = el.dataset.pws;
      search(el.dataset.pws);
    });
  });
}

$('history-clear-btn').addEventListener('click', () => {
  state.history = [];
  localStorage.removeItem('pws_history');
  renderHistory();
});

/* ── Events ── */
searchBtn.addEventListener('click', () => search(input.value));

clearBtn.addEventListener('click', () => {
  input.value = '';
  resultPanel.style.display = 'none';
  resultPanel.innerHTML = '';
  input.focus();
});

input.addEventListener('keydown', e => {
  if (e.key === 'Enter') {
    e.preventDefault();
    search(input.value);
  }
});

/* ── Utility ── */
function escHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function triggerPulse(el) {
  el.classList.remove('pulse');
  void el.offsetWidth;
  el.classList.add('pulse');
}
