// 住民公開ページ（ツール4）側のロジック。読み取り専用・書き込みは一切行わない。
// STRUCTURE（様式構造）・RECORDS（各事業の生データ配列）・PUBLIC_CONFIG（集約ツールが
// 設定した表示・集計・検索項目や説明文）はビルド時ではなく、集約ツールの
// 「住民公開用ページを書き出す」操作時にこのファイルへ直接注入される。
// ☑・メモはlocalStorageにのみ保存し、どこにも送信しない（バックアップはJSON書き出しのみ）。

function $(sel) { return document.querySelector(sel); }

function el(tag, attrs, children) {
  const e = document.createElement(tag);
  if (attrs) for (const k in attrs) {
    if (k === 'text') e.textContent = attrs[k];
    else if (k === 'html') e.innerHTML = attrs[k];
    else e.setAttribute(k, attrs[k]);
  }
  (children || []).forEach(c => c && e.appendChild(c));
  return e;
}

// filler_app.js/aggregator_app.jsと同じ形（STRUCTURE.cellsからgrid Mapを復元する）。
function buildStateFromStructure(structure) {
  const grid = new Map();
  structure.cells.forEach(c => {
    const info = Object.assign({}, c);
    for (let r = c.row; r <= c.row2; r++) {
      for (let cc = c.col; cc <= c.col2; cc++) {
        grid.set(r + ',' + cc, info);
      }
    }
  });
  return {
    grid,
    maxRow: structure.maxRow,
    maxCol: structure.maxCol,
    widths: structure.widths,
    heights: structure.heights,
    sections: structure.sections,
    fileName: structure.formTitle || 'jigyo',
    manualGroups: structure.manualGroups || [],
  };
}

const BASE_STATE = null; // 使わない（各所で使い捨てstateをbuildStateFromStructureから作る）

// extractFieldValues（filler_app.js）と同じ「使い捨てグリッドに読み込ませてDOM経由で読む」方式。
function scratchExtract(data, cellIds) {
  if (!cellIds || cellIds.length === 0) return {};
  const scratch = $('#viewer-scratch-root');
  const tempState = buildStateFromStructure(STRUCTURE);
  GridRender.renderGrid(scratch, tempState, { showGear: false });
  GridRender.loadDataIntoGrid(tempState, data || {});
  const out = {};
  cellIds.forEach((id) => {
    const inputEl = document.getElementById(id);
    out[id] = inputEl ? inputEl.value : '';
  });
  scratch.innerHTML = '';
  return out;
}

let DETAIL_STATE = null;

const STORAGE_KEY = (PUBLIC_CONFIG && PUBLIC_CONFIG.storageKey) || 'jimujigyou_viewer';

function loadLocalState() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return { checked: [], memos: {} };
    const parsed = JSON.parse(raw);
    return { checked: parsed.checked || [], memos: parsed.memos || {} };
  } catch (e) {
    return { checked: [], memos: {} };
  }
}

function saveLocalState() {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify({
      checked: Array.from(V.checked),
      memos: Object.fromEntries(V.memos),
    }));
  } catch (e) { /* localStorageが使えない環境では、バックアップは諦めて閲覧自体は継続する */ }
}

const V = {
  records: [],        // [{id, data, displayValues, searchText}]
  checked: new Set(),
  memos: new Map(),
  colFilters: new Map(),
  keyword: '',
  onlyChecked: false,
  sortKey: null,
  sortDir: 'asc',
  visibleCols: new Set(),
  showMemoColumn: false, // ON時：一覧にメモ本文をそのまま表示し、印刷にも含める
};

// 表示項目（PUBLIC_CONFIG.displayFields）は{label, ids}の配列。idsは単独欄なら1件、
// 繰り返し列（Excelの合計セルと同じ範囲）なら行数分。列見出し・絞り込み・並び替え・
// V.visibleCols等で扱いやすいよう、各項目に短い内部キー（df0, df1, ...）を振り直して
// 保持する（idsをそのまま連結した文字列だとDOM要素idとして冗長・不安定になるため）。
const DISPLAY_FIELDS = (PUBLIC_CONFIG.displayFields || []).map((f, idx) => ({ key: 'df' + idx, label: f.label, ids: f.ids || [] }));

function displayIds() { return DISPLAY_FIELDS.map(f => f.key); }
function displayFieldByKey(key) { return DISPLAY_FIELDS.find(f => f.key === key); }
function labelForFieldKey(key) {
  const f = displayFieldByKey(key);
  return f ? f.label : key;
}
function allDisplayCellIds() {
  const out = [];
  DISPLAY_FIELDS.forEach(f => out.push(...f.ids));
  return out;
}
function searchIds() {
  const s = PUBLIC_CONFIG.searchFields || [];
  return s.length > 0 ? s : allDisplayCellIds();
}
// 集計対象の実セルid（単独欄は1件、繰り返し列（Excelの合計セルと同じ範囲）は行数分）を
// すべて平らにしたもの。aggregator_app.jsのbuildAggregateFieldConfigsが単独欄・繰り返し列を
// 同じ{label,unit,ids}形に正規化しているため、種別を意識せず扱える。
function aggregateIds() {
  const out = [];
  (PUBLIC_CONFIG.aggregateFields || []).forEach(f => out.push(...(f.ids || [])));
  return out;
}

// 1項目分の表示値を作る：idsが1件（単独欄）ならその値をそのまま、複数件（繰り返し列）なら
// 数値として事業内で合計する（②集計対象と同じ考え方。繰り返し列は数値・金額型のみが
// 候補になっているため、複数件は必ず合計できる前提でよい）。
function computeFieldDisplayValue(rawValues, ids) {
  if (!ids || ids.length === 0) return '';
  if (ids.length === 1) return rawValues[ids[0]] || '';
  let sum = 0;
  let any = false;
  ids.forEach((id) => {
    const n = parseFloat(String(rawValues[id] || '').replace(/[,\s]/g, ''));
    if (!isNaN(n)) { sum += n; any = true; }
  });
  return any ? sum.toLocaleString('ja-JP') : '';
}

// ☑・メモの永続化キー。STRUCTURE.fileNameFields（ビルダーSTEP4で「ファイル名に含める項目」
// として選ばれた、事業名等の識別用セル＝filler_app.jsの書き出しファイル名と同じ元）の値を
// つないだ文字列を使う。配列インデックスと違い、事業の並び順や件数が次回の再集約で変わっても
// 「同じ事業」であれば同じキーになるため、☑・メモが別の事業に付け替わる事故を防げる。
// fileNameFieldsが未設定／該当欄が全て空欄の事業のみ、従来通りインデックスにフォールバックする。
function computePersistKey(rawValues, idx) {
  const ids = STRUCTURE.fileNameFields || [];
  const parts = ids.map(id => String(rawValues[id] || '').trim()).filter(v => v !== '');
  if (parts.length > 0) return parts.join('｜');
  return 'idx:' + idx;
}

function initRecords() {
  const allCellIds = Array.from(new Set([
    ...allDisplayCellIds(),
    ...searchIds(),
    ...aggregateIds(),
    ...(STRUCTURE.fileNameFields || []),
  ]));
  const seenKeyCounts = new Map();
  V.records = (RECORDS || []).map((data, idx) => {
    const rawValues = scratchExtract(data, allCellIds);
    const displayValues = {};
    DISPLAY_FIELDS.forEach((f) => { displayValues[f.key] = computeFieldDisplayValue(rawValues, f.ids); });
    const searchText = searchIds().map(id => rawValues[id] || '').join(' ').toLowerCase();
    // 同一キーが複数事業に発生した場合（識別用の項目が同じ値になる稀なケース）に、
    // ☑・メモが別事業と衝突しないよう連番を付けて一意化する。
    let persistKey = computePersistKey(rawValues, idx);
    const seenCount = seenKeyCounts.get(persistKey) || 0;
    if (seenCount > 0) persistKey = persistKey + '#' + (seenCount + 1);
    seenKeyCounts.set(persistKey, seenCount + 1);
    // rawValues（実セルidキー）はrenderSummaryの集計対象（PUBLIC_CONFIG.aggregateFields[].ids、
    // 実セルidの配列）を合計する際に使う。displayValues（表示項目のdfキー）とはキー体系が
    // 異なるため両方を保持する。idは画面内ナビゲーション（詳細画面の前へ／次へ、URLの
    // #detail-N)専用の一時的な連番、persistKeyがlocalStorage保存用の恒久キー。
    return { id: idx, persistKey, data, rawValues, displayValues, searchText };
  });
}

function toggleChecked(id) {
  if (V.checked.has(id)) V.checked.delete(id); else V.checked.add(id);
  saveLocalState();
}

function setMemo(id, text) {
  if (text) V.memos.set(id, text); else V.memos.delete(id);
  saveLocalState();
}

function currentDisplayIds() {
  return displayIds().filter(id => V.visibleCols.has(id));
}

function columnValue(record, id) { return record.displayValues[id] || ''; }

function uniqueColumnValues(id) {
  const set = new Set();
  V.records.forEach(r => set.add(columnValue(r, id)));
  return Array.from(set).sort((a, b) => String(a).localeCompare(String(b), 'ja'));
}

function matchesFilters(record) {
  if (V.onlyChecked && !V.checked.has(record.persistKey)) return false;
  if (V.keyword && record.searchText.indexOf(V.keyword.toLowerCase()) === -1) return false;
  for (const id of currentDisplayIds()) {
    const allowed = V.colFilters.get(id);
    if (!allowed) continue;
    if (!allowed.has(columnValue(record, id))) return false;
  }
  return true;
}

function filteredSortedRecords() {
  let list = V.records.filter(matchesFilters);
  if (V.sortKey) {
    list = list.slice().sort((a, b) => {
      const av = columnValue(a, V.sortKey);
      const bv = columnValue(b, V.sortKey);
      const an = parseFloat(String(av).replace(/[,\s]/g, ''));
      const bn = parseFloat(String(bv).replace(/[,\s]/g, ''));
      let cmp;
      if (!isNaN(an) && !isNaN(bn)) cmp = an - bn;
      else cmp = String(av).localeCompare(String(bv), 'ja');
      return V.sortDir === 'asc' ? cmp : -cmp;
    });
  }
  return list;
}

function hasActiveFilters() { return V.colFilters.size > 0; }

// 集計欄：文字を横並びに連結するだけだと目立たないため、各項目を「ラベル＋大きめの数値」の
// チップとして描画する（CSS側は#viewer-summary/.summary-item/.summary-value参照）。
function renderSummary(list) {
  const root = $('#viewer-summary');
  root.innerHTML = '';
  const items = [];
  if (PUBLIC_CONFIG.showRecordCount !== false) {
    const label = hasActiveFilters() || V.keyword || V.onlyChecked ? `件数（全${V.records.length}件中）` : '件数';
    items.push({ label, value: `${list.length}件` });
  }
  // f.ids：単独欄なら1件、繰り返し列（Excelの合計セルと同じ範囲）なら行数分。事業内で
  // まず合計してから、事業間でさらに合計する（＝行数に関わらずidsを全部足せば同じ結果）。
  (PUBLIC_CONFIG.aggregateFields || []).forEach((f) => {
    let sum = 0;
    list.forEach((r) => {
      (f.ids || []).forEach((id) => {
        const n = parseFloat(String(r.rawValues[id] || '').replace(/[,\s]/g, ''));
        if (!isNaN(n)) sum += n;
      });
    });
    items.push({ label: `${f.label}合計`, value: `${sum.toLocaleString('ja-JP')}${f.unit || ''}` });
  });
  items.forEach((item) => {
    root.appendChild(el('span', { class: 'summary-item' }, [
      el('span', { class: 'summary-label', text: item.label }),
      el('span', { class: 'summary-value', text: item.value }),
    ]));
  });
}

// filler_app.js（buildColFilterDetails）と同じ考え方：<details>をExcel風▼として使い、
// パネルはdocument.body直下にposition:fixedで配置する（#viewer-table-rootのoverflow-x:auto
// によるクリップを避けるため）。
let filterPanelNodes = [];
function removeFilterPanelNodes() {
  filterPanelNodes.forEach(p => p.remove());
  filterPanelNodes = [];
}

function buildColFilterDetails(id) {
  const values = uniqueColumnValues(id);
  const details = el('details', { class: 'col-filter' });
  const summary = el('summary', { text: '▼', title: '値で絞り込む' });
  details.appendChild(summary);
  const panel = el('div', { class: 'col-filter-panel' });
  panel.style.display = 'none';
  values.forEach((v, idx) => {
    const inputId = 'colfilter_' + id + '_' + idx;
    const cb = el('input', { type: 'checkbox', id: inputId });
    const currentAllowed = V.colFilters.get(id);
    if (!currentAllowed || currentAllowed.has(v)) cb.setAttribute('checked', 'checked');
    cb.addEventListener('change', () => {
      const allowed = new Set(V.colFilters.get(id) || values);
      if (cb.checked) allowed.add(v); else allowed.delete(v);
      V.colFilters.set(id, allowed);
      renderList();
    });
    panel.appendChild(el('label', {}, [cb, document.createTextNode(v === '' ? '（空欄）' : v)]));
  });
  document.body.appendChild(panel);
  filterPanelNodes.push(panel);
  details.addEventListener('toggle', () => {
    if (details.open) {
      const rect = summary.getBoundingClientRect();
      panel.style.position = 'fixed';
      panel.style.top = rect.bottom + 'px';
      panel.style.left = rect.left + 'px';
      panel.style.display = 'block';
    } else {
      panel.style.display = 'none';
    }
  });
  return details;
}

function buildMemoCell(record) {
  const details = el('details', { class: 'memo-popover' });
  const summary = el('summary', { text: V.memos.get(record.persistKey) ? '📝あり' : '📝', title: 'メモを編集（この端末にのみ保存）' });
  details.appendChild(summary);
  const panel = el('div', { class: 'memo-panel' });
  panel.style.display = 'none';
  const textarea = el('textarea', {});
  textarea.value = V.memos.get(record.persistKey) || '';
  textarea.addEventListener('input', () => {
    setMemo(record.persistKey, textarea.value);
    summary.textContent = textarea.value ? '📝あり' : '📝';
  });
  panel.appendChild(textarea);
  document.body.appendChild(panel);
  filterPanelNodes.push(panel);
  details.addEventListener('toggle', () => {
    if (details.open) {
      const rect = summary.getBoundingClientRect();
      panel.style.position = 'fixed';
      panel.style.top = rect.bottom + 'px';
      panel.style.left = rect.left + 'px';
      panel.style.display = 'block';
    } else {
      panel.style.display = 'none';
    }
  });
  return details;
}

function renderColToggle() {
  const root = $('#viewer-col-toggle');
  root.innerHTML = '';
  if (displayIds().length === 0) return;
  root.appendChild(el('span', { text: '表示する項目：' }));
  displayIds().forEach((id) => {
    const inputId = 'colvis_' + id;
    const cb = el('input', { type: 'checkbox', id: inputId });
    if (V.visibleCols.has(id)) cb.setAttribute('checked', 'checked');
    cb.addEventListener('change', () => {
      if (cb.checked) V.visibleCols.add(id); else V.visibleCols.delete(id);
      renderList();
    });
    root.appendChild(el('label', { for: inputId }, [cb, el('span', { text: labelForFieldKey(id) })]));
  });
}

function renderList() {
  const list = filteredSortedRecords();
  renderSummary(list);

  const tableRoot = $('#viewer-table-root');
  tableRoot.innerHTML = '';
  removeFilterPanelNodes();

  const ids = currentDisplayIds();
  const table = el('table', { class: 'viewer-table' });
  const thead = el('thead');
  const headRow = el('tr');
  ids.forEach((id) => {
    const th = el('th');
    const arrow = V.sortKey === id ? (V.sortDir === 'asc' ? ' ▲' : ' ▼') : '';
    const sortBtn = el('button', { class: 'sort-btn', type: 'button', text: labelForFieldKey(id) + arrow });
    sortBtn.addEventListener('click', () => {
      if (V.sortKey === id) V.sortDir = V.sortDir === 'asc' ? 'desc' : 'asc';
      else { V.sortKey = id; V.sortDir = 'asc'; }
      renderList();
    });
    th.appendChild(sortBtn);
    th.appendChild(buildColFilterDetails(id));
    headRow.appendChild(th);
  });
  headRow.appendChild(el('th', { text: '気になる／メモ' }));
  // 「メモを表示」ONのときだけ、メモ本文をそのまま読める列を追加する（画面表示と印刷の
  // 両方に反映される＝「表示中の事業を印刷」が見えている通りの内容になる）。
  if (V.showMemoColumn) headRow.appendChild(el('th', { text: 'メモ内容' }));
  headRow.appendChild(el('th', { text: '' }));
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = el('tbody');
  list.forEach((record) => {
    const tr = el('tr');
    ids.forEach(id => tr.appendChild(el('td', { text: record.displayValues[id] || '' })));

    const chkTd = el('td', { class: 'viewer-chk-cell' });
    const chk = el('input', { type: 'checkbox' });
    if (V.checked.has(record.persistKey)) chk.setAttribute('checked', 'checked');
    chk.addEventListener('change', () => toggleChecked(record.persistKey));
    chkTd.appendChild(chk);
    chkTd.appendChild(buildMemoCell(record));
    tr.appendChild(chkTd);

    if (V.showMemoColumn) {
      tr.appendChild(el('td', { class: 'viewer-memo-text-cell', text: V.memos.get(record.persistKey) || '' }));
    }

    const btnTd = el('td');
    const detailBtn = el('button', { type: 'button', text: '詳細' });
    detailBtn.addEventListener('click', () => openDetail(record.id));
    btnTd.appendChild(detailBtn);
    tr.appendChild(btnTd);

    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  tableRoot.appendChild(table);
}

// ---------------------------------------------------------------------------
// 詳細画面
// ---------------------------------------------------------------------------
function showScreen(name) {
  $('#list-root').style.display = name === 'list' ? '' : 'none';
  $('#detail-root').style.display = name === 'detail' ? '' : 'none';
}

function findRecord(id) { return V.records.find(r => r.id === id); }
function findRecordByPersistKey(key) { return V.records.find(r => r.persistKey === key); }

function renderDetail(id) {
  const record = findRecord(id);
  if (!record) return;

  const bar = $('#detail-summary-bar');
  bar.innerHTML = '';
  displayIds().forEach((fid) => {
    bar.appendChild(el('div', { class: 'detail-field' }, [
      el('span', { class: 'k', text: labelForFieldKey(fid) }),
      el('span', { class: 'v', text: record.displayValues[fid] || '' }),
    ]));
  });
  const chk = el('input', { type: 'checkbox', id: 'detail-check' });
  if (V.checked.has(record.persistKey)) chk.setAttribute('checked', 'checked');
  chk.addEventListener('change', () => toggleChecked(record.persistKey));
  bar.appendChild(el('label', { class: 'detail-check-label', for: 'detail-check' }, [chk, document.createTextNode('気になる')]));

  const memoBox = el('textarea', { id: 'detail-memo', placeholder: 'メモ（この端末にのみ保存されます）' });
  memoBox.value = V.memos.get(record.persistKey) || '';
  memoBox.addEventListener('input', () => setMemo(record.persistKey, memoBox.value));
  bar.appendChild(memoBox);

  const root = $('#detail-body-root');
  root.innerHTML = '';
  DETAIL_STATE = buildStateFromStructure(STRUCTURE);
  GridRender.renderGrid(root, DETAIL_STATE, { showGear: false, readonly: true });
  GridRender.loadDataIntoGrid(DETAIL_STATE, record.data || {});

  const list = filteredSortedRecords();
  const idx = list.findIndex(r => r.id === id);
  $('#detail-prev').disabled = idx <= 0;
  $('#detail-next').disabled = idx < 0 || idx >= list.length - 1;
  $('#detail-prev').onclick = () => { if (idx > 0) openDetail(list[idx - 1].id); };
  $('#detail-next').onclick = () => { if (idx >= 0 && idx < list.length - 1) openDetail(list[idx + 1].id); };
}

// URLのハッシュ部分だけを書き換える（一覧⇔詳細の画面遷移自体はこの関数を経由せず内部状態
// だけで行うため、失敗しても機能には影響しない。あくまでブックマーク・直リンク用のおまけ）。
// 実機確認で判明：file://で開いた場合、location.hash=・history.replaceStateのどちらで
// URLを書き換えても、環境によっては「Unsafe attempt to load URL...file: URLs are treated
// as unique security origins」としてブロックされ、コンソールにエラーが出続けることがある
// （http(s)://で配信した場合は問題なし、実機のPlaywrightで確認済み）。file://はローカルで
// 直接開く用途がほとんどでURL共有の意味も薄いため、file://のときはURL書き換え自体を
// 最初から試みない（try/catchで握りつぶすだけでは、ブロックされたこと自体がコンソールに
// エラーとして残り続けるため、根本的にスキップするほうが確実）。
const HASH_NAV_SUPPORTED = window.location.protocol !== 'file:';

function setUrlHash(hash) {
  if (!HASH_NAV_SUPPORTED) return;
  if (window.location.hash === hash) return;
  try {
    if (window.history && window.history.replaceState) {
      const base = window.location.pathname + window.location.search;
      window.history.replaceState(null, '', hash ? base + hash : base);
    } else {
      window.location.hash = hash;
    }
  } catch (e) { /* 万一ブロックされても表示は続行する */ }
}

function openDetail(id) {
  if (!findRecord(id)) return;
  setUrlHash('#detail-' + id);
  renderDetail(id);
  showScreen('detail');
}

function backToList() {
  setUrlHash('');
  showScreen('list');
  renderList();
}

function handleHashChange() {
  const m = window.location.hash.match(/^#detail-(\d+)$/);
  if (m) {
    const id = parseInt(m[1], 10);
    if (findRecord(id)) { renderDetail(id); showScreen('detail'); return; }
  }
  showScreen('list');
}

// ---------------------------------------------------------------------------
// 印刷（現在表示中＝一覧に見えている表をそのまま1枚の表として印刷）／単発印刷（詳細画面）
// 「☑を付けた事業だけ印刷したい」場合は、①「気になる事業だけ表示」で絞り込んでから
// ②この共通の印刷ボタンを押す、という2段階の操作にする（専用の別ボタンは持たない）。
// メモの有無は「メモを表示」トグル（V.showMemoColumn）に従う＝画面の見た目と一致させる。
// ---------------------------------------------------------------------------
function printFiltered() {
  const list = filteredSortedRecords();
  const root = $('#viewer-print-root');
  root.innerHTML = '';
  if (list.length === 0) return;

  const ids = currentDisplayIds();
  const table = el('table', { class: 'viewer-print-list-table' });
  const thead = el('thead');
  const headRow = el('tr');
  ids.forEach(id => headRow.appendChild(el('th', { text: labelForFieldKey(id) })));
  headRow.appendChild(el('th', { text: '気になる' }));
  if (V.showMemoColumn) headRow.appendChild(el('th', { text: 'メモ内容' }));
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = el('tbody');
  list.forEach((r) => {
    const tr = el('tr');
    ids.forEach(id => tr.appendChild(el('td', { text: r.displayValues[id] || '' })));
    tr.appendChild(el('td', { class: 'viewer-print-chk', text: V.checked.has(r.persistKey) ? '☑' : '' }));
    if (V.showMemoColumn) tr.appendChild(el('td', { text: V.memos.get(r.persistKey) || '' }));
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  const section = el('div', { class: 'bulk-print-record' });
  if (PUBLIC_CONFIG.title) section.appendChild(el('h3', { text: PUBLIC_CONFIG.title }));
  section.appendChild(table);
  root.appendChild(section);

  document.body.classList.add('bulk-printing');
  window.print();
}

function printSingleSnapshot(state, heading) {
  const table = GridRender.buildPrintTable(state);
  const root = $('#viewer-print-root');
  root.innerHTML = '';
  const section = el('div', { class: 'bulk-print-record' });
  if (heading) section.appendChild(el('h2', { text: heading }));
  section.appendChild(table);
  root.appendChild(section);
  document.body.classList.add('bulk-printing');
  window.print();
}

function cleanupBulkPrint() {
  document.body.classList.remove('bulk-printing');
  const root = $('#viewer-print-root');
  if (root) root.innerHTML = '';
}

// ---------------------------------------------------------------------------
// ☑・メモのバックアップ（JSON書き出し／読み込み）。個人用バックアップが目的で、
// 自治体への送信・集計は行わない（サーバー非経由・localStorageのみ）。
// ---------------------------------------------------------------------------
function downloadBlobSimple(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

function exportBackup() {
  const payload = { checked: Array.from(V.checked), memos: Object.fromEntries(V.memos) };
  const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const filename = `${PUBLIC_CONFIG.title || 'viewer'}_チェックとメモ_${dateStr}.json`;
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  downloadBlobSimple(blob, filename);
}

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsText(file, 'utf-8');
  });
}

async function importBackup(file) {
  try {
    const text = await readFileAsText(file);
    const parsed = JSON.parse(text);
    V.checked = new Set((parsed.checked || []).filter(key => findRecordByPersistKey(key)));
    V.memos = new Map(Object.entries(parsed.memos || {}).filter(([key]) => findRecordByPersistKey(key)));
    saveLocalState();
    renderList();
  } catch (e) { /* 壊れたJSONは無視 */ }
}

function init() {
  initRecords();
  V.visibleCols = new Set(displayIds());
  const local = loadLocalState();
  V.checked = new Set(local.checked);
  V.memos = new Map(Object.entries(local.memos));

  $('#viewer-title').textContent = PUBLIC_CONFIG.title || '';
  $('#viewer-description').textContent = PUBLIC_CONFIG.description || '';
  $('#viewer-source').textContent = PUBLIC_CONFIG.source || '';
  $('#viewer-updated').textContent = PUBLIC_CONFIG.updatedDate ? `　データ最終更新日：${PUBLIC_CONFIG.updatedDate}` : '';

  $('#viewer-keyword').addEventListener('input', (ev) => { V.keyword = ev.target.value; renderList(); });
  $('#viewer-only-checked').addEventListener('change', (ev) => { V.onlyChecked = ev.target.checked; renderList(); });
  $('#viewer-show-memo').addEventListener('change', (ev) => { V.showMemoColumn = ev.target.checked; renderList(); });
  $('#btn-print-checked').addEventListener('click', printFiltered);
  $('#btn-export-backup').addEventListener('click', exportBackup);
  $('#backup-file-input').addEventListener('change', (ev) => {
    if (ev.target.files[0]) importBackup(ev.target.files[0]);
    ev.target.value = '';
  });
  $('#btn-back-to-list').addEventListener('click', backToList);
  $('#btn-print-detail').addEventListener('click', () => printSingleSnapshot(DETAIL_STATE, PUBLIC_CONFIG.title));
  window.addEventListener('afterprint', cleanupBulkPrint);
  window.addEventListener('hashchange', handleHashChange);

  renderColToggle();
  renderList();
  handleHashChange();
}

if (typeof window !== 'undefined') {
  window.addEventListener('DOMContentLoaded', init);
  window.__app = {
    get V() { return V; },
    initRecords, renderList, renderColToggle, renderSummary,
    openDetail, backToList, renderDetail, handleHashChange,
    toggleChecked, setMemo, filteredSortedRecords,
    printFiltered, printSingleSnapshot, cleanupBulkPrint,
    exportBackup, importBackup, loadLocalState, saveLocalState,
    labelForFieldKey, displayIds, allDisplayCellIds,
    findRecord, findRecordByPersistKey, computePersistKey,
    get DETAIL_STATE() { return DETAIL_STATE; },
    init,
  };
}
