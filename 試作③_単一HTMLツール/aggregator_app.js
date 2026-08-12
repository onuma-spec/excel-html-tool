// 管理者用集約ツール（ツール3）側のロジック。
// ビルダー（ツール1）が埋め込んだ STRUCTURE（値ではなく確定済みの構造）と、
// VIEWER_TEMPLATE（住民公開ページの雛形テキスト、ビルド時にassemble.pyが埋め込む）を持つ。
// 各部署が書き出したJSON（ツール2の出力）を複数読み込み、住民公開設定を行った上で、
// STRUCTURE・読み込んだデータ・設定をVIEWER_TEMPLATEへ注入した単一HTMLを書き出す。

let CURRENT_STATE = null; // STRUCTUREから復元したgrid/sections/maxCol等（filler_app.jsのbuildStateFromStructureと同じ形）

// ツール4（住民公開ページ）のHTML全文。ビルド時（assemble.py）に埋め込まれる。
// 「__STRUCTURE__」「__RECORDS__」「__PUBLIC_CONFIG__」「__VIEWER_TITLE__」の4箇所だけが
// 未確定のまま残っており、doExportAsViewer()がそれぞれ置き換えてから書き出す。
const VIEWER_TEMPLATE = "__VIEWER_TEMPLATE_JSON__";

// 書き出し先フォルダ（Chromium限定・File System Access API）。builder_app.js/filler_app.jsと
// 同じ仕組みだが、このツール専用のハンドルとして別変数に保持する。
let EXPORT_DIR_HANDLE = null;

function $(sel) { return document.querySelector(sel); }

function setStatus(text) {
  const elx = $('#status');
  elx.textContent = text;
  if (!text) return;
  elx.classList.remove('status-flash');
  void elx.offsetWidth;
  elx.classList.add('status-flash');
  if (typeof elx.scrollIntoView === 'function') {
    elx.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
}

function sanitizeForFileName(s) {
  return String(s == null ? '' : s).trim().replace(/[\\/:*?"<>|\r\n\t]/g, '').slice(0, 60);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

async function saveToExportDir(blob, filename) {
  try {
    let exists = false;
    try { await EXPORT_DIR_HANDLE.getFileHandle(filename); exists = true; } catch (e) { /* ファイルが無い＝新規保存 */ }
    if (exists && !window.confirm(`「${filename}」は指定フォルダに既に存在します。上書きしますか？`)) {
      return 'cancelled';
    }
    const fileHandle = await EXPORT_DIR_HANDLE.getFileHandle(filename, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
    return 'dir';
  } catch (e) {
    downloadBlob(blob, filename);
    return 'download-failed';
  }
}

function saveBlob(blob, filename, onDone) {
  if (EXPORT_DIR_HANDLE) {
    saveToExportDir(blob, filename).then(onDone);
    return;
  }
  downloadBlob(blob, filename);
  onDone('download');
}

function statusMessageForSave(result, filename, downloadMessage) {
  if (result === 'dir') return `指定フォルダへ保存しました: ${filename}`;
  if (result === 'cancelled') return `書き出しを中止しました（「${filename}」は指定フォルダに既にあります）。`;
  if (result === 'download-failed') return `⚠️ 指定フォルダへの保存に失敗したため、通常のダウンロードで保存しました: ${filename}`;
  return downloadMessage;
}

async function pickExportDirectory() {
  if (!window.showDirectoryPicker) return;
  try {
    EXPORT_DIR_HANDLE = await window.showDirectoryPicker({ mode: 'readwrite' });
  } catch (e) {
    return;
  }
  $('#export-dir-status').textContent = `書き出し先：${EXPORT_DIR_HANDLE.name}（以後この集約ツールからの書き出しはここへ直接保存されます）`;
}

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

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsText(file, 'utf-8');
  });
}

// filler_app.jsのbuildStateFromStructureと同じ（STRUCTURE.cellsからgrid Mapを復元する）。
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
    fileName: structure.formTitle || 'shuyaku',
    manualGroups: structure.manualGroups || [],
  };
}

// STRUCTUREに含まれる全ての「単独の入力欄」候補（住民公開設定①〜③の共通プール）。
// グループ化された列（配列項目）は住民公開設定の対象外とする（1レコード＝1事業という
// 前提のシンプルな一覧・検索を優先し、配列項目まで一覧に列展開する複雑さは持たせない）。
function candidateSingles() {
  const { singles } = CoreLogic.findMappingTargets(CURRENT_STATE.grid, CURRENT_STATE.sections, CURRENT_STATE.maxCol, CURRENT_STATE.manualGroups);
  return singles;
}

function numericCandidateSingles() {
  return candidateSingles().filter(t => {
    const rt = t.cells[0].renderType;
    return rt === 'number' || rt === 'currency';
  });
}

// 繰り返し行（グループ化された列。例：Doの「決算見込額」8行分・「事業内容と活動実績」8行分）
// の全候補。findMappingTargetsのgroupsは、Excel上の合計セル（=SUM(...)数式）と同じ範囲を
// 指す生データの列そのもの（数式セルは書き出し時点でJSONに一切含まれない設計のため）。
function allGroupCandidates() {
  const { groups } = CoreLogic.findMappingTargets(CURRENT_STATE.grid, CURRENT_STATE.sections, CURRENT_STATE.maxCol, CURRENT_STATE.manualGroups);
  return groups;
}

// 繰り返し行のうち数値・金額型のもの（①表示項目・②集計対象で使用）。住民公開ページ側で
// 「合計」を再現するには、この列を事業ごとにまず合計し、事業間でさらに合計する必要がある
// （Excelの=SUM(...)と同じ範囲を合計するので数値的には同じ結果になる）。テキスト型の
// 繰り返し列を1つの数値に要約する明確な方法が無いため、表示・集計の対象は数値・金額型のみ。
function groupCandidatesNumeric() {
  return allGroupCandidates().filter(t => {
    const rt = t.cells[0].renderType;
    return rt === 'number' || rt === 'currency';
  });
}

// 集計対象（住民公開設定②）の候補：単独の入力欄＋繰り返し列の両方を、シート上の並び順
// （先頭セルの行・列）でまとめる。
function numericAggregateCandidates() {
  return numericCandidateSingles().concat(groupCandidatesNumeric())
    .sort((a, b) => (a.cells[0].row - b.cells[0].row) || (a.cells[0].col - b.cells[0].col));
}

// ③検索対象項目の候補：単独の入力欄＋繰り返し列（型を問わず全て）。②・①と違い、検索は
// 複数行のテキストを1本の検索用文字列にまとめて連結するだけでよく、数値でなくても
// （むしろテキスト型の繰り返し列こそ）意味があるため、型で絞り込まない。
function searchCandidatePool() {
  return candidateSingles().concat(allGroupCandidates())
    .sort((a, b) => (a.cells[0].row - b.cells[0].row) || (a.cells[0].col - b.cells[0].col));
}

// 集計対象の選択状態（CONFIG.aggregateFieldIds/aggregateUnits）のキー。単独欄はcellIdで
// 十分だが、繰り返し列はセルが複数あるため列番号＋グループ名で一意に識別する。
function aggregateKey(target) {
  return target.kind === 'group' ? ('g:' + target.col + ':' + target.groupLabel) : ('s:' + cellIdOf(target));
}

function cellIdOf(target) { return CoreLogic.cellId(target.cells[0]); }

function labelForTarget(target) {
  return target.cells[0].dbKey || target.autoName || CoreLogic.cellRef(target.cells[0]);
}

// cellId → 表示名。住民公開設定パネル・書き出し前の確認等、対象セルがtargetオブジェクトの
// 形で手元に無い場面向け（labelForTargetはtargetを直接受け取れる場面専用）。
function labelForCellId(id) {
  const target = candidateSingles().find(t => cellIdOf(t) === id);
  if (target) return labelForTarget(target);
  return id;
}

// ---------------------------------------------------------------------------
// STEP1：JSONファイルの読み込み（filler_app.jsのREVIEW.records方式を踏襲。
// レビュー欄・列フィルタ等は持たず、読み込んだ生データを保持するだけの簡易版）。
// ---------------------------------------------------------------------------
const AGG = { records: [], dirHandle: null };

function upsertRecords(newRecords) {
  const existingNames = new Set(AGG.records.map(r => r.fileName));
  let added = 0;
  newRecords.forEach((r) => {
    if (existingNames.has(r.fileName)) return;
    AGG.records.push({ fileName: r.fileName, data: r.data });
    existingNames.add(r.fileName);
    added++;
  });
  return added;
}

function removeRecord(fileName) {
  AGG.records = AGG.records.filter(r => r.fileName !== fileName);
}

function renderLoadedList() {
  $('#loaded-summary').textContent = `${AGG.records.length}件 読み込み済み`;
  const root = $('#loaded-table-root');
  root.innerHTML = '';
  if (AGG.records.length === 0) return;
  const table = el('table', { class: 'loaded-table' });
  table.appendChild(el('thead', {}, [el('tr', {}, [el('th', { text: 'ファイル名' }), el('th', { text: '' })])]));
  const tbody = el('tbody');
  AGG.records.forEach((r) => {
    const tr = el('tr');
    tr.appendChild(el('td', { text: r.fileName }));
    const btnTd = el('td', { class: 'loaded-remove-cell' });
    const removeBtn = el('button', { class: 'secondary', type: 'button', text: '削除' });
    removeBtn.addEventListener('click', () => {
      removeRecord(r.fileName);
      renderLoadedList();
      renderConfigDescriptionForm();
      setStatus(`「${r.fileName}」を読み込み対象から除きました（合計${AGG.records.length}件）。`);
    });
    btnTd.appendChild(removeBtn);
    tr.appendChild(btnTd);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  root.appendChild(table);
}

async function handleFileInput(fileList) {
  const files = Array.from(fileList || []);
  const newRecords = [];
  for (const file of files) {
    if (!file.name.toLowerCase().endsWith('.json')) continue;
    try {
      const text = await readFileAsText(file);
      newRecords.push({ fileName: file.name, data: JSON.parse(text) });
    } catch (e) { /* 壊れたJSONはスキップ */ }
  }
  const added = upsertRecords(newRecords);
  renderLoadedList();
  renderConfigDescriptionForm();
  setStatus(`${added}件読み込みました（合計${AGG.records.length}件）。`);
}

async function pickLoadDirectory() {
  if (!window.showDirectoryPicker) return;
  try {
    AGG.dirHandle = await window.showDirectoryPicker();
  } catch (e) {
    return;
  }
  $('#btn-refresh-load-dir').style.display = '';
  await scanLoadDirectory();
}

async function scanLoadDirectory() {
  if (!AGG.dirHandle) return;
  const newRecords = [];
  for await (const entry of AGG.dirHandle.values()) {
    if (entry.kind !== 'file' || !entry.name.toLowerCase().endsWith('.json')) continue;
    const file = await entry.getFile();
    try {
      const text = await readFileAsText(file);
      newRecords.push({ fileName: entry.name, data: JSON.parse(text) });
    } catch (e) { /* 壊れたJSONはスキップ */ }
  }
  const added = upsertRecords(newRecords);
  renderLoadedList();
  renderConfigDescriptionForm();
  setStatus(`フォルダを読み直しました（新規${added}件、合計${AGG.records.length}件）。`);
}

// ---------------------------------------------------------------------------
// STEP2：住民公開設定（①表示項目②集計対象③検索対象④説明文・出典）
// ---------------------------------------------------------------------------
const CONFIG = {
  displayFieldIds: new Set(),
  aggregateFieldIds: new Set(), // aggregateKey()の値（単独欄="s:cellId"／繰り返し列="g:col:groupLabel"）
  aggregateUnits: new Map(), // aggregateKey()の値 -> unit文字列
  showRecordCount: true,
  searchFieldIds: new Set(),
};

function defaultTitle() { return STRUCTURE.formTitle || '公表データ'; }
function defaultSource() { return `${defaultTitle()}（${AGG.records.length}件）を基に作成`; }
function defaultUpdatedDate() {
  const d = new Date();
  return d.toISOString().slice(0, 10);
}

// ①一覧・詳細に表示する項目の候補プール：単独欄（全型）＋繰り返し列（グループ、数値・金額型
// のみ）。②集計対象と同じ繰り返し列を対象にする（テキスト型の繰り返し列を1セルに要約する
// 明確な方法が無いため対象外のまま。数値・金額型は「事業内で合計して1つの数値として表示」
// という②と同じ考え方で表示できる）。
function displayCandidatePool() {
  return candidateSingles().concat(groupCandidatesNumeric())
    .sort((a, b) => (a.cells[0].row - b.cells[0].row) || (a.cells[0].col - b.cells[0].col));
}

// STEP4のbuildGenericFieldPicker（builder_app.js）と同じ見た目・考え方の簡易版。
// このツールでは事務局が全事業を見た後に一度だけ設定する想定のため、force-checked等の
// 特殊対応は持たせず、単純なcheckbox一覧にとどめる。
// opts.keyOf／opts.refLabel（省略可）：候補が単独欄のみのリスト（③検索対象等）では
// 既定のcellIdOf／単一セル番地のままでよいが、①のように繰り返し列も混ざるリストでは
// aggregateKey／aggregateRefLabel（後述）を渡して単独欄と区別できるキー・範囲表示にする。
function buildCheckboxPicker(candidates, opts) {
  if (candidates.length === 0) return el('p', { class: 'hint', text: '対象となる項目がありません。' });
  const keyOf = opts.keyOf || cellIdOf;
  const refLabel = opts.refLabel || (t => CoreLogic.cellRef(t.cells[0]));
  const wrap = el('div', { class: 'filename-field-picker' });
  if (opts.noteText) wrap.appendChild(el('p', { class: 'hint', text: opts.noteText }));
  const list = el('div', { class: 'filename-field-list' });
  candidates.forEach((target) => {
    const key = keyOf(target);
    const inputId = opts.idPrefix + '_' + key;
    const cb = el('input', { type: 'checkbox', id: inputId });
    if (opts.isChecked(key)) cb.setAttribute('checked', 'checked');
    cb.addEventListener('change', () => opts.onToggle(key, cb.checked));
    const label = `${labelForTarget(target)}（${refLabel(target)}）`;
    list.appendChild(el('label', { for: inputId, class: 'filename-field-item' }, [cb, el('span', { text: label })]));
  });
  wrap.appendChild(list);
  return wrap;
}

function renderDisplayFieldPicker() {
  const root = $('#config-display-root');
  root.innerHTML = '';
  root.appendChild(buildCheckboxPicker(displayCandidatePool(), {
    idPrefix: 'cfgdisplay',
    keyOf: aggregateKey,
    refLabel: aggregateRefLabel,
    noteText: '一覧表の列・詳細画面の概要欄に表示する項目を選びます（複数可）。「〜」（範囲表示）の項目は繰り返し行の列で、値はこの事業内で合計して1つの数値として表示します。',
    isChecked: (key) => CONFIG.displayFieldIds.has(key),
    onToggle: (key, checked) => { if (checked) CONFIG.displayFieldIds.add(key); else CONFIG.displayFieldIds.delete(key); },
  }));
}

function renderSearchFieldPicker() {
  const root = $('#config-search-root');
  root.innerHTML = '';
  root.appendChild(buildCheckboxPicker(searchCandidatePool(), {
    idPrefix: 'cfgsearch',
    keyOf: aggregateKey,
    refLabel: aggregateRefLabel,
    noteText: 'キーワード検索の対象にする項目を選びます（①の表示項目とは独立して選べます）。「〜」（範囲表示）の項目は繰り返し行の列で、行ごとの値をまとめて検索対象にします。',
    isChecked: (key) => CONFIG.searchFieldIds.has(key),
    onToggle: (key, checked) => { if (checked) CONFIG.searchFieldIds.add(key); else CONFIG.searchFieldIds.delete(key); },
  }));
}

// ラベル表示：単独欄はセル番地1つ、繰り返し列は範囲（例：K11:K18）で示す
// （builder_app.jsのmappingRefLabelと同じ考え方）。
function aggregateRefLabel(target) {
  if (target.kind !== 'group') return CoreLogic.cellRef(target.cells[0]);
  const first = CoreLogic.cellRef(target.cells[0]);
  const last = CoreLogic.cellRef(target.cells[target.cells.length - 1]);
  return `${first}:${last}`;
}

function renderAggregateFieldPicker() {
  const root = $('#config-aggregate-root');
  root.innerHTML = '';

  // ブロック1：レコード数表示の可否（他の選択肢と見分けやすいよう独立した白ブロックにする）。
  const countBlock = el('div', { class: 'filename-field-list' });
  const countItem = el('label', { class: 'agg-field-item' });
  const countCb = el('input', { type: 'checkbox', id: 'cfgagg-count' });
  if (CONFIG.showRecordCount) countCb.setAttribute('checked', 'checked');
  countCb.addEventListener('change', () => { CONFIG.showRecordCount = countCb.checked; });
  countItem.appendChild(countCb);
  countItem.appendChild(el('span', { text: '絞り込み結果のレコード数を表示する' }));
  countBlock.appendChild(countItem);
  root.appendChild(countBlock);

  // ブロック2：数値・金額項目（単独欄＋繰り返し列）。
  const numeric = numericAggregateCandidates();
  if (numeric.length === 0) {
    root.appendChild(el('p', { class: 'hint', text: '数値・金額型の項目がありません。' }));
    return;
  }
  root.appendChild(el('p', {
    class: 'hint',
    text: '合計を集計したい数値・金額項目を選び、単位（任意）を入力してください（例：千円）。'
      + '「〜」（範囲表示）の項目は繰り返し行の列で、この事業内でまず合計してから、事業間でさらに合計します'
      + '（Excelの合計セルと同じ範囲を集計するため、数式セル自体を選ぶ必要はありません）。',
  }));
  const list = el('div', { class: 'filename-field-list' });
  numeric.forEach((target) => {
    const key = aggregateKey(target);
    const inputId = 'cfgagg_' + key;
    const cb = el('input', { type: 'checkbox', id: inputId });
    if (CONFIG.aggregateFieldIds.has(key)) cb.setAttribute('checked', 'checked');
    cb.addEventListener('change', () => {
      if (cb.checked) CONFIG.aggregateFieldIds.add(key); else CONFIG.aggregateFieldIds.delete(key);
    });
    const unitInput = el('input', { type: 'text', placeholder: '単位（例：千円）', value: CONFIG.aggregateUnits.get(key) || '' });
    unitInput.addEventListener('input', () => { CONFIG.aggregateUnits.set(key, unitInput.value); });
    const label = `${labelForTarget(target)}（${aggregateRefLabel(target)}）`;
    list.appendChild(el('label', { for: inputId, class: 'agg-field-item' }, [cb, el('span', { text: label }), unitInput]));
  });
  root.appendChild(list);
}

// ④の初期値は「未入力なら埋める」だけでは不十分：STEP1でJSONを読み込む前（0件）に
// 一度自動入力されると、値が空欄でなくなるため、後から件数が変わっても再計算されない
// （出典欄「〜件を基に作成」が0件のまま固定されてしまう不具合）。そこで「前回この関数が
// 自動入力した値のまま（＝ユーザーが編集していない）」場合だけ最新の既定値で上書きする。
const AUTO_FILLED = { title: null, source: null, updated: null };
function fillIfUntouched(inputEl, key, defaultFn) {
  if (inputEl.value === '' || inputEl.value === AUTO_FILLED[key]) {
    inputEl.value = defaultFn();
    AUTO_FILLED[key] = inputEl.value;
  }
}
function renderConfigDescriptionForm() {
  fillIfUntouched($('#cfg-title'), 'title', defaultTitle);
  fillIfUntouched($('#cfg-source'), 'source', defaultSource);
  fillIfUntouched($('#cfg-updated'), 'updated', defaultUpdatedDate);
}

function renderConfigScreen() {
  renderDisplayFieldPicker();
  renderAggregateFieldPicker();
  renderSearchFieldPicker();
  renderConfigDescriptionForm();
}

// ---------------------------------------------------------------------------
// STEP3：住民公開用ページの書き出し
// ---------------------------------------------------------------------------
function slugTitle(s) {
  const t = sanitizeForFileName(s).replace(/\s+/g, '_');
  return t || 'viewer';
}

// 検索対象は、表示・集計と違って「1本の検索用文字列に連結するだけ」でよいため、単独欄・
// 繰り返し列を区別せず、選ばれた項目の実セルidを全部フラットに1つの配列へ詰めて返す
// （型を問わないので、テキスト型の繰り返し列（例：「事業内容と活動実績」8行）も対象にできる）。
function buildSearchFieldIds() {
  const out = [];
  searchCandidatePool()
    .filter(t => CONFIG.searchFieldIds.has(aggregateKey(t)))
    .forEach(t => out.push(...t.cells.map(c => CoreLogic.cellId(c))));
  return out;
}

// 集計対象は、単独欄・繰り返し列のどちらも{label, unit, ids}という同じ形に正規化して返す
// （idsは合計対象の実セルid配列。単独欄は要素1件、繰り返し列は行数分。住民公開ページ側は
// 種別を意識せず「idsの値を全部合計する」だけでよくなる）。
function buildAggregateFieldConfigs() {
  return numericAggregateCandidates()
    .filter(t => CONFIG.aggregateFieldIds.has(aggregateKey(t)))
    .map(t => ({
      label: labelForTarget(t),
      unit: CONFIG.aggregateUnits.get(aggregateKey(t)) || '',
      ids: t.cells.map(c => CoreLogic.cellId(c)),
    }));
}

// 表示項目も集計対象と同じ{label, ids}形に正規化する（idsは1件＝単独欄、複数件＝繰り返し列）。
// 住民公開ページ側は、ids.length===1ならそのまま値を表示し、複数件なら②と同じく
// 事業内で合計して1つの数値として表示する（テキスト型の繰り返し列は候補にそもそも
// 含めていないため、複数件は必ず数値・金額型）。
function buildDisplayFieldConfigs() {
  return displayCandidatePool()
    .filter(t => CONFIG.displayFieldIds.has(aggregateKey(t)))
    .map(t => ({
      label: labelForTarget(t),
      ids: t.cells.map(c => CoreLogic.cellId(c)),
    }));
}

function buildPublicConfig() {
  const title = ($('#cfg-title').value || '').trim() || defaultTitle();
  return {
    storageKey: 'jimujigyou_viewer_' + slugTitle(title),
    displayFields: buildDisplayFieldConfigs(),
    aggregateFields: buildAggregateFieldConfigs(),
    showRecordCount: CONFIG.showRecordCount,
    searchFields: buildSearchFieldIds(),
    title,
    description: ($('#cfg-description').value || '').trim(),
    source: ($('#cfg-source').value || '').trim() || defaultSource(),
    updatedDate: ($('#cfg-updated').value || '').trim() || defaultUpdatedDate(),
  };
}

function escapeScriptClose(jsonLiteral) {
  return jsonLiteral.replace(/<\/(script)/gi, '<\\/$1');
}

function doExportAsViewer() {
  if (AGG.records.length === 0) {
    setStatus('⚠ JSONファイルが1件も読み込まれていません。STEP1で読み込んでから書き出してください。');
    return;
  }
  const publicConfig = buildPublicConfig();
  const recordsData = AGG.records.map(r => r.data);
  const html = VIEWER_TEMPLATE
    .replace('/* __STRUCTURE__ */', escapeScriptClose(JSON.stringify(STRUCTURE)))
    .replace('/* __RECORDS__ */', escapeScriptClose(JSON.stringify(recordsData)))
    .replace('/* __PUBLIC_CONFIG__ */', escapeScriptClose(JSON.stringify(publicConfig)))
    .replace(/__VIEWER_TITLE__/g, publicConfig.title);
  const filename = `${publicConfig.title}_公開ページ.html`;
  const blob = new Blob([html], { type: 'text/html' });
  saveBlob(blob, filename, (result) => {
    setStatus(statusMessageForSave(result, filename, `住民公開用ページを書き出しました（${AGG.records.length}件）。`));
  });
}

// 読み込んだ全事業の生データを1つのJSONファイルにまとめて書き出す。住民公開ページとは
// 別に、庁内での再利用・バックアップ・他ツールでの二次利用（Excelでの二次加工、別の
// 集約ツールへの再読込等）を想定した汎用出力。ネスト構造をそのまま保つ（表形式が
// 欲しい場合はCSV書き出し（doExportAggregatedCsv）を使う）。
function doExportAggregatedJson() {
  if (AGG.records.length === 0) {
    setStatus('⚠ JSONファイルが1件も読み込まれていません。STEP1で読み込んでから書き出してください。');
    return;
  }
  const payload = {
    meta: {
      exported_at: new Date().toISOString(),
      source_title: STRUCTURE.formTitle || '',
      count: AGG.records.length,
    },
    records: AGG.records.map(r => ({ fileName: r.fileName, data: r.data })),
  };
  const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const filename = `${sanitizeForFileName(STRUCTURE.formTitle || 'jimujigyou')}_集約データ_${dateStr}.json`;
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  saveBlob(blob, filename, (result) => {
    setStatus(statusMessageForSave(result, filename, `集約データを書き出しました（${AGG.records.length}件）。`));
  });
}

// CSV書き出し用：単独欄＋繰り返し列（型を問わず全部）をシート上の並び順で列にする。
// 繰り返し列は、事業内の複数行の値を「／」区切りで1セルにまとめる（1事業＝1行のCSVに
// するため。行を分けたい場合はJSON書き出しの方を使う想定）。
function allFieldCandidatesSorted() {
  return candidateSingles().concat(allGroupCandidates())
    .sort((a, b) => (a.cells[0].row - b.cells[0].row) || (a.cells[0].col - b.cells[0].col));
}

// scratchExtract（viewer_app.js）と同じ「使い捨てグリッドに読み込ませてDOM経由で読む」方式。
// 集約ツールは通常データの値を読まない（住民公開設定はSTRUCTUREの構造だけを見る）ため、
// CSV書き出し専用にここへ実装する。
function scratchExtractForRecord(data, cellIds) {
  if (!cellIds || cellIds.length === 0) return {};
  const scratch = $('#agg-scratch-root');
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

// CSVフィールドのエスケープ（RFC4180準拠：カンマ・ダブルクォート・改行を含む場合のみ
// ダブルクォートで囲み、内部のダブルクォートは2つに重ねる）。本物のCSVファイルとして
// ダウンロードするため（コピー＆手動貼り付けではない）、Excel等の標準的なCSVパーサーが
// 正しく解釈できる。
function csvEscape(v) {
  const s = String(v == null ? '' : v);
  if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function doExportAggregatedCsv() {
  if (AGG.records.length === 0) {
    setStatus('⚠ JSONファイルが1件も読み込まれていません。STEP1で読み込んでから書き出してください。');
    return;
  }
  const columns = allFieldCandidatesSorted().map(t => ({
    label: labelForTarget(t),
    ids: t.cells.map(c => CoreLogic.cellId(c)),
  }));
  const allIds = columns.flatMap(c => c.ids);

  const rows = [['ファイル名', ...columns.map(c => c.label)]];
  AGG.records.forEach((r) => {
    const raw = scratchExtractForRecord(r.data, allIds);
    const row = [r.fileName];
    columns.forEach((c) => {
      const values = c.ids.map(id => raw[id] || '').filter(v => v !== '');
      row.push(values.join(' / '));
    });
    rows.push(row);
  });

  const csvText = rows.map(row => row.map(csvEscape).join(',')).join('\r\n');
  const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const filename = `${sanitizeForFileName(STRUCTURE.formTitle || 'jimujigyou')}_集約データ_${dateStr}.csv`;
  // 先頭にBOMを付与し、Excelで開いたときの文字化けを防ぐ。
  const blob = new Blob(['﻿' + csvText], { type: 'text/csv' });
  saveBlob(blob, filename, (result) => {
    setStatus(statusMessageForSave(result, filename, `CSVを書き出しました（${AGG.records.length}件）。`));
  });
}

function init() {
  CURRENT_STATE = buildStateFromStructure(STRUCTURE);

  if (window.showDirectoryPicker) {
    $('#btn-pick-load-dir').style.display = '';
    $('#export-dir-bar').style.display = 'flex';
  }

  $('#load-file-input').addEventListener('change', (ev) => {
    handleFileInput(ev.target.files);
    ev.target.value = '';
  });
  $('#btn-pick-load-dir').addEventListener('click', pickLoadDirectory);
  $('#btn-refresh-load-dir').addEventListener('click', scanLoadDirectory);
  $('#btn-pick-export-dir').addEventListener('click', pickExportDirectory);
  $('#btn-export-viewer').addEventListener('click', doExportAsViewer);
  $('#btn-export-aggregated-json').addEventListener('click', doExportAggregatedJson);
  $('#btn-export-aggregated-csv').addEventListener('click', doExportAggregatedCsv);

  renderLoadedList();
  renderConfigScreen();
}

if (typeof window !== 'undefined') {
  window.addEventListener('DOMContentLoaded', init);
  window.__app = {
    get AGG() { return AGG; },
    get CONFIG() { return CONFIG; },
    get CURRENT_STATE() { return CURRENT_STATE; },
    handleFileInput, upsertRecords, removeRecord, renderLoadedList,
    pickLoadDirectory, scanLoadDirectory,
    candidateSingles, numericCandidateSingles, allGroupCandidates, groupCandidatesNumeric, numericAggregateCandidates,
    displayCandidatePool, searchCandidatePool, aggregateKey, labelForCellId, labelForTarget, buildStateFromStructure,
    renderConfigScreen, buildPublicConfig, buildAggregateFieldConfigs, buildDisplayFieldConfigs, buildSearchFieldIds, doExportAsViewer, doExportAggregatedJson, doExportAggregatedCsv, slugTitle,
    allFieldCandidatesSorted, scratchExtractForRecord, csvEscape,
    saveBlob, pickExportDirectory,
    get EXPORT_DIR_HANDLE() { return EXPORT_DIR_HANDLE; },
    setExportDirHandle: (h) => { EXPORT_DIR_HANDLE = h; },
    init,
  };
}
