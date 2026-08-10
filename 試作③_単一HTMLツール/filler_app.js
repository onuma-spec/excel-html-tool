// 生成された入力フォーム（ツール2）側のロジック。
// ビルダー（ツール1）が埋め込んだ STRUCTURE（値ではなく確定済みの構造）を読み込み、
// グリッドを描画し、入力・書き出し・読み込みだけを行う。Excel読み込みや手直し機能は持たない。

let STATE = null;

// 書き出し先フォルダ（Chromium限定・File System Access API）。指定されていれば
// 通常のダウンロードダイアログを介さず直接このフォルダへ書き込む。未指定・非対応
// ブラウザでは従来通り<a download>でブラウザの既定のダウンロード先へ保存する
// （REVIEW.dirHandleと同じ発想だが、読込用と書き出し用でハンドルは分けている）。
let EXPORT_DIR_HANDLE = null;

function $(sel) { return document.querySelector(sel); }

// ファイル名に使えない文字（Windows/Mac共通で問題になりやすいもの）と改行・タブを除去し、
// 長すぎる値でファイル名が壊れないよう適度な長さに切り詰める。
function sanitizeForFileName(s) {
  return String(s == null ? '' : s).trim().replace(/[\\/:*?"<>|\r\n\t]/g, '').slice(0, 40);
}

// ファイル名は「フォームのタイトル」＋「STEP4で選ばれた項目の現在値（シート上の並び順）」＋
// 「書き出し日（YYYYMMDD、区切りなし）」を_で連結した形にする。
// 例：R7年度実施 事務事業評価_出産・子育て応援事業_こども未来部_こども未来課_20260809.json
// 項目が何も選ばれていない・全て空欄の場合はタイトル＋日付のみになる。
function buildExportFileName() {
  const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const title = sanitizeForFileName(STRUCTURE.formTitle);
  const values = (STRUCTURE.fileNameFields || [])
    .map(id => { const inputEl = document.getElementById(id); return inputEl ? sanitizeForFileName(inputEl.value) : ''; })
    .filter(v => v !== '');
  const parts = [];
  if (title) parts.push(title);
  parts.push(...values);
  parts.push(dateStr);
  return parts.join('_') + '.json';
}

// blobを通常の<a download>でダウンロードする（従来からの挙動・フォールバック用）。
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

// EXPORT_DIR_HANDLEへ直接書き込む。同名ファイルが既にあれば上書き確認をはさむ
// （getFileHandle(name,{create:true})は無警告で上書きしてしまうため）。フォルダへの
// 書き込みに失敗した場合（権限が切れた等）は、通常のダウンロードにフォールバックする。
// 戻り値："dir"（直接保存）｜"download"（フォールバック）｜"cancelled"（上書きを拒否）。
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
    return 'download';
  }
}

// EXPORT_DIR_HANDLEが指定されていれば非同期でフォルダへ直接書き込み、無ければ
// 従来通り即座にダウンロードする（同期）。onDoneには結果（"dir"|"download"|"cancelled"）
// が渡る。あえて2経路に分けているのは、常にasync/awaitへ統一すると、フォルダ未指定の
// 場合でも最低1マイクロタスク分の遅延が生まれ、書き出し直後に#statusを同期的に
// 確認している既存の挙動・テストが崩れるため（フォルダ未指定時は今まで通り同期のまま）。
function saveBlob(blob, filename, onDone) {
  if (EXPORT_DIR_HANDLE) {
    saveToExportDir(blob, filename).then(onDone);
    return;
  }
  downloadBlob(blob, filename);
  onDone('download');
}

// saveBlobの結果に応じたステータス文言を組み立てる。downloadMessageはフォルダ未指定時
// （従来通りダウンロードした場合）に表示する文言を呼び出し側から渡す。
function statusMessageForSave(result, filename, downloadMessage) {
  if (result === 'dir') return `指定フォルダへ保存しました: ${filename}`;
  if (result === 'cancelled') return `書き出しを中止しました（「${filename}」は指定フォルダに既にあります）。`;
  return downloadMessage;
}

// Chromium限定：書き出し先フォルダを1回指定する。以後、このフォームからのJSON書き出しは
// （通常画面・レビュー画面のどちらからでも）このフォルダへ直接保存される。
// window.showDirectoryPickerが無いブラウザ（Firefox/Safari）では、対応するボタン自体を
// 表示しない（レビュー画面の読込用フォルダ指定と同じ扱い）。
async function pickExportDirectory() {
  if (!window.showDirectoryPicker) return;
  try {
    EXPORT_DIR_HANDLE = await window.showDirectoryPicker({ mode: 'readwrite' });
  } catch (e) {
    return; // ユーザーがキャンセルした場合等。エラー表示はしない。
  }
  $('#export-dir-status').textContent = `書き出し先：${EXPORT_DIR_HANDLE.name}（以後このフォームのJSONはここへ直接保存されます）`;
}

// 通常グリッド・詳細画面のどちらでも同じ「1次/2次以降」の色分けを使うための共通Set。
function reviewFieldIdSet() {
  return new Set(STRUCTURE.reviewFields || []);
}

function buildStateFromStructure(structure) {
  const grid = new Map();
  structure.cells.forEach(c => {
    // isFormula/hasText/blocked等はビルダー側で確定済みの値をそのまま使う。
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
    fileName: structure.formTitle || 'nyuryoku',
    manualGroups: structure.manualGroups || [],
  };
}

// STEP4で指定された必須項目のうち、現在DOM上で空欄のものだけを返す（cellIdの配列）。
function findMissingRequiredFields() {
  return (STRUCTURE.requiredFields || []).filter((id) => {
    const inputEl = document.getElementById(id);
    return !inputEl || !String(inputEl.value || '').trim();
  });
}

// 未入力チェック・一覧表の列見出し等で、cellIdから人間が読めるラベルを求める。
// builder_app.jsのserializeStructure()と同じCoreLogic.findMappingTargets()を使い、
// 「dbKey（項目名マッピング済みならそれ）→行見出し文字/列見出し文字（autoName）→
// セル番地」の順でフォールバックする。以前はautoNameを見ていなかったため、
// 「1行1見出し1値」のようなdbKey未設定の項目がE5等の生セル番地でしか表示できなかった
// （行の見出し文字「部署名」等が本来使えたはずだった、という既知の手抜き実装）。
function labelForCellId(id) {
  const { singles } = CoreLogic.findMappingTargets(STATE.grid, STATE.sections, STATE.maxCol, STATE.manualGroups);
  const target = singles.find(t => CoreLogic.cellId(t.cells[0]) === id);
  if (target) return target.cells[0].dbKey || target.autoName || CoreLogic.cellRef(target.cells[0]);
  for (const info of STATE.grid.values()) {
    if (CoreLogic.cellId(info) === id) return info.dbKey || CoreLogic.cellRef(info);
  }
  return id;
}

// ---------------------------------------------------------------------------
// レビュー画面（複数事業のJSONを一括読込し、判定/コメント等のレビュー欄を
// 一覧表で入力する）。所管部署欄・レビュー欄は同じ様式（同じSTRUCTURE）に
// 含まれている前提のため、フィラー1個の中に「通常の1件入力画面」と
// 「複数件レビュー画面」の2モードを持たせ、ボタンで切り替える構成にしている。
// ---------------------------------------------------------------------------

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

// records: [{ fileName, data(読み込んだ生JSON), displayValues({cellId:value}), reviewValues({cellId:value}) }]
// visibleCols: 一覧表に列として表示するdisplayCandidateFieldsの部分集合（初期値は全件表示）。
// dirHandle: showDirectoryPicker()で選んだFileSystemDirectoryHandle（Chromium限定）。
// colFilters: Map<cellId, Set<選択中の値>>（Excel風のオートフィルタ。未操作の列はエントリ無し＝絞り込みなし）。
// statusFilter: 'all'|'done'|'undone'（レビュー欄が全て埋まっているかによる絞り込み）。
const REVIEW = { records: [], visibleCols: new Set(), dirHandle: null, colFilters: new Map(), statusFilter: 'all' };

// 詳細画面（openReviewDetail）で現在編集中のレコード・グリッド状態。
// 全セル誰でも編集可能という方針のため、詳細画面は所管部署の入力を閲覧するだけでなく
// 一件ずつ入力する経路としても使う。保存時にDETAIL_STATEから値を読み出し、
// DETAIL_RECORDへ書き戻す（破棄する場合は単に画面を閉じるだけでよい）。
let DETAIL_RECORD = null;
let DETAIL_STATE = null;

// STRUCTURE.cells（同じ様式）を土台に、渡されたdataを読み込んだ「使い捨てのグリッド」を
// #review-scratch-rootに描画→指定したcellIdの現在値を読み取る→片付ける、という
// 一連の処理をまとめる。既存のrenderGrid/loadDataIntoGridをそのまま再利用することで、
// JSONのネスト構造（セクション→グループ→項目）をたどる特別なロジックを新たに
// 書かずに済む（このアプリ全体で「値の出し入れは常にDOM経由」という設計を踏襲）。
function extractFieldValues(data, cellIds) {
  if (!cellIds || cellIds.length === 0) return {};
  // 同じSTRUCTURE（同じcellId体系）を使うレビュー詳細画面と同時に存在すると
  // document.getElementByIdが衝突するため、念のため詳細画面側も空にしておく。
  const detailRoot = $('#review-detail-root');
  if (detailRoot) detailRoot.innerHTML = '';
  const scratch = $('#review-scratch-root');
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

// record.reviewValues（一覧表で入力した現在値）を、所管部署の元データrecord.dataに
// 上書きマージしたJSONを組み立てる（書き出し用）。extractFieldValuesと同じ使い捨て
// グリッドの仕組みを使い、読み込み直後にレビュー欄だけDOM上で値を差し替えてから
// collectData()で丸ごと読み出す（＝ネスト構造の組み立てを新たに書かない）。
function mergedDataForExport(record) {
  const detailRoot = $('#review-detail-root');
  if (detailRoot) detailRoot.innerHTML = '';
  const scratch = $('#review-scratch-root');
  const tempState = buildStateFromStructure(STRUCTURE);
  GridRender.renderGrid(scratch, tempState, { showGear: false });
  GridRender.loadDataIntoGrid(tempState, record.data || {});
  (STRUCTURE.reviewFields || []).forEach((id) => {
    const inputEl = document.getElementById(id);
    if (inputEl) inputEl.value = record.reviewValues[id] || '';
  });
  const merged = GridRender.collectData(tempState);
  scratch.innerHTML = '';
  return merged;
}

// 既に読み込み済みのファイル名（fileName）は上書きしない。理由：レビュアーが
// 一覧表で入力中の内容（record.reviewValues）を、フォルダの「更新」や再選択のたびに
// 消してしまわないようにするため。新しいファイルだけを追加する仕様。
function upsertRecords(newRecords) {
  const existingNames = new Set(REVIEW.records.map(r => r.fileName));
  let added = 0;
  newRecords.forEach((r) => {
    if (existingNames.has(r.fileName)) return;
    const displayValues = extractFieldValues(r.data, STRUCTURE.displayCandidateFields || []);
    const reviewValues = extractFieldValues(r.data, STRUCTURE.reviewFields || []);
    REVIEW.records.push({ fileName: r.fileName, data: r.data, displayValues, reviewValues });
    existingNames.add(r.fileName);
    added++;
  });
  return added;
}

// このレコードのレビュー欄が全て入力済みかどうか（一覧表の「状態」列・進捗表示に使う）。
function isRecordComplete(record) {
  const fields = STRUCTURE.reviewFields || [];
  return fields.length > 0 && fields.every(id => String(record.reviewValues[id] || '').trim() !== '');
}

function buildExportFileNameForRecord(record) {
  const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const title = sanitizeForFileName(STRUCTURE.formTitle);
  const values = (STRUCTURE.fileNameFields || [])
    .map(id => sanitizeForFileName(record.displayValues[id]))
    .filter(v => v !== '');
  const parts = [];
  if (title) parts.push(title);
  parts.push(...values);
  parts.push(dateStr);
  return parts.join('_') + '.json';
}

function renderColToggle() {
  const root = $('#review-col-toggle');
  root.innerHTML = '';
  const candidates = STRUCTURE.displayCandidateFields || [];
  if (candidates.length === 0) return;
  root.appendChild(el('span', { text: '一覧に表示する項目：' }));
  candidates.forEach((id) => {
    const inputId = 'colvis_' + id;
    const checkbox = el('input', { type: 'checkbox', id: inputId });
    if (REVIEW.visibleCols.has(id)) checkbox.setAttribute('checked', 'checked');
    checkbox.addEventListener('change', () => {
      if (checkbox.checked) REVIEW.visibleCols.add(id);
      else REVIEW.visibleCols.delete(id);
      renderReviewTable();
    });
    root.appendChild(el('label', { for: inputId }, [checkbox, el('span', { text: labelForCellId(id) })]));
  });
}

function currentVisibleColIds() {
  return (STRUCTURE.displayCandidateFields || []).filter(id => REVIEW.visibleCols.has(id));
}

// 今読み込まれている全レコードから、その列（表示中の識別列）が実際に取りうる値の
// 一覧を重複なく求める（Excelのオートフィルタのドロップダウン一覧と同じ考え方）。
function uniqueColumnValues(id) {
  const set = new Set();
  REVIEW.records.forEach(r => set.add(r.displayValues[id] || ''));
  return Array.from(set).sort((a, b) => String(a).localeCompare(String(b), 'ja'));
}

// 列フィルタ（REVIEW.colFilters: Map<cellId, Set<選択中の値>>）とステータス絞り込み
// （REVIEW.statusFilter）の両方を満たすレコードだけを返す。列フィルタは複数列に
// またがる場合AND条件（Excelと同じ）。非表示列のフィルタは無視する（列を隠したら
// その列による絞り込みも見えなくなるため、直感に反しないようにする）。
function matchesColFilters(record) {
  for (const id of currentVisibleColIds()) {
    const allowed = REVIEW.colFilters.get(id);
    if (!allowed) continue; // このセッションでまだ操作されていない列＝絞り込みなし
    if (!allowed.has(record.displayValues[id] || '')) return false;
  }
  return true;
}

function matchesStatusFilter(record) {
  if (REVIEW.statusFilter === 'done') return isRecordComplete(record);
  if (REVIEW.statusFilter === 'undone') return !isRecordComplete(record);
  return true;
}

function filteredRecords() {
  return REVIEW.records.filter(r => matchesColFilters(r) && matchesStatusFilter(r));
}

function hasActiveFilters() {
  return REVIEW.colFilters.size > 0 || REVIEW.statusFilter !== 'all';
}

function updateReviewSummary(filtered) {
  const summaryRoot = $('#review-summary');
  const doneCount = filtered.filter(isRecordComplete).length;
  const base = `${filtered.length}件中 ${doneCount}件 レビュー完了`;
  summaryRoot.textContent = hasActiveFilters() ? `${base}（絞り込み中・全${REVIEW.records.length}件）` : base;
}

// 列見出しの▼（Excel風のオートフィルタ）。<details>を使うことで開閉状態の管理・
// 外側クリックでの自動close処理を独自に書かずに済ませる（ネイティブ挙動に委ねる）。
// チェック変更時はrenderReviewTable()（thead含む全体再描画）ではなくrenderReviewBody()
// （tbodyのみ再描画）だけを呼ぶ：thead全体を再描画すると<details>が毎回閉じてしまい、
// 複数の値を続けてチェック/解除する操作が成立しなくなるため。
function buildColFilterDetails(id) {
  const values = uniqueColumnValues(id);
  const details = el('details', { class: 'col-filter' });
  details.appendChild(el('summary', { text: '▼', title: '値で絞り込む' }));
  const panel = el('div', { class: 'col-filter-panel' });
  values.forEach((v) => {
    const inputId = 'colfilter_' + id + '_' + values.indexOf(v);
    const checkbox = el('input', { type: 'checkbox', id: inputId });
    const currentAllowed = REVIEW.colFilters.get(id);
    if (!currentAllowed || currentAllowed.has(v)) checkbox.setAttribute('checked', 'checked');
    checkbox.addEventListener('change', () => {
      const allowed = new Set(REVIEW.colFilters.get(id) || values);
      if (checkbox.checked) allowed.add(v); else allowed.delete(v);
      REVIEW.colFilters.set(id, allowed);
      renderReviewBody();
    });
    panel.appendChild(el('label', {}, [checkbox, document.createTextNode(v === '' ? '（空欄）' : v)]));
  });
  details.appendChild(panel);
  return details;
}

function renderReviewTable() {
  const tableRoot = $('#review-table-root');
  tableRoot.innerHTML = '';

  if (REVIEW.records.length === 0) {
    updateReviewSummary([]);
    return;
  }

  const visibleColIds = currentVisibleColIds();
  const reviewFieldIds = STRUCTURE.reviewFields || [];

  const table = el('table', { class: 'review-table' });
  const thead = el('thead');
  const headRow = el('tr');
  visibleColIds.forEach((id) => {
    const th = el('th');
    th.appendChild(document.createTextNode(labelForCellId(id) + ' '));
    th.appendChild(buildColFilterDetails(id));
    headRow.appendChild(th);
  });
  reviewFieldIds.forEach(id => headRow.appendChild(el('th', { text: labelForCellId(id) })));
  headRow.appendChild(el('th', { text: '状態' }));
  headRow.appendChild(el('th', { text: '' }));
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = el('tbody', { id: 'review-tbody' });
  table.appendChild(tbody);
  tableRoot.appendChild(table);

  renderReviewBody();
}

// tbodyだけを絞り込み結果で再描画する（theadは触らない。理由はbuildColFilterDetails参照）。
function renderReviewBody() {
  const tbody = document.getElementById('review-tbody');
  if (!tbody) return;
  tbody.innerHTML = '';

  const visibleColIds = currentVisibleColIds();
  const reviewFieldIds = STRUCTURE.reviewFields || [];
  const filtered = filteredRecords();

  filtered.forEach((record) => {
    const tr = el('tr');
    visibleColIds.forEach(id => tr.appendChild(el('td', { text: record.displayValues[id] || '' })));
    reviewFieldIds.forEach((id) => {
      const textarea = el('textarea', { id: 'review_' + record.fileName + '_' + id });
      textarea.value = record.reviewValues[id] || '';
      textarea.addEventListener('input', () => {
        record.reviewValues[id] = textarea.value;
        const statusTd = tr.querySelector('.review-status-cell');
        if (statusTd) {
          const complete = isRecordComplete(record);
          statusTd.textContent = complete ? '✅完了' : '未完了';
          statusTd.classList.toggle('review-status-done', complete);
        }
        updateReviewSummary(filteredRecords());
      });
      const td = el('td'); td.appendChild(textarea);
      tr.appendChild(td);
    });
    const statusTd = el('td', { class: 'review-status-cell' + (isRecordComplete(record) ? ' review-status-done' : ''), text: isRecordComplete(record) ? '✅完了' : '未完了' });
    tr.appendChild(statusTd);

    const btnTd = el('td', { class: 'review-btn-cell' });
    const detailBtn = el('button', { class: 'secondary', type: 'button', text: '詳細' });
    detailBtn.addEventListener('click', () => openReviewDetail(record));
    btnTd.appendChild(detailBtn);
    const exportBtn = el('button', { type: 'button', text: '📋 書き出す' });
    exportBtn.addEventListener('click', () => {
      const merged = mergedDataForExport(record);
      const filename = buildExportFileNameForRecord(record);
      const blob = new Blob([JSON.stringify(merged, null, 2)], { type: 'application/json' });
      saveBlob(blob, filename, (result) => {
        $('#status').textContent = statusMessageForSave(result, filename, `「${record.fileName}」のレビュー結果を書き出しました。`);
      });
    });
    btnTd.appendChild(exportBtn);
    tr.appendChild(btnTd);

    tbody.appendChild(tr);
  });

  updateReviewSummary(filtered);
}

function renderReviewList() {
  renderColToggle();
  renderReviewTable();
}

// 画面切替：'normal'（通常の1件入力）／'list'（レビュー一覧表）／'detail'（レビュー詳細・閲覧専用）。
// #grid-rootと#review-detail-rootは同じSTRUCTURE（同じcellId体系）を使うため、
// 同時に中身を持たせるとdocument.getElementByIdが衝突する。表示していない方は
// 必ず空にしておく（切替のたびにDOMを作り直すコストより、事故を防ぐことを優先する）。
function showScreen(name) {
  $('#actions').style.display = name === 'normal' ? 'flex' : 'none';
  $('#grid-root').style.display = name === 'normal' ? '' : 'none';
  $('#grid-hint').style.display = name === 'normal' ? '' : 'none';
  $('#review-root').style.display = name === 'list' ? '' : 'none';
  $('#review-detail-root-wrap').style.display = name === 'detail' ? '' : 'none';
}

function enterReviewMode() {
  $('#grid-root').innerHTML = '';
  showScreen('list');
  renderReviewList();
}

function exitReviewMode() {
  $('#review-detail-root').innerHTML = '';
  GridRender.renderGrid($('#grid-root'), STATE, { showGear: false, reviewFieldIds: reviewFieldIdSet() });
  showScreen('normal');
}

// 詳細画面は「所管部署の入力を閲覧するだけ」ではなく、通常グリッド画面と同じく
// 全セル編集可能にする（全セル誰でも入力できる、という方針）。1件ずつ入力したい
// レビュアー向けの経路として、一覧表のインライン編集と並ぶ選択肢になる。
function openReviewDetail(record) {
  DETAIL_RECORD = record;
  DETAIL_STATE = buildStateFromStructure(STRUCTURE);
  GridRender.renderGrid($('#review-detail-root'), DETAIL_STATE, { showGear: false, reviewFieldIds: reviewFieldIdSet() });
  GridRender.loadDataIntoGrid(DETAIL_STATE, record.data || {});
  showScreen('detail');
}

// 詳細画面での編集を破棄して一覧に戻る（保存しない）。
function backToReviewListFromDetail() {
  $('#review-detail-root').innerHTML = '';
  DETAIL_RECORD = null;
  DETAIL_STATE = null;
  showScreen('list');
}

// 詳細画面での編集内容を、対象レコードのdata/displayValues/reviewValuesへ書き戻し、
// 一覧を再描画してから一覧画面に戻る。書き戻すのはレビュー欄だけでなく全項目
// （所管部署欄も含む）で、詳細画面で行った変更はすべて反映される。
function saveDetailAndBackToList() {
  if (!DETAIL_RECORD || !DETAIL_STATE) return;
  const record = DETAIL_RECORD;
  const merged = GridRender.collectData(DETAIL_STATE);
  record.data = merged;
  record.displayValues = extractFieldValues(merged, STRUCTURE.displayCandidateFields || []);
  record.reviewValues = extractFieldValues(merged, STRUCTURE.reviewFields || []);
  backToReviewListFromDetail();
  renderReviewList();
  $('#status').textContent = `「${record.fileName}」の内容を保存しました。`;
}

// Chromium限定：フォルダを1回指定→以後「更新」ボタンで都度読み直せるようにする。
// window.showDirectoryPickerが無いブラウザ（Firefox/Safari）では、対応するボタン自体を
// 表示しない（フォールバックは複数ファイル選択ダイアログの選び直しに委ねる）。
async function pickReviewDirectory() {
  if (!window.showDirectoryPicker) return;
  try {
    REVIEW.dirHandle = await window.showDirectoryPicker();
  } catch (e) {
    return; // ユーザーがキャンセルした場合等。エラー表示はしない。
  }
  $('#review-btn-refresh').style.display = '';
  await scanReviewDirectory();
}

async function scanReviewDirectory() {
  if (!REVIEW.dirHandle) return;
  const newRecords = [];
  for await (const entry of REVIEW.dirHandle.values()) {
    if (entry.kind !== 'file' || !entry.name.toLowerCase().endsWith('.json')) continue;
    const file = await entry.getFile();
    try {
      const text = await readFileAsText(file);
      newRecords.push({ fileName: entry.name, data: JSON.parse(text) });
    } catch (e) { /* 壊れたJSONはスキップ */ }
  }
  const added = upsertRecords(newRecords);
  renderReviewList();
  $('#status').textContent = `フォルダを読み直しました（新規${added}件、合計${REVIEW.records.length}件）。`;
}

async function handleReviewFileInput(fileList) {
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
  renderReviewList();
  $('#status').textContent = `${added}件読み込みました（合計${REVIEW.records.length}件）。`;
}

function init() {
  STATE = buildStateFromStructure(STRUCTURE);
  GridRender.renderGrid($('#grid-root'), STATE, { showGear: false, reviewFieldIds: reviewFieldIdSet() });
  REVIEW.visibleCols = new Set(STRUCTURE.displayCandidateFields || []);

  // レビュー欄が1つも指定されていない様式では、レビュー画面自体を提供しない。
  // 色分けの凡例も同じ条件で出し分ける（区別する意味が無いため）。
  if ((STRUCTURE.reviewFields || []).length > 0) {
    $('#btn-open-review').style.display = '';
    $('#field-legend').style.display = '';
  }
  if (window.showDirectoryPicker) {
    $('#review-btn-pick-dir').style.display = '';
    // #export-dir-barはCSS側で`display:none`をスタイルシート宣言（インライン属性ではない）
    // として持っているため、.style.display=''（インライン指定の除去）では
    // スタイルシートのdisplay:noneにフォールバックしてしまい表示されない。
    // インラインで明示的に'flex'を指定してスタイルシートの宣言を上書きする必要がある。
    $('#export-dir-bar').style.display = 'flex';
  }

  $('#btn-open-review').addEventListener('click', enterReviewMode);
  $('#btn-close-review').addEventListener('click', exitReviewMode);
  $('#btn-back-to-review-list').addEventListener('click', backToReviewListFromDetail);
  $('#btn-save-detail').addEventListener('click', saveDetailAndBackToList);
  $('#btn-print-grid').addEventListener('click', () => window.print());
  $('#btn-print-detail').addEventListener('click', () => window.print());
  $('#review-file-load').addEventListener('change', (ev) => {
    handleReviewFileInput(ev.target.files);
    ev.target.value = '';
  });
  $('#review-status-select').addEventListener('change', (ev) => {
    REVIEW.statusFilter = ev.target.value;
    renderReviewBody();
  });
  $('#review-btn-pick-dir').addEventListener('click', pickReviewDirectory);
  $('#review-btn-refresh').addEventListener('click', scanReviewDirectory);
  $('#btn-pick-export-dir').addEventListener('click', pickExportDirectory);

  $('#btn-export').addEventListener('click', () => {
    const missing = findMissingRequiredFields();
    if (missing.length > 0) {
      const labels = missing.map(labelForCellId).join('、');
      const proceed = window.confirm(`必須項目が未入力です（${labels}）。このまま書き出しますか？`);
      if (!proceed) {
        $('#status').textContent = '書き出しを中止しました。未入力の必須項目を確認してください。';
        return;
      }
    }
    const data = GridRender.collectData(STATE);
    const filename = buildExportFileName();
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    saveBlob(blob, filename, (result) => {
      $('#status').textContent = statusMessageForSave(result, filename, '書き出しました。');
    });
  });

  $('#file-load-json').addEventListener('change', (ev) => {
    const file = ev.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result);
        GridRender.loadDataIntoGrid(STATE, data);
        $('#status').textContent = '読み込みました: ' + file.name;
      } catch (e) {
        $('#status').textContent = 'JSONの読み込みに失敗しました。';
      }
    };
    reader.readAsText(file, 'utf-8');
  });

  GridRender.setupPasteHandler(STATE, $('#status'));
}

if (typeof window !== 'undefined') {
  window.addEventListener('DOMContentLoaded', init);
  window.__app = {
    get STATE() { return STATE; },
    collectData: () => GridRender.collectData(STATE),
    loadDataIntoGrid: (data) => GridRender.loadDataIntoGrid(STATE, data),
    buildExportFileName, sanitizeForFileName,
    findMissingRequiredFields, labelForCellId,
    get REVIEW() { return REVIEW; },
    upsertRecords, extractFieldValues, mergedDataForExport, isRecordComplete,
    filteredRecords, renderReviewBody,
    buildExportFileNameForRecord, renderReviewList, enterReviewMode, exitReviewMode,
    openReviewDetail, backToReviewListFromDetail, saveDetailAndBackToList, handleReviewFileInput,
    pickReviewDirectory, scanReviewDirectory,
    saveBlob, pickExportDirectory,
    get EXPORT_DIR_HANDLE() { return EXPORT_DIR_HANDLE; },
    setExportDirHandle: (h) => { EXPORT_DIR_HANDLE = h; }, // テスト用（実ブラウザではpickExportDirectory経由）
  };
}
