// ビルダー（excel_form_builder.html）・生成される入力フォームの両方で共有する
// 「グリッド描画・データ収集/復元・数式再計算・貼り付け」ロジック。
// core_logic.js（CoreLogic.cellId／colNumToLetter／buildSectionObject）に依存する。
//
// state の形： { grid, maxRow, maxCol, widths, heights, sections, fileName }
//   - grid: Map<'row,col', CellInfo>（結合解決済み。core_logic.jsのbuildGrid/applyOverrides適用後のもの）
//   - widths/heights: 各列・各行のピクセル寸法の配列（1始まりの列・行番号に合わせ、widths[c-1]でアクセス）
//   - sections: [{title, row0, row1}, ...]（JSON書き出し/読み込みに必要）

(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.GridRender = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {

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

// ---------------------------------------------------------------------------
// グリッド描画
// ---------------------------------------------------------------------------

function buildInputControl(info, rowspan, colspan) {
  const type = info.renderType || 'textarea';
  const id = CoreLogic.cellId(info);
  if (type === 'select') {
    const select = el('select', { id });
    select.appendChild(el('option', { value: '' }, [document.createTextNode('（未選択）')]));
    (info.renderOptions || []).forEach(opt => select.appendChild(el('option', { value: opt }, [document.createTextNode(opt)])));
    return select;
  }
  if (type === 'number') {
    const input = el('input', { id });
    input.setAttribute('type', 'number');
    input.classList.add('input-number');
    return input;
  }
  if (type === 'currency') {
    const input = el('input', { id });
    input.setAttribute('type', 'text');
    input.classList.add('input-currency');
    input.addEventListener('input', () => {
      const digits = input.value.replace(/[^\d]/g, '');
      input.value = digits ? Number(digits).toLocaleString('ja-JP') : '';
    });
    return input;
  }
  return el('textarea', { id });
}

/**
 * state（grid/maxRow/maxCol/widths/heights）をrootElに描画する。
 * opts.showGear=true のとき各セルに⚙アイコンを付け、クリックでopts.onGear(info)を呼ぶ
 * （手直し機能を持つビルダー専用。生成済み入力フォーム側はfalseにする）。
 * opts.selected（Set<'row,col'>）を渡すと該当セルに .cell-selected を付ける。
 * opts.readonly=true のとき、全ての入力欄（input/textarea/select）にdisabledを付ける
 * （レビュー画面で所管部署の入力内容を閲覧専用で表示する用途）。
 * 戻り値：このグリッドに含まれる数式セルのinfo配列（再計算対象として呼び出し元が保持する）。
 */
function renderGrid(rootEl, state, opts) {
  opts = opts || {};
  const { grid, maxRow, maxCol, widths, heights, sections, manualGroups } = state;
  const formulaCells = [];
  rootEl.innerHTML = '';

  // 手動グループ（A）に属する行番号の集合。セルの背景に薄い印を付け、
  // どの範囲がすでにグループ化済みかを一目で分かるようにする。
  // 自動グループ化を解除しただけ（disabled）の範囲は「グループ」ではないので含めない。
  const groupedRows = new Set();
  (manualGroups || []).forEach(g => {
    if (g.disabled) return;
    for (let r = g.row0; r <= g.row1; r++) groupedRows.add(r);
  });
  // 自動検出（縦結合）でグループとみなされる行番号の集合。手動グループとは別の色で
  // 印を付け、「ここは縦結合から自動でグループ化されている」と一目で分かるようにする
  // （成果指標の見出し行がグループのレコードに紛れ込むような不具合を、色で早期発見できるように）。
  const autoGroupedRows = (sections && sections.length)
    ? CoreLogic.computeAutoGroupRows(grid, sections, maxCol, manualGroups)
    : new Set();

  const table = el('table', { class: 'excel-grid' });
  const colgroup = el('colgroup');
  colgroup.appendChild(el('col', { style: 'width:34px' }));
  for (let c = 1; c <= maxCol; c++) {
    colgroup.appendChild(el('col', { style: `width:${(widths && widths[c - 1]) || 64}px` }));
  }
  table.appendChild(colgroup);

  const theadRow = el('tr', { class: 'ruler-row' });
  theadRow.appendChild(el('th', { class: 'ruler-corner' }));
  for (let c = 1; c <= maxCol; c++) {
    theadRow.appendChild(el('th', { class: 'ruler-col', text: CoreLogic.colNumToLetter(c) }));
  }
  const thead = el('thead');
  thead.appendChild(theadRow);
  table.appendChild(thead);

  const tbody = el('tbody');
  for (let r = 1; r <= maxRow; r++) {
    const tr = el('tr', { style: `height:${(heights && heights[r - 1]) || 22}px` });
    tr.appendChild(el('th', { class: 'ruler-row-num', text: String(r) }));
    for (let c = 1; c <= maxCol; c++) {
      const info = grid.get(r + ',' + c);
      if (!info || info.row !== r || info.col !== c) continue;
      const rowspan = info.row2 - info.row + 1;
      const colspan = info.col2 - info.col + 1;
      const cellKey = info.row + ',' + info.col;
      const td = el('td', {
        rowspan: String(rowspan), colspan: String(colspan),
        id: 'td_R' + info.row + '_C' + info.col,
        'data-row': String(info.row), 'data-col': String(info.col),
      });

      if (opts.showGear) {
        const gear = el('button', { class: 'gear-btn', title: 'このセルの種類を設定', type: 'button', text: '⚙' });
        // mousedownの段階で止めないと、複数選択モードのドラッグ選択（td側のmousedown）が
        // 先に発火してしまい、ギアボタンをクリックしただけで意図せずセルが選択状態になる。
        gear.addEventListener('mousedown', (ev) => ev.stopPropagation());
        gear.addEventListener('click', (ev) => { ev.stopPropagation(); opts.onGear && opts.onGear(info); });
        td.appendChild(gear);
      }

      if (info.isFormula) {
        td.className = 'cell-calc';
        const span = el('span', { class: 'calc-val', id: 'calcval_' + CoreLogic.cellId(info), text: '(計算中)' });
        td.appendChild(span);
        formulaCells.push(info);
      } else if (info.blocked) {
        td.className = 'cell-blocked';
      } else if (info.hasText) {
        td.className = 'cell-label';
        if (info.fillColor) td.style.background = '#' + info.fillColor;
        const span = el('span', { text: String(info.value) });
        if (opts.onGear) span.addEventListener('dblclick', () => opts.onGear(info));
        td.appendChild(span);
      } else {
        td.className = 'cell-input';
        const inputEl = buildInputControl(info, rowspan, colspan);
        // opts.readonly：レビュー画面の詳細確認（所管部署の入力を閲覧するだけ）用。
        // renderType別に別々の分岐を持たず、描画後に一律disabledを付けるだけにする
        // （buildInputControl側に手を入れると通常モードとの分岐が増えて事故りやすいため）。
        if (opts.readonly) inputEl.setAttribute('disabled', 'disabled');
        td.appendChild(inputEl);
      }
      if (opts.selected && opts.selected.has(cellKey)) td.classList.add('cell-selected');
      // 手動グループが実際の出力で優先されるので、見た目もそれに合わせる
      // （両方に該当する行は手動グループの色を優先し、自動検出の色は付けない）。
      if (groupedRows.has(info.row)) td.classList.add('cell-grouped');
      else if (autoGroupedRows.has(info.row)) td.classList.add('cell-auto-grouped');
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  rootEl.appendChild(table);

  state.formulaCells = formulaCells;
  recalcAllFormulas(state);
  return formulaCells;
}

// ---------------------------------------------------------------------------
// 数式（SUM）の簡易評価
// ---------------------------------------------------------------------------

function parseCellRef(ref) {
  const m = ref.match(/^([A-Z]+)(\d+)$/);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col, row: parseInt(m[2], 10) };
}

function parseSumFormula(formula) {
  const m = String(formula || '').match(/^=SUM\(([A-Z]+\d+):([A-Z]+\d+)\)$/i);
  if (!m) return null;
  const a = parseCellRef(m[1]), b = parseCellRef(m[2]);
  if (!a || !b) return null;
  return { r0: Math.min(a.row, b.row), r1: Math.max(a.row, b.row), c0: Math.min(a.col, b.col), c1: Math.max(a.col, b.col) };
}

function evalFormulaInfo(state, info) {
  const range = parseSumFormula(info.formula);
  if (!range) return null;
  const { grid } = state;
  let sum = 0;
  const seen = new Set();
  for (let r = range.r0; r <= range.r1; r++) {
    for (let c = range.c0; c <= range.c1; c++) {
      const cellInfo = grid.get(r + ',' + c);
      if (!cellInfo) continue;
      const key = cellInfo.row + ',' + cellInfo.col;
      if (seen.has(key)) continue;
      seen.add(key);
      if (cellInfo.isFormula) continue;
      const elx = document.getElementById(CoreLogic.cellId(cellInfo));
      const raw = elx ? elx.value : (cellInfo.hasText ? cellInfo.value : '');
      const num = parseFloat(String(raw).replace(/[,\s]/g, ''));
      if (!isNaN(num)) sum += num;
    }
  }
  return sum;
}

function recalcAllFormulas(state) {
  (state.formulaCells || []).forEach(info => {
    const span = document.getElementById('calcval_' + CoreLogic.cellId(info));
    if (!span) return;
    const val = evalFormulaInfo(state, info);
    span.textContent = val === null ? '（自動計算・数式未対応）' : val.toLocaleString('ja-JP');
  });
}

// ---------------------------------------------------------------------------
// 書き出し／読み込み（DOM値 ⇄ JSON）
// ---------------------------------------------------------------------------

function domGetValue(info) {
  const elx = document.getElementById(CoreLogic.cellId(info));
  return elx ? elx.value : '';
}

function collectData(state) {
  const { grid, sections, maxCol, fileName, manualGroups } = state;
  const out = { meta: { exported_at: new Date().toISOString(), source_file: fileName } };
  for (const sec of sections) {
    out[sec.title] = CoreLogic.buildSectionObject(grid, sec.row0, sec.row1, maxCol, domGetValue, manualGroups);
  }
  return out;
}

function loadDataIntoGrid(state, data) {
  const { grid, sections, maxCol, manualGroups } = state;
  for (const sec of sections) {
    const secData = data[sec.title];
    if (!secData) continue;
    fillSection(grid, sec.row0, sec.row1, maxCol, secData, manualGroups);
  }
  recalcAllFormulas(state);
}

function fillSection(grid, row0, row1, maxCol, secData, manualGroups) {
  let r = row0;
  const units = [];
  while (r <= row1) {
    const r2 = rowSpanOfLocal(grid, r, maxCol);
    units.push([r, r2]);
    r = r2 + 1;
  }
  let i = 0;
  while (i < units.length) {
    const [u0, u1] = units[i];

    // 手動グループ（A）：core_logic.js の buildSectionObject と対称の実装
    // （書き出し側と読込側で判定条件が食い違うと、値が消える／混線するバグの原因になる）。
    const manualGroup = CoreLogic.findManualGroup(manualGroups, u0);
    // 自動グループ化の解除：グループ化せず、このユニットだけ通常の単発行として扱う
    // （core_logic.js の disabled 早期リターンと対称）。
    if (manualGroup && manualGroup.disabled) {
      // labelOfRowではなくbuildRowEntryを使う理由はrenderSectionFromData側の同種の
      // 呼び出しと同じ（「1行1見出し1値」がdbKeyで改名されている場合に対応するため）。
      const label = CoreLogic.buildRowEntry(grid, u0, u1, maxCol, 1).label;
      if (label && secData[label] !== undefined) fillRow(grid, u0, maxCol, 1, secData[label]);
      i++;
      continue;
    }
    if (manualGroup) {
      const groupEndRow = manualGroup.row1;
      const key = manualGroup.name;
      const childUnits = [];
      let j = i;
      while (j < units.length && units[j][0] <= groupEndRow) { childUnits.push(units[j]); j++; }
      fillGroupChildren(grid, maxCol, childUnits, secData[key], 1, secData);
      i = j;
      continue;
    }

    // 自動検出（縦結合）：core_logic.js の buildSectionObject と対称の実装。
    const groupInfo = grid.get(u0 + ',1');
    if (CoreLogic.isAutoGroupStart(groupInfo, u0, u1)) {
      let groupEndRow = groupInfo.row2;
      const overlappingManual = (manualGroups || [])
        .filter(g => g.row0 > u0 && g.row0 <= groupEndRow)
        .sort((a, b) => a.row0 - b.row0)[0];
      if (overlappingManual) groupEndRow = overlappingManual.row0 - 1;
      const key = String(groupInfo.value).replace(/\s+/g, '');
      const childUnits = [];
      let j = i;
      while (j < units.length && units[j][0] <= groupEndRow) { childUnits.push(units[j]); j++; }
      fillGroupChildren(grid, maxCol, childUnits, secData[key], 2, secData, true);
      i = j;
      continue;
    }

    // labelOfRowではなくbuildRowEntryを使う：「1行1見出し1値」の行が
    // dbKeyで改名されている場合、出力キーはrowLabelではなくdbKeyになる
    // （core_logic.jsのbuildRowEntry参照）。単純にA列の見出し文字だけを見る
    // labelOfRowのままだと、改名後のキーでsecDataを探せず値が読み込めなくなる。
    const label = CoreLogic.buildRowEntry(grid, u0, u1, maxCol, 1).label;
    if (label && secData[label] !== undefined) {
      fillRow(grid, u0, maxCol, 1, secData[label]);
    }
    i++;
  }
}

function rowSpanOfLocal(grid, r, maxCol) {
  let row2 = r;
  for (let c = 2; c <= maxCol; c++) {
    const info = grid.get(r + ',' + c);
    if (info && info.row === r) row2 = Math.max(row2, info.row2);
  }
  return row2;
}

// グループ（手動／自動検出どちらも）の子ユニット群に、書き出し済みのgroupValを流し込む。
// core_logic.jsのbuildSectionObject／groupChildrenToResultと対称の実装：
// 書き出し側は実データを持たない子（見出し行の残骸。CoreLogic.buildRowEntryの
// hasRealData=false）をレコードから除外しているため、読込側も同じ
// CoreLogic.buildRowEntry（label・hasRealDataの判定そのもの）を使って同じ子だけを
// 対象にする。ラベルの有無だけを見る簡易版（旧labelOfRow）では、実データを持たない
// 「ラベルなしの空行」まで拾ってしまい配列のインデックスがずれる場合があった。
// secDataはセクション全体のデータ（groupValの親）。無名の繰り返し行と「その他」の
// ような名前付き行が混在するグループでは、書き出し側は名前付きの行をグループの外側
// （secData直下、通常の単発行と同じ扱い）に出力しているため、読込側もそちらを見る。
// requireMultiple：core_logic.jsのgroupChildrenToResultと同じ意味・同じ呼び出し条件
// （自動検出グループの呼び出し元だけtrueを渡す）。子が1件だけなら辞書化せず、
// secData直下の独立キーとして復元する（既知バグの修正・読み込み側の対称実装）。
function fillGroupChildren(grid, maxCol, childUnits, groupVal, fromCol, secData, requireMultiple) {
  const entries = childUnits.map(([cu0, cu1]) => [cu0, CoreLogic.buildRowEntry(grid, cu0, cu1, maxCol, fromCol)]);
  const real = entries.filter(([, e]) => e.hasRealData);
  const minCount = requireMultiple ? 1 : 0;
  const allNamed = real.length > minCount && real.every(([, e]) => e.label);
  if (allNamed) {
    if (!groupVal) return;
    real.forEach(([cu0, e]) => {
      if (groupVal[e.label] !== undefined) fillRow(grid, cu0, maxCol, fromCol, groupVal[e.label]);
    });
    return;
  }
  if (Array.isArray(groupVal)) {
    const unlabeled = real.filter(([, e]) => !e.label);
    unlabeled.forEach(([cu0], idx) => fillRow(grid, cu0, maxCol, fromCol, groupVal[idx]));
  }
  const labeled = real.filter(([, e]) => e.label);
  labeled.forEach(([cu0, e]) => {
    if (secData && secData[e.label] !== undefined) fillRow(grid, cu0, maxCol, fromCol, secData[e.label]);
  });
}

function fillRow(grid, r0, maxCol, fromCol, rowVal) {
  if (rowVal === undefined || rowVal === null) return;
  const cells = [];
  const seen = new Set();
  for (let c = fromCol; c <= maxCol; c++) {
    const info = grid.get(r0 + ',' + c);
    if (!info || info.row !== r0 || info.isFormula) continue;
    const key = info.row + ',' + info.col;
    if (seen.has(key)) continue;
    seen.add(key);
    cells.push(info);
  }
  if (cells.length === 0) return;
  const first = cells[0];
  const valueCells = first.hasText ? cells.slice(1) : cells;
  if (typeof rowVal === 'object' && !Array.isArray(rowVal)) {
    // core_logic.js の buildRowEntry と対称のキー導出（見出しセルはスキップし、
    // 直前の見出し文字を既定の項目名として使う）。ずれると読込時に値が拾えなくなる。
    let precedingLabel = null;
    valueCells.forEach(c => {
      if (c.hasText) { precedingLabel = String(c.value).replace(/\s+/g, ''); return; }
      const v = rowVal[CoreLogic.defaultKeyFor(c, precedingLabel)];
      if (v !== undefined) setDomValue(c, v);
    });
  } else if (valueCells.length === 1) {
    setDomValue(valueCells[0], rowVal);
  }
}

function setDomValue(info, v) {
  const elx = document.getElementById(CoreLogic.cellId(info));
  if (elx) elx.value = v === null || v === undefined ? '' : v;
}

// ---------------------------------------------------------------------------
// 貼り付け対応（簡易版：タブ区切りテキストを、選択中セルを起点に貼り付ける）
// ---------------------------------------------------------------------------

function setupPasteHandler(state, statusEl) {
  document.addEventListener('paste', (ev) => {
    const active = document.activeElement;
    if (!active || !active.id || !active.id.startsWith('cell_')) return;
    const text = (ev.clipboardData || window.clipboardData).getData('text');
    if (!text || (text.indexOf('\t') === -1 && text.indexOf('\n') === -1)) return;
    ev.preventDefault();
    const m = active.id.match(/^cell_R(\d+)_C(\d+)$/);
    if (!m) return;
    const startRow = parseInt(m[1], 10), startCol = parseInt(m[2], 10);
    const rows = text.replace(/\r/g, '').split('\n').filter((_, i, arr) => !(i === arr.length - 1 && arr[i] === ''));
    rows.forEach((line, ri) => {
      const cols = line.split('\t');
      cols.forEach((val, ci) => {
        const target = document.getElementById(`cell_R${startRow + ri}_C${startCol + ci}`);
        if (target) target.value = val;
      });
    });
    recalcAllFormulas(state);
    if (statusEl) statusEl.textContent = `貼り付けました（${rows.length}行）。空欄でないセルの一部が対応する入力欄に反映されていない場合があります。結果をご確認ください。`;
  });
  document.addEventListener('input', (ev) => {
    if (ev.target && ev.target.id && ev.target.id.startsWith('cell_')) recalcAllFormulas(state);
  });
}

return {
  buildInputControl, renderGrid, collectData, loadDataIntoGrid, recalcAllFormulas,
  setupPasteHandler,
};
});
