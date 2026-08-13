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
 * opts.reviewFieldIds（Set<cellId>）を渡すと、入力セルごとに「1次入力欄」
 * （cell-field-primary）／「2次以降入力欄」（cell-field-review）のクラスを付ける
 * （STRUCTURE.reviewFieldsが1件も無い様式では区別する意味が無いため、Setが空なら何も付けない）。
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
        if (opts.reviewFieldIds && opts.reviewFieldIds.size > 0) {
          td.classList.add(opts.reviewFieldIds.has(CoreLogic.cellId(info)) ? 'cell-field-review' : 'cell-field-primary');
        }
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
// 印刷専用スナップショット：renderGridが作る編集用グリッド（ルーラー行・列付き、
// 値は<input>/<textarea>/<select>）とは別に、印刷にだけ使う読み取り専用テーブルを
// 組み立てる。ルーラーは含めず、値は全てテキストとして描画する（textareaの既定サイズに
// 引っぱられて行が間延びする問題を避けるため。実機確認で「1事業が9ページに分割される」
// という不具合として発覚した）。
// 呼び出し時点でstateがどこかのrootElへ既にrenderGrid()済みであること（cellIdから
// document.getElementByIdで値を引けること）が前提。domGetValue／evalFormulaInfoと同じ制約。
// <tr>は自前セルを持たない行でも必ず生成する（renderGridと同じ理由。省略すると
// rowspanが後続の無関係な行にはみ出す表示バグになる。上のregressionテスト参照）。
// ---------------------------------------------------------------------------
function buildPrintTable(state) {
  const { grid, maxRow, maxCol, widths } = state;
  const table = el('table', { class: 'print-grid' });
  const colgroup = el('colgroup');
  for (let c = 1; c <= maxCol; c++) {
    colgroup.appendChild(el('col', { style: `width:${(widths && widths[c - 1]) || 64}px` }));
  }
  table.appendChild(colgroup);

  const tbody = el('tbody');
  for (let r = 1; r <= maxRow; r++) {
    const tr = el('tr');
    for (let c = 1; c <= maxCol; c++) {
      const info = grid.get(r + ',' + c);
      if (!info || info.row !== r || info.col !== c) continue;
      const rowspan = info.row2 - info.row + 1;
      const colspan = info.col2 - info.col + 1;
      const td = el('td', { rowspan: String(rowspan), colspan: String(colspan) });
      if (info.isFormula) {
        td.className = 'print-cell-calc';
        const val = evalFormulaInfo(state, info);
        td.textContent = val === null ? '' : val.toLocaleString('ja-JP');
      } else if (info.blocked) {
        td.className = 'print-cell-blocked';
      } else if (info.hasText) {
        td.className = 'print-cell-label';
        if (info.fillColor) td.style.background = '#' + info.fillColor;
        td.textContent = String(info.value);
      } else {
        td.className = 'print-cell-input';
        const inputEl = document.getElementById(CoreLogic.cellId(info));
        td.textContent = inputEl ? inputEl.value : '';
      }
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  return table;
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
      fillTopLevelEntry(grid, u0, u1, maxCol, 1, secData);
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

    fillTopLevelEntry(grid, u0, u1, maxCol, 1, secData);
    i++;
  }
}

// トップレベルの単発行（グループに属さない行）をsecDataから読み込む。
// core_logic.jsのassignEntryと対称の実装：
// - hasRealData=falseの行（ラベル同士が隣接する見出しの残骸で、明示的な手直しもない行）
//   は書き出し側がそもそも出力していないため、読み込んでも復元すべき値が無くスキップする。
// - 複数値を持つ行は書き出し側で「文脈_項目名」の形にフラット化されているため、
//   nested dictとしてではなくsecData全体＋キー接頭辞（行見出し）としてfillRowへ渡す。
// - 「1行1見出し1値」の行がdbKeyで改名されている場合、出力キーはrowLabelではなく
//   dbKeyになる（core_logic.jsのbuildRowEntry参照）。単純にA列の見出し文字だけを見る
//   簡易版だと、改名後のキーでsecDataを探せず値が読み込めなくなる。
function fillTopLevelEntry(grid, u0, u1, maxCol, fromCol, secData) {
  const entry = CoreLogic.buildRowEntry(grid, u0, u1, maxCol, fromCol);
  if (!entry.label || !entry.hasRealData) return;
  if (entry.value && typeof entry.value === 'object' && !Array.isArray(entry.value)) {
    fillRow(grid, u0, maxCol, fromCol, secData, entry.label);
  } else if (secData[entry.label] !== undefined) {
    fillRow(grid, u0, maxCol, fromCol, secData[entry.label]);
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

// keyPrefix（省略可）：core_logic.jsのassignEntryがトップレベルの複数値行を
// 「文脈_項目名」にフラット化しているのを読み戻すための接頭辞。渡された場合、
// rowValは行専用のnested dictではなくsecData全体（フラットなまま）として扱い、
// 各セルは`${keyPrefix}_${ローカルキー}`で引く。グループの子レコード（元々nested dict
// のまま、フラット化していない）を読むfillGroupChildren経由の呼び出しでは渡さない。
function fillRow(grid, r0, maxCol, fromCol, rowVal, keyPrefix) {
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
      const localKey = CoreLogic.defaultKeyFor(c, precedingLabel);
      const fullKey = keyPrefix ? (keyPrefix + '_' + localKey) : localKey;
      const v = rowVal[fullKey];
      if (v !== undefined) setDomValue(c, v);
      precedingLabel = null;
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

// Excelのクリップボード形式（TSV）は、セル内に改行・タブ・ダブルクォート自体を含む場合、
// CSVと同じ規則でセル全体をダブルクォートで囲み、内部のダブルクォートは2つに重ねる
// （例：縦結合の複数行ラベル「R6年度\n(2024年度)」はコピー時に`"R6年度\n(2024年度)"`と
// 引用符付きで表現される）。素朴にsplit('\n')/split('\t')するだけだと、この引用符内の
// 改行を行の区切りと誤認し、そこから後ろの行が全て1行ずつずれて貼り付けられる不具合が
// 実機確認（複数行ラベルを含む様式）で発覚したため、RFC4180準拠のクォート規則を
// 踏まえて1文字ずつパースする。
function parseClipboardGrid(text) {
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  const s = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (inQuotes) {
      if (ch === '"') {
        if (s[i + 1] === '"') { field += '"'; i++; }
        else { inQuotes = false; }
      } else {
        field += ch;
      }
    } else if (ch === '"' && field === '') {
      inQuotes = true;
    } else if (ch === '\t') {
      row.push(field); field = '';
    } else if (ch === '\n') {
      row.push(field); field = '';
      rows.push(row); row = [];
    } else {
      field += ch;
    }
  }
  row.push(field);
  rows.push(row);
  // 貼り付けテキスト末尾の改行に由来する空行（1列だけの空文字列の行）は除く
  if (rows.length > 0 && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') {
    rows.pop();
  }
  return rows;
}

// (startRow, startCol)を起点に、パース済みのタブ区切りグリッドを流し込む共通処理。
// 対応する入力欄が無い座標（見出し・自動計算・結合セルの2マス目等）は
// document.getElementByIdがnullを返すため、何もせず静かにスキップされる
// （Excelの見出し行を含む範囲をそのまま貼り付けても壊れない、という前提の要）。
function applyPastedGrid(state, text, startRow, startCol) {
  const rows = parseClipboardGrid(text);
  rows.forEach((cols, ri) => {
    cols.forEach((val, ci) => {
      const target = document.getElementById(`cell_R${startRow + ri}_C${startCol + ci}`);
      if (target) target.value = val;
    });
  });
  recalcAllFormulas(state);
  return rows.length;
}

function setupPasteHandler(state, statusEl) {
  document.addEventListener('paste', (ev) => {
    const active = document.activeElement;
    if (!active || !active.id || !active.id.startsWith('cell_')) return;
    const text = (ev.clipboardData || window.clipboardData).getData('text');
    if (!text || (text.indexOf('\t') === -1 && text.indexOf('\n') === -1)) return;
    ev.preventDefault();
    const m = active.id.match(/^cell_R(\d+)_C(\d+)$/);
    if (!m) return;
    const rowCount = applyPastedGrid(state, text, parseInt(m[1], 10), parseInt(m[2], 10));
    if (statusEl) statusEl.textContent = `貼り付けました（${rowCount}行）。空欄でないセルの一部が対応する入力欄に反映されていない場合があります。結果をご確認ください。`;
  });
  document.addEventListener('input', (ev) => {
    if (ev.target && ev.target.id && ev.target.id.startsWith('cell_')) recalcAllFormulas(state);
  });
}

// 「まとめて貼り付け」ボタン用：クリック中のセルに依存せず、クリップボードの内容を
// グリッド左上（行1・列1＝Excelの実際の左上セルA1と対応）を基準に反映する。
// setupPasteHandler（クリックしたセルを起点にする方式）と違い、様式の左上が見出し等
// クリックできないセルであっても、Excel全体（見出し行含む）をそのまま貼り付けられる。
// navigator.clipboard.readTextはセキュアコンテキスト必須（Chromiumはfile://も対象に含む）。
// 呼び出し側（filler_app.js）でAPIの有無を確認し、無ければボタン自体を出さない設計とする。
async function pasteFromClipboardAtOrigin(state) {
  const text = await navigator.clipboard.readText();
  if (!text) throw new Error('clipboard-empty');
  return applyPastedGrid(state, text, 1, 1);
}

return {
  buildInputControl, renderGrid, buildPrintTable, collectData, loadDataIntoGrid, recalcAllFormulas,
  setupPasteHandler, applyPastedGrid, pasteFromClipboardAtOrigin, parseClipboardGrid,
};
});
