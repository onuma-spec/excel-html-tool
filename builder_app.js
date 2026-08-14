// ツール1（ビルダー）側のロジック。Excel読み込み・手直し・入力フォームのHTML書き出しを行う。
// データ入力そのものは行わない（それはツール2＝書き出された入力フォームの役割）。
// core_logic.js（CoreLogic）・grid_render.js（GridRender）・SheetJS（XLSX）に依存する。

let CURRENT = null; // { ws, grid, maxRow, maxCol, widths, heights, sections, fileName, manualGroups }
let OVERRIDES = {};  // cellId -> {kind,label,options,formula,dbKey}（このセッション内のみ保持）

// 書き出し先フォルダ（Chromium限定・File System Access API）。指定されていれば
// doExportAsForm()が通常のダウンロードダイアログを介さず直接このフォルダへ書き込む。
// filler_app.js側の同名機構と考え方は同じだが、ファイルは別々（ビルダー・入力フォームは
// 別のHTMLとして動くため）。
let EXPORT_DIR_HANDLE = null;

// ツール2（生成される入力フォーム）のHTML全文。ビルド時（assemble.py）に埋め込まれる。
// 「__STRUCTURE__」「__FORM_TITLE__」の2箇所だけが未確定のまま残っており、
// doExportAsForm()がそれぞれのフォーム固有の内容に置き換えてから書き出す。
const FILLER_TEMPLATE = "__FILLER_TEMPLATE_JSON__";

// ツール3（複数件のJSONを集約し、閲覧ページを書き出す集約ツール）のHTML全文。
// ビルド時（assemble.py）に埋め込まれる。「__STRUCTURE__」の1箇所だけが未確定のまま
// 残っており（この集約ツール自身が閲覧ページのVIEWER_TEMPLATEを内包しているため、
// ビューアー側のプレースホルダーはこの時点では触らない）、doExportAsAggregator()が
// 置き換えてから書き出す。
const AGGREGATOR_TEMPLATE = "__AGGREGATOR_TEMPLATE_JSON__";

// --- 複数選択（一括手直し）関連の状態。ツール1は常時このモード。 ---
let SELECTED = new Set(); // 'row,col'（アンカー座標）の集合
let DRAG_START_SNAPSHOT = null;
let DRAG_ANCHOR = null;
let DRAGGING = false;

function $(sel) { return document.querySelector(sel); }

// #statusへのメッセージ表示を一本化する。#statusはSTEP0直後の固定位置にあり、
// STEP4（書き出し）はページの下の方にあるため、ボタンから離れた場所での
// テキスト変化だけでは気づかれにくい（実機確認フィードバックで発覚）。
// ①class付け替えで背景色フラッシュのアニメーションを再生させ、②scrollIntoView()で
// #status自体を画面内に入れることで、どのSTEPからでも気づけるようにする。
// scrollIntoViewはjsdomに実装されていないため、存在チェックしてから呼ぶ。
function setStatus(text) {
  const el = $('#status');
  el.textContent = text;
  if (!text) return;
  el.classList.remove('status-flash');
  void el.offsetWidth; // 同じclassを続けて付け直しても再アニメーションするように強制リフロー
  el.classList.add('status-flash');
  if (typeof el.scrollIntoView === 'function') {
    el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
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

// ---------------------------------------------------------------------------
// STEPガイド（唯一の情報源）：作業場所に置くSTEPn見出しの文言と、その場で開ける
// 補足説明（長いものだけ）を、ここ1箇所だけで管理する。以前は「使い方ガイド」という
// 独立した折りたたみ1箇所と、作業場所のSTEPn見出し3箇所とに同じ内容を別々に手書きしており、
// 片方だけ直して他方を直し忘れる二重管理になっていた。見出しと説明を離れた場所に
// 重複させるのをやめ、各STEP見出しの直下にその場で表示する構成にした
// （見出しと説明が物理的に同じ場所にあるので、構造的に二重管理になりようがない）。
// label：作業場所の小さな.step-label見出し。全STEPが持つ。
// detailHtml：長い補足だけ持つ（今のところSTEP2のみ）。[data-step-detail]のある場所に
// 「詳しく見る」の折りたたみとして展開する（固定の開発者記述内容のみ・ユーザー入力は含まない）。
// ---------------------------------------------------------------------------
const STEP_GUIDE = [
  { label: 'STEP1：読み込み' },
  {
    label: 'STEP2：セルの設定を行う（下のフォーム（作成中））',
    detailHtml: `
      <ul>
        <li><strong>見出し・自動計算</strong>：見出し文（または合計行等の数式）が適切か確認する。⚙アイコンから「見出し」「数式」を選べる。見出しを空欄にすると、入力させたくない余白セルにできる。</li>
        <li><strong>入力欄</strong>：入力しやすい形式（テキスト／数値／金額／プルダウン）になっているか確認する。項目名（データベース上の列名）は、この後のSTEP3「データ構造確認」でまとめて設定する。</li>
        <li>同じ形式の行が繰り返される表（毎年の活動記録、品目ごとの明細など）は、グループとしてまとめられる。行範囲をドラッグまたはCtrl+クリックで選択→画面下部の「🗂 グループ化する」を押す→分かりやすいグループ名を入力して保存する。グループ化すると、各行がそれぞれ独立した記録として書き出されるようになる（グループ化しないと、2行目以降の内容が正しく取り出せないことがある）。</li>
        <li>なお、Excelのタテ結合がある範囲は、読み込み時に自動でグループ化されていることがあります（青緑の縦線）。意図と違う場合（見出し行を巻き込んでいる等）は、同じ範囲を選び直して「自動グループ化を解除する」から個別に無効化できます。</li>
      </ul>
    `,
  },
  { label: 'STEP3：データ構造確認で項目名を設定する' },
  { label: 'STEP4：項目の運用設定' },
  { label: 'STEP5：書き出す' },
];

// 作業場所の[data-step-label]要素にSTEP_GUIDEのlabelを流し込み、
// [data-step-detail]要素に（あれば）詳細折りたたみを組み立てる。init()の最初に一度呼ぶ。
function renderStepUi() {
  document.querySelectorAll('[data-step-label]').forEach((elx) => {
    const idx = Number(elx.dataset.stepLabel);
    if (STEP_GUIDE[idx]) elx.textContent = STEP_GUIDE[idx].label;
  });
  document.querySelectorAll('[data-step-detail]').forEach((container) => {
    const idx = Number(container.dataset.stepDetail);
    const step = STEP_GUIDE[idx];
    if (!step || !step.detailHtml) return;
    container.innerHTML = '';
    const details = el('details', { class: 'step-detail' });
    details.appendChild(el('summary', { text: '詳しく見る' }));
    details.appendChild(el('div', { html: step.detailHtml }));
    container.appendChild(details);
  });
}

// ---------------------------------------------------------------------------
// ファイル読み込み・列幅/行高の抽出
// ---------------------------------------------------------------------------

function extractWidths(ws, maxCol) {
  const cols = ws['!cols'];
  const out = [];
  for (let c = 1; c <= maxCol; c++) {
    const col = cols && cols[c - 1];
    out.push(col ? (col.wpx || (col.wch ? Math.round(col.wch * 7 + 5) : 64)) : 64);
  }
  return out;
}
function extractHeights(ws, maxRow) {
  const rows = ws['!rows'];
  const out = [];
  for (let r = 1; r <= maxRow; r++) {
    const row = rows && rows[r - 1];
    out.push(row ? (row.hpx || (row.hpt ? Math.round(row.hpt * 96 / 72) : 22)) : 22);
  }
  return out;
}

// 未到達セルが属する行が、自分自身の見出し（1列目のテキスト）を持っているかどうか。
// 見出しを持つ行は通常そのままJSON出力されるはずなので、それでも未到達ということは
// 他の行・同じ行の中の別の箇所と見出しの文字が重複していて上書きされている可能性が
// 高い（グループ化しても直らない）。renderCheckPanelが警告文の出し分けに使う。
// ※ 縦結合グループの内側にある行はanchor.row!==c.rowになるため対象外（見出しの有無に
//   関わらず「未グループ化の繰り返し行」として扱う）。
function rowHasOwnLabel(grid, c) {
  const anchor = grid.get(c.row + ',1');
  return !!(anchor && anchor.hasText && anchor.row === c.row);
}

// 書き出し前チェックで「入力しても消える」と判定されたセル（グループ化で直せるもの・
// セクション外で直せないもの・行見出しの重複で直せないものすべて）を、既定で
// 「見出し（文字は空欄）」＝入力不可の装飾セル扱いにする。実データが入るべき箇所は、
// 読み込んだ人が後から手直しでグループ化・種類変更・Excel側の見出し修正等をする前提
// （多くのテンプレートは無関係な空欄セルの方が多いため、「全部フラグを立てて選んで
// もらう」より「まず全部除外し、必要な箇所だけ復帰させる」方が手直しの手間が少ないと
// いう判断）。既に手直し済み（OVERRIDES設定済み）のセルは上書きしない
// （フォームHTML再読込時に既存の設定を壊さないため）。
// ※ 当初は行見出しが重複しているセル（rowHasOwnLabel）だけ自動ブロックの対象外に
//   していたが、実データ（複雑な集計・チェック欄が多い調査票）で試したところ、
//   行内に同じ見出し文字（例：チェック済みを表す「1」）が何十箇所も繰り返し出現する
//   様式では数千件の警告がブロックされずに残ってしまい、かえって使いづらくなることが
//   判明したため撤回した。行見出しの重複は「グループ化とは別の理由」であることを
//   renderCheckPanelの警告文で区別するに留め、自動ブロック自体は他のセルと同様に扱う。
function autoBlockUnreachableCells() {
  const { unreachable } = CoreLogic.findUnreachableCells(CURRENT.grid, CURRENT.sections, CURRENT.maxCol, CURRENT.manualGroups);
  let count = 0;
  unreachable.forEach(info => {
    const id = CoreLogic.cellId(info);
    if (OVERRIDES[id]) return;
    OVERRIDES[id] = { kind: 'label', label: '', dbKey: null };
    count++;
  });
  return count;
}

function handleFile(file) {
  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const data = new Uint8Array(ev.target.result);
      const wb = XLSX.read(data, { type: 'array', cellStyles: true });
      const wsName = wb.SheetNames[0];
      const ws = wb.Sheets[wsName];
      // 新しいファイルを読み込むたびに手直し内容をリセットする。
      // セル位置（行・列番号）だけで紐付いているため、リセットしないと
      // 別ファイル（別年度・別様式）を読み込んだ際、位置が一致する無関係な
      // セルに前のファイルの手直しが誤って適用されてしまう。
      OVERRIDES = {};
      SELECTED.clear();
      CURRENT = { mode: 'xlsx', ws, fileName: file.name, manualGroups: [] };
      rebuildAndRender();
      const blockedCount = autoBlockUnreachableCells();
      if (blockedCount > 0) rebuildAndRender();
      setStatus(`読み込みました: ${file.name}（${wsName}）`
        + (blockedCount > 0 ? ` ／ 書き出し前チェックに引っかかった${blockedCount}個のセルを自動的に「見出し（空白）」に設定しました。実際にデータが入る箇所は手直ししてください。` : ''));
    } catch (e) {
      console.error(e);
      setStatus('このファイルは読み込めませんでした。.xlsx形式かご確認ください。');
    }
  };
  reader.readAsArrayBuffer(file);
}

// ビルダーが書き出した入力フォームHTMLから STRUCTURE = {...}; の中身を取り出す。
// JSON.stringifyされたテキストなので、文字列リテラル内の { } はスキップしつつ
// 波かっこの深さで対応する終端位置を探す（値の中に ; や } が含まれても崩れないように）。
function extractStructureFromFillerHtml(text) {
  const marker = 'const STRUCTURE = ';
  const start = text.indexOf(marker);
  if (start === -1) return null;
  const braceStart = text.indexOf('{', start);
  if (braceStart === -1) return null;
  let depth = 0, inStr = false, esc = false, end = -1;
  for (let i = braceStart; i < text.length; i++) {
    const ch = text[i];
    if (inStr) {
      if (esc) esc = false;
      else if (ch === '\\') esc = true;
      else if (ch === '"') inStr = false;
      continue;
    }
    if (ch === '"') { inStr = true; continue; }
    if (ch === '{') depth++;
    else if (ch === '}') {
      depth--;
      if (depth === 0) { end = i + 1; break; }
    }
  }
  if (end === -1) return null;
  return JSON.parse(text.slice(braceStart, end));
}

// 既に書き出し済みの入力フォームHTML（今年度配布したフォーム等）を読み込み、
// それを土台にさらに手直しできるようにする（翌年度用フォームの作成を想定）。
// 手直し結果（プルダウン・数式・見出し変更等）は書き出し時点でセルにすでに焼き込まれて
// いるため、.xlsx読み込み時のようなOVERRIDES再適用は不要で、そのまま編集を続けられる。
function handleFormHtmlFile(file) {
  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const structure = extractStructureFromFillerHtml(ev.target.result);
      if (!structure || !Array.isArray(structure.cells)) throw new Error('構造データが見つかりません');
      OVERRIDES = {};
      SELECTED.clear();
      CURRENT = {
        mode: 'html',
        fileName: structure.formTitle || file.name.replace(/\.html?$/i, ''),
        baseCells: structure.cells,
        maxRow: structure.maxRow, maxCol: structure.maxCol,
        widths: structure.widths, heights: structure.heights,
        sections: structure.sections,
        // 手動グループ（A）は書き出し時にすでに確定済みの構造として焼き込まれているため、
        // OVERRIDESとは違い、そのまま引き継いで続きの編集ができるようにする。
        manualGroups: structure.manualGroups || [],
        // ファイル名に含める項目の選択も同様に、前回の書き出し内容を引き継いで続きから編集できるようにする。
        fileNameFieldIds: new Set(structure.fileNameFields || []),
        // 必須項目・2次入力欄・2次入力画面表示項目も同様に引き継ぐ。displayCandidateFieldsは
        // ファイル名項目も含んだ状態で書き出されているが、serializeStructure側で
        // fileNameFieldIdsとの和集合を毎回取り直すため、ここではそのまま復元してよい
        // （ファイル名項目分が重複してSetに入っても実害はない）。
        requiredFieldIds: new Set(structure.requiredFields || []),
        secondaryFieldIds: new Set(structure.secondaryFields || []),
        displayCandidateFieldIds: new Set(structure.displayCandidateFields || []),
      };
      rebuildAndRender();
      const blockedCount = autoBlockUnreachableCells();
      if (blockedCount > 0) rebuildAndRender();
      setStatus(`入力フォームを読み込みました: ${file.name}（このフォームをもとに手直しできます）`
        + (blockedCount > 0 ? ` ／ 書き出し前チェックに引っかかった${blockedCount}個のセルを自動的に「見出し（空白）」に設定しました。実際にデータが入る箇所は手直ししてください。` : ''));
    } catch (e) {
      console.error(e);
      setStatus('この入力フォームHTMLは読み込めませんでした。ビルダーで書き出したファイルかご確認ください。');
    }
  };
  reader.readAsText(file, 'utf-8');
}

function rebuildAndRender() {
  let grid;
  if (CURRENT.mode === 'html') {
    // 構造（行・列数、幅・高さ、セクション）は読み込み時点のものをそのまま維持する。
    grid = CoreLogic.buildGridFromCells(CURRENT.baseCells, CURRENT.maxRow, CURRENT.maxCol).grid;
    CoreLogic.applyOverrides(grid, OVERRIDES, CoreLogic.cellId);
  } else {
    const built = CoreLogic.buildGrid(CURRENT.ws);
    grid = built.grid;
    CoreLogic.applyOverrides(grid, OVERRIDES, CoreLogic.cellId);
    CURRENT.maxRow = built.maxRow;
    CURRENT.maxCol = built.maxCol;
    CURRENT.widths = extractWidths(CURRENT.ws, built.maxCol);
    CURRENT.heights = extractHeights(CURRENT.ws, built.maxRow);
    CURRENT.sections = CoreLogic.splitSections(grid, built.maxRow, built.maxCol);
  }
  CURRENT.grid = grid;
  GridRender.renderGrid($('#grid-root'), CURRENT, { showGear: true, onGear: handleGearClick, selected: SELECTED });
  // グリッドの状態が変わるたびに、常時表示中の書き出し前チェックパネルも最新化する
  // （手直し・グループ化のたびに都度「HTMLで書き出す」を押さなくても、常に今の状態を確認できるようにするため）。
  renderCheckPanel();
}

// 2個以上選択中に⚙を押した場合は、押したセルに関わらず選択中の全セルをまとめて設定する
// （「まとめて設定」ボタンと同じ挙動のショートカット）。1個以下ならそのセル単体を設定する。
function handleGearClick(info) {
  if (SELECTED.size >= 2) {
    openCellSettings(selectedInfos());
  } else {
    openCellSettings(info);
  }
}

// ---------------------------------------------------------------------------
// セル設定モーダル（見出し⇔入力欄の切替・型指定・プルダウン・数式編集を1つに統合）
// ---------------------------------------------------------------------------

const TYPE_LABELS = {
  label: '見出し',
  textarea: 'テキスト',
  number: '数値',
  currency: '金額（3桁区切り）',
  select: 'プルダウン',
  formula: '数式',
};

function currentKindOf(info) {
  if (info.isFormula) return 'formula';
  if (info.hasText || info.blocked) return 'label';
  return info.renderType || 'textarea';
}

// セル設定は「①セルの種類を選ぶ→②種類に応じた項目を（あれば）埋める」という
// ウィザード形式にする（一度に全項目を出す1枚のフォームより、非技術者には分かりやすいため）。
//   見出し     → 見出し文字 → 終わり
//   テキスト/数値/金額 → （追加入力なし）即終わり
//   プルダウン  → 選択肢 → 終わり
//   数式       → 数式の設定 → 終わり
// 項目名（dbKey）はこのウィザードでは扱わない。STEP3「データ構造確認」に一本化した
// （グループ化された列は複数セルにまたがるため、1セルずつ開くこのモーダルより、
// 構造全体を見ながら列単位でまとめて設定できるSTEP3の方が実態に合っているため）。
function openCellSettings(infoOrList) {
  closeModal();
  const infos = Array.isArray(infoOrList) ? infoOrList : [infoOrList];
  const primary = infos[0];
  const cellIds = infos.map(i => CoreLogic.cellId(i));
  const multi = infos.length > 1;

  const draft = {
    kind: currentKindOf(primary),
    label: primary.hasText ? String(primary.value) : '',
    options: (primary.renderOptions || []).join('\n'),
    formula: primary.formula || '=SUM()',
  };

  const overlay = el('div', { class: 'modal-overlay', id: 'modal-overlay' });
  const box = el('div', { class: 'modal-box' });
  overlay.appendChild(box);
  overlay.addEventListener('click', (ev) => { if (ev.target === overlay) closeModal(); });
  document.body.appendChild(overlay);

  let step = 'kind';

  function finish() {
    const override = { kind: draft.kind };
    if (draft.kind === 'label') override.label = draft.label;
    if (draft.kind === 'select') override.options = draft.options.split(/[\n,]+/).map(s => s.trim()).filter(Boolean);
    if (draft.kind === 'formula') override.formula = draft.formula;
    // 項目名（dbKey）はこのウィザードで設定しないが、STEP3で既に設定済みのセルの
    // 種類だけをここで変更した場合に消えてしまうと困るので、セルごとに現在のdbKeyを
    // そのまま引き継ぐ（infoは常に最新のCellInfoを指しているので、ov.dbKeyの発掘元として
    // OVERRIDESより確実）。
    infos.forEach((info) => {
      const id = CoreLogic.cellId(info);
      OVERRIDES[id] = Object.assign({}, override, { dbKey: info.dbKey || null });
    });
    closeModal();
    if (multi) { SELECTED.clear(); updateSelectionToolbar(); }
    rebuildAndRender();
  }

  function wizardNav({ onBack, onNext, nextLabel }) {
    const row = el('div', { class: 'm-btnrow' });
    if (onBack) {
      const backBtn = el('button', { class: 'secondary', text: '← 戻る', type: 'button' });
      backBtn.addEventListener('click', onBack);
      row.appendChild(backBtn);
    }
    const cancelBtn = el('button', { class: 'secondary', text: 'キャンセル', type: 'button' });
    cancelBtn.addEventListener('click', closeModal);
    row.appendChild(cancelBtn);
    const nextBtn = el('button', { text: nextLabel, type: 'button' });
    nextBtn.addEventListener('click', onNext);
    row.appendChild(nextBtn);
    return row;
  }

  function render() {
    box.innerHTML = '';
    const title = multi
      ? `セル設定（選択中の${infos.length}個をまとめて設定）`
      : `セル設定（${cellIds[0]}）`;
    box.appendChild(el('h3', { text: title }));

    if (step === 'kind') renderKindStep();
    else if (step === 'label') renderLabelStep();
    else if (step === 'options') renderOptionsStep();
    else if (step === 'formula') renderFormulaStep();
  }

  // 見出し・数式は、どちらもJSON出力の「値」にはならない＝項目名マッピングを持たない
  // （見出しはキー名そのものになる、数式は常に出力から除外される）という共通点があるため、
  // 選択画面でも「マッピングしない/する」の2グループに分けて示す。
  const KIND_GROUPS = [
    { title: '見出し・自動計算（データとしてマッピングしない）', kinds: ['label', 'formula'] },
    { title: '入力欄（データとしてマッピングする）', kinds: ['textarea', 'number', 'currency', 'select'] },
  ];

  function renderKindStep() {
    box.appendChild(el('p', { class: 'm-note', text: 'セルの種類を選んでください。' }));
    KIND_GROUPS.forEach(group => {
      box.appendChild(el('p', { class: 'wizard-group-title', text: group.title }));
      const choices = el('div', { class: 'wizard-choices' });
      group.kinds.forEach(k => {
        const isCurrent = k === draft.kind;
        const btn = el('button', {
          class: isCurrent ? '' : 'secondary', type: 'button',
          text: TYPE_LABELS[k] + (isCurrent ? '（現在の設定）' : ''),
        });
        btn.addEventListener('click', () => {
          draft.kind = k;
          if (k === 'label') { step = 'label'; render(); }
          else if (k === 'select') { step = 'options'; render(); }
          else if (k === 'formula') { step = 'formula'; render(); }
          else finish(); // テキスト/数値/金額は追加入力が無いので即確定する
        });
        choices.appendChild(btn);
      });
      box.appendChild(choices);
    });
    if (multi) {
      box.appendChild(el('p', { class: 'm-note', text: '見出しの文字・プルダウンの選択肢・数式は、選択した全セルに同じ内容が適用されます。' }));
    }
    const row = el('div', { class: 'm-btnrow' });
    const cancelBtn = el('button', { class: 'secondary', text: 'キャンセル', type: 'button' });
    cancelBtn.addEventListener('click', closeModal);
    row.appendChild(cancelBtn);
    box.appendChild(row);
  }

  function renderLabelStep() {
    box.appendChild(el('label', { text: '見出しの文字（空欄にすると入力不可の空白セルになります）' }));
    const input = el('input', { id: 'm-label', type: 'text', value: draft.label });
    box.appendChild(input);
    box.appendChild(wizardNav({
      onBack: () => { step = 'kind'; render(); },
      onNext: () => { draft.label = input.value; finish(); },
      nextLabel: '保存',
    }));
  }

  function renderOptionsStep() {
    box.appendChild(el('label', { text: 'プルダウンの選択肢（1行に1つ。Excel/スプレッドシートの列をそのまま貼り付け可、カンマ区切りも可）' }));
    const input = el('textarea', { id: 'm-options', rows: '8' });
    input.value = draft.options;
    box.appendChild(input);
    box.appendChild(wizardNav({
      onBack: () => { step = 'kind'; render(); },
      onNext: () => { draft.options = input.value; finish(); },
      nextLabel: '保存',
    }));
  }

  function renderFormulaStep() {
    box.appendChild(el('label', { text: '数式（例：=SUM(K11:L18)）' }));
    const input = el('input', { id: 'm-formula', type: 'text', value: draft.formula });
    box.appendChild(input);
    box.appendChild(el('p', {
      class: 'm-note',
      text: '対応しているのは =SUM(セル範囲) の形だけです（例：=SUM(K11:L18)）。'
        + '=K11+K12 のような式や、SUM以外の関数（AVERAGE等）は計算されません。'
        + 'セル範囲は、グリッド左端の行番号・上端の列アルファベットで確認できます。',
    }));
    box.appendChild(wizardNav({
      onBack: () => { step = 'kind'; render(); },
      onNext: () => { draft.formula = input.value; finish(); },
      nextLabel: '保存',
    }));
  }

  render();
}

function closeModal() {
  const existing = document.getElementById('modal-overlay');
  if (existing) existing.remove();
}

// ---------------------------------------------------------------------------
// 手動グループ化（A）：複数行を「親ラベル＋子レコードの配列」として明示的に
// 宣言する。Excelのタテ結合（R6年度等）があっても自動ではグループ化しない
// （見出し行とデータ行を区別できないため）。手動で宣言したグループは、
// このボタンで宣言したものだけがJSON出力上のグループになる。
// ---------------------------------------------------------------------------

// 選択中のセル群から、グループ化の対象となる行範囲（row0〜row1）を求める。
// マージセルの下端（row2）まで含めることで、選択が単位の途中で切れていても
// ユニット境界に揃うようにする。
function selectionRowRange() {
  const infos = selectedInfos();
  const row0 = Math.min(...infos.map(i => i.row));
  const row1 = Math.max(...infos.map(i => i.row2));
  return { row0, row1 };
}

function openManualGroupModal() {
  // 1個からでもグループ化できる（データが1行しかないグループに項目名を
  // 設定したい場合に対応するため。以前は2個以上が条件だったが、それだと単独行を
  // グループ化する手段がなかった）。
  if (SELECTED.size < 1) return;
  closeModal();
  const { row0, row1 } = selectionRowRange();
  const existingIdx = CURRENT.manualGroups.findIndex(g => g.row0 === row0 && g.row1 === row1 && !g.disabled);
  const existing = existingIdx >= 0 ? CURRENT.manualGroups[existingIdx] : null;
  // この選択の先頭行(row0)が、実際に自動検出（タテ結合）グループの先頭行になっているか。
  // なっていれば「自動グループ化を解除する」を選択肢として出せる。
  const autoRange = CoreLogic.getAutoGroupRange(CURRENT.grid, CURRENT.maxCol, row0);
  const disabledIdx = autoRange
    ? CURRENT.manualGroups.findIndex(g => g.row0 === autoRange.row0 && g.disabled)
    : -1;
  const isAutoDisabled = disabledIdx >= 0;

  const overlay = el('div', { class: 'modal-overlay', id: 'modal-overlay' });
  const box = el('div', { class: 'modal-box' });
  box.appendChild(el('h3', { text: `この範囲（${row0}〜${row1}行目）をグループ化する` }));
  box.appendChild(el('p', {
    class: 'm-note',
    text: '複数行の表（Excelでタテ結合が使われている表・いない表のどちらも）を、親ラベル＋子レコードの配列としてJSON出力できるようにします。'
      + '選択範囲は行単位で扱われます（列の選び方は結果に影響しません）。',
  }));
  box.appendChild(el('label', { text: '文脈（JSON出力時のキー名になります）' }));
  const nameInput = el('input', { id: 'm-group-name', type: 'text', value: existing ? existing.name : '' });
  box.appendChild(nameInput);

  const btnRow = el('div', { class: 'm-btnrow' });
  const saveBtn = el('button', { text: '保存', type: 'button' });
  saveBtn.addEventListener('click', () => {
    const name = nameInput.value.trim();
    if (!name) { nameInput.focus(); return; }
    // 1範囲=1グループの前提を保つため、選択範囲と重なる既存グループ（自動グループ化の
    // 解除マーカーを含む）は置き換える。
    CURRENT.manualGroups = CURRENT.manualGroups.filter(g => g.row1 < row0 || g.row0 > row1);
    CURRENT.manualGroups.push({ row0, row1, name });
    closeModal();
    SELECTED.clear();
    updateSelectionToolbar();
    rebuildAndRender();
  });
  const cancelBtn = el('button', { class: 'secondary', text: 'キャンセル', type: 'button' });
  cancelBtn.addEventListener('click', closeModal);
  btnRow.appendChild(saveBtn);
  btnRow.appendChild(cancelBtn);
  if (existing) {
    const delBtn = el('button', { class: 'secondary', text: 'グループ解除', type: 'button' });
    delBtn.addEventListener('click', () => {
      CURRENT.manualGroups.splice(existingIdx, 1);
      closeModal();
      SELECTED.clear();
      updateSelectionToolbar();
      rebuildAndRender();
    });
    btnRow.appendChild(delBtn);
  }
  box.appendChild(btnRow);

  // 自動グループ化（タテ結合）の解除／再有効化。すでに名前付きの手動グループが
  // この範囲にある場合は、そちらが優先されて自動検出自体が働かないので出さない。
  if (autoRange && !existing) {
    box.appendChild(el('p', {
      class: 'm-note modal-divider',
      text: `この範囲は現在、Excelのタテ結合から自動的に「${autoRange.label}」としてグループ化され`
        + `${isAutoDisabled ? 'ていません（解除済みです）' : 'ています'}（${autoRange.row0}〜${autoRange.row1}行目）。`,
    }));
    const toggleBtn = el('button', {
      class: 'secondary', type: 'button',
      text: isAutoDisabled ? '自動グループ化を有効に戻す' : '自動グループ化を解除する（グループ化しない）',
    });
    toggleBtn.addEventListener('click', () => {
      if (isAutoDisabled) {
        CURRENT.manualGroups.splice(disabledIdx, 1);
      } else {
        CURRENT.manualGroups = CURRENT.manualGroups.filter(g => g.row1 < autoRange.row0 || g.row0 > autoRange.row1);
        CURRENT.manualGroups.push({ row0: autoRange.row0, row1: autoRange.row1, name: null, disabled: true });
      }
      closeModal();
      SELECTED.clear();
      updateSelectionToolbar();
      rebuildAndRender();
    });
    box.appendChild(toggleBtn);
  }

  overlay.appendChild(box);
  overlay.addEventListener('click', (ev) => { if (ev.target === overlay) closeModal(); });
  document.body.appendChild(overlay);
}

// ---------------------------------------------------------------------------
// 複数選択（常時オン。ドラッグ／Ctrl+クリック）
// ---------------------------------------------------------------------------

function cellsInRect(a, b) {
  const minR = Math.min(a.row, b.row), maxR = Math.max(a.row, b.row);
  const minC = Math.min(a.col, b.col), maxC = Math.max(a.col, b.col);
  const out = [];
  const seen = new Set();
  for (let r = minR; r <= maxR; r++) {
    for (let c = minC; c <= maxC; c++) {
      const info = CURRENT.grid.get(r + ',' + c);
      if (!info) continue;
      const key = info.row + ',' + info.col;
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(info);
    }
  }
  return out;
}

function refreshSelectionHighlightDOM() {
  document.querySelectorAll('#grid-root .cell-selected').forEach(td => td.classList.remove('cell-selected'));
  SELECTED.forEach(key => {
    const [r, c] = key.split(',');
    const td = document.getElementById('td_R' + r + '_C' + c);
    if (td) td.classList.add('cell-selected');
  });
  updateSelectionToolbar();
}

function updateSelectionToolbar() {
  const bar = $('#selection-toolbar');
  if (!bar) return;
  if (SELECTED.size >= 1) {
    bar.style.display = 'flex';
    $('#selection-count').textContent = `選択中：${SELECTED.size}個`;
  } else {
    bar.style.display = 'none';
  }
}

function selectedInfos() {
  return Array.from(SELECTED).map(key => {
    const [r, c] = key.split(',');
    return CURRENT.grid.get(r + ',' + c);
  }).filter(Boolean);
}

function setupSelectionHandlers() {
  const gridRoot = $('#grid-root');

  gridRoot.addEventListener('mousedown', (ev) => {
    const td = ev.target.closest('td');
    if (!td || !td.dataset.row) return;
    DRAGGING = true;
    DRAG_ANCHOR = { row: +td.dataset.row, col: +td.dataset.col };
    DRAG_START_SNAPSHOT = (ev.ctrlKey || ev.metaKey) ? new Set(SELECTED) : new Set();
    SELECTED = new Set(DRAG_START_SNAPSHOT);
    cellsInRect(DRAG_ANCHOR, DRAG_ANCHOR).forEach(i => SELECTED.add(i.row + ',' + i.col));
    refreshSelectionHighlightDOM();
    ev.preventDefault();
  });

  gridRoot.addEventListener('mousemove', (ev) => {
    if (!DRAGGING) return;
    const td = ev.target.closest('td');
    if (!td || !td.dataset.row) return;
    const cur = { row: +td.dataset.row, col: +td.dataset.col };
    SELECTED = new Set(DRAG_START_SNAPSHOT);
    cellsInRect(DRAG_ANCHOR, cur).forEach(i => SELECTED.add(i.row + ',' + i.col));
    refreshSelectionHighlightDOM();
  });

  window.addEventListener('mouseup', (ev) => {
    if (!DRAGGING) return;
    DRAGGING = false;
    const td = ev.target.closest && ev.target.closest('td');
    if (!td || !td.dataset.row || !DRAG_ANCHOR) return;
    const cur = { row: +td.dataset.row, col: +td.dataset.col };
    const noDrag = cur.row === DRAG_ANCHOR.row && cur.col === DRAG_ANCHOR.col;
    // 実際にドラッグせず（＝同じセルでmousedownからmouseupまで動かず）、Ctrlを押した状態で
    // 「もともと選択済みだったセル」をクリックした場合だけ、そのセルを選択解除する。
    // （以前は「今回の操作でSELECTEDが1個増えたか」で判定していたが、すでに選択済みの
    // セルをクリックしてもSet.add()はサイズを増やさないため、常にfalseになり
    // 個別解除が一切効かないバグだった）
    if (noDrag && (ev.ctrlKey || ev.metaKey)) {
      const key = cur.row + ',' + cur.col;
      if (DRAG_START_SNAPSHOT.has(key)) {
        SELECTED.delete(key);
        refreshSelectionHighlightDOM();
      }
    }
  });

  $('#btn-selection-apply').addEventListener('click', () => {
    if (SELECTED.size < 1) return;
    openCellSettings(selectedInfos());
  });
  $('#btn-selection-group').addEventListener('click', openManualGroupModal);
  $('#btn-selection-clear').addEventListener('click', () => {
    SELECTED.clear();
    refreshSelectionHighlightDOM();
  });
}

// ---------------------------------------------------------------------------
// 入力フォーム（ツール2）としてのHTML書き出し
// ---------------------------------------------------------------------------

// 書き出すフォームのタイトル（<title>/<h1>・ダウンロードファイル名の元）。
// STEP4でユーザーが指定した場合はそれを優先し、未指定ならアップロードしたファイル名から
// 既定値を作る。CURRENT.customTitleに保持することで、パネルが再描画されても
// （renderCheckPanelはDOMを毎回作り直すため）入力済みの値が消えない。
function effectiveFormTitle() {
  if (CURRENT.customTitle && CURRENT.customTitle.trim()) return CURRENT.customTitle.trim();
  return (CURRENT.fileName || '入力フォーム').replace(/\.xlsx$/i, '');
}

function serializeStructure() {
  const seen = new Set();
  const cells = [];
  for (const info of CURRENT.grid.values()) {
    const key = info.row + ',' + info.col;
    if (seen.has(key)) continue;
    seen.add(key);
    cells.push({
      row: info.row, col: info.col, row2: info.row2, col2: info.col2,
      value: info.value, isFormula: info.isFormula, formula: info.formula,
      hasText: info.hasText, blocked: info.blocked, fillColor: info.fillColor,
      renderType: info.renderType, renderOptions: info.renderOptions,
      dbKey: info.dbKey,
    });
  }
  const formTitle = effectiveFormTitle();
  // ファイル名に含める項目・必須項目・2次入力欄・2次入力画面表示項目（いずれもcellIdの配列、
  // シート上の並び順）。findMappingTargetsを再実行してシート順（singlesは既にrow/col順）を
  // 基準にし、選択順（Set挿入順）に依存しないようにする。
  const { singles: allSingles } = CoreLogic.findMappingTargets(CURRENT.grid, CURRENT.sections, CURRENT.maxCol, CURRENT.manualGroups);
  const idsInSheetOrder = (set) => set
    ? allSingles.map(t => CoreLogic.cellId(t.cells[0])).filter(id => set.has(id))
    : [];
  const fileNameFields = idsInSheetOrder(CURRENT.fileNameFieldIds);
  const requiredFields = idsInSheetOrder(CURRENT.requiredFieldIds);
  const secondaryFields = idsInSheetOrder(CURRENT.secondaryFieldIds);
  // 2次入力画面表示項目：事務局が個別に選んだ項目に加え、ファイル名項目は常に自動的に候補へ
  // 含める（同じ項目を2回チェックさせない、というユーザー承認済みの仕様）。
  const displayCandidateSet = new Set([...(CURRENT.fileNameFieldIds || []), ...(CURRENT.displayCandidateFieldIds || [])]);
  const displayCandidateFields = idsInSheetOrder(displayCandidateSet);
  return {
    formTitle,
    fileNameFields, requiredFields, secondaryFields, displayCandidateFields,
    maxRow: CURRENT.maxRow, maxCol: CURRENT.maxCol,
    widths: CURRENT.widths, heights: CURRENT.heights,
    sections: CURRENT.sections.map(s => ({ title: s.title, row0: s.row0, row1: s.row1 })),
    cells,
    // 手動グループ（A）：入力フォーム（ツール2）側のJSON書き出し・読込にもそのまま必要なため、
    // 確定済み構造の一部として一緒に埋め込む。
    manualGroups: CURRENT.manualGroups || [],
  };
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
// 書き込みに失敗した場合（共有フォルダの権限が無い等）は、通常のダウンロードに
// フォールバックする。
// 戻り値："dir"（直接保存）｜"download-failed"（書き込み失敗によるフォールバック）｜
// "cancelled"（上書きを拒否）。"download-failed"を"download"（フォルダ未指定時の通常
// ダウンロード）と区別する理由も含め、filler_app.jsの同名関数群と設計は同一
// （詳細はそちらのコメント参照）。
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

// EXPORT_DIR_HANDLEが指定されていれば非同期でフォルダへ直接書き込み、無ければ従来通り
// 即座にダウンロードする（同期）。onDoneには結果（"dir"|"download"|"cancelled"）が渡る。
// フォルダ未指定時にあえて同期のままにしている理由はfiller_app.jsのsaveBlobと同じ
// （常にasync/awaitへ統一すると、書き出し直後に#statusを同期的に確認している
// 既存の挙動・テストが崩れるため）。
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

// Chromium限定：書き出し先フォルダを1回指定する。以後、このビルダーからの入力フォームHTML
// 書き出しはこのフォルダへ直接保存される。window.showDirectoryPickerが無いブラウザでは
// ボタン自体をrenderCheckPanel側で表示しない。
async function pickExportDirectory() {
  if (!window.showDirectoryPicker) return;
  try {
    EXPORT_DIR_HANDLE = await window.showDirectoryPicker({ mode: 'readwrite' });
  } catch (e) {
    return; // ユーザーがキャンセルした場合等。エラー表示はしない。
  }
  const statusEl = $('#export-dir-status');
  if (statusEl) statusEl.textContent = `書き出し先：${EXPORT_DIR_HANDLE.name}（以後この入力フォームの書き出しはここへ直接保存されます）`;
}

function doExportAsForm() {
  if (!CURRENT || !CURRENT.grid) return;
  const structure = serializeStructure();
  const html = FILLER_TEMPLATE
    .replace(/__FORM_TITLE__/g, structure.formTitle)
    .replace('/* __STRUCTURE__ */', JSON.stringify(structure));
  const filename = `${structure.formTitle}_入力フォーム.html`;
  const blob = new Blob([html], { type: 'text/html' });
  saveBlob(blob, filename, (result) => {
    setStatus(statusMessageForSave(result, filename, '入力フォームを書き出しました。'));
  });
}

// 集約ツール（ツール3）の書き出し。doExportAsForm()と同じ「テンプレート文字列注入＋
// Blobダウンロード」の仕組みだが、埋め込むのはSTRUCTURE（構造）だけで、データは
// 集約ツール自身がSTEP1で読み込む（このビルダーはデータを持たない）。
function doExportAsAggregator() {
  if (!CURRENT || !CURRENT.grid) return;
  const structure = serializeStructure();
  const html = AGGREGATOR_TEMPLATE
    .replace(/__FORM_TITLE__/g, structure.formTitle)
    .replace('/* __STRUCTURE__ */', JSON.stringify(structure));
  const filename = `${structure.formTitle}_集約ツール.html`;
  const blob = new Blob([html], { type: 'text/html' });
  saveBlob(blob, filename, (result) => {
    setStatus(statusMessageForSave(result, filename, '集約ツールを書き出しました。'));
  });
}

// STEP5の書き出すボタン（1つ）から、入力フォーム（ツール2）・集約ツール（ツール3）の
// 2ファイルをまとめて書き出す。両方とも今のSTRUCTURE（構造）だけから決まり、
// データを持たないため、片方だけ書き出して後で構造を変えると集約ツール側が
// 古いまま取り残される事故を防げる（doExportAsForm/doExportAsAggregatorは
// 既存テストが個別に検証しているためそのまま流用し、ここでは呼び出しを束ねるだけ）。
function statusMessageForBothSave(formResult, aggregatorResult) {
  if (formResult === 'cancelled' || aggregatorResult === 'cancelled') {
    return '書き出しを中止しました（指定フォルダに同名ファイルが既にあります）。';
  }
  if (formResult === 'download-failed' || aggregatorResult === 'download-failed') {
    return '⚠️ 指定フォルダへの保存に失敗したため、通常のダウンロードで保存しました（入力フォーム・集約ツール）。';
  }
  if (formResult === 'dir' && aggregatorResult === 'dir') {
    return '入力フォーム・集約ツールの2ファイルを指定フォルダへ保存しました。';
  }
  return '入力フォーム・集約ツールの2ファイルを書き出しました。';
}

function doExportAsBoth() {
  if (!CURRENT || !CURRENT.grid) return;
  const structure = serializeStructure();
  const formHtml = FILLER_TEMPLATE
    .replace(/__FORM_TITLE__/g, structure.formTitle)
    .replace('/* __STRUCTURE__ */', JSON.stringify(structure));
  const formFilename = `${structure.formTitle}_入力フォーム.html`;
  const formBlob = new Blob([formHtml], { type: 'text/html' });

  const aggregatorHtml = AGGREGATOR_TEMPLATE
    .replace(/__FORM_TITLE__/g, structure.formTitle)
    .replace('/* __STRUCTURE__ */', JSON.stringify(structure));
  const aggregatorFilename = `${structure.formTitle}_集約ツール.html`;
  const aggregatorBlob = new Blob([aggregatorHtml], { type: 'text/html' });

  saveBlob(formBlob, formFilename, (formResult) => {
    saveBlob(aggregatorBlob, aggregatorFilename, (aggregatorResult) => {
      setStatus(statusMessageForBothSave(formResult, aggregatorResult));
    });
  });
}

// ---------------------------------------------------------------------------
// 書き出し前チェック：現在の設定で入力しても消えてしまうセルがないかを検証し、
// 結果に応じてプレビュー・一括修正の導線・書き出し可否を提示する。
// ---------------------------------------------------------------------------

// 指定した行範囲に実在する全セルを選択状態にする（グループ化ショートカット用）。
function selectRowRangeAllColumns(row0, row1) {
  SELECTED = new Set();
  for (let r = row0; r <= row1; r++) {
    for (let c = 1; c <= CURRENT.maxCol; c++) {
      const info = CURRENT.grid.get(r + ',' + c);
      if (info) SELECTED.add(info.row + ',' + info.col);
    }
  }
  refreshSelectionHighlightDOM();
}

// 検出された「書き出せないセル」を、連続する行範囲にまとめる
// （修正アクションは常に「この行範囲をグループ化する」なので、行単位でまとめれば十分）。
function summarizeUnreachableRanges(unreachable) {
  const rows = Array.from(new Set(unreachable.map(c => c.row))).sort((a, b) => a - b);
  const ranges = [];
  let start = null, prev = null;
  rows.forEach(r => {
    if (start === null) { start = r; prev = r; return; }
    if (r === prev + 1) { prev = r; return; }
    ranges.push({ row0: start, row1: prev });
    start = r; prev = r;
  });
  if (start !== null) ranges.push({ row0: start, row1: prev });
  return ranges;
}

// ---------------------------------------------------------------------------
// プレビュー：生のJSONではなく、非技術者でも読める表形式に組み立てる。
// 各セクション内を「単純な項目（項目名→値）」と「グループ（配列）」に分け、
// グループは項目名を列見出しにした通常の表として描画する。
// ---------------------------------------------------------------------------
function stringifyLeaf(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') return JSON.stringify(v);
  return String(v);
}

function buildPreviewTables(output) {
  const container = el('div', { class: 'preview-tables' });

  Object.keys(output).forEach(secTitle => {
    const secData = output[secTitle];
    if (!secData || typeof secData !== 'object') return;

    const simpleRows = [];
    const groups = [];

    Object.keys(secData).forEach(key => {
      const val = secData[key];
      if (Array.isArray(val)) {
        groups.push({ name: key, records: val });
      } else if (val && typeof val === 'object') {
        // 「事業名」のように1項目に複数の値がぶら下がる複合項目。子を「親 > 子」の行として展開する。
        Object.keys(val).forEach(subKey => {
          simpleRows.push([`${key} > ${subKey}`, stringifyLeaf(val[subKey])]);
        });
      } else {
        simpleRows.push([key, stringifyLeaf(val)]);
      }
    });

    if (simpleRows.length === 0 && groups.length === 0) return;

    container.appendChild(el('h4', { class: 'preview-sec-title', text: secTitle }));

    if (simpleRows.length > 0) {
      const wrap = el('div', { class: 'preview-table-wrap' });
      const table = el('table', { class: 'preview-table' });
      // 「単独の入力欄」のマッピング表では「文脈」列と「項目名」列を分けて表示しているが、
      // こちらは1本の文字列に結合して見せる設計（親子関係を持つ行は「文脈 > 項目名」の形）
      // なので、見出しもその実態に合わせる（結合しない行＝文脈が無い行もあるが、その場合は
      // 単に項目名だけが入るので見出しとして矛盾はしない）。
      table.appendChild(el('thead', {}, [el('tr', {}, [el('th', { text: '文脈 > 項目名' }), el('th', { text: '値／セル' })])]));
      const tbody = el('tbody');
      simpleRows.forEach(([k, v]) => {
        tbody.appendChild(el('tr', {}, [el('td', { text: k }), el('td', { text: v })]));
      });
      table.appendChild(tbody);
      wrap.appendChild(table);
      container.appendChild(wrap);
    }

    groups.forEach(g => {
      container.appendChild(el('p', { class: 'preview-group-title', text: `🗂 ${g.name}（グループ・${g.records.length}件）` }));
      const cols = [];
      g.records.forEach(r => {
        if (r && typeof r === 'object' && !Array.isArray(r)) {
          Object.keys(r).forEach(k => { if (cols.indexOf(k) === -1) cols.push(k); });
        }
      });
      const wrap = el('div', { class: 'preview-table-wrap' });
      const table = el('table', { class: 'preview-table' });
      table.appendChild(el('thead', {}, [el('tr', {}, [el('th', { text: '#' })].concat(cols.map(c => el('th', { text: c }))))]));
      const tbody = el('tbody');
      g.records.forEach((r, idx) => {
        const cells = [el('td', { text: String(idx + 1) })];
        if (r && typeof r === 'object' && !Array.isArray(r)) {
          cols.forEach(c => cells.push(el('td', { text: r[c] !== undefined ? stringifyLeaf(r[c]) : '' })));
        } else {
          cells.push(el('td', { text: stringifyLeaf(r), colspan: String(cols.length || 1) }));
        }
        tbody.appendChild(el('tr', {}, cells));
      });
      table.appendChild(tbody);
      wrap.appendChild(table);
      container.appendChild(wrap);
    });
  });

  if (!container.firstChild) {
    container.appendChild(el('p', { class: 'm-note', text: '（データがありません）' }));
  }
  return container;
}

// ---------------------------------------------------------------------------
// 項目名（dbKey）マッピングの一覧・一括保存。「単独の入力欄」と「グループ化された列」を
// 別ブロックにする（保存すると何行に影響するかが本質的に異なるため）。
// グループ化された列は、同じグループ・同じ列のセルをまとめて1件として表示し、
// 保存するとその列の全行に同じ名前を一括適用する（1行だけ個別に名前を付けると
// 配列の要素ごとにキー名がバラバラになってしまう実害が実際にあったため）。
// 触っていない入力欄（表示時の値のまま）は保存時にスキップする
// （data-initialとの比較。既存の名前を保存のたびに無意味に書き戻さないため）。
// ---------------------------------------------------------------------------
function mappingInputId(target) {
  const prefix = target.kind === 'group' ? 'mapgroup_' : 'mapsingle_';
  return prefix + CoreLogic.cellId(target.cells[0]);
}

// グループの列は複数セルにまたがるため、セルごとにdbKeyがバラバラ（一部だけ設定済み・
// 名前が食い違っている等）というケースがありうる。空欄のまま何もしないと見た目上は
// 「まだ何も設定していない」のと区別が付かず気づきにくいため、mixedフラグを返して
// 呼び出し側でプレースホルダー等に明示する。
function mappingCurrentName(target) {
  const names = new Set(target.cells.map(c => c.dbKey || ''));
  if (names.size === 1) return { value: names.values().next().value, mixed: false };
  return { value: '', mixed: true };
}

function mappingRefLabel(target) {
  if (target.kind !== 'group') return CoreLogic.cellRef(target.cells[0]);
  // 「7行分」のような件数だけでは実際にどこのセルか分からないため、
  // グリッド上のセル番地の範囲（例：B11:B17）で示す。cellsは行の昇順に並んでいる
  // （buildSectionObjectの走査順そのままなので、行番号は必ず先頭が最小・末尾が最大になる）。
  const first = CoreLogic.cellRef(target.cells[0]);
  if (target.cells.length === 1) return `🗂 ${first}`;
  const last = CoreLogic.cellRef(target.cells[target.cells.length - 1]);
  return `🗂 ${first}:${last}`;
}

// 文脈列：グループ化された列はgroupLabelがそのまま実際のJSON出力の配列キーなので
// 正確に表示する。一方、単独の入力欄は「行自体の見出し（rowLabel）」だけを表示し、
// groupLabelは出さない。単独セルがたまたまグループの縦結合範囲の内側にあっても、
// 実際のJSON出力ではそのグループの配列には入らず独立したキーになる（K18＝「その他」の
// 実例）ため、「グループ名 > 行の見出し」のように連結して見せると、あたかも入れ子で
// あるかのような誤解を招いていた。行自体の見出しが無い場合のみ「（見出しの無い行）」とする。
function mappingContextLabel(target) {
  if (target.kind === 'group') return target.groupLabel;
  return target.rowLabel || '（見出しの無い行）';
}

function buildMappingBlock(title, noteText, targets) {
  if (targets.length === 0) return null;
  const wrap = el('div', { class: 'mapping-block' });
  wrap.appendChild(el('p', { class: 'mapping-block-title', text: title }));
  wrap.appendChild(el('p', { class: 'm-note', text: noteText }));

  const table = el('table', { class: 'mapping-table' });
  table.appendChild(el('thead', {}, [el('tr', {}, [
    el('th', { text: '対象セル' }), el('th', { text: '文脈' }), el('th', { text: '項目名' }),
  ])]));
  const tbody = el('tbody');
  targets.forEach((target) => {
    const row = el('tr', { class: 'mapping-item' });

    const refBtn = el('button', { class: 'secondary mapping-ref', type: 'button', title: 'このセルをグリッド上で確認する', text: mappingRefLabel(target) });
    refBtn.addEventListener('click', () => {
      const anchor = target.cells[0];
      const td = document.getElementById('td_R' + anchor.row + '_C' + anchor.col);
      if (td) td.scrollIntoView({ block: 'center', inline: 'center' });
    });
    row.appendChild(el('td', {}, [refBtn]));

    row.appendChild(el('td', { class: 'mapping-context', text: mappingContextLabel(target) }));

    const { value: currentName, mixed } = mappingCurrentName(target);
    // autoName：dbKey未設定のときにこのセルが実際に使っている既定名。見出し由来なら
    // その文字列、無ければcolN。フォールバックはグループの代表列番号／単独セルの列番号。
    const fallbackCol = target.kind === 'group' ? target.col : target.cells[0].col;
    const autoName = target.autoName || ('col' + fallbackCol);
    // 空欄のまま保存した場合に実際に使われる項目名をそのまま見せる（「〜のまま」という
    // 説明文言を付けると、あたかも別の名前が付くかのように誤読されるため、実際に
    // autoNameを入力したときと同じ見た目にする）。mixedは「空欄のままだと列内の行ごとに
    // 既存のバラバラな名前が残る」という別の状態なので、autoName単体では表現しきれず注記を残す。
    const placeholder = mixed ? `例: ${autoName}（列内で項目名が不統一。空欄のままだと行ごとに既存の名前が残ります）` : autoName;
    const input = el('input', { type: 'text', id: mappingInputId(target), value: currentName, placeholder });
    input.dataset.initial = currentName;
    // 空欄のまま何もしない状態と見分けが付くよう、行自体にも印を付ける
    // （プレースホルダーだけだと文字が薄く見落とされやすいため）。
    // colNのまま（＝Excelの見出しに由来しない、意味のある名前が無い）欄は、手直しの
    // 優先度が高いことが一目で分かるよう別途目立たせる（mixedとは別軸の状態なので、
    // 見た目が衝突しないよう互いに排他にする）。
    const isGeneric = /^col\d+$/.test(autoName);
    if (mixed) row.classList.add('mapping-mixed');
    else if (isGeneric && !currentName) row.classList.add('mapping-generic');
    row.appendChild(el('td', {}, [input]));

    tbody.appendChild(row);
  });
  table.appendChild(tbody);
  const listWrap = el('div', { class: 'mapping-list' });
  listWrap.appendChild(table);
  wrap.appendChild(listWrap);

  const saveBtn = el('button', { class: 'secondary mapping-save-btn', text: 'まとめて保存', type: 'button' });
  saveBtn.addEventListener('click', () => saveMapping(targets));
  wrap.appendChild(saveBtn);
  return wrap;
}

function saveMapping(targets) {
  let changed = 0;
  targets.forEach((target) => {
    const input = document.getElementById(mappingInputId(target));
    if (!input) return;
    const val = input.value.trim();
    if (val === input.dataset.initial) return; // 触っていない欄はスキップ
    target.cells.forEach((cell) => {
      const id = CoreLogic.cellId(cell);
      const existing = OVERRIDES[id] || { kind: currentKindOf(cell) };
      OVERRIDES[id] = Object.assign({}, existing, { dbKey: val || null });
    });
    changed++;
  });
  if (changed > 0) rebuildAndRender();
}

function buildMappingSection(singles, groups) {
  if (singles.length === 0 && groups.length === 0) return null;
  const wrap = el('div', { class: 'mapping-section' });
  const singleBlock = buildMappingBlock(
    `単独の入力欄（${singles.length}件）`,
    '対象セルに応じた項目名を入力し「まとめて保存」を押してください。色が付いている欄は自動名（colN）です、必要に応じて項目名を変更してください。',
    singles,
  );
  if (singleBlock) wrap.appendChild(singleBlock);
  const groupBlock = buildMappingBlock(
    `グループ化された入力欄（${groups.length}件）`,
    '対象セルに応じた項目名を入力し「まとめて保存」を押してください。',
    groups,
  );
  if (groupBlock) wrap.appendChild(groupBlock);
  return wrap;
}

// 各部署が入力フォーム（ツール2）側で「書き出す」を押したときのJSONファイル名に、
// どの項目の値を含めるかを選ぶUI（複数選択可）。グループ化された列（繰り返し行）は
// 値が複数あってファイル名にできないため対象外とし、singlesのみを候補にする。
// 選択順ではなく常にシート上の並び順（singlesは既にrow/col順にソート済み）で連結する
// ため、ユーザーが後から選択し直しても連結順が入れ替わらない。
// 選択状態はCURRENT.fileNameFieldIds（Set<cellId>）に保持し、パネル再描画をまたいで残す。
function buildFileNameFieldPicker(singles) {
  if (singles.length === 0) return null;
  const wrap = el('div', { class: 'filename-field-picker' });
  wrap.appendChild(el('p', { class: 'm-note', text: '各部署が入力後に書き出すJSONのファイル名に含める項目を選べます（複数可・シート上の並び順で連結されます）。選ばなければ従来通り日付だけのファイル名になります。' }));
  const list = el('div', { class: 'filename-field-list' });
  singles.forEach((target) => {
    const cellId = CoreLogic.cellId(target.cells[0]);
    const { value: currentName } = mappingCurrentName(target);
    const label = `${currentName || target.autoName || ('col' + target.cells[0].col)}（${CoreLogic.cellRef(target.cells[0])}）`;
    const inputId = 'filenamefield_' + cellId;
    const checkbox = el('input', { type: 'checkbox', id: inputId });
    if (CURRENT.fileNameFieldIds && CURRENT.fileNameFieldIds.has(cellId)) checkbox.setAttribute('checked', 'checked');
    checkbox.addEventListener('change', () => {
      if (!CURRENT.fileNameFieldIds) CURRENT.fileNameFieldIds = new Set();
      if (checkbox.checked) CURRENT.fileNameFieldIds.add(cellId);
      else CURRENT.fileNameFieldIds.delete(cellId);
      // 2次入力画面表示項目ピッカーのforceCheckedIds（ファイル名項目の自動候補化）を
      // 反映させるため、チェック直後にパネル全体を再描画する。
      renderCheckPanel();
    });
    const item = el('label', { for: inputId, class: 'filename-field-item' }, [checkbox, el('span', { text: label })]);
    list.appendChild(item);
  });
  wrap.appendChild(list);
  return wrap;
}

// 必須項目／2次入力欄／2次入力画面表示項目のピッカーを組み立てる汎用版。
// buildFileNameFieldPickerと見た目・構造は同じだが、対象のCURRENTプロパティ（storeKey）と
// チェックボックスのidプレフィックスだけを差し替えられるようにしてある
// （ファイル名項目は既存テスト・既存挙動を壊さないよう、専用関数のまま触らずに残した）。
// opts.forceCheckedIds（Set）を渡すと、そのcellIdは常にチェック済み・操作不可（disabled）で
// 表示する（2次入力画面表示項目のうち、ファイル名項目として選択済みのものを自動的に
// 候補へ含める＝重複してチェックさせない、という仕様のため）。
function buildGenericFieldPicker(singles, opts) {
  if (singles.length === 0) return null;
  const wrap = el('div', { class: 'filename-field-picker' });
  wrap.appendChild(el('p', { class: 'm-note', html: `<strong>${opts.title}</strong>` }));
  wrap.appendChild(el('p', { class: 'm-note', text: opts.noteText }));
  const list = el('div', { class: 'filename-field-list' });
  singles.forEach((target) => {
    const cellId = CoreLogic.cellId(target.cells[0]);
    const { value: currentName } = mappingCurrentName(target);
    const label = `${currentName || target.autoName || ('col' + target.cells[0].col)}（${CoreLogic.cellRef(target.cells[0])}）`;
    const inputId = opts.idPrefix + '_' + cellId;
    const forceChecked = opts.forceCheckedIds && opts.forceCheckedIds.has(cellId);
    const checkbox = el('input', { type: 'checkbox', id: inputId });
    const currentSet = CURRENT[opts.storeKey];
    if (forceChecked || (currentSet && currentSet.has(cellId))) checkbox.setAttribute('checked', 'checked');
    if (forceChecked) {
      checkbox.setAttribute('disabled', 'disabled');
    } else {
      checkbox.addEventListener('change', () => {
        if (!CURRENT[opts.storeKey]) CURRENT[opts.storeKey] = new Set();
        if (checkbox.checked) CURRENT[opts.storeKey].add(cellId);
        else CURRENT[opts.storeKey].delete(cellId);
      });
    }
    const labelText = forceChecked ? `${label}（ファイル名項目のため常に候補）` : label;
    const item = el('label', { for: inputId, class: 'filename-field-item' }, [checkbox, el('span', { text: labelText })]);
    list.appendChild(item);
  });
  wrap.appendChild(list);
  return wrap;
}

// このチェックはグリッド上のセル位置と見比べながら常時参照するものなので、
// グリッドを覆い隠す中央モーダルではなく、右側に常駐するサイドパネルとして
// ファイル読込直後から表示し、グリッドの状態が変わるたびに（rebuildAndRender経由で）
// 自動的に最新化する。ユーザーがこれを見ながら手直し・グループ化を進める想定。

// STEP3〜5はパネルが縦に長くなりがちなので、それぞれ折りたたみできるようにする。
// renderCheckPanel()は手直しのたびにパネル全体を作り直すため、開閉状態はここ
// （モジュールスコープ）で保持し、再描画のたびに復元する（既定はopen＝従来通り
// 全部開いた状態のまま、ユーザーが必要に応じて畳んで縦の長さを詰められるようにする）。
// STEP3の見出しは「データ構造確認」の直下という既存の位置を維持したいが、その中身
// （マッピング表）は書き出し前チェックの警告文の後に離れて置かれているため、native
// <details>（summary/contentが親子関係必須）ではなくクリックで表示切替する自前実装にする。
const STEP_PANEL_OPEN = { step3: true, step4: true, step5: true };

function buildStepLabel(key, guideIdx) {
  const label = el('p', {
    class: 'step-label step-toggle',
    text: (STEP_PANEL_OPEN[key] ? '▾ ' : '▸ ') + STEP_GUIDE[guideIdx].label,
  });
  label.addEventListener('click', () => {
    STEP_PANEL_OPEN[key] = !STEP_PANEL_OPEN[key];
    renderCheckPanel();
  });
  return label;
}

function buildStepContentWrap(key, contentElements) {
  const wrap = el('div', { class: 'step-content' });
  if (!STEP_PANEL_OPEN[key]) wrap.style.display = 'none';
  contentElements.forEach((elx) => { if (elx) wrap.appendChild(elx); });
  return wrap;
}

function renderCheckPanel() {
  if (!CURRENT || !CURRENT.grid) return;
  const { output, unreachable } = CoreLogic.findUnreachableCells(CURRENT.grid, CURRENT.sections, CURRENT.maxCol, CURRENT.manualGroups);
  // グループ化で直せるもの／どのセクションにも属さないため直せないものを分ける
  // （後者に「グループ化する」ボタンを出すと、機能しないのに解決策のように見えてしまうため）。
  const groupableAll = unreachable.filter(c => !c.outOfSection);
  const outOfSection = unreachable.filter(c => c.outOfSection);
  // groupableAllをさらに2種に分ける。行自体が見出し（1列目にテキストがある）を
  // 持っているのに未到達なセルは、「グループ化されていない」のが原因ではなく、
  // 他の行と見出しの文字が重複していて出力時に上書きされている可能性が高い
  // （行に見出しがあれば通常はそのままJSON出力されるため）。この場合はグループ化しても
  // 直らないので、「グループ化する」ボタンではなく別の案内を出す
  // （rowHasOwnLabelはautoBlockUnreachableCellsとも共有している判定）。
  const labelCollision = groupableAll.filter(c => rowHasOwnLabel(CURRENT.grid, c));
  const groupable = groupableAll.filter(c => !rowHasOwnLabel(CURRENT.grid, c));
  const ranges = summarizeUnreachableRanges(groupable);

  let panel = document.getElementById('check-panel');
  let savedScrollTop = 0;
  let savedPreviewScrollTop = 0;
  if (panel) {
    savedScrollTop = panel.scrollTop;
    // プレビュー表（.preview-tables）は独自のスクロール領域を持つ内側のコンテナで、
    // panel.innerHTML=''のたびに丸ごと作り直されるため、外側のpanel.scrollTopとは
    // 別にこちらのスクロール位置も保存しておかないと、手直しのたびに先頭へ戻ってしまう
    // （下へ下へ確認しながら作業する運用と相性が悪い）。
    const oldPreviewTables = panel.querySelector('.preview-tables');
    if (oldPreviewTables) savedPreviewScrollTop = oldPreviewTables.scrollTop;
    panel.innerHTML = '';
  } else {
    panel = el('div', { class: 'side-panel', id: 'check-panel' });
    document.body.classList.add('side-panel-open');
    document.body.appendChild(panel);
  }
  const box = panel;

  box.appendChild(el('h3', { text: 'データ構造確認' }));
  box.appendChild(buildStepLabel('step3', 2));
  box.appendChild(el('p', { class: 'step-memo', text: '💡 Excelの見出しは人が読むためのものですが、データベース化にはコンピュータが扱える固有の名前が必要です。ここでその名前を決めます。' }));
  box.appendChild(el('p', { class: 'm-note', text: 'セルの設定を行うたびに自動で更新されます。' }));

  if (unreachable.length > 0) {
    if (groupable.length > 0) {
      box.appendChild(el('p', {
        class: 'm-note m-warn',
        text: `⚠ 以下の${groupable.length}個のセルは、今の設定のままだと入力してもJSONに反映されません`
          + '（見出しのない行が、グループ化されていないことが主な原因です）。',
      }));
      const list = el('div', { class: 'unreachable-list' });
      ranges.forEach(rg => {
        const refs = groupable
          .filter(c => c.row >= rg.row0 && c.row <= rg.row1)
          .map(c => CoreLogic.cellRef(c));
        const item = el('div', { class: 'unreachable-item' });
        const label = refs.length > 6 ? refs.slice(0, 6).join('、') + ' 他' : refs.join('、');
        item.appendChild(el('span', { text: `${rg.row0}〜${rg.row1}行目（${label}）` }));
        const fixBtn = el('button', { class: 'secondary', text: 'この範囲を選択してグループ化する', type: 'button' });
        fixBtn.addEventListener('click', () => {
          selectRowRangeAllColumns(rg.row0, rg.row1);
          openManualGroupModal();
        });
        item.appendChild(fixBtn);
        list.appendChild(item);
      });
      box.appendChild(list);
    }
    if (outOfSection.length > 0) {
      const refs = outOfSection.map(c => CoreLogic.cellRef(c));
      const label = refs.length > 10 ? refs.slice(0, 10).join('、') + ' 他' : refs.join('、');
      box.appendChild(el('p', {
        class: 'm-note m-warn',
        text: `⚠ 以下の${outOfSection.length}個のセルは、見出し（Plan／Do等）より前にあるため、`
          + `現在の仕組みでは書き出し対象にできません（グループ化しても解決しません）：${label}`,
      }));
    }
    if (labelCollision.length > 0) {
      const refs = labelCollision.map(c => CoreLogic.cellRef(c));
      const label = refs.length > 10 ? refs.slice(0, 10).join('、') + ' 他' : refs.join('、');
      box.appendChild(el('p', {
        class: 'm-note m-warn',
        text: `⚠ 以下の${labelCollision.length}個のセルは、見出しの文字が他の箇所（別の行、または同じ行の中の別のセル）と重複しているため、`
          + `入力してもどちらか一方の内容で上書きされてしまいます（グループ化しても解決しません。`
          + `元のExcelファイルで見出しの文字を書き分けてから読み込み直してください）：${label}`,
      }));
    }
  }

  // STEP3：項目名マッピング。findUnreachableCellsの結果、そもそもJSONに含まれない
  // セル（unreachable）は先にグループ化すべき問題であり、ここで名前を付けても
  // 解決しないため除外する。
  const unreachableIds = new Set(unreachable.map(c => CoreLogic.cellId(c)));
  const { singles, groups } = CoreLogic.findMappingTargets(CURRENT.grid, CURRENT.sections, CURRENT.maxCol, CURRENT.manualGroups);
  const isReachable = (target) => !target.cells.some(c => unreachableIds.has(CoreLogic.cellId(c)));
  const mappingEl = buildMappingSection(singles.filter(isReachable), groups.filter(isReachable));
  if (mappingEl) box.appendChild(buildStepContentWrap('step3', [mappingEl]));

  // STEP4：項目の運用設定。STEP3で名前を付けた項目に対して、追加のフラグ（必須・
  // 2次入力画面表示・2次入力）を付ける作業なので、STEP3の直後に置く。
  box.appendChild(buildStepLabel('step4', 3));
  box.appendChild(el('p', { class: 'step-memo', text: '💡 必須化・2次入力・一覧表示といった、運用上の使い勝手を整えるための設定です（省略しても書き出しはできます）。' }));

  const reachableSingles = singles.filter(isReachable);
  const requiredFieldEl = buildGenericFieldPicker(reachableSingles, {
    idPrefix: 'requiredfield', storeKey: 'requiredFieldIds', title: '必須項目',
    noteText: '入力必須の項目を選べます（複数可）。書き出し前チェックで、未入力のまま書き出そうとした場合に警告します。',
  });

  const secondaryFieldEl = buildGenericFieldPicker(reachableSingles, {
    idPrefix: 'secondaryfield', storeKey: 'secondaryFieldIds', title: '2次入力欄',
    noteText: '2次入力者が判定・コメント等を書き込む項目を選べます（複数可）。この様式を複数まとめて2次入力する画面（一覧表）で、編集可能な列として扱われます。',
  });

  // 2次入力画面表示項目：ファイル名項目は常に候補へ自動的に含める（同じ項目を2回
  // チェックさせない、というユーザー承認済みの仕様）。実際に2次入力画面で列として
  // 表示するかどうかは、事務局ではなく2次入力者が画面上で選ぶ。
  const fileNameIdSet = new Set(
    reachableSingles
      .map(t => CoreLogic.cellId(t.cells[0]))
      .filter(id => CURRENT.fileNameFieldIds && CURRENT.fileNameFieldIds.has(id))
  );
  const displayCandidateEl = buildGenericFieldPicker(reachableSingles, {
    idPrefix: 'displaycandidatefield', storeKey: 'displayCandidateFieldIds', title: '2次入力画面表示項目',
    noteText: '2次入力画面（複数のデータを一覧表で確認する画面）で、見出し列として表示してよい項目を選べます（複数可）。ファイル名に含める項目は自動的に候補へ含まれます。実際にどれを表示するかは2次入力者が画面上で選びます。',
    forceCheckedIds: fileNameIdSet,
  });
  box.appendChild(buildStepContentWrap('step4', [requiredFieldEl, displayCandidateEl, secondaryFieldEl]));

  // STEP5：書き出す。
  box.appendChild(buildStepLabel('step5', 4));
  box.appendChild(el('p', { class: 'step-memo', text: '💡 ここまでの設定をもとに、入力フォーム・集約ツールを書き出します。' }));

  const titleRow = el('div', { class: 'form-title-row' });
  titleRow.appendChild(el('label', { for: 'form-title-input', text: 'フォームのタイトル（入力フォームの見出し・この様式ファイル自体のファイル名に使われます）' }));
  const titleInput = el('input', { type: 'text', id: 'form-title-input', value: effectiveFormTitle() });
  titleInput.addEventListener('input', () => { CURRENT.customTitle = titleInput.value; });
  titleRow.appendChild(titleInput);

  const fileNameFieldEl = buildFileNameFieldPicker(singles.filter(isReachable));

  const step5Content = [titleRow, fileNameFieldEl];
  if (unreachable.length === 0) {
    step5Content.push(el('p', { class: 'm-note', text: '✅ 現在の設定で、入力欄に入力した内容はすべて書き出されます。問題は見つかりませんでした。' }));
  }
  const btnRow = el('div', { class: 'm-btnrow' });
  if (unreachable.length === 0) {
    const okBtn = el('button', { text: '🌐 書き出す', type: 'button' });
    okBtn.addEventListener('click', doExportAsBoth);
    btnRow.appendChild(okBtn);
  } else {
    const forceBtn = el('button', { class: 'secondary', text: 'このまま書き出す（対象セルは空欄扱いになります）', type: 'button' });
    forceBtn.addEventListener('click', doExportAsBoth);
    btnRow.appendChild(forceBtn);
  }
  step5Content.push(btnRow);
  box.appendChild(buildStepContentWrap('step5', step5Content));

  // このパネルの主目的が「データ構造の確認」なので、プレビューは折りたたまずに常時表示する
  // （旧仕様では<details>で開閉式だったが、開く一手間自体をなくす）。
  const previewSection = el('div', { class: 'export-preview' });
  previewSection.appendChild(el('h4', { text: '現在の設定で書き出されるデータ構造' }));
  const previewTablesEl = buildPreviewTables(output);
  previewSection.appendChild(previewTablesEl);
  box.appendChild(previewSection);
  previewTablesEl.scrollTop = savedPreviewScrollTop;

  panel.scrollTop = savedScrollTop;
}

// ---------------------------------------------------------------------------
// ボタン配線
// ---------------------------------------------------------------------------

// .xlsx／.html（このビルダーで書き出した入力フォーム）どちらもここで受け付け、
// 拡張子だけを見て適切な読み込み処理に振り分ける（1つのドロップゾーンに統合）。
function handleAnyFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith('.xlsx')) {
    handleFile(file);
  } else if (name.endsWith('.html') || name.endsWith('.htm')) {
    handleFormHtmlFile(file);
  } else {
    setStatus('.xlsxファイル、またはこのビルダーで書き出した入力フォーム（.html）を選択してください。');
  }
}

function init() {
  renderStepUi();

  // 書き出し先フォルダの指定（Chromium限定）。ステータス表示と同じく、STEP0/STEP4の
  // どちらの書き出しにも関わる「STEP横断の共通機能」のため、特定のSTEPの中に置かず
  // 画面上部に固定表示する（filler_app.js側の#export-dir-barと同じ設計に統一した）。
  // 非対応ブラウザでは表示しない（インラインstyle.displayとスタイルシートのdisplay:noneが
  // 競合しないよう、明示的に'flex'をセットする必要がある）。
  if (window.showDirectoryPicker) {
    $('#export-dir-bar').style.display = 'flex';
  }
  $('#btn-pick-export-dir').addEventListener('click', pickExportDirectory);

  $('#file-input').addEventListener('change', (ev) => {
    if (ev.target.files[0]) handleAnyFile(ev.target.files[0]);
  });

  const dropZone = $('#drop-zone');
  dropZone.addEventListener('dragover', (ev) => { ev.preventDefault(); dropZone.classList.add('drag'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag'));
  dropZone.addEventListener('drop', (ev) => {
    ev.preventDefault();
    dropZone.classList.remove('drag');
    if (ev.dataTransfer.files[0]) handleAnyFile(ev.dataTransfer.files[0]);
  });

  // 「HTMLで書き出す」ボタンはグリッド側には置かず、常時表示の書き出し前チェックパネル
  // （renderCheckPanel内）の方に一本化してある。

  setupSelectionHandlers();
}

if (typeof window !== 'undefined') {
  window.addEventListener('DOMContentLoaded', init);
  window.__app = {
    handleFile, handleFormHtmlFile, handleAnyFile, extractStructureFromFillerHtml, rebuildAndRender, init,
    openCellSettings, closeModal, doExportAsForm, doExportAsAggregator, doExportAsBoth, serializeStructure,
    openManualGroupModal, selectionRowRange, renderCheckPanel,
    selectRowRangeAllColumns, summarizeUnreachableRanges, autoBlockUnreachableCells,
    saveBlob, pickExportDirectory,
    get OVERRIDES() { return OVERRIDES; },
    get SELECTED() { return SELECTED; },
    get CURRENT() { return CURRENT; },
    get EXPORT_DIR_HANDLE() { return EXPORT_DIR_HANDLE; },
    setExportDirHandle: (h) => { EXPORT_DIR_HANDLE = h; }, // テスト用（実ブラウザではpickExportDirectory経由）
  };
}
