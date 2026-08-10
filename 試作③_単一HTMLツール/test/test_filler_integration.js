// 入力フォーム（ツール2・filler_app.js）の結合テスト。
// ビルダーが書き出す入力フォームHTMLと同じ構成（core_logic.js + grid_render.js +
// `const STRUCTURE = {...};` + filler_app.js）でjsdomにページを作り、
// window.__app（filler_app.jsが公開するデバッグ用API）越しに検証する。

const path = require('path');
const fs = require('fs');
const { JSDOM } = require('jsdom');
const { runSuiteAsync, testAsync, assertEqual, assertTrue, assertFalse, summary } = require('./assert_mini');

const ROOT = path.join(__dirname, '..');
const SRC = {
  core: fs.readFileSync(path.join(ROOT, 'core_logic.js'), 'utf8'),
  grid: fs.readFileSync(path.join(ROOT, 'grid_render.js'), 'utf8'),
  filler: fs.readFileSync(path.join(ROOT, 'filler_app.js'), 'utf8'),
};
const FIXTURE = path.join(ROOT, '実機確認', '公開テンプレ調達', 'simple_moshikomi.xlsx');

function waitFor(predicate, timeoutMs, intervalMs) {
  timeoutMs = timeoutMs || 3000;
  intervalMs = intervalMs || 10;
  return new Promise((resolve, reject) => {
    const start = Date.now();
    (function check() {
      let ok = false;
      try { ok = !!predicate(); } catch (e) { ok = false; }
      if (ok) return resolve();
      if (Date.now() - start > timeoutMs) return reject(new Error('waitFor timeout'));
      setTimeout(check, intervalMs);
    })();
  });
}

// serializeStructure()（builder_app.js）と同じ形をNode単体で（DOM無しで）再現する。
// filler_app.jsが実際に受け取るデータ形そのままなので、フィクスチャからそのまま作れる。
function buildStructureFromFixture(fixturePath, formTitle) {
  const XLSX = require(path.join(ROOT, 'vendor', 'xlsx.core.min.js'));
  const CoreLogic = require(path.join(ROOT, 'core_logic.js'));
  const buf = fs.readFileSync(fixturePath);
  const wb = XLSX.read(buf, { type: 'buffer', cellStyles: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const { grid, maxRow, maxCol } = CoreLogic.buildGrid(ws);
  const sections = CoreLogic.splitSections(grid, maxRow, maxCol);
  const seen = new Set();
  const cells = [];
  for (const info of grid.values()) {
    const key = info.row + ',' + info.col;
    if (seen.has(key)) continue;
    seen.add(key);
    cells.push({
      row: info.row, col: info.col, row2: info.row2, col2: info.col2,
      value: info.value, isFormula: info.isFormula, formula: info.formula,
      hasText: info.hasText, blocked: info.blocked, fillColor: info.fillColor,
      renderType: info.renderType, renderOptions: info.renderOptions, dbKey: info.dbKey,
    });
  }
  return {
    formTitle: formTitle || 'テスト入力フォーム',
    maxRow, maxCol, widths: [], heights: [],
    sections: sections.map(s => ({ title: s.title, row0: s.row0, row1: s.row1 })),
    cells, manualGroups: [],
  };
}

// 「自前セルを持たない行でも<tr>が生成されるか」の回帰確認用。core_logic.jsの
// buildGridFromCells（フォームHTML再読込のときに使う経路）を経由させたいので、
// buildGrid由来ではなくbuildGridFromCells互換の生セル配列を直接組み立てる。
function buildRowspanRegressionStructure() {
  return {
    formTitle: 'rowspanテスト',
    maxRow: 4, maxCol: 2, widths: [], heights: [],
    sections: [{ title: 'シート', row0: 1, row1: 4 }],
    manualGroups: [],
    cells: [
      { row: 1, col: 1, row2: 4, col2: 1, value: '成果指標', isFormula: false, formula: '', hasText: true, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null },
      { row: 1, col: 2, row2: 1, col2: 2, value: '', isFormula: false, formula: '', hasText: false, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null },
      { row: 2, col: 2, row2: 4, col2: 2, value: '達成状況', isFormula: false, formula: '', hasText: true, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null },
    ],
  };
}

// labelForCellIdのautoNameフォールバック確認用。「1行1見出し1値」（A列に見出し、
// B列に値セルが1つだけ）という、dbKey未設定だと本来の項目名が拾えなくなっていた
// パターンをそのまま再現する（core_logic.jsのbuildRowEntryコメント参照）。
function buildSingleLabelRowStructure() {
  return {
    formTitle: 'ラベルテスト',
    maxRow: 1, maxCol: 2, widths: [], heights: [],
    sections: [{ title: 'シート', row0: 1, row1: 1 }],
    manualGroups: [],
    cells: [
      { row: 1, col: 1, row2: 1, col2: 1, value: '部署名', isFormula: false, formula: '', hasText: true, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null },
      { row: 1, col: 2, row2: 1, col2: 2, value: '', isFormula: false, formula: '', hasText: false, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null },
    ],
  };
}

// 実際のfiller_template.htmlと同じID体系を最小限そろえたfixture。レビュー画面用の
// 要素（#btn-open-review等）もinit()がaddEventListenerで参照するため、実テンプレートと
// 乖離しないよう一通り含めておく（過去に⚙アイコン追加でテストのDOM前提が崩れた教訓と同型）。
function newFillerPage(structure) {
  const html = `<!doctype html><html><body>
    <div id="export-dir-bar">
      <button id="btn-pick-export-dir">pickdir</button>
      <span id="export-dir-status"></span>
    </div>
    <div id="actions">
      <button id="btn-export">export</button>
      <input type="file" id="file-load-json">
      <button id="btn-open-review" style="display:none">review</button>
      <button id="btn-print-grid">print</button>
    </div>
    <div id="status"></div>
    <div id="grid-root"></div>
    <p id="grid-hint"></p>
    <p id="field-legend" style="display:none"></p>
    <div id="review-root" style="display:none">
      <div id="review-actions">
        <input type="file" id="review-file-load" multiple>
        <button id="review-btn-pick-dir" style="display:none">dir</button>
        <button id="review-btn-refresh" style="display:none">refresh</button>
        <button id="btn-close-review">close</button>
      </div>
      <div id="review-col-toggle"></div>
      <select id="review-status-select">
        <option value="all">all</option>
        <option value="undone">undone</option>
        <option value="done">done</option>
      </select>
      <div id="review-summary"></div>
      <button id="btn-bulk-print">bulkprint</button>
      <div id="review-table-root"></div>
    </div>
    <div id="review-detail-root-wrap" style="display:none">
      <button id="btn-save-detail">save</button>
      <button id="btn-back-to-review-list">back</button>
      <button id="btn-print-detail">print</button>
      <div id="review-detail-root"></div>
    </div>
    <div id="review-scratch-root" style="display:none"></div>
    <div id="bulk-print-root"></div>
    <script>${SRC.core}</script>
    <script>${SRC.grid}</script>
    <script>
      const STRUCTURE = ${JSON.stringify(structure)};
      ${SRC.filler}
    </script>
  </body></html>`;
  return new JSDOM(html, { runScripts: 'dangerously', url: 'http://localhost/' });
}

async function readyPage(structure) {
  const dom = newFillerPage(structure);
  await waitFor(() => dom.window.__app && dom.window.__app.STATE);
  return dom;
}

// レビュー画面のテスト用に、「あるページの通常グリッドに値を入れてcollectData()する」ことで
// 妥当なJSON（所管部署が書き出した想定のデータ）を作る。手でネスト構造を書き下ろすと
// buildSectionObjectの実際の出力と食い違うリスクがあるため、実ロジックに作らせる。
function buildSampleData(dom, values) {
  const win = dom.window;
  Object.keys(values).forEach((id) => {
    const elx = win.document.getElementById(id);
    if (elx) elx.value = values[id];
  });
  return win.__app.collectData();
}

(async () => {
  await runSuiteAsync('filler_app: STRUCTUREからのグリッド復元', async () => {
    await testAsync('実フィクスチャ由来のSTRUCTUREから、ビルダーと同じ行・列数でグリッドが描画される', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const { STATE } = dom.window.__app;
      assertEqual(STATE.maxRow, structure.maxRow);
      assertEqual(STATE.maxCol, structure.maxCol);
      assertEqual(dom.window.document.querySelectorAll('#grid-root tbody tr').length, structure.maxRow);
    });

    await testAsync('自前セルを持たない行でも<tr>が省略されない（buildGridFromCells経由でも回帰しない）', async () => {
      const structure = buildRowspanRegressionStructure();
      const dom = await readyPage(structure);
      const trs = dom.window.document.querySelectorAll('#grid-root tbody tr');
      assertEqual(trs.length, 4);
    });
  });

  await runSuiteAsync('filler_app: 入力→書き出し→読込の往復', async () => {
    await testAsync('入力した内容がcollectData()に反映され、loadDataIntoGrid()で復元できる', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      const { STATE } = win.__app;
      const c3 = STATE.grid.get('3,3');
      const input = win.document.getElementById(win.CoreLogic.cellId(c3));
      input.value = 'テスト太郎';
      const data = win.__app.collectData();
      assertEqual(data['シート']['申込日']['col3'], 'テスト太郎');

      input.value = '';
      win.__app.loadDataIntoGrid(data);
      assertEqual(win.document.getElementById(win.CoreLogic.cellId(c3)).value, 'テスト太郎');
    });

    await testAsync('書き出しボタンでJSONダウンロードがトリガーされ、ステータスが更新される（例外を投げない）', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      win.URL.createObjectURL = () => 'blob:mock';
      win.HTMLAnchorElement.prototype.click = function () {};
      win.document.getElementById('btn-export').click();
      assertEqual(win.document.getElementById('status').textContent, '書き出しました。');
    });
  });

  await runSuiteAsync('filler_app: 書き出しファイル名の組み立て（buildExportFileName）', async () => {
    await testAsync('fileNameFields未指定なら「タイトル_日付(YYYYMMDD).json」になる', async () => {
      const structure = buildStructureFromFixture(FIXTURE); // formTitle既定値「テスト入力フォーム」
      const dom = await readyPage(structure);
      const win = dom.window;
      const name = win.__app.buildExportFileName();
      assertTrue(/^テスト入力フォーム_\d{8}\.json$/.test(name), `タイトル＋YYYYMMDD形式のはず: "${name}"`);
    });

    await testAsync('fileNameFieldsで指定した項目の入力値が、タイトルの後ろに連結される', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3'; // 「申込日」col3（buildStructureFromFixtureの申込日行に実在するセル）
      structure.fileNameFields = [c3id];
      const dom = await readyPage(structure);
      const win = dom.window;
      win.document.getElementById(c3id).value = 'テスト事業';
      const name = win.__app.buildExportFileName();
      assertTrue(name.startsWith('テスト入力フォーム_テスト事業_'), `タイトルの後ろに項目の値が連結されるはず: "${name}"`);
      assertTrue(/_\d{8}\.json$/.test(name), `末尾はYYYYMMDD.jsonのはず: "${name}"`);
    });

    await testAsync('指定した項目が複数あれば、シート順（fileNameFieldsの並び順）のままタイトルの後ろに連結される', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.fileNameFields = [c3id, c4id];
      const dom = await readyPage(structure);
      const win = dom.window;
      win.document.getElementById(c3id).value = '事業A';
      win.document.getElementById(c4id).value = '担当課B';
      const name = win.__app.buildExportFileName();
      assertTrue(name.startsWith('テスト入力フォーム_事業A_担当課B_'), `タイトル→指定順の項目の順で連結されるはず: "${name}"`);
    });

    await testAsync('指定した項目が空欄なら、その項目は連結対象から除かれる（他の項目・タイトルは影響しない）', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.fileNameFields = [c3id, c4id];
      const dom = await readyPage(structure);
      const win = dom.window;
      win.document.getElementById(c3id).value = '';
      win.document.getElementById(c4id).value = '担当課B';
      const name = win.__app.buildExportFileName();
      assertTrue(name.startsWith('テスト入力フォーム_担当課B_'), `空欄の項目はスキップされ、タイトル＋値がある項目だけ使われるはず: "${name}"`);
    });

    await testAsync('sanitizeForFileNameはファイル名に使えない文字・改行を除去し、長さを切り詰める', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      assertEqual(win.__app.sanitizeForFileName('a/b:c*d?e"f<g>h|i'), 'abcdefghi');
      assertEqual(win.__app.sanitizeForFileName('改行\nタブ\t混在'), '改行タブ混在');
      assertEqual(win.__app.sanitizeForFileName('あ'.repeat(50)).length, 40);
    });
  });

  await runSuiteAsync('filler_app: 列見出しラベル（labelForCellId）', async () => {
    await testAsync('「1行1見出し1値」でdbKey未設定なら、生セル番地ではなく行見出し文字が返る', async () => {
      const structure = buildSingleLabelRowStructure();
      const dom = await readyPage(structure);
      const win = dom.window;
      const label = win.__app.labelForCellId('cell_R1_C2');
      assertEqual(label, '部署名', 'B1のcellRef「B1」ではなく、A1の見出し文字が返るはず');
    });

    await testAsync('dbKeyが設定されていれば、行見出し文字よりdbKeyが優先される', async () => {
      const structure = buildSingleLabelRowStructure();
      structure.cells[1].dbKey = 'department';
      const dom = await readyPage(structure);
      const win = dom.window;
      const label = win.__app.labelForCellId('cell_R1_C2');
      assertEqual(label, 'department');
    });

    await testAsync('該当セルが存在しないidを渡すと、そのidがそのまま返る', async () => {
      const structure = buildSingleLabelRowStructure();
      const dom = await readyPage(structure);
      const win = dom.window;
      assertEqual(win.__app.labelForCellId('cell_R99_C99'), 'cell_R99_C99');
    });
  });

  await runSuiteAsync('filler_app: 1次/2次入力欄の視覚的色分け', async () => {
    await testAsync('reviewFields未指定なら、どの入力セルにも色分けクラスが付かない', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      const tds = win.document.querySelectorAll('#grid-root td.cell-input');
      assertTrue(tds.length > 0);
      tds.forEach(td => {
        assertFalse(td.classList.contains('cell-field-primary'));
        assertFalse(td.classList.contains('cell-field-review'));
      });
    });

    await testAsync('reviewFields指定時、レビュー欄のセルはcell-field-review、それ以外の入力セルはcell-field-primaryになる', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const reviewTd = win.document.getElementById(c4id).closest('td');
      assertTrue(reviewTd.classList.contains('cell-field-review'));
      assertFalse(reviewTd.classList.contains('cell-field-primary'));

      const c3id = 'cell_R3_C3';
      const primaryTd = win.document.getElementById(c3id).closest('td');
      assertTrue(primaryTd.classList.contains('cell-field-primary'));
      assertFalse(primaryTd.classList.contains('cell-field-review'));
    });

    await testAsync('reviewFields指定時、凡例（#field-legend）が表示される。未指定なら非表示のまま', async () => {
      const withReview = await readyPage((() => { const s = buildStructureFromFixture(FIXTURE); s.reviewFields = ['cell_R3_C4']; return s; })());
      assertEqual(withReview.window.document.getElementById('field-legend').style.display, '');

      const withoutReview = await readyPage(buildStructureFromFixture(FIXTURE));
      assertEqual(withoutReview.window.document.getElementById('field-legend').style.display, 'none');
    });

    await testAsync('レビュー詳細画面（openReviewDetail）でも同じ色分けが適用される', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      structure.displayCandidateFields = [c3id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const data1 = buildSampleData(dom, { [c3id]: '事業A', [c4id]: '' });
      win.__app.upsertRecords([{ fileName: 'a.json', data: data1 }]);
      win.__app.openReviewDetail(win.__app.REVIEW.records[0]);
      const reviewTd = win.document.getElementById(c4id).closest('td');
      assertTrue(reviewTd.classList.contains('cell-field-review'), '詳細画面でもレビュー欄セルにcell-field-reviewが付くはず');
    });
  });

  await runSuiteAsync('filler_app: 印刷ボタン', async () => {
    await testAsync('通常画面の印刷ボタン（#btn-print-grid）をクリックしても例外を投げない', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      // jsdomはwindow.printを実装していないため、no-opで差し替えて呼び出しの配線だけ検証する
      let called = false;
      win.print = () => { called = true; };
      win.document.getElementById('btn-print-grid').click();
      assertTrue(called);
    });

    await testAsync('レビュー詳細画面の印刷ボタン（#btn-print-detail）をクリックしても例外を投げない', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const data1 = buildSampleData(dom, { [c3id]: '事業A', [c4id]: '' });
      win.__app.upsertRecords([{ fileName: 'a.json', data: data1 }]);
      win.__app.openReviewDetail(win.__app.REVIEW.records[0]);
      let called = false;
      win.print = () => { called = true; };
      win.document.getElementById('btn-print-detail').click();
      assertTrue(called);
    });
  });

  await runSuiteAsync('filler_app: 貼り付け', async () => {
    await testAsync('setupPasteHandlerが有効になっており、タブ区切りの貼り付けで複数セルが埋まる', async () => {
      // 貼り付けハンドラはタブ/改行を含まない単一値は無視する設計
      // （grid_render.jsのsetupPasteHandler参照）なので、タブ区切りで検証する。
      const structure = buildRowspanRegressionStructure();
      const dom = await readyPage(structure);
      const win = dom.window;
      const { STATE } = win.__app;
      const startCell = win.document.getElementById(win.CoreLogic.cellId(STATE.grid.get('1,2')));
      startCell.focus();
      const ev = new win.Event('paste', { bubbles: true, cancelable: true });
      ev.clipboardData = { getData: () => '入力値\t続き' };
      win.document.dispatchEvent(ev);
      assertEqual(startCell.value, '入力値');
    });
  });

  await runSuiteAsync('filler_app: 必須項目の未入力チェック', async () => {
    await testAsync('requiredFields未指定なら、未入力でも確認なしで書き出される', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      win.URL.createObjectURL = () => 'blob:mock';
      win.HTMLAnchorElement.prototype.click = function () {};
      let confirmCalled = false;
      win.confirm = () => { confirmCalled = true; return true; };
      win.document.getElementById('btn-export').click();
      assertFalse(confirmCalled, 'requiredFieldsが無ければconfirmは呼ばれないはず');
      assertEqual(win.document.getElementById('status').textContent, '書き出しました。');
    });

    await testAsync('必須項目が未入力なら確認ダイアログが出て、キャンセルすると書き出されない', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      structure.requiredFields = ['cell_R3_C3'];
      const dom = await readyPage(structure);
      const win = dom.window;
      win.URL.createObjectURL = () => 'blob:mock';
      let clicked = false;
      win.HTMLAnchorElement.prototype.click = function () { clicked = true; };
      win.confirm = () => false;
      win.document.getElementById('btn-export').click();
      assertFalse(clicked, 'キャンセルした場合は書き出されないはず');
      assertEqual(win.document.getElementById('status').textContent, '書き出しを中止しました。未入力の必須項目を確認してください。');
    });

    await testAsync('必須項目が未入力でも確認ダイアログでOKすれば書き出される', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      structure.requiredFields = ['cell_R3_C3'];
      const dom = await readyPage(structure);
      const win = dom.window;
      win.URL.createObjectURL = () => 'blob:mock';
      let clicked = false;
      win.HTMLAnchorElement.prototype.click = function () { clicked = true; };
      win.confirm = () => true;
      win.document.getElementById('btn-export').click();
      assertTrue(clicked);
      assertEqual(win.document.getElementById('status').textContent, '書き出しました。');
    });

    await testAsync('必須項目に値が入っていれば確認ダイアログは出ない', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3';
      structure.requiredFields = [c3id];
      const dom = await readyPage(structure);
      const win = dom.window;
      win.document.getElementById(c3id).value = '値あり';
      win.URL.createObjectURL = () => 'blob:mock';
      win.HTMLAnchorElement.prototype.click = function () {};
      let confirmCalled = false;
      win.confirm = () => { confirmCalled = true; return true; };
      win.document.getElementById('btn-export').click();
      assertFalse(confirmCalled);
    });
  });

  await runSuiteAsync('filler_app: レビュー画面（複数JSON読込・一覧表・詳細）', async () => {
    await testAsync('レビュー欄が指定されていれば「2次以降用一括入力画面へ移動する」ボタンが表示される', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      structure.reviewFields = ['cell_R3_C3'];
      const dom = await readyPage(structure);
      assertEqual(dom.window.document.getElementById('btn-open-review').style.display, '');
    });

    await testAsync('レビュー欄が無指定なら「2次以降用一括入力画面へ移動する」ボタンは表示されない', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      assertEqual(dom.window.document.getElementById('btn-open-review').style.display, 'none');
    });

    await testAsync('複数JSONを読み込むと件数分のレコードが追加され、同名ファイルは重複追加されない', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      structure.displayCandidateFields = [c3id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const data1 = buildSampleData(dom, { [c3id]: '事業A', [c4id]: '' });
      const data2 = buildSampleData(dom, { [c3id]: '事業B', [c4id]: '' });

      let added = win.__app.upsertRecords([{ fileName: 'a.json', data: data1 }, { fileName: 'b.json', data: data2 }]);
      assertEqual(added, 2);
      assertEqual(win.__app.REVIEW.records.length, 2);

      added = win.__app.upsertRecords([{ fileName: 'a.json', data: data1 }]);
      assertEqual(added, 0, '同名ファイルは既存レコードを上書きしないはず');
      assertEqual(win.__app.REVIEW.records.length, 2);
    });

    await testAsync('displayValues/reviewValuesが元データから正しく抽出される', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      structure.displayCandidateFields = [c3id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const data1 = buildSampleData(dom, { [c3id]: '事業A', [c4id]: '判定OK' });
      win.__app.upsertRecords([{ fileName: 'a.json', data: data1 }]);
      const record = win.__app.REVIEW.records[0];
      assertEqual(record.displayValues[c3id], '事業A');
      assertEqual(record.reviewValues[c4id], '判定OK');
    });

    await testAsync('isRecordComplete：レビュー欄が全て埋まっていれば完了、1つでも空なら未完了', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c3id, c4id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const complete = { fileName: 'x.json', data: {}, reviewValues: { [c3id]: '判定', [c4id]: 'コメント' } };
      const incomplete = { fileName: 'y.json', data: {}, reviewValues: { [c3id]: '判定', [c4id]: '' } };
      assertTrue(win.__app.isRecordComplete(complete));
      assertFalse(win.__app.isRecordComplete(incomplete));
    });

    await testAsync('一覧表：入力すると record.reviewValues と状態表示（未完了→完了）が更新される', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      structure.displayCandidateFields = [c3id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const data1 = buildSampleData(dom, { [c3id]: '事業A', [c4id]: '' });
      win.__app.upsertRecords([{ fileName: 'a.json', data: data1 }]);
      win.__app.enterReviewMode();

      const rows = win.document.querySelectorAll('#review-table-root tbody tr');
      assertEqual(rows.length, 1);
      assertEqual(rows[0].querySelector('.review-status-cell').textContent, '未完了');

      const textarea = rows[0].querySelector('textarea');
      textarea.value = '判定OK';
      textarea.dispatchEvent(new win.Event('input', { bubbles: true }));

      assertEqual(win.__app.REVIEW.records[0].reviewValues[c4id], '判定OK');
      assertEqual(rows[0].querySelector('.review-status-cell').textContent, '✅完了');
    });

    await testAsync('列表示トグルを外すと、一覧表の列数（ヘッダー・行とも）が1つ減る', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      structure.displayCandidateFields = [c3id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const data1 = buildSampleData(dom, { [c3id]: '事業A', [c4id]: '' });
      win.__app.upsertRecords([{ fileName: 'a.json', data: data1 }]);
      win.__app.enterReviewMode();

      const beforeCount = win.document.querySelectorAll('#review-table-root thead th').length;
      win.document.getElementById('colvis_' + c3id).click();
      const afterCount = win.document.querySelectorAll('#review-table-root thead th').length;
      assertEqual(afterCount, beforeCount - 1, '表示候補列を1つ外したのでヘッダーが1列減るはず');
      const bodyRowCells = win.document.querySelectorAll('#review-table-root tbody tr')[0].children.length;
      assertEqual(bodyRowCells, afterCount, 'ヘッダーと行のセル数は一致するはず');
    });

    await testAsync('詳細画面は所管部署の入力内容を表示し、編集可能（全セル誰でも入力できる方針のため読み取り専用ではない）', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      structure.displayCandidateFields = [c3id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const data1 = buildSampleData(dom, { [c3id]: '事業A', [c4id]: '' });
      win.__app.upsertRecords([{ fileName: 'a.json', data: data1 }]);
      win.__app.enterReviewMode();
      assertEqual(win.document.getElementById('grid-root').innerHTML, '', 'レビュー画面ではgrid-rootは空になっているはず');

      win.__app.openReviewDetail(win.__app.REVIEW.records[0]);
      assertEqual(win.document.getElementById('review-detail-root-wrap').style.display, '');
      const detailInput = win.document.getElementById(c3id);
      assertTrue(!!detailInput, '詳細画面にも同じcellIdの入力欄が描画されるはず');
      assertEqual(detailInput.value, '事業A');
      assertFalse(detailInput.disabled, '全セル編集可能の方針のため、詳細画面もdisabledではないはず');
    });

    await testAsync('詳細画面で保存せずに戻ると（backToReviewListFromDetail）、レコードの内容は変更されない', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      structure.displayCandidateFields = [c3id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const data1 = buildSampleData(dom, { [c3id]: '事業A', [c4id]: '' });
      win.__app.upsertRecords([{ fileName: 'a.json', data: data1 }]);
      win.__app.openReviewDetail(win.__app.REVIEW.records[0]);
      win.document.getElementById(c3id).value = '編集したが保存しない';

      win.__app.backToReviewListFromDetail();
      assertEqual(win.document.getElementById('review-detail-root').innerHTML, '');
      assertEqual(win.document.getElementById('review-root').style.display, '');
      assertEqual(win.__app.REVIEW.records[0].displayValues[c3id], '事業A', '保存しなかったので元の値のまま残るはず');
    });

    await testAsync('詳細画面で編集して保存すると（saveDetailAndBackToList）、record.data/displayValues/reviewValuesに反映される', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      structure.displayCandidateFields = [c3id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const data1 = buildSampleData(dom, { [c3id]: '事業A', [c4id]: '' });
      win.__app.upsertRecords([{ fileName: 'a.json', data: data1 }]);
      win.__app.openReviewDetail(win.__app.REVIEW.records[0]);
      win.document.getElementById(c3id).value = '事業A（詳細画面で修正）';
      win.document.getElementById(c4id).value = '詳細画面から入力したレビュー結果';

      win.__app.saveDetailAndBackToList();
      assertEqual(win.document.getElementById('review-detail-root-wrap').style.display, 'none', '保存後は一覧画面に戻るはず');
      const record = win.__app.REVIEW.records[0];
      assertEqual(record.displayValues[c3id], '事業A（詳細画面で修正）');
      assertEqual(record.reviewValues[c4id], '詳細画面から入力したレビュー結果');
      assertEqual(record.data['シート']['申込日']['col3'], '事業A（詳細画面で修正）', 'record.data自体も更新されるはず');
      assertTrue(win.document.getElementById('status').textContent.includes('保存しました'));
    });

    await testAsync('レビュー画面を閉じると、通常の入力グリッドが再描画される', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      structure.reviewFields = ['cell_R3_C4'];
      const dom = await readyPage(structure);
      const win = dom.window;
      win.__app.enterReviewMode();
      assertEqual(win.document.getElementById('grid-root').innerHTML, '');
      win.__app.exitReviewMode();
      assertTrue(win.document.getElementById('grid-root').innerHTML.length > 0, '通常画面に戻るとgrid-rootが再描画されるはず');
      assertEqual(win.document.getElementById('actions').style.display, 'flex');
    });

    await testAsync('mergedDataForExportは、所管部署の元データにレビュー欄の現在値をマージする', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const data1 = buildSampleData(dom, { [c3id]: '事業A', [c4id]: '' });
      win.__app.upsertRecords([{ fileName: 'a.json', data: data1 }]);
      const record = win.__app.REVIEW.records[0];
      record.reviewValues[c4id] = 'レビュー結果';
      const merged = win.__app.mergedDataForExport(record);
      assertEqual(merged['シート']['申込日']['col3'], '事業A', '所管部署の元データは維持されるはず');
      assertEqual(merged['シート']['申込日']['col4'], 'レビュー結果', 'レビュー欄の現在値がマージされるはず');
    });

    await testAsync('buildExportFileNameForRecordは、fileNameFieldsの値（displayValues経由）を使ってファイル名を組み立てる', async () => {
      const structure = buildStructureFromFixture(FIXTURE, 'テストタイトル');
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      structure.fileNameFields = [c3id];
      structure.displayCandidateFields = [c3id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const data1 = buildSampleData(dom, { [c3id]: '事業A', [c4id]: '' });
      win.__app.upsertRecords([{ fileName: 'a.json', data: data1 }]);
      const name = win.__app.buildExportFileNameForRecord(win.__app.REVIEW.records[0]);
      assertTrue(name.startsWith('テストタイトル_事業A_'), `タイトル＋事業名で始まるはず: "${name}"`);
    });

    await testAsync('複数ファイル選択（review-file-load）からの読込で、レコードが追加される', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      structure.displayCandidateFields = [c3id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const data1 = buildSampleData(dom, { [c3id]: '事業A', [c4id]: '' });
      const file = new win.File([JSON.stringify(data1)], 'a.json', { type: 'application/json' });
      const input = win.document.getElementById('review-file-load');
      Object.defineProperty(input, 'files', { value: [file], configurable: true });
      input.dispatchEvent(new win.Event('change'));
      await waitFor(() => win.__app.REVIEW.records.length === 1);
      assertEqual(win.__app.REVIEW.records[0].fileName, 'a.json');
    });
  });

  await runSuiteAsync('filler_app: 一覧の絞り込み（Excel風フィルタ・ステータス）', async () => {
    // 3件（部署A×2・部署B×1、うち1件はレビュー欄記入済み＝完了）を読み込んだ状態を作る共通処理。
    async function setupThreeRecords() {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      structure.displayCandidateFields = [c3id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const dataA1 = buildSampleData(dom, { [c3id]: '部署A', [c4id]: '' });
      const dataA2 = buildSampleData(dom, { [c3id]: '部署A', [c4id]: '' });
      const dataB1 = buildSampleData(dom, { [c3id]: '部署B', [c4id]: '' });
      win.__app.upsertRecords([
        { fileName: 'a1.json', data: dataA1 },
        { fileName: 'a2.json', data: dataA2 },
        { fileName: 'b1.json', data: dataB1 },
      ]);
      win.__app.REVIEW.records.find(r => r.fileName === 'a1.json').reviewValues[c4id] = '判定済み';
      win.__app.enterReviewMode();
      return { win, c3id, c4id };
    }

    await testAsync('絞り込みなし（初期状態）では全件が表示される', async () => {
      const { win } = await setupThreeRecords();
      const rows = win.document.querySelectorAll('#review-tbody tr');
      assertEqual(rows.length, 3);
      assertEqual(win.document.getElementById('review-summary').textContent, '3件中 1件 レビュー完了');
    });

    await testAsync('列フィルタ：ある値のチェックを外すと、その値を持つ行だけ一覧から消える（他の値は残る）', async () => {
      const { win, c3id } = await setupThreeRecords();
      const cb = [...win.document.querySelectorAll('.col-filter-panel label')]
        .find(l => l.textContent === '部署A').querySelector('input');
      cb.checked = false;
      cb.dispatchEvent(new win.Event('change'));

      const rows = win.document.querySelectorAll('#review-tbody tr');
      assertEqual(rows.length, 1, '部署Aの2件が絞り込まれ、部署Bの1件だけ残るはず');
      assertTrue(rows[0].textContent.includes('部署B'));
      assertTrue(win.document.getElementById('review-summary').textContent.includes('絞り込み中'));
    });

    await testAsync('列フィルタのチェック変更では、thead（<details>の開閉状態含む）は再描画されない', async () => {
      const { win } = await setupThreeRecords();
      const details = win.document.querySelector('.col-filter');
      details.setAttribute('open', 'open');
      const cb = details.querySelector('input');
      cb.checked = false;
      cb.dispatchEvent(new win.Event('change'));
      assertTrue(win.document.querySelector('.col-filter').hasAttribute('open'), 'tbodyだけの再描画なら<details>のopen状態は保たれるはず');
    });

    await testAsync('ステータス絞り込み：「完了」を選ぶとレビュー欄が埋まっている行だけ表示される', async () => {
      const { win } = await setupThreeRecords();
      const select = win.document.getElementById('review-status-select');
      select.value = 'done';
      select.dispatchEvent(new win.Event('change'));
      const rows = win.document.querySelectorAll('#review-tbody tr');
      assertEqual(rows.length, 1);
      assertTrue(rows[0].textContent.includes('部署A'));
    });

    await testAsync('ステータス絞り込み：「未完了」を選ぶとレビュー欄が空欄の行だけ表示される', async () => {
      const { win } = await setupThreeRecords();
      const select = win.document.getElementById('review-status-select');
      select.value = 'undone';
      select.dispatchEvent(new win.Event('change'));
      const rows = win.document.querySelectorAll('#review-tbody tr');
      assertEqual(rows.length, 2);
    });

    await testAsync('列フィルタとステータス絞り込みはAND条件で組み合わさる', async () => {
      const { win } = await setupThreeRecords();
      const cb = [...win.document.querySelectorAll('.col-filter-panel label')]
        .find(l => l.textContent === '部署B').querySelector('input');
      cb.checked = false;
      cb.dispatchEvent(new win.Event('change')); // 部署Aの2件（1完了・1未完了）だけに絞る

      const select = win.document.getElementById('review-status-select');
      select.value = 'done';
      select.dispatchEvent(new win.Event('change'));

      const rows = win.document.querySelectorAll('#review-tbody tr');
      assertEqual(rows.length, 1, '部署A かつ 完了、の1件だけになるはず');
    });

    await testAsync('列の表示を非表示に切り替えると、その列のフィルタは無視される（再表示前提の絞り込みは効かない）', async () => {
      const { win, c3id } = await setupThreeRecords();
      const cb = [...win.document.querySelectorAll('.col-filter-panel label')]
        .find(l => l.textContent === '部署A').querySelector('input');
      cb.checked = false;
      cb.dispatchEvent(new win.Event('change'));
      assertEqual(win.document.querySelectorAll('#review-tbody tr').length, 1, '絞り込みが効いている前提の確認');

      win.document.getElementById('colvis_' + c3id).click(); // 列を非表示化（renderReviewTableが全体再描画）
      assertEqual(win.document.querySelectorAll('#review-tbody tr').length, 3, '列を隠すとその列のフィルタは無視され全件表示に戻るはず');
    });
  });

  await runSuiteAsync('filler_app: 一括印刷（絞り込み結果を対象に連続印刷）', async () => {
    async function setupTwoRecordsForPrint() {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      structure.displayCandidateFields = [c3id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const dataA = buildSampleData(dom, { [c3id]: '部署A', [c4id]: '判定OK' });
      const dataB = buildSampleData(dom, { [c3id]: '部署B', [c4id]: '' });
      win.__app.upsertRecords([{ fileName: 'a.json', data: dataA }, { fileName: 'b.json', data: dataB }]);
      win.__app.enterReviewMode();
      return { win, c3id, c4id };
    }

    await testAsync('絞り込みなしでbulkPrintFilteredを呼ぶと、全レコード分のスナップショットが#bulk-print-rootに追加され印刷が呼ばれる', async () => {
      const { win } = await setupTwoRecordsForPrint();
      let printed = false;
      win.print = () => { printed = true; };
      win.__app.bulkPrintFiltered();
      const sections = win.document.querySelectorAll('#bulk-print-root .bulk-print-record');
      assertEqual(sections.length, 2);
      assertTrue(printed);
      assertTrue(win.document.body.classList.contains('bulk-printing'));
    });

    await testAsync('ステータス絞り込み中にbulkPrintFilteredを呼ぶと、絞り込んだ件数分だけ印刷対象になる', async () => {
      const { win } = await setupTwoRecordsForPrint();
      win.document.getElementById('review-status-select').value = 'done';
      win.document.getElementById('review-status-select').dispatchEvent(new win.Event('change'));
      win.print = () => {};
      win.__app.bulkPrintFiltered();
      const sections = win.document.querySelectorAll('#bulk-print-root .bulk-print-record');
      assertEqual(sections.length, 1, 'レビュー完了の1件だけが一括印刷の対象になるはず');
      assertTrue(sections[0].textContent.includes('a.json'));
    });

    await testAsync('複数レコード分のスナップショットを並べても、クローンされたテーブルにid属性が残らない（重複IDを避ける）', async () => {
      const { win } = await setupTwoRecordsForPrint();
      win.print = () => {};
      win.__app.bulkPrintFiltered();
      const idsRemaining = win.document.querySelectorAll('#bulk-print-root [id]');
      assertEqual(idsRemaining.length, 0, 'id属性を持つ要素が1つも残っていないはず');
      const tables = win.document.querySelectorAll('#bulk-print-root table.excel-grid');
      assertEqual(tables.length, 2, '2レコード分のテーブルが両方描画されているはず');
    });

    await testAsync('cleanupBulkPrintを呼ぶと、bulk-printingクラスと#bulk-print-rootの中身が消える', async () => {
      const { win } = await setupTwoRecordsForPrint();
      win.print = () => {};
      win.__app.bulkPrintFiltered();
      win.__app.cleanupBulkPrint();
      assertFalse(win.document.body.classList.contains('bulk-printing'));
      assertEqual(win.document.getElementById('bulk-print-root').innerHTML, '');
    });

    await testAsync('afterprintイベントで自動的に後片付けされる', async () => {
      const { win } = await setupTwoRecordsForPrint();
      win.print = () => {};
      win.__app.bulkPrintFiltered();
      assertTrue(win.document.body.classList.contains('bulk-printing'));
      win.dispatchEvent(new win.Event('afterprint'));
      assertFalse(win.document.body.classList.contains('bulk-printing'), 'afterprintでcleanupBulkPrintが呼ばれるはず');
    });
  });

  await runSuiteAsync('filler_app: 書き出し先フォルダの指定（File System Access API）', async () => {
    // showDirectoryPicker等はjsdomに実装されていないため、EXPORT_DIR_HANDLE相当の
    // オブジェクトを直接注入して書き込みロジックだけを検証する（フォルダ選択ダイアログ
    // 自体の自動テストは不可能。実機（Edge/Chrome）での手動確認が必要）。
    // jsdomのBlobには.text()が実装されていないため、FileReaderで読み取る
    // （filler_app.js自身のreadFileAsTextと同じ手法）。
    function blobToText(win, blob) {
      return new Promise((resolve, reject) => {
        const reader = new win.FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsText(blob);
      });
    }
    function makeMockDirHandle(win, existingFileNames) {
      const files = new Set(existingFileNames || []);
      const written = {}; // filename -> 書き込まれたテキスト内容
      return {
        name: 'mock-folder',
        written,
        async getFileHandle(name, opts) {
          const exists = files.has(name);
          if (!exists && !(opts && opts.create)) {
            const err = new Error('not found'); err.name = 'NotFoundError'; throw err;
          }
          files.add(name);
          return {
            async createWritable() {
              return {
                async write(blob) { written[name] = await blobToText(win, blob); },
                async close() {},
              };
            },
          };
        },
      };
    }
    // getFileHandleが常に例外を投げる壊れたハンドル（権限切れ等のシミュレーション用）。
    function makeFailingDirHandle() {
      return {
        name: 'broken-folder',
        async getFileHandle(name, opts) {
          if (opts && opts.create) throw new Error('NotAllowedError');
          const err = new Error('not found'); err.name = 'NotFoundError'; throw err;
        },
      };
    }

    await testAsync('showDirectoryPicker非対応環境（jsdom既定）では、#export-dir-barのinline displayは変更しない（実テンプレートのCSS側のdisplay:noneに委ねる）', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      // fixtureにはfiller_template.htmlのCSS（#export-dir-bar{display:none;...}）を
      // 含めていないため、ここで確認できるのは「JSがinline styleを一切いじっていない」
      // ことだけ（＝実テンプレートのCSSのdisplay:noneがそのまま効く前提）。
      assertEqual(dom.window.document.getElementById('export-dir-bar').style.display, '');
    });

    await testAsync('showDirectoryPicker対応環境では、#export-dir-barのinline displayに明示的に"flex"がセットされる', async () => {
      // 過去バグの回帰テスト：#export-dir-barはCSS側でdisplay:noneをスタイルシート宣言
      // （インライン属性ではない）として持っているため、.style.display=''（インライン指定の
      // 除去）ではスタイルシートのdisplay:noneにフォールバックしてしまい表示されなかった。
      // 「非空文字列を明示的にセットしているか」を確認することで、この種の回帰を検出する。
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = newFillerPage(structure);
      dom.window.showDirectoryPicker = async () => ({ name: 'mock' }); // 存在確認のみされるため中身は不要
      await waitFor(() => dom.window.__app && dom.window.__app.STATE);
      assertEqual(dom.window.document.getElementById('export-dir-bar').style.display, 'flex');
    });

    await testAsync('EXPORT_DIR_HANDLE未指定なら、saveBlobは同期的にダウンロードしonDoneに"download"を渡す', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      win.URL.createObjectURL = () => 'blob:mock';
      let clicked = false;
      win.HTMLAnchorElement.prototype.click = function () { clicked = true; };
      let result = null;
      win.__app.saveBlob(new win.Blob(['{}']), 'test.json', (r) => { result = r; });
      assertEqual(result, 'download', 'フォルダ未指定時は同期的にdownloadが返るはず');
      assertTrue(clicked);
    });

    await testAsync('EXPORT_DIR_HANDLEが指定されていれば、フォルダへ直接書き込みonDoneに"dir"を渡す', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      const dirHandle = makeMockDirHandle(win, []);
      win.__app.setExportDirHandle(dirHandle);
      const blob = new win.Blob([JSON.stringify({ a: 1 })]);
      const result = await new Promise((resolve) => {
        win.__app.saveBlob(blob, 'foo.json', resolve);
      });
      assertEqual(result, 'dir');
      assertEqual(JSON.parse(dirHandle.written['foo.json']), { a: 1 }, '指定フォルダへ正しい内容が書き込まれるはず');
    });

    await testAsync('同名ファイルが既にある場合、確認ダイアログでキャンセルすると上書きされずcancelledを返す', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      const dirHandle = makeMockDirHandle(win, ['foo.json']);
      win.__app.setExportDirHandle(dirHandle);
      win.confirm = () => false;
      const blob = new win.Blob(['新しい内容']);
      const result = await new Promise((resolve) => {
        win.__app.saveBlob(blob, 'foo.json', resolve);
      });
      assertEqual(result, 'cancelled');
      assertEqual(dirHandle.written['foo.json'], undefined, '上書きを拒否したので書き込まれていないはず');
    });

    await testAsync('同名ファイルが既にある場合、確認ダイアログでOKすると上書きしdirを返す', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      const dirHandle = makeMockDirHandle(win, ['foo.json']);
      win.__app.setExportDirHandle(dirHandle);
      win.confirm = () => true;
      const blob = new win.Blob(['新しい内容']);
      const result = await new Promise((resolve) => {
        win.__app.saveBlob(blob, 'foo.json', resolve);
      });
      assertEqual(result, 'dir');
      assertEqual(dirHandle.written['foo.json'], '新しい内容');
    });

    await testAsync('フォルダへの書き込みに失敗した場合はダウンロードにフォールバックする', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      win.__app.setExportDirHandle(makeFailingDirHandle());
      win.URL.createObjectURL = () => 'blob:mock';
      let clicked = false;
      win.HTMLAnchorElement.prototype.click = function () { clicked = true; };
      const result = await new Promise((resolve) => {
        win.__app.saveBlob(new win.Blob(['{}']), 'test.json', resolve);
      });
      assertEqual(result, 'download', '書き込み失敗時は通常のダウンロードにフォールバックするはず');
      assertTrue(clicked);
    });

    await testAsync('レビュー画面の個別「書き出す」ボタンもEXPORT_DIR_HANDLEを尊重する', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const c3id = 'cell_R3_C3', c4id = 'cell_R3_C4';
      structure.reviewFields = [c4id];
      structure.fileNameFields = [c3id];
      const dom = await readyPage(structure);
      const win = dom.window;
      const data1 = buildSampleData(dom, { [c3id]: '事業A', [c4id]: '' });
      win.__app.upsertRecords([{ fileName: 'a.json', data: data1 }]);
      win.__app.enterReviewMode();

      const dirHandle = makeMockDirHandle(win, []);
      win.__app.setExportDirHandle(dirHandle);
      const record = win.__app.REVIEW.records[0];
      const expectedName = win.__app.buildExportFileNameForRecord(record);

      const row = win.document.querySelector('#review-table-root tbody tr');
      const exportBtn = [...row.querySelectorAll('button')].find(b => b.textContent.includes('書き出す'));
      exportBtn.click();
      await waitFor(() => dirHandle.written[expectedName] !== undefined);
      assertTrue(win.document.getElementById('status').textContent.includes('指定フォルダへ保存しました'));
    });
  });

  const ok = summary();
  process.exit(ok ? 0 : 1);
})();
