// 集約ツール（ツール3・aggregator_app.js）の結合テスト。
// ビルダーが書き出す集約ツールHTMLと同じ構成（core_logic.js + grid_render.js +
// `const STRUCTURE = {...}; const VIEWER_TEMPLATE = "...";` + aggregator_app.js）で
// jsdomにページを作り、window.__app（aggregator_app.jsが公開するデバッグ用API）越しに検証する。

const path = require('path');
const fs = require('fs');
const { JSDOM } = require('jsdom');
const { runSuiteAsync, testAsync, assertEqual, assertTrue, assertFalse, summary } = require('./assert_mini');

const ROOT = path.join(__dirname, '..');
const SRC = {
  core: fs.readFileSync(path.join(ROOT, 'core_logic.js'), 'utf8'),
  grid: fs.readFileSync(path.join(ROOT, 'grid_render.js'), 'utf8'),
  aggregator: fs.readFileSync(path.join(ROOT, 'aggregator_app.js'), 'utf8'),
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

// test_filler_integration.jsのbuildStructureFromFixtureと同じ（serializeStructure()と
// 同じ形をNode単体で再現する）。
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
    formTitle: formTitle || 'テスト様式',
    maxRow, maxCol, widths: [], heights: [],
    sections: sections.map(s => ({ title: s.title, row0: s.row0, row1: s.row1 })),
    cells, manualGroups: [],
  };
}

function cellIdFor(structure, row, col) {
  return `cell_R${row}_C${col}`;
}

// 「事業名」（A列見出し）→ B列に値、という最小の1セクション様式。test_viewer_integration.jsの
// buildSimpleStructureと同じ考え方（fileNameFieldsによる同一事業判定のテスト専用。
// FIXTURE（simple_moshikomi.xlsx）はrenderType未設定の列が複数行にまたがり自動命名が
// 曖昧になるため、識別キーの検証にはこちらの手書きfixtureを使う）。
function buildFileNameFieldsStructure() {
  return {
    formTitle: 'テスト様式',
    maxRow: 1, maxCol: 2, widths: [], heights: [],
    sections: [{ title: 'シート', row0: 1, row1: 1 }],
    manualGroups: [],
    cells: [
      { row: 1, col: 1, row2: 1, col2: 1, value: '事業名', isFormula: false, formula: '', hasText: true, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null },
      { row: 1, col: 2, row2: 1, col2: 2, value: '', isFormula: false, formula: '', hasText: false, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null },
    ],
  };
}

// 実機確認用に保存済みの「R7年度実施 事務事業評価_入力フォーム.html」（Plan/Do/Check/Action
// 様式、Doセクションに「決算見込額」等8行の繰り返し列を持つ）からSTRUCTUREを取り出す。
// 繰り返し列（グループ）を含む実データでの検証に使う（simple_moshikomi.xlsxには
// 繰り返し行が無いため）。
function loadRealJimujigyouStructure() {
  const formHtml = fs.readFileSync(path.join(ROOT, '実機確認', 'R7年度実施 事務事業評価_入力フォーム.html'), 'utf8');
  const m = formHtml.match(/const STRUCTURE = (\{[\s\S]*?\});\n/);
  if (!m) throw new Error('STRUCTURE not found in R7年度実施 事務事業評価_入力フォーム.html');
  return JSON.parse(m[1]);
}

function buildViewerTemplateText() {
  const viewerApp = fs.readFileSync(path.join(ROOT, 'viewer_app.js'), 'utf8');
  return fs.readFileSync(path.join(ROOT, 'viewer_template.html'), 'utf8')
    .replace('/* __CORE_LOGIC__ */', SRC.core)
    .replace('/* __GRID_RENDER__ */', SRC.grid)
    .replace('/* __VIEWER_APP__ */', viewerApp);
}

// 実際のaggregator_template.htmlと同じID体系を最小限そろえたfixture。
function newAggregatorPage(structure, opts) {
  opts = opts || {};
  let aggregatorSrc = SRC.aggregator;
  if (opts.withRealViewerTemplate) {
    const viewerText = buildViewerTemplateText();
    const viewerJsonLiteral = JSON.stringify(viewerText).replace(/<\/(script)/gi, '<\\/$1');
    aggregatorSrc = aggregatorSrc.replace('"__VIEWER_TEMPLATE_JSON__"', viewerJsonLiteral);
  }
  const html = `<!doctype html><html><body>
    <div id="export-dir-bar">
      <button id="btn-pick-export-dir">pickdir</button>
      <span id="export-dir-status"></span>
    </div>
    <div id="status"></div>

    <div id="load-actions">
      <input type="file" id="load-file-input" multiple>
      <button id="btn-pick-load-dir" style="display:none">dir</button>
      <button id="btn-refresh-load-dir" style="display:none">refresh</button>
    </div>
    <div id="loaded-summary"></div>
    <div id="loaded-table-root"></div>

    <div id="config-display-root"></div>
    <div id="config-aggregate-root"></div>
    <div id="config-search-root"></div>
    <input type="text" id="cfg-title">
    <textarea id="cfg-description"></textarea>
    <input type="text" id="cfg-source">
    <input type="text" id="cfg-updated">

    <div id="export-actions">
      <button id="btn-export-viewer">export</button>
      <button id="btn-export-aggregated-json">export json</button>
      <button id="btn-export-aggregated-csv">export csv</button>
    </div>

    <div id="grid-root" style="display:none"></div>
    <div id="agg-scratch-root" style="display:none"></div>
    <script>${SRC.core}</script>
    <script>${SRC.grid}</script>
    <script>
      const STRUCTURE = ${JSON.stringify(structure)};
      ${aggregatorSrc}
    </script>
  </body></html>`;
  const dom = new JSDOM(html, { runScripts: 'dangerously', url: 'http://localhost/' });
  dom.window.URL.createObjectURL = () => 'blob:mock';
  dom.window.URL.revokeObjectURL = () => {};
  dom.window.HTMLAnchorElement.prototype.click = function () {};
  dom.window.Element.prototype.scrollIntoView = function () {};
  return dom;
}

async function readyPage(structure, opts) {
  const dom = newAggregatorPage(structure, opts);
  await waitFor(() => dom.window.__app && dom.window.__app.CURRENT_STATE);
  return dom;
}

(async () => {
  await runSuiteAsync('aggregator_app: STEP1（JSON読込）', async () => {
    await testAsync('handleFileInputで複数JSONを読み込むとAGG.recordsに追加され、loaded-summaryに件数が出る', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      const files = [
        new win.File([JSON.stringify({ シート: { col3: 'A事業' } })], 'a.json', { type: 'application/json' }),
        new win.File([JSON.stringify({ シート: { col3: 'B事業' } })], 'b.json', { type: 'application/json' }),
      ];
      await win.__app.handleFileInput(files);
      assertEqual(win.__app.AGG.records.length, 2);
      assertEqual(win.document.getElementById('loaded-summary').textContent, '2件 読み込み済み');
    });

    await testAsync('同じファイル名のJSONを再度読み込んでも重複追加されない（upsert）', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      const f1 = new win.File([JSON.stringify({ シート: { col3: 'A事業' } })], 'a.json', { type: 'application/json' });
      await win.__app.handleFileInput([f1]);
      await win.__app.handleFileInput([f1]);
      assertEqual(win.__app.AGG.records.length, 1);
    });

    await testAsync('fileNameFields未設定の様式では、内容が同じでもファイル名が違えば別レコードとして追加される（従来挙動）', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      const f1 = new win.File([JSON.stringify({ シート: { col3: 'A事業' } })], 'a_20260810.json', { type: 'application/json' });
      const f2 = new win.File([JSON.stringify({ シート: { col3: 'A事業' } })], 'a_20260811.json', { type: 'application/json' });
      await win.__app.handleFileInput([f1]);
      await win.__app.handleFileInput([f2]);
      assertEqual(win.__app.AGG.records.length, 2, 'fileNameFieldsが無い様式では内容の同一性を判定できないため、従来通り2件になる');
    });

    await testAsync('fileNameFields設定済みの様式では、同じ事業の再提出（内容修正・別ファイル名）は重複追加ではなく既存レコードの更新として扱われる', async () => {
      const structure = buildFileNameFieldsStructure();
      structure.fileNameFields = [cellIdFor(structure, 1, 2)];
      const dom = await readyPage(structure);
      const win = dom.window;
      // 部署が1回目に提出（ファイル名に書き出し日20260810を含む）
      const f1 = new win.File([JSON.stringify({ シート: { 事業名: '広報広聴事業' } })], 'テスト様式_広報広聴事業_20260810.json', { type: 'application/json' });
      await win.__app.handleFileInput([f1]);
      assertEqual(win.__app.AGG.records.length, 1);

      // 別の事業も1件読み込んでおく（更新対象と無関係なレコードが巻き込まれないことの確認用）
      const other = new win.File([JSON.stringify({ シート: { 事業名: '道路維持事業' } })], 'テスト様式_道路維持事業_20260810.json', { type: 'application/json' });
      await win.__app.handleFileInput([other]);
      assertEqual(win.__app.AGG.records.length, 2);

      // 同じ「広報広聴事業」を後日ファイル名（日付部分）を変えて再提出したケースを再現
      const f1Updated = new win.File([JSON.stringify({ シート: { 事業名: '広報広聴事業' } })], 'テスト様式_広報広聴事業_20260812.json', { type: 'application/json' });
      const result = await win.__app.handleFileInput([f1Updated]);
      assertEqual(win.__app.AGG.records.length, 2, 'ファイル名が違っても同じ事業と判定され、件数は増えないはず');
      const rec = win.__app.AGG.records.find(r => r.fileName === 'テスト様式_広報広聴事業_20260812.json');
      assertTrue(!!rec, '新しいファイル名で既存レコードが更新されているはず');
      assertFalse(win.__app.AGG.records.some(r => r.fileName === 'テスト様式_広報広聴事業_20260810.json'), '古いファイル名のレコードは残っていないはず（置き換え済み）');
      assertTrue(win.document.getElementById('status').textContent.includes('既存事業の更新'), 'ステータスに更新件数の注記が出るはず');
    });

    await testAsync('businessKeyForRecordは、fileNameFieldsの値をつないだ文字列を返す（fileNameFields未設定ならnull）', async () => {
      const withKeyStructure = buildFileNameFieldsStructure();
      withKeyStructure.fileNameFields = [cellIdFor(withKeyStructure, 1, 2)];
      const domWithKey = await readyPage(withKeyStructure);
      const key = domWithKey.window.__app.businessKeyForRecord({ シート: { 事業名: '広報広聴事業' } });
      assertEqual(key, '広報広聴事業');

      const withoutKeyStructure = buildFileNameFieldsStructure();
      const domWithoutKey = await readyPage(withoutKeyStructure);
      const noKey = domWithoutKey.window.__app.businessKeyForRecord({ シート: { 事業名: '広報広聴事業' } });
      assertEqual(noKey, null);
    });

    await testAsync('壊れたJSONはスキップされる', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      const bad = new win.File(['{invalid'], 'bad.json', { type: 'application/json' });
      const good = new win.File([JSON.stringify({ シート: {} })], 'good.json', { type: 'application/json' });
      await win.__app.handleFileInput([bad, good]);
      assertEqual(win.__app.AGG.records.length, 1);
    });

    await testAsync('removeRecordで読み込み済みレコードを除ける', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      const f1 = new win.File([JSON.stringify({ シート: {} })], 'a.json', { type: 'application/json' });
      await win.__app.handleFileInput([f1]);
      win.__app.removeRecord('a.json');
      assertEqual(win.__app.AGG.records.length, 0);
    });
  });

  await runSuiteAsync('aggregator_app: STEP2（住民公開設定）', async () => {
    await testAsync('candidateSinglesはSTRUCTUREの単独入力欄を返す（グループ化された列は含まない）', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      const singles = win.__app.candidateSingles();
      assertTrue(singles.length >= 2, 'このフィクスチャは単独の入力欄を2件以上持つはず');
    });

    await testAsync('①表示項目のチェックボックスをONにするとCONFIG.displayFieldIdsに入る', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      const target = win.__app.candidateSingles()[0];
      const key = win.__app.aggregateKey(target);
      const cb = win.document.getElementById('cfgdisplay_' + key);
      assertTrue(!!cb, '①のチェックボックスが描画されているはず');
      cb.click();
      assertTrue(win.__app.CONFIG.displayFieldIds.has(key));
      cb.click();
      assertFalse(win.__app.CONFIG.displayFieldIds.has(key));
    });

    await testAsync('②集計対象は数値・金額型の項目のみが候補になる（単独欄）', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      // 1件だけ数値型にしておく
      const numericCell = structure.cells.find(c => !c.hasText && !c.isFormula);
      numericCell.renderType = 'number';
      const dom = await readyPage(structure);
      const win = dom.window;
      const numeric = win.__app.numericCandidateSingles();
      assertEqual(numeric.length, 1);
      const key = win.__app.aggregateKey(numeric[0]);
      assertTrue(!!win.document.getElementById('cfgagg_' + key), '②に数値項目のチェックボックスがあるはず');
    });

    await testAsync('②単位を入力するとCONFIG.aggregateUnitsに保存される（単独欄）', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const numericCell = structure.cells.find(c => !c.hasText && !c.isFormula);
      numericCell.renderType = 'currency';
      const dom = await readyPage(structure);
      const win = dom.window;
      const numeric = win.__app.numericCandidateSingles();
      const key = win.__app.aggregateKey(numeric[0]);
      win.document.getElementById('cfgagg_' + key).click();
      const unitInput = win.document.querySelector(`label[for="cfgagg_${key}"] input[type=text]`);
      unitInput.value = '千円';
      unitInput.dispatchEvent(new win.Event('input'));
      assertEqual(win.__app.CONFIG.aggregateUnits.get(key), '千円');
    });

    await testAsync('②集計対象には、繰り返し行（グループ化された列）の数値・金額項目も候補として出る（Excelの合計セル代替）', async () => {
      // simple_moshikomi.xlsxには繰り返し行が無いため、実際に繰り返し行を持つ実機確認用の
      // 様式（R7年度実施 事務事業評価。Doセクションに「決算見込額」等の8行の繰り返し列がある）
      // から STRUCTURE を読み、実データで検証する。
      const structure = loadRealJimujigyouStructure();
      const dom = await readyPage(structure);
      const win = dom.window;
      const groups = win.__app.groupCandidatesNumeric();
      assertTrue(groups.length > 0, '数値・金額型の繰り返し列がgroupCandidatesNumericに含まれるはず');
      const combined = win.__app.numericAggregateCandidates();
      assertTrue(combined.some(t => t.kind === 'group'), '②の候補一覧（単独欄＋繰り返し列）にグループ種別が含まれるはず');

      const target = groups.find(t => t.groupLabel === 'R6年度(2024年度)' && t.cells[0].renderType === 'currency');
      assertTrue(!!target, '「決算見込額」列（8行）が候補にあるはず');
      const key = win.__app.aggregateKey(target);
      const cb = win.document.getElementById('cfgagg_' + key);
      assertTrue(!!cb, '②に繰り返し列のチェックボックスがあるはず');
      cb.click();
      assertTrue(win.__app.CONFIG.aggregateFieldIds.has(key));

      const cfg = win.__app.buildPublicConfig();
      const field = cfg.aggregateFields.find(f => f.ids.length === target.cells.length);
      assertTrue(!!field, '書き出し設定に、行数分（8件）のidsを持つ集計項目が含まれるはず');
      assertEqual(field.ids.length, 8);
    });

    await testAsync('③検索対象項目は①表示項目と独立して選べる', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      const target = win.__app.candidateSingles()[0];
      const key = win.__app.aggregateKey(target);
      win.document.getElementById('cfgsearch_' + key).click();
      assertTrue(win.__app.CONFIG.searchFieldIds.has(key));
      assertFalse(win.__app.CONFIG.displayFieldIds.has(key), '③をONにしても①には影響しないはず');
    });

    await testAsync('③検索対象項目には、テキスト型の繰り返し行（グループ化された列）も候補として出る（①②は数値・金額型のみのため対象外）', async () => {
      const structure = loadRealJimujigyouStructure();
      const dom = await readyPage(structure);
      const win = dom.window;
      const pool = win.__app.searchCandidatePool();
      const textGroup = pool.find(t => t.kind === 'group' && t.cells[0].renderType !== 'number' && t.cells[0].renderType !== 'currency');
      assertTrue(!!textGroup, 'テキスト型の繰り返し列（例：事業内容と活動実績）が③の候補にあるはず');
      assertFalse(win.__app.groupCandidatesNumeric().includes(textGroup), 'テキスト型のグループは①②の候補（数値・金額型のみ）には含まれないはず');

      const key = win.__app.aggregateKey(textGroup);
      const cb = win.document.getElementById('cfgsearch_' + key);
      assertTrue(!!cb, '③にテキスト型の繰り返し列のチェックボックスがあるはず');
      cb.click();

      const cfg = win.__app.buildPublicConfig();
      // buildSearchFieldIdsはフラットなセルidの配列を返す（表示・集計と違いlabelでまとめない）。
      const flatIds = textGroup.cells.map(c => win.CoreLogic.cellId(c));
      assertTrue(flatIds.every(id => cfg.searchFields.includes(id)), '選んだ繰り返し列の全セルidがsearchFieldsに含まれるはず');
    });

    await testAsync('④タイトル・出典・更新日は既定値が自動入力される', async () => {
      const structure = buildStructureFromFixture(FIXTURE, 'サンプル様式');
      const dom = await readyPage(structure);
      const win = dom.window;
      assertEqual(win.document.getElementById('cfg-title').value, 'サンプル様式');
      assertTrue(win.document.getElementById('cfg-source').value.includes('サンプル様式'));
      assertTrue(/^\d{4}-\d{2}-\d{2}$/.test(win.document.getElementById('cfg-updated').value));
    });

    await testAsync('④出典の件数は、STEP1でJSONを読み込んだ後（0件のまま自動入力された後）でも最新の件数に更新される', async () => {
      const structure = buildStructureFromFixture(FIXTURE, 'サンプル様式');
      const dom = await readyPage(structure);
      const win = dom.window;
      // STEP1未読込の時点（0件）で既定値が自動入力されている
      assertTrue(win.document.getElementById('cfg-source').value.includes('0件'), win.document.getElementById('cfg-source').value);
      await win.__app.handleFileInput([new win.File([JSON.stringify({ シート: {} })], 'a.json', { type: 'application/json' })]);
      assertTrue(win.document.getElementById('cfg-source').value.includes('1件'), win.document.getElementById('cfg-source').value);
    });

    await testAsync('④出典をユーザーが手動で編集した後は、STEP1でJSONを読み込んでも上書きされない', async () => {
      const structure = buildStructureFromFixture(FIXTURE, 'サンプル様式');
      const dom = await readyPage(structure);
      const win = dom.window;
      const sourceInput = win.document.getElementById('cfg-source');
      sourceInput.value = 'カスタム出典';
      sourceInput.dispatchEvent(new win.Event('input'));
      await win.__app.handleFileInput([new win.File([JSON.stringify({ シート: {} })], 'a.json', { type: 'application/json' })]);
      assertEqual(win.document.getElementById('cfg-source').value, 'カスタム出典');
    });
  });

  await runSuiteAsync('aggregator_app: STEP3（住民公開用ページの書き出し）', async () => {
    await testAsync('buildPublicConfigは選択順ではなくシート上の並び順でdisplayFieldsを返す', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      const singles = win.__app.candidateSingles();
      assertTrue(singles.length >= 2);
      const firstKey = win.__app.aggregateKey(singles[0]);
      const secondKey = win.__app.aggregateKey(singles[1]);
      win.document.getElementById('cfgdisplay_' + secondKey).click();
      win.document.getElementById('cfgdisplay_' + firstKey).click();
      const cfg = win.__app.buildPublicConfig();
      assertEqual(cfg.displayFields.map(f => f.ids[0]), [
        win.CoreLogic.cellId(singles[0].cells[0]),
        win.CoreLogic.cellId(singles[1].cells[0]),
      ]);
    });

    await testAsync('レコードが0件のままdoExportAsViewerを呼ぶと警告になり、書き出されない', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      let clicked = false;
      win.HTMLAnchorElement.prototype.click = function () { clicked = true; };
      win.__app.doExportAsViewer();
      assertFalse(clicked, 'レコード0件では書き出し（ダウンロード）は行われないはず');
      assertTrue(win.document.getElementById('status').textContent.includes('読み込まれていません'));
    });

    await testAsync('doExportAsViewerは例外を投げず、書き出し完了のステータスになる', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      await win.__app.handleFileInput([new win.File([JSON.stringify({ シート: {} })], 'a.json', { type: 'application/json' })]);
      win.__app.doExportAsViewer();
      assertEqual(win.document.getElementById('status').textContent, '住民公開用ページを書き出しました（1件）。');
    });

    await testAsync('書き出されるファイル名はタイトル＋「_公開ページ.html」', async () => {
      const structure = buildStructureFromFixture(FIXTURE, 'カスタム様式');
      const dom = await readyPage(structure);
      const win = dom.window;
      await win.__app.handleFileInput([new win.File([JSON.stringify({ シート: {} })], 'a.json', { type: 'application/json' })]);
      let capturedDownload = null;
      win.HTMLAnchorElement.prototype.click = function () { capturedDownload = this.download; };
      win.__app.doExportAsViewer();
      assertEqual(capturedDownload, 'カスタム様式_公開ページ.html');
    });

    await testAsync('実際のviewer_template.htmlを埋め込んだ状態で書き出すと、STRUCTURE・RECORDS・PUBLIC_CONFIGが注入されたHTMLになる', async () => {
      const structure = buildStructureFromFixture(FIXTURE, '実配布様式');
      const dom = await readyPage(structure, { withRealViewerTemplate: true });
      const win = dom.window;
      await win.__app.handleFileInput([
        new win.File([JSON.stringify({ シート: { col3: '事業A' } })], 'a.json', { type: 'application/json' }),
        new win.File([JSON.stringify({ シート: { col3: '事業B' } })], 'b.json', { type: 'application/json' }),
      ]);
      let capturedBlob = null;
      win.URL.createObjectURL = (blob) => { capturedBlob = blob; return 'blob:mock'; };
      win.HTMLAnchorElement.prototype.click = function () {};
      win.__app.doExportAsViewer();
      assertTrue(!!capturedBlob, '書き出されたBlobが捕捉できているはず');
      const html = await new Promise((resolve, reject) => {
        const fr = new win.FileReader();
        fr.onload = () => resolve(fr.result);
        fr.onerror = reject;
        fr.readAsText(capturedBlob);
      });
      assertTrue(html.includes('<title>実配布様式</title>'), '住民公開ページのtitleに反映されているはず');
      assertTrue(html.includes('"formTitle":"実配布様式"'), 'STRUCTUREが注入されているはず');
      assertTrue(html.includes('事業A') && html.includes('事業B'), 'RECORDSが注入されているはず');
      assertFalse(html.includes('__STRUCTURE__') || html.includes('__RECORDS__') || html.includes('__PUBLIC_CONFIG__'), 'プレースホルダーが残っていないはず');
    });

    await testAsync('レコードが0件のままdoExportAggregatedJsonを呼ぶと警告になり、書き出されない', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      let clicked = false;
      win.HTMLAnchorElement.prototype.click = function () { clicked = true; };
      win.__app.doExportAggregatedJson();
      assertFalse(clicked, 'レコード0件では書き出しは行われないはず');
      assertTrue(win.document.getElementById('status').textContent.includes('読み込まれていません'));
    });

    await testAsync('doExportAggregatedJsonは、読み込んだ全事業の生データを1つのJSONにまとめて書き出す', async () => {
      const structure = buildStructureFromFixture(FIXTURE, 'サンプル様式');
      const dom = await readyPage(structure);
      const win = dom.window;
      await win.__app.handleFileInput([
        new win.File([JSON.stringify({ シート: { col3: 'A事業' } })], 'a.json', { type: 'application/json' }),
        new win.File([JSON.stringify({ シート: { col3: 'B事業' } })], 'b.json', { type: 'application/json' }),
      ]);
      let capturedBlob = null;
      let capturedDownload = null;
      win.URL.createObjectURL = (blob) => { capturedBlob = blob; return 'blob:mock'; };
      win.HTMLAnchorElement.prototype.click = function () { capturedDownload = this.download; };
      win.__app.doExportAggregatedJson();
      assertTrue(!!capturedDownload && capturedDownload.includes('サンプル様式') && capturedDownload.endsWith('.json'));
      const text = await new Promise((resolve, reject) => {
        const fr = new win.FileReader();
        fr.onload = () => resolve(fr.result);
        fr.onerror = reject;
        fr.readAsText(capturedBlob);
      });
      const payload = JSON.parse(text);
      assertEqual(payload.meta.count, 2);
      assertEqual(payload.meta.source_title, 'サンプル様式');
      assertEqual(payload.records.length, 2);
      assertEqual(payload.records[0].fileName, 'a.json');
      assertEqual(payload.records[0].data, { シート: { col3: 'A事業' } });
      assertEqual(payload.records[1].data, { シート: { col3: 'B事業' } });
    });

    await testAsync('レコードが0件のままdoExportAggregatedCsvを呼ぶと警告になり、書き出されない', async () => {
      const structure = buildStructureFromFixture(FIXTURE);
      const dom = await readyPage(structure);
      const win = dom.window;
      let clicked = false;
      win.HTMLAnchorElement.prototype.click = function () { clicked = true; };
      win.__app.doExportAggregatedCsv();
      assertFalse(clicked, 'レコード0件では書き出しは行われないはず');
      assertTrue(win.document.getElementById('status').textContent.includes('読み込まれていません'));
    });

    await testAsync('doExportAggregatedCsvは、単独欄・繰り返し列を列にしたBOM付きUTF-8のCSVを書き出す（値のカンマ・改行はダブルクォートでエスケープ）', async () => {
      const structure = loadRealJimujigyouStructure();
      const dom = await readyPage(structure);
      const win = dom.window;
      const singles = win.__app.candidateSingles();
      const nameField = singles.find(t => t.rowLabel === '事業名' && t.autoName === 'col2');
      const nameId = win.CoreLogic.cellId(nameField.cells[0]);

      // 実際の入力フォームと同じ形（DOM入力→collectData）でJSONを作る（手書きだと構造が
      // 食い違うリスクがあるため、他のテストと同じ手法を使う）。
      const makeRecord = (label) => {
        const scratch = win.document.getElementById('agg-scratch-root');
        const tempState = win.__app.buildStateFromStructure(structure);
        win.GridRender.renderGrid(scratch, tempState, { showGear: false });
        win.document.getElementById(nameId).value = label;
        const data = win.GridRender.collectData(tempState);
        scratch.innerHTML = '';
        return data;
      };

      await win.__app.handleFileInput([
        new win.File([JSON.stringify(makeRecord('カンマ,と"クォート"を含む事業名'))], 'a.json', { type: 'application/json' }),
      ]);

      let capturedBlob = null;
      let capturedDownload = null;
      win.URL.createObjectURL = (blob) => { capturedBlob = blob; return 'blob:mock'; };
      win.HTMLAnchorElement.prototype.click = function () { capturedDownload = this.download; };
      win.__app.doExportAggregatedCsv();
      assertTrue(!!capturedDownload && capturedDownload.endsWith('.csv'));
      // FileReader.readAsText()はデコード時にBOMを自動的に取り除く（TextDecoderの標準挙動）
      // ため、BOMの有無はバイト列（readAsArrayBuffer）側で確認する。Excel等がUTF-8と正しく
      // 認識できるかはこの生バイトの並びで決まるため、テキストとして読んだ結果を見ても分からない。
      const buf = await new Promise((resolve, reject) => {
        const fr = new win.FileReader();
        fr.onload = () => resolve(fr.result);
        fr.onerror = reject;
        fr.readAsArrayBuffer(capturedBlob);
      });
      const bytes = new Uint8Array(buf);
      assertEqual([bytes[0], bytes[1], bytes[2]].join(','), '239,187,191', '先頭にUTF-8のBOM（EF BB BF）が付いているはず');

      const text = await new Promise((resolve, reject) => {
        const fr = new win.FileReader();
        fr.onload = () => resolve(fr.result);
        fr.onerror = reject;
        fr.readAsText(capturedBlob);
      });
      const lines = text.split('\r\n');
      assertTrue(lines[0].startsWith('ファイル名,'), '1行目はヘッダー行のはず');
      assertTrue(lines[0].includes('事業名'), 'ヘッダーに単独欄のラベルが含まれるはず');
      assertTrue(lines[0].includes('決算見込額'), 'ヘッダーに繰り返し列のラベルも含まれるはず（Excelの合計セルと同じ範囲）');
      assertTrue(lines[1].includes('"カンマ,と""クォート""を含む事業名"'), 'カンマとダブルクォートを含む値が正しくエスケープされているはず');
    });
  });

  const ok = summary();
  process.exit(ok ? 0 : 1);
})();
