// 住民公開ページ（ツール4・viewer_app.js）の結合テスト。
// core_logic.js + grid_render.js + `const STRUCTURE/RECORDS/PUBLIC_CONFIG = {...};` +
// viewer_app.js という、集約ツールが実際に書き出す構成でjsdomにページを作り、
// window.__app（viewer_app.jsが公開するデバッグ用API）越しに検証する。

const path = require('path');
const fs = require('fs');
const { JSDOM } = require('jsdom');
const { runSuiteAsync, testAsync, assertEqual, assertTrue, assertFalse, summary } = require('./assert_mini');

const ROOT = path.join(__dirname, '..');
const SRC = {
  core: fs.readFileSync(path.join(ROOT, 'core_logic.js'), 'utf8'),
  grid: fs.readFileSync(path.join(ROOT, 'grid_render.js'), 'utf8'),
  viewer: fs.readFileSync(path.join(ROOT, 'viewer_app.js'), 'utf8'),
};

// viewer_template.htmlの<style>...</style>をそのまま抜き出す。手書きの簡易fixtureに独自の
// display:none等を仮置きすると、実テンプレートのCSSとJSの組み合わせ不整合（実機確認で発覚した
// 「#detail-rootがスタイルシート側でdisplay:noneのため、JS側でstyle.display=''にリセットしても
// スタイルシートにフォールバックして表示されない」不具合）をテストが検出できなくなるため、
// 必ず本物のCSSを使う。
const VIEWER_TEMPLATE_HTML = fs.readFileSync(path.join(ROOT, 'viewer_template.html'), 'utf8');
const VIEWER_STYLE_BLOCK = (VIEWER_TEMPLATE_HTML.match(/<style>[\s\S]*?<\/style>/) || [''])[0];

// getComputedStyle（スタイルシートとの合成結果）で判定する。element.style.display
// （JSが直接セットした値）だけを見ると、「JSは''をセットしたつもりでも、スタイルシート側の
// display:noneにフォールバックして実際には非表示のまま」という不具合を見逃す
// （今回の実機確認で発覚した不具合そのもの）。
function isVisible(win, el) {
  return win.getComputedStyle(el).display !== 'none';
}

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

// 「事業名」（A列見出し）→ B列に値、という最小の1セクション様式。
// filler_app.jsのbuildSingleLabelRowStructure等と同じ、手書きcells配列によるfixture。
function buildSimpleStructure() {
  return {
    formTitle: 'テスト様式',
    maxRow: 3, maxCol: 2, widths: [], heights: [],
    sections: [{ title: 'シート', row0: 1, row1: 3 }],
    manualGroups: [],
    cells: [
      { row: 1, col: 1, row2: 1, col2: 1, value: '事業名', isFormula: false, formula: '', hasText: true, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null },
      { row: 1, col: 2, row2: 1, col2: 2, value: '', isFormula: false, formula: '', hasText: false, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null },
      { row: 2, col: 1, row2: 2, col2: 1, value: '担当課', isFormula: false, formula: '', hasText: true, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null },
      { row: 2, col: 2, row2: 2, col2: 2, value: '', isFormula: false, formula: '', hasText: false, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null },
      { row: 3, col: 1, row2: 3, col2: 1, value: '予算', isFormula: false, formula: '', hasText: true, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null },
      { row: 3, col: 2, row2: 3, col2: 2, value: '', isFormula: false, formula: '', hasText: false, blocked: false, fillColor: null, renderType: 'number', renderOptions: null, dbKey: null },
    ],
  };
}

function cellIdFor(structure, row, col) {
  return `cell_R${row}_C${col}`;
}

// 繰り返し行（グループ化された列。例：Excelの合計セルと同じ範囲の「明細」2行）を持つ様式。
// aggregateFieldsのids（複数セル）を、事業内でまず合計してから事業間で合計する挙動の検証用。
function buildGroupStructure() {
  return {
    formTitle: 'グループテスト様式',
    maxRow: 5, maxCol: 2, widths: [], heights: [],
    sections: [{ title: 'シート', row0: 1, row1: 5 }],
    manualGroups: [{ row0: 4, row1: 5, name: '明細' }],
    cells: [
      { row: 1, col: 1, row2: 1, col2: 1, value: '事業名', isFormula: false, formula: '', hasText: true, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null },
      { row: 1, col: 2, row2: 1, col2: 2, value: '', isFormula: false, formula: '', hasText: false, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null },
      { row: 4, col: 2, row2: 4, col2: 2, value: '', isFormula: false, formula: '', hasText: false, blocked: false, fillColor: null, renderType: 'currency', renderOptions: null, dbKey: null },
      { row: 5, col: 2, row2: 5, col2: 2, value: '', isFormula: false, formula: '', hasText: false, blocked: false, fillColor: null, renderType: 'currency', renderOptions: null, dbKey: null },
    ],
  };
}

function buildGroupSampleRecords() {
  return [
    { シート: { 事業名: '事業α', 明細: [{ col2: '1000' }, { col2: '1500' }] } },
    { シート: { 事業名: '事業β', 明細: [{ col2: '2000' }, { col2: '500' }] } },
  ];
}

function buildSampleRecords() {
  return [
    { シート: { 事業名: '広報広聴事業', 担当課: '広報課', 予算: '1000' } },
    { シート: { 事業名: '子育て支援事業', 担当課: 'こども未来課', 予算: '3000' } },
    { シート: { 事業名: '道路維持事業', 担当課: '土木課', 予算: '500' } },
  ];
}

function defaultPublicConfig(structure, overrides) {
  const cfg = {
    storageKey: 'jimujigyou_viewer_test',
    displayFields: [
      { label: '事業名', ids: [cellIdFor(structure, 1, 2)] },
      { label: '担当課', ids: [cellIdFor(structure, 2, 2)] },
    ],
    aggregateFields: [{ label: '予算', unit: '千円', ids: [cellIdFor(structure, 3, 2)] }],
    showRecordCount: true,
    searchFields: [cellIdFor(structure, 1, 2)],
    title: 'テスト公開ページ',
    description: 'テスト用の説明文です。',
    source: 'テスト様式（3件）を基に作成',
    updatedDate: '2026-08-12',
  };
  return Object.assign(cfg, overrides || {});
}

function newViewerPage(structure, records, publicConfig, url) {
  const html = `<!doctype html><html><head>${VIEWER_STYLE_BLOCK}</head><body>
    <h1 id="viewer-title"></h1>
    <p id="viewer-description"></p>
    <p><span id="viewer-source"></span><span id="viewer-updated"></span></p>

    <div id="list-root">
      <div id="viewer-search-bar">
        <input type="text" id="viewer-keyword">
        <label for="viewer-only-checked"><input type="checkbox" id="viewer-only-checked"></label>
        <label for="viewer-show-memo"><input type="checkbox" id="viewer-show-memo"></label>
      </div>
      <div id="viewer-col-toggle"></div>
      <div id="viewer-summary"></div>
      <div id="viewer-table-root"></div>
      <div class="action-bar">
        <button id="btn-print-checked">print</button>
        <button id="btn-export-backup">backup</button>
        <input type="file" id="backup-file-input">
      </div>
    </div>

    <div id="detail-root" style="display:none">
      <button id="btn-back-to-list">back</button>
      <div id="detail-nav">
        <button id="detail-prev">prev</button>
        <button id="detail-next">next</button>
        <button id="btn-print-detail">print</button>
      </div>
      <div id="detail-summary-bar"></div>
      <div id="detail-body-root"></div>
    </div>

    <div id="viewer-print-root"></div>
    <div id="viewer-scratch-root"></div>
    <script>${SRC.core}</script>
    <script>${SRC.grid}</script>
    <script>
      const STRUCTURE = ${JSON.stringify(structure)};
      const RECORDS = ${JSON.stringify(records)};
      const PUBLIC_CONFIG = ${JSON.stringify(publicConfig)};
      ${SRC.viewer}
    </script>
  </body></html>`;
  const dom = new JSDOM(html, { runScripts: 'dangerously', url: url || 'http://localhost/' });
  dom.window.URL.createObjectURL = () => 'blob:mock';
  dom.window.URL.revokeObjectURL = () => {};
  dom.window.HTMLAnchorElement.prototype.click = function () {};
  dom.window.Element.prototype.scrollIntoView = function () {};
  return dom;
}

async function readyPage(structure, records, publicConfig, url) {
  const dom = newViewerPage(structure, records, publicConfig || defaultPublicConfig(structure), url);
  await waitFor(() => dom.window.__app && dom.window.__app.V && dom.window.__app.V.records.length >= 0);
  await waitFor(() => dom.window.document.getElementById('viewer-table-root').innerHTML.length > 0);
  return dom;
}

(async () => {
  await runSuiteAsync('viewer_app: 初期化・一覧表示', async () => {
    await testAsync('レコードごとにdisplayFieldsの値が読み取られ、一覧表に行が並ぶ', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      const rows = win.document.querySelectorAll('#viewer-table-root tbody tr');
      assertEqual(rows.length, 3);
      assertTrue(win.document.getElementById('viewer-table-root').textContent.includes('広報広聴事業'));
    });

    await testAsync('タイトル・説明文・出典・更新日がヘッダーに反映される', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      assertEqual(win.document.getElementById('viewer-title').textContent, 'テスト公開ページ');
      assertEqual(win.document.getElementById('viewer-description').textContent, 'テスト用の説明文です。');
      assertTrue(win.document.getElementById('viewer-source').textContent.includes('テスト様式'));
      assertTrue(win.document.getElementById('viewer-updated').textContent.includes('2026-08-12'));
    });

    await testAsync('集計サマリーに件数と合計（単位付き）が表示される', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      const summaryText = win.document.getElementById('viewer-summary').textContent;
      assertTrue(summaryText.includes('3件'), summaryText);
      assertTrue(summaryText.includes('4,500') && summaryText.includes('千円'), summaryText);
    });

    await testAsync('集計対象が繰り返し行（ids複数件）の場合、事業内でまず合計してから事業間で合計する（Excelの合計セル代替）', async () => {
      const structure = buildGroupStructure();
      const publicConfig = {
        storageKey: 'jimujigyou_viewer_grouptest',
        displayFields: [cellIdFor(structure, 1, 2)],
        aggregateFields: [{ label: '明細', unit: '千円', ids: [cellIdFor(structure, 4, 2), cellIdFor(structure, 5, 2)] }],
        showRecordCount: true,
        searchFields: [],
        title: 'グループ集計テスト', description: '', source: '', updatedDate: '2026-08-12',
      };
      const dom = await readyPage(structure, buildGroupSampleRecords(), publicConfig);
      const win = dom.window;
      const summaryText = win.document.getElementById('viewer-summary').textContent;
      // 事業α(1000+1500)＋事業β(2000+500) = 5000
      assertTrue(summaryText.includes('5,000') && summaryText.includes('千円'), summaryText);
    });
  });

  await runSuiteAsync('viewer_app: 検索・絞り込み・並び替え', async () => {
    await testAsync('キーワード検索で検索対象項目にヒットする行だけに絞り込まれる', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      win.document.getElementById('viewer-keyword').value = '子育て';
      win.document.getElementById('viewer-keyword').dispatchEvent(new win.Event('input'));
      const rows = win.document.querySelectorAll('#viewer-table-root tbody tr');
      assertEqual(rows.length, 1);
      assertTrue(win.document.getElementById('viewer-table-root').textContent.includes('子育て支援事業'));
    });

    await testAsync('検索結果が絞り込まれると集計サマリーの合計も絞り込み後の値になる', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      win.document.getElementById('viewer-keyword').value = '広報';
      win.document.getElementById('viewer-keyword').dispatchEvent(new win.Event('input'));
      const summaryText = win.document.getElementById('viewer-summary').textContent;
      assertTrue(summaryText.includes('1,000'), summaryText);
      assertFalse(summaryText.includes('4,500'), summaryText);
    });

    await testAsync('列見出しの並び替えボタンをクリックすると昇順→降順に切り替わる（クリックのたびに先頭行が入れ替わる）', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      // renderList()のたびにtheadごと作り直されるため、クリック後は毎回ボタンを取り直す
      // （古い参照を保持し続けると、差し替え後のDOMに付いていないボタンを操作することになる）。
      const sortBtnAt = (idx) => [...win.document.querySelectorAll('.sort-btn')][idx];
      sortBtnAt(1).click(); // 担当課列
      const ascFirst = win.document.querySelector('#viewer-table-root tbody tr').textContent;
      assertTrue(sortBtnAt(1).textContent.includes('▲'), '1回目のクリックで昇順マークが付くはず');
      sortBtnAt(1).click();
      const descFirst = win.document.querySelector('#viewer-table-root tbody tr').textContent;
      assertTrue(sortBtnAt(1).textContent.includes('▼'), '2回目のクリックで降順マークが付くはず');
      assertTrue(ascFirst !== descFirst, '昇順と降順で先頭行が入れ替わるはず');
    });

    await testAsync('「気になる事業だけ表示」をONにすると☑した行だけに絞り込まれる', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      const firstCheckbox = win.document.querySelector('#viewer-table-root tbody tr td.viewer-chk-cell input[type=checkbox]');
      firstCheckbox.click();
      win.document.getElementById('viewer-only-checked').checked = true;
      win.document.getElementById('viewer-only-checked').dispatchEvent(new win.Event('change'));
      const rows = win.document.querySelectorAll('#viewer-table-root tbody tr');
      assertEqual(rows.length, 1);
    });
  });

  await runSuiteAsync('viewer_app: ☑・メモのlocalStorage永続化', async () => {
    await testAsync('☑をONにするとlocalStorageに保存され、ページ再読込後も復元される', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      const cb = win.document.querySelector('#viewer-table-root tbody tr td.viewer-chk-cell input[type=checkbox]');
      cb.click();
      assertTrue(win.__app.V.checked.size === 1);
      const raw = win.localStorage.getItem('jimujigyou_viewer_test');
      assertTrue(!!raw, 'localStorageに保存されているはず');
      const parsed = JSON.parse(raw);
      assertEqual(parsed.checked.length, 1);
    });

    await testAsync('メモを入力するとsetMemoでV.memosに保存され、localStorageにも反映される', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      const key = win.__app.V.records[0].persistKey;
      win.__app.setMemo(key, '住民として気になる');
      const raw = JSON.parse(win.localStorage.getItem('jimujigyou_viewer_test'));
      assertEqual(raw.memos[key], '住民として気になる');
    });

    await testAsync('exportBackup/importBackupの往復で☑・メモが復元される', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      const key = win.__app.V.records[1].persistKey;
      win.__app.toggleChecked(key);
      win.__app.setMemo(key, 'メモA');
      const payload = { checked: Array.from(win.__app.V.checked), memos: Object.fromEntries(win.__app.V.memos) };

      // 新しいページ（localStorage未設定の状態）にバックアップを読み込ませて復元できるか確認
      const dom2 = await readyPage(structure, buildSampleRecords());
      const win2 = dom2.window;
      assertEqual(win2.__app.V.checked.size, 0, '新しいページは未チェックのはず');
      const file = new win2.File([JSON.stringify(payload)], 'backup.json', { type: 'application/json' });
      await win2.__app.importBackup(file);
      assertTrue(win2.__app.V.checked.has(key));
      assertEqual(win2.__app.V.memos.get(key), 'メモA');
    });

    await testAsync('バックアップの☑・メモは、事業の並び順が変わっても同じ事業に正しく復元される（persistKeyの検証）', async () => {
      const structure = buildSimpleStructure();
      // ビルダーSTEP4相当：「事業名」をファイル名（＝識別）用フィールドとして指定した状態を再現。
      // これが無いとpersistKeyは配列indexにフォールバックし、この検証自体が成立しない。
      structure.fileNameFields = [cellIdFor(structure, 1, 2)];
      const records = buildSampleRecords();
      const dom = await readyPage(structure, records);
      const win = dom.window;
      // 「道路維持事業」（元は配列3番目=index2）にチェック・メモを付ける
      const target = win.__app.V.records.find(r => r.data.シート.事業名 === '道路維持事業');
      win.__app.toggleChecked(target.persistKey);
      win.__app.setMemo(target.persistKey, '並び順が変わっても付いてくるはず');
      const payload = { checked: Array.from(win.__app.V.checked), memos: Object.fromEntries(win.__app.V.memos) };

      // 同じ3事業を並び替えた（「道路維持事業」が先頭に来る）状態で再集約されたと想定
      const reordered = [records[2], records[0], records[1]];
      const dom2 = await readyPage(structure, reordered);
      const win2 = dom2.window;
      const file = new win2.File([JSON.stringify(payload)], 'backup.json', { type: 'application/json' });
      await win2.__app.importBackup(file);
      const restoredTarget = win2.__app.V.records.find(r => r.data.シート.事業名 === '道路維持事業');
      const restoredOther = win2.__app.V.records.find(r => r.data.シート.事業名 === '広報広聴事業');
      assertTrue(win2.__app.V.checked.has(restoredTarget.persistKey), '並び順が変わっても道路維持事業に☑が復元されるはず');
      assertEqual(win2.__app.V.memos.get(restoredTarget.persistKey), '並び順が変わっても付いてくるはず');
      assertFalse(win2.__app.V.checked.has(restoredOther.persistKey), '別の事業（旧index0）に☑が誤って付かないはず');
    });
  });

  await runSuiteAsync('viewer_app: 詳細画面', async () => {
    await testAsync('「詳細」ボタンで詳細画面に切り替わり、概要バーに表示項目の値が出る', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      win.__app.openDetail(0);
      assertFalse(isVisible(win, win.document.getElementById('list-root')), '一覧は隠れているはず');
      assertTrue(isVisible(win, win.document.getElementById('detail-root')), '詳細は実際に描画（表示）されているはず');
      assertTrue(win.document.getElementById('detail-summary-bar').textContent.includes('広報広聴事業'));
    });

    await testAsync('詳細画面の本文はGridRender.renderGridで読み取り専用描画され、値が入っている', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      win.__app.openDetail(0);
      const input = win.document.getElementById(cellIdFor(structure, 1, 2));
      assertTrue(!!input, '本文グリッドにセルが描画されているはず');
      assertEqual(input.value, '広報広聴事業');
      assertTrue(input.disabled, '住民公開ページの詳細は読み取り専用のはず');
    });

    await testAsync('前へ／次へで、現在の絞り込み・並び替え順の中を移動できる', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      win.__app.openDetail(0);
      assertTrue(win.document.getElementById('detail-prev').disabled, '先頭なのでprevは無効のはず');
      win.document.getElementById('detail-next').click();
      assertTrue(win.document.getElementById('detail-summary-bar').textContent.includes('子育て支援事業'));
    });

    await testAsync('一覧に戻ると絞り込み・並び替えの状態が保たれる', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      win.document.getElementById('viewer-keyword').value = '広報';
      win.document.getElementById('viewer-keyword').dispatchEvent(new win.Event('input'));
      win.__app.openDetail(0);
      win.document.getElementById('btn-back-to-list').click();
      assertTrue(isVisible(win, win.document.getElementById('list-root')));
      const rows = win.document.querySelectorAll('#viewer-table-root tbody tr');
      assertEqual(rows.length, 1, 'キーワードでの絞り込みが一覧復帰後も残っているはず');
    });

    await testAsync('URLハッシュ（#detail-N）で直接詳細画面を開ける', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      win.location.hash = '#detail-2';
      win.dispatchEvent(new win.Event('hashchange'));
      assertTrue(isVisible(win, win.document.getElementById('detail-root')));
      assertTrue(win.document.getElementById('detail-summary-bar').textContent.includes('道路維持事業'));
    });

    await testAsync('念のため：history.replaceStateが例外を投げても、詳細画面の表示自体は続行する', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      win.history.replaceState = () => { throw new Error('Unsafe attempt to load URL (simulated)'); };
      win.__app.openDetail(1);
      assertTrue(isVisible(win, win.document.getElementById('detail-root')), 'replaceStateが例外を投げても詳細画面へは切り替わるはず');
      assertTrue(win.document.getElementById('detail-summary-bar').textContent.includes('子育て支援事業'));
    });

    await testAsync('file://で開いた場合はURL書き換え自体を試みない（実機確認で発見：location.hash=・history.replaceStateのどちらもfile://ではブロックされうるため）', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords(), null, 'file:///C:/dummy/path/%E3%83%86%E3%82%B9%E3%83%88.html');
      const win = dom.window;
      let replaceStateCalled = false;
      win.history.replaceState = () => { replaceStateCalled = true; };
      win.__app.openDetail(1);
      assertFalse(replaceStateCalled, 'file://ではhistory.replaceStateを呼ばないはず');
      assertTrue(isVisible(win, win.document.getElementById('detail-root')), '詳細画面自体は正常に開けるはず');
      assertTrue(win.document.getElementById('detail-summary-bar').textContent.includes('子育て支援事業'));

      win.__app.backToList();
      assertFalse(replaceStateCalled, '一覧に戻る際もURL書き換えを試みないはず');
      assertTrue(isVisible(win, win.document.getElementById('list-root')));
    });
  });

  await runSuiteAsync('viewer_app: 印刷', async () => {
    await testAsync('printFilteredは、一覧に見えている表と同じ列構成（表示項目＋気になる）で1枚の表を作る', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      let printed = false;
      win.print = () => { printed = true; };
      win.__app.printFiltered();
      assertTrue(printed);
      const rows = win.document.querySelectorAll('#viewer-print-root table.viewer-print-list-table tbody tr');
      assertEqual(rows.length, 3, '☑を1件も付けていなくても、表示中の3件すべてが対象になるはず');
      // 既定では「メモを表示」がOFFなので、メモ列は無いはず（見出しがheadの2列＝表示項目2＋気になる1）
      const headers = win.document.querySelectorAll('#viewer-print-root table.viewer-print-list-table thead th');
      assertEqual(headers.length, 3, '事業名・担当課・気になる の3列のはず（メモ列は非表示）');
      win.__app.cleanupBulkPrint();
    });

    await testAsync('「気になる事業だけ表示」で絞り込んでからprintFilteredを呼ぶと、☑を付けた事業だけが表の行になる', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      win.__app.toggleChecked(win.__app.V.records[0].persistKey);
      win.__app.toggleChecked(win.__app.V.records[2].persistKey);
      win.document.getElementById('viewer-only-checked').checked = true;
      win.document.getElementById('viewer-only-checked').dispatchEvent(new win.Event('change'));
      let printed = false;
      win.print = () => { printed = true; };
      win.__app.printFiltered();
      assertTrue(printed);
      const rows = win.document.querySelectorAll('#viewer-print-root table.viewer-print-list-table tbody tr');
      assertEqual(rows.length, 2, '「気になる事業だけ表示」との組み合わせで☑2件だけが対象になるはず');
      win.__app.cleanupBulkPrint();
    });

    await testAsync('「メモを表示」をONにすると、一覧にもメモ内容の列が出て、印刷にも同じ列が含まれる', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      win.__app.setMemo(win.__app.V.records[0].persistKey, 'テストメモ本文');
      win.document.getElementById('viewer-show-memo').checked = true;
      win.document.getElementById('viewer-show-memo').dispatchEvent(new win.Event('change'));
      assertTrue(win.document.getElementById('viewer-table-root').textContent.includes('テストメモ本文'), '画面の一覧にメモ本文が見えるはず');

      let printed = false;
      win.print = () => { printed = true; };
      win.__app.printFiltered();
      assertTrue(printed);
      const headers = [...win.document.querySelectorAll('#viewer-print-root table.viewer-print-list-table thead th')].map(th => th.textContent);
      assertTrue(headers.includes('メモ内容'), '印刷にもメモ内容の列見出しがあるはず');
      assertTrue(win.document.querySelector('#viewer-print-root table.viewer-print-list-table').textContent.includes('テストメモ本文'), '印刷にもメモ本文が含まれるはず');
      win.__app.cleanupBulkPrint();
    });

    await testAsync('絞り込み結果が0件ならprintFilteredを呼んでも印刷対象は作られない', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      win.document.getElementById('viewer-keyword').value = '該当なしのキーワード';
      win.document.getElementById('viewer-keyword').dispatchEvent(new win.Event('input'));
      let printed = false;
      win.print = () => { printed = true; };
      win.__app.printFiltered();
      assertFalse(printed);
    });

    await testAsync('詳細画面の単発印刷（printSingleSnapshot）は例外を投げない', async () => {
      const structure = buildSimpleStructure();
      const dom = await readyPage(structure, buildSampleRecords());
      const win = dom.window;
      win.__app.openDetail(0);
      let printed = false;
      win.print = () => { printed = true; };
      win.document.getElementById('btn-print-detail').click();
      assertTrue(printed);
      win.__app.cleanupBulkPrint();
    });
  });

  const ok = summary();
  process.exit(ok ? 0 : 1);
})();
