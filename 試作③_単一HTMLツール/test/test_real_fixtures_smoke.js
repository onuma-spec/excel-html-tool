// 実際に配布されている公開Excelテンプレート（実機確認/公開テンプレ調達/*.xlsx）を
// ビルダーの実処理パイプライン全体（読込→自動ブロック→書き出し）に通し、
// 「どんな実在フォームでも例外を投げず、書き出し前チェックが必ず0件に収束する」という
// アプリの根幹の保証（autoBlockUnreachableCellsの安全網）を横断的に確認する回帰テスト。
// 個々のセル値の厳密な突き合わせはしない（それはtest_builder_integration.jsの役割）。
//
// このテストがカバーする実在ファイルは、以前のセッションでPlaywright実機確認済みの
// ものを含め、`実機確認/公開テンプレ調達/`にまとめて置かれている（内閣官房・香川県・
// 総務省・厚労省・自作サンプル）。手動で都度ブラウザ確認する代わりに、ここで
// 自動回帰の対象にする（[[project-gyosei-data-publish-tool-idea]] 2026-08-09メモの
// 「残り2ファイルの実機検証は未実施」を、厳密な実機確認ではなく自動スモークテストとして
// 最低限カバーする位置づけ）。

const path = require('path');
const fs = require('fs');
const { JSDOM } = require('jsdom');
const { runSuiteAsync, testAsync, assertTrue, assertEqual, summary } = require('./assert_mini');

const ROOT = path.join(__dirname, '..');
const SRC = {
  vendor: fs.readFileSync(path.join(ROOT, 'vendor', 'xlsx.core.min.js'), 'utf8'),
  core: fs.readFileSync(path.join(ROOT, 'core_logic.js'), 'utf8'),
  grid: fs.readFileSync(path.join(ROOT, 'grid_render.js'), 'utf8'),
  builder: fs.readFileSync(path.join(ROOT, 'builder_app.js'), 'utf8'),
};
const FIXTURE_DIR = path.join(ROOT, '実機確認', '公開テンプレ調達');

function waitFor(predicate, timeoutMs, intervalMs) {
  timeoutMs = timeoutMs || 5000;
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

function newBuilderPage() {
  const html = `<!doctype html><html><body>
    <div id="export-dir-bar">
      <button id="btn-pick-export-dir">pickdir</button>
      <span id="export-dir-status"></span>
    </div>
    <div id="status"></div>
    <label id="drop-zone"><input type="file" id="file-input"></label>
    <div id="grid-root"></div>
    <div id="selection-toolbar">
      <span id="selection-count"></span>
      <button id="btn-selection-apply"></button>
      <button id="btn-selection-group"></button>
      <button id="btn-selection-clear"></button>
    </div>
    <script>${SRC.vendor}</script>
    <script>${SRC.core}</script>
    <script>${SRC.grid}</script>
    <script>${SRC.builder}</script>
  </body></html>`;
  const dom = new JSDOM(html, { runScripts: 'dangerously', url: 'http://localhost/' });
  dom.window.URL.createObjectURL = () => 'blob:mock';
  dom.window.URL.revokeObjectURL = () => {};
  dom.window.HTMLAnchorElement.prototype.click = function () {};
  dom.window.__app.init();
  return dom;
}

async function loadFixture(dom, fixturePath) {
  const win = dom.window;
  const prevGrid = win.__app.CURRENT && win.__app.CURRENT.grid;
  const buf = fs.readFileSync(fixturePath);
  const file = new win.File([buf], path.basename(fixturePath));
  win.__app.handleFile(file);
  await waitFor(() => win.__app.CURRENT && win.__app.CURRENT.grid && win.__app.CURRENT.grid !== prevGrid);
}

(async () => {
  const files = fs.existsSync(FIXTURE_DIR)
    ? fs.readdirSync(FIXTURE_DIR).filter(f => f.toLowerCase().endsWith('.xlsx'))
    : [];

  await runSuiteAsync('実在テンプレートのスモークテスト（読込→自動ブロック→書き出し）', async () => {
    if (files.length === 0) {
      await testAsync('フィクスチャフォルダが見つかった（0件は異常）', () => {
        assertTrue(false, `${FIXTURE_DIR} に.xlsxが1件も見つからなかった`);
      });
      return;
    }

    for (const fname of files) {
      await testAsync(`${fname}: 読込〜書き出し前チェック0件〜HTML書き出しまで例外なく完了する`, async () => {
        const dom = newBuilderPage();
        const win = dom.window;
        await loadFixture(dom, path.join(FIXTURE_DIR, fname));

        const { CURRENT } = win.__app;
        assertEqual(CURRENT.mode, 'xlsx');
        assertTrue(CURRENT.maxRow > 0 && CURRENT.maxCol > 0, '行・列数が正の値であるべき');
        assertTrue(CURRENT.sections.length >= 1, 'セクションが最低1つはあるべき');

        const { unreachable } = win.CoreLogic.findUnreachableCells(CURRENT.grid, CURRENT.sections, CURRENT.maxCol, CURRENT.manualGroups);
        assertEqual(unreachable.length, 0, '読込直後の自動ブロックにより、未解決のunreachableセルは無いはず（autoBlockUnreachableCellsの安全網）');

        // 現在の設定でJSON出力を組み立てても例外が起きないこと（実データが無い状態なので
        // 値の中身は検証しない。構造の組み立て自体が壊れていないかだけを見る）。
        const out = {};
        CURRENT.sections.forEach(sec => {
          out[sec.title] = win.CoreLogic.buildSectionObject(CURRENT.grid, sec.row0, sec.row1, CURRENT.maxCol, null, CURRENT.manualGroups);
        });
        assertTrue(Object.keys(out).length === CURRENT.sections.length);

        // 入力フォームとして書き出しても例外が起きないこと。
        win.__app.doExportAsForm();
        assertEqual(win.document.getElementById('status').textContent, '入力フォームを書き出しました。');

        // 書き出した構造を、フォームHTMLとして再読込しても例外が起きないこと
        // （extractStructureFromFillerHtmlの波かっこマッチングが実データでも壊れないか）。
        const structure = win.__app.serializeStructure();
        const fillerText = `<script>\nconst STRUCTURE = ${JSON.stringify(structure)};\n</script>`;
        const parsed = win.__app.extractStructureFromFillerHtml(fillerText);
        assertTrue(!!parsed && Array.isArray(parsed.cells) && parsed.cells.length === structure.cells.length);
      });
    }
  });

  const ok = summary();
  process.exit(ok ? 0 : 1);
})();
