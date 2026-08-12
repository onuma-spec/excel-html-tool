// フルパイプラインのスモークテスト：
// 実際にビルド済みのexcel_form_builder.html（`python assemble.py`の成果物、リポジトリに
// コミットされているファイルそのもの）を読み込み、実Excel（simple_moshikomi.xlsx）→
// ビルダー →「集約ツールを書き出す」→ 集約ツールに複数JSONを読み込ませ住民公開設定→
// 「住民公開用ページを書き出す」→ 住民公開ページで一覧・検索・詳細が動く、という
// 4ツールの受け渡し全体を1本のテストで通す。
//
// 個々のツールの単体的な結合テスト（test_builder/filler/aggregator/viewer_integration.js）
// は各ファイルを直接インラインした簡易fixtureで検証しているが、こちらは
// 「assemble.pyが実際に組み立てたファイル」を直接使うことで、テンプレート文字列注入の
// 入れ子（filler/aggregator/viewerの3階層）に置換漏れ・二重宣言等が無いかを検出する。
// このテストを実行する前に `python assemble.py` を実行し、excel_form_builder.html を
// 最新化しておくこと（builder_app.js等のソースを変更したら、まずassemble.pyを再実行する）。

const path = require('path');
const fs = require('fs');
const { JSDOM } = require('jsdom');
const { runSuiteAsync, testAsync, assertEqual, assertTrue, summary } = require('./assert_mini');

const ROOT = path.join(__dirname, '..');
const BUILDER_HTML_PATH = path.join(ROOT, 'excel_form_builder.html');
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

function mockDownloads(win) {
  win.URL.createObjectURL = () => 'blob:mock';
  win.URL.revokeObjectURL = () => {};
  win.HTMLAnchorElement.prototype.click = function () {};
  win.Element.prototype.scrollIntoView = function () {};
}

async function readBlobAsText(win, blob) {
  return new Promise((resolve, reject) => {
    const fr = new win.FileReader();
    fr.onload = () => resolve(fr.result);
    fr.onerror = reject;
    fr.readAsText(blob);
  });
}

function openHtmlPage(html) {
  const dom = new JSDOM(html, { runScripts: 'dangerously', url: 'http://localhost/' });
  mockDownloads(dom.window);
  return dom;
}

async function loadFixtureIntoBuilder(dom) {
  const win = dom.window;
  const buf = fs.readFileSync(FIXTURE);
  const file = new win.File([buf], 'simple_moshikomi.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  win.__app.handleFile(file);
  await waitFor(() => win.__app.CURRENT && win.__app.CURRENT.grid);
}

(async () => {
  await runSuiteAsync('フルパイプライン：Excel → ビルダー → 集約ツール → 住民公開ページ', async () => {
    await testAsync('excel_form_builder.htmlが存在し、ビルド時プレースホルダーが置換済みである（実行前に python assemble.py が必要）', async () => {
      assertTrue(fs.existsSync(BUILDER_HTML_PATH), 'excel_form_builder.htmlが見つかりません。先に `python assemble.py` を実行してください。');
      const html = fs.readFileSync(BUILDER_HTML_PATH, 'utf8');
      // __STRUCTURE__・__FORM_TITLE__・__RECORDS__・__PUBLIC_CONFIG__・__VIEWER_TITLE__は
      // 実行時（各ツールの「書き出す」操作）に埋め込まれる想定の実行時プレースホルダーであり、
      // ビルド成果物にJS文字列リテラルとして残っていて正しい（ここでは対象にしない）。
      ['__CORE_LOGIC__', '__GRID_RENDER__', '__BUILDER_APP__', '__SHEETJS__',
        '__FILLER_TEMPLATE_JSON__', '__AGGREGATOR_TEMPLATE_JSON__', '__AGGREGATOR_APP__',
        '__VIEWER_APP__', '__VIEWER_TEMPLATE_JSON__', '__FILLER_APP__']
        .forEach((placeholder) => {
          assertTrue(!html.includes(placeholder), `ビルド時プレースホルダー ${placeholder} が置換されずに残っている`);
        });
    });

    await testAsync('ビルダーでExcelを読み込み「集約ツールを書き出す」を押すと、集約ツールとして機能するHTMLが得られる', async () => {
      const builderDom = openHtmlPage(fs.readFileSync(BUILDER_HTML_PATH, 'utf8'));
      const builderWin = builderDom.window;
      await loadFixtureIntoBuilder(builderDom);

      let capturedBlob = null;
      builderWin.URL.createObjectURL = (blob) => { capturedBlob = blob; return 'blob:mock'; };
      builderWin.__app.doExportAsAggregator();
      assertTrue(!!capturedBlob, '集約ツールのBlobが捕捉できているはず');
      const aggregatorHtml = await readBlobAsText(builderWin, capturedBlob);
      // 集約ツール自身の`const STRUCTURE = ...;`は実体（読み込んだ様式のJSON）に
      // 置き換わっているべき。ただし住民公開ページ（ツール4）用のVIEWER_TEMPLATEは
      // この時点ではまだ雛形のまま埋め込まれているのが正しい仕様で、そちらは
      // `const STRUCTURE = /* __STRUCTURE__ */;`という同一の未置換テキストを1つ内包し続ける
      // （doExportAsViewer実行時に置き換わる）ため、「残っていないこと」ではなく
      // 「実体が書き込まれていること」を確認する。
      assertTrue(aggregatorHtml.includes('"formTitle":"simple_moshikomi"'), '集約ツール自身のSTRUCTUREに実際の様式データが注入されているはず');

      const aggDom = openHtmlPage(aggregatorHtml);
      const aggWin = aggDom.window;
      await waitFor(() => aggWin.__app && aggWin.__app.CURRENT_STATE);
      assertTrue(aggWin.__app.candidateSingles().length > 0, '集約ツールが元の様式の入力欄候補を認識できているはず');
    });

    await testAsync('集約ツールにJSONを読み込み住民公開設定をして書き出すと、一覧・検索・詳細が動く住民公開ページが得られる（4ツール通しの受け渡し）', async () => {
      const builderDom = openHtmlPage(fs.readFileSync(BUILDER_HTML_PATH, 'utf8'));
      const builderWin = builderDom.window;
      await loadFixtureIntoBuilder(builderDom);

      let aggregatorBlob = null;
      builderWin.URL.createObjectURL = (blob) => { aggregatorBlob = blob; return 'blob:mock'; };
      builderWin.__app.doExportAsAggregator();
      const aggregatorHtml = await readBlobAsText(builderWin, aggregatorBlob);

      const aggDom = openHtmlPage(aggregatorHtml);
      const aggWin = aggDom.window;
      await waitFor(() => aggWin.__app && aggWin.__app.CURRENT_STATE);

      // 各部署が入力フォーム（ツール2）で書き出したJSON相当のデータを作る。手でネスト構造を
      // 書き下ろすと（行見出しの有無で「フラットなキー」か「行見出し配下の辞書」かが変わる
      // buildRowEntry/buildSectionObjectの規則と）食い違うリスクがあるため、
      // test_filler_integration.jsのbuildSampleDataと同じく、STRUCTURE由来の使い捨てグリッドに
      // 実際にDOM入力してGridRender.collectData()で作らせる（実ロジックに作らせる）。
      const singles = aggWin.__app.candidateSingles();
      assertTrue(singles.length >= 2, 'このフィクスチャは単独の入力欄を2件以上持つはず');
      const firstId = aggWin.CoreLogic.cellId(singles[0].cells[0]);
      const firstKey = aggWin.__app.aggregateKey(singles[0]);
      // window.STRUCTURE（トップレベルconst宣言）はグローバルオブジェクトのプロパティには
      // ならないため、aggregator_app.jsが公開しているCURRENT_STATE（init()時に
      // buildStateFromStructure(STRUCTURE)した結果）をそのまま使う。
      const makeRecord = (label) => {
        const scratch = aggWin.document.getElementById('agg-scratch-root');
        const tempState = aggWin.__app.CURRENT_STATE;
        aggWin.GridRender.renderGrid(scratch, tempState, { showGear: false });
        aggWin.document.getElementById(firstId).value = label;
        const data = aggWin.GridRender.collectData(tempState);
        scratch.innerHTML = '';
        return data;
      };

      const files = [
        new aggWin.File([JSON.stringify(makeRecord('A事業'))], 'a.json', { type: 'application/json' }),
        new aggWin.File([JSON.stringify(makeRecord('B事業'))], 'b.json', { type: 'application/json' }),
        new aggWin.File([JSON.stringify(makeRecord('C事業'))], 'c.json', { type: 'application/json' }),
      ];
      await aggWin.__app.handleFileInput(files);
      assertEqual(aggWin.__app.AGG.records.length, 3);

      // 住民公開設定①表示項目に1件目の入力欄を選ぶ
      aggWin.document.getElementById('cfgdisplay_' + firstKey).click();
      aggWin.document.getElementById('cfg-title').value = 'パイプラインテスト公開ページ';
      aggWin.document.getElementById('cfg-title').dispatchEvent(new aggWin.Event('input'));

      let viewerBlob = null;
      aggWin.URL.createObjectURL = (blob) => { viewerBlob = blob; return 'blob:mock'; };
      aggWin.__app.doExportAsViewer();
      assertTrue(!!viewerBlob, '住民公開ページのBlobが捕捉できているはず');
      const viewerHtml = await readBlobAsText(aggWin, viewerBlob);
      assertTrue(viewerHtml.includes('<title>パイプラインテスト公開ページ</title>'));

      const viewerDom = openHtmlPage(viewerHtml);
      const viewerWin = viewerDom.window;
      await waitFor(() => viewerWin.__app && viewerWin.__app.V && viewerWin.document.getElementById('viewer-table-root').innerHTML.length > 0);

      const rows = viewerWin.document.querySelectorAll('#viewer-table-root tbody tr');
      assertEqual(rows.length, 3, '3件のJSONを読み込んだ集約ツールから書き出したページには3行あるはず');
      assertTrue(viewerWin.document.getElementById('viewer-table-root').textContent.includes('A事業'));

      viewerWin.document.getElementById('viewer-keyword').value = 'B事業';
      viewerWin.document.getElementById('viewer-keyword').dispatchEvent(new viewerWin.Event('input'));
      assertEqual(viewerWin.document.querySelectorAll('#viewer-table-root tbody tr').length, 1);

      viewerWin.__app.openDetail(1);
      // element.style.display（JSが直接セットした値）だけでなくgetComputedStyleで実際の
      // 描画結果を確認する。スタイルシート側に同名セレクタのdisplay:none宣言が別途あると、
      // JS側でstyle.display=''にリセットしてもそちらにフォールバックして実際には非表示のまま
      // ということがある（実機確認で発覚した不具合、viewer_template.htmlの#detail-root参照）。
      assertTrue(viewerWin.getComputedStyle(viewerWin.document.getElementById('detail-root')).display !== 'none', '詳細画面が実際に描画されているはず');
      assertTrue(viewerWin.document.getElementById('detail-summary-bar').textContent.includes('B事業'));
    });
  });

  const ok = summary();
  process.exit(ok ? 0 : 1);
})();
