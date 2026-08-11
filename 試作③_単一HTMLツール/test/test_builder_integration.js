// excel_form_builder.html（ツール1）の結合テスト。
// core_logic.js／grid_render.js／vendor/xlsx.core.min.js／builder_app.js を
// 実際のHTMLページと同じ形（インラインscript）でjsdom内に読み込み、
// window.__app（builder_app.jsが公開するデバッグ用API）越しに実際のファイル読込・
// セル設定ウィザード・手動グループ化・書き出しを駆動する。
//
// なぜ実Excelを書き込んで検証しないか：xlsx.core.min.js（無料版SheetJS）は
// セルの塗りつぶし色をwrite()で書き出せない（書いても読み直すとpatternType:'none'に
// 戻ってしまう）ことを事前確認済み。そのため塗りつぶしに依存する見出し判定を
// 正しく検証するには、実際にExcelで作られた本物のテンプレート
// （実機確認/公開テンプレ調達/simple_moshikomi.xlsx）を読み込んで使う。

const path = require('path');
const fs = require('fs');
const { JSDOM } = require('jsdom');
const { runSuiteAsync, testAsync, assertEqual, assertTrue, assertFalse, summary } = require('./assert_mini');

const ROOT = path.join(__dirname, '..');
const SRC = {
  vendor: fs.readFileSync(path.join(ROOT, 'vendor', 'xlsx.core.min.js'), 'utf8'),
  core: fs.readFileSync(path.join(ROOT, 'core_logic.js'), 'utf8'),
  grid: fs.readFileSync(path.join(ROOT, 'grid_render.js'), 'utf8'),
  builder: fs.readFileSync(path.join(ROOT, 'builder_app.js'), 'utf8'),
};
const FIXTURE = path.join(ROOT, '実機確認', '公開テンプレ調達', 'simple_moshikomi.xlsx');

// preInit（省略可）：dom.window.__app.init()を呼ぶ前にwindowへ設定を差し込みたい場合に使う
// （例：showDirectoryPickerの有無はinit()実行時に一度だけ判定されるため、それより前に
// 差し込む必要があるテストのため）。
function newBuilderPage(preInit) {
  const html = `<!doctype html><html><body>
    <div id="export-dir-bar">
      <button id="btn-pick-export-dir">pickdir</button>
      <span id="export-dir-status"></span>
    </div>
    <div id="status"></div>
    <p class="step-label" data-step-label="0"></p>
    <label id="drop-zone"><input type="file" id="file-input"></label>
    <p class="step-label" data-step-label="1"></p>
    <div data-step-detail="1"></div>
    <div id="grid-root"></div>
    <div id="selection-toolbar">
      <span id="selection-count"></span>
      <button id="btn-selection-apply">まとめて設定</button>
      <button id="btn-selection-group">グループ化する</button>
      <button id="btn-selection-clear">選択解除</button>
    </div>
    <script>${SRC.vendor}</script>
    <script>${SRC.core}</script>
    <script>${SRC.grid}</script>
    <script>${SRC.builder}</script>
  </body></html>`;
  const dom = new JSDOM(html, { runScripts: 'dangerously', url: 'http://localhost/' });
  // jsdomはURL.createObjectURLを実装していない。doExportAsForm()のBlobダウンロード
  // 部分だけが必要とするので、テストのために最小限のダミーを差し込む。
  dom.window.URL.createObjectURL = () => 'blob:mock';
  dom.window.URL.revokeObjectURL = () => {};
  // doExportAsForm()が<a href="blob:...">をclick()する。jsdomは実ナビゲーションを
  // 実装しておらず、そのままだと「Not implemented: navigation」がコンソールに出るだけの
  // ノイズになるので、テストではダウンロード用のクリックを無害化する。
  dom.window.HTMLAnchorElement.prototype.click = function () {};
  // jsdomはレイアウト計算をしないためscrollIntoViewを実装していない（実ブラウザには存在する
  // 標準API）。「このセルをグリッド上で確認する」ボタンのクリックで例外にならないよう補う。
  dom.window.Element.prototype.scrollIntoView = function () {};
  if (preInit) preInit(dom.window);
  dom.window.__app.init();
  return dom;
}

// newBuilderPage()はbuilder_app.jsを未加工のまま読み込むため、FILLER_TEMPLATE定数は
// プレースホルダー文字列"__FILLER_TEMPLATE_JSON__"のままで、実際のfiller_template.html
// の中身（<title>__FORM_TITLE__</title>等）は含まれていない。書き出されるフォームHTMLの
// 実際の中身（タイトルの反映等）を検証するテストだけは、assemble.py（build_filler_template_text
// → build_builder_html）と同じ手順でFILLER_TEMPLATEに本物のfiller_template.htmlを埋め込んだ
// builder_app.jsを組み立ててから使う。
function newBuilderPageWithRealFillerTemplate() {
  const fillerTemplate = fs.readFileSync(path.join(ROOT, 'filler_template.html'), 'utf8')
    .replace('/* __CORE_LOGIC__ */', SRC.core)
    .replace('/* __GRID_RENDER__ */', SRC.grid)
    .replace('/* __FILLER_APP__ */', fs.readFileSync(path.join(ROOT, 'filler_app.js'), 'utf8'));
  const fillerJsonLiteral = JSON.stringify(fillerTemplate).replace(/<\/(script)/gi, '<\\/$1');
  const builderWithFiller = SRC.builder.replace('"__FILLER_TEMPLATE_JSON__"', fillerJsonLiteral);

  const html = `<!doctype html><html><body>
    <div id="export-dir-bar">
      <button id="btn-pick-export-dir">pickdir</button>
      <span id="export-dir-status"></span>
    </div>
    <div id="status"></div>
    <p class="step-label" data-step-label="0"></p>
    <label id="drop-zone"><input type="file" id="file-input"></label>
    <p class="step-label" data-step-label="1"></p>
    <div data-step-detail="1"></div>
    <div id="grid-root"></div>
    <div id="selection-toolbar">
      <span id="selection-count"></span>
      <button id="btn-selection-apply">まとめて設定</button>
      <button id="btn-selection-group">グループ化する</button>
      <button id="btn-selection-clear">選択解除</button>
    </div>
    <script>${SRC.vendor}</script>
    <script>${SRC.core}</script>
    <script>${SRC.grid}</script>
    <script>${builderWithFiller}</script>
  </body></html>`;
  const dom = new JSDOM(html, { runScripts: 'dangerously', url: 'http://localhost/' });
  dom.window.URL.createObjectURL = () => 'blob:mock';
  dom.window.URL.revokeObjectURL = () => {};
  dom.window.HTMLAnchorElement.prototype.click = function () {};
  dom.window.Element.prototype.scrollIntoView = function () {};
  dom.window.__app.init();
  return dom;
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
      if (Date.now() - start > timeoutMs) return reject(new Error('waitFor timeout: ' + predicate));
      setTimeout(check, intervalMs);
    })();
  });
}

async function loadFixture(dom, fixturePath) {
  const win = dom.window;
  // CURRENT.gridは初回読込後は常に非nullのままなので、単純な非null待ちだと
  // 2回目以降の読込では「まだ前回のgridのまま」という状態を「完了」と誤判定してしまう。
  // gridオブジェクトの参照が変わったことをもって完了とみなす。
  const prevGrid = win.__app.CURRENT && win.__app.CURRENT.grid;
  const buf = fs.readFileSync(fixturePath);
  const file = new win.File([buf], path.basename(fixturePath));
  win.__app.handleFile(file);
  await waitFor(() => win.__app.CURRENT && win.__app.CURRENT.grid && win.__app.CURRENT.grid !== prevGrid);
}

function clickButton(win, text) {
  const btns = Array.from(win.document.querySelectorAll('.modal-box button'));
  const btn = btns.find(b => b.textContent.trim().startsWith(text));
  if (!btn) throw new Error(`button "${text}" not found. available: ${btns.map(b => JSON.stringify(b.textContent)).join(', ')}`);
  btn.click();
}

function setModalInput(win, id, value) {
  const el = win.document.getElementById(id);
  if (!el) throw new Error('modal input not found: #' + id);
  el.value = value;
}

(async () => {
  await runSuiteAsync('builder_app: 実ファイル読込（simple_moshikomi.xlsx）', async () => {
    await testAsync('読み込み後、CURRENTが構築されグリッドが描画される', async () => {
      const dom = newBuilderPage();
      await loadFixture(dom, FIXTURE);
      const { CURRENT } = dom.window.__app;
      assertEqual(CURRENT.mode, 'xlsx');
      assertEqual(CURRENT.maxRow, 12);
      assertEqual(CURRENT.maxCol, 6);
      assertEqual(CURRENT.sections.map(s => s.title), ['シート'], 'このテンプレートは塗り無し単独見出しが無いため単一セクションになるはず');
      const table = dom.window.document.querySelector('#grid-root table');
      assertTrue(!!table, 'グリッドのtableが描画されているべき');
      assertEqual(dom.window.document.querySelectorAll('#grid-root tbody tr').length, 12);
    });

    await testAsync('書き出し前チェックに引っかかるセルは読込時に自動でblocked化され、最終的に0件になる', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const { CURRENT } = win.__app;
      const { unreachable } = win.CoreLogic.findUnreachableCells(CURRENT.grid, CURRENT.sections, CURRENT.maxCol, CURRENT.manualGroups);
      assertEqual(unreachable.length, 0, 'autoBlockUnreachableCellsにより、読込直後の時点で未解決のunreachableは無いはず');
      assertTrue(win.document.getElementById('status').textContent.indexOf('自動的に') !== -1, '自動ブロックした旨がstatusに表示されるべき');
    });

    await testAsync('新しいファイルを読み込むとOVERRIDESがリセットされる', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const c3 = win.__app.CURRENT.grid.get('3,3');
      win.__app.OVERRIDES[win.CoreLogic.cellId(c3)] = { kind: 'textarea', dbKey: 'memo1' };
      win.__app.rebuildAndRender();
      assertTrue(Object.keys(win.__app.OVERRIDES).length > 0, '手直し後はOVERRIDESが非空のはず');

      // 同じファイルをもう一度読み込む
      await loadFixture(dom, FIXTURE);
      const remaining = Object.keys(win.__app.OVERRIDES).filter(k => win.__app.OVERRIDES[k].dbKey === 'memo1');
      assertEqual(remaining.length, 0, '再読込でmemo1の手直しは残っていないはず（自動ブロック分だけが新たに入る）');
    });
  });

  await runSuiteAsync('builder_app: セル設定ウィザード（種類のみ・項目名はデータ構造確認に一本化）', async () => {
    await testAsync('テキスト/数値/金額/見出しは、追加入力なしで種類ボタンを押した瞬間に確定する', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const c3 = win.__app.CURRENT.grid.get('3,3');
      win.__app.openCellSettings(c3);
      clickButton(win, '数値');
      assertFalse(!!win.document.getElementById('modal-overlay'), 'テキスト/数値/金額は追加ステップなしで即座にモーダルが閉じるはず');
      const override = win.__app.OVERRIDES[win.CoreLogic.cellId(c3)];
      assertEqual(override.kind, 'number');
    });

    await testAsync('種類だけを変更しても、既に設定済みの項目名(dbKey)は消えない', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const c3 = win.__app.CURRENT.grid.get('3,3');
      // データ構造確認のマッピング機能で先に項目名を設定しておく（後続テストで詳しく検証する経路）
      const id = win.CoreLogic.cellId(c3);
      win.__app.OVERRIDES[id] = { kind: 'textarea', dbKey: 'moushikomibi_memo' };
      win.__app.rebuildAndRender();

      // その後、ウィザードで種類だけ「数値」に変更する
      const target = win.__app.CURRENT.grid.get('3,3');
      win.__app.openCellSettings(target);
      clickButton(win, '数値');

      const override = win.__app.OVERRIDES[id];
      assertEqual(override.kind, 'number', '種類は変更されるべき');
      assertEqual(override.dbKey, 'moushikomibi_memo', '項目名は種類変更で消えずに引き継がれるべき');
    });

    await testAsync('数式種類は数式ステップの後に確定し、JSON出力から除外される（過去の実装バグの回帰）', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const d3 = win.__app.CURRENT.grid.get('3,4'); // 「申込日」行の別の空欄セルを数式セルに転用
      win.__app.openCellSettings(d3);
      clickButton(win, '数式');
      const formulaInput = win.document.getElementById('m-formula');
      assertTrue(!!formulaInput, '数式入力欄が表示されているべき');
      formulaInput.value = '=SUM(C3:C3)';
      clickButton(win, '保存');

      const override = win.__app.OVERRIDES[win.CoreLogic.cellId(d3)];
      assertEqual(override.kind, 'formula');
      assertEqual(override.dbKey, null, '数式セルの項目名はnull（未設定）のはず');

      const { CURRENT } = win.__app;
      const target = CURRENT.grid.get('3,4');
      assertTrue(target.isFormula, '手直し後、このセルはisFormula=trueになっているべき');
      const out = win.CoreLogic.buildSectionObject(CURRENT.grid, CURRENT.sections[0].row0, CURRENT.sections[0].row1, CURRENT.maxCol, null, CURRENT.manualGroups);
      assertFalse('col4' in out['申込日'], '数式化したセルは値として出力されないはず');
    });

    await testAsync('プルダウン種類は改行・カンマ混在の選択肢テキストを分割して保持し、選択肢入力後に即確定する', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const e3 = win.__app.CURRENT.grid.get('3,5');
      win.__app.openCellSettings(e3);
      clickButton(win, 'プルダウン');
      setModalInput(win, 'm-options', '北棟\n南棟,別館');
      clickButton(win, '保存'); // 選択肢入力の次はもう項目名ステップが無く「保存」になる

      assertFalse(!!win.document.getElementById('modal-overlay'), '選択肢を保存したら即座にモーダルが閉じるはず');
      const override = win.__app.OVERRIDES[win.CoreLogic.cellId(e3)];
      assertEqual(override.options, ['北棟', '南棟', '別館']);
      const selectEl = win.document.getElementById(win.CoreLogic.cellId(win.__app.CURRENT.grid.get('3,5')));
      assertEqual(selectEl.tagName, 'SELECT');
      assertEqual(selectEl.options.length, 4);
    });
  });

  await runSuiteAsync('builder_app: 項目名マッピング（データ構造確認・単独/グループ2ブロック）', async () => {
    function findMappingBlock(win, titlePrefix) {
      return Array.from(win.document.querySelectorAll('#check-panel .mapping-block'))
        .find(b => b.querySelector('.mapping-block-title').textContent.startsWith(titlePrefix));
    }

    await testAsync('「単独の入力欄」ブロックに一覧が表示され、CoreLogic.findMappingTargetsのsinglesと件数が一致する', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const { CURRENT } = win.__app;
      const { unreachable } = win.CoreLogic.findUnreachableCells(CURRENT.grid, CURRENT.sections, CURRENT.maxCol, CURRENT.manualGroups);
      const unreachableIds = new Set(unreachable.map(c => win.CoreLogic.cellId(c)));
      const { singles, groups } = win.CoreLogic.findMappingTargets(CURRENT.grid, CURRENT.sections, CURRENT.maxCol, CURRENT.manualGroups);
      assertTrue(singles.length > 0, 'このフィクスチャは単独の入力欄を複数持つはず');
      assertEqual(groups.length, 0, 'このフィクスチャにはグループが無いはず');

      const block = findMappingBlock(win, '単独の入力欄');
      assertTrue(!!block, '単独の入力欄ブロックが表示されているべき');
      const items = block.querySelectorAll('.mapping-item');
      assertEqual(items.length, singles.length);
      assertFalse(!!findMappingBlock(win, 'グループ化された列'), 'グループが無いのでグループブロックは表示されないはず');

      // 先頭項目（3行目「申込日」のC3セル）
      const firstItem = items[0];
      assertEqual(firstItem.querySelector('.mapping-ref').textContent, 'C3');
      assertEqual(firstItem.querySelector('.mapping-context').textContent, '申込日');
      assertEqual(firstItem.querySelector('input').placeholder, 'col3 のまま');
    });

    await testAsync('単独ブロックで名前を打ち込み保存すると反映され、触らなかった欄は変化しない（タッチ検知）', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const c3 = win.__app.CURRENT.grid.get('3,3');

      const block = findMappingBlock(win, '単独の入力欄');
      const input = block.querySelector(`#mapsingle_${win.CoreLogic.cellId(c3)}`);
      input.value = 'moushikomibi_memo';
      block.querySelector('.mapping-save-btn').click();

      const override = win.__app.OVERRIDES[win.CoreLogic.cellId(c3)];
      assertEqual(override.dbKey, 'moushikomibi_memo');

      const { CURRENT } = win.__app;
      const out = win.CoreLogic.buildSectionObject(CURRENT.grid, CURRENT.sections[0].row0, CURRENT.sections[0].row1, CURRENT.maxCol, null, CURRENT.manualGroups);
      assertEqual(out['申込日']['moushikomibi_memo'], '');
      assertFalse('col3' in out['申込日'], 'col3という自動名はもう使われないはず');
      assertTrue('col4' in out['申込日'], '触っていない欄（D3）は自動名colNを維持するはず');

      // 触らなかった欄については、保存操作全体を通じてOVERRIDESに新規エントリが増えていないこと
      const d3id = win.CoreLogic.cellId(win.__app.CURRENT.grid.get('3,4'));
      assertFalse(d3id in win.__app.OVERRIDES, 'タッチしていないセルにはOVERRIDESが作られないはず');
    });

    await testAsync('既に設定済みの項目名も一覧に表示され、書き換えて保存するとリネームできる', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const c3 = win.__app.CURRENT.grid.get('3,3');
      const id = win.CoreLogic.cellId(c3);
      win.__app.OVERRIDES[id] = { kind: 'textarea', dbKey: 'old_name' };
      win.__app.rebuildAndRender();

      const block = findMappingBlock(win, '単独の入力欄');
      const input = block.querySelector(`#mapsingle_${id}`);
      assertEqual(input.value, 'old_name', '既存の項目名が入力欄に反映されているべき');
      input.value = 'new_name';
      block.querySelector('.mapping-save-btn').click();

      assertEqual(win.__app.OVERRIDES[id].dbKey, 'new_name');
    });

    await testAsync('行内の直前見出しで名前が決まる欄も一覧に表示され、colNのまま（自動名）の欄だけ目立つ色が付く（D36/E36問題の対応）', async () => {
      // A1="事業評価"(行ラベル) B1="区分"(行内見出し) C1・D1=見出し無しの値セルが2個連続。
      // C1の項目名はB1の見出し文字("区分")由来、D1はcolNフォールバック。
      function cell(row, col, opts) {
        return Object.assign({ row, col, row2: row, col2: col, value: '', isFormula: false, formula: '', hasText: false, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null }, opts || {});
      }
      const structure = {
        formTitle: '見出し由来テスト', maxRow: 1, maxCol: 4, widths: [], heights: [],
        sections: [{ title: 'シート', row0: 1, row1: 1 }],
        cells: [
          cell(1, 1, { value: '事業評価', hasText: true }),
          cell(1, 2, { value: '区分', hasText: true }),
          cell(1, 3),
          cell(1, 4),
        ],
        manualGroups: [],
      };
      const dom = newBuilderPage();
      const win = dom.window;
      const fillerHtmlText = `<script>\nconst STRUCTURE = ${JSON.stringify(structure)};\n</script>`;
      const file = new win.File([fillerHtmlText], 'preceding_label_test.html', { type: 'text/html' });
      win.__app.handleFormHtmlFile(file);
      await waitFor(() => win.__app.CURRENT && win.__app.CURRENT.mode === 'html');

      const block = findMappingBlock(win, '単独の入力欄');
      assertTrue(!!block, '単独の入力欄ブロックが表示されているべき（以前は見出し由来のC1は非表示だった）');

      const c1id = win.CoreLogic.cellId(win.__app.CURRENT.grid.get('1,3'));
      const d1id = win.CoreLogic.cellId(win.__app.CURRENT.grid.get('1,4'));
      const c1Input = block.querySelector('#mapsingle_' + c1id);
      const d1Input = block.querySelector('#mapsingle_' + d1id);
      assertTrue(!!c1Input, '見出し由来のC1も一覧に表示されるべき');
      assertTrue(!!d1Input, 'colN由来のD1も引き続き表示されるべき');

      assertTrue(c1Input.placeholder.includes('区分'), `C1のプレースホルダーは見出し文字「区分」由来のはず: "${c1Input.placeholder}"`);
      assertTrue(d1Input.placeholder.includes('col4'), `D1のプレースホルダーはcolNのままのはず: "${d1Input.placeholder}"`);

      assertFalse(c1Input.closest('.mapping-item').classList.contains('mapping-generic'), '見出し由来で既に意味のある名前がある欄には目立たせる印を付けない');
      assertTrue(d1Input.closest('.mapping-item').classList.contains('mapping-generic'), 'colNのまま＝意味のある名前が無い欄には目立たせる印を付ける');
    });

    await testAsync('グループ化された列は1件にまとまり、名前を付けるとグループの全行に一括適用される（過去の不具合の回帰）', async () => {
      // 手動グループ（3行×2列、いずれも無名の値セル）を持つ入力フォームHTMLを模したSTRUCTUREを
      // 直接構築する。過去に実機で発見された不具合：グループ内の1行だけに名前を付けると、
      // 他の行には伝播せず配列の要素ごとにキー名がバラバラになっていた。
      function blankCell(row, col) {
        return { row, col, row2: row, col2: col, value: '', isFormula: false, formula: '', hasText: false, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null };
      }
      const cells = [];
      for (let r = 1; r <= 3; r++) { cells.push(blankCell(r, 1)); cells.push(blankCell(r, 2)); }
      const structure = {
        formTitle: 'グループテスト', maxRow: 3, maxCol: 2, widths: [], heights: [],
        sections: [{ title: 'シート', row0: 1, row1: 3 }],
        cells,
        manualGroups: [{ row0: 1, row1: 3, name: '実績グループ' }],
      };

      const dom = newBuilderPage();
      const win = dom.window;
      const fillerHtmlText = `<script>\nconst STRUCTURE = ${JSON.stringify(structure)};\n</script>`;
      const file = new win.File([fillerHtmlText], 'group_test.html', { type: 'text/html' });
      win.__app.handleFormHtmlFile(file);
      await waitFor(() => win.__app.CURRENT && win.__app.CURRENT.mode === 'html');

      const { singles, groups } = win.CoreLogic.findMappingTargets(win.__app.CURRENT.grid, win.__app.CURRENT.sections, win.__app.CURRENT.maxCol, win.__app.CURRENT.manualGroups);
      assertEqual(singles.length, 0);
      assertEqual(groups.length, 2, '2列（col1・col2）ぶんの2件にまとまるはず');

      const block = findMappingBlock(win, 'グループ化された列');
      assertTrue(!!block, 'グループブロックが表示されているべき');
      const items = block.querySelectorAll('.mapping-item');
      assertEqual(items.length, 2);
      // 「N行分」という件数表示ではどこのセルか分からないため、実際のセル範囲で示す
      // （A1:A3のように、行数ではなく実セル番地で表示することでグリッド上を探しやすくする）。
      assertEqual(items[0].querySelector('.mapping-ref').textContent, '🗂 A1:A3');
      assertEqual(items[0].querySelector('.mapping-context').textContent, '実績グループ');

      // col1の列（1件目）に名前を付けて保存する
      const col1Cells = groups.find(g => g.col === 1).cells;
      const inputId = 'mapgroup_' + win.CoreLogic.cellId(col1Cells[0]);
      block.querySelector('#' + inputId).value = 'kekka';
      block.querySelector('.mapping-save-btn').click();

      // グループの3行すべてに同じ項目名が適用されていること（=過去の不具合の解消）
      col1Cells.forEach((cell) => {
        assertEqual(win.__app.OVERRIDES[win.CoreLogic.cellId(cell)].dbKey, 'kekka', `${win.CoreLogic.cellRef(cell)}にも同じ項目名が適用されるべき`);
      });
      const { CURRENT } = win.__app;
      const out = win.CoreLogic.buildSectionObject(CURRENT.grid, CURRENT.sections[0].row0, CURRENT.sections[0].row1, CURRENT.maxCol, null, CURRENT.manualGroups);
      const records = out['実績グループ'];
      assertTrue(Array.isArray(records) && records.length === 3);
      records.forEach((rec) => assertTrue('kekka' in rec, '配列の全レコードで同じキー名になっているべき（バラバラにならない）'));
    });

    await testAsync('グループの列内で項目名が一部だけ設定済み（不統一）な場合、プレースホルダーと行の色で示す', async () => {
      function blankCell(row, col) {
        return { row, col, row2: row, col2: col, value: '', isFormula: false, formula: '', hasText: false, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null };
      }
      const cells = [];
      for (let r = 1; r <= 3; r++) { cells.push(blankCell(r, 1)); cells.push(blankCell(r, 2)); }
      const structure = {
        formTitle: '不統一テスト', maxRow: 3, maxCol: 2, widths: [], heights: [],
        sections: [{ title: 'シート', row0: 1, row1: 3 }],
        cells,
        manualGroups: [{ row0: 1, row1: 3, name: '実績グループ' }],
      };
      const dom = newBuilderPage();
      const win = dom.window;
      const fillerHtmlText = `<script>\nconst STRUCTURE = ${JSON.stringify(structure)};\n</script>`;
      const file = new win.File([fillerHtmlText], 'mixed_test.html', { type: 'text/html' });
      win.__app.handleFormHtmlFile(file);
      await waitFor(() => win.__app.CURRENT && win.__app.CURRENT.mode === 'html');

      // col1列の1行目だけ先に項目名を設定しておく（2・3行目は未設定のまま）＝不統一な状態を作る
      const row1Col1 = win.__app.CURRENT.grid.get('1,1');
      win.__app.OVERRIDES[win.CoreLogic.cellId(row1Col1)] = { kind: 'textarea', dbKey: 'kekka' };
      win.__app.rebuildAndRender();

      const block = findMappingBlock(win, 'グループ化された列');
      const { groups } = win.CoreLogic.findMappingTargets(win.__app.CURRENT.grid, win.__app.CURRENT.sections, win.__app.CURRENT.maxCol, win.__app.CURRENT.manualGroups);
      const col1Cells = groups.find(g => g.col === 1).cells;
      const inputId = 'mapgroup_' + win.CoreLogic.cellId(col1Cells[0]);
      const input = block.querySelector('#' + inputId);
      const row = input.closest('.mapping-item');

      assertEqual(input.value, '', '一致する共通の名前が無いので入力欄は空のはず');
      assertTrue(input.placeholder.includes('不統一'), `プレースホルダーに不統一である旨が出るべき: "${input.placeholder}"`);
      assertTrue(row.classList.contains('mapping-mixed'), '行に不統一を示すクラスが付くべき');

      // 触らずに保存しても、既存の設定（1行目のkekka）が壊れないこと
      block.querySelector('.mapping-save-btn').click();
      assertEqual(win.__app.OVERRIDES[win.CoreLogic.cellId(row1Col1)].dbKey, 'kekka', '触っていない不統一な列は保存しても既存設定を維持するはず');
    });
  });

  await runSuiteAsync('builder_app: 手動グループ化（複数選択→グループ化する）', async () => {
    await testAsync('選択した行範囲を手動グループとして宣言でき、JSON出力に反映される', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      // 「利用希望日」(7行目)・「利用時間帯」(8行目)の2行をまとめてグループ化する
      win.__app.SELECTED.add('7,1');
      win.__app.SELECTED.add('8,1');
      win.__app.openManualGroupModal();
      setModalInput(win, 'm-group-name', '利用details');
      clickButton(win, '保存');

      const { CURRENT } = win.__app;
      assertTrue(CURRENT.manualGroups.some(g => g.row0 === 7 && g.row1 === 8 && g.name === '利用details'), 'manualGroupsに新しいグループが追加されるべき');
      assertEqual(win.__app.SELECTED.size, 0, '保存後は選択が解除されるはず');

      const out = win.CoreLogic.buildSectionObject(CURRENT.grid, CURRENT.sections[0].row0, CURRENT.sections[0].row1, CURRENT.maxCol, null, CURRENT.manualGroups);
      assertTrue('利用details' in out, 'グループ名がJSON出力のキーになっているべき');
      // 「利用希望日」「利用時間帯」はどちらも自分自身のラベルを持つ行なので、
      // グループ化後は無名の配列ではなく「ラベル→値」の辞書としてまとめられる
      // （groupChildrenToResultの「全員ラベル付きなら辞書」分岐）。
      assertFalse(Array.isArray(out['利用details']), '子行が両方ラベル付きなので配列ではなく辞書になるはず');
      assertTrue('利用希望日' in out['利用details'] && '利用時間帯' in out['利用details']);
    });
  });

  await runSuiteAsync('builder_app: 行見出しの重複による上書き（警告文をグループ化未対応と区別する）', async () => {
    await testAsync('見出し文字が重複する行は「グループ化しても解決しません」という専用の警告になる', async () => {
      function labelCell(row, col, text) {
        return { row, col, row2: row, col2: col, value: text, isFormula: false, formula: '', hasText: true, blocked: false, fillColor: 'DDEBF7', renderType: null, renderOptions: null, dbKey: null };
      }
      function blankCell(row, col) {
        return { row, col, row2: row, col2: col, value: '', isFormula: false, formula: '', hasText: false, blocked: false, fillColor: null, renderType: null, renderOptions: null, dbKey: null };
      }
      // 「備考」という見出しが1・3行目に重複している（2行目は無関係な行）。
      // 出力時は後から処理された3行目の値で上書きされ、1行目の値セルが未到達になる。
      const cells = [
        labelCell(1, 1, '備考'), blankCell(1, 2),
        labelCell(2, 1, '達成状況'), blankCell(2, 2),
        labelCell(3, 1, '備考'), blankCell(3, 2),
      ];
      const structure = {
        formTitle: '重複見出しテスト', maxRow: 3, maxCol: 2, widths: [], heights: [],
        sections: [{ title: 'シート', row0: 1, row1: 3 }],
        cells, manualGroups: [],
      };
      const dom = newBuilderPage();
      const win = dom.window;
      const fillerHtmlText = `<script>\nconst STRUCTURE = ${JSON.stringify(structure)};\n</script>`;
      const file = new win.File([fillerHtmlText], 'dup_label_test.html', { type: 'text/html' });
      win.__app.handleFormHtmlFile(file);
      await waitFor(() => win.__app.CURRENT && win.__app.CURRENT.mode === 'html');

      // 読込直後はautoBlockUnreachableCellsが（重複であっても他の未到達セルと同様に）
      // B1を自動的に見出し（空白）へブロック済みのはず。この自動ブロックは他の様式
      // （項目名として同じ短い数値・記号が行内に大量に繰り返し出現する複雑な調査票等）で
      // 外すと逆に大量の警告が残ってしまうことが実データ検証で判明したため、意図的に
      // 全カテゴリ共通のままにしてある。ここでは警告文の出し分けロジック自体を検証したいので、
      // 手直しでB1のブロックを解除し、あえて未到達な状態を作ってから確認する。
      const { CURRENT } = win.__app;
      const b1 = CURRENT.grid.get('1,2');
      delete win.__app.OVERRIDES[win.CoreLogic.cellId(b1)];
      win.__app.rebuildAndRender();

      const { unreachable } = win.CoreLogic.findUnreachableCells(CURRENT.grid, CURRENT.sections, CURRENT.maxCol, CURRENT.manualGroups);
      assertEqual(unreachable.length, 1);
      assertEqual(win.CoreLogic.cellRef(unreachable[0]), 'B1');

      const panelText = Array.from(win.document.querySelectorAll('#check-panel .m-warn')).map(e => e.textContent).join('\n');
      assertTrue(panelText.includes('見出しの文字が他の箇所') && panelText.includes('重複している'), `専用の警告文が出るべき: "${panelText}"`);
      assertTrue(panelText.includes('B1'), 'どのセルが該当するか示されるべき');
      assertFalse(panelText.includes('グループ化されていないことが主な原因'), '未グループ化が原因という一般的な文言は出ないはず（原因が違うため）');

      // 「グループ化する」ボタン（unreachable-list）はこのケース向けには出ないこと
      // （グループ化しても重複自体は解決しないため、誤った解決策を提示しない）。
      assertEqual(win.document.querySelectorAll('#check-panel .unreachable-list .unreachable-item').length, 0);
    });
  });

  await runSuiteAsync('builder_app: 入力フォームHTMLとしての書き出し・再読込', async () => {
    await testAsync('serializeStructure()の内容が構造として妥当（cells/sections/manualGroups）', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const structure = win.__app.serializeStructure();
      assertEqual(structure.maxRow, 12);
      assertEqual(structure.maxCol, 6);
      assertEqual(structure.sections.map(s => s.title), ['シート']);
      assertTrue(Array.isArray(structure.cells) && structure.cells.length > 0);
      assertTrue(Array.isArray(structure.manualGroups));
    });

    await testAsync('doExportAsForm()は例外を投げず、書き出し完了のステータスになる（目立たせるためのstatus-flashクラスも付く）', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE); // この時点でファイル読込メッセージによりstatus-flashは既に付いている
      const status = win.document.getElementById('status');
      win.__app.doExportAsForm();
      assertEqual(status.textContent, '入力フォームを書き出しました。');
      assertTrue(status.classList.contains('status-flash'), '書き出し後もstatus-flashクラスが付いているはず');
    });

    await testAsync('STEP4のタイトル欄で書き換えると、書き出されるフォームのh1・titleとダウンロードファイル名に反映される', async () => {
      // 実際のfiller_template.htmlの中身（__FORM_TITLE__プレースホルダー）を検証したいので、
      // newBuilderPage()ではなくassemble.py相当の組み立てを行うヘルパーを使う。
      const dom = newBuilderPageWithRealFillerTemplate();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);

      const titleInput = win.document.getElementById('form-title-input');
      assertTrue(!!titleInput, 'STEP4にタイトル入力欄が表示されているべき');
      assertEqual(titleInput.value, 'simple_moshikomi', '既定値はアップロードしたファイル名（拡張子なし）のはず');

      titleInput.value = 'カスタムタイトル';
      titleInput.dispatchEvent(new win.Event('input'));
      assertEqual(win.__app.CURRENT.customTitle, 'カスタムタイトル');
      assertEqual(win.__app.serializeStructure().formTitle, 'カスタムタイトル');

      let capturedBlob = null;
      let capturedDownload = null;
      win.URL.createObjectURL = (blob) => { capturedBlob = blob; return 'blob:mock'; };
      win.HTMLAnchorElement.prototype.click = function () { capturedDownload = this.download; };
      win.__app.doExportAsForm();
      assertEqual(capturedDownload, 'カスタムタイトル_入力フォーム.html');
      const html = await new Promise((resolve, reject) => {
        const fr = new win.FileReader();
        fr.onload = () => resolve(fr.result);
        fr.onerror = reject;
        fr.readAsText(capturedBlob);
      });
      assertTrue(html.includes('<title>カスタムタイトル</title>'), 'titleタグに反映されているべき');
      assertTrue(html.includes('カスタムタイトル</h1>'), 'h1に反映されているべき');
    });

    await testAsync('ファイル名に含める項目を選ぶと、選んだ順ではなくシート上の並び順でfileNameFieldsに入る', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const { singles } = win.CoreLogic.findMappingTargets(win.__app.CURRENT.grid, win.__app.CURRENT.sections, win.__app.CURRENT.maxCol, win.__app.CURRENT.manualGroups);
      assertTrue(singles.length >= 2, 'このフィクスチャは単独の入力欄を2件以上持つはず');
      const firstId = win.CoreLogic.cellId(singles[0].cells[0]);
      const secondId = win.CoreLogic.cellId(singles[1].cells[0]);

      // わざとシート順とは逆に（2件目→1件目の順で）チェックする
      win.document.getElementById('filenamefield_' + secondId).click();
      win.document.getElementById('filenamefield_' + firstId).click();

      const structure = win.__app.serializeStructure();
      assertEqual(structure.fileNameFields, [firstId, secondId], 'クリックした順ではなく、シート上の並び順（row/col順）で並ぶべき');
    });

    await testAsync('ファイル名項目の選択も、書き出し済みフォームHTMLの再読込で引き継がれる', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const { singles } = win.CoreLogic.findMappingTargets(win.__app.CURRENT.grid, win.__app.CURRENT.sections, win.__app.CURRENT.maxCol, win.__app.CURRENT.manualGroups);
      const targetId = win.CoreLogic.cellId(singles[0].cells[0]);
      win.document.getElementById('filenamefield_' + targetId).click();

      const structure = win.__app.serializeStructure();
      assertEqual(structure.fileNameFields, [targetId]);

      const fillerHtmlText = `<script>\nconst STRUCTURE = ${JSON.stringify(structure)};\n</script>`;
      const file = new win.File([fillerHtmlText], 'filename_field_test.html', { type: 'text/html' });
      win.__app.handleFormHtmlFile(file);
      await waitFor(() => win.__app.CURRENT && win.__app.CURRENT.mode === 'html');

      assertTrue(win.__app.CURRENT.fileNameFieldIds.has(targetId), '再読込後もfileNameFieldIdsに引き継がれているべき');
      assertEqual(win.__app.serializeStructure().fileNameFields, [targetId], '再書き出し時も同じ項目がfileNameFieldsに含まれるべき');
    });

    await testAsync('必須項目・レビュー欄をチェックすると、選んだ順ではなくシート上の並び順でstructureに入る', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const { singles } = win.CoreLogic.findMappingTargets(win.__app.CURRENT.grid, win.__app.CURRENT.sections, win.__app.CURRENT.maxCol, win.__app.CURRENT.manualGroups);
      assertTrue(singles.length >= 2, 'このフィクスチャは単独の入力欄を2件以上持つはず');
      const firstId = win.CoreLogic.cellId(singles[0].cells[0]);
      const secondId = win.CoreLogic.cellId(singles[1].cells[0]);

      win.document.getElementById('requiredfield_' + firstId).click();
      win.document.getElementById('reviewfield_' + secondId).click();
      win.document.getElementById('reviewfield_' + firstId).click();

      const structure = win.__app.serializeStructure();
      assertEqual(structure.requiredFields, [firstId]);
      assertEqual(structure.reviewFields, [firstId, secondId], 'クリックした順ではなく、シート上の並び順で並ぶべき');
    });

    await testAsync('一覧表示候補項目は、ファイル名項目を自動的に候補へ含み、そのチェックは操作できない', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const { singles } = win.CoreLogic.findMappingTargets(win.__app.CURRENT.grid, win.__app.CURRENT.sections, win.__app.CURRENT.maxCol, win.__app.CURRENT.manualGroups);
      const firstId = win.CoreLogic.cellId(singles[0].cells[0]);
      const secondId = win.CoreLogic.cellId(singles[1].cells[0]);

      win.document.getElementById('filenamefield_' + firstId).click();
      // ファイル名項目のチェックは一覧表示候補ピッカーの強制チェック状態には即時反映されない
      // （どちらもrenderCheckPanelの次回描画時に反映される作り）ため、明示的に再描画する。
      win.__app.renderCheckPanel();
      const forcedCheckbox = win.document.getElementById('displaycandidatefield_' + firstId);
      assertTrue(forcedCheckbox.checked, 'ファイル名項目は自動的に一覧表示候補にもチェックされるべき');
      assertTrue(forcedCheckbox.disabled, 'ファイル名項目由来のチェックは操作できないようにするべき');

      win.document.getElementById('displaycandidatefield_' + secondId).click();
      const structure = win.__app.serializeStructure();
      assertEqual(structure.displayCandidateFields, [firstId, secondId], 'ファイル名項目＋追加選択分がシート順で入るはず');
    });

    await testAsync('必須項目・レビュー欄・一覧表示候補項目の選択も、書き出し済みフォームHTMLの再読込で引き継がれる', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const { singles } = win.CoreLogic.findMappingTargets(win.__app.CURRENT.grid, win.__app.CURRENT.sections, win.__app.CURRENT.maxCol, win.__app.CURRENT.manualGroups);
      const targetId = win.CoreLogic.cellId(singles[0].cells[0]);
      win.document.getElementById('requiredfield_' + targetId).click();
      win.document.getElementById('reviewfield_' + targetId).click();
      win.document.getElementById('displaycandidatefield_' + targetId).click();

      const structure = win.__app.serializeStructure();
      assertEqual(structure.requiredFields, [targetId]);
      assertEqual(structure.reviewFields, [targetId]);
      assertEqual(structure.displayCandidateFields, [targetId]);

      const fillerHtmlText = `<script>\nconst STRUCTURE = ${JSON.stringify(structure)};\n</script>`;
      const file = new win.File([fillerHtmlText], 'required_review_test.html', { type: 'text/html' });
      win.__app.handleFormHtmlFile(file);
      await waitFor(() => win.__app.CURRENT && win.__app.CURRENT.mode === 'html');

      assertTrue(win.__app.CURRENT.requiredFieldIds.has(targetId), '再読込後もrequiredFieldIdsに引き継がれているべき');
      assertTrue(win.__app.CURRENT.reviewFieldIds.has(targetId), '再読込後もreviewFieldIdsに引き継がれているべき');
      assertEqual(win.__app.serializeStructure().requiredFields, [targetId]);
      assertEqual(win.__app.serializeStructure().reviewFields, [targetId]);
      assertEqual(win.__app.serializeStructure().displayCandidateFields, [targetId]);
    });

    await testAsync('手直し・グループ化済みの入力フォームHTMLを読み込むと、その状態のまま続きから編集できる', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);

      const c3 = win.__app.CURRENT.grid.get('3,3');
      win.__app.OVERRIDES[win.CoreLogic.cellId(c3)] = { kind: 'textarea', dbKey: 'moushikomibi_memo' };
      win.__app.rebuildAndRender();

      win.__app.SELECTED.add('7,1');
      win.__app.SELECTED.add('8,1');
      win.__app.openManualGroupModal();
      setModalInput(win, 'm-group-name', '利用details');
      clickButton(win, '保存');

      const structure = win.__app.serializeStructure();
      // ビルダーが書き出す入力フォームHTMLと同じ形（`const STRUCTURE = {...};`）を
      // 模したテキストを作り、handleFormHtmlFileの抽出ロジックに読み込ませる。
      const fillerHtmlText = `<script>\nconst STRUCTURE = ${JSON.stringify(structure)};\nconst FORM_TITLE = "テスト";\n</script>`;
      const file2 = new win.File([fillerHtmlText], '申込書_入力フォーム.html', { type: 'text/html' });
      win.__app.handleFormHtmlFile(file2);
      await waitFor(() => win.__app.CURRENT && win.__app.CURRENT.mode === 'html');

      const { CURRENT } = win.__app;
      assertEqual(CURRENT.maxRow, 12);
      assertTrue(CURRENT.manualGroups.some(g => g.name === '利用details'), '手動グループが引き継がれているはず');
      const restored = CURRENT.grid.get('3,3');
      assertEqual(restored.dbKey, 'moushikomibi_memo', '項目名の手直しが引き継がれているはず');
    });
  });

  await runSuiteAsync('builder_app: extractStructureFromFillerHtml（波かっこの深さ判定の頑健性）', async () => {
    await testAsync('値の中に { } ; や エスケープ済み引用符が含まれても正しく抽出できる', () => {
      const dom = newBuilderPage();
      const win = dom.window;
      const tricky = {
        formTitle: '変な値のテスト',
        note: '波かっこ{と}セミコロン;を含む値。エスケープされた引用符\\"もある。',
        cells: [{ row: 1, col: 1, value: '{"nested":true}' }],
      };
      const text = `前置きテキスト\nconst STRUCTURE = ${JSON.stringify(tricky)};\nconst 続き = 1;`;
      const parsed = win.__app.extractStructureFromFillerHtml(text);
      assertEqual(parsed, tricky);
    });

    await testAsync('マーカー文字列が無ければnullを返す', () => {
      const dom = newBuilderPage();
      const win = dom.window;
      assertEqual(win.__app.extractStructureFromFillerHtml('関係ないテキスト'), null);
    });
  });

  // STEP見出しの二重管理対策（[data-step-label]・[data-step-detail]・書き出し前チェック
  // パネル側のSTEP3見出しが、いずれもSTEP_GUIDE1箇所から生成されることの回帰確認。
  await runSuiteAsync('builder_app: STEPガイド（唯一の情報源からの生成）', async () => {
    await testAsync('作業場所の[data-step-label]がSTEP1/STEP2の見出し文言で埋まる', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      const label0 = win.document.querySelector('[data-step-label="0"]').textContent;
      const label1 = win.document.querySelector('[data-step-label="1"]').textContent;
      assertTrue(label0.startsWith('STEP1'), `STEP1見出しが空のままではないか: "${label0}"`);
      assertTrue(label1.startsWith('STEP2'), `STEP2見出しが空のままではないか: "${label1}"`);
    });

    await testAsync('STEP2の[data-step-detail]に「詳しく見る」の折りたたみが生成される', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      const details = win.document.querySelector('[data-step-detail="1"] details.step-detail');
      assertTrue(!!details, 'STEP2の詳細折りたたみが生成されているべき');
      assertEqual(win.document.querySelector('[data-step-detail="1"] summary').textContent, '詳しく見る');
      assertTrue(details.querySelectorAll('li').length >= 4, '箇条書き4項目が含まれているべき');
    });

    await testAsync('書き出し前チェックパネルにSTEP3（マッピング）・STEP4（書き出し）の見出しが表示される（作業場所に見出しが無い問題の解消）', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const panel = win.document.getElementById('check-panel');
      assertEqual(panel.querySelector('h3').textContent, 'データ構造確認', 'パネル自体の見出しはSTEP番号と独立して先頭にあるはず');

      const stepLabels = Array.from(panel.querySelectorAll('p.step-label'));
      const step3 = stepLabels.find(e => e.textContent.startsWith('STEP3'));
      const step4 = stepLabels.find(e => e.textContent.startsWith('STEP4'));
      assertTrue(!!step3, 'STEP3の見出しがパネル内に無い');
      assertTrue(!!step4, 'STEP4の見出しがパネル内に無い');
      // STEP3見出しの直後にマッピングセクションが、STEP4見出しの直後にタイトル入力欄が続き、
      // その後（タイトル欄・ファイル名項目欄を経て）書き出すボタンが現れること
      assertTrue(step3.nextElementSibling.classList.contains('mapping-section'), 'STEP3の直後はマッピングセクションであるべき');
      assertTrue(step4.nextElementSibling.classList.contains('form-title-row'), 'STEP4の直後はタイトル入力欄であるべき');
      let sib = step4.nextElementSibling;
      let foundButton = false;
      while (sib) {
        if (sib.querySelector && sib.querySelector('button')) { foundButton = true; break; }
        sib = sib.nextElementSibling;
      }
      assertTrue(foundButton, 'STEP4のブロック内のどこかに書き出すボタンがあるべき');
    });
  });

  await runSuiteAsync('builder_app: 書き出し先フォルダの指定（File System Access API）', async () => {
    // jsdomのBlobには.text()が実装されていないため、FileReaderで読み取る
    // （filler側のテストと同じ手法）。
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
      const written = {};
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

    await testAsync('showDirectoryPicker非対応環境（jsdom既定）では、#export-dir-barのinline displayは変更しない（実テンプレートのCSS側のdisplay:noneに委ねる）', async () => {
      // #status・#export-dir-barはSTEP0/STEP4どちらの書き出しにも関わるSTEP横断の
      // 共通機能のため、STEP固有のcheck-panelの中ではなく画面上部に固定表示する
      // （filler_app.js側の#export-dir-barと同じ設計に統一した）。
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      assertEqual(win.document.getElementById('export-dir-bar').style.display, '');
    });

    await testAsync('showDirectoryPicker対応環境では、#export-dir-barのinline displayに明示的に"flex"がセットされる', async () => {
      // showDirectoryPickerの有無はinit()実行時に一度だけ判定されるため、
      // newBuilderPage()のpreInitフックでinit()より前に差し込む。
      const dom = newBuilderPage((win) => { win.showDirectoryPicker = async () => ({ name: 'mock' }); });
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      assertEqual(win.document.getElementById('export-dir-bar').style.display, 'flex');
    });

    await testAsync('EXPORT_DIR_HANDLE未指定なら、doExportAsForm()は従来通り同期的にダウンロードする', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      let clicked = false;
      win.HTMLAnchorElement.prototype.click = function () { clicked = true; };
      win.__app.doExportAsForm();
      assertTrue(clicked);
      assertEqual(win.document.getElementById('status').textContent, '入力フォームを書き出しました。');
    });

    await testAsync('EXPORT_DIR_HANDLEが指定されていれば、doExportAsForm()はフォルダへ直接書き込む', async () => {
      // newBuilderPage()はFILLER_TEMPLATEがプレースホルダー文字列のままなので、
      // <title>等の実際のHTML構造を確認するにはnewBuilderPageWithRealFillerTemplate()を使う
      // （既存の「STEP4のタイトル欄で書き換えると…」テストと同じ理由）。
      const dom = newBuilderPageWithRealFillerTemplate();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const dirHandle = makeMockDirHandle(win, []);
      win.__app.setExportDirHandle(dirHandle);
      const filename = `${win.__app.serializeStructure().formTitle}_入力フォーム.html`;
      win.__app.doExportAsForm();
      await waitFor(() => dirHandle.written[filename] !== undefined);
      assertTrue(dirHandle.written[filename].includes('<title>'), '書き込まれた内容はフォームHTMLであるはず');
      assertTrue(win.document.getElementById('status').textContent.includes('指定フォルダへ保存しました'));
    });

    await testAsync('同名ファイルが既にある場合、確認ダイアログでキャンセルすると上書きされない', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      const filename = `${win.__app.serializeStructure().formTitle}_入力フォーム.html`;
      const dirHandle = makeMockDirHandle(win, [filename]);
      win.__app.setExportDirHandle(dirHandle);
      win.confirm = () => false;
      win.__app.doExportAsForm();
      await waitFor(() => win.document.getElementById('status').textContent.includes('書き出しを中止しました'));
      assertEqual(dirHandle.written[filename], undefined);
    });

    await testAsync('saveBlobはEXPORT_DIR_HANDLEが指定されていれば"dir"、無ければ同期的に"download"を返す', async () => {
      const dom = newBuilderPage();
      const win = dom.window;
      await loadFixture(dom, FIXTURE);
      let syncResult = null;
      win.HTMLAnchorElement.prototype.click = function () {};
      win.__app.saveBlob(new win.Blob(['{}']), 'a.json', (r) => { syncResult = r; });
      assertEqual(syncResult, 'download', 'フォルダ未指定時は同期的にdownloadが返るはず');

      const dirHandle = makeMockDirHandle(win, []);
      win.__app.setExportDirHandle(dirHandle);
      const asyncResult = await new Promise((resolve) => {
        win.__app.saveBlob(new win.Blob(['{"x":1}']), 'b.json', resolve);
      });
      assertEqual(asyncResult, 'dir');
      assertEqual(JSON.parse(dirHandle.written['b.json']), { x: 1 });
    });
  });

  const ok = summary();
  process.exit(ok ? 0 : 1);
})();
