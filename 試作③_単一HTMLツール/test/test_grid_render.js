// grid_render.js（グリッド描画・書き出し/読込・数式再計算・貼り付け）の回帰テスト。
// jsdomでdocument/windowを用意した上で、実際のDOM操作を伴うコードパスを検証する。

const path = require('path');
const { JSDOM } = require('jsdom');
const { mockSheet } = require('./helpers/mockWs');
const { runSuite, test, assertEqual, assertTrue, assertFalse, summary } = require('./assert_mini');

const dom = new JSDOM('<!doctype html><html><body><div id="grid-root"></div></body></html>', { url: 'http://localhost/' });
global.window = dom.window;
global.document = dom.window.document;

const CoreLogic = require(path.join(__dirname, '..', 'core_logic.js'));
global.CoreLogic = CoreLogic; // grid_render.jsはUMDでrequireされず、グローバル参照を前提にしている
const GridRender = require(path.join(__dirname, '..', 'grid_render.js'));

function freshRoot() {
  document.body.innerHTML = '<div id="grid-root"></div>';
  return document.getElementById('grid-root');
}

function stateFromWs(ws, sections) {
  const { grid, maxRow, maxCol } = CoreLogic.buildGrid(ws);
  return { grid, maxRow, maxCol, widths: [], heights: [], sections: sections || CoreLogic.splitSections(grid, maxRow, maxCol), manualGroups: [] };
}

runSuite('grid_render: renderGridのrowspan/<tr>整合性（過去の表示ズレバグの回帰）', () => {
  // 「成果指標」のような縦結合ラベル(A1:A4, rowspan4)の内側に、さらに短い縦結合
  // (B2:B4, rowspan3)がある場合、4行目はどの列の「自前セル」も持たない
  // （A・B列とも別の行にアンカーがある）。かつてはこの行の<tr>自体を省略しており、
  // rowspanが後続の無関係な行にはみ出す表示バグがあった。
  test('自前セルを持たない行でも<tr>は必ず生成される（<tr>数=maxRow）', () => {
    const ws = mockSheet({
      maxRow: 4, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: '成果指標', rowspan: 4 },
        { row: 2, col: 2, v: '達成状況', rowspan: 3 },
      ],
    });
    const state = stateFromWs(ws);
    const root = freshRoot();
    GridRender.renderGrid(root, state, { showGear: false });
    const trs = root.querySelectorAll('tbody tr');
    assertEqual(trs.length, 4, '4行分の<tr>が生成されるべき（4行目が自前セルを持たなくても省略されない）');
  });

  test('結合セルのrowspan属性は実際の結合幅と一致する', () => {
    const ws = mockSheet({
      maxRow: 4, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: '成果指標', rowspan: 4 },
        { row: 2, col: 2, v: '達成状況', rowspan: 3 },
      ],
    });
    const state = stateFromWs(ws);
    const root = freshRoot();
    GridRender.renderGrid(root, state, { showGear: false });
    const tdA = document.getElementById('td_R1_C1');
    const tdB = document.getElementById('td_R2_C2');
    assertEqual(tdA.getAttribute('rowspan'), '4');
    assertEqual(tdB.getAttribute('rowspan'), '3');
  });

  test('ルーラー行・ルーラー列（行番号・列アルファベット）が生成される', () => {
    const ws = mockSheet({ maxRow: 2, maxCol: 2, cells: [{ row: 1, col: 1, v: 'A' }] });
    const state = stateFromWs(ws);
    const root = freshRoot();
    GridRender.renderGrid(root, state, { showGear: false });
    const rulerCols = Array.from(root.querySelectorAll('.ruler-col')).map(e => e.textContent);
    assertEqual(rulerCols, ['A', 'B']);
    const rulerRows = Array.from(root.querySelectorAll('.ruler-row-num')).map(e => e.textContent);
    assertEqual(rulerRows, ['1', '2']);
  });
});

runSuite('grid_render: collectData / loadDataIntoGrid の往復', () => {
  test('入力した値がJSON書き出しに反映され、読み込みで正しく復元される', () => {
    const ws = mockSheet({
      maxRow: 3, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: 'Plan' },
        { row: 2, col: 1, v: '事業名', fill: 'DDEBF7' },
        { row: 3, col: 1, v: '担当課', fill: 'DDEBF7' },
      ],
    });
    const state = stateFromWs(ws);
    state.fileName = 'test.xlsx';
    const root = freshRoot();
    GridRender.renderGrid(root, state, { showGear: false });

    const inputEvent = document.getElementById(CoreLogic.cellId(state.grid.get('2,2')));
    const inputDept = document.getElementById(CoreLogic.cellId(state.grid.get('3,2')));
    inputEvent.value = '広報広聴事業';
    inputDept.value = '広報課';

    const data = GridRender.collectData(state);
    assertEqual(data['Plan'], { 事業名: '広報広聴事業', 担当課: '広報課' });

    // 読み込み側：一旦クリアしてから復元
    inputEvent.value = '';
    inputDept.value = '';
    GridRender.loadDataIntoGrid(state, data);
    assertEqual(inputEvent.value, '広報広聴事業');
    assertEqual(inputDept.value, '広報課');
  });

  test('グループ（縦結合の繰り返し行）も書き出し→読込で往復する', () => {
    const ws = mockSheet({
      maxRow: 3, maxCol: 2,
      cells: [{ row: 1, col: 1, v: '活動実績', rowspan: 3 }],
    });
    const state = stateFromWs(ws);
    state.fileName = 'test.xlsx';
    const root = freshRoot();
    GridRender.renderGrid(root, state, { showGear: false });

    const c1 = document.getElementById(CoreLogic.cellId(state.grid.get('1,2')));
    const c2 = document.getElementById(CoreLogic.cellId(state.grid.get('2,2')));
    const c3 = document.getElementById(CoreLogic.cellId(state.grid.get('3,2')));
    c1.value = '4月実施'; c2.value = '7月実施'; c3.value = '10月実施';

    const data = GridRender.collectData(state);
    assertEqual(data['シート']['活動実績'], [{ col2: '4月実施' }, { col2: '7月実施' }, { col2: '10月実施' }]);

    c1.value = ''; c2.value = ''; c3.value = '';
    GridRender.loadDataIntoGrid(state, data);
    assertEqual([c1.value, c2.value, c3.value], ['4月実施', '7月実施', '10月実施']);
  });

  // 既知バグ（自動グループが手動グループに範囲を奪われて1行だけ残る）の修正が、
  // core_logic.js（書き出し）だけでなくgrid_render.js（読込）でも対称に効いていることの確認。
  // 対称性が崩れると、書き出しは直っていても読込側だけ古いままの値を探しに行って
  // データが復元されない、という新種のバグになりかねない。
  test('自動グループと手動グループが重なるケースも書き出し→読込で往復する（既知バグ修正の対称性確認）', () => {
    const ws = mockSheet({
      maxRow: 3, maxCol: 3,
      cells: [
        { row: 1, col: 1, v: '確認', rowspan: 3 },
        { row: 1, col: 2, v: '回答理由' },
      ],
    });
    const state = stateFromWs(ws);
    state.manualGroups = [{ row0: 2, row1: 3, name: '手動子' }];
    state.fileName = 'test.xlsx';
    const root = freshRoot();
    GridRender.renderGrid(root, state, { showGear: false });

    const reasonInput = document.getElementById(CoreLogic.cellId(state.grid.get('1,3')));
    const c2 = document.getElementById(CoreLogic.cellId(state.grid.get('2,2')));
    const c3 = document.getElementById(CoreLogic.cellId(state.grid.get('3,2')));
    reasonInput.value = '特に問題なし';
    c2.value = 'A'; c3.value = 'B';

    const data = GridRender.collectData(state);
    assertEqual(data['シート']['確認'], []);
    assertEqual(data['シート']['回答理由'], '特に問題なし');
    assertTrue(Array.isArray(data['シート']['手動子']) && data['シート']['手動子'].length === 2);

    reasonInput.value = ''; c2.value = ''; c3.value = '';
    GridRender.loadDataIntoGrid(state, data);
    assertEqual(reasonInput.value, '特に問題なし', '「回答理由」（グループ外の独立キー）が正しく復元されるはず');
    assertEqual(c2.value, 'A');
    assertEqual(c3.value, 'B');
  });

  // 「1行1見出し1値」の行がSTEP3でdbKeyリネームされた場合の読込側の対称性確認。
  // core_logic.jsのbuildRowEntryはdbKeyがあればそれを出力キーにするよう修正済みだが、
  // grid_render.js側が旧labelOfRow（A列の見出し文字だけを見る簡易版）のままだと、
  // 改名後のキーでsecDataを探せず値が復元できない、という新種のバグになりかねない。
  test('「1行1見出し1値」の行をdbKeyでリネームしても、書き出し→読込で正しく往復する', () => {
    const ws = mockSheet({
      maxRow: 1, maxCol: 2,
      cells: [{ row: 1, col: 1, v: '事業の目的', fill: 'DDEBF7' }],
    });
    const state = stateFromWs(ws);
    state.fileName = 'test.xlsx';
    const root = freshRoot();
    GridRender.renderGrid(root, state, { showGear: false });

    const target = state.grid.get('1,2');
    CoreLogic.applyOverrides(state.grid, { [CoreLogic.cellId(target)]: { kind: 'textarea', dbKey: 'purpose' } }, CoreLogic.cellId);
    GridRender.renderGrid(root, state, { showGear: false }); // applyOverrides後のdbKeyをgridへ反映

    const input = document.getElementById(CoreLogic.cellId(state.grid.get('1,2')));
    input.value = 'こういう目的です';

    const data = GridRender.collectData(state);
    assertEqual(data['シート'], { purpose: 'こういう目的です' }, '出力キーは行見出しではなくdbKey（改名後の名前）になるべき');

    input.value = '';
    GridRender.loadDataIntoGrid(state, data);
    assertEqual(input.value, 'こういう目的です', '改名後のキーで正しく値を探し出し復元できるべき');
  });
});

runSuite('grid_render: 数式（SUM）の再計算', () => {
  test('SUM範囲内の入力値が変わると数式セルの表示が再計算される', () => {
    const ws2 = mockSheet({
      maxRow: 3, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: '決算見込額' },
        { row: 2, col: 1, v: '内特財' },
        { row: 3, col: 1, v: '合計' },
        { row: 3, col: 2, f: 'SUM(B1:B2)' },
      ],
    });
    const state = stateFromWs(ws2, [{ title: 'シート', row0: 1, row1: 3 }]);
    const root = freshRoot();
    const formulaCells = GridRender.renderGrid(root, state, { showGear: false });
    assertEqual(formulaCells.length, 1);

    const b1 = document.getElementById(CoreLogic.cellId(state.grid.get('1,2')));
    const b2 = document.getElementById(CoreLogic.cellId(state.grid.get('2,2')));
    b1.value = '1000';
    b2.value = '2000';
    GridRender.recalcAllFormulas(state);

    const span = document.getElementById('calcval_' + CoreLogic.cellId(state.grid.get('3,2')));
    assertEqual(span.textContent, (1000 + 2000).toLocaleString('ja-JP'));
  });

  test('数式セル自身は合計に二重計上されない（数式セルどうしの連鎖はスキップ）', () => {
    const ws = mockSheet({
      maxRow: 3, maxCol: 2,
      cells: [
        { row: 1, col: 2, v: '' }, // 100を入れる
        { row: 2, col: 2, f: 'SUM(B1:B1)' }, // 小計（=100のはず）
        { row: 3, col: 2, f: 'SUM(B1:B2)' }, // 合計（小計を含めるとB2の数式セル自体は無視されるべき）
      ],
    });
    const state = stateFromWs(ws, [{ title: 'シート', row0: 1, row1: 3 }]);
    const root = freshRoot();
    GridRender.renderGrid(root, state, { showGear: false });
    const b1 = document.getElementById(CoreLogic.cellId(state.grid.get('1,2')));
    b1.value = '100';
    GridRender.recalcAllFormulas(state);
    const span2 = document.getElementById('calcval_' + CoreLogic.cellId(state.grid.get('2,2')));
    const span3 = document.getElementById('calcval_' + CoreLogic.cellId(state.grid.get('3,2')));
    assertEqual(span2.textContent, '100');
    // B2は数式セルなのでSUM(B1:B2)からは除外され、合計はB1の100のみになるべき
    // （evalFormulaInfoの `if (cellInfo.isFormula) continue;` の回帰確認）
    assertEqual(span3.textContent, '100');
  });
});

runSuite('grid_render: buildInputControlの入力型（過去の三項演算子コピペミスの回帰）', () => {
  test('数値型セルはtype=numberになる（type=textのままではない）', () => {
    const info = { renderType: 'number', row: 1, col: 1 };
    const control = GridRender.buildInputControl(info);
    assertEqual(control.tagName, 'INPUT');
    assertEqual(control.getAttribute('type'), 'number');
  });

  test('金額型セルはtype=textのまま、入力イベントで3桁区切りに整形される', () => {
    const info = { renderType: 'currency', row: 2, col: 1 };
    const control = GridRender.buildInputControl(info);
    document.body.appendChild(control);
    assertEqual(control.getAttribute('type'), 'text');
    control.value = 'あいう1000000';
    control.dispatchEvent(new window.Event('input', { bubbles: true }));
    assertEqual(control.value, '1,000,000');
    control.remove();
  });

  test('プルダウン型は選択肢付きのselectになる', () => {
    const info = { renderType: 'select', renderOptions: ['A', 'B', 'C'], row: 3, col: 1 };
    const control = GridRender.buildInputControl(info);
    assertEqual(control.tagName, 'SELECT');
    assertEqual(control.options.length, 4); // 未選択 + A/B/C
  });

  test('種類未指定はtextareaになる（1行/複数行を統合した設計）', () => {
    const info = { renderType: null, row: 4, col: 1 };
    const control = GridRender.buildInputControl(info);
    assertEqual(control.tagName, 'TEXTAREA');
  });
});

runSuite('grid_render: 貼り付け（タブ区切りテキスト）', () => {
  test('選択中セルを起点に、タブ・改行区切りの貼り付け内容が複数セルへ展開される', () => {
    const ws = mockSheet({
      maxRow: 2, maxCol: 2,
      cells: [], // 全セル入力欄
    });
    const state = stateFromWs(ws, [{ title: 'シート', row0: 1, row1: 2 }]);
    const root = freshRoot();
    GridRender.renderGrid(root, state, { showGear: false });
    GridRender.setupPasteHandler(state, null);

    const startCell = document.getElementById(CoreLogic.cellId(state.grid.get('1,1')));
    startCell.focus();
    assertEqual(document.activeElement, startCell, '貼り付け先の起点セルにフォーカスが当たっている前提');

    const ev = new window.Event('paste', { bubbles: true, cancelable: true });
    ev.clipboardData = { getData: () => 'a\tb\nc\td' };
    document.dispatchEvent(ev);

    assertEqual(document.getElementById(CoreLogic.cellId(state.grid.get('1,1'))).value, 'a');
    assertEqual(document.getElementById(CoreLogic.cellId(state.grid.get('1,2'))).value, 'b');
    assertEqual(document.getElementById(CoreLogic.cellId(state.grid.get('2,1'))).value, 'c');
    assertEqual(document.getElementById(CoreLogic.cellId(state.grid.get('2,2'))).value, 'd');
  });
});

if (require.main === module) {
  process.exit(summary() ? 0 : 1);
}
