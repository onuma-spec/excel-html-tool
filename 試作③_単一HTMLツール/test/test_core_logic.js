// core_logic.js（構造抽出・セクション分割・グループ化・書き出し前チェック）の回帰テスト。
// DOM不要の純ロジックなのでNode単体で動く。実Excelを介さず、SheetJS風のワークシート
// オブジェクトを直接組み立てて（mockSheet）、buildGrid以降の挙動を検証する。

const path = require('path');
const CoreLogic = require(path.join(__dirname, '..', 'core_logic.js'));
const { mockSheet } = require('./helpers/mockWs');
const { runSuite, test, assertEqual, assertTrue, assertFalse } = require('./assert_mini');

function extract(ws, manualGroups) {
  return CoreLogic.extractStructure(ws, manualGroups || []);
}

runSuite('core_logic: 基本の行・セクション分解', () => {
  // 見出し（塗りつぶし無し・単独行・1列目のみ値あり）が1つもない場合は
  // 全体を単一セクション「シート」として扱う（フォールバック）。
  test('見出し行が無ければ単一セクション「シート」にフォールバックする', () => {
    const ws = mockSheet({
      maxRow: 2, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: '申込日', fill: 'DDEBF7' },
      ],
    });
    const { sections } = extract(ws);
    assertEqual(sections.map(s => s.title), ['シート']);
  });

  test('塗りつぶし無しの単独行だけが見出し（セクション区切り）として検出される', () => {
    const ws = mockSheet({
      maxRow: 6, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: 'Plan' },              // 見出し（塗りなし）
        { row: 2, col: 1, v: '事業名', fill: 'DDEBF7' }, // 通常のラベル行（塗りあり）
        { row: 4, col: 1, v: 'Do' },                 // 見出し（塗りなし）
        { row: 5, col: 1, v: '活動内容', fill: 'DDEBF7' },
      ],
    });
    const { sections } = extract(ws);
    assertEqual(sections.map(s => s.title), ['Plan', 'Do']);
  });

  test('単純なラベル→値の行がセクションオブジェクトに正しく入る（値セルが空欄の場合）', () => {
    const ws = mockSheet({
      maxRow: 2, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: 'Plan' },
        { row: 2, col: 1, v: '担当課', fill: 'DDEBF7' },
        // 2行目のB列は空欄の入力セル
      ],
    });
    const { data } = extract(ws);
    assertEqual(data['Plan'], { 担当課: '' });
  });

  // 値セル自身がテキストを持つ場合（B2に既に「広報広聴事業」という記入済みの値がある）は、
  // 「もう1つのラベルが並んでいるだけの見出し行」（判定／判定理由のようなパターン）との
  // 区別がテキストの有無だけでは付かない。明示的な手直し（項目名の設定・種類の変更）が
  // 無い限りは実データなしとして扱い、出力から除外する（既知の限界。buildRowEntry参照）。
  test('値セルにテキストがあっても、明示的な手直しが無ければ実データとして扱わない（判定／判定理由問題の対策）', () => {
    const ws = mockSheet({
      maxRow: 2, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: 'Plan' },
        { row: 2, col: 1, v: '事業名', fill: 'DDEBF7' },
        { row: 2, col: 2, v: '広報広聴事業' }, // 記入済みの値 or 隣接するラベルの残骸か区別不能
      ],
    });
    const { data } = extract(ws);
    assertEqual(data['Plan'], {}, '明示的な手直しが無い限り出力しない');
  });

  test('値セルにテキストがあっても、STEP2で明示的に種類を設定すれば実データとして出力される', () => {
    const ws = mockSheet({
      maxRow: 2, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: 'Plan' },
        { row: 2, col: 1, v: '事業名', fill: 'DDEBF7' },
        { row: 2, col: 2, v: '広報広聴事業' },
      ],
    });
    const { grid, sections, maxCol } = extract(ws);
    const target = grid.get('2,2');
    CoreLogic.applyOverrides(grid, { [CoreLogic.cellId(target)]: { kind: 'textarea' } }, CoreLogic.cellId);
    // applyOverridesで入力欄に変更すると、Excel由来の元のテキストは破棄され
    // （生成されるフォームでは実際に入力可能な空欄になるべきため）、hasTextもfalseになる。
    // よってプレビュー用のdefaultGetValueでは値は空文字列になるが、重要なのは
    // 「事業名」というキー自体がちゃんと出力に含まれるようになった点（除外されない）。
    const out = CoreLogic.buildSectionObject(grid, sections[0].row0, sections[0].row1, maxCol, undefined, []);
    assertEqual(out, { 事業名: '' }, '明示的な手直し後はキーが出力に含まれるべき（除外されなくなる）');
  });

  test('縦結合の見出し行（複数行にまたがる）は見出し扱いされない', () => {
    // 1列目が複数行にまたがる縦結合の場合、見出し行の条件「単独行」に反するため
    // セクション区切りにはならない（自動グループのラベルとして扱われる別ロジック）。
    const ws = mockSheet({
      maxRow: 3, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: 'R6年度', rowspan: 2 },
        { row: 1, col: 2, v: '' },
        { row: 2, col: 2, v: '' },
      ],
    });
    const { sections } = extract(ws);
    assertEqual(sections.map(s => s.title), ['シート']);
  });
});

runSuite('core_logic: 自動グループ化（縦結合ラベル）', () => {
  test('縦結合ラベルの子行が無名なら配列としてまとめられる', () => {
    const ws = mockSheet({
      maxRow: 4, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: '活動実績', rowspan: 3 }, // 1列目3行分の縦結合ラベル
        // row1-3のB列は空欄の入力セル（無名の繰り返し行）
      ],
    });
    const { data } = extract(ws);
    assertEqual(data['シート']['活動実績'], [{ col2: '' }, { col2: '' }, { col2: '' }]);
  });

  // 「無名の繰り返し行」と「その他」のような名前付きの単発行が混在するケース。
  // 過去に名前付きの子を配列に混ぜて捨てていたバグの回帰テスト（groupChildrenToResult）。
  // maxCol=3のため、無名行（1〜3行目）もC列を持ち{col2,col3}の2キー辞書になる
  // （矩形グリッドである以上、4行目に合わせて全行が同じ列数を持つのは仕様どおり）。
  test('繰り返し行に「その他」等の名前付き行が混じっても、データを失わずグループ外に出す', () => {
    const ws = mockSheet({
      maxRow: 4, maxCol: 3,
      cells: [
        { row: 1, col: 1, v: '活動実績', rowspan: 4 },
        // row1-3: B・C列に値なし（無名の繰り返し行）
        { row: 4, col: 2, v: 'その他' },   // 4行目だけラベル付き
        // row4: C列が値セル
      ],
    });
    const { data } = extract(ws);
    assertEqual(data['シート']['活動実績'], [
      { col2: '', col3: '' }, { col2: '', col3: '' }, { col2: '', col3: '' },
    ]);
    assertEqual(data['シート']['その他'], '');
  });

  test('自動グループの子行の値セルにdbKeyを設定すると、そのキー名で出力される', () => {
    const ws = mockSheet({
      maxRow: 3, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: '活動実績', rowspan: 3 },
      ],
    });
    const { grid, sections, maxCol } = extract(ws);
    const target = grid.get('1,2'); // 1行目B列の入力セル
    CoreLogic.applyOverrides(grid, { [CoreLogic.cellId(target)]: { kind: 'textarea', dbKey: 'memo' } }, CoreLogic.cellId);
    const out = CoreLogic.buildSectionObject(grid, sections[0].row0, sections[0].row1, maxCol, null, []);
    assertEqual(out['活動実績'][0], { memo: '' });
  });
});

runSuite('core_logic: 手動グループ化', () => {
  // splitSections（見出し検出）を介さず、buildGridの結果に直接buildSectionObjectを
  // 適用する（見出し検出の副作用に左右されない、純粋なグループ化ロジックの単体テスト）。
  test('手動グループは1行からでも宣言でき、名前付き辞書またはラベルとして出力される', () => {
    const ws = mockSheet({
      maxRow: 1, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: '備考', fill: 'DDEBF7' },
      ],
    });
    const { grid, maxCol } = CoreLogic.buildGrid(ws);
    const manualGroups = [{ row0: 1, row1: 1, name: 'メモ欄' }];
    const out = CoreLogic.buildSectionObject(grid, 1, 1, maxCol, null, manualGroups);
    // 子が1件でもラベル付きなら辞書化される
    assertEqual(out['メモ欄'], { 備考: '' });
  });

  test('手動グループは縦結合が無い範囲でも、無名の子行を配列にまとめられる', () => {
    const ws = mockSheet({
      maxRow: 3, maxCol: 2,
      cells: [
        // 縦結合なし。3行とも1列目・2列目が普通の入力セル（ラベル無し）
      ],
    });
    const { grid, maxCol } = CoreLogic.buildGrid(ws);
    const manualGroups = [{ row0: 1, row1: 3, name: '繰り返し項目' }];
    const out = CoreLogic.buildSectionObject(grid, 1, 3, maxCol, null, manualGroups);
    // 手動グループはfromCol=1（1列目もそのまま値として扱う）なので、
    // 自動検出（fromCol=2でグループラベル列を消費する）とはキー構成が異なる。
    assertEqual(out['繰り返し項目'], [{ col1: '', col2: '' }, { col1: '', col2: '' }, { col1: '', col2: '' }]);
  });

  // 手動グループはfindManualGroupがisAutoGroupStartより先に判定されるため、u0が一致する
  // 範囲では必ず手動側のキー名が採用される。ただし手動グループはfromCol=1（1列目を
  // 消費しない）なので、縦結合ラベルの先頭行だけは「ラベル文字→隣のセル」という
  // 通常の単発行としても解釈され、グループ名とは別に元のラベル文字がもう1つの
  // トップレベル項目として漏れ出す（これは既知の仕様上の癖であり、バグではない：
  // 縦結合列に手動グループを重ねる場合は、事前に「自動グループ化を解除する」で
  // 縦結合自体の効果を無効化してから使うのが正しい運用）。
  test('手動グループは同じ範囲の自動検出（縦結合）より優先してキー名を決める', () => {
    const ws = mockSheet({
      maxRow: 3, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: '活動実績', rowspan: 3 },
      ],
    });
    const { grid, maxCol } = CoreLogic.buildGrid(ws);
    const manualGroups = [{ row0: 1, row1: 3, name: '手動優先' }];
    const out = CoreLogic.buildSectionObject(grid, 1, 3, maxCol, null, manualGroups);
    assertTrue(Array.isArray(out['手動優先']), '手動グループのキー名で配列が出力されるべき');
    assertEqual(out['活動実績'], '', '縦結合ラベル文字自体は先頭行の値として別途漏れ出す（既知の仕様）');
  });

  test('自動グループ化の解除（disabled）は縦結合ラベルを先頭行だけの通常ラベル行として扱う', () => {
    const ws = mockSheet({
      maxRow: 2, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: '活動実績', rowspan: 2 },
      ],
    });
    const { grid, maxCol } = CoreLogic.buildGrid(ws);
    const manualGroups = [{ row0: 1, row1: 2, name: null, disabled: true }];
    const out = CoreLogic.buildSectionObject(grid, 1, 2, maxCol, null, manualGroups);
    // グループとしては出力されない（配列ではない）が、先頭行はラベル→値として残る。
    // 2行目の値セルは、どのキーにも対応付かず出力から漏れる
    // （findUnreachableCellsで検出できることを別テストで確認済み）。
    assertEqual(out, { '活動実績': '' });
  });
});

runSuite('core_logic: 数式セルの除外', () => {
  test('数式セルはdbKeyを設定しても値としては出力されない（合計行の混入防止）', () => {
    const ws = mockSheet({
      maxRow: 2, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: '合計' },
        { row: 1, col: 2, f: 'SUM(A1:A1)' },
      ],
    });
    const { grid, sections, maxCol } = extract(ws);
    const target = grid.get('1,2');
    assertTrue(target.isFormula, '数式セルとして認識されているはず');
    CoreLogic.applyOverrides(grid, { [CoreLogic.cellId(target)]: { kind: 'formula', dbKey: 'total', formula: '=SUM(A1:A1)' } }, CoreLogic.cellId);
    const out = CoreLogic.buildSectionObject(grid, sections[0].row0, sections[0].row1, maxCol, null, []);
    // 数式セルが値として出力されないので、「合計」行は値の無いラベルのみの行となり、
    // rest.length===0によりトップレベルには一切出力されない。
    assertFalse('合計' in out, '数式セルしか無い行はトップレベルに出力されないはず');
    assertFalse('total' in out, 'dbKeyを設定しても数式セルは出力対象にならないはず');
  });
});

runSuite('core_logic: 見出し文字を空欄にする手直し（blocked）', () => {
  test('見出しを空欄にすると入力不可セル(blocked)になり、幽霊セクションを生まない', () => {
    const ws = mockSheet({
      maxRow: 2, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: '仕切り' }, // 元は見出し扱い候補（塗りなし単独行）
        { row: 2, col: 1, v: '事業名', fill: 'DDEBF7' },
      ],
    });
    const { grid } = CoreLogic.buildGrid(ws);
    const target = grid.get('1,1');
    CoreLogic.applyOverrides(grid, { [CoreLogic.cellId(target)]: { kind: 'label', label: '' } }, CoreLogic.cellId);
    assertFalse(target.hasText, '空欄見出しはhasTextがfalseのまま');
    assertTrue(target.blocked, 'blockedフラグが立つ');

    const sections = CoreLogic.splitSections(grid, 2, 2);
    // 1行目が見出しとして検出されなくなったので、「仕切り」というセクションは生まれない
    assertFalse(sections.some(s => s.title === '仕切り'), '空欄化した見出しがセクションとして残ってはいけない');
  });

  test('label override に空文字列以外を渡すと、旧「||」バグのように意図しない文字列に化けない', () => {
    // 過去のバグ: `ov.label || String(...||'ラベル')` で空文字列が偽値とみなされ
    // 「ラベル」という文字列に化けていた。正しい実装ではlabelそのものが保持される。
    const ws = mockSheet({ maxRow: 1, maxCol: 1, cells: [{ row: 1, col: 1, v: '旧見出し' }] });
    const { grid } = CoreLogic.buildGrid(ws);
    const target = grid.get('1,1');
    CoreLogic.applyOverrides(grid, { [CoreLogic.cellId(target)]: { kind: 'label', label: '新見出し' } }, CoreLogic.cellId);
    assertEqual(target.value, '新見出し');
    assertFalse(target.blocked, '文字が入っている場合はblockedにならない');
  });
});

runSuite('core_logic: 書き出し前チェック（findUnreachableCells）', () => {
  test('見出しより前（どのセクションにも属さない）の入力セルはoutOfSection扱いで検出される', () => {
    const ws = mockSheet({
      maxRow: 3, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: '' }, // 見出しより前の孤立した入力セル（見出しでは無い＝値なし）
        { row: 2, col: 1, v: 'Plan' },
        { row: 3, col: 1, v: '事業名', fill: 'DDEBF7' },
      ],
    });
    const { grid, sections, maxCol } = extract(ws);
    const { unreachable } = CoreLogic.findUnreachableCells(grid, sections, maxCol, []);
    const outOfSectionCells = unreachable.filter(c => c.outOfSection);
    assertTrue(outOfSectionCells.length >= 1, '見出し以前の入力セルがoutOfSectionとして検出されるべき');
  });

  test('グループ化されていない無名の繰り返し行はunreachableとして検出され、グループ化すると解消する', () => {
    // maxCol=1にして、見出し行（Do、1列目のみ）に余計な空欄列を持たせない
    // （2列目以降があると見出し行自体もoutOfSectionなunreachableとして混ざり、
    // 「グループ化で解消するかどうか」の検証がぼやけるため）。
    const ws = mockSheet({
      maxRow: 3, maxCol: 1,
      cells: [
        { row: 1, col: 1, v: 'Do' },
        // row2,3: ラベル無しの単発行（本来は繰り返しデータだが未グループ化）
      ],
    });
    const { grid, sections, maxCol } = extract(ws);
    const before = CoreLogic.findUnreachableCells(grid, sections, maxCol, []);
    assertTrue(before.unreachable.length > 0, 'グループ化前はunreachableが検出されるべき');

    const manualGroups = [{ row0: 2, row1: 3, name: '繰り返し' }];
    const after = CoreLogic.findUnreachableCells(grid, sections, maxCol, manualGroups);
    assertEqual(after.unreachable.length, 0, 'グループ化すれば解消するはず');
  });

  test('数式セル・見出しセル・blockedセルは入力対象外なのでunreachableに含まれない', () => {
    const ws = mockSheet({
      maxRow: 2, maxCol: 3,
      cells: [
        // 見出し行は全列に横結合し、隣の空欄列がoutOfSectionな孤立セルとして
        // 混ざらないようにする（実際のテンプレートでも見出し行は全列結合が通例）。
        { row: 1, col: 1, v: 'Plan', colspan: 3 },
        { row: 2, col: 1, v: '合計', fill: 'DDEBF7' }, // 塗りありにして見出し誤検出を防ぐ
        { row: 2, col: 2, f: 'SUM(A2:A2)' },
        { row: 2, col: 3, v: '固定ラベル' },
      ],
    });
    const { grid, sections, maxCol } = extract(ws);
    assertEqual(sections.map(s => s.title), ['Plan'], '合計行が誤って新しい見出しにならないこと');
    const { unreachable } = CoreLogic.findUnreachableCells(grid, sections, maxCol, []);
    assertEqual(unreachable.length, 0, '数式・ラベルのみの行にunreachableな入力セルは無いはず');
  });
});

runSuite('core_logic: 自動グループと手動グループの重なり（既知バグの修正）', () => {
  // [[project-gyosei-data-publish-tool-idea]] の2026-08-09メモに記録されていた既知の不具合。
  // 自動検出グループ（縦結合）と手動グループが同じ範囲で重なり、手動グループが範囲の
  // 後半だけを奪うと、自動側に「見出し由来の1エントリだけの辞書」という残骸が残っていた
  // （本来は配列であるべき箇所が {ラベル: 値} という辞書になってしまっていた）。
  //
  // 修正：groupChildrenToResultにrequireMultipleオプションを追加し、自動検出グループ
  // からの呼び出しだけ「子が2件以上無ければ辞書化しない」よう変更。1件だけ残った
  // ラベル付きの子は、手動グループを使わなかった場合（＝縦結合の全範囲がそのまま
  // 自動グループとして処理された場合）と同じ経路（extraEntries）で、グループの外側の
  // 独立キーとして出力されるようになった。
  test('自動グループが手動グループに範囲を奪われて1行だけ残っても、辞書化せず独立キーとして分離される', () => {
    const ws = mockSheet({
      maxRow: 3, maxCol: 3,
      cells: [
        { row: 1, col: 1, v: '確認', rowspan: 3 }, // 縦結合ラベル：1〜3行目
        { row: 1, col: 2, v: '回答理由' },          // 自動グループに残る1行目：ラベル付きの残骸
        // row1のC列は入力セル
        // row2,3のB列は入力セル（手動グループ側で処理される）
      ],
    });
    const { grid, sections, maxCol } = extract(ws);
    // 手動グループが2〜3行目だけを奪う→自動グループ側には1行目だけが残る
    const manualGroups = [{ row0: 2, row1: 3, name: '手動子' }];
    const out = CoreLogic.buildSectionObject(grid, sections[0].row0, sections[0].row1, maxCol, null, manualGroups);
    // 修正後の挙動：「確認」は（残った実データが無いので）空配列、
    // 1行目の値は「回答理由」という独立したキーとしてグループの外側に出力される。
    assertEqual(out['確認'], []);
    assertEqual(out['回答理由'], '');
    assertTrue(Array.isArray(out['手動子']) && out['手動子'].length === 2, '手動グループ側は影響を受けず2件の配列のまま');
  });

  // 一方、手動グループが最初から1行だけを対象に宣言する運用（例：備考欄1行をメモ欄として
  // 辞書化する）は、requireMultipleの対象外なので従来どおり辞書化される（退行していないことの確認）。
  test('手動グループが最初から1行だけの場合は、引き続き辞書として出力される（退行していないこと）', () => {
    const ws = mockSheet({
      maxRow: 1, maxCol: 2,
      cells: [{ row: 1, col: 1, v: '備考', fill: 'DDEBF7' }],
    });
    const { grid, maxCol } = CoreLogic.buildGrid(ws);
    const manualGroups = [{ row0: 1, row1: 1, name: 'メモ欄' }];
    const out = CoreLogic.buildSectionObject(grid, 1, 1, maxCol, null, manualGroups);
    assertEqual(out['メモ欄'], { 備考: '' });
  });

  // 自動グループが手動グループと重ならず、素直に複数行のまま残る場合は今まで通り
  // 辞書として出力されることの確認（requireMultipleを付けても正常系は壊れていない）。
  test('自動グループが手動グループと重ならず2行以上のラベル付き子を持つ場合は、引き続き辞書として出力される', () => {
    // maxCol=3にする理由：fromCol=2（グループの1列目を消費した残り）で
    // 「行自身のラベル(col2)＋値セル(col3)」の2セルが無いと、rest.length===0で
    // そもそも実データ無し扱いになってしまうため（col2だけだとラベルの置き場しかない）。
    const ws = mockSheet({
      maxRow: 2, maxCol: 3,
      cells: [
        { row: 1, col: 1, v: '確認', rowspan: 2 },
        { row: 1, col: 2, v: '理由A' },
        { row: 2, col: 2, v: '理由B' },
      ],
    });
    const { grid, sections, maxCol } = extract(ws);
    const out = CoreLogic.buildSectionObject(grid, sections[0].row0, sections[0].row1, maxCol, null, []);
    assertEqual(out['確認'], { 理由A: '', 理由B: '' });
  });
});

runSuite('core_logic: defaultKeyFor（項目名の決定規則）', () => {
  test('dbKeyが最優先、次に行内の直前ラベル、どちらも無ければcolN', () => {
    const infoWithDbKey = { dbKey: 'foo', col: 5 };
    assertEqual(CoreLogic.defaultKeyFor(infoWithDbKey, '区分'), 'foo');
    const infoNoDbKey = { dbKey: null, col: 5 };
    assertEqual(CoreLogic.defaultKeyFor(infoNoDbKey, '区分'), '区分');
    assertEqual(CoreLogic.defaultKeyFor(infoNoDbKey, null), 'col5');
  });
});

runSuite('core_logic: findMappingTargets（項目名マッピング対象セルの検出・グループ列単位の集約）', () => {
  test('グループ外のセルはsingles、グループ内のセルは列単位でgroupsにまとめられる', () => {
    const ws = mockSheet({
      maxRow: 3, maxCol: 3,
      cells: [
        { row: 1, col: 1, v: '成果指標', rowspan: 2 }, // 自動グループ：1〜2行目
        { row: 3, col: 1, v: '達成状況', fill: 'DDEBF7' }, // 通常のラベル行
      ],
    });
    const { grid, sections, maxCol } = extract(ws);
    const { singles, groups } = CoreLogic.findMappingTargets(grid, sections, maxCol, []);

    // 達成状況行のB3・C3はグループ外なので、1セル=1件のsinglesに入る
    const singleRefs = singles.map(s => CoreLogic.cellRef(s.cells[0]));
    assertTrue(singleRefs.includes('B3') && singleRefs.includes('C3'), 'ラベル付き行の値セルはsinglesに入るべき');
    const b3 = singles.find(s => CoreLogic.cellRef(s.cells[0]) === 'B3');
    assertEqual(b3.rowLabel, '達成状況');
    assertEqual(b3.groupLabel, null);

    // 成果指標グループのB列（B1,B2）は1件（列単位）にまとまる。C列も別の1件。
    assertEqual(groups.length, 2, '2列（B列・C列）ぶんの2件にまとまるべき');
    const bCol = groups.find(g => g.col === 2);
    assertEqual(bCol.groupLabel, '成果指標');
    assertEqual(bCol.cells.map(c => CoreLogic.cellRef(c)), ['B1', 'B2'], 'B列の2行が1件のcellsにまとまるべき');
  });

  test('項目名(dbKey)を設定済みのセルも対象に含まれる（一覧で見直せるように）', () => {
    // ここでは値セルを2つ（B1・C1）用意しdefaultKeyFor（colN自動命名）を経由させている。
    // 値セルが1つだけの行（「1行1見出し1値」）も対象に含まれることは、別テスト
    // 「『1行1見出し1値』の行の値セルも対象に含まれる」で確認する。
    const ws = mockSheet({
      maxRow: 1, maxCol: 3,
      cells: [{ row: 1, col: 1, v: '事業名', fill: 'DDEBF7' }],
    });
    const { grid, sections, maxCol } = extract(ws);
    const target = grid.get('1,2');
    CoreLogic.applyOverrides(grid, { [CoreLogic.cellId(target)]: { kind: 'textarea', dbKey: 'jigyoumei' } }, CoreLogic.cellId);
    const { singles } = CoreLogic.findMappingTargets(grid, sections, maxCol, []);
    assertEqual(singles.length, 2, 'dbKey設定済みでも対象に含まれるはず（未設定のみに絞らない）');
    const named = singles.find(s => CoreLogic.cellRef(s.cells[0]) === 'B1');
    assertEqual(named.cells[0].dbKey, 'jigyoumei');
  });

  test('行内の直前見出し（「区分」等）で名前が決まるセルも一覧に含まれ、autoNameに見出し文字が入る', () => {
    const ws = mockSheet({
      maxRow: 1, maxCol: 3,
      cells: [
        { row: 1, col: 1, v: '事業評価', fill: 'DDEBF7' },
        { row: 1, col: 2, v: '区分' }, // 行内見出し（値としては出力されない）
        // col3が「区分」の値として扱われる → キー名は自動名ではなく「区分」
      ],
    });
    const { grid, sections, maxCol } = extract(ws);
    const { singles, groups } = CoreLogic.findMappingTargets(grid, sections, maxCol, []);
    assertEqual(groups.length, 0);
    assertEqual(singles.length, 1, '直前見出しでキー名が決まるセルも一覧に含まれるようになった（データ構造確認パネルで手直しできるように）');
    assertEqual(CoreLogic.cellRef(singles[0].cells[0]), 'C1');
    assertEqual(singles[0].autoName, '区分', 'autoNameには見出し文字そのものが入る（colNへのフォールバックではない）');
  });

  // D36/E36問題（見出しの直後に値セルが2個連続すると、両方が同じキー名になり片方が
  // 上書きで消えていた）の回帰テスト。precedingLabelを1回使ったらリセットするようにした
  // ことで、1個目は見出し由来のキー、2個目はcolNにフォールバックし、衝突しなくなった。
  test('見出しの直後に値セルが2個連続しても、1個目は見出し由来・2個目はcolNになりキーが衝突しない', () => {
    const ws = mockSheet({
      maxRow: 1, maxCol: 4,
      cells: [
        { row: 1, col: 1, v: '事業評価', fill: 'DDEBF7' },
        { row: 1, col: 2, v: '区分' }, // 見出し（col3のキー名になる）
        // col3, col4は見出しなしの連続する値セル
      ],
    });
    const { grid, sections, maxCol } = extract(ws);
    const out = CoreLogic.buildSectionObject(grid, sections[0].row0, sections[0].row1, maxCol, null, []);
    // 複数値を持つ行は「文脈_項目名」の形にフラット化されるため、トップレベルの
    // キー数で衝突の有無を確認する（本来は「事業評価」に入れ子だったもの）。
    assertEqual(Object.keys(out).length, 2, '衝突せず両方の値が出力されるはず');
    assertEqual(out['事業評価_区分'] !== undefined, true);

    const { singles } = CoreLogic.findMappingTargets(grid, sections, maxCol, []);
    const c1 = singles.find(s => CoreLogic.cellRef(s.cells[0]) === 'C1');
    const d1 = singles.find(s => CoreLogic.cellRef(s.cells[0]) === 'D1');
    assertEqual(c1.autoName, '区分', '見出し直後の1個目は見出し由来のまま');
    assertEqual(d1.autoName, 'col4', '2個目は見出しが使い回されずcolNにフォールバックする');
  });

  // 実機（kagawa_nicchu_yousiki3.xlsx）で発覚した不具合の回帰テスト：グループの中でも
  // 行自体がラベルを持つ場合（法人名／事業所名のように、繰り返しではなく個別の項目が
  // たまたま同じグループ・同じ列に並んでいるだけのケース）は、列単位でまとめてはいけない
  // （まとめると、法人名と事業所名という全く別の値に同じ項目名を強制することになる）。
  test('グループ内でも行自体がラベルを持つ場合は列単位でまとめず、個別のsingleとして扱う', () => {
    // col1=縦結合グループラベル、col2=各行自身のラベル（法人名／事業所名／所在地）、
    // col3・col4=値セル2つ（値セルが1つだけの行はrest.length===1分岐に入りcollectorに
    // 積まれないため、実際にcollectorへ積まれる最小構成として2つ用意する）。
    const ws = mockSheet({
      maxRow: 3, maxCol: 4,
      cells: [
        { row: 1, col: 1, v: '1.施設概要', rowspan: 3 },
        { row: 1, col: 2, v: '法人名' },
        { row: 2, col: 2, v: '事業所名' },
        { row: 3, col: 2, v: '所在地' },
      ],
    });
    const { grid, sections, maxCol } = extract(ws);
    const { singles, groups } = CoreLogic.findMappingTargets(grid, sections, maxCol, []);
    assertEqual(groups.length, 0, '行自体がラベルを持つので列単位のグループにはならないはず');
    // 3行×2値セル(col3,col4)=6件。同じ列番号(col3同士・col4同士)でもcells.length===1のまま
    // （＝法人名のcol3と事業所名のcol3が誤って1件に混ざっていない）ことが本質的な確認点。
    assertEqual(singles.length, 6);
    assertTrue(singles.every(s => s.cells.length === 1), '行が違えば同じ列番号でも別々のsingleであるべき（誤って合算されない）');
    assertTrue(singles.every(s => s.groupLabel === '1.施設概要'), '文脈としてgroupLabelは保持されるべき');
    const rowLabels = new Set(singles.map(s => s.rowLabel));
    assertEqual(Array.from(rowLabels).sort(), ['事業所名', '所在地', '法人名']);
  });

  test('通常のJSON出力（collector省略）には一切影響しない', () => {
    const ws = mockSheet({
      maxRow: 2, maxCol: 2,
      cells: [{ row: 1, col: 1, v: '成果指標', rowspan: 2 }],
    });
    const { data } = extract(ws);
    assertEqual(data['シート']['成果指標'], [{ col2: '' }, { col2: '' }]);
  });

  // 「1行1見出し1値」（例：事業の目的／事業の概要のような、A列に見出し・値セルが1個だけの行）は、
  // JSONキーが行見出し文字そのものに一意に決まるため、当初はSTEP3の項目名マッピング・STEP4の
  // 各種ピッカー（ファイル名項目／必須項目／レビュー欄／一覧表示候補項目）のどれにも対象として
  // 出てこない設計だった。しかしこのパターンの行だけで構成されたシンプルな様式では、
  // 必須項目に何も選べなくなってしまう（実データ：湖西市様式のC7/C8で発覚）ため、
  // 他のセルと同様にsinglesへ含めるよう修正した。
  test('「1行1見出し1値」の行の値セルも対象に含まれる（必須項目・レビュー欄等のピッカーが空にならないように）', () => {
    const ws = mockSheet({
      maxRow: 1, maxCol: 2,
      cells: [{ row: 1, col: 1, v: '事業の目的', fill: 'DDEBF7' }],
    });
    const { grid, sections, maxCol } = extract(ws);
    const { singles, groups } = CoreLogic.findMappingTargets(grid, sections, maxCol, []);
    assertEqual(groups.length, 0);
    assertEqual(singles.length, 1, '値セルが1つだけの行もsinglesに含まれるべき');
    assertEqual(CoreLogic.cellRef(singles[0].cells[0]), 'B1');
    assertEqual(singles[0].rowLabel, '事業の目的', '文脈（rowLabel）には行自身の見出しが入る');
    assertEqual(singles[0].autoName, '事業の目的', '初期の項目名候補（autoName）は行見出し文字そのもの');
  });

  test('「1行1見出し1値」の行はdbKeyで改名すると、JSON出力キーが行見出しではなく改名後の名前になる', () => {
    const ws = mockSheet({
      maxRow: 1, maxCol: 2,
      cells: [{ row: 1, col: 1, v: '事業の目的', fill: 'DDEBF7' }],
    });
    const { grid, sections, maxCol } = extract(ws);
    const target = grid.get('1,2');
    // 改名前：行見出しがそのままキーになる
    const before = CoreLogic.buildSectionObject(grid, sections[0].row0, sections[0].row1, maxCol, () => 'テスト値', []);
    assertEqual(before, { 事業の目的: 'テスト値' });

    CoreLogic.applyOverrides(grid, { [CoreLogic.cellId(target)]: { kind: 'textarea', dbKey: 'purpose' } }, CoreLogic.cellId);
    const after = CoreLogic.buildSectionObject(grid, sections[0].row0, sections[0].row1, maxCol, () => 'テスト値', []);
    assertEqual(after, { purpose: 'テスト値' }, '改名後はdbKeyがそのままJSON出力キーになるべき');
  });

  test('「1行1見出し1値」でも値セルが見出し文字（実データなし）ならsinglesに含めない', () => {
    // 先頭セルに続けてもう1つラベルが並ぶだけの見出し行（実データを持たない）は、
    // 従来通り選択の対象にしない。
    // col1にfillを付けないと、splitSectionsが「単独行・塗りつぶしなし」を
    // セクション見出し行と誤認し、この行自体が本文の外に出てしまう（他のテストと同じ注意点）。
    const ws = mockSheet({
      maxRow: 1, maxCol: 2,
      cells: [
        { row: 1, col: 1, v: '判定', fill: 'DDEBF7' },
        { row: 1, col: 2, v: '判定理由' },
      ],
    });
    const { grid, sections, maxCol } = extract(ws);
    const { singles, groups } = CoreLogic.findMappingTargets(grid, sections, maxCol, []);
    assertEqual(singles.length, 0, '実データを持たない見出し行はsinglesに含めるべきではない');
    assertEqual(groups.length, 0);
  });
});

if (require.main === module) {
  const { summary } = require('./assert_mini');
  process.exit(summary() ? 0 : 1);
}
