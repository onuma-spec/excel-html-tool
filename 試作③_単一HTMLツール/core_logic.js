// 事務事業評価データ公表支援ツール ①②のコア処理（ブラウザ・Node両方で動く共通ロジック）
// 前提：ロック状態・データ入力規則は読めない（SheetJS無料版の制約、ユーザー承認済み）。
// 「値の有無」「結合セル」「数式かどうか（cell.f）」だけで構造を判定する簡易版。

(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.CoreLogic = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {

// ---------------------------------------------------------------------------
// 1. 結合セルを解決した論理グリッドを作る
// ---------------------------------------------------------------------------

function colLetterToNum(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function parseRef(ref) {
  const m = ref.match(/^([A-Z]+)(\d+)$/);
  return { col: colLetterToNum(m[1]), row: parseInt(m[2], 10) };
}

function buildGrid(ws) {
  const ref = ws['!ref'];
  const range = XLSXUtilsDecodeRange(ref);
  const merges = ws['!merges'] || [];

  // (row,col) -> アンカー(row0,col0,row1,col1)
  const anchorOf = new Map();
  for (const m of merges) {
    const r0 = m.s.r + 1, c0 = m.s.c + 1, r1 = m.e.r + 1, c1 = m.e.c + 1;
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        anchorOf.set(r + ',' + c, [r0, c0, r1, c1]);
      }
    }
  }

  const grid = new Map();
  const infoCache = new Map();
  for (let r = range.s.r + 1; r <= range.e.r + 1; r++) {
    for (let c = range.s.c + 1; c <= range.e.c + 1; c++) {
      const key = r + ',' + c;
      const anchor = anchorOf.get(key);
      const ar = anchor ? anchor[0] : r;
      const ac = anchor ? anchor[1] : c;
      const ar2 = anchor ? anchor[2] : r;
      const ac2 = anchor ? anchor[3] : c;
      const infoKey = ar + ',' + ac;
      if (!infoCache.has(infoKey)) {
        const addr = colNumToLetter(ac) + ar;
        const cell = ws[addr];
        const isFormula = !!(cell && cell.f);
        const value = isFormula ? '' : (cell && cell.v !== undefined ? cell.v : '');
        const formula = isFormula ? ('=' + cell.f) : '';
        const hasFill = !!(cell && cell.s && cell.s.patternType);
        const fillColor = (hasFill && cell.s.fgColor && cell.s.fgColor.rgb) ? cell.s.fgColor.rgb : null;
        infoCache.set(infoKey, {
          row: ar, col: ac, row2: ar2, col2: ac2,
          value, isFormula, formula, hasFill, fillColor,
          renderType: null, renderOptions: null,
          // 項目名マッピング（B）：JSON出力時のキー名を「col11」等の自動名から
          // 任意の名前に差し替える。未設定はnull（自動名のまま）。
          dbKey: null,
          // 手直しで「見出しの文字を空欄」に設定した結果としてのみtrueになる。
          // 入力不可の空白セル（余白・仕切り）として描画され、構造上も
          // 通常の空欄入力欄と同じ扱い（hasText=false）だが編集はできない。
          blocked: false,
          hasText: (typeof value === 'string' && value.trim() !== '') ||
                    (typeof value === 'number' && !isFormula),
        });
      }
      grid.set(key, infoCache.get(infoKey));
    }
  }
  return { grid, maxRow: range.e.r + 1, maxCol: range.e.c + 1 };
}

/**
 * 手直し（見出し⇔入力欄の切替・型指定・プルダウン・数式編集）をgridに適用する。
 * override = { kind: 'label'|'textarea'|'number'|'currency'|'select'|'formula',
 *              label, options, formula }
 * gridを直接書き換えるので、この後にsplitSections/buildSectionObjectを呼べば
 * 手直し後の構造がそのままJSON出力にも反映される（描画と書き出しで二重管理しない）。
 */
function applyOverrides(grid, overridesById, cellIdFn) {
  for (const info of new Set(grid.values())) {
    const ov = overridesById[cellIdFn(info)];
    if (!ov) continue;
    // 項目名マッピング（B）：セルの種類とは独立した設定なので、種類の分岐に関係なく
    // 常に反映する（未設定はnull＝JSON出力は従来通り自動名「col11」等のまま）。
    info.dbKey = ov.dbKey || null;
    if (ov.kind === 'label') {
      const text = String(ov.label || '').trim();
      if (text) {
        info.hasText = true;
        info.isFormula = false;
        info.value = ov.label;
        info.blocked = false;
      } else {
        // 見出しの文字を空欄にした場合：入力不可の空白セル（余白・仕切り）にする。
        // hasTextはfalseのまま（＝JSONの項目名にもセクション見出しにもならない）。
        info.hasText = false;
        info.isFormula = false;
        info.blocked = true;
      }
    } else if (ov.kind === 'formula') {
      info.isFormula = true;
      info.hasText = false;
      info.blocked = false;
      info.formula = ov.formula || info.formula;
    } else {
      info.hasText = false;
      info.isFormula = false;
      info.blocked = false;
      info.renderType = ov.kind;
      info.renderOptions = ov.options || null;
    }
  }
}

/**
 * 書き出し済みの入力フォームHTMLから取り出したSTRUCTURE（serializeStructureの出力）を、
 * buildGridと同じ形のgrid Mapに復元する。手直し結果（renderType/blocked/formula等）は
 * すでにcellsに焼き込み済みなのでそのまま使い、この上にさらに新しいOVERRIDESを重ねられる。
 */
function buildGridFromCells(cells, maxRow, maxCol) {
  const grid = new Map();
  cells.forEach(c => {
    const info = {
      row: c.row, col: c.col, row2: c.row2, col2: c.col2,
      value: c.value, isFormula: !!c.isFormula, formula: c.formula || '',
      hasFill: !!c.fillColor, fillColor: c.fillColor || null,
      renderType: c.renderType || null, renderOptions: c.renderOptions || null,
      dbKey: c.dbKey || null,
      blocked: !!c.blocked, hasText: !!c.hasText,
    };
    for (let r = c.row; r <= c.row2; r++) {
      for (let cc = c.col; cc <= c.col2; cc++) {
        grid.set(r + ',' + cc, info);
      }
    }
  });
  return { grid, maxRow, maxCol };
}

function colNumToLetter(n) {
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function XLSXUtilsDecodeRange(ref) {
  const [a, b] = ref.split(':');
  const pa = parseRef(a), pb = parseRef(b || a);
  return { s: { r: pa.row - 1, c: pa.col - 1 }, e: { r: pb.row - 1, c: pb.col - 1 } };
}

// ---------------------------------------------------------------------------
// 2. セクション分割：1列目だけに値があり他の列が空の行を「見出し行」とみなす
//    （ロック状態やフォントに頼らず、値の配置パターンだけで判定する簡易版）
// ---------------------------------------------------------------------------

function splitSections(grid, maxRow, maxCol) {
  const markers = [];
  const seen = new Set();
  for (let r = 1; r <= maxRow; r++) {
    const a1 = grid.get(r + ',1');
    if (!a1 || a1.row !== r || !a1.hasText) continue;
    // 複数行にまたがる縦結合（「R6年度」等の外側グループラベル）は見出し行ではない。
    // 見出し行は1行だけの短いラベル（例：Do／Check）という前提を置く。
    if (a1.row2 !== a1.row) continue;
    // 「他の列が空欄」だけでは、空欄テンプレートでは普通のラベル行
    // （例：事業の目的＋空欄の入力欄）と区別がつかない。塗りつぶし色の有無を
    // 追加の判定材料にする（このツールの前提：見出し行だけ無地、ラベル行は
    // 塗りつぶしあり。SheetJSはロック状態は読めないが塗りつぶし色は読める）。
    if (a1.hasFill) continue;
    // 塗りつぶしなし・単独行・値ありという条件だけで見出し行と判定する
    // （「他の列が空欄」という追加条件は、見出し行と同じ行にたまたま入っている
    // 補足メモ（「（単位：千円）」等）1つで簡単に崩れるため、あえて課さない）。
    if (!seen.has(r)) {
      seen.add(r);
      markers.push({ row: r, title: String(a1.value) });
    }
  }
  if (markers.length === 0) {
    return [{ title: 'シート', row0: 1, row1: maxRow }];
  }
  const sections = [];
  markers.forEach((m, i) => {
    const row1 = i + 1 < markers.length ? markers[i + 1].row - 1 : maxRow;
    sections.push({ title: m.title, row0: m.row + 1, row1 });
  });
  return sections;
}

// ---------------------------------------------------------------------------
// 3. 行ユニット化＋構造の入れ子JSON化（ラベル→値、グループ→子リストを再帰的に組み立てる）
// ---------------------------------------------------------------------------

function rowSpanOf(grid, r, maxCol) {
  // 1列目（外側グループラベル列になりやすい）を除外し、2列目以降の結合状況だけで
  // その行の実質的な高さを求める（PDF/Excel版の試作で得た教訓を踏襲）。
  let row2 = r;
  for (let c = 2; c <= maxCol; c++) {
    const info = grid.get(r + ',' + c);
    if (info && info.row === r) row2 = Math.max(row2, info.row2);
  }
  return row2;
}

function firstCellOfRow(grid, r0, maxCol, fromCol) {
  for (let c = fromCol; c <= maxCol; c++) {
    const info = grid.get(r0 + ',' + c);
    if (info && info.row === r0) return info;
  }
  return null;
}

function cellsOfRow(grid, r0, maxCol, fromCol) {
  const out = [];
  const seen = new Set();
  for (let c = fromCol; c <= maxCol; c++) {
    const info = grid.get(r0 + ',' + c);
    if (!info || info.row !== r0) continue;
    const key = info.row + ',' + info.col;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(info);
  }
  return out;
}

function buildUnits(grid, row0, row1, maxCol) {
  const units = [];
  let r = row0;
  while (r <= row1) {
    const r2 = rowSpanOf(grid, r, maxCol);
    units.push([r, r2]);
    r = r2 + 1;
  }
  return units;
}

/**
 * セクション本文（row0〜row1）を、入れ子のJSONオブジェクトに変換する。
 * - 手動グループ（人が🗂ボタンで明示的に宣言した行範囲）があれば、その範囲の
 *   子ユニット群をまとめる（子が全て無名＝配列、子に個別ラベルあり＝辞書）。
 *   グループ化はこの手動宣言だけが起点で、Excelのタテ結合の有無からは自動検出しない
 *   （タテ結合は見出し行とデータ行を区別できないため）。
 * - それ以外の単発ユニットは「1列目のラベル → 2列目以降の値（複数あれば辞書）」として扱う。
 * - 数式セル（合計行等）は除外する。
 */
function cellId(info) {
  return 'cell_R' + info.row + '_C' + info.col;
}

// getValue(info) は「このセルの値をどこから読むか」を差し替え可能にする。
// 省略時はテンプレート自身の値（プレビュー用）。フォーム入力後の書き出し時は
// 対応するDOM要素（input/textarea、id=cellId(info)）から読む関数を渡す。
function defaultGetValue(info) {
  return info.hasText ? info.value : '';
}

// 手動グループ化（A）：ある行(u0)がユーザーが明示的に宣言したグループの開始行かどうかを返す。
// {row0,row1,name} の配列から、row0が一致するものを探す（1範囲=1グループの前提）。
function findManualGroup(manualGroups, u0) {
  return (manualGroups || []).find(g => g.row0 === u0) || null;
}

// 自動検出（縦結合）：ユニット(u0,u1)が「1列目に複数ユニットにまたがる縦結合ラベルを
// 持つグループの先頭」かどうかを判定する。buildSectionObject（書き出し）・
// renderGrid（グリッド描画時のハイライト）の両方から同じ条件で呼べるよう共通化する。
function isAutoGroupStart(groupInfo, u0, u1) {
  return !!(groupInfo && groupInfo.row === u0 && groupInfo.row2 > u1 && groupInfo.hasText);
}

// 行rowが、自動検出グループ（縦結合）の範囲内に含まれているかどうかを判定して返す
// （UIが「この選択範囲は自動グループ化を解除できます」を出すかどうかの判定に使う）。
// 先頭行（縦結合のアンカー行）である必要はない。1列目のセルはExcelの結合により
// アンカー行のCellInfoをそのまま指しているため、rowがどこであってもgrid.get(row+',1')は
// アンカー行の情報を返す（buildGridのanchorOf解決による）。これは自動グループが
// 現在disabledかどうかにも左右されない、Excel自体の構造だけを見た判定である。
// 該当しなければnull。該当すればラベル名・自動検出としての範囲（row0〜row1）を返す。
function getAutoGroupRange(grid, maxCol, row) {
  const groupInfo = grid.get(row + ',1');
  if (!groupInfo || !groupInfo.hasText) return null;
  const anchorRow = groupInfo.row;
  const u1 = rowSpanOf(grid, anchorRow, maxCol);
  if (!isAutoGroupStart(groupInfo, anchorRow, u1)) return null;
  return { row0: anchorRow, row1: groupInfo.row2, label: String(groupInfo.value).replace(/\s+/g, '') };
}

// グリッド全体のうち、自動検出（縦結合）によってグループの一部とみなされる行番号の集合を返す。
// 手動グループが同じ範囲を明示的に宣言している場合はそちらが優先されるため、その分は除く
// （buildSectionObjectの優先順位・打ち切りロジックとズレると、実際には手動グループとして
// 出力される行が自動グループの色で表示される、といった見た目と実体の食い違いが起きる）。
function computeAutoGroupRows(grid, sections, maxCol, manualGroups) {
  const rows = new Set();
  for (const sec of sections) {
    const units = buildUnits(grid, sec.row0, sec.row1, maxCol);
    let i = 0;
    while (i < units.length) {
      const [u0, u1] = units[i];
      const manualGroup = findManualGroup(manualGroups, u0);
      if (manualGroup && manualGroup.disabled) {
        // 自動グループ化を解除した範囲：グループではないので色も付けない。このユニットだけ
        // スキップして次へ（buildSectionObjectの解除ブランチと同じ単位で進める）。
        i++;
        continue;
      }
      if (manualGroup) {
        let j = i;
        while (j < units.length && units[j][0] <= manualGroup.row1) j++;
        i = j;
        continue;
      }
      const groupInfo = grid.get(u0 + ',1');
      if (isAutoGroupStart(groupInfo, u0, u1)) {
        let groupEndRow = groupInfo.row2;
        const overlappingManual = (manualGroups || [])
          .filter(g => g.row0 > u0 && g.row0 <= groupEndRow)
          .sort((a, b) => a.row0 - b.row0)[0];
        if (overlappingManual) groupEndRow = overlappingManual.row0 - 1;
        let j = i;
        while (j < units.length && units[j][0] <= groupEndRow) {
          for (let r = units[j][0]; r <= units[j][1]; r++) rows.add(r);
          j++;
        }
        i = j;
        continue;
      }
      i++;
    }
  }
  return rows;
}

// グループの子ユニット群（buildRowEntryの結果配列）を、JSON出力用の辞書または配列に
// まとめる。まず実データを持たない子（hasRealData=false。見出し行の残骸）を除外する。
// 残った子が全員ラベル付きなら辞書（groupValue）にする。
// 無名の繰り返し行と「その他」のような名前付きの単発行が混在する場合、無名の子だけを
// 配列（groupValue）にし、名前付きの子はグループの外側で通常の単発行として扱えるよう
// extraEntriesで返す（配列に混ぜると要素の形がバラバラになり、ラベルの情報も失われる
// ため。以前はここで名前付きの子を捨てておりデータが消えていた）。
//
// opts.requireMultiple（自動検出グループ専用。手動グループでは使わない）：
// 「全員ラベル付きなら辞書」の判定に、子が2件以上あることを追加で要求する。
// 自動検出グループは、手動グループに範囲を奪われて残りが1件だけになることがあり
// （既知の不具合として報告・修正済み）、その1件がたまたまラベル付きだと
// 「見出し由来の残骸1件だけの辞書」が生成されてしまっていた。手動グループなしで
// 縦結合そのものが検出される場合は必ず2ユニット以上あるため（isAutoGroupStartが
// row2>u1を要求する）、requireMultipleをtrueにしても正常なケースには影響しない。
// 一方、手動グループ（人が明示的に1行だけを選んで宣言する運用が正当にある。
// 例：備考欄1行をメモ欄として辞書化する）ではrequireMultipleを立てないことで、
// 1件だけの辞書化を引き続き許容する。
function groupChildrenToResult(children, opts) {
  const minCount = (opts && opts.requireMultiple) ? 1 : 0;
  const real = children.filter(c => c.hasRealData);
  if (real.length > minCount && real.every(c => c.label)) {
    const dict = {};
    real.forEach(c => { dict[c.label] = c.value; });
    return { groupValue: dict, extraEntries: [] };
  }
  const unlabeled = real.filter(c => !c.label);
  const labeled = real.filter(c => c.label);
  return { groupValue: unlabeled.map(c => c.value), extraEntries: labeled };
}

// collector（省略可）：値セルが実際に使う項目名を、そのセル（CellInfo）と一緒に
// 記録したい場合に配列を渡す（findMappingTargets専用）。渡すと、defaultKeyFor()による
// 自動命名（colN）が使われたセルの{cell, rowLabel, groupLabel}がここに積まれる。
// 通常のJSON書き出し（collector省略）には一切影響しない。
// 単発行のentryをresultへ反映する。hasRealData=falseは出力しない
// （ラベル同士が隣接する見出し行の残骸を、明示的な手直しなしでは実データとして
// 扱わない。buildRowEntryのコメント参照）。
// 複数値を持つ行（value がプレーンオブジェクト）は入れ子にせず「文脈_項目名」の形に
// フラット化する。文脈だけを親キーにした入れ子だと、列位置由来の自動名（colN）が
// 他の行の自動名と衝突しやすい（同じ列番号は何十行にもわたって繰り返し出現するため）。
// フラット化すれば文脈自体がキーの一部になり、天然の名前空間として機能する。
function assignEntry(result, entry) {
  if (!entry.label || !entry.hasRealData) return;
  if (entry.value && typeof entry.value === 'object' && !Array.isArray(entry.value)) {
    Object.keys(entry.value).forEach((subKey) => {
      result[entry.label + '_' + subKey] = entry.value[subKey];
    });
  } else {
    result[entry.label] = entry.value;
  }
}

function buildSectionObject(grid, row0, row1, maxCol, getValue, manualGroups, collector) {
  getValue = getValue || defaultGetValue;
  const units = buildUnits(grid, row0, row1, maxCol);
  const result = {};
  let i = 0;
  while (i < units.length) {
    const [u0, u1] = units[i];

    // 手動グループ（A）：Excelの縦結合の有無に頼らず、人が明示的に「この範囲は
    // 親ラベル＋子レコードの配列」と宣言したもの。自動検出（縦結合ラベル）より優先する。
    // 子行は1列目もそのまま値として扱う（縦結合ラベルのように1列目を消費しないため fromCol=1）。
    const manualGroup = findManualGroup(manualGroups, u0);
    // 自動グループ化の解除（disabled:true）：この行から始まる自動検出グループを
    // 「グループ化しない」と明示的に宣言したもの。グループとしてはまとめず、
    // このユニット1つだけ通常の単発行として処理し、次のユニットへ進む
    // （後続のユニットはそもそもisAutoGroupStartの対象外なので、ここで1回スキップ
    // すれば十分。自動検出の判定自体には手を加えない）。
    if (manualGroup && manualGroup.disabled) {
      const entry = buildRowEntry(grid, u0, u1, maxCol, 1, getValue, collector, null);
      assignEntry(result, entry);
      i++;
      continue;
    }
    if (manualGroup) {
      const groupEndRow = manualGroup.row1;
      const childUnits = [];
      let j = i;
      while (j < units.length && units[j][0] <= groupEndRow) { childUnits.push(units[j]); j++; }

      const children = childUnits.map(([cu0, cu1]) => buildRowEntry(grid, cu0, cu1, maxCol, 1, getValue, collector, manualGroup.name));
      const { groupValue, extraEntries } = groupChildrenToResult(children);
      result[manualGroup.name] = groupValue;
      // 配列に混ぜられなかった名前付きの子（「その他」等）は、グループの外側で
      // 通常の単発行として出力する（データを失わないため）。
      extraEntries.forEach(e => assignEntry(result, e));
      i = j;
      continue;
    }

    // 自動検出（縦結合）：1列目に複数ユニットにまたがる縦結合ラベル（例：「R6年度」
    // 「成果指標」）があれば、そのラベルをキーに子ユニット群を自動的にまとめる。
    // 手動グループより後に判定するので、同じ範囲を手動グループで明示的に宣言すれば
    // そちらが優先される。
    const groupInfo = grid.get(u0 + ',1');
    if (isAutoGroupStart(groupInfo, u0, u1)) {
      let groupEndRow = groupInfo.row2;
      // 自動検出の範囲内から手動グループが開始している場合、そのままでは自動グループが
      // 手動グループの行まで丸ごと飲み込んでしまい、手動グループ側のu0が一生ループに
      // 現れない（このifブロックがjを一気に進めてしまうため）。手動グループの開始行の
      // 手前で自動グループを打ち切り、後続のループで手動グループ側のu0を正しく迎える。
      const overlappingManual = (manualGroups || [])
        .filter(g => g.row0 > u0 && g.row0 <= groupEndRow)
        .sort((a, b) => a.row0 - b.row0)[0];
      if (overlappingManual) groupEndRow = overlappingManual.row0 - 1;

      const childUnits = [];
      let j = i;
      while (j < units.length && units[j][0] <= groupEndRow) { childUnits.push(units[j]); j++; }

      const key = String(groupInfo.value).replace(/\s+/g, '');
      const children = childUnits.map(([cu0, cu1]) => buildRowEntry(grid, cu0, cu1, maxCol, 2, getValue, collector, key));
      // requireMultiple:true — 自動検出グループが手動グループに範囲を奪われて1件だけ
      // 残った場合に、その1件を辞書化せずextraEntries（グループ外の独立キー）へ回す
      // ための指定（groupChildrenToResultのコメント参照。既知バグの修正）。
      const { groupValue, extraEntries } = groupChildrenToResult(children, { requireMultiple: true });
      result[key] = groupValue;
      extraEntries.forEach(e => assignEntry(result, e));
      i = j;
      continue;
    }

    const entry = buildRowEntry(grid, u0, u1, maxCol, 1, getValue, collector, null);
    assignEntry(result, entry);
    // ラベルなしの単発行（グループにも属さない）は、見出しどうしの間の余白行など
    // 意味を持たない行であることがほとんどのため、トップレベルには出力しない
    // （グループ内の無名行は上の自動検出／手動グループの分岐で別途、配列として正しく保持される）。
    i++;
  }
  return result;
}

// 値セルの既定キー名を求める：①明示的なdbKeyが最優先、②なければ同じ行内で直前に
// 現れた見出しセルの文字（「事業類型」「区分」のように、行の途中に埋め込まれた
// 見出しをそのまま項目名として使う）、③どちらもなければ従来通りcolN。
// core_logic.js（書き出し）・grid_render.js（読込）の両方で同じ規則を使う必要があるため
// 共有関数として切り出す（対称実装が崩れるとデータ消失・混線の原因になるという既存の教訓）。
function defaultKeyFor(c, precedingLabel) {
  return c.dbKey || precedingLabel || ('col' + c.col);
}

/**
 * 1つの行ユニットを {label, value, hasRealData} に変換する。
 * fromCol=1のときは1列目自体をラベル候補とする。fromCol=2のときは
 * （呼び出し元がすでにグループラベルとして1列目を消費済みなので）2列目から見る。
 * hasRealData：この行に実際の入力欄（空欄セル）が1つでも含まれるかどうか。
 * 「判定／判定理由」のような、値側もラベルしか持たない見出し行（グループの先頭に
 * 混ざりやすい）を、繰り返しグループの処理側でレコードから除外するために使う。
 */
// collector/context：findMappingTargets専用（buildSectionObjectのコメント参照）。
// context には、この行が属するグループのキー名（無ければnull）を渡す。
// dbKey設定済みかどうかに関わらず収集する（findMappingTargets側で「未設定のみ」
// 「設定済みも含めて全部」を判断する）。
function buildRowEntry(grid, u0, u1, maxCol, fromCol, getValue, collector, context) {
  getValue = getValue || defaultGetValue;
  const cells = cellsOfRow(grid, u0, maxCol, fromCol).filter(c => !c.isFormula);
  if (cells.length === 0) return { label: null, value: null, hasRealData: false };

  const first = cells[0];
  if (first.hasText) {
    // 先頭セルがラベル。残りのセルを値として集める。
    const rowLabel = String(first.value).replace(/\s+/g, '');
    const rest = cells.slice(1);
    if (rest.length === 0) return { label: null, value: null, hasRealData: false };
    if (rest.length === 1) {
      const valueCell = rest[0];
      if (valueCell.hasText && !valueCell.dbKey && !valueCell.renderType) {
        // rest[0]自体がラベル（＝先頭セルに続けてもう1つラベルが並ぶだけの見出し行）である
        // 可能性と、既に記入済みの本物のデータである可能性を、テキストの有無だけでは
        // 区別できない（実データを持つセルもhasText=trueになるため）。明示的な手直し
        // （項目名の設定・種類の変更のいずれか）が無い限りは実データなしとして扱う。
        // 名前を選ぶ余地もないためcollectorには積まない。
        return { label: rowLabel, value: getValue(valueCell), hasRealData: false };
      }
      // 「1行1見出し1値」：JSONキーは行見出し文字がそのまま既定名になり曖昧さがないため、
      // 当初はSTEP3の項目名マッピング・STEP4の各種ピッカーどちらにも出さない設計だった。
      // しかしこのパターンの行だけで構成された様式では選択肢が0件になってしまうため、
      // 他のセルと同様にcollectorへ積んで選択・改名の対象にする
      // （dbKeyで改名されればそちらを出力キーとして優先する）。
      const autoName = rowLabel;
      if (collector && !valueCell.blocked) collector.push({ cell: valueCell, rowLabel, groupLabel: context || null, autoName });
      return { label: valueCell.dbKey || rowLabel, value: getValue(valueCell), hasRealData: true };
    }
    const dict = {};
    let precedingLabel = null;
    let hasRealData = false;
    rest.forEach((c) => {
      // 行の途中に埋め込まれた見出しセル（例：「事業類型」「区分」）は、それ自体は
      // 値として出力せず、直後の1個の値セルの既定の項目名として使う。次の値セルからは
      // 新しい見出しに出会うまでcolNにフォールバックする（見出しが使い回されて複数の
      // 値セルが同じキーになり、片方が上書きで消える事故を防ぐ）。
      if (c.hasText) { precedingLabel = String(c.value).replace(/\s+/g, ''); return; }
      hasRealData = true;
      const key = defaultKeyFor(c, precedingLabel);
      dict[key] = getValue(c);
      // dbKey未設定時にこのセルが今まさに使っている既定名（見出し由来かcolNか）。
      // findMappingTargets側で「見出し由来だから目立たせない／colNだから目立たせる」の
      // 判定に使う（builder_app.jsのbuildMappingBlock参照）。
      const autoName = precedingLabel || ('col' + c.col);
      if (collector && !c.blocked) collector.push({ cell: c, rowLabel, groupLabel: context || null, autoName });
      precedingLabel = null;
    });
    return { label: rowLabel, value: dict, hasRealData };
  }
  // 先頭セルが空欄（＝ラベルなしの行。繰り返し項目行など）→ 行全体を1つの無名レコードとして返す
  const dict = {};
  let precedingLabel = null;
  cells.forEach(c => {
    if (c.hasText) { precedingLabel = String(c.value).replace(/\s+/g, ''); return; }
    const key = defaultKeyFor(c, precedingLabel);
    dict[key] = getValue(c);
    const autoName = precedingLabel || ('col' + c.col);
    if (collector && !c.blocked) collector.push({ cell: c, rowLabel: null, groupLabel: context || null, autoName });
    precedingLabel = null;
  });
  return { label: null, value: dict, hasRealData: true };
}

// ---------------------------------------------------------------------------
// メイン：workbook（SheetJSのws）→ 構造化オブジェクト
// ---------------------------------------------------------------------------

function extractStructure(ws, manualGroups) {
  const { grid, maxRow, maxCol } = buildGrid(ws);
  const sections = splitSections(grid, maxRow, maxCol);
  const data = {};
  const meta = { sections: sections.map(s => s.title) };
  for (const sec of sections) {
    data[sec.title] = buildSectionObject(grid, sec.row0, sec.row1, maxCol, null, manualGroups);
  }
  return { grid, maxRow, maxCol, sections, data, meta };
}

// ---------------------------------------------------------------------------
// 書き出し前チェック：現在のグループ化・項目名の設定で、各入力セルの値が
// 実際にJSON出力へ反映されるかどうかを検証する。
// 「推測でラベルなし行を探す」のではなく、各入力セルに固有の目印（セル番地。
// 例：K42）を仮に入れて実際の出力ロジック（buildSectionObject）を走らせ、
// その目印が出力に現れなかったセル＝「入力しても消える」セルとして検出する。
// 既存の静的なラベル文字（hasText）はそのまま使うので、プレビューにも使える。
// ---------------------------------------------------------------------------
function cellRef(info) {
  return colNumToLetter(info.col) + info.row;
}

function findUnreachableCells(grid, sections, maxCol, manualGroups) {
  const markerOf = (info) => (info.hasText ? info.value : cellRef(info));

  const output = {};
  for (const sec of sections) {
    output[sec.title] = buildSectionObject(grid, sec.row0, sec.row1, maxCol, markerOf, manualGroups);
  }

  const reached = new Set();
  const walk = (v) => {
    if (v === null || v === undefined) return;
    if (typeof v === 'string') { reached.add(v); return; }
    if (Array.isArray(v)) { v.forEach(walk); return; }
    if (typeof v === 'object') { Object.keys(v).forEach(k => walk(v[k])); return; }
  };
  walk(output);

  const seen = new Set();
  const unreachable = [];
  for (const info of grid.values()) {
    const key = info.row + ',' + info.col;
    if (seen.has(key)) continue;
    seen.add(key);
    // 対象は「入力欄」のみ（見出し・数式・入力不可セルは元々入力対象ではない）
    if (info.isFormula || info.hasText || info.blocked) continue;
    if (!reached.has(cellRef(info))) {
      // どのセクション（Plan/Do/Check/Action等）にも属さない行（見出しより前の行等）は、
      // グループ化してもbuildSectionObjectの処理対象外のままなので直らない。
      // 「グループ化すれば直る」と誤った案内をしないよう区別しておく。
      const inSection = sections.some(s => info.row >= s.row0 && info.row <= s.row1);
      unreachable.push(Object.assign({}, info, { outOfSection: !inSection }));
    }
  }
  unreachable.sort((a, b) => (a.row - b.row) || (a.col - b.col));
  return { output, unreachable };
}

// ---------------------------------------------------------------------------
// 項目名（dbKey）マッピングの対象になる値セルを一覧化する。実際の書き出しロジック
// （buildSectionObject）にcollectorを差し込んで走らせるだけなので、書き出し結果と
// 定義がズレることがない。行内の直前見出し（「区分」等）で名前が決まるセルも、
// colN任せのセルと同じく一覧に含める（既に意味のある名前が付いているかどうかに関わらず、
// データ構造確認パネルで手直しできるようにするため）。各エントリのautoName（見出し由来か
// colNか）を使って、呼び出し側（builder_app.js）が「既に見出し由来の名前がある」ものと
// 「colNのまま＝目立たせるべき」ものを描画で区別する。
// findUnreachableCells（そもそもJSON出力に含まれないセル）とは別の観点：
// こちらは「出力はされるが、名前をどうするか」の対象。出力に含まれないセルまで
// 混ぜると紛らわしいので、unreachableなセルは呼び出し側で除外すること
// （builder_app.jsのrenderCheckPanel参照）。
//
// グループ（自動検出・手動どちらも）に属し、かつ行自体にラベルが無い
// （＝無名の繰り返し行＝配列の1要素になる行）セルだけを、同じグループ・同じ列で
// 1件にまとめて返す（{ kind:'group', cells:[...], groupLabel, col }）。
// この種の繰り返し行は「同じ意味の項目が複数行に並ぶ」ことがほとんどで、
// 行ごとに別々の項目名を付けると配列の要素ごとにキー名がバラバラになってしまう
// （実際にこのバグが起きた実例で発覚）。列単位でまとめて1つの名前を適用する方が
// 実態に合っている。
// 一方、グループの中にあっても行自体がラベルを持つ場合（例：「1.施設概要」という
// グループの中の「法人名」「事業所名」の各行）は、同じ列でも行ごとに全く別の意味の
// 値であり、列単位でまとめると無関係な行に同じ名前を強制することになってしまう。
// この場合は1件＝1セルのsingleとして扱う（groupLabelは文脈表示のためだけに残す）。
// グループに全く属さないセルも同様に1件＝1セルのsingleとして返す。
// ---------------------------------------------------------------------------
function findMappingTargets(grid, sections, maxCol, manualGroups) {
  const collector = [];
  for (const sec of sections) {
    buildSectionObject(grid, sec.row0, sec.row1, maxCol, null, manualGroups, collector);
  }
  const seen = new Set();
  const flat = [];
  collector.forEach((entry) => {
    const key = entry.cell.row + ',' + entry.cell.col;
    if (seen.has(key)) return;
    seen.add(key);
    flat.push(entry);
  });

  const singles = [];
  // 'groupLabel|col' -> { groupLabel, col, cells: [], autoName }
  // autoNameは列の代表値としてバケット作成時（＝先頭行）の値を採用する。同じ列なら
  // 通常は行によらず同じ見出しパターンになるため（実データで確認済み）、これで十分。
  const groupBuckets = new Map();
  flat.forEach((entry) => {
    if (entry.groupLabel && !entry.rowLabel) {
      const bucketKey = entry.groupLabel + '|' + entry.cell.col;
      if (!groupBuckets.has(bucketKey)) groupBuckets.set(bucketKey, { groupLabel: entry.groupLabel, col: entry.cell.col, cells: [], autoName: entry.autoName });
      groupBuckets.get(bucketKey).cells.push(entry.cell);
    } else {
      singles.push({ kind: 'single', cells: [entry.cell], rowLabel: entry.rowLabel, groupLabel: entry.groupLabel, autoName: entry.autoName });
    }
  });
  const groups = Array.from(groupBuckets.values()).map(b => ({
    kind: 'group', cells: b.cells, rowLabel: null, groupLabel: b.groupLabel, col: b.col, autoName: b.autoName,
  }));

  const byFirstCell = (a, b) => (a.cells[0].row - b.cells[0].row) || (a.cells[0].col - b.cells[0].col);
  singles.sort(byFirstCell);
  groups.sort(byFirstCell);
  return { singles, groups };
}

return {
  buildGrid, buildGridFromCells, splitSections, buildSectionObject, extractStructure,
  colNumToLetter, cellId, cellRef, defaultGetValue, applyOverrides, findManualGroup,
  findUnreachableCells, findMappingTargets, defaultKeyFor, isAutoGroupStart, computeAutoGroupRows,
  getAutoGroupRange, buildRowEntry, groupChildrenToResult,
};
});
