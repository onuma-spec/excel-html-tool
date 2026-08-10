// core_logic.js の buildGrid が期待するSheetJS風ワークシートオブジェクトを
// 実際のExcelファイルを介さず直接組み立てるテスト用ヘルパー。
// buildGridが読むのは ws[addr]（v/f/s.patternType/s.fgColor.rgb）と
// ws['!ref']・ws['!merges'] だけなので、これらだけを最小構成で作る。

function colNumToLetter(n) {
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * spec.maxRow, spec.maxCol
 * spec.cells: [{row, col, rowspan=1, colspan=1, v, f, fill}]
 *   - v: セルの値（文字列 or 数値）。省略＝空欄の入力セル。
 *   - f: 数式文字列（'SUM(...)' 等、buildGridが'='を付けて格納する）。
 *   - fill: 塗りつぶし色（例 'FFFF00'）。指定すると見出し（ラベル）扱いの判定材料になる。
 *   - rowspan/colspan: 2以上を指定すると!mergesに登録される。
 */
function mockSheet(spec) {
  const ws = {};
  const merges = [];
  (spec.cells || []).forEach((c) => {
    const rowspan = c.rowspan || 1;
    const colspan = c.colspan || 1;
    const addr = colNumToLetter(c.col) + c.row;
    const cellObj = {};
    if (c.f !== undefined) cellObj.f = c.f;
    else if (c.v !== undefined) cellObj.v = c.v;
    if (c.fill) cellObj.s = { patternType: 'solid', fgColor: { rgb: c.fill } };
    ws[addr] = cellObj;
    if (rowspan > 1 || colspan > 1) {
      merges.push({
        s: { r: c.row - 1, c: c.col - 1 },
        e: { r: c.row - 1 + rowspan - 1, c: c.col - 1 + colspan - 1 },
      });
    }
  });
  ws['!ref'] = 'A1:' + colNumToLetter(spec.maxCol) + spec.maxRow;
  ws['!merges'] = merges;
  return ws;
}

module.exports = { mockSheet, colNumToLetter };
