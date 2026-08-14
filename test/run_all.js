// 回帰テスト一式をまとめて実行する。各テストファイルは別プロセスで実行する
// （test_grid_render.js等がglobal.document/global.windowを書き換えるため、
// 同一プロセス内で複数ファイルをrequireするとグローバル汚染で干渉しかねない。
// プロセス分離が最も単純で安全）。
//
// 使い方: npm test  または  node test/run_all.js

const { execFileSync } = require('child_process');
const path = require('path');

const FILES = [
  'test_core_logic.js',
  'test_grid_render.js',
  'test_builder_integration.js',
  'test_filler_integration.js',
  'test_aggregator_integration.js',
  'test_viewer_integration.js',
  'test_full_pipeline_smoke.js',
  'test_real_fixtures_smoke.js',
];

let allOk = true;
const summaryLines = [];

for (const f of FILES) {
  const full = path.join(__dirname, f);
  console.log(`\n${'='.repeat(70)}\n${f}\n${'='.repeat(70)}`);
  try {
    const out = execFileSync('node', [full], { encoding: 'utf8' });
    process.stdout.write(out);
    const m = out.match(/合計: (\d+)件成功 \/ (\d+)件失敗/);
    summaryLines.push(`${f}: ${m ? `${m[1]}件成功 / ${m[2]}件失敗` : '(集計行が見つかりません)'}`);
  } catch (e) {
    allOk = false;
    if (e.stdout) process.stdout.write(e.stdout);
    if (e.stderr) process.stderr.write(e.stderr);
    const out = String(e.stdout || '');
    const m = out.match(/合計: (\d+)件成功 \/ (\d+)件失敗/);
    summaryLines.push(`${f}: ${m ? `${m[1]}件成功 / ${m[2]}件失敗` : '異常終了（集計行なし）'} ← FAIL`);
  }
}

console.log(`\n${'='.repeat(70)}\n全体サマリー\n${'='.repeat(70)}`);
summaryLines.forEach(l => console.log('  ' + l));
console.log(allOk ? '\n✅ 全ファイル成功' : '\n❌ 失敗したファイルがあります');

process.exit(allOk ? 0 : 1);
