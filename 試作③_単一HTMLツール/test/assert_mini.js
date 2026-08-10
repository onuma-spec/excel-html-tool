// 依存ゼロの最小テストランナー（describe/it/assert）。
// 各テストファイルは `runSuite(name, fn)` で登録し、fn内で `test(label, fn)` を呼ぶ。

const results = { pass: 0, fail: 0, failures: [] };
let currentSuite = '';

function runSuite(name, fn) {
  currentSuite = name;
  console.log(`\n## ${name}`);
  fn();
}

function test(label, fn) {
  try {
    fn();
    results.pass++;
    console.log(`  ok - ${label}`);
  } catch (e) {
    results.fail++;
    results.failures.push(`[${currentSuite}] ${label}: ${e.message}`);
    console.log(`  FAIL - ${label}`);
    console.log(`         ${e.message}`);
  }
}

function assertEqual(actual, expected, msg) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a !== e) {
    throw new Error(`${msg || 'assertEqual'}\n         actual:   ${a}\n         expected: ${e}`);
  }
}

function assertTrue(cond, msg) {
  if (!cond) throw new Error(msg || 'assertTrue failed');
}

function assertFalse(cond, msg) {
  if (cond) throw new Error(msg || 'assertFalse failed');
}

// 非同期版（ファイル読込等、FileReaderの完了を待つ必要があるテスト用）。
// 既存の同期版test/runSuiteとは独立させ、既存テストファイルの実行順序保証を壊さない。
async function runSuiteAsync(name, fn) {
  currentSuite = name;
  console.log(`\n## ${name}`);
  await fn();
}

async function testAsync(label, fn) {
  try {
    await fn();
    results.pass++;
    console.log(`  ok - ${label}`);
  } catch (e) {
    results.fail++;
    results.failures.push(`[${currentSuite}] ${label}: ${e.message}`);
    console.log(`  FAIL - ${label}`);
    console.log(`         ${e.stack || e.message}`);
  }
}

function summary() {
  console.log(`\n合計: ${results.pass}件成功 / ${results.fail}件失敗`);
  if (results.failures.length) {
    console.log('\n--- 失敗一覧 ---');
    results.failures.forEach(f => console.log('  - ' + f));
  }
  return results.fail === 0;
}

module.exports = { runSuite, test, runSuiteAsync, testAsync, assertEqual, assertTrue, assertFalse, summary, results };
