'use strict';
/* 構築スクリプトを、模擬キントーンに対して通しで動かす。実際のキントーンには一切つながない。 */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFile } = require('node:child_process');
const JSZip = require('jszip');
const { createMock } = require('./mock-kintone');
const { build } = require('../scripts/make-test-templates');

const ROOT = path.join(__dirname, '..');
function run(script, args, env) {
  return new Promise((resolve) => {
    execFile('node', [path.join(ROOT, 'scripts', script)].concat(args || []), { env: Object.assign({}, process.env, env), cwd: ROOT }, (err, stdout, stderr) => {
      resolve({ code: err ? err.code : 0, out: stdout + stderr });
    });
  });
}

async function withMock(fn) {
  const mock = createMock();
  const url = await mock.listen();
  // 本物のバックアップ（backup/）に影響されないよう、テストごとに空の保存先を使う
  const env = { KINTONE_BACKUP_DIR: fs.mkdtempSync(path.join(os.tmpdir(), 'kd-backup-')), KINTONE_BASE_URL: url, KINTONE_DOMAIN: 'mock.example', KINTONE_USER: 'u', KINTONE_PASSWORD: 'p', KINTONE_APP_ID: '575' };
  try { await fn(mock, env); } finally { await mock.close(); }
}

test('create-apps：--yes なしは何も作らない／--yes で書類マスタを作成・初期データ19件・確認OK／2回目は止まる', async () => {
  await withMock(async (mock, env) => {
    let r = await run('create-apps.js', ['--space=7'], env);
    assert.strictEqual(r.code, 0, r.out);
    assert.match(r.out, /確認だけです/);
    assert.strictEqual(Object.keys(mock.state.apps).length, 0);

    r = await run('create-apps.js', ['--space=7', '--yes'], env);
    assert.strictEqual(r.code, 0, r.out);
    assert.match(r.out, /OK 項目・件数・スペースとも想定どおり/);
    const apps = Object.values(mock.state.apps);
    assert.deepStrictEqual(apps.map((a) => a.name), ['書類マスタ'], '作るのは書類マスタだけ（案件書類アプリは作らない）');
    assert.ok(apps.every((a) => a.spaceId === '7'), 'スペース内に作る');
    const master = apps.find((a) => a.name === '書類マスタ');
    assert.strictEqual(mock.state.records[master.appId].length, 19);
    assert.deepStrictEqual(Object.keys(master.fields), ['工程', '表示順', '書類名', '説明', '形式', '雛形', '見られる立場', '対象', '使用状況']);
    const first = mock.state.records[master.appId][0];
    assert.deepStrictEqual(first['見られる立場'].value, ['営業', 'TS計画推進', '工事部']);
    assert.ok(mock.state.records[master.appId].every((r) => r['対象'].value === '共通'), '初期の書類は、すべて対象「共通」');
    assert.deepStrictEqual(Object.keys(master.fields['対象'].options), ['共通', '橋梁', '一般道路']);

    const before = mock.state.log.length;
    r = await run('create-apps.js', ['--space=7', '--yes'], env);
    assert.strictEqual(r.code, 2, r.out);
    assert.match(r.out, /すでにあります/);
    assert.ok(!mock.state.log.slice(before).some((l) => l.startsWith('POST /k/v1/preview/app')), '2回目は何も作らない');
  });
});

test('create-apps：スペース番号なしはエラー', async () => {
  await withMock(async (mock, env) => {
    const r = await run('create-apps.js', [], Object.assign({}, env, { KINTONE_SPACE_ID: '' }));
    assert.notStrictEqual(r.code, 0);
  });
});

test('認証情報がない場合は、止まって環境変数名を知らせる', async () => {
  const r = await run('create-apps.js', ['--space=7'], { KINTONE_USER: '', KINTONE_PASSWORD: '', KINTONE_DOMAIN: 'x', KINTONE_BASE_URL: '' });
  assert.notStrictEqual(r.code, 0);
  assert.match(r.out, /KINTONE_USER|\.env/);
});

test('backup → apply-customize：バックアップなしは実行しない／既存を残して追加／同名は入れ替え', async () => {
  await withMock(async (mock, env) => {
    mock.addApp('575', '案件くん');
    const oldKey = mock.addFile(Buffer.from('/* 既存のJS */'), 'existing.js');
    mock.state.apps['575'].customize.desktop.js = [{ type: 'URL', url: 'https://example.com/a.js' }, { type: 'FILE', file: { fileKey: oldKey, name: 'existing.js' } }];
    mock.state.apps['575'].customize.desktop.css = [{ type: 'URL', url: 'https://example.com/a.css' }];
    // すでにある backup/ を壊さないよう、テスト用に別の作業場所で動かす
    const backupDir = env.KINTONE_BACKUP_DIR;
    const had = fs.existsSync(backupDir) ? fs.readdirSync(backupDir) : [];

    let r = await run('apply-customize.js', [], env);
    assert.strictEqual(r.code, 0, r.out);
    assert.match(r.out, /既存の設定：3件（すべて残ります）/);
    assert.match(r.out, /今回追加／更新/);

    r = await run('apply-customize.js', ['--yes'], env);
    assert.strictEqual(r.code, 2, r.out);
    assert.match(r.out, /バックアップがありません/);
    assert.ok(!mock.state.log.includes('PUT /k/v1/preview/app/customize'));

    r = await run('backup-app.js', [], env);
    assert.strictEqual(r.code, 0, r.out);
    const made = fs.readdirSync(backupDir).filter((n) => !had.includes(n));
    assert.strictEqual(made.length, 1);
    const files = fs.readdirSync(path.join(backupDir, made[0], 'files'));
    assert.ok(files.some((f) => f.endsWith('existing.js')), '既存のJSファイル本体も保存する');

    r = await run('apply-customize.js', ['--yes'], env);
    assert.strictEqual(r.code, 0, r.out);
    const c = mock.state.apps['575'].customize;
    const names = (list) => list.map((x) => x.type === 'URL' ? x.url : x.file.name);
    assert.deepStrictEqual(names(c.desktop.js), ['https://example.com/a.js', 'existing.js', 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js', 'doc-button.js', 'taiyo-open-app.js']);
    assert.deepStrictEqual(names(c.mobile.js), ['https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js', 'doc-button.js', 'round-buttons.js']);
    assert.deepStrictEqual(names(c.desktop.css), ['https://example.com/a.css'], '既存のCSSは残る');

    r = await run('apply-customize.js', ['--yes'], env);   // 2回目：重複して増えない
    assert.strictEqual(r.code, 0, r.out);
    assert.deepStrictEqual(names(mock.state.apps['575'].customize.desktop.js), names(c.desktop.js));

    made.forEach((n) => fs.rmSync(path.join(backupDir, n), { recursive: true, force: true }));
  });
});

test('e2e-check：雛形→差し込み→案件くんの書類保管表に保存→中身確認→後始末', async () => {
  await withMock(async (mock, env) => {
    const tpl = await build();
    // 「段落をまたぐ {{ }}」を含まない、きれいな雛形（正常系用）
    const z0 = await JSZip.loadAsync(tpl.docx);
    z0.file('word/document.xml', (await z0.file('word/document.xml').async('string')).replace('<w:p><w:r><w:t>{{存在しない項目}}</w:t></w:r></w:p>', '').replace('<w:p><w:r><w:t>{{案件</w:t></w:r></w:p><w:p><w:r><w:t>名}}</w:t></w:r></w:p>', ''));
    const cleanDocx = await z0.generateAsync({ type: 'nodebuffer' });
    // 案件くん（575）：項目名と項目コードが違う
    mock.addApp('575', '案件くん', { fields: { c_name: { type: 'SINGLE_LINE_TEXT', code: 'c_name', label: '案件名' }, c_addr: { type: 'SINGLE_LINE_TEXT', code: 'c_addr', label: '現場住所' }, c_client: { type: 'SINGLE_LINE_TEXT', code: 'c_client', label: '発注者' }, c_memo: { type: 'MULTI_LINE_TEXT', code: 'c_memo', label: '備考' } } });
    mock.addRecord('575', { c_name: { type: 'SINGLE_LINE_TEXT', value: '東京Aビル 改修工事' }, c_addr: { type: 'SINGLE_LINE_TEXT', value: '東京都千代田区1-1' }, c_client: { type: 'SINGLE_LINE_TEXT', value: '株式会社<テスト>&サンプル' }, c_memo: { type: 'MULTI_LINE_TEXT', value: '1行目\n2行目' }, 書類保管: { type: 'SUBTABLE', value: [] } });
    mock.addApp('10', '書類マスタ');
    const key = mock.addFile(cleanDocx, '【テスト】雛形.docx');
    const badKey = mock.addFile(tpl.docx, '【テスト】雛形.docx');
    mock.addRecord('10', { 書類名: { value: '打合せ議事次第' }, 工程: { value: '打合せ・見積・提案' }, 使用状況: { value: '使用中' }, 雛形: { value: [{ fileKey: key, name: '【テスト】雛形.docx' }] } });

    let r = await run('e2e-check.js', ['--record=1', '--doc=打合せ議事次第', '--cleanup'], env);
    assert.strictEqual(r.code, 0, r.out);
    assert.match(r.out, /案件名「東京Aビル 改修工事」が書類に入っています/);
    assert.strictEqual(mock.state.records['575'][0]['書類保管'].value.length, 0, '--cleanup で確認用の行を取り除く');

    // --cleanup なしで2回：版が v1 → v2 になる
    await run('e2e-check.js', ['--record=1', '--doc=打合せ議事次第'], env);
    r = await run('e2e-check.js', ['--record=1', '--doc=打合せ議事次第'], env);
    assert.strictEqual(r.code, 0, r.out);
    const rows = mock.state.records['575'][0]['書類保管'].value;
    assert.deepStrictEqual(rows.map((x) => x.value['書類保管_版'].value), ['1', '2']);
    const f = rows[1].value['書類保管_ファイル'].value[0];
    assert.strictEqual(f.name, '東京Aビル 改修工事_打合せ議事次第_v2.docx');
    const zip = await JSZip.loadAsync(mock.state.files[f.fileKey].buf);
    const xml = await zip.file('word/document.xml').async('string');
    assert.match(xml, /打合せ議事次第 v2 /);

    // 異常系：名前違い・段落をまたぐ {{ }} が残る雛形は、NG として知らせる
    mock.addRecord('10', { 書類名: { value: '問題のある雛形' }, 工程: { value: '受注・準備' }, 使用状況: { value: '使用中' }, 雛形: { value: [{ fileKey: badKey, name: '【テスト】雛形.docx' }] } });
    r = await run('e2e-check.js', ['--record=1', '--doc=問題のある雛形', '--cleanup'], env);
    assert.strictEqual(r.code, 5, r.out);
    assert.match(r.out, /NG .*存在しない項目/);
    assert.match(r.out, /NG .*案件名\}\}/);
  });
});

test('add-doc-storage：確認だけでは変えない／バックアップなしは止まる／足すと書類保管表ができ、既存は変わらない／2回目は止まる', async () => {
  await withMock(async (mock, env) => {
    mock.addApp('575', '案件くん', { fields: { c_name: { type: 'SINGLE_LINE_TEXT', code: 'c_name', label: '案件名' } } });
    const backupDir = env.KINTONE_BACKUP_DIR;
    const had = fs.existsSync(backupDir) ? fs.readdirSync(backupDir) : [];
    try {
      let r = await run('add-doc-storage.js', [], env);
      assert.strictEqual(r.code, 0, r.out);
      assert.match(r.out, /確認だけです/);
      assert.deepStrictEqual(Object.keys(mock.state.apps['575'].fields), ['c_name']);

      if (!had.some((n) => n.startsWith('575_'))) {
        r = await run('add-doc-storage.js', ['--yes'], env);
        assert.strictEqual(r.code, 2, r.out);
        assert.match(r.out, /バックアップがありません/);
        assert.deepStrictEqual(Object.keys(mock.state.apps['575'].fields), ['c_name']);
        r = await run('backup-app.js', ['575'], env);
        assert.strictEqual(r.code, 0, r.out);
      }
      r = await run('add-doc-storage.js', ['--yes'], env);
      assert.strictEqual(r.code, 0, r.out);
      assert.match(r.out, /OK 「書類保管」表が追加されました/);
      const f = mock.state.apps['575'].fields;
      assert.deepStrictEqual(Object.keys(f), ['c_name', '書類保管'], '既存の項目は残り、表が1つ足される');
      assert.strictEqual(f['書類保管'].type, 'SUBTABLE');
      assert.deepStrictEqual(Object.keys(f['書類保管'].fields), ['書類保管_書類名', '書類保管_工程', '書類保管_版', '書類保管_作成日', '書類保管_ファイル']);

      r = await run('add-doc-storage.js', ['--yes'], env);
      assert.strictEqual(r.code, 2, r.out);
      assert.match(r.out, /すでにあります/);
    } finally {
      fs.readdirSync(backupDir).filter((n) => !had.includes(n)).forEach((n) => fs.rmSync(path.join(backupDir, n), { recursive: true, force: true }));
    }
  });
});

test('template-tool：list／get／put（確認だけは変えない・差し替えで旧版が残る・旧版は3つまで）', async () => {
  await withMock(async (mock, env) => {
    const tpl = await build();
    mock.addApp('10', '書類マスタ');
    const k1 = mock.addFile(tpl.docx, '議事次第.docx');
    mock.addRecord('10', { 書類名: { value: '打合せ議事次第' }, 工程: { value: '打合せ・見積・提案' }, 表示順: { value: '210' }, 使用状況: { value: '使用中' }, 雛形: { value: [{ fileKey: k1, name: '議事次第.docx' }] } });
    mock.addRecord('10', { 書類名: { value: '見積書' }, 工程: { value: '打合せ・見積・提案' }, 表示順: { value: '220' }, 使用状況: { value: '使用中' }, 雛形: { value: [] } });
    const work = fs.mkdtempSync(path.join(os.tmpdir(), 'kd-tpl-'));
    let r = await run('template-tool.js', ['list'], env);
    assert.strictEqual(r.code, 0, r.out); assert.match(r.out, /打合せ議事次第 ｜ 議事次第\.docx ｜/); assert.match(r.out, /見積書 ｜ 雛形なし/);
    r = await run('template-tool.js', ['get', '--doc=打合せ議事次第', '--out=' + work], env);
    assert.strictEqual(r.code, 0, r.out); assert.ok(fs.readFileSync(path.join(work, '議事次第.docx')).equals(tpl.docx), '取り出した中身が同じ');

    // 直したファイル（{{議事録メモ}} を足したもの）
    const z = await JSZip.loadAsync(tpl.docx);
    z.file('word/document.xml', (await z.file('word/document.xml').async('string')).replace('</w:body>', '<w:p><w:r><w:t>{{議事録メモ}}</w:t></w:r></w:p></w:body>'));
    const fixed = path.join(work, '議事次第_改.docx'); fs.writeFileSync(fixed, await z.generateAsync({ type: 'nodebuffer' }));
    r = await run('template-tool.js', ['put', '--doc=打合せ議事次第', '--file=' + fixed], env);
    assert.strictEqual(r.code, 0, r.out); assert.match(r.out, /確認だけです/); assert.match(r.out, /議事録メモ/);
    assert.strictEqual(mock.state.records['10'][0]['雛形'].value.length, 1, '確認だけでは変わらない');
    // 4回差し替える：先頭が最新・旧版は最新3つまで
    for (let i = 0; i < 4; i++) { r = await run('template-tool.js', ['put', '--doc=打合せ議事次第', '--file=' + fixed, '--yes'], env); assert.strictEqual(r.code, 0, r.out); assert.match(r.out, /OK 差し替え後の雛形/); }
    const names = mock.state.records['10'][0]['雛形'].value.map((f) => f.name);
    assert.strictEqual(names[0], '議事次第_改.docx', '先頭が、使われる雛形');
    assert.strictEqual(names.length, 4, '最新1つ＋旧版3つ');
    assert.ok(names.slice(1).every((n) => /^旧版_\d{8}_/.test(n)), names.join(','));
    r = await run('template-tool.js', ['get', '--doc=存在しない'], env);
    assert.strictEqual(r.code, 1, r.out);
  });
});

test('add-approval-fields：表がなければ止まる／確認だけは変えない／足すと列が増え既存は残る／2回目は何もしない', async () => {
  await withMock(async (mock, env) => {
    mock.addApp('575', '案件くん', { fields: { c_name: { type: 'SINGLE_LINE_TEXT', code: 'c_name', label: '案件名' } } });
    mock.state.apps['575'].previewFields = JSON.parse(JSON.stringify(mock.state.apps['575'].fields));
    let r = await run('add-approval-fields.js', [], env);
    assert.strictEqual(r.code, 2, r.out); assert.match(r.out, /見つかりません/);
    // 書類保管の表がある状態にする（add-doc-storage と同じ手順）
    const backupDir = env.KINTONE_BACKUP_DIR;
    r = await run('backup-app.js', ['575'], env); assert.strictEqual(r.code, 0, r.out);
    r = await run('add-doc-storage.js', ['--yes'], env); assert.strictEqual(r.code, 0, r.out);
    mock.state.apps['575'].previewFields = JSON.parse(JSON.stringify(mock.state.apps['575'].fields));
    r = await run('add-approval-fields.js', [], env);
    assert.strictEqual(r.code, 0, r.out); assert.match(r.out, /確認だけです/); assert.match(r.out, /承認状況/);
    assert.strictEqual(Object.keys(mock.state.apps['575'].fields['書類保管'].fields).length, 5, '確認だけでは変わらない');
    fs.rmSync(backupDir, { recursive: true, force: true }); fs.mkdirSync(backupDir);
    r = await run('add-approval-fields.js', ['--yes'], env);
    assert.strictEqual(r.code, 2, r.out); assert.match(r.out, /バックアップがありません/);
    r = await run('backup-app.js', ['575'], env); assert.strictEqual(r.code, 0, r.out);
    r = await run('add-approval-fields.js', ['--yes'], env);
    assert.strictEqual(r.code, 0, r.out); assert.match(r.out, /OK 列を足しました（5 → 9列/);
    const cols = Object.keys(mock.state.apps['575'].fields['書類保管'].fields);
    assert.deepStrictEqual(cols, ['書類保管_書類名', '書類保管_工程', '書類保管_版', '書類保管_作成日', '書類保管_ファイル', '書類保管_承認状況', '書類保管_承認者', '書類保管_承認日時', '書類保管_PDF']);
    mock.state.apps['575'].previewFields = JSON.parse(JSON.stringify(mock.state.apps['575'].fields));
    r = await run('add-approval-fields.js', ['--yes'], env);
    assert.strictEqual(r.code, 0, r.out); assert.match(r.out, /すでにあります/);
  });
});
