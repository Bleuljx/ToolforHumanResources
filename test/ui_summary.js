/* ui_summary.js — 复现用户场景：localStorage 里有旧的"未匹配"配置 → 上传 → 打开导入面板 */
const fs = require('fs');
const path = require('path');
const WS = 'C:/Users/liujx121/.workbuddy/binaries/node/workspace/node_modules';
const { JSDOM } = require(path.join(WS, 'jsdom'));
const root = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'dist/人力工具集成平台.html'), 'utf8');

const dom = new JSDOM(html, {
  runScripts: 'dangerously', url: 'https://local.test/', pretendToBeVisual: true,
  beforeParse(w) {
    // 预置旧版本残留的 localStorage 配置：应发/实发合计是"未匹配"状态
    const stale = {
      version: 1,
      attendance: {}, category: {},
      importCfg: {
        task: '', date: '',
        map: {
          'WY应发合计': { value: '', conf: 'none' },
          'WY实发合计': { value: '', conf: 'none' }
        }
      }
    };
    w.localStorage.setItem('salaryTool.cfg.v1', JSON.stringify(stale));
  }
});
const win = dom.window, doc = win.document;
win.confirm = function () { return true; };   // 宽表未配齐的确认框默认继续
const sleep = ms => new Promise(r => setTimeout(r, ms));
async function waitFor(fn, timeout) {
  const t0 = Date.now();
  while (Date.now() - t0 < (timeout || 20000)) { if (fn()) return true; await sleep(80); }
  return false;
}
let failed = 0;
const assert = (c, m) => { console.log((c ? '✓ ' : '✗ ') + m); if (!c) failed++; };
function dropFile(p, name) {
  const buf = fs.readFileSync(p);
  const f = new win.File([new Uint8Array(buf)], name || path.basename(p),
    { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const ev = new win.Event('drop', { bubbles: true, cancelable: true });
  ev.dataTransfer = { files: [f] };
  doc.querySelector('#drop').dispatchEvent(ev);
}
(async function () {
  await waitFor(() => doc.querySelector('#drop'));
  await sleep(250);
  dropFile(path.join(__dirname, 'sample_summary.xlsx'));
  assert(await waitFor(() => doc.querySelectorAll('.file-item').length === 1), '样例表解析进列表');
  // 打开第三个面板（非循环项导入）
  doc.querySelectorAll('.tile')[2].click();
  await sleep(300);
  const rows = Array.from(doc.querySelectorAll('#panel-import .map-table tbody tr'));
  const find = n => rows.find(r => r.querySelector('.col-out') && r.querySelector('.col-out').textContent === n);
  for (const n of ['WY应发合计']) {
    const tr = find(n);
    assert(!!tr, '面板列出列「' + n + '」');
    if (tr) {
      const badge = tr.querySelector('.badge').textContent;
      const inp = tr.querySelector('.opt-input').value;
      assert(badge === '不报送', '「' + n + '」匹配标记=不报送（实际:' + badge + '）');
      assert(inp === '不报送', '「' + n + '」输入框=不报送（实际:' + inp + '）');
    }
  }
  // WY实发合计 在本样例里位于区间末尾之后：不列入映射面板，也不应产生"区间外项目"警告
  assert(!find('WY实发合计'), '「WY实发合计」在区间外，不进映射面板');
  const hint = doc.querySelector('#panel-import .hint');
  assert(!/未匹配/.test(hint ? hint.textContent : ''), '提示里不再出现"还有 N 项未匹配"');
  // 点击开始生成，确认不阻断且导出文件不含应发/实发
  const t = doc.querySelector('#impTask'); if (t) { t.value = 'RWH-01'; t.dispatchEvent(new win.Event('change', { bubbles: true })); }
  const d = doc.querySelector('#impDate'); if (d) { d.value = '2026-11-17'; d.dispatchEvent(new win.Event('change', { bubbles: true })); }
  await sleep(150);
  const btn = doc.querySelector('#btnRun'); if (btn) btn.click();
  await sleep(2500);
  const toast = doc.querySelector('#toast');
  console.log('toast:', toast ? toast.textContent : '(none)');
  const cards = doc.querySelectorAll('.result-card');
  assert(cards.length === 1, '生成出结果卡片（未被未匹配阻断）');
  process.exit(failed ? 1 : 0);
})();
