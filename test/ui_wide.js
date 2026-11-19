/* ui_wide.js — jsdom 端到端：宽表（非循环项导入(宽表)）配置面板 + 生成 + 下载 */
const fs = require('fs');
const path = require('path');
const WS = process.env.NODE_PATH_WORKSPACE || 'C:/Users/liujx121/.workbuddy/binaries/node/workspace/node_modules';
const { JSDOM } = require(path.join(WS, 'jsdom'));

const root = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'dist/人力工具集成平台.html'), 'utf8');
const outDir = path.join(__dirname, 'out_wide');
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
else fs.readdirSync(outDir).forEach(f => { try { fs.unlinkSync(path.join(outDir, f)); } catch (e) { } });

const dom = new JSDOM(html, { runScripts: 'dangerously', url: 'https://local.test/', pretendToBeVisual: true });
const win = dom.window, doc = win.document;

const blobs = [], downloads = [];
win.URL.createObjectURL = function (b) { blobs.push(b); return 'blob:test/' + blobs.length; };
win.URL.revokeObjectURL = function () { };
win.HTMLAnchorElement.prototype.click = function () {
  downloads.push({ name: this.download, idx: parseInt(String(this.href).split('/').pop(), 10) });
};
let confirmMsgs = [];
win.confirm = function (msg) { confirmMsgs.push(String(msg)); return true; };

const sleep = ms => new Promise(r => setTimeout(r, ms));
async function waitFor(fn, timeout) {
  const t0 = Date.now();
  while (Date.now() - t0 < (timeout || 20000)) { if (fn()) return true; await sleep(80); }
  return false;
}
let failed = 0;
function assert(c, m) { console.log((c ? '✓ ' : '✗ ') + m); if (!c) failed++; }

function dropFile(p) {
  const buf = fs.readFileSync(p);
  const f = new win.File([new Uint8Array(buf)], path.basename(p),
    { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const ev = new win.Event('drop', { bubbles: true, cancelable: true });
  ev.dataTransfer = { files: [f] };
  doc.querySelector('#drop').dispatchEvent(ev);
}

async function saveBlob(dl, fallbackName) {
  if (!dl) return null;
  const b = blobs[dl.idx - 1];
  const ab = await b.arrayBuffer();
  const name = dl.name || fallbackName;
  fs.writeFileSync(path.join(outDir, name), Buffer.from(ab));
  return name;
}

(async function () {
  await waitFor(() => doc.querySelector('#drop'));
  await sleep(250);

  // 1. 四个关系配置瓦片
  const tiles = doc.querySelectorAll('#tiles .tile');
  assert(tiles.length === 4, '关系配置 4 个瓦片（一行4个），实际 ' + tiles.length);
  const wideTile = doc.querySelector('.tile[data-panel="wide"]');
  assert(!!wideTile, '宽表瓦片存在');

  // 2. 上传算薪表
  dropFile(path.join(__dirname, 'wide_sample.xlsx'));
  await waitFor(() => doc.querySelectorAll('.file-item').length === 1);
  await sleep(200);

  // 3. 打开宽表面板
  wideTile.dispatchEvent(new win.MouseEvent('click', { bubbles: true }));
  await sleep(250);
  const sels = Array.from(doc.querySelectorAll('#panel-wide select[data-wide]'));
  assert(sels.length === 28, '宽表映射 28 个项目列下拉，实际 ' + sels.length);
  function selVal(tpl) {
    const s = doc.querySelector('#panel-wide select[data-wide="' + tpl + '"]');
    return s ? s.value : null;
  }
  assert(selVal('基本工资') === '@baseSalary', '基本工资 → 考勤字段@baseSalary（实际 ' + selVal('基本工资') + '）');
  assert(selVal('定额值班费') === '@dutyFee', '定额值班费 → @dutyFee（自动计算值）');
  assert(selVal('租房补贴') === 'WJ租房补贴', '租房补贴 → WJ租房补贴（去前缀匹配）');
  assert(selVal('考核罚款') === 'WY考核罚款', '考核罚款 → WY考核罚款');
  assert(selVal('五险一金合计扣款（个人）') === '五险一金合计扣款（个人）', '个人五险一金精确匹配');
  const unmatched = ['商业险', '高温补贴', '推荐奖励'].filter(t => !selVal(t));
  assert(unmatched.length === 3, '商业险/高温补贴/推荐奖励 未匹配（会走确认弹窗）');
  assert(/列未选择取数来源/.test(doc.querySelector('#panel-wide').textContent), '面板有未匹配提示');

  // 4. 手动给商业险选「服务费」再清空（验证手选生效）
  const syx = doc.querySelector('#panel-wide select[data-wide="商业险"]');
  syx.value = '服务费';
  syx.dispatchEvent(new win.Event('change', { bubbles: true }));
  await sleep(200);
  assert(selVal('商业险') === '服务费', '商业险手动映射到服务费');
  const syx2 = doc.querySelector('#panel-wide select[data-wide="商业险"]');
  syx2.value = '';
  syx2.dispatchEvent(new win.Event('change', { bubbles: true }));
  await sleep(200);

  // 5. 业务日期
  const wd = doc.querySelector('#wideDate');
  wd.value = '2026-11-16';
  wd.dispatchEvent(new win.Event('change', { bubbles: true }));
  await sleep(150);

  // 6. 填窄表参数并生成（宽表未匹配列 → 应弹确认框）
  const task = doc.querySelector('#impTask');
  task.value = 'RWH-2026-11'; task.dispatchEvent(new win.Event('change', { bubbles: true }));
  const dateEl = doc.querySelector('#impDate');
  dateEl.value = '2026-11-16'; dateEl.dispatchEvent(new win.Event('change', { bubbles: true }));
  await sleep(150);

  doc.querySelector('#btnRun').dispatchEvent(new win.MouseEvent('click', { bubbles: true }));
  const okRun = await waitFor(() => doc.querySelectorAll('.result-card').length === 1);
  assert(okRun, '生成结果卡片');
  if (!okRun) { console.error(doc.querySelector('#results').textContent); process.exit(1); }
  assert(confirmMsgs.some(m => m.indexOf('宽表还有以下列没有选择取数来源') >= 0), '宽表未配齐弹了确认框');
  assert(/商业险/.test(confirmMsgs.find(m => m.indexOf('宽表') >= 0) || ''), '确认框列出商业险');

  // 7. 结果卡片统计与下载
  const card = doc.querySelector('.result-card');
  assert(/4宽表行/.test(card.textContent.replace(/\s+/g, '')), '结果卡片显示宽表行 4');
  const dlWide = card.querySelector('[data-dl="3"]');
  assert(!!dlWide && /宽表/.test(dlWide.textContent), '第 4 个下载按钮为宽表');
  dlWide.dispatchEvent(new win.MouseEvent('click', { bubbles: true }));
  await sleep(100);
  const wideDl = downloads.find(d => /非循环项导入\(宽表\)/.test(d.name));
  assert(!!wideDl, '宽表文件名=非循环项导入(宽表).xlsx（实际 ' + (wideDl && wideDl.name) + '）');
  if (wideDl) {
    const saved = await saveBlob(wideDl);
    assert(!!saved, '宽表 xlsx 已保存到输出目录');
  }

  console.log('\n' + (failed ? '有 ' + failed + ' 项失败' : 'ALL PASS') + '，输出：' + outDir);
  process.exit(failed ? 1 : 0);
})();
