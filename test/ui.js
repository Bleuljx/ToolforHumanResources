/* ui.js — jsdom 端到端：两个蝶城、含脏数据、三个下载 + ZIP + 问题清单 */
const fs = require('fs');
const path = require('path');
const WS = process.env.NODE_PATH_WORKSPACE || 'C:/Users/liujx121/.workbuddy/binaries/node/workspace/node_modules';
const { JSDOM } = require(path.join(WS, 'jsdom'));

const root = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'dist/人力工具集成平台.html'), 'utf8');
const outDir = path.join(__dirname, 'out');
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
win.confirm = function () { return true; };

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

  // ---- 上传两个蝶城 ----
  dropFile(path.join(__dirname, '麓湖蝶城2026年11月算薪表.xlsx'));
  await waitFor(() => doc.querySelectorAll('.file-item').length === 1);
  dropFile(path.join(__dirname, '花城蝶外2026年11月算薪表.xlsx'));
  const ok2 = await waitFor(() => doc.querySelectorAll('.file-item').length === 2, 10000);
  assert(ok2, '两个文件都解析进列表');
  const cities = Array.from(doc.querySelectorAll('.fi-city')).map(i => i.value);
  assert(cities[0] === '麓湖蝶城' && cities[1] === '花城蝶外', '蝶城识别：' + cities.join(' / '));

  // ---- 打开映射面板，确认自动匹配已有值 ----
  doc.querySelectorAll('.tile')[2].dispatchEvent(new win.MouseEvent('click', { bubbles: true }));
  await sleep(200);
  const inputs = Array.from(doc.querySelectorAll('#panel-import .opt-input'));
  const filled = inputs.filter(i => i.value && i.value !== '不报送');
  assert(filled.length >= 7, '映射表已自动填值（' + filled.length + '/' + inputs.length + '）：' + inputs.map(i => i.value).join(' | '));
  const badges = Array.from(doc.querySelectorAll('#panel-import .badge')).map(b => b.textContent);
  assert(badges.every(b => ['一致','去前缀','归一','推测','未匹配'].indexOf(b) >= 0), '匹配状态标记正常：' + badges.join(','));
  const impLabels = Array.from(doc.querySelectorAll('#panel-import .col-out')).map(td => td.textContent);
  assert(impLabels.filter(t => t.indexOf('考勤') >= 0).length === 7, '映射面板列出 7 个考勤字段：' + impLabels.filter(t => t.indexOf('考勤') >= 0).join(' | '));
  const baseInput = doc.querySelector('#panel-import .opt-input[data-src="@baseSalary"]');
  assert(!!baseInput, '「基本薪资」考勤字段映射输入存在');
  if (baseInput) {
    baseInput.value = '基本工资';
    baseInput.dispatchEvent(new win.Event('change', { bubbles: true }));
  }
  await sleep(250);
  const baseMapped = (JSON.parse(win.localStorage.getItem('salaryTool.cfg.v1') || '{}').importCfg?.map?.['@baseSalary'] || {}).value;
  assert(baseMapped === '基本工资', '基本薪资手动映射已保存：' + baseMapped);

  // ---- 工资类目一键表面板：备注兜底区间 + 忽略列 ----
  doc.querySelectorAll('.tile')[1].dispatchEvent(new win.MouseEvent('click', { bubbles: true }));
  await sleep(250);
  const fbS = doc.querySelector('#panel-category select[data-fb="start"]');
  const fbE = doc.querySelector('#panel-category select[data-fb="end"]');
  assert(!!fbS && !!fbE, '备注兜底区间两个下拉存在');
  assert(fbS && Array.from(fbS.options).some(o => o.value === '其他补款'), '兜底起下拉含「其他补款」选项');
  const igItems = doc.querySelectorAll('#panel-category .ig-item');
  assert(igItems.length >= 30, '忽略列勾选列表出全部表头（' + igItems.length + ' 列）');
  const dayCb = doc.querySelector('#panel-category input[data-dayauto]');
  assert(!!dayCb && dayCb.checked, '考勤明细区域默认自动识别');
  const dayCellTxt = doc.querySelector('#panel-category .day-cell');
  assert(!!dayCellTxt && /「1」~「31」共 31 列/.test(dayCellTxt.textContent), '考勤明细区域自动识别到 1~31 日共 31 列：' + (dayCellTxt ? dayCellTxt.textContent.trim().slice(0, 80) : ''));
  const cbHour = doc.querySelector('#panel-category input[data-ig="工时"]');
  assert(!!cbHour && cbHour.checked, '「工时」按默认配置已勾选');
  const cbSup = doc.querySelector('#panel-category input[data-ig="供应商"]');
  if (cbSup) { cbSup.checked = true; cbSup.dispatchEvent(new win.Event('change', { bubbles: true })); }
  await sleep(300);
  const savedIg = (JSON.parse(win.localStorage.getItem('salaryTool.cfg.v1') || '{}').ignoreCols) || [];
  assert(savedIg.indexOf('供应商') >= 0, '勾选忽略列已保存到配置：' + JSON.stringify(savedIg));
  const search = doc.querySelector('#igSearch');
  search.value = '工时'; search.dispatchEvent(new win.Event('input', { bubbles: true }));
  const visible = Array.from(doc.querySelectorAll('#panel-category .ig-item')).filter(l => l.style.display !== 'none');
  assert(visible.length >= 1 && visible.every(l => (l.dataset.name || '').indexOf('工时') >= 0), '忽略列搜索过滤生效（' + visible.length + ' 项可见）');
  // 恢复面板状态：重新打开导入面板（后面还要用）
  doc.querySelectorAll('.tile')[2].dispatchEvent(new win.MouseEvent('click', { bubbles: true }));
  await sleep(200);

  // ---- 填参数 + 生成 ----
  const task = doc.querySelector('#impTask');
  task.value = 'RWH-2026-11'; task.dispatchEvent(new win.Event('change', { bubbles: true }));
  const dateEl = doc.querySelector('#impDate');
  dateEl.value = '2026-11-16'; dateEl.dispatchEvent(new win.Event('change', { bubbles: true }));
  await sleep(150);

  doc.querySelector('#btnRun').dispatchEvent(new win.MouseEvent('click', { bubbles: true }));
  const okRun = await waitFor(() => doc.querySelectorAll('.result-card').length === 2);
  assert(okRun, '生成两个蝶城的结果卡片');
  if (!okRun) { console.error(doc.querySelector('#results').textContent); process.exit(1); }

  const cards = Array.from(doc.querySelectorAll('.result-card'));
  console.log('   卡片1:', cards[0].textContent.replace(/\s+/g, ' ').slice(0, 130));
  console.log('   卡片2:', cards[1].textContent.replace(/\s+/g, ' ').slice(0, 130));

  // ---- 逐个下载 ----
  for (let c = 0; c < cards.length; c++) {
    const btns = cards[c].querySelectorAll('[data-dl]');
    for (let i = 0; i < btns.length; i++) {
      const before = downloads.length;
      btns[i].dispatchEvent(new win.MouseEvent('click', { bubbles: true }));
      await sleep(60);
      if (downloads.length > before) await saveBlob(downloads[downloads.length - 1]);
    }
  }
  const excel = fs.readdirSync(outDir).filter(f => f.endsWith('.xlsx'));
  assert(excel.length === 8, '导出 8 个 xlsx（2 蝶城 × 4 表）：' + excel.join(', '));
  assert(excel.some(f => /非循环项导入\(宽表\)/.test(f)), '含非循环项导入(宽表).xlsx');
  assert(excel.every(f => /蝶城|蝶外/.test(f)), '文件名带蝶城前缀');

  // ---- ZIP ----
  const zipBtn = cards[0].querySelector('[data-zip]');
  zipBtn.dispatchEvent(new win.MouseEvent('click', { bubbles: true }));
  const okZip = await waitFor(() => downloads.some(d => /\.zip$/.test(d.name)), 10000);
  assert(okZip, 'ZIP 导出成功');
  if (okZip) await saveBlob(downloads.filter(d => /\.zip$/.test(d.name))[0]);

  // ---- 问题清单 ----
  const pbBtn = doc.querySelector('#btnProblems');
  assert(!!pbBtn, '有问题时出现“下载问题清单”按钮');
  if (pbBtn) {
    pbBtn.dispatchEvent(new win.MouseEvent('click', { bubbles: true }));
    await waitFor(() => downloads.some(d => /问题清单/.test(d.name)), 8000);
    await saveBlob(downloads.filter(d => /问题清单/.test(d.name))[0]);
  }

  // ---- 配置持久化 ----
  const saved = JSON.parse(win.localStorage.getItem('salaryTool.cfg.v1') || '{}');
  assert(saved.importCfg && saved.importCfg.task === 'RWH-2026-11', '配置记住核算任务');

  console.log('\n' + (failed ? '有 ' + failed + ' 项失败' : '全部通过') + '，输出：' + outDir);
  process.exit(failed ? 1 : 0);
})();
