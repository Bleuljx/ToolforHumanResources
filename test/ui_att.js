/* ui_att.js — jsdom 端到端：考勤处理模块（SHR+RM+花名册 → 蝶城结果文件+汇总） */
const fs = require('fs');
const path = require('path');
const WS = process.env.NODE_PATH_WORKSPACE || 'C:/Users/liujx121/.workbuddy/binaries/node/workspace/node_modules';
const { JSDOM } = require(path.join(WS, 'jsdom'));

const root = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'dist/人力工具集成平台.html'), 'utf8');
const outDir = path.join(__dirname, 'att_out');
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
else fs.readdirSync(outDir).forEach(f => { try { fs.unlinkSync(path.join(outDir, f)); } catch (e) { } });

const dom = new JSDOM(html, { runScripts: 'dangerously', url: 'https://local.test/', pretendToBeVisual: true });
const win = dom.window, doc = win.document;

const blobs = [], downloads = [];
win.URL.createObjectURL = function (b) { blobs.push(b); return 'blob:test/' + blobs.length; };
win.URL.revokeObjectURL = function () { };
win.HTMLAnchorElement.prototype.click = function () {
  downloads.push({ name: this.download, blob: blobs[parseInt(String(this.href).split('/').pop(), 10) - 1] });
};

const sleep = ms => new Promise(r => setTimeout(r, ms));
async function waitFor(fn, timeout) {
  const t0 = Date.now();
  while (Date.now() - t0 < (timeout || 20000)) { if (fn()) return true; await sleep(60); }
  return false;
}
let failed = 0;
function assert(c, m) { console.log((c ? '✓ ' : '✗ ') + m); if (!c) failed++; }

function dropInto(zone, p) {
  const buf = fs.readFileSync(p);
  const f = new win.File([new Uint8Array(buf)], path.basename(p),
    { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const ev = new win.Event('drop', { bubbles: true, cancelable: true });
  ev.dataTransfer = { files: [f] };
  doc.querySelector('#drop-' + zone).dispatchEvent(ev);
}

async function saveBlob(d) {
  const ab = await d.blob.arrayBuffer();
  fs.writeFileSync(path.join(outDir, d.name), Buffer.from(ab));
  return path.join(outDir, d.name);
}

(async () => {
  await sleep(500);

  // 1. 上传三个文件
  dropInto('shr', path.join(__dirname, '钟公庙蝶城10月外包考勤.xlsx'));
  dropInto('rm', path.join(__dirname, '钟公庙蝶城10月RM考勤.xlsx'));
  dropInto('roster', path.join(__dirname, '人员花名册10月.xlsx'));
  await sleep(600);

  assert(doc.querySelectorAll('#list-shr .file-item').length === 1, 'SHR 文件卡片出现');
  assert(doc.querySelectorAll('#list-rm .file-item').length === 1, 'RM 文件卡片出现');
  assert(doc.querySelectorAll('#list-roster .file-item').length === 1, '花名册卡片出现');
  assert(/识别到 9 名员工/.test(doc.querySelector('#list-roster').textContent), '花名册识别 9 人（自动猜映射）');
  assert(doc.querySelectorAll('#machinePanel [data-machine]').length >= 4, '机师岗清单列出岗位');
  const machineChecked = doc.querySelector('#machinePanel input[data-machine][checked], #machinePanel input[data-machine]:checked');
  assert(!!machineChecked, '含「机师」的岗位已默认勾选');

  // 2. 生成
  doc.querySelector('#btnAttRun').click();
  const ok = await waitFor(() => /生成完成/.test(doc.querySelector('#toast').textContent), 30000);
  assert(ok, '生成完成提示出现');

  const btns = Array.from(doc.querySelectorAll('#attResults [data-att-dl]'));
  assert(btns.length === 4, '生成 4 个下载按钮（钟公庙外包 + 钟公庙机师岗 + 海潮映月机师岗 + 汇总总表），实际 ' + btns.length);
  btns.forEach(b => b.click());
  await sleep(100);

  const dls = downloads.filter(d => d.name.endsWith('.xlsx'));
  assert(dls.length === 4, '生成 4 个 xlsx，实际 ' + dls.length);
  for (const d of dls) await saveBlob(d);
  const cityFile = dls.find(d => /钟公庙蝶城2026-10-01~10-05\.xlsx/.test(d.name));
  const rmFile = dls.find(d => /钟公庙蝶城机师岗2026-10-01~10-05\.xlsx/.test(d.name));
  const rm2File = dls.find(d => /海潮映月蝶城机师岗/.test(d.name));
  const sumFile = dls.find(d => /汇总/.test(d.name));
  assert(!!cityFile && cityFile.name === '【考勤差异】钟公庙蝶城2026-10-01~10-05.xlsx', '外包文件名=蝶城+核算周期（实际 ' + (cityFile && cityFile.name) + '）');
  assert(!!rmFile && rmFile.name === '【考勤差异】钟公庙蝶城机师岗2026-10-01~10-05.xlsx', '机师岗文件名=蝶城+机师岗+核算周期（实际 ' + (rmFile && rmFile.name) + '）');
  assert(!!rm2File, 'RM 原始文件里混入的其他蝶城（海潮映月）按记录拆分出独立机师岗文件');
  assert(!!sumFile && sumFile.name === '【考勤差异】10月迟到早退汇总.xlsx', '汇总总表文件名正确');

  // 问题清单在最后
  const cards = Array.from(doc.querySelectorAll('#attResults .result-card'));
  const lastCard = cards[cards.length - 1];
  assert(/问题清单/.test(lastCard.textContent), '问题清单卡在结果区最后');

  // 3. 数据统计看板 + 在职无出勤明细表（独立卡片，不进问题清单）
  const probText = doc.querySelector('#attResults').textContent;
  assert(!/已从外包考勤剔除/.test(probText), '问题清单不含机师岗剔除记录');
  assert(!/打卡校准/.test(probText), '问题清单不含打卡校准 info 记录');
  assert(/数据统计/.test(probText) && /主表行数/.test(probText) && /扣款合计（元）/.test(probText), '结果区含数据统计看板');
  assert(/在职无出勤人员明细/.test(probText), '结果区含在职无出勤明细表标题');
  assert(/吴十/.test(probText), '在职无出勤明细列出吴十');
  const noAttCard = cards.filter(c => /在职无出勤人员明细/.test(c.textContent))[0];
  assert(!!noAttCard && /<table/.test(noAttCard.innerHTML) && /蝶城/.test(noAttCard.textContent), '无出勤明细以表格展示（含蝶城/姓名列）');
  // 问题清单卡片：只含系统计算问题（周九花名册未匹配），不再含无出勤
  const probCard = cards.filter(c => /<span class="rc-city">问题清单<\/span>/.test(c.innerHTML))[0];
  assert(!!probCard, '问题清单卡存在（有花名册未匹配问题）');
  assert(/未匹配到该员工/.test(probCard.textContent), '问题清单含花名册未匹配提示（周九）');
  assert(!/在职无出勤|无任何考勤记录/.test(probCard.textContent), '问题清单不含在职无出勤条目');

  // 4. 用 openpyxl 校验导出文件（verify_att.py）
  console.log(failed ? 'FAILURES: ' + failed : 'ALL_UI_ASSERTS_PASSED');
  process.exit(failed ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
