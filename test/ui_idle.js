/* ui_idle.js — jsdom 端到端：出勤异常核对模块（考勤导出+花名册+上期反馈 → 异常明细） */
const fs = require('fs');
const path = require('path');
const WS = process.env.NODE_PATH_WORKSPACE || 'C:/Users/liujx121/.workbuddy/binaries/node/workspace/node_modules';
const { JSDOM } = require(path.join(WS, 'jsdom'));

const root = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'dist/人力工具集成平台.html'), 'utf8');
const outDir = path.join(__dirname, 'idle_out');
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
let failed = 0;
function assert(c, m) { console.log((c ? '✓ ' : '✗ ') + m); if (!c) failed++; }

function dropInto(zoneKey, p) {
  const buf = fs.readFileSync(p);
  const f = new win.File([new Uint8Array(buf)], path.basename(p),
    { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const ev = new win.Event('drop', { bubbles: true, cancelable: true });
  ev.dataTransfer = { files: [f] };
  doc.querySelector('#idleDrop-' + zoneKey).dispatchEvent(ev);
}

async function saveBlob(d) {
  const ab = await d.blob.arrayBuffer();
  fs.writeFileSync(path.join(outDir, d.name), Buffer.from(ab));
  return path.join(outDir, d.name);
}

(async () => {
  await sleep(500);

  // 0. 导航到超7天未出勤
  doc.querySelector('.nav-item[data-view="absent"]').click();
  await sleep(100);
  assert(!doc.querySelector('#viewAbsent').classList.contains('hidden'), '视图切换到超7天未出勤');
  assert(doc.querySelectorAll('#idleZones .att-zone').length === 3, '三个上传区（考勤统计/花名册/上期反馈）');

  // 1. 上传三个文件
  dropInto('att', path.join(__dirname, '考勤统计结果导出.xlsx'));
  dropInto('emp', path.join(__dirname, '员工信息维护.xlsx'));
  dropInto('prev', path.join(__dirname, '上期结果表.xlsx'));
  await sleep(700);

  const zonesTxt = doc.querySelector('#idleZones').textContent;
  assert(/识别到 9 条考勤记录/.test(zonesTxt), '考勤导出识别 9 条记录');
  assert(/识别到 12 名员工/.test(zonesTxt), '花名册识别 12 名员工');
  assert(/识别到 5 条上期反馈/.test(zonesTxt), '上期反馈识别 5 条');

  // 2. 开始核对
  doc.querySelector('#btnIdleRun').click();
  await sleep(200);
  const resTxt = doc.querySelector('#idleResults').textContent;

  // 3. 周期：月份 2026-10，满勤 20 天 → 10-01 ~ 10-20 共 20 天
  assert(/2026-10-01 ~ 2026-10-20/.test(resTxt) && /共 <\/?[^>]*>?20\s*天|共\s*20\s*天/.test(resTxt.replace(/<[^>]+>/g, '')), '核算周期 2026-10-01~10-20 共20天（按数据填）');

  // 4. KPI
  const kpis = Array.from(doc.querySelectorAll('#idleResults .idle-kpi')).map(k => k.textContent.replace(/\s/g, ''));
  assert(kpis.some(t => /本期异常人数1人/.test(t)), 'KPI 本期异常=1 人');
  assert(kpis.some(t => /合计无出勤天数12天/.test(t)), 'KPI 合计无出勤=12 天');
  assert(kpis.some(t => /导入人数（去重）9人/.test(t)), 'KPI 导入人数=9 人');
  assert(kpis.some(t => /无考勤记录1人/.test(t)), 'KPI 无考勤记录=1 人');

  // 5. 主表内容：王某 12 天 / 已扣请假 6
  const tbTxt = doc.querySelector('#idleTb').textContent;
  assert(/钟公庙蝶城/.test(tbTxt) && /王某/.test(tbTxt), '异常表含钟公庙蝶城分组和王某');
  assert(/12/.test(tbTxt), '王某无出勤=12 天（18-6）');
  assert(/-6/.test(tbTxt), '已扣请假显示 -6');
  assert(!/李某/.test(tbTxt), '李某请假扣完被剔除（不在异常表）');
  assert(!/赵某/.test(tbTxt) && !/孙某/.test(tbTxt) && !/冯某/.test(tbTxt), '离职/入职晚/范围外的人不进异常表');
  assert(!doc.querySelector('#idleTb').textContent.includes('钱某'), '满勤的钱某不上榜');

  // 6. 蝶城→项目分组行
  assert(doc.querySelector('#idleTb .idle-g-city'), '蝶城分组行存在');
  assert(doc.querySelector('#idleTb .idle-g-proj'), '项目分组行存在');

  // 7. 已剔除备查：4 人（赵某+郑某 本期离职，张某+刘某 上期已报离/离职）
  const leftTxt = (doc.querySelector('#idleLeftCard') || { textContent: '' }).textContent;
  assert(/已剔除/.test(leftTxt) && /4/.test(leftTxt), '已剔除备查 4 人');
  assert(/赵某/.test(leftTxt) && /郑某/.test(leftTxt) && /张某/.test(leftTxt) && /刘某/.test(leftTxt), '备查名单含赵某/郑某/张某/刘某');

  // 8. 页脚统计
  const footR = doc.querySelector('#idleFootR').textContent;
  assert(/扣减 2 人、剔除已报离 2 人/.test(footR), '页脚：上期反馈扣减2人/剔除已报离2人');
  assert(/已剔除本期离职 2 人/.test(footR), '页脚：本期离职剔除2人');
  assert(/排除名单命中 1 人/.test(footR), '页脚：排除名单（机师岗）命中1人（陈某）');
  assert(/排除 2 人/.test(footR), '页脚：花名册校验排除2人（孙某/周某）');

  // 9. 无考勤记录 Tab：吴某 17 天
  const noTab = doc.querySelector('#idleTabs [data-tab="noatt"]');
  assert(!!noTab && /1/.test(noTab.textContent), '无考勤记录 Tab 存在（1 人）');
  noTab.click();
  await sleep(100);
  const noTxt = doc.querySelector('#idleTabNoatt').textContent;
  assert(/吴某/.test(noTxt) && /17/.test(noTxt), '无考勤记录吴某 17 天（20-3）');
  doc.querySelector('#idleTabs [data-tab="main"]').click();

  // 10. 筛选与搜索
  const selCity = doc.querySelector('#idleCity');
  assert(selCity && Array.from(selCity.options).some(o => o.text === '钟公庙蝶城'), '蝶城下拉含钟公庙蝶城');
  const kw = doc.querySelector('#idleKw');
  kw.value = '不存在的名字';
  kw.dispatchEvent(new win.Event('input', { bubbles: true }));
  await sleep(50);
  assert(doc.querySelector('#idleCnt').textContent === '0', '搜索无匹配 → 0 人');
  kw.value = '王某';
  kw.dispatchEvent(new win.Event('input', { bubbles: true }));
  await sleep(50);
  assert(doc.querySelector('#idleCnt').textContent === '1', '搜索王某 → 1 人');
  kw.value = '';
  kw.dispatchEvent(new win.Event('input', { bubbles: true }));

  // 11. 导出
  doc.querySelector('#btnIdleExport').click();
  await sleep(300);
  const dl = downloads.find(d => /^超7天无出勤异常明细-\d{12}\.xlsx$/.test(d.name));
  assert(!!dl, '导出文件名=超7天无出勤异常明细-时间戳.xlsx');
  if (dl) await saveBlob(dl);

  console.log(failed ? ('FAILED: ' + failed) : 'ALL PASS');
  process.exit(failed ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
