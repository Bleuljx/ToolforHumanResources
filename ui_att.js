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
  // config: checkbox list + 2 inputs, no select/chips/sub
  const checks = doc.querySelectorAll('#machineChecks [data-post]');
  assert(checks.length >= 4, 'checkbox list >= 4');
  assert(Array.from(checks).some(c => c.checked), 'default checked exists');
  assert(!!doc.querySelector('#machineAll'), 'has select-all btn');
  assert(!!doc.querySelector('#machineNone'), 'has deselect-all btn');
  assert(!doc.querySelector('.ex-card'), 'no green ex-card');
  assert(!doc.querySelector('.ex-all'), 'no old ex-all');
  assert(!/剔除岗位/.test(doc.querySelector('#machinePanel').textContent), 'no pinwei label');
  assert(!/上传外包考勤后在此列出岗位/.test(html), 'no empty hint');
  assert(doc.querySelector('#viewAttendance .att-panels').firstElementChild.id === 'machinePanel', 'machine panel first');
  assert(/att-panels\{display:flex;flex-direction:column/.test(html), 'vertical css');
  assert(/rmap-grid\{display:grid;grid-template-columns:repeat\(3/.test(html), 'rmap 3col css');
  assert(/ex-checks\{display:flex;flex-wrap:wrap/.test(html), 'checks wrap css');
  assert(!!doc.querySelector('#machineExtraPosts'), 'has posts input');
  assert(!!doc.querySelector('#machineExtraNames'), 'has names input');
  assert(!doc.querySelector('#machineSel'), 'no select');
  assert(!doc.querySelector('.ex-chip') && !doc.querySelector('.ex-sub'), 'no chips/no sub');
  assert(doc.querySelectorAll('.itip').length >= 3, 'has itip icons');
  assert(doc.querySelectorAll('#rosterPanel [data-rmap]').length === 6, '6 rmap');
  assert(!/说明/.test(doc.querySelector('#rosterPanel').textContent), 'no tip column');
  const vc = doc.querySelector('#viewAttendance').textContent;
  assert(/花名册字段映射/.test(vc) && /外包考勤岗位剔除/.test(vc), 'two sec-h titles');
  assert(/1fr\) minmax\(0,1fr\) minmax\(0,1\.6fr\)/.test(html), 'card ratio');
  const C = win.AttCore;
  assert(C.splitTokens('机师岗, 张三 李四、王五；赵六').length === 5, 'split 5');
  assert(C.makeExcludeFilter({ machineExtraNames: '张三 李四' })({ post: '保洁岗', name: '李四' }) === true, 'name box hits name');
  assert(C.makeExcludeFilter({ machineExtraNames: '张三 李四' })({ post: '李四', name: '孙八' }) === false, 'name box not post');
  assert(C.makeExcludeFilter({ machineExtraPosts: '维保机师' })({ post: '维保机师', name: '孙八' }) === true, 'post box hits post');
  assert(C.makeExcludeFilter({ machineExtraPosts: '维保机师' })({ post: '保洁岗', name: '维保机师' }) === false, 'post box not name');
  assert(C.makeExcludeFilter({ machineShifts: ['机师岗'] })({ post: '保洁岗', name: '钱七' }) === false, 'unmatched kept');
  // 2. 生成
  doc.querySelector('#btnAttRun').click();
  const ok = await waitFor(() => /生成完成/.test(doc.querySelector('#toast').textContent), 30000);
  assert(ok, '生成完成提示出现');

  const btns = Array.from(doc.querySelectorAll('#attResults [data-att-group]'));
  assert(btns.length === 3, '生成 3 个模块下载按钮（外包/机师岗RM/数据汇总），实际 ' + btns.length);
  // 新版：批量下载按钮逐个触发，间隔 120ms，等足够时间下完 4 个文件
  for (const b of btns) b.click();
  await sleep(1200);

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
  assert(!!sumFile && sumFile.name === '【考勤差异】10月差异数据汇总.xlsx', '汇总总表文件名正确（差异数据汇总）');

  // 问题清单在最后
  const cards = Array.from(doc.querySelectorAll('#attResults .result-card'));
  const lastCard = cards[cards.length - 1];
  assert(/问题清单/.test(lastCard.textContent), '问题清单卡在结果区最后');

  // 3. 数据统计看板 + 在职无出勤明细表（独立卡片，不进问题清单）
  const probText = doc.querySelector('#attResults').textContent;
  assert(!/已从外包考勤剔除/.test(probText), '问题清单不含机师岗剔除记录');
  assert(!/打卡校准/.test(probText), '问题清单不含打卡校准 info 记录');
  assert(/外包考勤结果/.test(probText) && /机师岗RM考勤结果/.test(probText) && /数据汇总/.test(probText)
    && /扣款合计（元）/.test(probText), '结果区含三模块大卡（外包/机师岗RM/数据汇总+扣款合计）');
  // 回归：汇总统计 15 列，扣款合计应按金额列（样例 20+10+20=50），不能按次数列
  const sumCard = Array.from(doc.querySelectorAll('#attResults .module-card'))
    .find(c => /数据汇总/.test(c.querySelector('.mc-title').textContent));
  assert(!!sumCard && /<b>50<\/b><span>扣款合计（元）<\/span>/.test(sumCard.innerHTML),
    '数据汇总卡扣款合计=50（实际 ' + (sumCard && sumCard.textContent.match(/\d+(?=扣款合计（元）)/)) + '）');
  // 小字/问号/闪烁：去掉「N 个文件」与 title 提示，悬浮改挂在数字统计区
  assert(!/个文件/.test(doc.querySelector('#attResults').innerHTML), '三张卡均不再显示「N 个文件」小字');
  assert(!/悬浮查看文件清单/.test(doc.querySelector('#attResults').innerHTML), '不再有原生 title 提示（避免闪烁）');
  assert(!/cursor:help/.test(html) && !/\.mc-count\{/.test(html), '去掉 help 问号光标与橙色小字样式');
  const hosts = Array.from(doc.querySelectorAll('#attResults .filelist-host'));
  assert(hosts.length === 2 && hosts.every(e => e.classList.contains('mc-stats')),
    '只有外包/机师岗两张卡的数字统计区可悬浮看文件清单（数据汇总卡无）');
  assert(doc.querySelectorAll('#attResults .filelist-tip').length === 2, '两张卡各带一个文件清单气泡');
  assert(/pointer-events:none/.test(html) && /\.filelist-host:hover \.filelist-tip/.test(html),
    '气泡禁用鼠标事件且由数字区悬浮触发（不闪烁）');
  assert(/在职无出勤人员明细/.test(probText), '结果区含在职无出勤明细表标题');
  assert(/吴十/.test(probText), '在职无出勤明细列出吴十');
  const subArea = doc.querySelector('#attResults').innerHTML;
  const noAttTable = Array.from(doc.querySelectorAll('#attResults .att-subhead')).filter(e => /在职无出勤人员明细/.test(e.textContent))[0];
  assert(!!noAttTable && /蝶城/.test(doc.querySelector('#attResults').textContent) && /<table/.test(subArea), '无出勤明细以表格展示（含蝶城/姓名列）');
  // 问题清单卡片：只含系统计算问题（周九花名册未匹配），不再含无出勤
  const probCard = cards.filter(c => /<span class="rc-city">问题清单<\/span>/.test(c.innerHTML))[0];
  assert(!!probCard, '问题清单卡存在（有花名册未匹配问题）');
  assert(/未匹配到该员工/.test(probCard.textContent), '问题清单含花名册未匹配提示（周九）');
  assert(!/在职无出勤|无任何考勤记录/.test(probCard.textContent), '问题清单不含在职无出勤条目');

  // 4. 用 openpyxl 校验导出文件（verify_att.py）
  console.log(failed ? 'FAILURES: ' + failed : 'ALL_UI_ASSERTS_PASSED');
  process.exit(failed ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
