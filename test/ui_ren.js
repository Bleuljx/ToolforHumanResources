/* ui_ren.js — jsdom 端到端：文件重命名模块
 * 场景：花名册4人（含两名张三）+ 三个模块 zip
 * 覆盖：zip 解包 / 姓名匹配 / 编码匹配 / 撞名按项目&调度中心消歧 /
 *       未匹配失败 / 目标重名冲突跳过 / 缺失检查 / 问题清单弹窗+xlsx / zip 导出
 */
const fs = require('fs');
const path = require('path');
const WS = process.env.NODE_PATH_WORKSPACE || 'C:/Users/liujx121/.workbuddy/binaries/node/workspace/node_modules';
const { JSDOM } = require(path.join(WS, 'jsdom'));

const root = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'dist/人力工具集成平台.html'), 'utf8');
const downloads = [];

const dom = new JSDOM(html, { runScripts: 'dangerously', url: 'https://local.test/', pretendToBeVisual: true });
const win = dom.window, doc = win.document;

const blobs = [];
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

function dropInto(sel, p) {
  const buf = fs.readFileSync(p);
  const f = new win.File([new Uint8Array(buf)], path.basename(p),
    { type: 'application/octet-stream' });
  const ev = new win.Event('drop', { bubbles: true, cancelable: true });
  ev.dataTransfer = { files: [f] };
  doc.querySelector(sel).dispatchEvent(ev);
}

function utf8Bytes(str) { return Array.from(new (require('util').TextEncoder)().encode(str)); }
function zipContains(u8, str) {
  const needle = utf8Bytes(str);
  outer: for (let i = 0; i <= u8.length - needle.length; i++) {
    for (let j = 0; j < needle.length; j++) if (u8[i + j] !== needle[j]) continue outer;
    return true;
  }
  return false;
}

(async () => {
  await sleep(500);

  // 1. 切到文件重命名视图
  doc.querySelector('.nav-item[data-view="rename"]').click();
  assert(!doc.querySelector('#viewRename').classList.contains('hidden'), '视图切换到文件重命名');
  assert(!!doc.querySelector('#renRosterBox .drop'), '花名册上传窗存在');
  const zones = doc.querySelectorAll('#renGrid .drop');
  assert(zones.length === 9, '9 个资料模块窗口（其他档案材料合并后），实际 ' + zones.length);
  const segs = doc.querySelectorAll('#renTplSegs .ren-seg');
  assert(segs.length === 1 && segs[0].querySelector('.ren-seg-sel').value === '姓名',
    '默认自定义段=姓名（可编辑段只有姓名，排在前）');
  const locks = doc.querySelectorAll('#renTplSegs .ren-seg-lock');
  assert(locks.length === 2 && locks[0].textContent === '工号' && locks[1].textContent === '类型',
    '固定尾段=工号-类型，锁定不可编辑且排在最后');
  assert(!locks[0].querySelector('.ren-del') && !locks[0].querySelector('.ren-seg-sel'),
    '锁定段没有删除键和下拉框');
  assert(doc.querySelectorAll('#renTplSegs .ren-seg-c').length === 2, '段间显示 2 个独立连接符 -（姓名-工号-类型）');
  assert(doc.querySelectorAll('#renTplSegs .ren-del').length === 1 &&
    !!doc.querySelector('#renTplSegs .ren-seg-wrap .ren-del'), '删除键只在自定义段输入框右上角（小 ×）');
  assert(doc.querySelector('#renSegAdd').textContent.includes('新增段'), '新增段按钮在「重命名规则」标题后');
  assert((doc.querySelector('.ren-rule-h .ren-tpl .itip')||{}).getAttribute?.('data-tip')?.includes?.('工号-类型'), '标题旁悬浮提示：工号-类型固定在最后');
  assert(!!doc.querySelector('#renTplPreview'), '完整规则预览存在');
  assert(doc.querySelector('#renExt').value === '', '后缀输入框默认留空（用原文件后缀）');
  assert(doc.querySelector('#renSep').value === '-', '连接符输入框默认 -');
  assert(doc.querySelector('.ren-rule-h .ren-tpl').textContent, '规则区标题存在');
  assert(!doc.querySelector('#renTplSegs').textContent.includes('{'), '分段模板无花括号');
  // 改连接符 → 段间连接符与预览同步
  const sepInput = doc.querySelector('#renSep');
  sepInput.value = '_';
  sepInput.dispatchEvent(new win.Event('input', { bubbles: true }));
  assert(doc.querySelectorAll('#renTplSegs .ren-seg-c').length === 2 &&
    doc.querySelectorAll('#renTplSegs .ren-seg-c')[0].textContent === '_', '连接符改为 _ 后段间连接符同步');
  // 清空连接符 → 回退 -
  sepInput.value = '';
  sepInput.dispatchEvent(new win.Event('input', { bubbles: true }));
  assert(doc.querySelectorAll('#renTplSegs .ren-seg-c')[0].textContent === '-', '连接符清空自动回退 -');
  // 加一段 → 固定文字段（自定义段都在工号前面）
  doc.querySelector('#renSegAdd').click();
  const segAdd = doc.querySelectorAll('#renTplSegs .ren-seg')[1];
  segAdd.querySelector('.ren-seg-sel').value = '__text';
  segAdd.querySelector('.ren-seg-sel').dispatchEvent(new win.Event('change', { bubbles: true }));
  assert(doc.querySelectorAll('#renTplSegs .ren-seg').length === 2 &&
    !!doc.querySelector('#renTplSegs .ren-seg-txt') &&
    doc.querySelectorAll('#renTplSegs .ren-seg-lock').length === 2, '加一段并切成固定文字，锁定尾段仍在最后');
  // 删除自定义固定文字段回到 1
  doc.querySelectorAll('#renTplSegs .ren-del')[1].click();
  assert(doc.querySelectorAll('#renTplSegs .ren-seg').length === 1 &&
    doc.querySelectorAll('#renTplSegs .ren-seg-lock').length === 2, '删除自定义段后剩 1 段 + 锁定尾段');
  assert(doc.querySelector('#btnRenProblems').disabled, '问题清单明细按钮初始禁用');
  const sw = doc.querySelector('#renMissingOn');
  assert(!!sw && !sw.checked, '缺失检查开关默认关闭');
  assert(doc.querySelector('#renChkList').classList.contains('hidden'), '开关关闭时勾选区隐藏');
  sw.click();
  assert(!doc.querySelector('#renChkList').classList.contains('hidden'), '开关打开后勾选区显示');
  assert(doc.querySelectorAll('[data-ren-chk]').length === 8 &&
    !doc.querySelector('[data-ren-chk="投保证明"]'), '缺失检查勾选区 8 项且不含投保证明（特殊模块）');

  // 2. 花名册
  dropInto('#renDrop-roster', path.join(__dirname, 'ren_roster.xlsx'));
  await sleep(400);
  assert(/识别到 4 名员工/.test(doc.querySelector('#renRosterBox').textContent), '花名册识别 4 名员工');
  assert(/^张三-R001-身份证正面\.jpg/.test(doc.querySelector('#renTplPreview').textContent), '规则预览用花名册第一人渲染成 张三-R001-身份证正面.jpg');
  // 字段映射：4 项自动猜 + 可改
  const maps = doc.querySelectorAll('[data-ren-map]');
  assert(maps.length === 4, '花名册字段映射 4 项（工号/姓名/蝶城/项目）');
  assert(maps[0].value === '2' && maps[1].value === '3' && maps[2].value === '0' && maps[3].value === '1',
    '映射自动猜对：工号=员工编码 姓名=员工姓名 蝶城=调度中心 项目=员工项目');
  maps[1].value = '0';
  maps[1].dispatchEvent(new win.Event('change'));
  assert(/^甬江蝶城-R001-身份证正面/.test(doc.querySelector('#renTplPreview').textContent),
    '改姓名映射后按新列重读（预览姓名变成调度中心列的值）');
  const maps2 = doc.querySelectorAll('[data-ren-map]');
  maps2[1].value = '3';
  maps2[1].dispatchEvent(new win.Event('change'));
  assert(/^张三-R001-身份证正面\.jpg/.test(doc.querySelector('#renTplPreview').textContent), '映射改回后恢复 张三');

  // 3. 资料模块 zip
  dropInto('#renGrid .drop[data-mod="身份证正面"]', path.join(__dirname, 'ren_身份证正面.zip'));
  await sleep(500);
  assert(/身份证正面 · 4 个文件/.test(doc.querySelector('#renGrid').textContent), '身份证正面 zip 解包 4 个文件');
  dropInto('#renGrid .drop[data-mod="劳务合同协议"]', path.join(__dirname, 'ren_劳务合同协议.zip'));
  dropInto('#renGrid .drop[data-mod="投保证明"]', path.join(__dirname, 'ren_投保证明.zip'));
  await sleep(500);
  assert(/劳务合同协议 · 2 个文件/.test(doc.querySelector('#renGrid').textContent), '劳务合同协议 zip 解包 2 个文件');

  // 3.5 投保证明特殊模块：上传 2 个文件 → 直接阻断
  doc.querySelector('#btnRenRun').click();
  await sleep(200);
  assert(/投保证明.*只需上传 1 个原始文件/.test(doc.querySelector('#toast').textContent), '投保证明 2 个文件触发阻断提示');
  assert(doc.querySelector('#btnRenProblems').disabled, '阻断后问题清单按钮回到禁用');
  // 清空投保证明窗口，改传 1 个原始文件
  doc.querySelector('[data-ren-clear-mod="投保证明"]').click();
  await sleep(100);
  dropInto('#renGrid .drop[data-mod="投保证明"]', path.join(__dirname, 'ren_投保证明1.zip'));
  await sleep(400);
  assert(/投保证明 · 1 个文件/.test(doc.querySelector('#renGrid').textContent), '投保证明换成 1 个原始文件');

  // 4. 勾选缺失检查：身份证正面（开关已在上一步打开）
  doc.querySelector('[data-ren-chk="身份证正面"]').click();
  assert(doc.querySelector('[data-ren-chk="身份证正面"]').checked, '身份证正面纳入缺失检查');

  // 5. 开始重命名（投保证明 1 个文件 → 全员复制 4 份）
  doc.querySelector('#btnRenRun').click();
  const ok = await waitFor(() => /重命名完成/.test(doc.querySelector('#toast').textContent), 20000);
  assert(ok, '重命名完成提示出现');
  assert(/成功 8 个文件/.test(doc.querySelector('#toast').textContent) &&
    /问题 4 条/.test(doc.querySelector('#toast').textContent), '完成提示：成功 8 / 问题 4（实际：' + doc.querySelector('#toast').textContent + '）');
  const sum = doc.querySelector('#renResults').textContent;
  assert(/成功重命名\s*8/.test(sum), '结果摘要 成功重命名 8');
  assert(/覆盖\s*4/.test(sum), '结果摘要 覆盖 4 人');
  assert(/问题\s*4/.test(sum), '结果摘要 问题 4 条');
  assert(/合计 \d+(\.\d+)?(KB|MB|B)/.test(sum), '结果提示行显示合计大小（实际：' + sum.slice(0, 120) + '）');
  assert(sum.includes('共 8 个文件'), '结果提示行显示共 8 个文件');
  assert(!doc.querySelector('#btnRenProblems').disabled, '问题清单明细按钮已启用');

  // 6. 问题清单弹窗
  doc.querySelector('#btnRenProblems').click();
  await sleep(100);
  const body = doc.querySelector('#renModalBody').textContent;
  assert(!doc.querySelector('#renModal').classList.contains('hidden'), '问题清单弹窗打开');
  assert(body.includes('周九.jpg'), '问题清单含 周九.jpg（未匹配）');
  assert(body.includes('未匹配到花名册员工'), '周九失败原因=未匹配到花名册员工');
  assert((body.match(/同名同姓无法区分/g) || []).length >= 1, '纯姓名张三.jpg 撞名进清单（无消歧段）');
  assert((body.match(/缺少「身份证正面」/g) || []).length === 2, '缺失检查 2 条：张三(R001)+王五(R004) 缺正面');
  assert(body.includes('张三-R003-身份证正面.jpg') || body.includes('R003'), '消歧结果落到 R003（张三-项目乙）');

  // 7. 导出 zip（平铺，投保证明全员复制在内）
  const nDl0 = downloads.length;
  doc.querySelector('#btnRenZip').click();
  await waitFor(() => downloads.length > nDl0, 15000);
  const zipDl = downloads.find(d => /^重命名结果-/.test(d.name));
  assert(!!zipDl, '导出 zip 文件名=重命名结果-时间戳.zip（实际 ' + (zipDl && zipDl.name) + '）');
  if (zipDl) {
    const u8 = new Uint8Array(await zipDl.blob.arrayBuffer());
    assert(zipContains(u8, '李四-R002-身份证正面.jpg'), 'zip 内含 李四-R002-身份证正面.jpg');
    assert(zipContains(u8, '张三-R003-身份证正面.jpg'), 'zip 内含 张三-R003-身份证正面.jpg（项目消歧）');
    assert(zipContains(u8, '张三-R001-劳务合同协议.pdf'), 'zip 内含 张三-R001-劳务合同协议.pdf（调度中心消歧）');
    assert(zipContains(u8, '张三-R001-投保证明.jpg'), '投保证明复制：张三-R001-投保证明.jpg');
    assert(zipContains(u8, '李四-R002-投保证明.jpg'), '投保证明复制：李四-R002-投保证明.jpg');
    assert(zipContains(u8, '王五-R004-投保证明.jpg'), '投保证明复制：王五-R004-投保证明.jpg');
    assert(!zipContains(u8, '周九'), 'zip 内不含周九（未匹配）');
    assert(!zipContains(u8, '张三-R001-身份证正面.jpg'), 'zip 内不含 R001 正面（纯姓名撞名失败）');
  }

  // 8. 后缀强制覆盖：填 pdf → 全部文件 .pdf 结尾
  const extInput = doc.querySelector('#renExt');
  extInput.value = 'pdf';
  extInput.dispatchEvent(new win.Event('input', { bubbles: true }));
  doc.querySelector('#btnRenRun').click();
  await waitFor(() => /重命名完成/.test(doc.querySelector('#toast').textContent), 20000);
  const nDl2 = downloads.length;
  doc.querySelector('#btnRenZip').click();
  await waitFor(() => downloads.length > nDl2, 15000);
  const zip2 = downloads.slice(nDl2).find(d => /^重命名结果-/.test(d.name));
  if (zip2) {
    const u82 = new Uint8Array(await zip2.blob.arrayBuffer());
    assert(zipContains(u82, '李四-R002-身份证正面.pdf'), '后缀覆盖：李四-R002-身份证正面.pdf');
    assert(!zipContains(u82, '李四-R002-身份证正面.jpg'), '后缀覆盖后不再有 .jpg');
  } else assert(false, '后缀覆盖导出 zip 下载成功');

  // 9. 自动拆包：每包限 2 个文件 → 8 个成功文件拆多包
  win.RenApp._test.setLimits(1024 * 1024, 2);
  const nDl3 = downloads.length;
  doc.querySelector('#btnRenZip').click();
  await sleep(1200);
  const pack1 = downloads.slice(nDl3).find(d => /^重命名结果-1-/.test(d.name));
  const pack2 = downloads.slice(nDl3).find(d => /^重命名结果-2-/.test(d.name));
  assert(!!pack1 && !!pack2, '自动拆包成多个 zip（-1 / -2 均存在）');
  assert(/已拆分成 \d+ 个压缩包/.test(doc.querySelector('#toast').textContent) && /已导出 \d+ 个文件/.test(doc.querySelector('#toast').textContent), '完成 toast 汇总：已导出 N 个文件（大小 · 拆分包数）（实际：' + doc.querySelector('#toast').textContent + '）');
  assert(doc.querySelector('#btnRenZip').textContent === '导出 zip' && !doc.querySelector('#btnRenZip').disabled, '打包完成后按钮恢复「导出 zip」');
  win.RenApp._test.defaults();

  // 10. 问题清单 xlsx
  extInput.value = '';
  extInput.dispatchEvent(new win.Event('input', { bubbles: true }));
  doc.querySelector('#btnRenProbXlsx').click();
  await waitFor(() => downloads.some(d => /^重命名问题清单-/.test(d.name)), 15000);
  const xlsx = downloads.find(d => /^重命名问题清单-/.test(d.name));
  assert(!!xlsx, '问题清单 xlsx 下载成功（实际 ' + (xlsx && xlsx.name) + '）');

  // 9. 关闭弹窗
  doc.querySelector('#renModalClose').click();
  assert(doc.querySelector('#renModal').classList.contains('hidden'), '弹窗可关闭');

  console.log(failed ? ('\nFAILED: ' + failed) : '\nALL PASS');
  process.exit(failed ? 1 : 0);
})().catch(e => { console.error('FATAL', e); process.exit(1); });
