/* ============================================================
 * idle-app.js — 出勤异常核对界面（超7天未出勤）
 * 三个上传区：① 考勤统计结果导出 ② 人员花名册 ③ 上期反馈结果
 * 核对设置（核算周期 / 口径门槛 / 排除名单）→ 开始核对 → 结果看板
 * 结果：KPI + 已剔除备查 + 出勤异常 / 无考勤记录 双 Tab（蝶城→项目分组）
 * 导出：超7天无出勤异常明细-时间戳.xlsx（异常主表 + 在职无出勤子表）
 * ============================================================ */
(function () {
  'use strict';

  var XLSX = window.XLSX;
  var I = window.IdleCore;
  var W = window.XlsxWriter;
  var Z = window.ZipUtil;

  var state = {
    att: null, emp: null, prev: null,       // 各区解析结果
    prevIds: [], prevById: {}, prevByName: {},
    ran: false,
    cfg: { mode: 'gap', thr: 7, scopeOn: true, leaveOut: true, showNoRec: true, excl: '机师岗' },
    res: null,
    ui: { tab: 'main', city: '', kw: '', kw2: '' }
  };

  function $(sel, ctx) { return (ctx || document).querySelector(sel); }
  function $$(sel, ctx) { return Array.prototype.slice.call((ctx || document).querySelectorAll(sel)); }
  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  var toastTimer = null;
  function toast(msg, type) {
    var t = $('#toast');
    t.textContent = msg;
    t.className = 'toast show' + (type ? ' ' + type : '');
    clearTimeout(toastTimer);
    setTimeout(function () { t.className = 'toast'; }, 2600);
  }
  function downloadBlob(blob, name) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 800);
  }

  /* ---------------- 上传区 ---------------- */
  var ZONES = [
    { key: 'att', label: '① 考勤统计结果导出', sub: '含 SHRid / 计薪天数 / 应服务工时', req: true },
    { key: 'emp', label: '② 人员花名册 / 员工信息维护', sub: '含 员工编码 / 入司日期 / 离职日期 / 职位', req: false },
    { key: 'prev', label: '③ 上期反馈结果（可选）', sub: '本工具上期导出的结果表，用于扣请假/剔除已报离', req: false }
  ];

  function zoneHtml(z) {
    var d = state[z.key];
    var info;
    if (!d) {
      info = '';
    } else if (d.error) {
      info = '<div class="file-item"><div class="fi-head"><div class="fi-name"><span class="fi-dot" style="background:var(--red)"></span>' +
        '<span class="fi-nm">' + esc(d.fileName) + '</span></div>' +
        '<button class="btn-ghost" data-idle-clear="' + z.key + '">移除</button></div>' +
        '<div class="att-file-meta" style="color:var(--red)">⚠ ' + esc(d.error) + '</div></div>';
    } else {
      var txt = z.key === 'att' ? ('识别到 ' + d.rows.length + ' 条考勤记录 · 工作表「' + esc(d.sheet) + '」') :
        z.key === 'emp' ? ('识别到 ' + d.rows.length + ' 名员工 · 工作表「' + esc(d.sheet) + '」') :
        ('识别到 ' + (d.count || 0) + ' 条上期反馈 · 工作表「' + esc(d.sheet) + '」');
      info = '<div class="file-item"><div class="fi-head"><div class="fi-name"><span class="fi-dot" style="background:var(--green)"></span>' +
        '<span class="fi-nm">' + esc(d.fileName) + '</span></div>' +
        '<button class="btn-ghost" data-idle-clear="' + z.key + '">移除</button></div>' +
        '<div class="att-file-meta">' + txt + '</div></div>';
    }
    return '<div class="att-zone"><div class="drop" id="idleDrop-' + z.key + '">' +
      '<div class="d-main">' + z.label + '</div>' +
      '<div class="d-sub" title="' + esc(z.sub) + '">' + esc(z.sub) + '</div>' +
      '<div class="d-sub">' + (z.req ? '必填' : '选填') + ' · 拖拽或点击选择</div>' +
      '<input type="file" id="idlePick-' + z.key + '" accept=".xlsx,.xls" class="hidden"></div>' +
      (info ? '<div class="file-list">' + info + '</div>' : '') + '</div>';
  }

  function renderZones() {
    var box = $('#idleZones');
    if (!box) return;
    box.innerHTML = ZONES.map(zoneHtml).join('');
    ZONES.forEach(function (z) {
      var drop = $('#idleDrop-' + z.key), picker = $('#idlePick-' + z.key);
      if (!drop) return;
      drop.addEventListener('click', function (e) { if (!e.target.closest('input')) picker.click(); });
      picker.addEventListener('change', function () {
        if (picker.files.length) handleFile(z.key, picker.files[0]);
        picker.value = '';
      });
      drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('drag'); });
      drop.addEventListener('dragleave', function () { drop.classList.remove('drag'); });
      drop.addEventListener('drop', function (e) {
        e.preventDefault(); drop.classList.remove('drag');
        if (e.dataTransfer.files.length) handleFile(z.key, e.dataTransfer.files[0]);
      });
    });
    $$('[data-idle-clear]').forEach(function (b) {
      b.addEventListener('click', function () { clearZone(b.dataset.idleClear); });
    });
  }

  function handleFile(key, file) {
    if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) { toast('请选择 .xlsx / .xls 文件', 'warn'); return; }
    var reader = new FileReader();
    reader.onload = function () {
      var wb;
      try { wb = XLSX.read(new Uint8Array(reader.result), { type: 'array' }); }
      catch (e) { toast('解析失败：' + file.name, 'warn'); return; }
      var parsed = I.parseFile(wb, key, file.name);
      if (parsed.error) {
        state[key] = { fileName: file.name, error: parsed.error };
        toast(file.name + '：' + parsed.error, 'warn');
      } else {
        state[key] = parsed;
        if (key === 'prev') { collectPrev(parsed); }
        if (parsed.warnings && parsed.warnings.length) toast(file.name + '：' + parsed.warnings[0], 'warn');
        state.ran = false;
      }
      renderZones();
      renderResults();
    };
    reader.readAsArrayBuffer(file);
  }

  function collectPrev(p) {
    (p.ids || []).forEach(function (v) {
      if (state.prevIds.indexOf(v) < 0) state.prevIds.push(v);
      if (!state.prevById[v]) state.prevById[v] = p.byId[v];
    });
    Object.keys(p.byId || {}).forEach(function (k) { if (!state.prevById[k]) state.prevById[k] = p.byId[k]; });
    Object.keys(p.byName || {}).forEach(function (k) { if (!state.prevByName[k]) state.prevByName[k] = p.byName[k]; });
  }

  function clearZone(key) {
    state[key] = null;
    if (key === 'prev') { state.prevIds = []; state.prevById = {}; state.prevByName = {}; }
    state.ran = false;
    renderZones();
    renderResults();
  }

  /* ---------------- 核对 ---------------- */
  function readCfg() {
    var c = state.cfg;
    c.mode = ($('#idleMode') || {}).value || 'gap';
    var t = parseFloat(($('#idleThr') || {}).value);
    c.thr = isFinite(t) && t > 0 ? t : 7;
    c.scopeOn = !$('#idleChkScope') || $('#idleChkScope').checked;
    c.leaveOut = !$('#idleChkLeave') || $('#idleChkLeave').checked;
    c.showNoRec = !$('#idleChkNoRec') || $('#idleChkNoRec').checked;
    c.excl = (($('#idleExcl') || {}).value || '').trim();
    c.periodStart = ($('#idlePStart') || {}).value ? I.parseDate($('#idlePStart').value) : null;
    c.periodEnd = ($('#idlePEnd') || {}).value ? I.parseDate($('#idlePEnd').value) : null;
    return c;
  }

  function doRun() {
    if (!state.att || state.att.error || !state.att.rows || !state.att.rows.length) {
      toast('先导入「① 考勤统计结果导出」', 'warn');
      return;
    }
    var cfg = readCfg();
    state.res = I.compute({
      att: state.att.rows,
      emp: (state.emp && state.emp.rows) || [],
      prevIds: state.prevIds, prevById: state.prevById, prevByName: state.prevByName,
      thr: cfg.thr, mode: cfg.mode, scopeOn: cfg.scopeOn, leaveOut: cfg.leaveOut,
      showNoRec: cfg.showNoRec, excl: cfg.excl,
      periodStart: cfg.periodStart, periodEnd: cfg.periodEnd
    });
    state.ran = true;
    syncPeriodInputs();
    renderResults();
    toast('核对完成：异常 ' + state.res.anomalies.length + ' 人，无考勤记录 ' + state.res.noRecord.length + ' 人', 'ok');
  }

  function syncPeriodInputs() {
    var r = state.res;
    if (!r) return;
    var ps = $('#idlePStart'), pe = $('#idlePEnd');
    if (ps && !ps.value) ps.value = I.fmtDate(r.periodStart);
    if (pe && !pe.value) pe.value = I.fmtDate(r.periodEnd);
  }

  /* ---------------- 渲染 ---------------- */
  function kpiHtml(hl, label, value, unit) {
    return '<div class="idle-kpi' + (hl ? ' hl' : '') + '"><div class="ik-lb">' + label +
      '</div><div class="ik-vv">' + value + '<small>' + unit + '</small></div></div>';
  }

  function renderResults() {
    var box = $('#idleResults');
    if (!box) return;
    var r = state.res;
    if (!r || !state.ran) {
      var ready = state.att && !state.att.error && state.att.rows && state.att.rows.length;
      box.innerHTML = ready ? '' : '';
      return;
    }
    var fmtD = I.fmtDay;
    var h = '';

    /* 周期提示 */
    var warn = '';
    if (r.maxPay > 0 && r.spanDays > r.maxPay + 1) {
      warn = '<div class="idle-warnbar">这份考勤数据的<b>计薪天数最高只有 ' + fmtD(r.maxPay) + ' 天</b>，说明它只统计到 <b>' +
        I.fmtDate(r.dataEnd) + '</b>；你现在选的周期有 <b>' + r.spanDays + ' 天</b>，按这个算每人至少缺 ' +
        fmtD(r.spanDays - r.maxPay) + ' 天，<b>会全员上榜</b>。点「按数据填」改成 ' + I.fmtDate(r.periodStart) + ' ~ ' +
        I.fmtDate(r.dataEnd) + '，或重新导出一份覆盖完整周期的数据。</div>';
    }
    h += warn;
    h += '<div class="idle-scopebar">本期 <b>' + I.fmtDate(r.periodStart) + ' ~ ' + I.fmtDate(r.periodEnd) + '</b>　共 <b>' + r.spanDays +
      ' 天</b>（每天计为应出勤，不扣周末/法定假）　·　数据计薪天数最高 ' + fmtD(r.maxPay) + ' 天 → 实际统计到 <b>' + I.fmtDate(r.dataEnd) + '</b>' +
      (r.scope && r.scope.cities.length ? '　·　导入范围：<b>' + esc(r.scope.cities.length === 1 ? r.scope.cities[0] : r.scope.cities.length + ' 个蝶城') + '</b> / ' + r.scope.projCount + ' 个项目' : '') + '</div>';

    /* KPI */
    var allGap = r.anomalies.reduce(function (s, x) { return s + x.gap; }, 0);
    h += '<div class="idle-kpis">' +
      kpiHtml(true, '本期异常人数', r.anomalies.length, '人') +
      kpiHtml(false, '合计无出勤天数', fmtD(allGap), '天') +
      kpiHtml(false, '导入人数（去重）', r.peopleCount, '人') +
      kpiHtml(false, '无考勤记录', r.noRecord.length, '人') +
      '</div>';

    /* 已剔除备查 */
    var leftRows = r.leftOut.concat(r.reportedOut);
    if (leftRows.length) {
      h += '<details class="result-card" id="idleLeftCard"><summary class="idle-leftsum">已剔除 <b>' + leftRows.length +
        '</b> 人（本期离职 / 上期已报离，点开看名单）</summary><div class="idle-lefthint">共 <b>' + leftRows.length + '</b> 人，原本会算出 <b>' +
        fmtD(leftRows.reduce(function (s, x) { return s + x.gap; }, 0)) + '</b> 天无出勤，已从明细和导出里剔除，只在这里备查：' +
        '· <b>本期离职 ' + r.leftOut.length + ' 人</b>：离职日期落在本期，离职后不再要求出勤，算出来「无出勤 = 整段在职天数」是假异常；' +
        '· <b>上期已报离 / 已确认离职 ' + r.reportedOut.length + ' 人</b>：上期结果表里已核实，本期不再重复上榜。</div>' +
        '<div class="idle-twrap" style="max-height:320px"><table class="att-table"><thead><tr>' +
        '<th>蝶城</th><th>项目</th><th>员工编码</th><th>姓名</th><th>入职时间</th><th>离职时间</th><th>剔除原因</th><th>原算出无出勤</th>' +
        '</tr></thead><tbody>' +
        leftRows.map(function (x) {
          return '<tr><td>' + esc(x.city) + '</td><td>' + esc(x.project) + '</td><td>' + esc(x.id) + '</td><td>' + esc(x.name) + '</td>' +
            '<td>' + esc(x.joinTxt || '—') + '</td><td>' + esc(x.endTxt || '—') + '</td><td>' + esc(x.why || '') + '</td>' +
            '<td style="font-weight:600">' + fmtD(x.gap) + '</td></tr>';
        }).join('') + '</tbody></table></div></details>';
    }

    /* 结果卡：双 Tab */
    h += '<div class="result-card"><div class="idle-tabs" id="idleTabs">' +
      tabHtml('main', '出勤异常', r.anomalies.length) +
      (r.noRecord.length ? tabHtml('noatt', '无考勤记录', r.noRecord.length) : '') +
      '</div>';

    /* 主表工具栏 */
    h += '<div class="idle-filters" id="idleFiltersMain">' +
      '<input id="idleKw" type="text" placeholder="搜索姓名 / 工号 / 项目">' +
      '<select id="idleCity"><option value="">全部蝶城</option></select>' +
      '<div class="idle-cnt">共 <b id="idleCnt">0</b> 人</div>' +
      '<button class="btn btn-line" id="btnIdleExport" style="margin-left:auto">导出 Excel</button>' +
      '</div>';
    h += '<div class="idle-twrap"><table class="att-table" id="idleTbMain"><thead><tr>' +
      '<th>蝶城</th><th>项目</th><th>员工编码</th><th>姓名</th><th>入职时间</th><th>离职时间</th><th>无出勤天数</th><th>已扣请假</th>' +
      '</tr></thead><tbody id="idleTb"></tbody></table></div>';
    h += '<div class="idle-tfoot"><div id="idleFootL"></div><div id="idleFootR"></div></div>';

    /* 无考勤记录表 */
    if (r.noRecord.length) {
      h += '<div id="idleTabNoatt" style="display:none"><div class="idle-scopebar" style="margin-top:0">这 ' + r.noRecord.length +
        ' 人来自<b>「人员花名册」</b>（按 员工编码 = SHRid 比对），<b>考勤统计表里没有他们的工号</b>，无出勤天数按<b>本期在职的日历天数估算</b>。常见原因：劳务外包不在考勤采集范围、已离职被清档、长期请假、考勤导出漏了这一批人。</div>' +
        '<div class="idle-twrap"><table class="att-table"><thead><tr>' +
        '<th>蝶城</th><th>项目</th><th>员工编码</th><th>姓名</th><th>入职时间</th><th>离职时间</th><th>无出勤天数（估算）</th><th>已扣请假</th>' +
        '</tr></thead><tbody id="idleTbNo"></tbody></table></div></div>';
    }
    h += '</div>';

    box.innerHTML = h;
    renderMainTable();
    renderNoTable();
    bindResults();
  }

  function tabHtml(id, lb, n) {
    return '<button class="idle-tab' + (state.ui.tab === id ? ' on' : '') + '" data-tab="' + id + '">' + lb + ' <b>' + n + '</b></button>';
  }

  function gapClass(g) {
    if (g >= 7) return 'idle-gap';
    if (g >= 3) return 'idle-gap mid';
    return 'idle-gap sm';
  }

  function filteredMain() {
    var r = state.res;
    var kw = state.ui.kw, city = state.ui.city;
    return r.anomalies.filter(function (x) {
      if (city && x.city !== city) return false;
      if (!kw) return true;
      return (x.name + ' ' + x.id + ' ' + x.project + ' ' + x.city).toLowerCase().indexOf(kw) >= 0;
    });
  }

  /* 蝶城 → 项目 分组表 */
  function groupRows(rows, est) {
    var html = [];
    var i = 0;
    while (i < rows.length) {
      var city = rows[i].city, j = i;
      while (j < rows.length && rows[j].city === city) j++;
      var cityRows = rows.slice(i, j);
      var citySum = cityRows.reduce(function (s, x) { return s + x.gap; }, 0);
      html.push('<tr class="idle-g-city"><td colspan="8">' + esc(city) +
        '<span class="idle-meta">' + cityRows.length + ' 人 · 合计 ' + I.fmtDay(citySum) + ' 天</span></td></tr>');
      var k = 0;
      while (k < cityRows.length) {
        var proj = cityRows[k].project, m = k;
        while (m < cityRows.length && cityRows[m].project === proj) m++;
        var pros = cityRows.slice(k, m);
        var pSum = pros.reduce(function (s, x) { return s + x.gap; }, 0);
        html.push('<tr class="idle-g-proj"><td></td><td colspan="7">' + esc(proj) +
          '<span class="idle-meta">' + pros.length + ' 人 · ' + I.fmtDay(pSum) + ' 天</span></td></tr>');
        pros.forEach(function (x) {
          html.push('<tr>' +
            '<td>' + esc(x.city) + '</td>' +
            '<td>' + esc(x.project) + '</td>' +
            '<td>' + esc(x.id) + '</td>' +
            '<td>' + esc(x.name) + '</td>' +
            '<td>' + esc(x.joinTxt || '—') + '</td>' +
            '<td>' + esc(x.endTxt || '—') + '</td>' +
            '<td class="num"><span class="' + gapClass(x.gap) + '"' + (x.est ? ' title="考勤统计表里没有这个工号，无出勤天数按本期在职日历天数估算"' : '') + '>' + I.fmtDay(x.gap) + '</span></td>' +
            '<td class="num">' + (x.deducted > 0 ? '<span class="idle-ded">-' + I.fmtDay(x.deducted) + '</span>' : '—') + '</td>' +
            '</tr>');
        });
        k = m;
      }
      i = j;
    }
    return html.join('');
  }

  function renderMainTable() {
    var r = state.res;
    if (!r) return;
    var tb = $('#idleTb');
    if (!tb) return;
    var rows = filteredMain();
    $('#idleCnt').textContent = rows.length;
    tb.innerHTML = rows.length ? groupRows(rows) : '<tr><td colspan="8"><div class="att-empty">没有符合条件的异常人员</div></td></tr>';
    var totalGap = rows.reduce(function (s, x) { return s + x.gap; }, 0);
    $('#idleFootL').textContent = (rows.length === r.anomalies.length
      ? ('共 ' + rows.length + ' 人，合计无出勤 ' + I.fmtDay(totalGap) + ' 天')
      : ('筛选后 ' + rows.length + ' / ' + r.anomalies.length + ' 人，合计 ' + I.fmtDay(totalGap) + ' 天'));
    var parts = [state.cfg.mode === 'low' ? ('筛选：实出勤 < ' + state.cfg.thr + ' 天') : ('筛选：无出勤 ≥ ' + state.cfg.thr + ' 天')];
    if (state.emp && state.emp.rows) parts.push('已按花名册校验在职月份（排除 ' + r.stats.filteredOut + ' 人）');
    else parts.push('未导入花名册，未做入职/离职校验');
    if (state.prev && !state.prev.error) parts.push('已按上期反馈扣减 ' + r.stats.deducted + ' 人、剔除已报离 ' + r.stats.reportedOut + ' 人');
    if (state.emp && state.cfg.leaveOut) parts.push('已剔除本期离职 ' + r.leftOut.length + ' 人');
    if (r.stats.exclN) parts.push('排除名单命中 ' + r.stats.exclN + ' 人');
    $('#idleFootR').textContent = parts.join(' · ');
  }

  function renderNoTable() {
    var r = state.res;
    if (!r || !r.noRecord.length) return;
    var tb = $('#idleTbNo');
    if (!tb) return;
    tb.innerHTML = groupRows(r.noRecord, true) || '<tr><td colspan="8"><div class="att-empty">没有无考勤记录的人</div></td></tr>';
  }

  function bindResults() {
    var r = state.res;
    /* Tab */
    $$('#idleTabs .idle-tab').forEach(function (t) {
      t.addEventListener('click', function () {
        state.ui.tab = t.dataset.tab;
        $$('#idleTabs .idle-tab').forEach(function (x) { x.classList.toggle('on', x.dataset.tab === state.ui.tab); });
        var showMain = state.ui.tab === 'main';
        $('#idleFiltersMain').style.display = showMain ? '' : 'none';
        $('#idleTbMain').parentElement.style.display = showMain ? '' : 'none';
        var noatt = $('#idleTabNoatt');
        if (noatt) noatt.style.display = showMain ? 'none' : '';
        // 页脚跟主表走
        $('#idleFootL').parentElement.style.display = showMain ? '' : 'none';
      });
    });
    /* 筛选 */
    var kw = $('#idleKw'), city = $('#idleCity');
    if (kw) {
      kw.value = state.ui.kw;
      kw.addEventListener('input', function () { state.ui.kw = kw.value.trim().toLowerCase(); renderMainTable(); });
    }
    if (city) {
      var keep = state.ui.city;
      var cities = [];
      r.anomalies.forEach(function (x) { if (cities.indexOf(x.city) < 0) cities.push(x.city); });
      cities.sort(function (a, b) { return String(a).localeCompare(String(b), 'zh-Hans-CN'); });
      city.innerHTML = '<option value="">全部蝶城</option>' + cities.map(function (c) { return '<option>' + esc(c) + '</option>'; }).join('');
      city.value = keep;
      if (city.value !== keep) state.ui.city = keep = '';
      city.addEventListener('change', function () { state.ui.city = city.value; renderMainTable(); });
    }
    /* 导出 */
    var ex = $('#btnIdleExport');
    if (ex) ex.addEventListener('click', doExport);
  }

  /* ---------------- 导出 ---------------- */
  var COLS = [
    { key: 'city', header: '蝶城', width: 18 },
    { key: 'project', header: '项目', width: 30 },
    { key: 'id', header: '员工编码', width: 18 },
    { key: 'name', header: '姓名', width: 12 },
    { key: 'joinTxt', header: '入职时间', width: 14 },
    { key: 'endTxt', header: '离职时间', width: 14 },
    { key: 'gap', header: '无出勤天数', width: 13, num: true },
    { key: 'reason', header: '异常原因', width: 12, blank: true },
    { key: 'leaveFrom', header: '请假开始时间', width: 15, blank: true },
    { key: 'leaveTo', header: '请假结束时间', width: 15, blank: true },
    { key: 'leaveDays', header: '请假天数', width: 11, num: true, blank: true },
    { key: 'reported', header: '是否已报离', width: 12, blank: true },
    { key: 'note', header: '备注', width: 24, blank: true },
    { key: 'deducted', header: '已扣请假', width: 11, num: true }
  ];

  function toRow(x) {
    return [
      x.city, x.project, x.id, x.name, x.joinTxt || '', x.endTxt || '',
      x.gap,
      x.reason || '', x.leaveFrom || '', x.leaveTo || '',
      x.leaveDays > 0 ? x.leaveDays : '',
      x.reported || '', x.note || '',
      x.deducted > 0 ? x.deducted : ''
    ];
  }

  function sheetStyleFn(rows) {
    // 反馈填写区（7~12 列）空单元格浅黄底（样式9=FFFFCC 加粗）；无出勤天数标红（样式8）
    return function (r, cc) {
      if (cc === 6) return rows[r][6] != null && rows[r][6] !== '' ? 8 : 2;
      if (cc >= 7 && cc <= 12) {
        var v = rows[r][cc];
        if (v == null || v === '') return 9;
        return cc === 10 ? 3 : 2;
      }
      if (cc === 13) return rows[r][13] !== '' ? 3 : 2;
      return 2;
    };
  }

  function doExport() {
    var r = state.res;
    if (!r || !state.ran) return;
    var mainRows = r.anomalies.slice();
    var noRows = state.cfg.showNoRec ? r.noRecord.slice() : [];
    if (!mainRows.length && !noRows.length) { toast('没有可导出的内容', 'warn'); return; }
    var sheets = [];
    if (mainRows.length) {
      var rowsMain = mainRows.map(toRow);
      sheets.push({
        name: '出勤异常', header: COLS.map(function (c) { return c.header; }), rows: rowsMain,
        headerStyle: 7, styleFn: sheetStyleFn(rowsMain), freeze: { x: 0, y: 1 }, autoFilter: true,
        rowHeight: 24, cols: COLS.map(function (c) { return c.width; })
      });
    }
    if (noRows.length) {
      var rowsNo = noRows.map(toRow);
      sheets.push({
        name: '在职无出勤记录', header: COLS.map(function (c) { return c.header; }), rows: rowsNo,
        headerStyle: 7, styleFn: sheetStyleFn(rowsNo), freeze: { x: 0, y: 1 }, autoFilter: true,
        rowHeight: 24, cols: COLS.map(function (c) { return c.width; })
      });
    }
    W.build({ zipImpl: Z, sheets: sheets }).then(function (bytes) {
      var now = new Date();
      function p2(n) { return n < 10 ? '0' + n : '' + n; }
      var stamp = '' + now.getFullYear() + p2(now.getMonth() + 1) + p2(now.getDate()) + p2(now.getHours()) + p2(now.getMinutes());
      downloadBlob(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
        '超7天无出勤异常明细-' + stamp + '.xlsx');
      toast('已导出（异常 ' + mainRows.length + ' 条' + (noRows.length ? '，子表「在职无出勤记录」' + noRows.length + ' 条' : '') + '）', 'ok');
    });
  }

  /* ---------------- 事件绑定 ---------------- */
  function bind() {
    if (!$('#viewAbsent')) return;
    renderZones();
    var run = $('#btnIdleRun');
    if (run) run.addEventListener('click', doRun);
    /* 上榜口径分段按钮：高亮当前选中项 */
    $$('#idleModeSeg button').forEach(function (b) {
      b.addEventListener('click', function () {
        $$('#idleModeSeg button').forEach(function (x) { x.classList.remove('on'); });
        b.classList.add('on');
        var h = $('#idleMode');
        if (h) h.value = b.dataset.mode;
        state.ran = false;
      });
    });
    /* 上期已扣请假的行更新提示等：排除名单输入即时重算 */
    var excl = $('#idleExcl');
    if (excl) {
      try { var saved = localStorage.getItem('idleExcl'); if (saved != null) { excl.value = saved; state.cfg.excl = saved; } } catch (e) { }
      excl.addEventListener('input', function () {
        state.cfg.excl = excl.value.trim();
        try { localStorage.setItem('idleExcl', excl.value); } catch (e) { }
        if (state.ran) { state.ran = false; }
      });
    }
    var de = $('#btnIdleDataEnd');
    if (de) de.addEventListener('click', function () {
      if (!state.att || state.att.error || !state.att.rows) { toast('先导入「① 考勤统计结果导出」', 'warn'); return; }
      // 反推：计薪天数最高值 → 数据截止日
      var maxPay = 0;
      state.att.rows.forEach(function (x) { if (x.pay > maxPay) maxPay = x.pay; });
      var mb2 = I.monthBounds((state.att.rows.filter(function (x) { return x.month; }).pop() || {}).month || '');
      var ps = $('#idlePStart'), pe = $('#idlePEnd');
      if (ps) ps.value = I.fmtDate(mb2.start);
      if (pe) {
        var d = new Date(mb2.start.getTime());
        d.setDate(d.getDate() + Math.ceil(maxPay) - 1);
        var last = new Date(mb2.end.getFullYear(), mb2.end.getMonth(), mb2.end.getDate());
        if (d > last) d = last;
        pe.value = I.fmtDate(d);
      }
      state.ran = false;
      toast('已按数据填周期，请重新点「开始核对」');
    });
  }

  document.addEventListener('DOMContentLoaded', bind);
  window.IdleApp = {
    onShow: function () {
      renderZones();
      renderResults();
    }
  };
})();
