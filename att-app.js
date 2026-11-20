/* ============================================================
 * att-app.js — 考勤处理模块界面
 * 三个上传区（外包SHR / RM / 花名册）+ 花名册映射 + 机师岗勾选 + 生成与下载
 * ============================================================ */
(function () {
  'use strict';

  var XLSX = window.XLSX;
  var A = window.AttCore;
  var W = window.XlsxWriter;
  var Z = window.ZipUtil;
  var CFG_KEY = 'salaryTool.att.v1';
  var FILE_ID = 0;

  var state = {
    shr: [],        // [{id, name, buffer, parsed, sheetName}]
    rm: [],
    roster: null,   // {id, name, buffer, wb, rows, headers}
    cfg: null,
    results: null,  // {files:[{name,blob,city}], cities, summary, problems}
    machineCandidates: [],  // [{name, checked}]
    rosterGuessed: false
  };

  /* ---------------- 小工具 ---------------- */
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
    toastTimer = setTimeout(function () { t.className = 'toast'; }, 2600);
  }

  function deepMerge(base, extra) {
    var out = JSON.parse(JSON.stringify(base));
    if (!extra) return out;
    Object.keys(extra).forEach(function (k) {
      if (extra[k] && typeof extra[k] === 'object' && !Array.isArray(extra[k]) &&
          out[k] && typeof out[k] === 'object' && !Array.isArray(out[k])) out[k] = deepMerge(out[k], extra[k]);
      else out[k] = extra[k];
    });
    return out;
  }
  function loadCfg() {
    try { return deepMerge(A.DEFAULT_CONFIG, JSON.parse(localStorage.getItem(CFG_KEY) || '{}')); }
    catch (e) { return JSON.parse(JSON.stringify(A.DEFAULT_CONFIG)); }
  }
  var saveTimer = null;
  function saveCfg() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      try { localStorage.setItem(CFG_KEY, JSON.stringify(state.cfg)); } catch (e) { /* 忽略 */ }
    }, 200);
  }

  function downloadBlob(blob, name) {
    if (window.navigator && window.navigator.msSaveOrOpenBlob) { window.navigator.msSaveOrOpenBlob(blob, name); return; }
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 800);
  }

  /* ---------------- 文件接收 ---------------- */

  function addFiles(zone, fileList) {
    var files = Array.prototype.slice.call(fileList).filter(function (f) { return /\.(xlsx|xlsm|xls)$/i.test(f.name); });
    if (!files.length) { toast('请选择 .xlsx / .xlsm / .xls 文件', 'warn'); return; }
    var pending = files.length;
    files.forEach(function (f) {
      var reader = new FileReader();
      reader.onload = function () {
        var buf = reader.result;
        if (zone === 'roster') {
          var wb;
          try { wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true }); }
          catch (e) { toast('花名册解析失败：' + f.name, 'warn'); done(); return; }
          state.roster = { id: ++FILE_ID, name: f.name, buffer: buf, wb: wb, rows: [], headers: [] };
          refreshRoster();
        } else {
          var parsed;
          try {
            var w = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true });
            parsed = A.parseAttFile(w, f.name, zone);
          } catch (e) { parsed = { fileName: f.name, type: zone, error: '解析失败：' + e.message, headers: [], rows: [] }; }
          state[zone].push({ id: ++FILE_ID, name: f.name, buffer: buf, parsed: parsed });
          refreshMachine();
        }
        done();
      };
      reader.readAsArrayBuffer(f);
    });
    function done() { if (--pending === 0) renderAll(); }
  }

  function removeItem(zone, id) {
    if (zone === 'roster') { state.roster = null; }
    else state[zone] = state[zone].filter(function (x) { return x.id !== id; });
    refreshMachine();
    renderAll();
  }

  /* ---------------- 花名册映射 ---------------- */

  function rosterHeaders() {
    if (!state.roster || !state.roster.wb) return [];
    var ws = state.roster.wb.Sheets[state.roster.wb.SheetNames[0]];
    if (!ws || !ws['!ref']) return [];
    var range = ws['!ref'];
    var colN = range.match(/:([A-Z]+)/)[1], n = 0;
    for (var i = 0; i < colN.length; i++) n = n * 26 + colN.charCodeAt(i) - 64;
    var endRowM = range.match(/(\d+):[A-Z]+(\d+)$/);
    var endRow = endRowM ? Number(endRowM[2]) : 1;
    var out = [];
    for (var r = 1; r <= Math.min(6, endRow); r++) {
      var vals = [];
      for (var c = 1; c <= n; c++) {
        var addr = String.fromCharCode(64 + Math.min(c, 26)) + r;
        var cell = ws[addr];
        vals.push(cell ? String(cell.v != null ? cell.v : cell.w || '').trim() : '');
      }
      var hits = 0;
      A.ROSTER_FIELDS.forEach(function (f) {
        if (vals.indexOf(f.label) >= 0 || f.aliases.some(function (a) { return vals.indexOf(a) >= 0; })) hits++;
      });
      if (hits >= 2) return vals.filter(function (v) { return v; });
    }
    return [];
  }

  function refreshRoster() {
    if (!state.roster) return;
    var hs = rosterHeaders();
    state.roster.headers = hs;
    if (!state.rosterGuessed) {
      var guess = A.guessRosterMap(hs);
      Object.keys(guess).forEach(function (k) { state.cfg.rosterMap[k] = guess[k]; });
      state.rosterGuessed = true;
      saveCfg();
    }
    reparseRoster();
  }

  function reparseRoster() {
    if (!state.roster || !state.roster.wb) return;
    try { state.roster.rows = A.parseRoster(state.roster.wb, state.cfg.rosterMap); }
    catch (e) { state.roster.rows = []; }
  }

  /* ---------------- 机师岗候选 ---------------- */

  function refreshMachine() {
    var posts = {}, order = [];
    state.shr.forEach(function (f) {
      (f.parsed.rows || []).forEach(function (r) {
        var p = r.post || '';
        if (!p) return;
        if (!posts[p]) { posts[p] = true; order.push(p); }
      });
    });
    var known = {};
    state.machineCandidates.forEach(function (m) { known[m.name] = m.checked; });
    state.machineCandidates = order.map(function (p) {
      return { name: p, checked: p in known ? known[p] : p.indexOf('机师') >= 0 };
    });
    state.cfg.machineShifts = state.machineCandidates.filter(function (m) { return m.checked; }).map(function (m) { return m.name; });
    saveCfg();
  }

  /* ---------------- 渲染 ---------------- */

  function renderAll() {
    renderZones();
    renderRosterPanel();
    renderMachinePanel();
    renderPeriod();
  }

  function zoneHtml(zone, title, sub, multiple) {
    var list = zone === 'roster' ? (state.roster ? [state.roster] : []) : state[zone];
    var h = '<div class="att-zone"><div class="drop" id="drop-' + zone + '">' +
      '<div class="d-main">' + title + '</div><div class="d-sub">' + sub + '</div>' +
      '<input type="file" id="picker-' + zone + '" accept=".xlsx,.xlsm,.xls"' + (multiple ? ' multiple' : '') + ' class="hidden"></div>' +
      '<div class="file-list" id="list-' + zone + '">';
    list.forEach(function (item) {
      var err = item.parsed && item.parsed.error;
      var info = '';
      if (item === state.roster) info = item.rows.length ? '识别到 ' + item.rows.length + ' 名员工' : '尚未配置字段映射';
      else info = err ? ('⚠ ' + esc(err)) : ('工作表「' + esc(item.parsed.sheetName) + '」· ' + item.parsed.rows.length + ' 行 · 蝶城：' + esc(item.parsed.city));
      h += '<div class="file-item"><div class="fi-head"><div class="fi-name"><span class="fi-dot" style="background:' +
        (err ? 'var(--red)' : 'var(--green)') + '"></span><span class="fi-nm">' + esc(item.name) + '</span></div>' +
        '<button class="btn-ghost" data-remove="' + zone + ':' + item.id + '">移除</button></div>' +
        '<div class="att-file-meta">' + info + '</div></div>';
    });
    h += '</div></div>';
    return h;
  }

  function renderZones() {
    var box = $('#attZones');
    if (!box) return;
    box.innerHTML =
      zoneHtml('shr', '外包考勤（SHR）', '含「SHR员工考勤」工作表的原始导出', true) +
      zoneHtml('rm', '机师岗 RM 考勤', '原始导出，取第 1 张工作表', true) +
      zoneHtml('roster', '人员花名册', '全量人员一张表（用于匹配项目/在职）', false);
    ['shr', 'rm', 'roster'].forEach(function (zone) {
      (function (zone) {
        var drop = $('#drop-' + zone), picker = $('#picker-' + zone);
        drop.addEventListener('click', function (e) { if (e.target === drop || !e.target.closest('input')) picker.click(); });
        picker.addEventListener('change', function () { addFiles(zone, picker.files); picker.value = ''; });
        drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('drag'); });
        drop.addEventListener('dragleave', function () { drop.classList.remove('drag'); });
        drop.addEventListener('drop', function (e) {
          e.preventDefault(); drop.classList.remove('drag');
          addFiles(zone, e.dataTransfer.files);
        });
      })(zone);
    });
    $$('[data-remove]').forEach(function (b) {
      b.addEventListener('click', function () {
        var p = b.dataset.remove.split(':');
        removeItem(p[0], Number(p[1]));
      });
    });
  }

  function renderRosterPanel() {
    var box = $('#rosterPanel');
    if (!box) return;
    if (!state.roster) {
      box.innerHTML = '';
      return;
    }
    var hs = state.roster.headers;
    var h = '<table class="map-table"><thead><tr><th style="width:150px">花名册字段</th><th>对应列名</th><th>说明</th></tr></thead><tbody>';
    A.ROSTER_FIELDS.forEach(function (f) {
      var cur = state.cfg.rosterMap[f.key] || '';
      h += '<tr><td><b>' + esc(f.label) + '</b></td>' +
        '<td><select data-rmap="' + f.key + '">' +
        '<option value="">（自动识别）</option>' +
        hs.map(function (x) { return '<option value="' + esc(x) + '"' + (x === cur ? ' selected' : '') + '>' + esc(x) + '</option>'; }).join('') +
        '</select></td>' +
        '<td class="col-tip">' + (f.key === 'code' ? '工号优先匹配，姓名兜底' : f.key === 'city' ? '按此字段归属蝶城（无出勤名单）' : f.key === 'hire' || f.key === 'leave' ? '用于判定考勤期内是否在职' : '') + '</td></tr>';
    });
    h += '</tbody></table><div class="hint">已识别 ' + (state.roster.rows.length) + ' 名员工。</div>';
    box.innerHTML = h;
    $$('[data-rmap]', box).forEach(function (sel) {
      sel.addEventListener('change', function () {
        state.cfg.rosterMap[sel.dataset.rmap] = sel.value;
        saveCfg();
        reparseRoster();
        renderAll();
      });
    });
  }

  function renderMachinePanel() {
    var box = $('#machinePanel');
    if (!box) return;
    if (!state.shr.length) {
      box.innerHTML = '';
      return;
    }
    var h = '<div class="ig-head"><b>机师岗名单</b><span>勾选的岗位从外包考勤剔除（RM 考勤不受影响）：</span></div><div class="ig-list">';
    state.machineCandidates.forEach(function (m, i) {
      h += '<label class="ig-item"><input type="checkbox" data-machine="' + i + '"' + (m.checked ? ' checked' : '') + '><span title="' + esc(m.name) + '">' + esc(m.name) + '</span></label>';
    });
    h += '</div>';
    box.innerHTML = h;
    $$('[data-machine]', box).forEach(function (cb) {
      cb.addEventListener('change', function () {
        state.machineCandidates[Number(cb.dataset.machine)].checked = cb.checked;
        state.cfg.machineShifts = state.machineCandidates.filter(function (m) { return m.checked; }).map(function (m) { return m.name; });
        saveCfg();
      });
    });
  }

  function renderPeriod() {
    var box = $('#attPeriod');
    if (!box) return;
    var cities = {}, multiAny = false;
    ['shr', 'rm'].forEach(function (zone) {
      state[zone].forEach(function (f) {
        if (!f.parsed || f.parsed.error) return;
        // 蝶城优先取源表蝶城列/花名册匹配，而不是文件名
        var inf;
        try { inf = A.inferFileCityFor(f.parsed, state.roster ? state.roster.rows : []); }
        catch (e) { inf = { city: f.parsed.city, multi: false }; }
        if (inf.multi) multiAny = true;
        cities[inf.city] = (cities[inf.city] || 0) + 1;
      });
    });
    var keys = Object.keys(cities);
    if (!keys.length) { box.innerHTML = '<span class="sub">上传考勤文件后自动按蝶城分组处理</span>'; return; }
    box.innerHTML = '<span class="sub">将生成 ' + keys.length + ' 个蝶城结果：' +
      keys.map(function (k) { return esc(k) + '（' + cities[k] + ' 个文件）'; }).join('、') +
      (multiAny ? '；含多蝶城的文件会按每条记录的蝶城自动拆分归组' : '') + '</span>';
  }

  /* ---------------- 生成 ---------------- */

  function run() {
    if (!state.shr.length && !state.rm.length) { toast('请先上传外包考勤或 RM 考勤文件', 'warn'); return; }
    if (!state.roster) { toast('请先上传人员花名册（补项目、算在职无出勤都依赖它）', 'warn'); return; }
    if (!state.roster.rows.length) { toast('花名册未识别到员工行，请检查字段映射', 'warn'); return; }

    var btn = $('#btnAttRun');
    btn.disabled = true;
    btn.textContent = '生成中…';

    setTimeout(function () {
      doRun().catch(function (e) {
        toast('生成失败：' + e.message, 'warn');
        console.error(e);
      }).then(function () {
        btn.disabled = false;
        btn.textContent = '开始生成';
      });
    }, 30);
  }

  // 让出主线程一次：浏览器有机会重绘进度文字，避免大文件处理时弹「页面无响应」
  function yieldUI() { return new Promise(function (res) { setTimeout(res, 0); }); }

  function doRun() {
    var files = [];
    state.shr.forEach(function (f) { files.push({ type: 'shr', parsed: f.parsed }); });
    state.rm.forEach(function (f) { files.push({ type: 'rm', parsed: f.parsed }); });

    var btn = $('#btnAttRun');
    var res = A.processAll(files, state.roster.rows, state.cfg);
    var specs = [];
    res.cities.forEach(function (c) {
      A.cityFileSpecs(c).forEach(function (f) {
        specs.push({ name: f.name, city: c.city, kind: 'city', sheets: f.sheets });
      });
    });
    var months = res.cities.map(function (c) { return c.monthLabel; }).filter(function (x) { return x; });
    var monthLabel = months.length ? months.sort()[0] : '';
    specs.push({ name: '【考勤差异】' + monthLabel + '迟到早退汇总.xlsx', city: '', kind: 'summary', sheets: A.summarySpecs(res.summary) });

    // 逐个生成：每个文件之间让出主线程，按钮显示进度；最终产物与原来 Promise.all 完全一致
    return specs.reduce(function (chain, sp, i) {
      return chain.then(function (arr) {
        btn.textContent = '生成中 ' + (i + 1) + '/' + specs.length;
        return yieldUI().then(function () {
          // 每个 xlsx 的 zip 内部仍按文件回调进度；这里把最外层进度接到按钮
          return W.build({
            zipImpl: Z, sheets: sp.sheets,
            onProgress: function (done, total) {
              btn.textContent = '生成中 ' + (i + 1) + '/' + specs.length + ' (' + done + '/' + total + ')';
            }
          }).then(function (bytes) { arr.push(bytes); return arr; });
        });
      });
    }, Promise.resolve([])).then(function (arr) {
      var files2 = arr.map(function (bytes, i) {
        return {
          name: specs[i].name, kind: specs[i].kind, city: specs[i].city,
          blob: new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
        };
      });
      state.results = { files: files2, cities: res.cities, summary: res.summary, problems: res.problems };
      renderResults();
      toast('生成完成：' + files2.length + ' 个文件', 'ok');
      return files2;
    });
  }

  /* ---------------- 结果渲染 ---------------- */

  function renderResults() {
    var box = $('#attResults');
    if (!state.results) { box.innerHTML = ''; return; }
    var r = state.results;
    var h = '';

    /* 每个蝶城一张结果卡（可能含外包/机师岗两个文件） */
    r.cities.forEach(function (c) {
      var cFiles = r.files.map(function (f, i) { return { f: f, i: i }; })
        .filter(function (x) { return x.f.kind === 'city' && x.f.city === c.city; });
      h += '<div class="result-card"><div class="rc-head">' +
        '<div><span class="rc-city">' + esc(c.city) + '</span>' +
        '<span class="rc-file">考勤期间 ' + esc(c.periodStart || '') + ' ~ ' + esc(c.periodEnd || '') + ' · ' + c.peopleCount + ' 人</span>' +
        '</div><div class="rc-actions">' +
        cFiles.map(function (x) {
          return '<button class="btn" data-att-dl="' + x.i + '">' + (x.f.name.indexOf('机师岗') >= 0 ? '下载机师岗' : '下载外包') + '</button>';
        }).join('') +
        '</div></div>';
      h += '<div class="rc-stats">' +
        '<div><b>' + (c.shrRows.length + c.rmRows.length) + '</b><span>主表行数</span></div>' +
        '<div><b>' + c.diffRows.length + '</b><span>考勤差异</span></div>' +
        '<div><b>' + c.statRows.length + '</b><span>有迟到早退的人数</span></div>' +
        '<div><b>' + roundDed(c.statRows) + '</b><span>扣款合计（元）</span></div>' +
        '<div><b>' + c.noAttendance.length + '</b><span>在职无出勤</span></div>' +
        '</div>';
      h += '</div>';
    });

    /* 跨蝶城汇总 */
    r.files.forEach(function (f, i) {
      if (f.kind !== 'summary') return;
      h += '<div class="result-card"><div class="rc-head">' +
        '<div><span class="rc-city">' + esc(f.name) + '</span><span class="rc-file">跨蝶城汇总</span></div>' +
        '<div class="rc-actions"><button class="btn" data-att-dl="' + i + '">下载 xlsx</button></div></div>';
      h += '<div class="rc-stats">' +
        '<div><b>' + r.summary.statRows.length + '</b><span>统计行数</span></div>' +
        '<div><b>' + r.summary.detailRows.length + '</b><span>迟到早退明细</span></div>' +
        '<div><b>' + r.summary.noAttendance.length + '</b><span>在职无出勤</span></div>' +
        '</div></div>';
    });

    /* 数据统计看板 + 在职无出勤明细 */
    var T = { rows: 0, diff: 0, stat: 0, ded: 0, noAtt: 0 };
    r.cities.forEach(function (c) {
      T.rows += c.shrRows.length + c.rmRows.length;
      T.diff += c.diffRows.length;
      T.stat += c.statRows.length + (c.statRowsRm || []).length;
      T.ded += roundDed(c.statRows) + roundDed(c.statRowsRm || []);
      T.noAtt += c.noAttendance.length;
    });
    T.ded = Math.round(T.ded * 100) / 100;
    var noAttAll = [];
    r.cities.forEach(function (c) {
      c.noAttendance.forEach(function (a) { noAttAll.push(a); });
    });

    h += '<div class="result-card"><div class="rc-head"><div><span class="rc-city">数据统计</span>' +
      '<span class="rc-file">全部蝶城汇总</span></div></div>';
    h += '<div class="rc-stats">' +
      '<div><b>' + T.rows + '</b><span>主表行数</span></div>' +
      '<div><b>' + T.diff + '</b><span>考勤差异</span></div>' +
      '<div><b>' + T.stat + '</b><span>有迟到早退的人数</span></div>' +
      '<div><b>' + T.ded + '</b><span>扣款合计（元）</span></div>' +
      '<div><b>' + T.noAtt + '</b><span>在职无出勤</span></div>' +
      '</div>';

    h += '<div class="att-subhead"><div><span class="t">在职无出勤人员明细</span><span class="n">' +
      noAttAll.length + ' 人（花名册在册、考勤期内无任何考勤记录）</span></div>' +
      (noAttAll.length ? '<button class="btn btn-line" id="btnNoAtt">下载明细.xlsx</button>' : '') + '</div>';
    if (noAttAll.length) {
      h += '<table class="att-table"><thead><tr>' +
        '<th>蝶城</th><th>姓名</th><th>员工编码</th><th>项目</th><th>入职日期</th><th>离职日期</th></tr></thead><tbody>' +
        noAttAll.map(function (a) {
          return '<tr><td>' + esc(a[0]) + '</td><td>' + esc(a[1]) + '</td><td>' + esc(a[2] || '—') + '</td><td>' +
            esc(a[3] || '—') + '</td><td>' + esc(a[4] || '—') + '</td><td>' + esc(a[5] || '—') + '</td></tr>';
        }).join('') + '</tbody></table>';
    } else {
      h += '<div class="att-empty">无在职无出勤人员</div>';
    }
    h += '</div>';

    /* 问题清单放最后（只显示系统计算产生的问题） */
    if (r.problems.length) {
      h += '<div class="result-card"><div class="rc-head"><div><span class="rc-city">问题清单</span>' +
        '<span class="rc-file">' + r.problems.length + ' 条</span></div>' +
        '<div class="rc-actions"><button class="btn btn-line" id="btnAttProblems">下载问题清单.xlsx</button></div></div>' +
        '<div class="att-problems">' +
        r.problems.slice(0, 30).map(function (p) {
          return '<div class="att-prob-item sev-' + p.sev + '">' +
            '<span class="badge ' + (p.sev === 'error' ? 'bad' : p.sev === 'warn' ? 'warn' : 'gray') + '">' +
            (p.sev === 'error' ? '错误' : p.sev === 'warn' ? '提示' : '信息') + '</span> ' +
            (p.city ? '[' + esc(p.city) + '] ' : '') + esc(p.msg) + '</div>';
        }).join('') +
        (r.problems.length > 30 ? '<div class="att-prob-item">…其余 ' + (r.problems.length - 30) + ' 条见下载的问题清单</div>' : '') +
        '</div></div>';
    }

    box.innerHTML = h;

    $$('[data-att-dl]', box).forEach(function (b) {
      b.addEventListener('click', function () {
        var f = r.files[Number(b.dataset.attDl)];
        if (f) downloadBlob(f.blob, f.name);
      });
    });
    var pb = $('#btnAttProblems');
    if (pb) pb.addEventListener('click', function () { downloadProblems(r.problems); });
    var nb = $('#btnNoAtt');
    if (nb) nb.addEventListener('click', function () {
      W.build({
        zipImpl: Z,
        sheets: [{
          name: '在职无出勤',
          header: ['蝶城', '姓名', '员工编码', '项目', '入职日期', '离职日期'],
          rows: noAttAll,
          cols: [22, 10, 14, 16, 12, 12], headerStyle: 5, dataStyle: 2
        }]
      }).then(function (bytes) {
        downloadBlob(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), '在职无出勤人员明细.xlsx');
      });
    });
  }

  function roundDed(statRows) {
    var s = 0;
    statRows.forEach(function (r) { s += r[12]; });
    return Math.round(s * 100) / 100;
  }

  function downloadProblems(problems) {
    var LV = { error: '错误', warn: '提示', info: '信息' };
    W.build({
      zipImpl: Z,
      sheets: [{
        name: '问题清单',
        header: ['蝶城', '工作表', '员工姓名', '说明', '级别'],
        rows: problems.map(function (p) { return [p.city, p.sheet, p.name, p.msg, LV[p.sev] || p.sev]; }),
        cols: [12, 14, 10, 80, 8], headerStyle: 5, dataStyle: 2
      }]
    }).then(function (bytes) {
      downloadBlob(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), '考勤问题清单.xlsx');
    });
  }

  /* ---------------- 事件绑定 ---------------- */

  function bind() {
    $('#btnAttRun').addEventListener('click', run);
    var zb = $('#btnAttZipTop');
    if (zb) zb.addEventListener('click', function () {
      var r = state.results;
      if (!r || !r.files.length) { toast('请先点击「开始生成」', 'warn'); return; }
      Promise.all(r.files.map(function (f) {
        return f.blob.arrayBuffer().then(function (ab) { return { name: f.name, data: new Uint8Array(ab) }; });
      })).then(function (entries) {
        return Z.zip(entries);
      }).then(function (u8) {
        downloadBlob(new Blob([u8], { type: 'application/zip' }), '考勤处理结果.zip');
      });
    });
  }

  /* ---------------- 初始化 ---------------- */

  document.addEventListener('DOMContentLoaded', function () {
    if (!$('#viewAttendance')) return;
    state.cfg = loadCfg();
    bind();
    renderAll();
  });

  window.AttApp = { onShow: function () { state.cfg = loadCfg(); renderAll(); } };
})();
