/* ============================================================
 * att-app.js — 考勤处理模块界面
 * 三个上传区（外包SHR / RM / 花名册）+ 花名册映射 + 外包考勤岗位剔除 + 生成与下载
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
    postOptions: [],  // 外包考勤里出现的岗位名（下拉多选候选）
    seenPosts: {},    // 已自动带入过「机师」岗位，避免用户取消后被重新加回
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

  /* ---------------- 剔除岗位候选 ---------------- */

  function refreshMachine() {
    var posts = {}, order = [];
    state.shr.forEach(function (f) {
      (f.parsed.rows || []).forEach(function (r) {
        var p = r.post || '';
        if (!p || posts[p]) return;
        posts[p] = true; order.push(p);
      });
    });
    state.postOptions = order.sort(function (a, b) { return a.localeCompare(b, 'zh'); });
    var shifts = state.cfg.machineShifts || (state.cfg.machineShifts = []);
    // 首次出现的岗位名含「机师」时自动带入（用户手动取消后不再自动加回）
    state.postOptions.forEach(function (p) {
      if (state.seenPosts[p]) return;
      state.seenPosts[p] = 1;
      if (p.indexOf('机师') >= 0 && shifts.indexOf(p) < 0) shifts.push(p);
    });
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
    if (!state.roster) { box.innerHTML = ''; return; }
    var hs = state.roster.headers;
    // 紧凑 2 列网格：标签在上、下拉在下，不再用表格和「说明」列
    var h = '<div class="sec-h">花名册字段映射</div><div class="rmap-grid">';
    A.ROSTER_FIELDS.forEach(function (f) {
      var cur = state.cfg.rosterMap[f.key] || '';
      h += '<label class="rmap-item"><span>' + esc(f.label) + '</span>' +
        '<select data-rmap="' + f.key + '"><option value="">（自动识别）</option>' +
        hs.map(function (x) { return '<option value="' + esc(x) + '"' + (x === cur ? ' selected' : '') + '>' + esc(x) + '</option>'; }).join('') +
        '</select></label>';
    });
    h += '</div><div class="rmap-foot">已识别 <b>' + state.roster.rows.length + '</b> 名员工</div>';
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

  /* 当前生效的剔除项：下拉岗位 + 「其他」输入框的词，归一化后去重 */
  function exItems() {
    var out = [], seen = {};
    (state.cfg.machineShifts || []).forEach(function (p) {
      var k = A.norm(p); if (!k || seen[k]) return;
      seen[k] = 1; out.push({ v: p, t: 'post' });
    });
    A.splitTokens(state.cfg.machineExtraText).forEach(function (t) {
      var k = A.norm(t); if (!k || seen[k]) return;
      seen[k] = 1; out.push({ v: t, t: 'extra' });
    });
    return out;
  }

  function renderChips() {
    var box = $('#machineChips');
    if (!box) return;
    var items = exItems();
    var head = $('#machineChipsHead');
    if (head) head.textContent = '当前剔除 ' + items.length + ' 项';
    box.innerHTML = items.length ? items.map(function (x) {
      return '<span class="ex-chip"><span title="' + esc(x.v) + '">' + esc(x.v) + '</span>' +
        '<button type="button" data-chip-t="' + x.t + '" data-chip-v="' + esc(x.v) + '" title="移除">×</button></span>';
    }).join('') : '<span class="ex-empty">未设置，外包考勤全部保留</span>';
    $$('[data-chip-t]', box).forEach(function (b) {
      b.addEventListener('click', function () {
        var v = b.dataset.chipV;
        if (b.dataset.chipT === 'post') {
          state.cfg.machineShifts = (state.cfg.machineShifts || []).filter(function (p) { return p !== v; });
        } else {
          var left = A.splitTokens(state.cfg.machineExtraText).filter(function (t) { return A.norm(t) !== A.norm(v); });
          state.cfg.machineExtraText = left.join(', ');
        }
        saveCfg();
        renderMachinePanel();
      });
    });
  }

  function renderMachinePanel() {
    var box = $('#machinePanel');
    if (!box) return;
    var selSet = {};
    (state.cfg.machineShifts || []).forEach(function (p) { selSet[p] = 1; });
    var h = '<div class="sec-h">外包考勤岗位剔除</div>' +
      '<div class="ex-card">' +
        '<div class="ex-sub">选中的岗位、或「其他」里填写的岗位名 / 员工姓名，会从外包考勤中剔除；RM 考勤不受影响。</div>' +
        '<label class="ex-fld"><span>① 剔除岗位 <em>（Ctrl / Shift 可多选）</em></span>' +
          '<select id="machineSel" class="ex-sel" multiple size="5">' +
          (state.postOptions.length
            ? state.postOptions.map(function (p) {
                return '<option value="' + esc(p) + '"' + (selSet[p] ? ' selected' : '') + '>' + esc(p) + '</option>';
              }).join('')
            : '<option value="" disabled>上传外包考勤后自动列出岗位</option>') +
          '</select></label>' +
        '<label class="ex-fld"><span>② 其他 <em>（岗位名或姓名，逗号 / 空格分隔）</em></span>' +
          '<input id="machineExtra" class="ex-input" type="text" placeholder="如：机师岗, 维保机师 或 张三 李四" value="' +
          esc(state.cfg.machineExtraText || '') + '"></label>' +
        '<div class="ex-chips-head" id="machineChipsHead"></div>' +
        '<div class="ex-chips" id="machineChips"></div>' +
      '</div>';
    box.innerHTML = h;
    var sel = $('#machineSel', box);
    if (sel) sel.addEventListener('change', function () {
      state.cfg.machineShifts = Array.prototype.slice.call(sel.options)
        .filter(function (o) { return o.selected && o.value; })
        .map(function (o) { return o.value; });
      saveCfg();
      renderChips();
    });
    var inp = $('#machineExtra', box);
    if (inp) inp.addEventListener('input', function () {
      state.cfg.machineExtraText = inp.value;
      saveCfg();
      renderChips();
    });
    renderChips();
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
    specs.push({ name: '【考勤差异】' + monthLabel + '差异数据汇总.xlsx', city: '', kind: 'summary', sheets: A.summarySpecs(res.summary, res.cities) });

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

    /* 三个模块的文件分组 */
    var outFiles = [], rmFiles = [], sumFiles = [];
    r.files.forEach(function (f, i) {
      var item = { f: f, i: i };
      if (f.kind === 'summary') sumFiles.push(item);
      else if (f.name.indexOf('机师岗') >= 0) rmFiles.push(item);
      else outFiles.push(item);
    });

    /* 各组差异数：外包=shr差异、机师岗=rm差异 */
    var outDiff = 0, rmDiff = 0, outStat = 0, rmStat = 0, outDed = 0, rmDed = 0;
    r.cities.forEach(function (c) {
      // 用 姓名|日期 集合区分差异行属于外包还是机师岗（避免逐行 O(n²) 扫描）
      var shrKey = {}, rmKey = {};
      (c.shrRows || []).forEach(function (s) { shrKey[s[0] + '|' + s[3]] = 1; });
      (c.rmRows || []).forEach(function (s) { rmKey[s[1] + '|' + s[3]] = 1; });
      (c.diffRows || []).forEach(function (x) {
        var k = x[0] + '|' + x[3];
        if (shrKey[k]) outDiff++;
        else if (rmKey[k]) rmDiff++;
      });
      outStat += c.statRows.length;
      rmStat += (c.statRowsRm || []).length;
      outDed += roundDed(c.statRows);
      rmDed += roundDed(c.statRowsRm || []);
    });
    outDed = Math.round(outDed * 100) / 100;
    rmDed = Math.round(rmDed * 100) / 100;

    function fileListTip(items) {
      if (!items.length) return '';
      return '<div class="filelist-tip">' + items.map(function (x) {
        return '<div>' + esc(x.f.name) + '</div>';
      }).join('') + '</div>';
    }
    /* 一张模块大卡：标题 + 下载按钮 + 数字统计区。
       文件清单挂在「数字统计区」悬浮，不在标题下再显示「N 个文件」小字；
       气泡本身不接收鼠标事件，避免盖住鼠标引起 显示/隐藏 来回闪烁。 */
    function bigCard(title, items, stats, btnLabel, dlAttr, showList) {
      var tip = showList ? fileListTip(items) : '';
      h += '<div class="module-card"><div class="mc-head">' +
        '<span class="mc-title">' + title + '</span>' +
        '<button class="btn" ' + dlAttr + '>' + btnLabel + '</button></div>' +
        '<div class="mc-stats' + (tip ? ' filelist-host' : '') + '">' + stats.map(function (s) {
          return '<div class="mc-stat"><b' + (tip && s[2] ? ' class="mc-file-num"' : '') + '>' + s[0] +
            '</b><span>' + s[1] + '</span></div>';
        }).join('') + tip + '</div></div>';
    }
    h += '<div class="module-cards">';
    bigCard('外包考勤结果', outFiles,
      [['' + outDiff, '考勤差异数'], ['' + outFiles.length, '已生成文件数', true]],
      '下载外包考勤', 'data-att-group="out"', true);
    bigCard('机师岗RM考勤结果', rmFiles,
      [['' + rmDiff, '考勤差异数'], ['' + rmFiles.length, '已生成文件数', true]],
      '下载机师岗RM考勤', 'data-att-group="rm"', true);
    var sumStat = r.summary || { statRows: [], detailRows: [], noAttendance: [] };
    // 数据汇总固定只有 1 个文件，不再显示文件数，也不挂文件清单
    bigCard('数据汇总', sumFiles,
      [['' + sumStat.statRows.length, '统计行数'], ['' + sumStat.detailRows.length, '迟到早退明细'],
       ['' + (sumStat.statRows.length + sumStat.detailRows.length + sumStat.noAttendance.length), '人员合计'],
       ['' + A.statDed(sumStat.statRows), '扣款合计（元）'], ['' + sumStat.noAttendance.length, '在职无出勤']],
      '下载汇总结果', 'data-att-group="sum"', false);
    h += '</div>';

    /* 在职无出勤明细（保持现状，放在三卡下面） */
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
    // 三大模块批量下载（非 zip，逐个触发浏览器下载，按钮显示 1/N 进度）
    $$('[data-att-group]', box).forEach(function (b) {
      b.addEventListener('click', function () {
        var grp = b.dataset.attGroup;
        var items = grp === 'out' ? outFiles : grp === 'rm' ? rmFiles : sumFiles;
        if (!items.length) { toast('该模块没有可下载的文件', 'warn'); return; }
        toast('浏览器可能询问「是否允许多个文件下载」，请点允许，约 ' + items.length + ' 个文件', 'warn');
        var orig = b.textContent, i = 0;
        b.disabled = true;
        function next() {
          if (i >= items.length) { b.textContent = orig; b.disabled = false; return; }
          b.textContent = '下载中 ' + (i + 1) + '/' + items.length;
          downloadBlob(items[i].f.blob, items[i].f.name);
          i++;
          setTimeout(next, 120);   // 间隔触发，避免浏览器把连续下载当作误操作拦截
        }
        next();
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
    // 统一走 AttCore：自动兼容「蝶城统计表 14 列」与「汇总总表 15 列（带蝶城）」
    return A.statDed(statRows);
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
      var orig = zb.textContent, done = 0, total = r.files.length;
      zb.disabled = true;
      zb.textContent = '打包中 0/' + total;
      // 逐个读取+打包进度，避免大文件打包时看起来卡死
      var entries = [];
      function next(i) {
        if (i >= total) {
          zb.textContent = '压缩中…';
          return Z.zip(entries, function (d, t) { zb.textContent = '压缩中 ' + d + '/' + t; }).then(function (u8) {
            downloadBlob(new Blob([u8], { type: 'application/zip' }), '考勤处理结果.zip');
            zb.textContent = orig; zb.disabled = false;
          });
        }
        var f = r.files[i];
        zb.textContent = '打包中 ' + (i + 1) + '/' + total;
        return f.blob.arrayBuffer().then(function (ab) {
          entries.push({ name: f.name, data: new Uint8Array(ab) });
          return next(i + 1);
        });
      }
      next(0).catch(function (e) { toast('打包失败：' + e.message, 'warn'); zb.textContent = orig; zb.disabled = false; });
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
