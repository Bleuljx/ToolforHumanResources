/* ============================================================
 * app.js — 界面与交互
 * ============================================================ */
(function () {
  'use strict';

  var XLSX = window.XLSX;
  var C = window.SalaryCore;
  var W = window.XlsxWriter;
  var Z = window.ZipUtil;
  var OPTIONS = window.SALARY_OPTIONS.options;
  var IMP_HEADERS = window.SALARY_OPTIONS.headers;
  var CFG_KEY = 'salaryTool.cfg.v1';

  var state = {
    files: [],
    cfg: loadConfig(),
    results: [],
    seq: 1
  };

  /* ---------------- 配置存取 ---------------- */
  function deepMerge(base, extra) {
    var out = JSON.parse(JSON.stringify(base));
    Object.keys(extra || {}).forEach(function (k) {
      if (extra[k] && typeof extra[k] === 'object' && !Array.isArray(extra[k])) {
        out[k] = deepMerge(out[k] || {}, extra[k]);
      } else if (extra[k] !== undefined) {
        out[k] = extra[k];
      }
    });
    return out;
  }
  function loadConfig() {
    try {
      var raw = localStorage.getItem(CFG_KEY);
      if (raw) return deepMerge(C.DEFAULT_CONFIG, JSON.parse(raw));
    } catch (e) { }
    return JSON.parse(JSON.stringify(C.DEFAULT_CONFIG));
  }
  var saveTimer = null;
  function saveConfig() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      try { localStorage.setItem(CFG_KEY, JSON.stringify(state.cfg)); } catch (e) { }
    }, 150);
  }

  /* ---------------- 工具 ---------------- */
  function $(sel, ctx) { return (ctx || document).querySelector(sel); }
  function $$(sel, ctx) { return Array.prototype.slice.call((ctx || document).querySelectorAll(sel)); }
  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function fmtMoney(n) {
    var v = Number(n) || 0;
    return '¥' + v.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  var toastTimer = null;
  function toast(msg, type) {
    var el = $('#toast');
    el.textContent = msg;
    el.className = 'toast show' + (type ? ' ' + type : '');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { el.className = 'toast'; }, 3200);
  }
  function headersOf() {
    var f = state.files[0];
    return f && f.parsed ? f.parsed.headers : [];
  }
  function optionsHtml(list, current, emptyLabel) {
    var html = '<option value="">' + esc(emptyLabel || '（不取值）') + '</option>';
    list.forEach(function (h) {
      var name = typeof h === 'string' ? h : h.name;
      html += '<option value="' + esc(name) + '"' + (name === current ? ' selected' : '') + '>' + esc(name) + '</option>';
    });
    return html;
  }

  /* ---------------- 文件上传 ---------------- */
  function handleFiles(fileList) {
    var arr = Array.prototype.slice.call(fileList);
    if (!arr.length) return;
    var pending = arr.length;
    arr.forEach(function (f) {
      if (!/\.(xlsx|xlsm|xls)$/i.test(f.name)) {
        toast('不支持的文件类型：' + f.name, 'warn');
        if (--pending === 0) renderFiles();
        return;
      }
      var reader = new FileReader();
      reader.onload = function (e) {
        try {
          var parsed = C.parseFile(XLSX, e.target.result, f.name, null);
          state.files.push({
            id: state.seq++,
            name: f.name,
            buffer: e.target.result,
            parsed: parsed,
            city: parsed.city
          });
        } catch (err) {
          toast('解析失败：' + f.name + '（' + (err && err.message ? err.message : err) + '）', 'warn');
        }
        if (--pending === 0) { renderFiles(); renderPanels(); }
      };
      reader.readAsArrayBuffer(f);
    });
  }

  function reparseFile(item) {
    var parsed = C.parseFile(XLSX, item.buffer, item.name, {
      sheetName: item.parsed.sheetName,
      headerRow: item.parsed.headerRow,
      dataStartRow: item.parsed.dataStartRow,
      city: item.city
    });
    item.parsed = parsed;
    renderFiles();
    renderPanels();
  }

  function renderFiles() {
    var box = $('#fileList');
    if (!state.files.length) {
      box.innerHTML = '';
      $('#drop').classList.remove('has-file');
      return;
    }
    $('#drop').classList.add('has-file');
    box.innerHTML = state.files.map(function (f) {
      var p = f.parsed;
      var sheets = p.sheetNames.map(function (n) {
        return '<option value="' + esc(n) + '"' + (n === p.sheetName ? ' selected' : '') + '>' + esc(n) + '</option>';
      }).join('');
      return '' +
        '<div class="file-item" data-id="' + f.id + '">' +
        '  <div class="fi-head">' +
        '    <span class="fi-name"><span class="fi-dot"></span><span class="fi-nm" title="' + esc(f.name) + '">' + esc(f.name) + '</span></span>' +
        '    <button class="btn-ghost fi-del" data-act="del">移除</button>' +
        '  </div>' +
        '  <div class="fi-grid">' +
        '    <label><span>蝶城（文件名）</span><input class="fi-city" data-act="city" value="' + esc(f.city) + '"></label>' +
        '    <label><span>工作表</span><select data-act="sheet">' + sheets + '</select></label>' +
        '    <label><span>表头行</span><input type="number" min="1" data-act="hrow" value="' + p.headerRow + '"></label>' +
        '    <label><span>数据起始行</span><input type="number" min="1" data-act="drow" value="' + p.dataStartRow + '"></label>' +
        '    <div class="fi-meta">识别到 <b>' + p.headers.length + '</b> 列 · 首列「' + esc(p.headers.length ? p.headers[0].name : '-') + '」</div>' +
        '  </div>' +
        '</div>';
    }).join('');
  }

  /* ---------------- 配置面板 ---------------- */
  function openPanels() { return state.openPanels || (state.openPanels = {}); }

  function openWidePanel() {
    $$('.panel').forEach(function (p) { p.classList.remove('open'); });
    $$('.tile').forEach(function (x) { x.classList.remove('active'); });
    var p = $('#panel-wide'), t = $('.tile[data-panel="wide"]');
    if (p) p.classList.add('open');
    if (t) t.classList.add('active');
    p && p.scrollIntoView && p.scrollIntoView({ block: 'nearest' });
  }

  function renderPanels() {
    renderAttendancePanel();
    renderCategoryPanel();
    renderImportPanel();
    renderWidePanel();
  }

  function renderAttendancePanel() {
    var box = $('#panel-attendance .panel-inner');
    if (!box) return;
    var hs = headersOf();
    var hint = hs.length ? '' : '<div class="hint">先上传算薪表，这里会列出可选的源列。</div>';
    var rows = C.ATTENDANCE_KEYS.map(function (pair) {
      var key = pair[0], label = pair[1];
      var fixed = (key === 'reason' || key === 'lateCount' || key === 'missCount');
      var cur = state.cfg.attendance[key] || '';
      return '<tr>' +
        '<td class="col-out">' + esc(label) + '</td>' +
        '<td>' + (fixed
          ? '<select data-key="' + key + '">' + optionsHtml(hs, cur, '（留空，暂不取数）') + '</select>'
          : '<select data-key="' + key + '">' + optionsHtml(hs, cur, '（不取值）') + '</select>') +
        '</td>' +
        '<td class="col-tip">' + (fixed ? '占位：接了考勤结果表后再启用' : '') + '</td>' +
        '</tr>';
    }).join('');
    box.innerHTML = '<table class="map-table"><thead><tr><th>报送表列</th><th>取自算薪表</th><th>说明</th></tr></thead><tbody>' + rows + '</tbody></table>' + hint;
  }

  function renderCategoryPanel() {
    var box = $('#panel-category .panel-inner');
    if (!box) return;
    var hs = headersOf();
    var cc = state.cfg.category;
    function row(key, label, color) {
      var r = cc[key] || {};
      return '<tr>' +
        '<td><span class="tag ' + color + '">' + esc(r.type) + '</span></td>' +
        '<td><select data-range="' + key + '" data-side="start">' + optionsHtml(hs, r.start, '（请选择起始项目）') + '</select></td>' +
        '<td class="arrow">→</td>' +
        '<td><select data-range="' + key + '" data-side="end">' + optionsHtml(hs, r.end, '（请选择终止项目）') + '</select></td>' +
        '</tr>';
    }
    // 忽略列：所有已上传文件表头去重 + 手动勾选状态
    var igNames = state.cfg.ignoreCols || [];
    var seen = {}, allHeads = [];
    state.files.forEach(function (f) {
      (f.parsed.headers || []).forEach(function (h) {
        if (!seen[h.name]) { seen[h.name] = 1; allHeads.push(h); }
      });
    });
    var igItems = allHeads.length
      ? allHeads.map(function (h) {
        return '<label class="ig-item" data-name="' + esc(h.name) + '">' +
          '<input type="checkbox" data-ig="' + esc(h.name) + '"' + (igNames.indexOf(h.name) >= 0 ? ' checked' : '') + '>' +
          '<span title="' + esc(h.name) + '">' + esc(h.name) + '</span></label>';
      }).join('')
      : '<span class="ig-empty">上传算薪表后，这里会列出可选的列</span>';
    var yellowN = 0;
    state.files.forEach(function (f) { yellowN += (f.parsed.yellowCols || []).length; });

    // 考勤明细区域：自动识别（默认）或手动指定起止表头
    var dayAuto = cc.dayAuto !== false;
    var dayArea = null;
    state.files.forEach(function (f) { if (!dayArea && f.parsed.attendArea) dayArea = f.parsed.attendArea; });
    var dayTip;
    if (dayAuto) {
      dayTip = dayArea
        ? '自动识别到「' + esc(dayArea.names[0]) + '」~「' + esc(dayArea.names[dayArea.names.length - 1]) + '」共 ' + dayArea.names.length + ' 列：不进导入模板、不提示区间外'
        : '未识别到纯数字（1~31）的连续表头列，可取消勾选手动指定';
    } else {
      dayTip = '';
    }
    var dayCell = '<label class="chk-line"><input type="checkbox" data-dayauto' + (dayAuto ? ' checked' : '') + '> 自动识别 1~31 日考勤明细列</label>' +
      (dayAuto ? '<span class="col-tip">　' + dayTip + '</span>'
               : '<span class="day-manual"><select data-day="start">' + optionsHtml(hs, cc.dayStart, '（起始列）') + '</select>' +
                 '<span class="arrow">→</span>' +
                 '<select data-day="end">' + optionsHtml(hs, cc.dayEnd, '（终止列）') + '</select></span>');

    box.innerHTML = '<table class="map-table"><thead><tr><th>类别</th><th>起始项目</th><th></th><th>终止项目</th></tr></thead><tbody>' +
      row('r1', '津补贴', 'tag-green') + row('r2', '扣罚类', 'tag-red') + row('r3', '成本类', 'tag-blue') +
      '<tr><td><span class="tag tag-red">个人扣款</span></td><td colspan="2"><select data-deduct>' + optionsHtml(hs, cc.deduct, '（请选择）') + '</select></td><td class="col-tip">金额原值报送</td></tr>' +
      '<tr><td><span class="tag tag-gray">备注兜底</span></td>' +
      '<td><select data-fb="start">' + optionsHtml(hs, cc.fbStart, '（不启用兜底）') + '</select></td>' +
      '<td class="arrow">→</td>' +
      '<td><select data-fb="end">' + optionsHtml(hs, cc.fbEnd, '（不启用兜底）') + '</select></td></tr>' +
      '<tr><td><span class="tag tag-gray">考勤明细</span></td><td colspan="3" class="day-cell">' + dayCell + '</td></tr>' +
      '</tbody></table>' +
      '<div class="ig-block">' +
      '  <div class="ig-head"><b>忽略列</b>' +
      '    <span>勾选后这些列<b>一键表和导入模板都不导出</b>，也不提示「区间外项目」' + (yellowN ? '（表头标黄的 ' + yellowN + ' 列已自动忽略）' : '') + '：</span>' +
      '    <input id="igSearch" placeholder="搜索列名…">' +
      '  </div>' +
      '  <div class="ig-list">' + igItems + '</div>' +
      '</div>';
  }

  function spanHeaders() {
    var cc = state.cfg.category;
    if (!state.files.length) return [];
    var seen = {}, list = [];
    state.files.forEach(function (f) {
      var hs = f.parsed.headers;
      if (!hs.length) return;
      var s = C.colIndexByName(hs, cc.r1.start), e = C.colIndexByName(hs, cc.r3.end);
      var d = C.colIndexByName(hs, cc.deduct);
      if (!s || !e) return;
      if (e < s) { var t = s; s = e; e = t; }
      hs.forEach(function (h) {
        if ((h.col >= s && h.col <= e) || (d && h.col === d)) {
          if (!seen[h.name]) { seen[h.name] = 1; list.push(h); }
        }
      });
    });
    // 7 个考勤字段也进映射面板（key 带 @ 前缀，不会与真实表头重名）
    C.PSEUDO_FIELDS.forEach(function (pf) {
      if (!seen[pf.key]) { seen[pf.key] = 1; list.push({ name: pf.key, label: '考勤 · ' + pf.label, autoFrom: pf.from, pseudo: true }); }
    });
    return list;
  }

  var CONF_LABEL = {
    exact: ['一致', 'ok'], norm: ['归一', 'ok'], prefix: ['去前缀', 'ok'],
    guess: ['推测', 'warn'], none: ['未匹配', 'bad'], manual: ['手选', 'ok'], skip: ['不报送', 'gray']
  };

  function renderImportPanel() {
    var box = $('#panel-import .panel-inner');
    if (!box) return;
    var ic = state.cfg.importCfg;
    var hs = headersOf();
    var span = spanHeaders();
    var rows = '';
    var unresolved = 0;
    var changed = false;

    if (!span.length) {
      rows = '<tr><td colspan="4" class="col-tip">上传算薪表并配好三大区间后，这里会列出需要映射的工资项目。</td></tr>';
    } else {
      span.forEach(function (h) {
        var item = ic.map[h.name];
        var curVal0 = item && typeof item === 'object' ? item.value : (item || '');
        // 首次见到、或上次是“未匹配”状态的列：重新自动匹配，让用户一打开就看到最新结果
        if (item === undefined || curVal0 === '') {
          if (!h.pseudo && C.isSummaryCol(h.name)) {
            item = { value: '__skip__', conf: 'skip' };
          } else {
            var mm = C.matchOption(h.autoFrom || h.name, OPTIONS);
            item = { value: mm.value, conf: mm.conf };
          }
          ic.map[h.name] = item;
          changed = true;
        }
        var val = item && typeof item === 'object' ? item.value : (item || '');
        var conf = item && typeof item === 'object' ? item.conf : (val ? 'manual' : 'none');
        if (val === '__skip__') conf = 'skip';
        var meta = CONF_LABEL[conf] || CONF_LABEL.none;
        if (!val) unresolved++;
        rows += '<tr class="' + (conf === 'guess' ? 'row-warn' : '') + '">' +
          '<td class="col-out" title="' + esc(h.name) + '">' + esc(h.label || h.name) + '</td>' +
          '<td><input class="opt-input" list="optList" data-src="' + esc(h.name) + '" value="' +
          esc(val === '__skip__' ? '不报送' : val) + '" placeholder="搜索或选择标准薪酬项目"></td>' +
          '<td><span class="badge ' + meta[1] + '">' + meta[0] + '</span></td>' +
          '<td class="col-tip">' + (val === '__skip__' ? '不进入导入模板' : (val ? '→ ' + esc(val) : '必须选择后才能导出')) + '</td>' +
          '</tr>';
      });
    }

    box.innerHTML = '' +
      '<div class="import-params">' +
      '  <label><span>绑定核算任务</span><input id="impTask" placeholder="可填任务编码或名称，可留空" value="' + esc(ic.task || '') + '"></label>' +
      '  <label><span>业务日期（yyyy-MM-dd）</span><input id="impDate" type="date" value="' + esc(ic.date || '') + '"></label>' +
      '  <button class="btn-ghost" id="btnAutoMap">重新自动匹配</button>' +
      '</div>' +
      '<div class="map-wrap"><table class="map-table wide"><thead><tr><th>算薪表项目</th><th>映射到标准薪酬项目</th><th>匹配</th><th>说明</th></tr></thead><tbody>' + rows + '</tbody></table></div>' +
      (unresolved ? '<div class="hint"><b class="warn-text">还有 ' + unresolved + ' 项未匹配，需选择后才能生成</b></div>' : '');
    if (changed) saveConfig();
  }

  function autoMapAll(force) {
    var span = spanHeaders();
    if (!span.length) { toast('请先上传算薪表并配置区间', 'warn'); return; }
    var map = {};
    span.forEach(function (h) {
      var cur = state.cfg.importCfg.map[h.name];
      var curVal = cur && typeof cur === 'object' ? cur.value : cur;
      if (!force && curVal) { map[h.name] = cur; return; }
      if (!h.pseudo && C.isSummaryCol(h.name)) { map[h.name] = { value: '__skip__', conf: 'skip' }; return; }
      var m = C.matchOption(h.autoFrom || h.name, OPTIONS);
      map[h.name] = { value: m.value, conf: m.conf };
    });
    state.cfg.importCfg.map = map;
    saveConfig();
    renderImportPanel();
  }

  /* ---------------- 宽表（非循环项导入(宽表)）配置面板 ---------------- */
  function wideOptions(list, current) {
    var html = '<option value="">（不取值）</option>';
    list.forEach(function (h) {
      var name = typeof h === 'string' ? h : h.name;
      var lb = (typeof h === 'object' && h.label) ? h.label : name;
      html += '<option value="' + esc(name) + '"' + (name === current ? ' selected' : '') + '>' + esc(lb) + '</option>';
    });
    return html;
  }

  function renderWidePanel() {
    var box = $('#panel-wide .panel-inner');
    if (!box) return;
    var wc = state.cfg.wideCfg;
    // 宽表可取算薪表任意列（不限三大区间）+ 7 个考勤字段
    var allHeads = headersOf().slice();
    C.PSEUDO_FIELDS.forEach(function (pf) {
      allHeads.push({ name: pf.key, label: '考勤 · ' + pf.label });
    });
    var rows = '';
    var unresolved = 0;
    var changed = false;

    if (!allHeads.length) {
      rows = '<tr><td colspan="3" class="col-tip">上传算薪表后，这里会列出宽表 28 个项目列的取数来源（7 个考勤字段自动带出）。</td></tr>';
    } else {
      C.WIDE_ITEMS.forEach(function (tpl) {
        var cur = wc.map[tpl];
        if (cur === undefined || cur === null || cur === '') {
          if (C.WIDE_PSEUDO[tpl]) cur = C.WIDE_PSEUDO[tpl];
          else {
            var mh = null;
            state.files.forEach(function (f) {
              if (mh || !f.parsed.headers) return;
              mh = C.matchWideHeader(tpl, f.parsed.headers);
            });
            cur = mh ? mh.name : '';
          }
          wc.map[tpl] = cur;
          changed = true;
        }
        var ok = !!cur;
        if (!ok) unresolved++;
        rows += '<tr>' +
          '<td class="col-out" title="' + esc(tpl) + '">' + esc(tpl) + '</td>' +
          '<td><select data-wide="' + esc(tpl) + '">' + wideOptions(allHeads, cur) + '</select></td>' +
          '<td class="col-tip">' + (ok ? '→ ' + esc(cur) : '未选择，导出时该列留空') + '</td>' +
          '</tr>';
      });
    }

    box.innerHTML = '' +
      '<div class="import-params">' +
      '  <label><span>业务日期（yyyy-MM-dd）</span><input id="wideDate" type="date" value="' + esc(wc.date || '') + '"></label>' +
      '  <button class="btn-ghost" id="btnWideAutoMap">重新自动匹配</button>' +
      '</div>' +
      '<div class="map-wrap"><table class="map-table wide"><thead><tr><th>宽表列</th><th>取自算薪表</th><th>说明</th></tr></thead><tbody>' + rows + '</tbody></table></div>' +
      (unresolved ? '<div class="hint"><b class="warn-text">还有 ' + unresolved + ' 列未选择取数来源，生成时会弹出确认，未选择的列导出留空</b></div>' : '');
    if (changed) saveConfig();
  }

  function autoMapWide(force) {
    var wc = state.cfg.wideCfg;
    var map = {};
    var hs = [];
    state.files.forEach(function (f) { if (!hs.length && f.parsed.headers) hs = f.parsed.headers; });
    C.WIDE_ITEMS.forEach(function (tpl) {
      var cur = wc.map[tpl];
      var curVal = cur == null ? '' : String(cur);
      if (!force && curVal) { map[tpl] = curVal; return; }
      if (C.WIDE_PSEUDO[tpl]) { map[tpl] = C.WIDE_PSEUDO[tpl]; return; }
      var mh = hs.length ? C.matchWideHeader(tpl, hs) : null;
      map[tpl] = mh ? mh.name : '';
    });
    wc.map = map;
    saveConfig();
    renderWidePanel();
  }

  /* ---------------- 结果生成 ---------------- */
  function attCols() {
    var cols = [6, 10, 18, 10, 12];
    for (var i = 0; i < 12; i++) cols.push(5.7);
    return cols;
  }

  function buildXlsxFiles(res, city) {
    var attPromise = W.build({
      zipImpl: Z,
      sheets: [{
        name: '考勤报送',
        header: res.attendance.header,
        rows: res.attendance.rows,
        cols: attCols(),
        headerHeight: 62,
        rowHeight: 31.2,
        freezeHeader: true,
        autoFilter: true,
        dataStyle: 2
      }]
    });
    var catPromise = W.build({
      zipImpl: Z,
      sheets: [{
        name: '工资类目一键表',
        header: res.category.header,
        rows: res.category.rows,
        cols: [6, 10, 18, 10, 12, 12, 10, 24, 12, 34, 10],
        headerHeight: 30,
        rowHeight: 26,
        freezeHeader: true,
        autoFilter: true,
        numFmtCols: [8],
        dataStyle: 2
      }]
    });
    var impRows = res.imp.rows;
    var impPromise = W.build({
      zipImpl: Z,
      sheets: [{
        name: '非循环项导入模板',
        header: res.imp.header,
        rows: impRows,
        cols: [14, 12, 18, 20, 28, 14, 42],
        headerHeight: 30,
        rowHeight: 20,
        headerStyle: 5,
        dataStyle: 2,
        validations: [{ sqref: 'E2:E' + Math.max(501, impRows.length + 1), formula: '_opt0!$A$1:$A$' + OPTIONS.length }]
      }, {
        name: '_opt0',
        hidden: true,
        rows: OPTIONS.map(function (o) { return [o]; }),
        dataStyle: 4
      }]
    });

    var wideRows = res.wide.rows;
    // 数量列（原值样式）：应休月休天数/应出勤天数/实际出勤天数/加班时长（小时）
    var qtySet = {};
    res.wide.header.forEach(function (h, i) { if (C.WIDE_QTY[h]) qtySet[i] = 1; });
    var widePromise = W.build({
      zipImpl: Z,
      sheets: [{
        name: '非循环项导入模板',
        header: res.wide.header,
        rows: wideRows,
        cols: [16, 12, 20, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 42],
        headerHeight: 20,
        rowHeight: 18,
        freeze: { x: 2, y: 1 },
        headerStyleFn: function (c) { return (c === 0 || c === 2) ? 16 : 15; },
        dataStyle: 17,
        styleFn: function (r, cc) {
          if (qtySet[cc]) return 17;
          var v = wideRows[r] ? wideRows[r][cc] : '';
          return (v !== '' && v != null) ? 18 : 17;
        }
      }]
    });

    return Promise.all([attPromise, catPromise, impPromise, widePromise]).then(function (arr) {
      return [
        { key: 'attendance', label: '考勤报送', base: '考勤报送.xlsx', bytes: arr[0] },
        { key: 'category', label: '工资类目一键表', base: '工资类目一键表.xlsx', bytes: arr[1] },
        { key: 'import', label: '非循环项导入', base: '非循环项导入.xlsx', bytes: arr[2] },
        { key: 'wide', label: '非循环项导入(宽表)', base: '非循环项导入(宽表).xlsx', bytes: arr[3] }
      ].map(function (f) {
        return {
          key: f.key, label: f.label, base: f.base,
          name: city + f.base,
          blob: new Blob([f.bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
        };
      });
    });
  }

  function problemsXlsx(problems) {
    var header = ['来源文件', '工作表', '行号', '员工姓名', '问题类型', '详细说明', '级别'];
    var LV = { error: '阻断', 'block-import': '阻断导入', warn: '提示', info: '信息' };
    var rows = problems.map(function (p) {
      return [p.file, p.sheet, p.row, p.name, p.type, p.detail, LV[p.level] || p.level];
    });
    return W.build({
      zipImpl: Z,
      sheets: [{
        name: '问题清单',
        header: header,
        rows: rows,
        cols: [30, 20, 8, 12, 16, 60, 12],
        headerHeight: 26,
        rowHeight: 20,
        freezeHeader: true,
        autoFilter: true,
        headerStyle: 5,
        dataStyle: 6
      }]
    });
  }

  function run() {
    if (!state.files.length) { toast('请先上传算薪表', 'warn'); return; }
    var cfg = state.cfg;
    // 首轮自动映射（只补空项）
    autoMapAll(false);

    var allProblems = [];
    var pre = [];
    state.files.forEach(function (f) {
      state.cfg.importCfg.map = state.cfg.importCfg.map || {};
      var res = C.processFile(XLSX, f.parsed, cfg, { importHeaders: IMP_HEADERS });
      res.city = f.city;
      res.parsedCity = f.parsed.city;
      pre.push(res);
      allProblems = allProblems.concat(res.problems);
      // 记录每列是否被用到（用于跨文件映射完整性）
    });

    var blocking = allProblems.filter(function (p) { return p.level === 'block-import'; });
    var fatal = allProblems.filter(function (p) { return p.level === 'error'; });
    if (fatal.length) {
      toast('有文件无法处理，请查看问题清单', 'warn');
    }
    if (blocking.length) {
      var names = [];
      blocking.forEach(function (p) {
        var m = String(p.detail || '').match(/「([^」]+)」/);
        var nm = m ? m[1] : String(p.detail || '').replace(/，.*$/, '');
        if (nm && names.indexOf(nm) < 0) names.push(nm);
      });
      var shown = names.slice(0, 10).map(function (n) { return '· ' + n; }).join('\n');
      var more = names.length > 10 ? '\n…等共 ' + names.length + ' 项' : '';
      var ok = window.confirm('以下薪酬项目没有映射到标准项目，导入模板会缺这些数据：\n\n' + shown + more +
        '\n\n请到「非循环项导入模板」卡片点开后，在映射列表中给这些项目选择标准项（未匹配的会标红）。\n\n要继续生成吗？');
      if (!ok) { showProblems(allProblems); return; }
    }

    // 宽表列未配齐：弹确认框列出，确认后仍可导出（未选列留空）
    var wideMissing = C.WIDE_ITEMS.filter(function (t) {
      var v = state.cfg.wideCfg.map[t];
      return v == null || v === '';
    });
    if (wideMissing.length) {
      var wShown = wideMissing.slice(0, 12).map(function (n) { return '· ' + n; }).join('\n');
      var wMore = wideMissing.length > 12 ? '\n…等共 ' + wideMissing.length + ' 列' : '';
      var wOk = window.confirm('宽表还有以下列没有选择取数来源，导出时这些列会留空：\n\n' + wShown + wMore +
        '\n\n可到「非循环项导入(宽表)」卡片里选择；要继续生成吗？');
      if (!wOk) { renderPanels(); openWidePanel(); return; }
    }

    var btn = $('#btnRun');
    btn.disabled = true;
    btn.textContent = '正在生成…';

    var jobs = pre.map(function (res) {
      return buildXlsxFiles(res, res.city).then(function (files) {
        return {
          city: res.city,
          file: res.file,
          stats: res.stats,
          problems: res.problems,
          files: files
        };
      });
    });

    Promise.all(jobs).then(function (results) {
      state.results = results;
      state.problems = allProblems;
      renderResults();
      toast('生成完成，共 ' + results.length + ' 个蝶城', 'ok');
    }).catch(function (err) {
      console.error(err);
      toast('生成失败：' + (err && err.message ? err.message : err), 'warn');
    }).then(function () {
      btn.disabled = false;
      btn.textContent = '开始生成';
    });
  }

  function renderResults() {
    var box = $('#results');
    if (!state.results.length) { box.innerHTML = ''; return; }
    var html = state.results.map(function (r, i) {
      var warnCount = r.problems.filter(function (p) { return p.level !== 'info'; }).length;
      return '' +
        '<div class="result-card" data-i="' + i + '">' +
        '  <div class="rc-head">' +
        '    <div><span class="rc-city">' + esc(r.city) + '</span>' +
        '      <span class="rc-file">' + esc(r.file.fileName) + '</span></div>' +
        '    <span class="rc-tag ' + (warnCount ? 'warn' : 'ok') + '">' + (warnCount ? warnCount + ' 条提示' : '无异常') + '</span>' +
        '  </div>' +
        '  <div class="rc-stats">' +
        '    <div><b>' + r.stats.employees + '</b><span>员工</span></div>' +
        '    <div><b>' + r.stats.attRows + '</b><span>考勤行</span></div>' +
        '    <div><b>' + r.stats.catRows + '</b><span>一键表行</span></div>' +
        '    <div><b>' + r.stats.impRows + '</b><span>导入行</span></div>' +
        '    <div><b>' + (r.stats.wideRows || 0) + '</b><span>宽表行</span></div>' +
        '    <div><b>' + fmtMoney(r.stats.amount) + '</b><span>金额合计</span></div>' +
        '  </div>' +
        '  <div class="rc-actions">' +
        '    <button class="btn" data-dl="0">考勤报送.xlsx</button>' +
        '    <button class="btn" data-dl="1">工资类目一键表.xlsx</button>' +
        '    <button class="btn" data-dl="2">非循环项导入.xlsx</button>' +
        '    <button class="btn" data-dl="3">非循环项导入(宽表).xlsx</button>' +
        '    <button class="btn btn-line" data-zip="1">打包 ZIP</button>' +
        '    <button class="btn btn-line" data-dir="1">保存到文件夹</button>' +
        '  </div>' +
        '</div>';
    }).join('');

    var pTotal = (state.problems || []).length;
    html += '<div class="result-foot">' +
      '<span>共 ' + state.results.length + ' 个蝶城 · ' + pTotal + ' 条问题记录</span>' +
      '<span class="foot-actions">' +
      (state.results.length > 1 ? '<button class="btn btn-line" id="btnAllDir">全部保存到文件夹</button>' : '') +
      (pTotal ? '<button class="btn btn-line" id="btnProblems">下载问题清单.xlsx</button>' : '') +
      '</span></div>';
    box.innerHTML = html;
  }

  function showProblems(problems) {
    state.problems = problems;
    if (!problems.length) return;
    renderResults();
  }

  /* ---------------- 导出 ---------------- */
  function downloadBlob(blob, name) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      URL.revokeObjectURL(url);
      a.remove();
    }, 1500);
  }

  function pickDirectory() {
    if (typeof window.showDirectoryPicker !== 'function') return Promise.resolve({ unsupported: true });
    return window.showDirectoryPicker({ mode: 'readwrite' })
      .then(function (d) { return { dir: d }; })
      .catch(function (e) {
        if (e && e.name === 'AbortError') return { cancelled: true };
        return { unsupported: true };
      });
  }

  function saveToFolder() {
    pickDirectory().then(function (res) {
      if (res.cancelled) return;
      if (res.unsupported || !res.dir) { fallbackDownloadAll(); return; }
      var dir = res.dir;
      var chain = Promise.resolve();
      state.results.forEach(function (r) {
        chain = chain.then(function () {
          var target = state.results.length > 1
            ? dir.getDirectoryHandle(r.city, { create: true })
            : Promise.resolve(dir);
          return target.then(function (d) {
            var inner = Promise.resolve();
            r.files.forEach(function (f) {
              inner = inner.then(function () {
                return d.getFileHandle(f.name, { create: true })
                  .then(function (fh) { return fh.createWritable(); })
                  .then(function (w) { return w.write(f.blob).then(function () { return w.close(); }); });
              });
            });
            return inner;
          });
        });
      });
      chain.then(function () {
        toast('已保存 ' + state.results.length + ' 个蝶城的报表', 'ok');
      }).catch(function (err) {
        console.error(err);
        toast('保存失败：' + (err && err.message ? err.message : err), 'warn');
      });
    });
  }

  function fallbackDownloadAll() {
    var i = 0;
    function next() {
      if (i >= state.results.length) { toast('浏览器不支持选择文件夹，已逐个下载到“下载”文件夹', 'warn'); return; }
      var r = state.results[i++];
      r.files.forEach(function (f, k) { setTimeout(function () { downloadBlob(f.blob, f.name); }, k * 250); });
      setTimeout(next, r.files.length * 260 + 300);
    }
    next();
  }

  function zipResult(r) {
    return Promise.all(r.files.map(function (f) {
      return f.blob.arrayBuffer().then(function (ab) { return { name: f.base, data: new Uint8Array(ab) }; });
    })).then(function (entries) {
      return Z.zip(entries);
    }).then(function (u8) {
      downloadBlob(new Blob([u8], { type: 'application/zip' }), r.city + '工资报送.zip');
    });
  }

  /* ---------------- 事件绑定 ---------------- */
  function bind() {
    var drop = $('#drop');
    var picker = $('#filePicker');

    drop.addEventListener('click', function () { picker.click(); });
    picker.addEventListener('change', function () { handleFiles(picker.files); picker.value = ''; });
    ['dragenter', 'dragover'].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('drag'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('drag'); });
    });
    drop.addEventListener('drop', function (e) {
      if (e.dataTransfer && e.dataTransfer.files) handleFiles(e.dataTransfer.files);
    });

    // 文件列表操作
    $('#fileList').addEventListener('click', function (e) {
      var btn = e.target.closest('[data-act="del"]');
      if (!btn) return;
      var id = Number(btn.closest('.file-item').dataset.id);
      state.files = state.files.filter(function (f) { return f.id !== id; });
      renderFiles(); renderPanels();
    });
    $('#fileList').addEventListener('change', function (e) {
      var el = e.target;
      var item = state.files.find(function (f) { return f.id === Number(el.closest('.file-item').dataset.id); });
      if (!item) return;
      var act = el.dataset.act;
      if (act === 'city') { item.city = el.value.trim() || item.parsed.city; }
      if (act === 'sheet') { item.parsed.sheetName = el.value; reparseFile(item); return; }
      if (act === 'hrow') { item.parsed.headerRow = Math.max(1, Number(el.value) || 1); reparseFile(item); return; }
      if (act === 'drow') { item.parsed.dataStartRow = Math.max(1, Number(el.value) || 1); reparseFile(item); return; }
      saveConfig();
    });

    // 关系配置瓦片
    $('#tiles').addEventListener('click', function (e) {
      var t = e.target.closest('.tile');
      if (!t) return;
      var key = t.dataset.panel;
      var panel = $('#panel-' + key);
      var isOpen = panel.classList.contains('open');
      $$('.panel').forEach(function (p) { p.classList.remove('open'); });
      $$('.tile').forEach(function (x) { x.classList.remove('active'); });
      if (!isOpen) { panel.classList.add('open'); t.classList.add('active'); }
    });

    // 考勤映射
    $('#panel-attendance').addEventListener('change', function (e) {
      var sel = e.target.closest('select[data-key]');
      if (!sel) return;
      state.cfg.attendance[sel.dataset.key] = sel.value;
      saveConfig();
    });

    // 区间配置
    $('#panel-category').addEventListener('change', function (e) {
      var sel = e.target.closest('select[data-range]');
      var ded = e.target.closest('select[data-deduct]');
      var fb = e.target.closest('select[data-fb]');
      var ig = e.target.closest('input[data-ig]');
      var dayAuto = e.target.closest('input[data-dayauto]');
      var daySel = e.target.closest('select[data-day]');
      if (dayAuto) {
        state.cfg.category.dayAuto = dayAuto.checked;
        saveConfig(); renderPanels();
      } else if (daySel) {
        state.cfg.category[daySel.dataset.day === 'start' ? 'dayStart' : 'dayEnd'] = daySel.value;
        saveConfig(); renderPanels();
      } else if (sel) {
        var key = sel.dataset.range, side = sel.dataset.side;
        state.cfg.category[key][side] = sel.value;
        saveConfig(); renderPanels();
      } else if (ded) {
        state.cfg.category.deduct = ded.value;
        saveConfig(); renderPanels();
      } else if (fb) {
        state.cfg.category[fb.dataset.fb === 'start' ? 'fbStart' : 'fbEnd'] = fb.value;
        saveConfig(); renderPanels();
      } else if (ig) {
        var arr = state.cfg.ignoreCols || (state.cfg.ignoreCols = []);
        var nm = ig.dataset.ig;
        var i = arr.indexOf(nm);
        if (ig.checked && i < 0) arr.push(nm);
        if (!ig.checked && i >= 0) arr.splice(i, 1);
        saveConfig();
      }
    });
    // 忽略列搜索：本地过滤，不重渲染
    $('#panel-category').addEventListener('input', function (e) {
      if (e.target.id !== 'igSearch') return;
      var q = e.target.value.trim();
      $$('#panel-category .ig-item').forEach(function (lab) {
        var n = lab.dataset.name || '';
        lab.style.display = (!q || n.indexOf(q) >= 0) ? '' : 'none';
      });
    });

    // 导入模板参数与映射
    $('#panel-import').addEventListener('change', function (e) {
      var t = e.target;
      if (t.id === 'impTask') { state.cfg.importCfg.task = t.value.trim(); saveConfig(); return; }
      if (t.id === 'impDate') { state.cfg.importCfg.date = t.value; saveConfig(); return; }
      if (t.classList.contains('opt-input')) {
        var src = t.dataset.src;
        var v = t.value.trim();
        var val = '', conf = 'none';
        if (v === '不报送' || v === '__skip__') { val = '__skip__'; conf = 'skip'; }
        else if (!v) { val = ''; conf = 'none'; }
        else if (OPTIONS.indexOf(v) >= 0) { val = v; conf = 'manual'; }
        else {
          var m = C.matchOption(v, OPTIONS);
          if (m.value) { val = m.value; conf = 'guess'; t.value = m.value; }
          else { val = ''; conf = 'none'; toast('“' + v + '”不在标准项目里，请从下拉中选', 'warn'); t.value = ''; }
        }
        state.cfg.importCfg.map[src] = { value: val, conf: conf };
        saveConfig();
        renderImportPanel();
      }
    });
    $('#panel-import').addEventListener('click', function (e) {
      if (e.target.id === 'btnAutoMap') { autoMapAll(true); toast('已按表头重新自动匹配', 'ok'); }
    });

    // 宽表参数与映射
    $('#panel-wide').addEventListener('change', function (e) {
      var t = e.target;
      if (t.id === 'wideDate') { state.cfg.wideCfg.date = t.value; saveConfig(); return; }
      if (t.dataset && t.dataset.wide !== undefined && t.tagName === 'SELECT') {
        state.cfg.wideCfg.map[t.dataset.wide] = t.value;
        saveConfig();
        renderWidePanel();
      }
    });
    $('#panel-wide').addEventListener('click', function (e) {
      if (e.target.id === 'btnWideAutoMap') { autoMapWide(true); toast('已按表头重新自动匹配', 'ok'); }
    });

    // 生成
    $('#btnRun').addEventListener('click', run);

    // 结果操作
    $('#results').addEventListener('click', function (e) {
      var card = e.target.closest('.result-card');
      if (card) {
        var r = state.results[Number(card.dataset.i)];
        if (!r) return;
        var dl = e.target.closest('[data-dl]');
        if (dl) { var f = r.files[Number(dl.dataset.dl)]; if (f) downloadBlob(f.blob, f.name); return; }
        if (e.target.closest('[data-zip]')) { zipResult(r); return; }
        if (e.target.closest('[data-dir]')) {
          saveToFolder();
          return;
        }
      }
      if (e.target.id === 'btnProblems') {
        problemsXlsx(state.problems || []).then(function (u8) {
          downloadBlob(new Blob([u8], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), '问题清单.xlsx');
        });
      }
      if (e.target.id === 'btnAllDir') saveToFolder();
    });

    // 配置导入导出（侧栏底部）——覆盖全部四个模块的设置
    function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
    function lsJson(k) { try { return JSON.parse(lsGet(k) || 'null'); } catch (e) { return null; } }

    // 用户手册：用模板生成独立网页，新标签页打开（Blob 方式，单文件照样离线可用）
    $('#btnHelp').addEventListener('click', function () {
      var tpl = document.getElementById('helpTpl');
      if (!tpl) { toast('手册模板缺失，请使用完整版本的工具', 'warn'); return; }
      var css = [
        '*{box-sizing:border-box}',
        'body{margin:0;background:#eef2ef;font-family:"Microsoft YaHei","PingFang SC","Segoe UI",sans-serif;color:#26313c;padding:28px 16px}',
        '.help-wrap{max-width:880px;margin:0 auto;background:#fff;border-radius:14px;box-shadow:0 10px 30px rgba(0,0,0,.08);padding:30px 36px 36px}',
        '.help-h{display:flex;align-items:center;gap:10px;border-bottom:2px solid #217346;padding-bottom:12px;margin-bottom:8px}',
        '.help-h .ico{font-size:22px}',
        '.help-h h1{margin:0;font-size:20px;color:#217346}',
        '.help-h .sub{font-size:12.5px;color:#8ba094;margin-left:auto}',
        '.help-body h4{margin:18px 0 8px;font-size:15px;color:#217346;border-left:3px solid #217346;padding-left:8px}',
        '.help-body p{margin:6px 0;font-size:13.5px;color:#3c4854;line-height:1.8}',
        '.help-body ol{margin:6px 0 10px;padding-left:22px;font-size:13.5px;color:#3c4854;line-height:1.9}',
        '.help-body .help-tip{background:#fff6e5;border:1px solid #f0dfb8;border-radius:8px;padding:8px 12px;font-size:13px;color:#7a5b16;margin:8px 0;line-height:1.8}',
        '.help-body b{color:#26313c}'
      ].join('\n');
      var html = '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
        '<title>用户手册 · 人力工具集成平台</title><style>' + css + '</style></head><body>' +
        tpl.innerHTML + '</body></html>';
      var url = URL.createObjectURL(new Blob([html], { type: 'text/html;charset=utf-8' }));
      var w = window.open(url, '_blank');
      if (!w) toast('浏览器拦截了新标签页，请允许本页弹窗后重试', 'warn');
      setTimeout(function () { URL.revokeObjectURL(url); }, 60000);
    });

    $('#btnExportCfg').addEventListener('click', function () {
      var all = {
        _app: 'humanToolsCfg', _ver: 1,
        salary: state.cfg,
        attendance: lsJson('salaryTool.att.v1'),
        idle: { excl: lsGet('idleExcl') },
        rename: {
          missingOn: lsGet('renMissingOn'),
          missingChk: lsJson('renMissingChk'),
          tplSegs: lsJson('renTplSegs'),
          ext: lsGet('renExt'),
          sep: lsGet('renSep')
        }
      };
      var blob = new Blob([JSON.stringify(all, null, 2)], { type: 'application/json' });
      downloadBlob(blob, '人力工具配置.json');
      toast('已导出全部模块配置', 'ok');
    });
    $('#btnImportCfg').addEventListener('click', function () { $('#cfgPicker').click(); });
    $('#cfgPicker').addEventListener('change', function (e) {
      var f = e.target.files[0];
      if (!f) return;
      var reader = new FileReader();
      reader.onload = function (ev) {
        try {
          var all = JSON.parse(ev.target.result);
          if (!all || all._app !== 'humanToolsCfg') { toast('配置文件格式不正确', 'warn'); return; }
          // 薪资报送
          if (all.salary) {
            state.cfg = deepMerge(C.DEFAULT_CONFIG, all.salary);
            saveConfig();
            renderPanels();
          }
          // 考勤处理
          if (all.attendance) {
            try { localStorage.setItem('salaryTool.att.v1', JSON.stringify(all.attendance)); } catch (err) { }
          }
          // 超7天未出勤
          if (all.idle) {
            try { localStorage.setItem('idleExcl', all.idle.excl || ''); } catch (err) { }
            var excl = $('#idleExcl');
            if (excl && all.idle.excl != null) excl.value = all.idle.excl;
          }
          // 文件重命名
          if (all.rename) {
            try {
              if (all.rename.missingOn != null) localStorage.setItem('renMissingOn', all.rename.missingOn);
              if (all.rename.missingChk) localStorage.setItem('renMissingChk', JSON.stringify(all.rename.missingChk));
              if (all.rename.tplSegs) localStorage.setItem('renTplSegs', JSON.stringify(all.rename.tplSegs));
              if (all.rename.ext != null) localStorage.setItem('renExt', all.rename.ext);
              if (all.rename.sep != null) localStorage.setItem('renSep', all.rename.sep);
            } catch (err) { }
          }
          // 各模块刷新
          if (window.RenApp) window.RenApp.onShow();
          if (window.AttApp) window.AttApp.onShow();
          if (window.IdleApp) window.IdleApp.onShow();
          toast('配置已导入（含全部模块设置）', 'ok');
        } catch (err) { toast('配置文件格式不正确', 'warn'); }
      };
      reader.readAsText(f);
      e.target.value = '';
    });

    // 侧边导航
    $$('.nav-item').forEach(function (n) {
      n.addEventListener('click', function () {
        $$('.nav-item').forEach(function (x) { x.classList.remove('active'); });
        n.classList.add('active');
        var view = n.dataset.view;
        var TITLES = { salary: '薪资报送', attendance: '考勤处理', absent: '超7天未出勤', rename: '文件重命名' };
        $('#topTitle').textContent = TITLES[view] || n.textContent.trim();
        $('#viewSalary').classList.toggle('hidden', view !== 'salary');
        $('#viewAttendance').classList.toggle('hidden', view !== 'attendance');
        $('#viewAbsent').classList.toggle('hidden', view !== 'absent');
        $('#viewRename').classList.toggle('hidden', view !== 'rename');
        /* 顶栏操作按钮按视图显隐 */
        $$('[data-topview]').forEach(function (b) { b.classList.toggle('hidden', b.dataset.topview !== view); });
        if (view === 'attendance' && window.AttApp) window.AttApp.onShow();
        if (view === 'absent' && window.IdleApp) window.IdleApp.onShow();
        if (view === 'rename' && window.RenApp) window.RenApp.onShow();
      });
    });
  }

  /* ---------------- 初始化 ---------------- */
  document.addEventListener('DOMContentLoaded', function () {
    var dl = document.createElement('datalist');
    dl.id = 'optList';
    dl.innerHTML = OPTIONS.map(function (o) { return '<option value="' + esc(o) + '">'; }).join('') +
      '<option value="不报送">';
    document.body.appendChild(dl);

    bind();
    renderFiles();
    renderPanels();
  });
})();
