/* ============================================================
 * ren-app.js — 文件重命名界面
 * 布局（用户图）：花名册大窗置顶 + 右侧「开始重命名」「问题清单明细」按钮
 *   资料模块一行 2 个（10 个窗口），支持多文件或 zip 解包
 * 改名模板可自定义（8 个变量）；缺失检查勾选区默认不勾 + localStorage 记忆
 * 结果：摘要 + 导出 zip（平铺）；问题清单弹窗（原文件名/失败原因/改正建议 + xlsx）
 * ============================================================ */
(function () {
  'use strict';

  var XLSX = window.XLSX;
  var R = window.RenCore;
  var W = window.XlsxWriter;
  var Z = window.ZipUtil;

  var state = {
    roster: null,
    rosterMeta: null,     // {sheet, aoa, hIdx, cols, headers, fileName} 字段映射元数据
    files: [],            // {module, name, data}
    ran: false,
    res: null,
    chk: {},              // module -> bool（缺失检查勾选）
    tplCfg: null          // {segs:[{t,v}], ext} 分段改名模板
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
  function stamp() {
    var n = new Date();
    function p2(x) { return x < 10 ? '0' + x : '' + x; }
    return '' + n.getFullYear() + p2(n.getMonth() + 1) + p2(n.getDate()) + p2(n.getHours()) + p2(n.getMinutes());
  }
  function readAsArrayBuffer(file) {
    return new Promise(function (res, rej) {
      var r = new FileReader();
      r.onload = function () { res(r.result); };
      r.onerror = function () { rej(new Error('读取失败')); };
      r.readAsArrayBuffer(file);
    });
  }

  /* ---------------- 花名册区（含字段映射） ---------------- */
  var ROSTER_FIELDS = [['id', '工号'], ['name', '姓名'], ['center', '蝶城'], ['project', '项目']];

  function renderRoster() {
    var box = $('#renRosterBox');
    if (!box) return;
    var d = state.roster;
    var info = !d
      ? ''
      : d.error
        ? '<div class="file-item"><div class="fi-head"><div class="fi-name"><span class="fi-dot" style="background:var(--red)"></span>' +
          '<span class="fi-nm">' + esc(d.fileName || '') + '</span></div><button class="btn-ghost" data-ren-clear-roster="1">移除</button></div>' +
          '<div class="att-file-meta" style="color:var(--red)">⚠ ' + esc(d.error) + '</div></div>'
        : '<div class="file-item"><div class="fi-head"><div class="fi-name"><span class="fi-dot" style="background:var(--green)"></span>' +
          '<span class="fi-nm">' + esc(d.fileName) + '</span></div><button class="btn-ghost" data-ren-clear-roster="1">移除</button></div>' +
          '<div class="att-file-meta">识别到 ' + d.rows.length + ' 名员工 · 工作表「' + esc(d.sheet) + '」</div></div>';
    var mapHtml = renderRosterMap();
    box.innerHTML = '<div class="drop drop-wide" id="renDrop-roster">' +
      '<div class="d-main">人员花名册上传</div>' +
      '<div class="d-sub">必填 · 含 调度中心 / 员工项目 / 员工编码 / 员工姓名</div>' +
      '<div class="d-sub">拖入人员花名册 Excel（.xlsx / .xls / .xlsm），或点击选择</div>' +
      '<input type="file" id="renPick-roster" accept=".xlsx,.xls,.xlsm" class="hidden"></div>' +
      (info ? '<div class="file-list">' + info + '</div>' : '') +
      mapHtml;

    var drop = $('#renDrop-roster'), picker = $('#renPick-roster');
    drop.addEventListener('click', function (e) { if (!e.target.closest('input')) picker.click(); });
    picker.addEventListener('change', function () {
      if (picker.files.length) handleRoster(picker.files[0]);
      picker.value = '';
    });
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('drag'); });
    drop.addEventListener('dragleave', function () { drop.classList.remove('drag'); });
    drop.addEventListener('drop', function (e) {
      e.preventDefault(); drop.classList.remove('drag');
      if (e.dataTransfer.files.length) handleRoster(e.dataTransfer.files[0]);
    });
    var clr = $('[data-ren-clear-roster]');
    if (clr) clr.addEventListener('click', function () {
      state.roster = null; state.rosterMeta = null; state.ran = false; state.res = null;
      renderRoster(); renderResults();
    });
    $$('[data-ren-map]').forEach(function (sel) {
      sel.addEventListener('change', function () {
        var meta = state.rosterMeta;
        if (!meta) return;
        meta.cols[sel.dataset.renMap] = parseInt(sel.value, 10);
        var r = R.extractRoster(meta, meta.cols);
        if (r.error) {
          state.roster = { error: r.error, fileName: meta.fileName };
          toast(r.error, 'warn');
        } else {
          state.roster = { sheet: r.sheet, rows: r.rows, fileName: meta.fileName };
          toast('已按新映射读取到 ' + r.rows.length + ' 名员工', 'ok');
        }
        state.ran = false; state.res = null;
        renderRoster(); renderResults();
      });
    });
    updateTplPreview();
  }

  /** 字段映射下拉：自动猜 + 可改（工号/姓名/蝶城/项目） */
  function renderRosterMap() {
    var meta = state.rosterMeta;
    if (!meta || !meta.headers || !meta.headers.length) return '';
    var items = ROSTER_FIELDS.map(function (f) {
      var cur = meta.cols[f[0]];
      var opts = '<option value="-1"' + (cur < 0 ? ' selected' : '') + '>— 不映射 —</option>';
      meta.headers.forEach(function (h, i) {
        if (h === '') return;
        opts += '<option value="' + i + '"' + (cur === i ? ' selected' : '') + '>' + esc(h) + '</option>';
      });
      return '<label class="ren-map-it"><span class="ren-map-lb">' + f[1] + '</span>' +
        '<select data-ren-map="' + f[0] + '">' + opts + '</select></label>';
    }).join('');
    return '<div class="ren-map">' + items + '</div>';
  }

  function handleRoster(file) {
    if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) { toast('花名册请选择 .xlsx / .xls 文件', 'warn'); return; }
    readAsArrayBuffer(file).then(function (buf) {
      var wb;
      try { wb = XLSX.read(new Uint8Array(buf), { type: 'array' }); }
      catch (e) { toast('解析失败：' + file.name, 'warn'); return; }
      var meta = R.parseRosterMeta(wb);
      state.ran = false; state.res = null;
      if (meta.error) {
        state.rosterMeta = null;
        state.roster = { error: meta.error, fileName: file.name };
        renderRoster(); renderResults();
        toast(file.name + '：' + meta.error, 'warn');
        return;
      }
      meta.fileName = file.name;
      state.rosterMeta = meta;
      var r = R.extractRoster(meta, meta.cols);
      if (r.error) {
        state.roster = { error: r.error, fileName: file.name };
        toast(file.name + '：' + r.error, 'warn');
      } else {
        state.roster = { sheet: r.sheet, rows: r.rows, fileName: file.name };
        toast('花名册识别到 ' + r.rows.length + ' 名员工', 'ok');
      }
      renderRoster(); renderResults();
    });
  }

  /* ---------------- 资料模块区 ---------------- */
  function renderGrid() {
    var box = $('#renGrid');
    if (!box) return;
    box.innerHTML = R.MODULES.map(function (m) {
      var n = state.files.filter(function (f) { return f.module === m; }).length;
      var info = n
        ? '<div class="file-list"><div class="file-item"><div class="fi-head"><div class="fi-name"><span class="fi-dot" style="background:var(--green)"></span>' +
          '<span class="fi-nm">' + m + ' · ' + n + ' 个文件</span></div>' +
          '<button class="btn-ghost" data-ren-clear-mod="' + esc(m) + '">清空</button></div></div></div>'
        : '';
      return '<div class="att-zone"><div class="drop" id="renDrop-mod" data-mod="' + esc(m) + '">' +
        '<div class="d-main">' + esc(m) + '</div><div class="d-sub">拖拽或点击选择</div>' +
        '<input type="file" id="renPick-mod" data-mod="' + esc(m) + '" multiple class="hidden"></div>' + info + '</div>';
    }).join('');

    $$('[data-ren-clear-mod]').forEach(function (b) {
      b.addEventListener('click', function () {
        var m = b.dataset.renClearMod;
        state.files = state.files.filter(function (f) { return f.module !== m; });
        state.ran = false; state.res = null;
        renderGrid(); renderResults();
      });
    });
    $$('#renGrid .drop').forEach(function (drop) {
      var m = drop.dataset.mod;
      var picker = drop.querySelector('input');
      drop.addEventListener('click', function (e) { if (!e.target.closest('input')) picker.click(); });
      picker.addEventListener('change', function () {
        if (picker.files.length) handleModuleFiles(m, Array.prototype.slice.call(picker.files));
        picker.value = '';
      });
      drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('drag'); });
      drop.addEventListener('dragleave', function () { drop.classList.remove('drag'); });
      drop.addEventListener('drop', function (e) {
        e.preventDefault(); drop.classList.remove('drag');
        if (e.dataTransfer.files.length) handleModuleFiles(m, Array.prototype.slice.call(e.dataTransfer.files));
      });
    });
  }

  function handleModuleFiles(module, files) {
    var chain = Promise.resolve();
    var added = 0, errs = [];
    files.forEach(function (f) {
      chain = chain.then(function () {
        return readAsArrayBuffer(f).then(function (buf) {
          var u8 = new Uint8Array(buf);
          if (/\.zip$/i.test(f.name)) {
            return Z.unzip(u8).then(function (entries) {
              entries.forEach(function (e) {
                var base = e.name.split('/').pop();
                if (!base || base.charAt(0) === '.') return;
                state.files.push({ module: module, name: base, data: e.data });
                added++;
              });
            }).catch(function (err) {
              errs.push(f.name + '：' + err.message);
            });
          }
          if (/\.(xlsx|xlsm|xls)$/i.test(f.name) && !window.confirm('「' + f.name + '」是 Excel 文件，确定要作为「' + module + '」资料重命名吗？')) return;
          state.files.push({ module: module, name: f.name, data: u8 });
          added++;
        });
      });
    });
    chain.then(function () {
      state.ran = false; state.res = null;
      renderGrid(); renderResults();
      if (errs.length) toast(errs[0], 'warn');
      else if (added) toast(module + '：已加入 ' + added + ' 个文件', 'ok');
    });
  }

  /* ---------------- 缺失检查（开关 + 勾选区） ---------------- */
  function renderChk() {
    var box = $('#renChk');
    if (!box) return;
    var saved = {};
    try { saved = JSON.parse(localStorage.getItem('renMissingChk') || '{}'); } catch (e) { }
    var on = false;
    try { on = localStorage.getItem('renMissingOn') === '1'; } catch (e) { }
    state.missingOn = on;
    var list = R.MODULES.filter(function (m) { return m !== R.DUP_MODULE; }).map(function (m) {
      var chk = saved[m] === true;
      state.chk[m] = chk;
      return '<label class="idle-chk"><input type="checkbox" data-ren-chk="' + esc(m) + '"' + (chk ? ' checked' : '') + '>' + esc(m) + '</label>';
    }).join('');
    box.innerHTML = '<label class="ren-switch"><input type="checkbox" id="renMissingOn"' + (on ? ' checked' : '') + '>' +
      '<span class="rs-track"></span><span class="rs-lb">缺失检查</span></label>' +
      '<span class="ren-chk-all"><button type="button" class="btn-mini" id="renChkAll">全选</button>' +
      '<button type="button" class="btn-mini" id="renChkNone">全不选</button></span>' +
      '<div class="ren-chk-list' + (on ? '' : ' hidden') + '" id="renChkList">' + list + '</div>';
    var sw = $('#renMissingOn');
    if (sw) sw.addEventListener('change', function () {
      state.missingOn = sw.checked;
      try { localStorage.setItem('renMissingOn', sw.checked ? '1' : '0'); } catch (e) { }
      var l = $('#renChkList');
      if (l) l.classList.toggle('hidden', !sw.checked);
    });
    var bAll = $('#renChkAll'), bNone = $('#renChkNone');
    function setAll(v) {
      $$('[data-ren-chk]').forEach(function (c) {
        c.checked = v;
        state.chk[c.dataset.renChk] = v;
      });
      saveChk();
    }
    if (bAll) bAll.addEventListener('click', function () { setAll(true); });
    if (bNone) bNone.addEventListener('click', function () { setAll(false); });
    $$('[data-ren-chk]').forEach(function (c) {
      c.addEventListener('change', function () {
        state.chk[c.dataset.renChk] = c.checked;
        saveChk();
      });
    });
  }
  function saveChk() {
    try { localStorage.setItem('renMissingChk', JSON.stringify(state.chk)); } catch (e) { }
  }

  /* ---------------- 改名模板（分段式：自定义段在前，工号-类型 固定尾段在后） ----------------
   * 上传系统规则：工号-类型 必须连接在一起且固定在文件名最后（后缀之前），
   * 所以「工号」「类型」是锁定尾段不可编辑，新增/自定义段都排在工号前面 */
  var TPL_VARS = ['姓名', '项目', '调度中心', '日期', '序号'];
  var TPL_DEFAULT = [{ t: 'var', v: '姓名' }];

  function loadTplCfg() {
    var segs = TPL_DEFAULT, ext = '', sep = '-';
    try {
      var s = JSON.parse(localStorage.getItem('renTplSegs') || 'null');
      if (Array.isArray(s) && s.length) segs = s.filter(function (x) { return x && (x.t === 'text' ? true : TPL_VARS.indexOf(x.v) >= 0); });
      ext = String(localStorage.getItem('renExt') || '').trim();
      sep = String(localStorage.getItem('renSep') || '').trim();
    } catch (e) { }
    if (!segs.length) segs = TPL_DEFAULT.slice();
    if (!sep) sep = '-';
    return { segs: segs, ext: ext, sep: sep };
  }
  function saveTplCfg(cfg) {
    try {
      localStorage.setItem('renTplSegs', JSON.stringify(cfg.segs));
      localStorage.setItem('renExt', cfg.ext);
      localStorage.setItem('renSep', cfg.sep);
    } catch (e) { }
  }

  function renderTplBuilder() {
    var box = $('#renTplSegs');
    if (!box) return;
    var cfg = state.tplCfg;
    var sepH = '<span class="ren-seg-c">' + esc(cfg.sep || '-') + '</span>';
    var html = cfg.segs.map(function (s, i) {
      var isText = s.t === 'text';
      var opts = TPL_VARS.map(function (v) {
        return '<option value="' + esc(v) + '"' + (!isText && s.v === v ? ' selected' : '') + '>' + esc(v) + '</option>';
      }).join('');
      var seg = '<span class="ren-seg" data-idx="' + i + '">' +
        '<span class="ren-seg-wrap">' +
        '<select class="ren-seg-sel">' + '<option value="__text"' + (isText ? ' selected' : '') + '>固定文字</option>' + opts + '</select>' +
        (isText ? '<input type="text" class="ren-seg-txt" value="' + esc(s.v || '') + '" placeholder="填固定文字">' : '') +
        '<button type="button" class="ren-del" title="删除此段">×</button>' +
        '</span></span>';
      return (i ? sepH : '') + seg;
    }).join('');
    if (!cfg.segs.length) html = '<span class="ren-seg-empty">暂无自定义段，</span>';
    // 固定尾段：工号-类型 连接在一起、固定在最后（后缀之前），锁定不可编辑
    html += (cfg.segs.length ? sepH : '') +
      '<span class="ren-seg-lock" title="固定尾段，不可修改">工号</span>' + sepH +
      '<span class="ren-seg-lock" title="固定尾段，不可修改">类型</span>';
    box.innerHTML = html;
    var selAll = $$('.ren-seg-sel', box);
    function refreshRow(idx, val) {
      var seg = state.tplCfg.segs[idx];
      if (val === '__text') { seg.t = 'text'; seg.v = ''; }
      else { seg.t = 'var'; seg.v = val; }
      renderTplBuilder();
      saveTplCfg(state.tplCfg);
    }
    selAll.forEach(function (sel, i) {
      sel.addEventListener('change', function () { refreshRow(i, sel.value); });
    });
    $$('.ren-seg-txt', box).forEach(function (inp, i) {
      inp.addEventListener('input', function () {
        state.tplCfg.segs[i].t = 'text';
        state.tplCfg.segs[i].v = inp.value;
        saveTplCfg(state.tplCfg);
      });
    });
    $$('.ren-del', box).forEach(function (b, i) {
      b.addEventListener('click', function () {
        state.tplCfg.segs.splice(i, 1);
        renderTplBuilder();
        saveTplCfg(state.tplCfg);
      });
    });
    updateTplPreview();
  }

  var addBinded = false;
  function bindAddBtn() {
    if (addBinded) return;
    addBinded = true;
    var add = $('#renSegAdd');
    if (add) add.addEventListener('click', function () {
      state.tplCfg.segs.push({ t: 'var', v: '姓名' });
      renderTplBuilder();
      saveTplCfg(state.tplCfg);
    });
  }

  /** 完整规则真实示例预览：花名册第一人渲染；无花名册用示例值 */
  function updateTplPreview() {
    var el = $('#renTplPreview');
    if (!el) return;
    var emp, src = '示例';
    if (state.roster && !state.roster.error && state.roster.rows && state.roster.rows.length) {
      emp = state.roster.rows[0];
      src = state.roster.rows[0].name + '(' + state.roster.rows[0].id + ')';
    } else {
      emp = { id: 'R001', name: '张三', project: '项目甲', center: '甬江蝶城' };
    }
    if (!emp.seq) emp.seq = 1;
    var n = new Date();
    var dateStr = '' + n.getFullYear() + R.pad2(n.getMonth() + 1) + R.pad2(n.getDate());
    var t = R.buildTarget({ segs: state.tplCfg.segs, ext: state.tplCfg.ext || 'jpg', sep: state.tplCfg.sep }, emp, R.MODULES[0], '', dateStr);
    el.textContent = t ? (t + '（按 ' + src + ' 渲染）') : '（当前规则为空，请添加有效段）';
  }

  function bindExtInput() {
    var ext = $('#renExt');
    if (!ext) return;
    ext.value = state.tplCfg.ext;
    ext.addEventListener('input', function () {
      state.tplCfg.ext = ext.value.trim();
      saveTplCfg(state.tplCfg);
    });
    var sep = $('#renSep');
    if (!sep) return;
    sep.value = state.tplCfg.sep;
    sep.addEventListener('input', function () {
      var v = sep.value.trim();
      if (!v) v = '-';
      state.tplCfg.sep = v;
      renderTplBuilder();
      saveTplCfg(state.tplCfg);
    });
  }

  /* ---------------- 开始重命名 ---------------- */
  function readTpl() {
    return state.tplCfg || { segs: TPL_DEFAULT.slice(), ext: '', sep: '-' };
  }

  function doRun() {
    if (!state.roster || state.roster.error || !state.roster.rows) { toast('先上传人员花名册', 'warn'); return; }
    if (!state.files.length) { toast('先在下面上传资料文件', 'warn'); return; }
    var tpl = readTpl();
    var n = new Date();
    var dateStr = '' + n.getFullYear() + R.pad2(n.getMonth() + 1) + R.pad2(n.getDate());
    var res = R.matchAll(state.files, state.roster.rows, tpl, dateStr);
    if (res.error) {
      state.ran = false; state.res = null;
      renderResults();
      toast(res.error, 'warn');
      return;
    }
    var checked = state.missingOn ? R.MODULES.filter(function (m) { return state.chk[m]; }) : [];
    if (checked.length) res.fails = res.fails.concat(R.findMissing(res, state.roster.rows, checked));
    res.tpl = tpl;
    state.res = res;
    state.ran = true;
    renderResults();
    toast('重命名完成：成功 ' + res.ok.length + ' 个文件，问题 ' + res.fails.length + ' 条', res.fails.length ? 'warn' : 'ok');
  }

  /* ---------------- 结果渲染 ---------------- */
  function renderResults() {
    var box = $('#renResults');
    if (!box) return;
    var btnP = $('#btnRenProblems');
    if (btnP) btnP.disabled = !state.ran;
    var r = state.res;
    if (!r || !state.ran) { box.innerHTML = ''; return; }
    var failRows = r.fails;
    var pk = packableFiles(r.ok);
    var nPacks = pk.files.length ? planPacks(pk.files).length : 0;
    var sizeInfo = pk.files.length
      ? '共 ' + pk.files.length + ' 个文件 · 合计 ' + fmtSize(pk.total) +
        (nPacks > 1 ? '，超过单包 300MB 限制，导出时预计拆分成 ' + nPacks + ' 个压缩包' : '')
      : '';
    box.innerHTML = '<div class="result-card"><div class="ren-res-summary">' +
      '<span>花名册 <b>' + (state.roster ? state.roster.rows.length : 0) + '</b> 人 · 成功重命名 <b class="ren-ok">' + r.ok.length +
      '</b> 个文件 · 覆盖 <b>' + r.empCount + '</b> 人 · 问题 <b class="ren-bad">' + failRows.length + '</b> 条</span>' +
      '<button class="btn btn-lg" id="btnRenZip">导出 zip</button></div>' +
      (failRows.length
        ? '<div class="ren-res-warn">有 ' + failRows.length + ' 条问题没有重命名，点上方「问题清单明细」查看原因与改正建议；导出的 zip 不含这些文件。' + (sizeInfo ? '<br>' + sizeInfo + '。</div>' : '</div>')
        : '<div class="ren-res-warn ok">全部文件重命名成功，可直接导出 zip（' + sizeInfo + '）。</div>') +
      '</div>';
    var z = $('#btnRenZip');
    if (z) z.addEventListener('click', doExportZip);
  }

  /* ---------------- 导出 zip（单附件 50MB 校验 + 按模块分组自动拆包） ---------------- */
  var MAX_FILE = 50 * 1024 * 1024;        // 单个附件 50MB（上传系统限制）
  var MAX_TOTAL = 300 * 1024 * 1024;      // 单包解压后合计 300MB（用户实测 500MB 会上传超时，收紧防超时）
  var MAX_COUNT = 500;                    // 单包附件数（上传系统限制）

  function fmtSize(b) {
    if (b >= 1024 * 1024) return (b / 1024 / 1024).toFixed(1).replace(/\.0$/, '') + 'MB';
    if (b >= 1024) return Math.round(b / 1024) + 'KB';
    return b + 'B';
  }
  /** 可打包文件（剔除超 50MB 的）+ 总大小 */
  function packableFiles(ok) {
    var files = [], total = 0;
    (ok || []).forEach(function (x) {
      if (x.data && x.data.length > MAX_FILE) return;
      files.push(x);
      total += x.data ? x.data.length : 0;
    });
    return { files: files, total: total };
  }
  /** 按模块分组装箱：每包 ≤500 个文件且解压后 ≤300MB */
  function planPacks(files) {
    var byMod = {};
    files.forEach(function (x) { (byMod[x.module] = byMod[x.module] || []).push(x); });
    var packs = [], cur = null;
    function newPack() { cur = { list: [], size: 0 }; packs.push(cur); }
    newPack();
    R.MODULES.forEach(function (m) {
      var arr = byMod[m];
      if (!arr || !arr.length) return;
      if (cur.list.length && (cur.list.length + arr.length > MAX_COUNT || cur.size + sumSize(arr) > MAX_TOTAL)) newPack();
      arr.forEach(function (x) {
        if (cur.list.length >= MAX_COUNT || cur.size + x.data.length > MAX_TOTAL) newPack();
        cur.list.push(x);
        cur.size += x.data.length;
      });
    });
    return packs;
  }

  function doExportZip() {
    var r = state.res;
    if (!r || !r.ok.length) { toast('没有可导出的文件', 'warn'); return; }
    var big = [], files = [];
    r.ok.forEach(function (x) {
      if (x.data && x.data.length > MAX_FILE) {
        big.push(x);
        r.fails.push({ file: x.file, module: x.module, reason: '文件大小 ' + (x.data.length / 1024 / 1024).toFixed(1) + 'MB，超过单附件 50MB 限制', advice: '压缩或拆分该文件后重新上传' });
      } else files.push(x);
    });
    if (big.length) {
      state.res = r;
      renderResults();
      openModal();
    }
    if (!files.length) { toast('没有可打包的文件（全部超过 50MB）', 'warn'); return; }

    var packs = planPacks(files);
    var totalSize = sumSize(files);

    var zbtn = $('#btnRenZip');
    var ts = stamp();
    var single = packs.length === 1;
    var chain = Promise.resolve();
    var packed = 0, packedBytes = 0;
    if (zbtn) { zbtn.disabled = true; }
    packs.forEach(function (p, i) {
      chain = chain.then(function () {
        var name = single ? '重命名结果-' + ts + '.zip' : '重命名结果-' + (i + 1) + '-' + ts + '.zip';
        toast('正在打包第 ' + (i + 1) + '/' + packs.length + ' 个压缩包…', 'info');
        function tick(done, total) {
          if (!zbtn) return;
          var pct = total ? Math.round((packed + done) / files.length * 100) : 0;
          zbtn.textContent = '正在打包 ' + pct + '% (' + (packed + done) + '/' + files.length + ')';
        }
        tick(0, 1);
        return Z.zip(p.list.map(function (x) { return { name: x.target, data: x.data }; }), tick).then(function (u8) {
          downloadBlob(new Blob([u8], { type: 'application/zip' }), name);
          packed += p.list.length;
          packedBytes += u8.length;
        });
      });
    });
    chain.then(function () {
      if (zbtn) { zbtn.disabled = false; zbtn.textContent = '导出 zip'; }
      toast('已导出 ' + files.length + ' 个文件（' + fmtSize(totalSize) +
        (single ? '' : ' · 已拆分成 ' + packs.length + ' 个压缩包') + '）', 'ok');
    }).catch(function (e) {
      if (zbtn) { zbtn.disabled = false; zbtn.textContent = '导出 zip'; }
      toast('打包失败：' + (e && e.message || e), 'warn');
    });
  }
  function sumSize(arr) {
    var s = 0;
    arr.forEach(function (x) { s += x.data ? x.data.length : 0; });
    return s;
  }

  /* ---------------- 问题清单弹窗 ---------------- */
  function openModal() {
    var r = state.res;
    if (!r) { toast('先点「开始重命名」', 'warn'); return; }
    var body = $('#renModalBody');
    var rows = r.fails;
    body.innerHTML = rows.length
      ? '<div class="idle-twrap"><table class="att-table"><thead><tr>' +
        '<th>#</th><th>原文件名</th><th>模块</th><th>失败原因</th><th>改正建议</th>' +
        '</tr></thead><tbody>' +
        rows.map(function (x, i) {
          return '<tr><td>' + (i + 1) + '</td><td>' + esc(x.file) + '</td><td>' + esc(x.module || '') + '</td>' +
            '<td style="color:var(--red)">' + esc(x.reason) + '</td><td>' + esc(x.advice) + '</td></tr>';
        }).join('') + '</tbody></table></div>'
      : '<div class="att-empty">没有问题，全部重命名成功 🎉</div>';
    $('#renModal').classList.remove('hidden');
  }
  function closeModal() { $('#renModal').classList.add('hidden'); }

  function doExportProblems() {
    var r = state.res;
    if (!r) return;
    var rows = r.fails.map(function (x, i) { return [i + 1, x.file, x.module || '', x.reason, x.advice]; });
    if (!rows.length) { toast('没有问题，无需导出', 'warn'); return; }
    W.build({
      zipImpl: Z,
      sheets: [{
        name: '重命名问题清单',
        header: ['序号', '原文件名', '模块', '失败原因', '改正建议'],
        rows: rows, headerStyle: 7, freeze: { x: 0, y: 1 }, autoFilter: true,
        rowHeight: 24, cols: [8, 34, 16, 44, 44]
      }]
    }).then(function (bytes) {
      downloadBlob(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
        '重命名问题清单-' + stamp() + '.xlsx');
    });
  }

  /* ---------------- 绑定 ---------------- */
  function bind() {
    if (!$('#viewRename')) return;
    state.tplCfg = loadTplCfg();
    renderRoster();
    renderGrid();
    renderChk();
    renderTplBuilder();
    bindAddBtn();
    bindExtInput();
    var run = $('#btnRenRun');
    if (run) run.addEventListener('click', doRun);
    var pb = $('#btnRenProblems');
    if (pb) pb.addEventListener('click', openModal);
    var mc = $('#renModalClose');
    if (mc) mc.addEventListener('click', closeModal);
    var mask = $('#renModal');
    if (mask) mask.addEventListener('click', function (e) { if (e.target === mask) closeModal(); });
    var px = $('#btnRenProbXlsx');
    if (px) px.addEventListener('click', doExportProblems);
  }

  document.addEventListener('DOMContentLoaded', bind);
  window.RenApp = {
    onShow: function () {
      state.tplCfg = loadTplCfg();
      renderRoster();
      renderGrid();
      renderChk();
      renderTplBuilder();
      bindAddBtn();
      bindExtInput();
      renderResults();
    },
    _test: {
      setLimits: function (total, count) { MAX_TOTAL = total; MAX_COUNT = count; },
      defaults: function () { MAX_TOTAL = 300 * 1024 * 1024; MAX_COUNT = 500; }
    }
  };
})();
