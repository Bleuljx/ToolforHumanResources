/* ============================================================
 * core.js — 薪资报送核心逻辑（不依赖 DOM，可在 Node 里测试）
 * 职责：识别表头/工作表、列映射自动匹配、三张表的数据计算
 * ============================================================ */
(function (root, factory) {
  var api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.SalaryCore = api;
})(typeof window !== 'undefined' ? window : globalThis, function () {
  'use strict';

  /* ---------------- 默认配置 ---------------- */
  var DEFAULT_CONFIG = {
    version: 1,
    attendance: {
      city: '蝶城', project: '项目', name: '姓名', code: '员工编码',
      restDays: 'WY当月应休月休天数', shouldAttend: 'WY应出勤天数',
      actualAttend: 'WY实际出勤天数', baseSalary: '基本薪资',
      hourlyRate: 'WJ定额小时加班费', otHours: 'WJ加班时长（小时）',
      otFee: 'WY加班费', remark: 'WY备注',
      reason: '', lateCount: '', missCount: ''
    },
    category: {
      r1: { start: 'WJ租房补贴', end: '工资补发', type: '津补贴' },
      r2: { start: 'WY考核罚款', end: 'WY其他扣款', type: '扣罚类' },
      r3: { start: '五险一金合计扣款（单位）', end: '增值税', type: '成本类' },
      deduct: '五险一金合计扣款（个人）',
      // 备注兜底区间：这段列的单元格没有批注时，用源表备注列的值兜底填充
      fbStart: '其他补款', fbEnd: '其他扣款',
      // 考勤明细区域（1~31 日的纯数字表头列）：自动识别，可手动指定起止表头名
      dayAuto: true, dayStart: '', dayEnd: ''
    },
    importCfg: { task: '', date: '', map: {} },
    // 宽表（非循环项导入(宽表)）：map = 模板列名 → 算薪表列名（@ 开头为考勤字段自动值）
    wideCfg: { date: '', map: {} },
    // 忽略列：这些列即使有金额也不提示「区间外项目」（表头标黄的列会自动忽略，无需配置）
    ignoreCols: ['工时']
  };

  var ATTENDANCE_KEYS = [
    ['city', '蝶城'], ['project', '所在项目'], ['name', '姓名'], ['code', '员工编码'],
    ['restDays', '员工固定月休'], ['shouldAttend', 'WY应出勤天数'], ['actualAttend', 'WY实际出勤天数'],
    ['baseSalary', '基本薪资（算定额值班费用）'], ['hourlyRate', '标准 元/小时'],
    ['otHours', '加班小时数'], ['otFee', 'WY加班费'], ['remark', 'WY备注'],
    ['reason', '加班理由说明'], ['lateCount', '迟到次数'], ['missCount', '漏打卡次数']
  ];

  var CANDIDATES = {
    city: ['蝶城', '所属蝶城', '城市', '蝶城/蝶外'],
    project: ['项目', '所在项目', '项目名称'],
    name: ['姓名', '员工姓名', '人员姓名'],
    code: ['员工编码', '工号', '人员编码', '员工编号'],
    restDays: ['WY当月应休月休天数', '当月应休月休天数', '应休月休天数', '月休天数'],
    shouldAttend: ['WY应出勤天数', '应出勤天数'],
    actualAttend: ['WY实际出勤天数', '实际出勤天数', '实际出勤'],
    baseSalary: ['基本薪资', '基本工资', '月薪', '基础工资'],
    hourlyRate: ['WJ定额小时加班费', '定额小时加班费', '标准 元/小时', '加班费标准'],
    otHours: ['WJ加班时长（小时）', '加班时长（小时）', '加班小时数', '加班时长'],
    otFee: ['WY加班费', '加班费'],
    remark: ['WY备注', '备注'],
    reason: ['加班理由说明', '加班理由'],
    lateCount: ['迟到次数'],
    missCount: ['漏打卡次数']
  };

  var HEADER_KEYWORDS = ['姓名', '员工编码', '工号', '基本薪资', '岗位', '项目', '蝶城',
    '应出勤', '实际出勤', '部门', '身份证', '供应商'];

  /* ---------------- 文本工具 ---------------- */
  function norm(s) {
    if (s == null) return '';
    return String(s)
      .replace(/[\s\u3000]/g, '')
      .replace(/[（）]/g, function (m) { return m === '（' ? '(' : ')'; })
      .replace(/[，]/g, ',')
      .replace(/[：]/g, ':')
      .toLowerCase();
  }
  function stripPrefix(s) {
    return String(s == null ? '' : s).replace(/^[Ww][IJYijy]/, '').replace(/[\s\u3000]/g, '');
  }
  function toStr(v) {
    if (v == null) return '';
    if (typeof v === 'string') return v.trim();
    return String(v).trim();
  }
  function toNum(v) {
    if (typeof v === 'number') return isFinite(v) ? v : NaN;
    if (v == null) return NaN;
    var s = String(v).replace(/[\s\u3000,，￥¥$元]/g, '');
    if (s === '') return NaN;
    if (!/^[-+]?(\d+\.?\d*|\.\d+)$/.test(s)) return NaN;
    return Number(s);
  }
  function round2(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }

  function dice(a, b) {
    a = String(a); b = String(b);
    if (!a || !b) return 0;
    if (a === b) return 1;
    var A = {}, B = {}, i;
    for (i = 0; i < a.length; i++) A[a[i]] = (A[a[i]] || 0) + 1;
    for (i = 0; i < b.length; i++) B[b[i]] = (B[b[i]] || 0) + 1;
    var inter = 0;
    for (var k in A) if (B[k]) inter += Math.min(A[k], B[k]);
    return (2 * inter) / (a.length + b.length);
  }

  /* ---------------- 工作表与表头识别 ---------------- */
  function detectHeaderRow(XLSX, ws, maxScan) {
    if (!ws || !ws['!ref']) return -1;
    var range = XLSX.utils.decode_range(ws['!ref']);
    var limit = Math.min(range.e.r, (maxScan || 12) - 1);
    var best = -1, bestScore = 0;
    for (var r = range.s.r; r <= limit; r++) {
      var matches = 0, nonEmpty = 0;
      for (var c = range.s.c; c <= Math.min(range.e.c, 80); c++) {
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        if (!cell || cell.v == null || cell.v === '') continue;
        nonEmpty++;
        var txt = String(cell.v);
        for (var k = 0; k < HEADER_KEYWORDS.length; k++) {
          if (txt.indexOf(HEADER_KEYWORDS[k]) >= 0) { matches++; break; }
        }
      }
      var score = matches * 10 + Math.min(nonEmpty, 30) * 0.3;
      if (matches >= 2 && score > bestScore) { bestScore = score; best = r; }
    }
    return best;
  }

  function pickSheet(XLSX, wb) {
    var names = wb.SheetNames.slice();
    var i, s;
    for (i = 0; i < names.length; i++) if (names[i].indexOf('新公式') >= 0) return names[i];
    for (i = 0; i < names.length; i++) if (names[i].indexOf('算薪') >= 0) return names[i];
    for (i = 0; i < names.length; i++) {
      s = detectHeaderRow(XLSX, wb.Sheets[names[i]]);
      if (s >= 0) return names[i];
    }
    return names[0];
  }

  function readHeaders(XLSX, ws, headerRow0) {
    var out = [];
    if (!ws['!ref']) return out;
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (var c = range.s.c; c <= range.e.c; c++) {
      var cell = ws[XLSX.utils.encode_cell({ r: headerRow0, c: c })];
      if (!cell || cell.v == null || String(cell.v).trim() === '') continue;
      out.push({ col: c + 1, idx: c, name: String(cell.v).trim() });
    }
    return out;
  }

  function cellOf(XLSX, ws, row0, col0) {
    return ws[XLSX.utils.encode_cell({ r: row0, c: col0 })];
  }

  /** 读取文件：返回结构信息（不含生成） */
  function parseFile(XLSX, buffer, fileName, opts) {
    var wb = XLSX.read(new Uint8Array(buffer), { type: 'array', cellComments: true, cellDates: true, cellStyles: true });
    var sheetName = (opts && opts.sheetName) || pickSheet(XLSX, wb);
    var ws = wb.Sheets[sheetName];
    var hRow = detectHeaderRow(XLSX, ws);
    var headerRow = (opts && opts.headerRow) || (hRow >= 0 ? hRow + 1 : 3);
    var headers = readHeaders(XLSX, ws, headerRow - 1);
    var dataStart = (opts && opts.dataStartRow) || guessDataStart(XLSX, ws, headerRow, headers);
    return {
      fileName: fileName,
      wb: wb,
      sheetNames: wb.SheetNames.slice(),
      sheetName: sheetName,
      headerRow: headerRow,
      dataStartRow: dataStart,
      headers: headers,
      yellowCols: detectYellowCols(XLSX, ws, headerRow - 1),
      attendArea: detectAttendArea(headers),
      city: (opts && opts.city) || detectCity(fileName)
    };
  }

  /* ---------------- 表头标黄列识别 ---------------- */
  // 黄色系判定：红绿分量都高、蓝分量明显偏低（淡黄/金黄/橙色都算），绿色、蓝色、灰白不算
  function yellowish(rgbHex, tint) {
    var h = String(rgbHex == null ? '' : rgbHex).replace('#', '');
    if (h.length > 6) h = h.slice(-6);
    if (!/^[0-9A-Fa-f]{6}$/.test(h)) return false;
    var r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16);
    var t = Number(tint) || 0;
    if (t > 0) { r = r * (1 - t) + 255 * t; g = g * (1 - t) + 255 * t; b = b * (1 - t) + 255 * t; }
    else if (t < 0) { r = r * (1 + t); g = g * (1 + t); b = b * (1 + t); }
    r = Math.round(r); g = Math.round(g); b = Math.round(b);
    return r > 140 && g > 140 && ((r + g) / 2 - b) > 30 && (r - g) > -25;
  }
  function headerFgColor(cell) {
    var s = cell && cell.s;
    if (!s) return null;
    // SheetJS 读出的样式可能是扁平结构（s.patternType/s.fgColor）或嵌套结构（s.patternFill.fgColor）
    var pt = s.patternType, fg = s.fgColor;
    if (s.patternFill) {
      pt = pt || s.patternFill.patternType;
      fg = fg || s.patternFill.fgColor;
    }
    if (s.fill && typeof s.fill === 'object') {
      var inner = s.fill.patternFill || s.fill;
      pt = pt || inner.patternType;
      fg = fg || inner.fgColor;
    }
    if (pt && pt !== 'solid') return null;
    return fg || null;
  }
  function detectYellowCols(XLSX, ws, headerRow0) {
    if (!ws || !ws['!ref']) return [];
    var range = XLSX.utils.decode_range(ws['!ref']);
    var out = [];
    for (var c = range.s.c; c <= range.e.c; c++) {
      var cell = ws[XLSX.utils.encode_cell({ r: headerRow0, c: c })];
      if (!cell || cell.v == null || String(cell.v).trim() === '') continue;
      var fg = headerFgColor(cell);
      if (!fg) continue;
      if (yellowish(fg.rgb, fg.tint)) out.push({ col: c + 1, name: String(cell.v).trim() });
    }
    return out;
  }

  function guessDataStart(XLSX, ws, headerRow, headers) {
    var nameCol = findCol(headers, ['姓名', '员工姓名']);
    if (!nameCol) return headerRow + 2;
    function isGuideRow(r) {
      // 「填写要求」等填写说明行：姓名列常填「张三」、编码列填「取自SHR」等提示文案
      var limit = Math.min(headers.length, 15);
      for (var c = 1; c <= limit; c++) {
        var cell = cellOf(XLSX, ws, r - 1, c - 1);
        if (!cell || cell.v == null) continue;
        if (/填写要求|取自SHR|取自shr|调度中心全称|和战图一致/.test(String(cell.v))) return true;
      }
      return false;
    }
    for (var r = headerRow + 1; r <= headerRow + 4; r++) {
      if (isGuideRow(r)) continue;
      var cell = cellOf(XLSX, ws, r - 1, nameCol - 1);
      if (cell && cell.v != null && String(cell.v).trim() !== '') return r;
    }
    return headerRow + 2;
  }

  function detectCity(fileName) {
    var base = String(fileName || '').replace(/\.(xlsx|xlsm|xls)$/i, '');
    var m = base.match(/^(.+?(?:蝶城|蝶外))/);
    if (m) return m[1];
    return base.slice(0, 20);
  }

  /* ---------------- 列定位 ---------------- */
  function findCol(headers, candidates) {
    var i, j;
    for (i = 0; i < candidates.length; i++) {
      for (j = 0; j < headers.length; j++) if (headers[j].name === candidates[i]) return headers[j].col;
    }
    for (i = 0; i < candidates.length; i++) {
      for (j = 0; j < headers.length; j++) if (norm(headers[j].name) === norm(candidates[i])) return headers[j].col;
    }
    for (i = 0; i < candidates.length; i++) {
      var nc = norm(candidates[i]);
      for (j = 0; j < headers.length; j++) {
        var nh = norm(headers[j].name);
        if (nh.indexOf(nc) >= 0 || nc.indexOf(nh) >= 0) return headers[j].col;
      }
    }
    return 0;
  }

  function colIndexByName(headers, name) {
    if (!name) return 0;
    if (/^\d+$/.test(name)) return parseInt(name, 10);
    if (/^[A-Za-z]+$/.test(name)) {
      var n = 0;
      for (var i = 0; i < name.length; i++) n = n * 26 + (name.toUpperCase().charCodeAt(i) - 64);
      return n;
    }
    for (var j = 0; j < headers.length; j++) if (headers[j].name === name) return headers[j].col;
    for (var k = 0; k < headers.length; k++) if (norm(headers[k].name) === norm(name)) return headers[k].col;
    return 0;
  }

  /* ---------------- 批注 ---------------- */
  function stripAuthor(text, author) {
    var t = String(text == null ? '' : text).replace(/\r/g, '');
    if (author) {
      var an = String(author);
      if (t.indexOf(an + ':') === 0) t = t.slice(an.length + 1);
      else if (t.indexOf(an + '：') === 0) t = t.slice(an.length + 1);
      else if (t.indexOf(an + ': ') === 0) t = t.slice(an.length + 2);
      else if (t.indexOf(an + '： ') === 0) t = t.slice(an.length + 2);
    }
    // 去掉「某人: 内容」前缀（作者名可能很长、也可能含数字/工号）；
    // 仅排除「纯数字」和「时间 12:30」这类误伤情况
    var m = t.match(/^([^\n:：]{1,30})[:：]\s*/);
    if (m && !/^\d+$/.test(m[1]) && !/^\d{1,2}:\d{2}/.test(t)) t = t.slice(m[0].length);
    return t.replace(/[\n\r]+/g, ' ').trim();
  }

  function readComment(XLSX, ws, row0, col0) {
    var cell = cellOf(XLSX, ws, row0, col0);
    if (!cell || !cell.c || !cell.c.length) return '';
    // 一格可能有多段批注（多人追加/富文本被拆段），全部合并，避免只取到作者行丢内容
    var parts = [];
    cell.c.forEach(function (c) {
      var t = c.t || c.h || (c.r ? String(c.r).replace(/<[^>]+>/g, '') : '');
      t = String(t == null ? '' : t).replace(/\r/g, '').trim();
      if (t) parts.push(t);
    });
    if (!parts.length) return '';
    var t0 = parts[0];
    var s = stripAuthor(parts.join('\n'), cell.c[0].a);
    s = s.replace(/[\n\r]+/g, ' ').trim();
    // 作者前缀剥空时用原文兜底，避免整条批注丢失
    return s || t0.replace(/[\n\r]+/g, ' ').trim();
  }

  /* ---------------- 自动匹配 662 项标准薪酬项目 ---------------- */
  function matchOption(headerName, options) {
    var h = toStr(headerName);
    if (!h) return { value: '', conf: 'none' };
    var i;
    for (i = 0; i < options.length; i++) if (options[i] === h) return { value: options[i], conf: 'exact' };
    var nh = norm(h);
    for (i = 0; i < options.length; i++) if (norm(options[i]) === nh) return { value: options[i], conf: 'norm' };
    var sh = stripPrefix(h), nsh = norm(sh);
    if (sh !== h) {
      for (i = 0; i < options.length; i++) if (options[i] === sh || norm(options[i]) === nsh) return { value: options[i], conf: 'prefix' };
    }
    var best = '', bestScore = 0;
    for (i = 0; i < options.length; i++) {
      var no = norm(options[i]);
      var sc = 0;
      if (no === nsh || no === nh) sc = 1;
      else if (no && (no.indexOf(nsh) >= 0 || nsh.indexOf(no) >= 0)) sc = 0.75 + 0.2 * (Math.min(no.length, nsh.length) / Math.max(no.length, nsh.length));
      else sc = dice(nsh, no);
      if (sc > bestScore) { bestScore = sc; best = options[i]; }
    }
    if (best && bestScore >= 0.62) return { value: best, conf: 'guess' };
    return { value: '', conf: 'none' };
  }

  // 汇总列（应发合计/实发合计等）：只是核对用的合计数，不属于薪酬项目，不进导入模板
  function isSummaryCol(name) {
    var n = norm(name);
    return /(应发|实发|实付|应付)/.test(n) && /(合计|总计|小计)/.test(n);
  }

  /* ---------------- 导入模板的 7 个考勤字段 ----------------
   * 除三大区间项目外，这些考勤字段也要梳理进非循环项导入模板（值非 0 才导）。
   * key 用 @ 前缀避免与真实表头重名；from 用于对 662 项做自动匹配的参考名；
   * srcKey 指向考勤配置里的源列（dutyFee 是自动计算值，无源列）。
   */
  var PSEUDO_FIELDS = [
    { key: '@baseSalary',   label: '基本薪资',              from: '基本薪资',         srcKey: 'baseSalary' },
    { key: '@restDays',     label: '当月应休天数',          from: '当月应休月休天数', srcKey: 'restDays' },
    { key: '@shouldAttend', label: '应出勤天数',            from: '应出勤天数',       srcKey: 'shouldAttend' },
    { key: '@actualAttend', label: '实际出勤天数',          from: '实际出勤天数',     srcKey: 'actualAttend' },
    { key: '@dutyFee',      label: '定额值班费（自动计算）', from: '定额值班费',       srcKey: null },
    { key: '@hourlyRate',   label: '定额小时加班费',        from: '定额小时加班费',   srcKey: 'hourlyRate' },
    { key: '@otHours',      label: '加班时长（小时）',      from: '加班时长（小时）', srcKey: 'otHours' }
  ];

  /* ---------------- 考勤明细区域识别（1~31 日） ----------------
   * 表头为纯数字 1~31 的连续列视为考勤明细区域：不进导入模板、不提示区间外。
   * 取最长的一段连续数字列。
   */
  function detectAttendArea(headers) {
    var best = null, run = null;
    headers.forEach(function (h) {
      var t = String(h.name == null ? '' : h.name).trim();
      if (/^\d{1,2}$/.test(t) && +t >= 1 && +t <= 31) {
        if (run && h.col === run.endCol + 1) { run.endCol = h.col; run.names.push(t); }
        else run = { startCol: h.col, endCol: h.col, names: [t] };
        if (!best || run.names.length > best.names.length) best = run;
      } else run = null;
    });
    return best;
  }

  function autoMapImport(headers, spanCols, options, existing) {
    var map = {};
    spanCols.forEach(function (h) {
      var old = existing && existing[h.name];
      var oldVal = old ? old.value : '';
      // 已有人工/有效映射（非空、非未匹配态）则保留
      var keepOld = !!old && oldVal !== '' && (oldVal === '__skip__' || old.conf !== 'none');
      if (keepOld) { map[h.name] = old; return; }
      if (isSummaryCol(h.name)) { map[h.name] = { value: '__skip__', conf: 'skip' }; return; }
      var m = matchOption(h.name, options);
      map[h.name] = { value: m.value, conf: m.conf };
    });
    return map;
  }

  /* ---------------- 宽表（非循环项导入(宽表)）----------------
   * 一人一行 32 列；28 个薪资项目列按映射取数；备注列汇总全行批注。
   * 7 个考勤字段默认走 PSEUDO 自动值（定额值班费为计算值），其余列自动按
   * 归一化/去前缀匹配算薪表列，映射可在页面手动改。
   */
  var WIDE_ITEMS = [
    '基本工资', '应休月休天数', '应出勤天数', '实际出勤天数', '定额值班费', '定额小时加班费', '加班时长（小时）',
    '租房补贴', '高温补贴', '交通住宿补助', '专项补贴', '垃圾点位专项补贴', '法定假补贴', '其他补款',
    '推荐奖励', '考核激励', '其他奖金', '全勤奖励', '工资补发',
    '考核罚款', '考勤扣款', '上月扣回', '其他扣款', '五险一金合计扣款（个人）',
    '五险一金合计扣款（单位）', '商业险', '服务费', '增值税'
  ];
  var WIDE_HEADER = ['员工编码', '员工姓名', '业务日期(yyyy-MM-dd)'].concat(WIDE_ITEMS, ['备注']);
  // 数量列（不按金额格式化，原值写入）
  var WIDE_QTY = { '应休月休天数': 1, '应出勤天数': 1, '实际出勤天数': 1, '加班时长（小时）': 1 };
  // 7 个考勤字段的默认映射（@key 见 PSEUDO_FIELDS；定额值班费为自动计算值）
  var WIDE_PSEUDO = {
    '基本工资': '@baseSalary', '应休月休天数': '@restDays', '应出勤天数': '@shouldAttend',
    '实际出勤天数': '@actualAttend', '定额值班费': '@dutyFee',
    '定额小时加班费': '@hourlyRate', '加班时长（小时）': '@otHours'
  };

  // 模板列名 → 算薪表表头自动匹配：同名 → 归一化同名 → 去前缀同名 → 包含
  function matchWideHeader(tplName, headers) {
    var nt = norm(tplName), st = norm(stripPrefix(tplName));
    var i, h;
    for (i = 0; i < headers.length; i++) if (headers[i].name === tplName) return headers[i];
    for (i = 0; i < headers.length; i++) if (norm(headers[i].name) === nt) return headers[i];
    for (i = 0; i < headers.length; i++) if (norm(stripPrefix(headers[i].name)) === st) return headers[i];
    for (i = 0; i < headers.length; i++) {
      h = norm(stripPrefix(headers[i].name));
      if (h && st && (h.indexOf(st) >= 0 || st.indexOf(h) >= 0)) return headers[i];
    }
    return null;
  }

  // 金额写进备注的格式：原值去尾零（200 → "200"，200.50 → "200.5"，-100 → "-100"）
  function fmtAmt(n) {
    var r = Math.round(n * 100) / 100;
    return String(r);
  }

  /* ---------------- 主处理：生成三张表数据 ---------------- */
  function processFile(XLSX, file, cfg, options) {
    var ws = file.wb.Sheets[file.sheetName];
    var headers = file.headers;
    var problems = [];
    var hRow0 = file.headerRow - 1;
    var startRow = file.dataStartRow;
    var range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
    var lastRow = range.e.r + 1;

    function addProblem(row, name, type, detail, level) {
      problems.push({
        file: file.fileName, sheet: file.sheetName, row: row || '',
        name: name || '', type: type, detail: detail || '', level: level || 'warn'
      });
    }

    // 关键列定位
    var am = cfg.attendance || {};
    var col = {};
    var missing = [];
    ATTENDANCE_KEYS.forEach(function (pair) {
      var key = pair[0];
      var want = am[key];
      var idx = colIndexByName(headers, want);
      if (!idx && CANDIDATES[key]) idx = findCol(headers, CANDIDATES[key]);
      col[key] = idx;
      if (!idx && key !== 'reason' && key !== 'lateCount' && key !== 'missCount') missing.push(pair[1]);
    });
    if (!col.name) {
      addProblem('', '', '表头识别失败', '未找到“姓名”列，无法处理该文件（请检查表头行号）', 'error');
      return { file: file, attendance: null, category: null, imp: null, problems: problems, stats: null };
    }
    if (missing.length) addProblem('', '', '列缺失', '未匹配到列：' + missing.join('、'), 'warn');

    // 三大区间列定位
    var cc = cfg.category || {};
    var r1s = colIndexByName(headers, cc.r1 && cc.r1.start), r1e = colIndexByName(headers, cc.r1 && cc.r1.end);
    var r2s = colIndexByName(headers, cc.r2 && cc.r2.start), r2e = colIndexByName(headers, cc.r2 && cc.r2.end);
    var r3s = colIndexByName(headers, cc.r3 && cc.r3.start), r3e = colIndexByName(headers, cc.r3 && cc.r3.end);
    var dcol = colIndexByName(headers, cc.deduct);
    if (!r1s || !r1e || !r3s || !r3e) {
      addProblem('', '', '区间配置无效', '三大类别的起始/终止项目未能在表头中找到，请到【工资类目一键表】配置里重新选择', 'error');
    }

    // 备注兜底区间（没有批注时用源表备注兜底的列范围）
    var fbS = colIndexByName(headers, cc.fbStart) || (cc.fbStart ? findCol(headers, [cc.fbStart]) : 0);
    var fbE = colIndexByName(headers, cc.fbEnd) || (cc.fbEnd ? findCol(headers, [cc.fbEnd]) : 0);
    if (fbS && fbE && fbE < fbS) { var ft = fbS; fbS = fbE; fbE = ft; }
    function fbCovers(colNo) { return !!(fbS && fbE && colNo >= fbS && colNo <= fbE); }
    var spanStart = r1s || 1, spanEnd = r3e || 1;
    if (spanEnd < spanStart) { var t = spanStart; spanStart = spanEnd; spanEnd = t; }

    // 忽略列（共用一份名单）：一键表、导入模板都不导出，也不触发区间外提示
    var ignoreSet = {};
    (cfg.ignoreCols || []).forEach(function (n) { if (n) ignoreSet[norm(n)] = 1; });

    // 考勤明细区域（1~31 日）：自动识别，可手动指定起止表头名
    function headerColByName(headers, name) {
      var n = norm(name);
      for (var i = 0; i < headers.length; i++) if (norm(headers[i].name) === n) return headers[i].col;
      return 0;
    }
    var daySet = {};
    if (cc.dayAuto === false && cc.dayStart && cc.dayEnd) {
      var dS = headerColByName(headers, cc.dayStart), dE = headerColByName(headers, cc.dayEnd);
      if (dS && dE) {
        if (dE < dS) { var dt = dS; dS = dE; dE = dt; }
        for (var dc = dS; dc <= dE; dc++) daySet[dc] = 1;
      }
    } else if (file.attendArea) {
      for (var dc2 = file.attendArea.startCol; dc2 <= file.attendArea.endCol; dc2++) daySet[dc2] = 1;
    }

    // 报送列清单（区间内 + 个人扣款列）
    var spanCols = [];
    headers.forEach(function (h) {
      if (ignoreSet[norm(h.name)]) return;                 // 忽略列：两表都不导出
      if (h.col >= spanStart && h.col <= spanEnd) spanCols.push(h);
      else if (dcol && h.col === dcol && !(h.col >= spanStart && h.col <= spanEnd)) spanCols.push(h);
    });
    spanCols.sort(function (a, b) { return a.col - b.col; });

    // 每列是否有数据（用于映射校验）
    var hasData = {};
    spanCols.forEach(function (h) { hasData[h.name] = false; });

    var attHeader = ['序号', '蝶城', '所在项目', '姓名', '员工编码',
      '员工固定月休\n（填写月休额度）',
      'WY应出勤天数\n（月自然天数-WY当月应休月休天数）',
      'WY实际出勤天数\n（系统实际打卡上班天数）',
      '月休加班天数\n（值班天数）',
      'WY定额值班费', '标准 元/小时', '加班小时数', 'WY加班费',
      '加班理由说明', '迟到次数', '漏打卡次数', 'WY备注'];
    var catHeader = ['序号', '蝶城', '所在项目', '姓名', '员工编码', '岗位', '类型', '报送项目', '金额', 'WY备注', '填写说明'];
    var impHeader = (options && options.importHeaders) || ['员工编码', '员工姓名', '绑定核算任务', '业务日期(yyyy-MM-dd)', '薪酬项目', '项目值', '备注说明'];

    var attRows = [], catRows = [], impRows = [], wideRows = [];
    var serial = 0, totalAmount = 0, skipped = 0;
    var importMap = (cfg.importCfg && cfg.importCfg.map) || {};
    var taskVal = (cfg.importCfg && cfg.importCfg.task) || '';
    var dateVal = (cfg.importCfg && cfg.importCfg.date) || '';
    var pseudoMissing = {};   // 有数据但未映射的考勤字段（整个文件只提示一次）

    // 宽表列解析（模板列 → 算薪表列 / 考勤字段自动值）
    var wideMap = (cfg.wideCfg && cfg.wideCfg.map) || {};
    var wideDate = (cfg.wideCfg && cfg.wideCfg.date) || '';
    var wideResolvers = WIDE_ITEMS.map(function (tpl) {
      var mv = wideMap[tpl];
      if (mv === undefined || mv === null || mv === '') {
        if (WIDE_PSEUDO[tpl]) return { pseudo: WIDE_PSEUDO[tpl] };
        var mh = matchWideHeader(tpl, headers);
        return mh ? { col: mh.col } : null;
      }
      if (mv.charAt(0) === '@') return { pseudo: mv };
      var idx = colIndexByName(headers, mv) || findCol(headers, [mv]) || 0;
      return idx ? { col: idx } : null;
    });

    for (var r = startRow; r <= lastRow; r++) {
      var r0 = r - 1;
      var nameCell = col.name ? cellOf(XLSX, ws, r0, col.name - 1) : null;
      var nameVal = nameCell ? toStr(nameCell.v) : '';
      if (nameVal === '') {
        // 该行是否有其它内容（有内容才提示，避免空行刷屏）
        var anyVal = false;
        for (var q = 0; q < spanCols.length; q++) {
          var cv = cellOf(XLSX, ws, r0, spanCols[q].col - 1);
          if (cv && cv.v != null && String(cv.v).trim() !== '') { anyVal = true; break; }
        }
        if (anyVal) { skipped++; addProblem(r, '', '姓名为空', '该行有金额但姓名为空，已跳过', 'warn'); }
        continue;
      }
      serial++;

      var restDays = col.restDays ? toNum(valAt(XLSX, ws, r0, col.restDays)) : NaN;
      var should = col.shouldAttend ? toNum(valAt(XLSX, ws, r0, col.shouldAttend)) : NaN;
      var actual = col.actualAttend ? toNum(valAt(XLSX, ws, r0, col.actualAttend)) : NaN;
      var basePay = col.baseSalary ? toNum(valAt(XLSX, ws, r0, col.baseSalary)) : NaN;
      var hourly = col.hourlyRate ? toNum(valAt(XLSX, ws, r0, col.hourlyRate)) : NaN;
      var otHours = col.otHours ? toNum(valAt(XLSX, ws, r0, col.otHours)) : NaN;
      var otFee = col.otFee ? toNum(valAt(XLSX, ws, r0, col.otFee)) : NaN;

      if (!isFinite(should)) should = 0;
      if (!isFinite(actual)) actual = 0;
      if (!isFinite(basePay)) basePay = 0;
      if (should === 0 && basePay > 0) addProblem(r, nameVal, '应出勤为0', '应出勤天数为0，定额值班费按0处理', 'warn');

      var otDays = actual - should;
      if (otDays < 0) otDays = 0;
      var dutyFee = should > 0 ? Math.round(basePay / should) : 0;

      // 7 个考勤字段取值（与考勤报送同源：定额值班费为自动计算值）
      var pseudoVals = {
        '@baseSalary': basePay, '@restDays': restDays, '@shouldAttend': should,
        '@actualAttend': actual, '@dutyFee': dutyFee, '@hourlyRate': hourly, '@otHours': otHours
      };

      var srcRemark = col.remark ? toStr(valAt(XLSX, ws, r0, col.remark)) : '';
      // 考勤报送「WY备注」：直接使用源表备注列的值，不拼接批注

      attRows.push([
        serial,
        col.city ? toStr(valAt(XLSX, ws, r0, col.city)) : '',
        col.project ? toStr(valAt(XLSX, ws, r0, col.project)) : '',
        nameVal,
        col.code ? toStr(valAt(XLSX, ws, r0, col.code)) : '',
        isFinite(restDays) ? restDays : '',
        should,
        actual,
        otDays,
        dutyFee,
        isFinite(hourly) ? hourly : '',
        isFinite(otHours) ? otHours : '',
        isFinite(otFee) ? otFee : '',
        col.reason ? toStr(valAt(XLSX, ws, r0, col.reason)) : '',
        col.lateCount ? toStr(valAt(XLSX, ws, r0, col.lateCount)) : '',
        col.missCount ? toStr(valAt(XLSX, ws, r0, col.missCount)) : '',
        srcRemark
      ]);

      // 一键表 / 导入模板
      for (var j = 0; j < spanCols.length; j++) {
        var h = spanCols[j];
        if (isSummaryCol(h.name)) continue;   // 应发合计/实发合计等汇总列：核对用，不报送
        var raw = cellOf(XLSX, ws, r0, h.col - 1);
        if (!raw || raw.v == null || String(raw.v).trim() === '') continue;
        var num = toNum(raw.v);
        if (!isFinite(num)) {
          addProblem(r, nameVal, '金额非数字', '项目「' + h.name + '」的值“' + String(raw.v) + '”不是数字，已跳过', 'warn');
          continue;
        }
        if (num === 0) continue;
        hasData[h.name] = true;

        var type;
        if (dcol && h.col === dcol) { type = '扣罚类'; }
        else if (h.col >= (r1s || 0) && h.col <= (r1e || 0)) { type = '津补贴'; }
        else if (h.col >= (r2s || 0) && h.col <= (r2e || 0)) { type = '扣罚类'; }
        else if (h.col >= (r3s || 0) && h.col <= (r3e || 0)) { type = '成本类'; }
        else continue;

        // 扣款一律用源表原值（正负号不改）：导入系统有公式自动计算，改动正负会导致计算错误
        var amount = round2(num);
        var cellCmt = readComment(XLSX, ws, r0, h.col - 1);
        // 一键表「WY备注」与导入模板「备注说明」同规则：
        // 优先取该行对应源表单元格的批注；没有批注且列在「备注兜底区间」内 → 用源表备注列的值兜底；否则留空
        var remarkVal = cellCmt || '';
        if (!remarkVal && srcRemark && fbCovers(h.col)) remarkVal = srcRemark;

        catRows.push([
          catRows.length + 1,
          col.city ? toStr(valAt(XLSX, ws, r0, col.city)) : '',
          col.project ? toStr(valAt(XLSX, ws, r0, col.project)) : '',
          nameVal,
          col.code ? toStr(valAt(XLSX, ws, r0, col.code)) : '',
          headersMapPost(headers, ws, XLSX, r0, file) ,
          type,
          h.name,
          amount,
          remarkVal,
          ''
        ]);
        totalAmount += amount;

        var mapped = importMap[h.name];
        var target = mapped && typeof mapped === 'object' ? mapped.value : mapped;
        if (!target) {
          addProblem(r, nameVal, '薪酬项目未映射', '项目「' + h.name + '」未匹配到标准薪酬项目，导入模板将缺少这条数据', 'block-import');
        } else if (target !== '__skip__') {
          impRows.push([
            col.code ? toStr(valAt(XLSX, ws, r0, col.code)) : '',
            nameVal,
            taskVal,
            dateVal,
            target,
            amount,
            remarkVal
          ]);
        }
      }

      // 7 个考勤字段 → 导入模板（值非 0 才导；备注只取对应源单元格的批注）
      var codeVal = col.code ? toStr(valAt(XLSX, ws, r0, col.code)) : '';
      PSEUDO_FIELDS.forEach(function (pf) {
        var pv = pseudoVals[pf.key];
        if (!isFinite(pv) || pv === 0) return;             // 非 0 才导
        var srcNo = pf.srcKey ? col[pf.srcKey] : 0;
        var cmt = srcNo ? readComment(XLSX, ws, r0, srcNo - 1) : '';
        var pm = importMap[pf.key];
        var pTarget = pm && typeof pm === 'object' ? pm.value : pm;
        if (!pTarget) { pseudoMissing[pf.key] = true; return; }
        if (pTarget === '__skip__') return;
        impRows.push([codeVal, nameVal, taskVal, dateVal, pTarget, round2(pv), cmt]);
      });

      // ---- 宽表（一人一行）：28 项目列 + 备注汇总全行批注 ----
      // 批注段 = 批注内容 + 金额（扣罚类区间/个人扣款列写成负数）；多条分号隔开；
      // 批注之间按原文归一化去重；批注内容本身已含数字（如「迟到扣10元」）时不再拼单元格金额；
      // 与源备注列去重按「剥离金额后的内容」比较（星星徽章50 = 星星徽章50元），保留更长的一条（保住金额信息）
      var wideSegs = [], wideKeys = [];
      function remarkKey(s) {
        return norm(s).replace(/[0-9.+\-,:;；、()（）%元整多只]/g, '');
      }
      headers.forEach(function (h) {
        if (h.col === col.name || h.col === col.code || h.col === col.city ||
            h.col === col.project || h.col === col.remark ||
            h.col === col.reason || h.col === col.lateCount || h.col === col.missCount) return;
        if (daySet[h.col]) return;
        if (isSummaryCol(h.name)) return;
        var cmt = readComment(XLSX, ws, r0, h.col - 1).replace(/[；;]+\s*$/, '').trim();
        if (!cmt) return;
        var ckey = norm(cmt);
        if (ckey && wideKeys.indexOf(ckey) >= 0) return;
        var seg = cmt;
        // 批注内容已含数字（金额往往已写在批注里，如「扣10元」「+300」）→ 不再拼单元格金额
        if (!/\d/.test(cmt)) {
          var raw = cellOf(XLSX, ws, r0, h.col - 1);
          // 公式单元格无缓存值时退回 w（格式化文本），避免批注在、金额丢
          var rawv = raw ? (raw.v != null ? raw.v : raw.w) : null;
          if (rawv != null && String(rawv).trim() !== '') {
            var cn = toNum(rawv);
            if (isFinite(cn) && cn !== 0) {   // 金额为 0 时不拼数字，避免「…每次100」+「0」变成「…每次1000」
              var isPen = (dcol && h.col === dcol) || (h.col >= (r2s || 0) && h.col <= (r2e || 0));
              seg = cmt + fmtAmt(isPen ? -Math.abs(cn) : cn);
            }
          }
        }
        wideSegs.push(seg);
        if (ckey) wideKeys.push(ckey);
      });
      var wideRemark;
      if (srcRemark) {
        var srcClean = String(srcRemark).trim().replace(/[；;]+$/, '');
        var sk = remarkKey(srcClean);
        var dupIdx = sk ? -1 : -1;
        if (sk) {
          for (var wi2 = 0; wi2 < wideSegs.length; wi2++) {
            if (remarkKey(wideSegs[wi2]) === sk) { dupIdx = wi2; break; }
          }
        }
        if (dupIdx >= 0) {
          // 内容相同：保留更长的一条（源备注通常带金额更完整）
          if (norm(srcClean).length > norm(wideSegs[dupIdx]).length) wideSegs[dupIdx] = srcClean;
        } else {
          wideSegs.push(srcClean);   // 无文字内容的源备注（如纯金额）也原样保留
        }
      }
      wideRemark = wideSegs.join('；');

      var wideVals = WIDE_ITEMS.map(function (tpl, wi) {
        var rs = wideResolvers[wi];
        if (!rs) return '';
        if (rs.pseudo) {
          var pv2 = pseudoVals[rs.pseudo];
          if (!isFinite(pv2)) return '';
          return WIDE_QTY[tpl] ? pv2 : round2(pv2);
        }
        var raw2 = cellOf(XLSX, ws, r0, rs.col - 1);
        if (!raw2 || raw2.v == null || String(raw2.v).trim() === '') return '';
        var n2 = toNum(raw2.v);
        if (!isFinite(n2)) return toStr(raw2.v);
        return WIDE_QTY[tpl] ? n2 : round2(n2);
      });
      wideRows.push([codeVal, nameVal, wideDate].concat(wideVals, [wideRemark]));
    }

    // 考勤字段有数据但未映射到标准项目：整个文件只提示一次，阻断导入
    PSEUDO_FIELDS.forEach(function (pf) {
      if (pseudoMissing[pf.key]) {
        addProblem('', '', '薪酬项目未映射', '考勤字段「' + pf.label + '」未匹配到标准薪酬项目，导入模板将缺少这些数据（请到【非循环项导入模板】配置里选择）', 'block-import');
      }
    });

    // 区间末列之后、且不属于考勤报送取数列的“疑似新增工资项目”：提示，避免静默丢数
    // 忽略：表头标黄的列（自动识别）+ 手动忽略列 + 考勤明细区域（1~31 日）
    var yellowSet = {};
    (file.yellowCols || []).forEach(function (y) { yellowSet[y.col] = 1; });
    var usedCols = {};
    Object.keys(col).forEach(function (k) { if (col[k]) usedCols[col[k]] = 1; });
    var META = ['序号', '供应商', '备注栏', '入职日期', '离职日期', '身份证', '部门', '备注',
      '定额值班费', '应休月休', '应出勤', '实际出勤', '定额小时加班费', '加班时长', '加班小时',
      '基本薪资', '基本工资', '月薪', '岗位'];
    headers.forEach(function (h) {
      if (h.col <= spanEnd) return;                       // 区间内已处理
      if (dcol && h.col === dcol) return;                  // 个人扣款列
      if (usedCols[h.col]) return;                         // 考勤报送已取数
      if (yellowSet[h.col]) return;                        // 表头标黄的列：约定不报送，不提示
      if (ignoreSet[norm(h.name)]) return;                 // 手动忽略列
      if (daySet[h.col]) return;                           // 考勤明细区域（1~31 日）
      if (isSummaryCol(h.name)) return;                    // 应发/实发合计等汇总列：核对用，不提示
      var nh = norm(h.name);
      for (var mi = 0; mi < META.length; mi++) {
        if (nh.indexOf(norm(META[mi])) >= 0) return;
      }
      for (var rr = startRow; rr <= lastRow; rr++) {
        var v = toNum(valAt(XLSX, ws, rr - 1, h.col));
        if (isFinite(v) && v !== 0) {
          addProblem('', '', '区间外项目',
            '列「' + h.name + '」有金额，但没有被划进津补贴/扣罚类/成本类任何一个区间，所以没有导出到一键表。' +
            '如果该列本来就不需要报送，可忽略此提示；如需报送请调整区间，或把它加入忽略列', 'warn');
          break;
        }
      }
    });

    if (!dateVal) addProblem('', '', '业务日期为空', '非循环项导入模板需要“业务日期”，请在上方填写后再导出', 'block-import');
    if (!taskVal) addProblem('', '', '核算任务为空', '“绑定核算任务”为空（系统允许为空，如需绑定请填写）', 'info');

    return {
      file: file,
      attendance: { header: attHeader, rows: attRows },
      category: { header: catHeader, rows: catRows },
      imp: { header: impHeader, rows: impRows },
      wide: { header: WIDE_HEADER, rows: wideRows },
      problems: problems,
      stats: { employees: serial, attRows: attRows.length, catRows: catRows.length, impRows: impRows.length, wideRows: wideRows.length, amount: round2(totalAmount), skipped: skipped }
    };
  }

  function valAt(XLSX, ws, r0, colNum) {
    var cell = cellOf(XLSX, ws, r0, colNum - 1);
    if (!cell) return '';
    return cell.v == null ? '' : cell.v;
  }

  function headersMapPost(headers, ws, XLSX, r0, file) {
    // 岗位列
    var idx = findCol(headers, ['岗位', '岗位名称', '职位']);
    if (!idx) return '';
    return toStr(valAt(XLSX, ws, r0, idx));
  }

  return {
    DEFAULT_CONFIG: DEFAULT_CONFIG,
    ATTENDANCE_KEYS: ATTENDANCE_KEYS,
    CANDIDATES: CANDIDATES,
    parseFile: parseFile,
    processFile: processFile,
    detectCity: detectCity,
    detectHeaderRow: detectHeaderRow,
    pickSheet: pickSheet,
    readHeaders: readHeaders,
    findCol: findCol,
    colIndexByName: colIndexByName,
    matchOption: matchOption,
    matchWideHeader: matchWideHeader,
    WIDE_ITEMS: WIDE_ITEMS,
    WIDE_HEADER: WIDE_HEADER,
    WIDE_QTY: WIDE_QTY,
    WIDE_PSEUDO: WIDE_PSEUDO,
    fmtAmt: fmtAmt,
    autoMapImport: autoMapImport,
    isSummaryCol: isSummaryCol,
    PSEUDO_FIELDS: PSEUDO_FIELDS,
    detectAttendArea: detectAttendArea,
    norm: norm,
    stripPrefix: stripPrefix,
    toNum: toNum,
    round2: round2
  };
});
