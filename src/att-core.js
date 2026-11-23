/* ============================================================
 * att-core.js — 考勤处理模块核心逻辑
 * 输入：外包(SHR)考勤原始导出 / 机师岗RM考勤原始导出 / 人员花名册
 * 输出：每蝶城一个结果文件（主表+差异明细+迟到早退明细+统计+4张工时矩阵）
 *       + 跨文件汇总总表（迟到早退统计汇总/明细汇总/在职无出勤）
 * 规则来源：SHR人脸及RM考勤批量处理工具.xlsm 两个宏 + 2026-11-17 需求确认
 * ============================================================ */
(function (root, factory) {
  var api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.AttCore = api;
})(typeof window !== 'undefined' ? window : globalThis, function () {
  'use strict';

  /* ---------------- 基础工具 ---------------- */

  function norm(s) {
    if (s == null) return '';
    var t = String(s);
    var out = '';
    for (var i = 0; i < t.length; i++) {
      var c = t.charCodeAt(i);
      if (c >= 0xFF01 && c <= 0xFF5E) out += String.fromCharCode(c - 0xFEE0); // 全角→半角
      else if (c === 0x3000) out += ' ';
      else out += t.charAt(i);
    }
    return out.replace(/\s+/g, '').toLowerCase();
  }

  function toStr(v) { return v == null ? '' : String(v).trim(); }

  function round1(n) { return Math.round(n * 10) / 10; }
  function round2(n) { return Math.round(n * 100) / 100; }

  function fmt1(n) { return (Math.round(n * 10) / 10).toFixed(1); }
  function fmt2(n) { return (Math.round(n * 100) / 100).toFixed(2); }

  /* "06:43" / Date / 时间小数 → 当日分钟数；失败返回 null */
  function parseHM(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) {
      if (v.getFullYear() <= 1900) {
        // 纯时间单元格：SheetJS 以 UTC 构造，必须用 UTC 取值，否则时区会偏移
        var uh = v.getUTCHours(), um = v.getUTCMinutes(), us = v.getUTCSeconds();
        return uh * 60 + um + (us >= 30 ? 1 : 0);
      }
      // 带日期的时间（如打卡记录）也只取时分
      return v.getHours() * 60 + v.getMinutes() + (v.getSeconds() >= 30 ? 1 : 0);
    }
    if (typeof v === 'number' && isFinite(v)) {
      if (v < 1) return Math.round(v * 1440);               // Excel 时间小数
      if (v < 1440) return Math.round(v);                   // 已经是分钟数
      return null;
    }
    var t = String(v).trim();
    if (t.indexOf('缺卡') >= 0) return '缺卡';
    var m = t.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (m) {
      var mm = parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
      if (m[3]) mm += parseInt(m[3], 10) >= 30 ? 1 : 0;
      return mm;
    }
    return null;
  }

  function isMiss(v) { return v === '缺卡'; }

  /* 考勤差异明细纳入规则（纯函数，可单测）：
     1) 签到、签退都缺卡 → 一律排除（最高优先级，即使考勤结果非正常）
     2) 考勤结果非'正常'（迟到/早退/旷工/休等）→ 纳入
     3) 恰好一个缺卡 → 纳入
     4) 其余不纳入 */
  function isDiff(r) {
    var res = toStr(r.result), cinM = isMiss(r.cin), coutM = isMiss(r.cout);
    if (cinM && coutM) return false;
    if (res && res !== '正常') return true;
    if (cinM && !coutM) return true;
    if (!cinM && coutM) return true;
    return false;
  }

  /* 分钟 → "HH:MM" */
  function minToHM(v) {
    if (v == null || v === '' || v === '缺卡') return v == null ? '' : v;
    var m = ((Math.round(v) % 1440) + 1440) % 1440;
    var h = Math.floor(m / 60), mm = m % 60;
    return (h < 10 ? '0' : '') + h + ':' + (mm < 10 ? '0' : '') + mm;
  }

  /* 日期统一为 'yyyy-mm-dd'；失败 null */
  function parseDateKey(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date && !isNaN(v.getTime())) {
      var y = v.getFullYear(), m = v.getMonth() + 1, d = v.getDate();
      return y + '-' + (m < 10 ? '0' : '') + m + '-' + (d < 10 ? '0' : '') + d;
    }
    if (typeof v === 'number' && isFinite(v) && v > 20000 && v < 80000) {
      var ms = Math.round((v - 25569) * 86400 * 1000);
      return parseDateKey(new Date(ms));
    }
    var t = String(v).trim();
    var m = t.match(/^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})/);
    if (m) return m[1] + '-' + pad2(m[2]) + '-' + pad2(m[3]);
    return null;
  }
  function pad2(s) { s = String(s); return s.length < 2 ? '0' + s : s; }

  /* 班次 "07:00-11:00" / "7:00-11:00" → {start,end}（end<start 视为跨午夜，end+1440）；失败 null */
  function parseShift(v) {
    var t = toStr(v);
    if (!t || t.indexOf('-') < 0) return null;
    var parts = t.split('-');
    if (parts.length < 2) return null;
    var a = parseHM(parts[0]), b = parseHM(parts[1]);
    if (typeof a !== 'number' || typeof b !== 'number') return null;
    if (b <= a) b += 1440;                                  // 跨午夜（如 20:00-06:00）
    return { start: a, end: b };
  }

  function shiftHours(sh) {
    if (!sh) return null;
    var e = sh.end <= sh.start ? sh.end + 1440 : sh.end;      // 跨天班次（如 20:00-04:00）
    return (e - sh.start) / 60;
  }

  /* 跨午夜班次的签退归一：
     ① 班次跨天（下班<上班）且签退早于上班时间 → +24h
     ② 深夜班次（下班≥22:00）签退落在凌晨（与下班差超12小时）→ +24h
     避免 23:59 下班、凌晨 00:10 打卡被误算成早退 1429 分钟 */
  function adjCout(cout, sh) {
    if (typeof cout !== 'number' || !sh) return cout;
    if (sh.end < sh.start && cout < sh.start) return cout + 1440;
    if (sh.end >= 1320 && sh.end - cout > 720) return cout + 1440;
    return cout;
  }

  /* 蝶城名：文件名截取到"蝶城/蝶外"为止 */
  function detectCity(fileName) {
    var base = String(fileName || '').replace(/\.[^.]+$/, '');
    var m = base.match(/^(.*?(?:蝶城|蝶外))/);
    if (m) return m[1];
    m = base.match(/^(.{2,12}?)(?:\d{4}|\d{2}月|[0-9])/);
    if (m) return m[1];
    return base;
  }

  /* 迟到分钟：签到-上班；[-30,0] 容差→0；缺卡/空→0 */
  function calcLate(cin, start) {
    if (start == null) return 0;
    if (cin == null || isMiss(cin)) return 0;
    var d = cin - start;
    if (d <= 0 && d >= -30) return 0;
    return Math.round(d);
  }

  /* 早退分钟：下班-签退（跨午夜班次签退自动+24h）；[-30,0] 容差→0；缺卡/空→0 */
  function calcEarly(cout, sh) {
    if (!sh) return 0;
    if (cout == null || isMiss(cout)) return 0;
    var d = sh.end - adjCout(cout, sh);
    if (d <= 0 && d >= -30) return 0;
    return Math.round(d);
  }

  /* 扣款档：≤5 不扣；5~15 扣10；15~120 扣20；>120 不扣 */
  function deduction(mins) {
    if (typeof mins !== 'number' || !isFinite(mins)) return 0;
    if (mins <= 5) return 0;
    if (mins <= 15) return 10;
    if (mins <= 120) return 20;
    return 0;
  }

  /* dice 相似度（模糊匹配姓名用） */
  function dice(a, b) {
    if (a === b) return 1;
    if (!a || !b || a.length < 2 || b.length < 2) return 0;
    if (a.length + b.length > 40) return 0;
    var map = {}, i, t;
    for (i = 0; i < a.length - 1; i++) { t = a.substr(i, 2); map[t] = (map[t] || 0) + 1; }
    var hit = 0;
    for (i = 0; i < b.length - 1; i++) { t = b.substr(i, 2); if (map[t]) { hit++; map[t]--; } }
    return 2 * hit / (a.length - 1 + b.length - 1);
  }

  function digitsOnly(s) { return String(s || '').replace(/[^0-9]/g, ''); }

  /* ---------------- 配置 ---------------- */

  var DEFAULT_CONFIG = {
    rosterMap: { code: '', name: '', project: '', city: '', hire: '', leave: '' }, // 花名册列名映射
    machineShifts: [],        // 复选框勾选的剔除岗位（SHR 中这些岗位整行剔除）
    machineExtraPosts: '',    // 「其他岗位名」输入框原文，逗号/空格/顿号/分号分隔
    machineExtraNames: '',    // 「员工姓名」输入框原文，逗号/空格/顿号/分号分隔
    machineExtraText: ''      // 旧版单一「其他」输入框（兼容旧配置，按岗位名+姓名都试）
  };

  /* ---------------- 花名册 ---------------- */

  var ROSTER_FIELDS = [
    { key: 'code', label: '员工编码/工号', aliases: ['员工编码', '工号', '员工编号', '员工号', '员工ID', '账号', '员工账号'] },
    { key: 'name', label: '姓名', aliases: ['员工姓名', '姓名', '名字'] },  // 花名册常同时有「姓名」「员工姓名」两列，员工姓名才是真实姓名
    { key: 'project', label: '项目', aliases: ['项目', '项目名称', '所属项目', '服务项目'] },
    { key: 'city', label: '蝶城/组织', aliases: ['蝶城', '蝶城名称', '所属蝶城', '组织', '所属组织', '区域', '城市'] },
    { key: 'hire', label: '入职日期', aliases: ['入职日期', '入职时间', '入职'] },
    { key: 'leave', label: '离职日期', aliases: ['离职日期', '离职时间', '离职'] }
  ];

  /* 根据表头猜花名册字段映射 {key: 列名} */
  function guessRosterMap(headers) {
    var map = {}, used = {};
    ROSTER_FIELDS.forEach(function (f) {
      for (var i = 0; i < f.aliases.length; i++) {
        var hit = null;
        headers.forEach(function (h) {
          if (hit) return;
          var n = norm(h);
          if (n && n === norm(f.aliases[i]) && !used[n]) hit = h;
        });
        if (hit == null) {
          headers.forEach(function (h) {
            if (hit) return;
            var n = norm(h);
            if (n && n.indexOf(norm(f.aliases[i])) >= 0 && !used[n]) hit = h;
          });
        }
        if (hit != null) { map[f.key] = hit; used[norm(hit)] = 1; return; }
      }
    });
    return map;
  }

  /* 解析花名册 → 行数组 [{code,name,project,city,hire,leave}] */
  function parseRoster(wb, map) {
    var rows = [];
    wb.SheetNames.forEach(function (sn) {
      var ws = wb.Sheets[sn];
      if (!ws || !ws['!ref']) return;
      var range = ws['!ref'];
      var R = lastRowOf(range);
      // 表头行：第一行能在前 30 列内认出 >=2 个映射字段的行
      var headerRow = -1, headers = [];
      for (var r = 1; r <= Math.min(R, 6); r++) {
        var vals = [], hit = 0;
        for (var c = 1; c <= 40; c++) {
          var cell = ws[String.fromCharCode(64 + Math.min(c, 26)) + r] || (c > 26 ? null : null);
          vals.push(toStr(cell && cell.v != null ? cell.v : (cell ? cell.w : '')));
        }
        var mapped = Object.keys(map).filter(function (k) { return map[k] && vals.indexOf(toStr(map[k])) >= 0; });
        if (mapped.length >= 2) { headerRow = r; headers = vals; break; }
      }
      if (headerRow < 0) return;
      var colIdx = {};
      headers.forEach(function (h, i) { colIdx[h] = i; });
      for (var rr = headerRow + 1; rr <= R; rr++) {
        function gv(colName) {
          if (!colName) return '';
          var ci = colIdx[colName];
          if (ci == null) return '';
          var addr = xlsxAddr(ci, rr);
          var cell = ws[addr];
          return cell ? toStr(cell.v != null ? cell.v : cell.w) : '';
        }
        var rec = {
          code: gv(map.code), name: gv(map.name), project: gv(map.project),
          city: gv(map.city), hire: parseDateKey(gv(map.hire)) || toStr(gv(map.hire)),
          leave: parseDateKey(gv(map.leave)) || toStr(gv(map.leave)),
          raw: true
        };
        if (!rec.code && !rec.name) continue;
        if (rec.hire && !/^\d{4}-/.test(rec.hire)) rec.hire = '';
        if (rec.leave && !/^\d{4}-/.test(rec.leave)) rec.leave = '';
        rows.push(rec);
      }
    });
    return rows;
  }

  /* 0 基列号 → A1 地址（SheetJS 键） */
  function xlsxAddr(col0, row1) {
    var n = col0 + 1, s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s + row1;
  }

  /* 'A1:P9' → 最后一行号 9 */
  function lastRowOf(range) {
    var m = String(range).match(/(\d+):[A-Z]+(\d+)$/);
    return m ? Number(m[2]) : 1;
  }

  /* ---------------- 考勤文件解析 ---------------- */

  var SHR_FIELDS = {
    name: ['员工姓名', '姓名'], code: ['员工编号', '编号', '工号', '员工编码'],
    post: ['岗位名称', '岗位'], date: ['日期', '考勤日期'], project: ['项目名称', '项目'],
    shift: ['班次时间', '班次'], cin: ['签到时间', '签到'], cout: ['签退时间', '签退'],
    theory: ['理论工时'], result: ['考勤结果'], abn: ['异常考勤'],
    buka: ['补卡工时'], payh: ['计薪时长'], acth: ['实际工时']
  };
  var RM_FIELDS = {
    code: ['账号', '员工编号', '工号'], name: ['姓名', '员工姓名'],
    post: ['岗位名称', '岗位'], date: ['日期', '考勤日期'], city: ['蝶城', '蝶城名称'],
    shift: ['班次时间', '班次'], cin: ['签到时间', '签到'], cout: ['签退时间', '签退'],
    acth: ['实际工时'], result: ['考勤结果'], abn: ['异常考勤'],
    payh: ['计薪时长'], buka: ['补卡工时'], absent: ['旷工'],
    ot1: ['工作日加班'], ot2: ['休息日加班'], ot3: ['法定节假日']
  };

  /* 解析一个考勤文件（type: 'shr' | 'rm'） */
  function parseAttFile(wb, fileName, type) {
    var city = detectCity(fileName);
    var sheetName = null, ws = null;
    if (type === 'shr') {
      wb.SheetNames.forEach(function (sn) { if (!ws && sn.indexOf('SHR员工考勤') >= 0) { ws = wb.Sheets[sn]; sheetName = sn; } });
      if (!ws) { ws = wb.Sheets[wb.SheetNames[0]]; sheetName = wb.SheetNames[0]; }
    } else {
      ws = wb.Sheets[wb.SheetNames[0]]; sheetName = wb.SheetNames[0];
    }
    if (!ws || !ws['!ref']) return { fileName: fileName, type: type, city: city, error: '没有可用工作表', headers: [], rows: [] };

    var range = ws['!ref'];
    var lastRow = lastRowOf(range);
    var colCount = colCountOf(range);
    var fields = type === 'shr' ? SHR_FIELDS : RM_FIELDS;

    // 表头行自动识别：前 6 行里命中字段数最多的行
    var bestRow = -1, bestHits = 0, bestVals = null;
    for (var r = 1; r <= Math.min(lastRow, 6); r++) {
      var vals = rowVals(ws, r, colCount), hits = 0;
      Object.keys(fields).forEach(function (k) {
        for (var i = 0; i < fields[k].length; i++) if (vals.indexOf(toStr(fields[k][i])) >= 0) { hits++; return; }
      });
      if (hits > bestHits) { bestHits = hits; bestRow = r; bestVals = vals; }
    }
    if (bestRow < 0 || bestHits < 5) {
      return { fileName: fileName, type: type, city: city, error: '表头识别失败（需要至少识别 5 个标准列）', headers: [], rows: [] };
    }

    var colMap = {};
    Object.keys(fields).forEach(function (k) {
      for (var i = 0; i < fields[k].length; i++) {
        var idx = bestVals.indexOf(toStr(fields[k][i]));
        if (idx >= 0) { colMap[k] = idx; return; }
      }
    });

    var headers = [];
    for (var c = 0; c < colCount; c++) headers.push(bestVals[c] || '');

    var rows = [];
    for (var rr = bestRow + 1; rr <= lastRow; rr++) {
      var raw = rowVals(ws, rr, colCount);
      var name = colMap.name != null ? toStr(raw[colMap.name]) : '';
      if (!name) continue;                                    // 姓名为空的行跳过
      var sh = parseShift(colMap.shift != null ? raw[colMap.shift] : '');
      rows.push({
        type: type,
        name: name,
        code: colMap.code != null ? toStr(raw[colMap.code]) : '',
        post: colMap.post != null ? toStr(raw[colMap.post]) : '',
        date: parseDateKey(colMap.date != null ? raw[colMap.date] : '') || '',
        project: colMap.project != null ? toStr(raw[colMap.project]) : '',
        city: colMap.city != null ? toStr(raw[colMap.city]) : '',
        shiftRaw: colMap.shift != null ? toStr(raw[colMap.shift]) : '',
        sh: sh,
        cin: colMap.cin != null ? parseHM(raw[colMap.cin]) : null,
        cout: colMap.cout != null ? parseHM(raw[colMap.cout]) : null,
        theory: colMap.theory != null ? numOr(raw[colMap.theory]) : null,
        result: colMap.result != null ? toStr(raw[colMap.result]) : '',
        abn: colMap.abn != null ? toStr(raw[colMap.abn]) : '',
        buka: colMap.buka != null ? numOr(raw[colMap.buka]) : 0,
        payh: colMap.payh != null ? numOr(raw[colMap.payh]) : null,
        acth: colMap.acth != null ? numOr(raw[colMap.acth]) : null,
        absent: colMap.absent != null ? raw[colMap.absent] : '',
        ot1: colMap.ot1 != null ? raw[colMap.ot1] : '',
        ot2: colMap.ot2 != null ? raw[colMap.ot2] : '',
        ot3: colMap.ot3 != null ? raw[colMap.ot3] : ''
      });
    }
    return { fileName: fileName, type: type, city: city, sheetName: sheetName, headers: headers, rows: rows };
  }

  function rowVals(ws, r, n) {
    var out = [];
    for (var c = 0; c < n; c++) {
      var cell = ws[xlsxAddr(c, r)];
      if (!cell) { out.push(''); continue; }
      if (cell.v instanceof Date) out.push(cell.v);
      else out.push(cell.v != null ? cell.v : toStr(cell.w));
    }
    return out;
  }
  function numOr(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return v;
    var n = parseFloat(String(v).replace(/[^\d.\-]/g, ''));
    return isNaN(n) ? null : n;
  }
  function colCountOf(range) {
    var m = range.match(/^[A-Z]+(\d+):([A-Z]+)(\d+)$/);
    if (!m) return 30;
    var s = 0, t = m[2];
    for (var i = 0; i < t.length; i++) s = s * 26 + (t.charCodeAt(i) - 64);
    return s;
  }

  /* ---------------- 主处理 ---------------- */

  var UNI_HEAD = ['姓名', '员工编码', '岗位', '日期', '项目', '班次时间', '签到时间', '签退时间', '理论工时', '实际工时', '考勤结果', '异常考勤', '计薪时长', '补卡工时', '工时', '旷工', '迟到(分钟)', '早退(分钟)', '迟到扣款', '早退扣款', '扣款合计', '工作日加班', '休息日加班', '法定节假日'];

  /*
   * files: [{type:'shr'|'rm', parsed:{rows,...}}]
   * rosterRows: parseRoster 结果
   * cfg: {rosterMap(已不需要，解析时用), machineShifts:[岗位名...]}
   */
  /* 推断一个考勤文件的蝶城：
     ① 源表记录自带蝶城列 → 取众数；
     ② 无蝶城列 → 按姓名匹配花名册，取花名册蝶城众数；
     ③ 兜底退回文件名截取 */
  function inferFileCity(parsed, rosterIdx) {
    var rows = parsed.rows || [];
    function majority(counter, raws) {
      var bk = null, bn = 0;
      Object.keys(counter).forEach(function (k) { if (counter[k] > bn) { bn = counter[k]; bk = k; } });
      return bk != null ? raws[bk] : null;
    }
    var cnt = {}, raws = {}, withCity = 0;
    rows.forEach(function (r) {
      var c = toStr(r.city);
      if (!c) return;
      withCity++;
      var k = norm(c);
      cnt[k] = (cnt[k] || 0) + 1;
      raws[k] = c;
    });
    if (withCity > 0) {
      var m1 = majority(cnt, raws);
      if (m1) return { city: m1, multi: Object.keys(cnt).length > 1, how: '源表蝶城列' };
    }
    var cnt2 = {}, raws2 = {}, matched = 0;
    rows.forEach(function (r) {
      var m = matchRoster(rosterIdx, r.code, r.name, '');
      if (m && toStr(m.row.city)) {
        matched++;
        var k = norm(m.row.city);
        cnt2[k] = (cnt2[k] || 0) + 1;
        raws2[k] = toStr(m.row.city);
      }
    });
    if (matched > 0) {
      var m2 = majority(cnt2, raws2);
      if (m2) return { city: m2, multi: Object.keys(cnt2).length > 1, how: '花名册匹配' };
    }
    return { city: parsed.city || detectCity(parsed.fileName), multi: false, how: '文件名' };
  }

  function processAll(files, rosterRows, cfg) {
    cfg = cfg || {};
    var problems = [];
    function addProblem(city, sheet, name, msg, sev) {
      problems.push({ city: city || '', sheet: sheet || '', name: name || '', msg: msg, sev: sev || 'warn' });
    }

    var rosterIdx = buildRosterIndex(rosterRows || []);

    // 1. 按蝶城分组 —— 记录级：一条记录一个蝶城
    //    一个原始导出文件可能混多个蝶城（如 RM 原始表含全部蝶城的机师岗），
    //    每条记录按：记录自身蝶城列 → 花名册匹配蝶城 → 文件推断蝶城 兜底；
    //    再把互相包含的蝶城名归并到最长名（如「钟公庙」→「宁波钟公庙调度中心」）
    var inferredCities = [];
    files.forEach(function (f) {
      if (f.parsed.error) return;
      f.parsed._inf = inferFileCity(f.parsed, rosterIdx);
      inferredCities.push(f.parsed._inf.city);
    });
    var namePool = inferredCities.slice();
    (rosterIdx.rows || []).forEach(function (r) { if (r.city) namePool.push(r.city); });
    var canonCache = {};
    function canonCity(name) {
      var key = String(name == null ? '' : name);
      if (key in canonCache) return canonCache[key];
      var best = name;
      namePool.forEach(function (n) {
        if (n && cityMatch(name, n) && String(n).length > String(best).length) best = n;
      });
      canonCache[key] = best;
      return best;
    }

    var byCity = {};
    files.forEach(function (f) {
      if (f.parsed.error) { addProblem(detectCity(f.parsed.fileName), '', f.parsed.fileName, f.parsed.error, 'error'); return; }
      f.parsed.rows.forEach(function (r) {
        var rc = toStr(r.city);
        if (!rc) {
          var m = matchRoster(rosterIdx, r.code, r.name, '');
          if (m && toStr(m.row.city)) rc = toStr(m.row.city);
        }
        var c = canonCity(rc || f.parsed._inf.city);
        if (!byCity[c]) byCity[c] = { shrRows: [], rmRows: [] };
        byCity[c][f.parsed.type === 'shr' ? 'shrRows' : 'rmRows'].push(r);
      });
    });

    // 全局考勤索引：在职无出勤要匹配外包+机师岗所有考勤记录（不限于本蝶城文件）
    var globalAttRows = [];
    files.forEach(function (f) { if (!f.parsed.error) globalAttRows = globalAttRows.concat(f.parsed.rows || []); });
    var globalAttIdx = buildAttIdx(globalAttRows);

    var cities = [];
    Object.keys(byCity).sort().forEach(function (city) {
      cities.push(processCity(city, byCity[city].shrRows, byCity[city].rmRows, rosterIdx, cfg, addProblem, globalAttIdx));
    });

    // 2. 跨文件汇总总表
    var summary = buildSummary(cities, addProblem);
    // 问题清单只保留在职无出勤明细与核算/匹配类问题，剔除 info 级记录（打卡校准、机师岗剔除等）
    problems = problems.filter(function (p) { return p.sev !== 'info'; });
    return { cities: cities, summary: summary, problems: problems };
  }

  function buildRosterIndex(rows) {
    var byCode = {}, byName = {}, byNameList = {};
    rows.forEach(function (r) {
      var c = norm(r.code), n = norm(r.name);
      if (c && !byCode[c]) byCode[c] = r;
      var dc = digitsOnly(c);
      if (dc && !byCode['#' + dc]) byCode['#' + dc] = r;
      if (n) {
        if (!byName[n]) byName[n] = r;
        byNameList[n] = (byNameList[n] || 0) + 1;
      }
    });
    return { rows: rows, byCode: byCode, byName: byName, byNameList: byNameList };
  }

  /* 花名册匹配：编码精确 → 编码去前缀/前导零 → 姓名精确(唯一) → 蝶城+姓名精确 → 蝶城内姓名模糊 → 全局姓名模糊；
     返回 {row, conf} 或 null */
  function matchRoster(idx, code, name, city) {
    var c = norm(code), n = norm(name);
    if (c && idx.byCode[c]) return { row: idx.byCode[c], conf: 'code' };
    var dc = digitsOnly(c);
    if (dc && idx.byCode['#' + dc]) return { row: idx.byCode['#' + dc], conf: 'code-loose' };
    if (n && idx.byName[n]) {
      if (idx.byNameList[n] === 1) return { row: idx.byName[n], conf: 'name' };
    }
    // 蝶城+姓名 兜底：同名多人时按蝶城缩小范围（源表无员工编码时的主要匹配方式）
    if (n && city) {
      var cands = idx.rows.filter(function (r) {
        return norm(r.name) === n && cityMatch(r.city, city);
      });
      if (cands.length === 1) return { row: cands[0], conf: 'name-city' };
      if (cands.length > 1) return null;                       // 同蝶城还有同名：不敢乱配
    }
    if (n) {
      var pool = city ? idx.rows.filter(function (r) { return cityMatch(r.city, city); }) : idx.rows;
      var searchRows = pool.length ? pool : idx.rows;
      var best = null, bestScore = 0;
      searchRows.forEach(function (r) {
        var s = dice(n, norm(r.name));
        if (s > bestScore) { bestScore = s; best = r; }
      });
      if (best && bestScore >= 0.85) return { row: best, conf: city ? 'name-city-fuzzy' : 'name-fuzzy' };
    }
    return null;
  }

  /* 反向：花名册里的人是否出现在考勤里 */
  function inAttendance(idx, attIdx, r) {
    var c = norm(r.code), n = norm(r.name);
    if (c && attIdx.byCode[c]) return true;
    var dc = digitsOnly(c);
    if (dc && attIdx.byCode['#' + dc]) return true;
    if (n && attIdx.byName[n]) return true;
    if (n) {
      // 原逻辑：遍历全部考勤行，dice(n, 考勤归一化名)>=0.85 且 显示名完全相同 才算命中。
      // 既然显示名必须完全相同，其归一化名也相同，dice 必为 1（>=0.85），
      // 因此等价于「考勤里存在显示名完全相同的记录」。改为按显示名建索引 O(1) 命中，结果一致。
      var disp = toStr(r.name);
      if (attIdx.byDisplay && attIdx.byDisplay[disp]) return true;
    }
    return false;
  }

  function buildAttIdx(rows) {
    var byCode = {}, byName = {}, byDisplay = {}, nameList = [], seenName = {};
    rows.forEach(function (r) {
      var c = norm(r.code), n = norm(r.name), disp = toStr(r.name);
      if (c && !byCode[c]) byCode[c] = 1;
      var dc = digitsOnly(c);
      if (dc && !byCode['#' + dc]) byCode['#' + dc] = 1;
      if (n && !byName[n]) byName[n] = 1;
      if (disp && !byDisplay[disp]) byDisplay[disp] = 1;
      // nameList 仅供外部读取/兼容；inAttendance 已改用 byDisplay，此处仍去重降低内存
      if (n) {
        var key = n + '\u0001' + disp;
        if (!seenName[key]) { seenName[key] = 1; nameList.push({ n: n, name: r.name }); }
      }
    });
    return { byCode: byCode, byName: byName, byDisplay: byDisplay, nameList: nameList };
  }

  /* 「其他」输入框：逗号 / 空格 / 顿号 / 分号分隔，去重后返回原词 */
  function splitTokens(s) {
    var out = [], seen = {};
    String(s == null ? '' : s).split(/[,，、;；\s]+/).forEach(function (t) {
      var v = t.trim();
      if (!v) return;
      var k = norm(v);
      if (!k || seen[k]) return;
      seen[k] = 1; out.push(v);
    });
    return out;
  }

  /* 外包考勤剔除判定（三个来源各自独立）：
   *   machineShifts      复选框勾选的岗位        → 只按岗位匹配
   *   machineExtraPosts  「其他岗位名」输入框    → 只按岗位匹配
   *   machineExtraNames  「员工姓名」输入框      → 只按姓名匹配
   *   machineExtraText   旧版单一「其他」输入框  → 兼容旧配置，岗位名与姓名都试 */
  function makeExcludeFilter(cfg) {
    var postSet = {}, nameSet = {};
    (cfg.machineShifts || []).forEach(function (p) { if (p) postSet[norm(p)] = 1; });
    splitTokens(cfg.machineExtraPosts).forEach(function (t) { postSet[norm(t)] = 1; });
    splitTokens(cfg.machineExtraNames).forEach(function (t) { nameSet[norm(t)] = 1; });
    splitTokens(cfg.machineExtraText).forEach(function (t) {   // 旧配置兼容
      postSet[norm(t)] = 1; nameSet[norm(t)] = 1;
    });
    return function (r) {
      if (r.post && postSet[norm(r.post)]) return true;
      if (r.name && nameSet[norm(r.name)]) return true;
      return false;
    };
  }

  function processCity(city, shrRowsIn, rmRowsIn, rosterIdx, cfg, addProblem, globalAttIdx) {
    var isExcluded = makeExcludeFilter(cfg);

    var allRows = [];

    /* --- SHR：剔除指定岗位 / 指定员工（不写入问题清单） --- */
    shrRowsIn.forEach(function (r) {
      if (isExcluded(r)) return;
      allRows.push(r);
    });
    rmRowsIn.forEach(function (r) { allRows.push(r); });
    allRows.sort(function (a, b) {
      var d = (a.date || '').localeCompare(b.date || '');
      if (d) return d;
      return toStr(a.name).localeCompare(toStr(b.name), 'zh');
    });

    /* --- 花名册匹配（补项目/蝶城 + 矩阵项目/入职/离职） --- */
    var matchCache = {};
    function rosterOf(r) {
      var key = norm(r.code) + '|' + norm(r.name) + '|' + norm(r.city);
      if (!(key in matchCache)) {
        var m = matchRoster(rosterIdx, r.code, r.name, r.city || city);
        matchCache[key] = m || null;
        if (!m) {
          addProblem(city, r.type === 'shr' ? 'SHR员工考勤' : 'RM考勤', r.name,
            '花名册中未匹配到该员工（编码 ' + (r.code || '空') + '、蝶城 ' + (r.city || city || '未知') + '），「项目」留空、蝶城用源表值', 'warn');
        } else if (m.conf === 'name-fuzzy' || m.conf === 'name-city-fuzzy') {
          addProblem(city, '', r.name, '花名册按姓名模糊匹配（相似度≥0.85），请人工核对：' + m.row.name + ' / ' + m.row.city + ' / ' + m.row.project, 'warn');
        }
      }
      return matchCache[key];
    }
    /* 记录生效蝶城：花名册蝶城优先，其次源表蝶城列，最后文件级蝶城 */
    function effCity(r) {
      var m = rosterOf(r);
      if (m && toStr(m.row.city)) return toStr(m.row.city);
      return r.city || city;
    }

    /* --- 打卡校准（SHR + RM 统一）：同一天同人按班次开始排序 --- */
    var dayGroups = {};
    allRows.forEach(function (r) {
      if (!r.date || !r.sh) return;
      var k = r.type + '|' + r.date + '|' + norm(r.name) + '|' + norm(r.code);
      (dayGroups[k] = dayGroups[k] || []).push(r);
    });
    Object.keys(dayGroups).forEach(function (k) {
      var g = dayGroups[k].slice().sort(function (a, b) { return a.sh.start - b.sh.start; });
      for (var i = 0; i < g.length; i++) {
        var ri = g[i];
        if (typeof ri.cout !== 'number' || !ri.sh) continue;
        if (adjCout(ri.cout, ri.sh) - ri.sh.end <= 30) continue; // 下班卡没超 30 分钟：正常（跨午夜签退已归一）
        var reassigned = false;
        for (var j = i + 1; j < g.length; j++) {
          var rj = g[j];
          if (!rj.sh) continue;
          if (!(rj.cin == null || isMiss(rj.cin))) continue;   // 下一班次签到已有卡：不动
          if (Math.abs(ri.cout - rj.sh.start) <= 30) {          // 落在下一班次上班±30分钟内 → 改归属
            addProblem(city, rj.type === 'shr' ? 'SHR员工考勤' : 'RM考勤', ri.name,
              '打卡校准：' + ri.date + ' ' + minToHM(ri.cout) +
              ' 原被记为「' + ri.shiftRaw + '」的下班卡，已归还给「' + rj.shiftRaw + '」的上班卡', 'info');
            rj.cin = ri.cout; rj.calIn = true;
            ri.cout = '缺卡'; ri.calOut = true;
            reassigned = true;
            break;
          }
        }
        if (!reassigned) {
          // 超过本班次下班30分钟以上、又不能归还下一班次：视作无效打卡，本班次记缺卡
          addProblem(city, ri.type === 'shr' ? 'SHR员工考勤' : 'RM考勤', ri.name,
            '打卡校准：' + ri.date + ' ' + minToHM(typeof ri.cout === 'number' ? ri.cout : 0) +
            ' 超过「' + ri.shiftRaw + '」下班时间30分钟以上，视作无效打卡，本班次签退记缺卡', 'info');
          ri.cout = '缺卡'; ri.calOut = true;
        }
      }
    });

    /* --- 全表重算迟到/早退/扣款；校准行连带重算派生列 --- */
    allRows.forEach(function (r) {
      r.late = calcLate(r.cin, r.sh ? r.sh.start : null);
      r.early = calcEarly(r.cout, r.sh);
      r.dedL = deduction(r.late);
      r.dedE = deduction(r.early);
      r.dedT = r.dedL + r.dedE;
      if (!(r.calIn || r.calOut)) return;

      var theory = (typeof r.theory === 'number' && r.theory > 0) ? r.theory : round1(shiftHours(r.sh) || 0);
      var cinMiss = r.cin == null || isMiss(r.cin);
      var coutMiss = r.cout == null || isMiss(r.cout);
      if (cinMiss || coutMiss) {
        var half = theory > 0 ? theory / 2 : 0;
        r.result = '旷工';
        r.abn = '旷工' + fmt1(half) + '小时';
        r.payh = round2(half);
        r.acth = r.type === 'rm' ? 0 : r.acth;
      } else {
        r.payh = round2(Math.max(0, Math.min(theory - r.late / 60 - r.early / 60, theory)));
        var coutAdj = adjCout(r.cout, r.sh);                    // 跨午夜下班卡归一
        if (r.type === 'rm') r.acth = round2((coutAdj - r.cin) / 60);
        if (r.late > 0 && r.early > 0) { r.result = '迟到'; r.abn = '迟到' + fmt2(r.late / 60) + '小时'; }
        else if (r.late > 0) { r.result = '迟到'; r.abn = '迟到' + fmt2(r.late / 60) + '小时'; }
        else if (r.early > 0) { r.result = '早退'; r.abn = '早退' + fmt2(r.early / 60) + '小时'; }
        else { r.result = '正常'; r.abn = '-'; }
      }
      if (r.type === 'shr') r.gongshi = round2((r.buka || 0) + (r.payh || 0));
      else r.gongshi = null;
    });
    /* 非校准行也补算 工时 = 补卡+计薪（SHR 输出列需要） */
    allRows.forEach(function (r) {
      if (r.type === 'shr' && r.gongshi == null) r.gongshi = round2((r.buka || 0) + (r.payh || 0));
    });

    /* --- 期间 --- */
    var minD = null, maxD = null;
    allRows.forEach(function (r) {
      if (!r.date) return;
      if (minD == null || r.date < minD) minD = r.date;
      if (maxD == null || r.date > maxD) maxD = r.date;
    });

    /* --- 主表行 --- */
    var shrRows = [], rmRows = [];
    allRows.forEach(function (r) {
      var m = rosterOf(r);
      // 项目：花名册匹配优先（源表无员工编码时按 蝶城+姓名 兜底匹配），匹配不到再用源表自己的
      var proj = (m && toStr(m.row.project)) ? toStr(m.row.project) : (r.project || '');
      var cinHM = isMiss(r.cin) ? '缺卡' : minToHM(r.cin);
      var coutHM = isMiss(r.cout) ? '缺卡' : minToHM(r.cout);
      if (r.type === 'shr') {
        shrRows.push([
          r.name, r.code, r.post, r.date, proj, r.shiftRaw,
          cinHM, coutHM,
          typeof r.theory === 'number' ? round1(r.theory) : r.theory,
          r.result || '', r.abn || '',
          r.buka == null ? '' : round1(r.buka),
          r.payh == null ? '' : round2(r.payh),
          r.gongshi == null ? '' : round2(r.gongshi),
          r.sh ? minToHM(r.sh.start) : '', r.sh ? minToHM(r.sh.end % 1440 || r.sh.end) : '',
          r.late, r.early, r.dedL, r.dedE, r.dedT
        ]);
      } else {
        rmRows.push([
          r.code, r.name, r.post, r.date,
          proj,                                         // 花名册补的项目
          effCity(r), r.shiftRaw, cinHM, coutHM,
          r.acth == null ? '' : round2(r.acth),
          r.result || '', r.abn || '',
          r.payh == null ? '' : round2(r.payh),
          r.buka == null ? '' : round1(r.buka),
          r.absent, r.late, r.early, r.ot1, r.ot2, r.ot3,
          r.sh ? minToHM(r.sh.start) : '', r.sh ? minToHM(r.sh.end) : '',
          r.dedL, r.dedE, r.dedT
        ]);
      }
    });
    shrRows.sort(function (a, b) { var d = toStr(a[0]).localeCompare(toStr(b[0]), 'zh'); return d || toStr(a[3]).localeCompare(toStr(b[3])); });
    rmRows.sort(function (a, b) { var d = toStr(a[1]).localeCompare(toStr(b[1]), 'zh'); return d || toStr(a[3]).localeCompare(toStr(b[3])); });

    /* --- 考勤差异明细（合并两类，纳入规则见模块级 isDiff） --- */
    var UNI_HEAD_ = UNI_HEAD;
    function toUni(r) {
      var m = rosterOf(r);
      var proj = (m && toStr(m.row.project)) ? toStr(m.row.project) : (r.project || '');
      var cinHM = isMiss(r.cin) ? '缺卡' : minToHM(r.cin);
      var coutHM = isMiss(r.cout) ? '缺卡' : minToHM(r.cout);
      return [
        r.name, r.code, r.post, r.date,
        proj,
        r.shiftRaw, cinHM, coutHM,
        typeof r.theory === 'number' ? round1(r.theory) : (r.type === 'rm' && r.sh ? round1(shiftHours(r.sh)) : ''),
        r.acth == null ? '' : round2(r.acth),
        r.result || '', r.abn || '',
        r.payh == null ? '' : round2(r.payh),
        r.buka == null ? '' : round1(r.buka),
        r.gongshi == null ? '' : round2(r.gongshi),
        r.type === 'rm' ? r.absent : '',
        r.late, r.early, r.dedL, r.dedE, r.dedT, r.ot1, r.ot2, r.ot3
      ];
    }
    var diffRows = allRows.filter(isDiff).map(toUni);
    diffRows.sort(function (a, b) {
      var d = toStr(a[2]).localeCompare(toStr(b[2]), 'zh');
      if (d) return d;
      d = toStr(a[0]).localeCompare(toStr(b[0]), 'zh');
      return d || toStr(a[3]).localeCompare(toStr(b[3]));
    });

    /* --- 迟到早退明细 --- */
    var detailRows = allRows.filter(function (r) { return r.late !== 0 || r.early !== 0; }).map(toUni);
    detailRows.sort(function (a, b) {
      var d = toStr(a[0]).localeCompare(toStr(b[0]), 'zh');
      return d || toStr(a[3]).localeCompare(toStr(b[3]));
    });

    /* --- 迟到早退统计（13 列，宏口径：只导有迟到早退的人 + 合计行） --- */
    function buildStat(rows) {
      var statMap = {}, statOrder = [];
      rows.forEach(function (r) {
        var key = toStr(r.name) + '|' + toStr(r.code);
        if (!statMap[key]) {
          var m = rosterOf(r);
          statMap[key] = {
            name: r.name, code: r.code,
            // 编码后加项目、项目后为岗位：项目取花名册优先、源表兜底；岗位取考勤行岗位（两类都有岗位列）
            project: (m && toStr(m.row.project)) ? toStr(m.row.project) : toStr(r.project || ''),
            post: toStr(r.post || ''),
            l5: 0, l15: 0, l120: 0, e5: 0, e15: 0, e120: 0, ded: 0
          };
          statOrder.push(key);
        }
        var s = statMap[key];
        if (r.late > 0 && r.late <= 5) s.l5++;
        else if (r.late > 5 && r.late <= 15) s.l15++;
        else if (r.late > 15 && r.late <= 120) s.l120++;
        if (r.early > 0 && r.early <= 5) s.e5++;
        else if (r.early > 5 && r.early <= 15) s.e15++;
        else if (r.early > 15 && r.early <= 120) s.e120++;
        s.ded += r.dedT;
      });
      var out = [];
      statOrder.forEach(function (k) {
        var s = statMap[k];
        var total = s.l5 + s.l15 + s.l120 + s.e5 + s.e15 + s.e120;
        if (total <= 0) return;
        out.push([s.name, s.code, s.project, s.post, s.l5, s.l15, s.l120, s.e5, s.e15, s.e120,
          s.l5 + s.l15 + s.l120, s.e5 + s.e15 + s.e120, total, round2(s.ded)]);
      });
      return out;
    }
    var statRows = buildStat(allRows);

    /* --- 机师岗自己的迟到早退明细/统计（放进机师岗文件） --- */
    var rmOnlyRows = allRows.filter(function (r) { return r.type === 'rm'; });
    var detailRowsRm = rmOnlyRows.filter(function (r) { return r.late !== 0 || r.early !== 0; }).map(toUni);
    detailRowsRm.sort(function (a, b) {
      var d = toStr(a[0]).localeCompare(toStr(b[0]), 'zh');
      return d || toStr(a[3]).localeCompare(toStr(b[3]));
    });
    var statRowsRm = buildStat(rmOnlyRows);

    /* --- 工时矩阵（计薪工时/理论工时/打卡工时/工时比对） ---
       通用构建：rows=考勤行数组，产出按人×日期的 计薪/理论/打卡 三张矩阵
       计薪工时=照搬打卡折算（1班次0.5/卡、≥2班次0.25/卡、封顶1）
       打卡工时=（计薪时长+补卡工时+工作日加班+休息日加班+法定节假日）按天求和 */
    var days = monthDays(minD);
    function buildMatrices(rows) {
      var ppl = [], idx = {};
      rows.forEach(function (r) {
        var key = norm(r.code) + '|' + norm(r.name);
        if (!idx[key]) {
          var m = rosterOf(r);
          idx[key] = {
            name: r.name, code: r.code,
            project: m ? m.row.project : (r.project || ''),
            hire: m ? m.row.hire : '', leave: m ? m.row.leave : '',
            days: {}
          };
          ppl.push(key);
        }
        var p = idx[key];
        if (!r.date) return;
        var d = p.days[r.date] = p.days[r.date] || { cards: 0, shifts: 0, theory: 0, theoryN: 0, punch: 0, punchN: 0 };
        d.shifts++;
        if (typeof r.cin === 'number') d.cards++;
        if (typeof r.cout === 'number') d.cards++;
        var th = null;
        if (typeof r.theory === 'number' && r.theory > 0) th = r.theory;
        else if (r.sh) th = round1(shiftHours(r.sh));
        if (th != null) { d.theory += th; d.theoryN++; }
        var pv = numOr(r.payh), bv = numOr(r.buka);
        var o1 = numOr(r.ot1), o2 = numOr(r.ot2), o3 = numOr(r.ot3);
        if (pv != null || bv != null || o1 != null || o2 != null || o3 != null) {
          d.punch += (pv || 0) + (bv || 0) + (o1 || 0) + (o2 || 0) + (o3 || 0);
          d.punchN++;
        }
      });
      /* 矩阵表排序：项目为主Key、姓名为次Key */
      ppl.sort(function (a, b) {
        var pa = idx[a], pb = idx[b];
        var dd = toStr(pa.project).localeCompare(toStr(pb.project), 'zh');
        return dd || toStr(pa.name).localeCompare(toStr(pb.name), 'zh');
      });
      var pay = [], theory = [], punch = [];
      ppl.forEach(function (key) {
        var p = idx[key];
        var payRow = [p.project, p.hire, p.leave, p.name];
        var thRow = [p.project, p.hire, p.leave, p.name];
        var punRow = [p.project, p.hire, p.leave, p.name];
        days.forEach(function (dd) {
          var d = p.days[fullDate(minD, dd)];
          if (!d) { payRow.push(''); thRow.push(''); punRow.push(''); return; }
          var weight = d.shifts <= 1 ? 0.5 : 0.25;
          payRow.push(round2(Math.min(d.cards * weight, 1)));
          thRow.push(d.theoryN > 0 ? round1(d.theory) : '');
          punRow.push(d.punchN > 0 ? round2(d.punch) : '');
        });
        pay.push(payRow); theory.push(thRow); punch.push(punRow);
      });
      return { people: ppl, pay: pay, theory: theory, punch: punch };
    }
    /* 外包文件矩阵：全部人员（SHR剔除机师岗 + RM机师岗）汇总 */
    var allMat = buildMatrices(allRows);
    var people = allMat.people;
    var payMatrix = allMat.pay, theoryMatrix = allMat.theory;
    /* 机师岗文件矩阵：仅 RM 人员 */
    var rmMat = buildMatrices(rmOnlyRows);

    /* --- 在职无出勤（匹配外包+机师岗全部考勤记录） --- */
    var noAtt = [];
    (rosterIdx.rows || []).forEach(function (r) {
      if (!cityMatch(r.city, city)) return;
      var hire = toStr(r.hire), leave = toStr(r.leave);
      if (hire && hire > maxD) return;                        // 期末后入职
      if (leave && leave < minD) return;                      // 期初前离职
      if (inAttendance(rosterIdx, globalAttIdx, r)) return;
      noAtt.push([city, r.name, r.code, r.project, hire, leave]);
    });
    noAtt.sort(function (a, b) { return toStr(a[1]).localeCompare(toStr(b[1]), 'zh'); });
    /* 在职无出勤不进问题清单（问题清单只放系统计算问题），明细在结果看板展示 */

    var monthLabel = minD ? (parseInt(minD.slice(5, 7), 10) + '月') : '';
    var periodLabel = '';
    if (minD) periodLabel = maxD && maxD !== minD
      ? (minD.slice(0, 4) === maxD.slice(0, 4) ? minD + '~' + maxD.slice(5) : minD + '~' + maxD)
      : minD;
    return {
      city: city, periodStart: minD, periodEnd: maxD, monthLabel: monthLabel, periodLabel: periodLabel,
      shrRows: shrRows, rmRows: rmRows,
      diffRows: diffRows, detailRows: detailRows, statRows: statRows,
      detailRowsRm: detailRowsRm, statRowsRm: statRowsRm,
      payMatrix: payMatrix, theoryMatrix: theoryMatrix, matrixDays: days,
      rmPayMatrix: rmMat.pay, rmTheoryMatrix: rmMat.theory, rmPunchMatrix: rmMat.punch,
      noAttendance: noAtt,
      peopleCount: people.length,
      fileNameShr: '【考勤差异】' + city + periodLabel + '.xlsx',
      fileNameRm: '【考勤差异】' + city + '机师岗' + periodLabel + '.xlsx'
    };
  }

  function cityMatch(rosterCity, fileCity) {
    var a = norm(rosterCity), b = norm(fileCity);
    if (!a || !b) return false;
    return a.indexOf(b) >= 0 || b.indexOf(a) >= 0;
  }

  function monthDays(minD) {
    if (!minD) return [];
    var y = parseInt(minD.slice(0, 4), 10), m = parseInt(minD.slice(5, 7), 10);
    var n = new Date(y, m, 0).getDate();
    var out = [];
    for (var i = 1; i <= n; i++) out.push(i);
    return out;
  }
  function fullDate(minD, day) {
    return minD.slice(0, 7) + '-' + pad2(day);
  }

  /* ---------------- 汇总总表 ---------------- */

  var SUMMARY_STAT_HEAD = ['蝶城', '员工姓名', '员工编号', '项目', '岗位', '迟到<=5分钟次数', '迟到<=15分钟次数', '迟到15-120分钟',
    '早退<=5分钟次数', '早退<=15分钟次数', '早退15-120分钟', '合计迟到次数', '合计早退次数', '迟到早退合计', '扣款合计'];
  var SUMMARY_DETAIL_HEAD = ['蝶城'].concat(UNI_HEAD);

  function buildSummary(cities) {
    var statRows = [], detailRows = [], noAtt = [];
    cities.forEach(function (c) {
      c.statRows.forEach(function (r) { statRows.push([c.city].concat(r)); });
      c.detailRows.forEach(function (r) { detailRows.push([c.city].concat(r)); });
      c.noAttendance.forEach(function (r) { noAtt.push(r); });
    });
    return { statRows: statRows, detailRows: detailRows, noAttendance: noAtt };
  }

  /* 统计行的 10 个数值列（迟到<=5/<=15/15-120、早退同、合计迟到/早退/总次数、扣款合计）。
     蝶城统计表是 14 列，汇总总表前面多一列「蝶城」是 15 列——统一从这里取数，
     否则会按旧下标把「迟到早退合计（次数）」当成「扣款合计（元）」。 */
  function statNums(r) {
    var d = r.length >= 15 ? r.slice(1) : r, out = [];
    for (var i = 4; i <= 13; i++) {
      var v = d[i];
      out.push(typeof v === 'number' ? v : (Number(v) || 0));
    }
    return out;
  }
  /* 扣款合计：取统计行最后一列金额，兼容带/不带蝶城列两种行 */
  function statDed(rows) {
    var s = 0;
    (rows || []).forEach(function (r) { s += statNums(r)[9]; });
    return round2(s);
  }

  /* ---------------- 导出工作簿 spec ---------------- */

  var SHR_HEAD = ['员工姓名', '员工编号', '岗位名称', '日期', '项目名称', '班次时间', '签到时间', '签退时间',
    '理论工时', '考勤结果', '异常考勤', '补卡工时', '计薪时长', '工时', '上班时间', '下班时间',
    '迟到(分钟)', '早退(分钟)', '迟到扣款', '早退扣款', '扣款合计'];
  var RM_HEAD = ['账号', '姓名', '岗位', '日期', '项目', '蝶城', '班次', '签到', '签退', '实际工时',
    '考勤结果', '异常考勤', '计薪时长', '补卡工时', '旷工', '迟到(分钟)', '早退(分钟)',
    '工作日加班', '休息日加班', '法定节假日', '上班时间', '下班时间', '迟到扣款', '早退扣款', '扣款合计'];
  var MATRIX_HEAD4 = ['项目', '入职', '离职', '姓名'];

  function attStyleShr(r, c, v) {
    var base = c === 8 ? 11 : (c === 12 || c === 13) ? 3 : 2;
    if ((c === 6 || c === 7) && String(v).indexOf('缺卡') >= 0) return 8;
    if (c === 9 && String(v).indexOf('旷工') >= 0) return 8;
    if ((c === 16 || c === 17) && typeof v === 'number' && v !== 0) return 8;
    if (c >= 18 && c <= 20 && typeof v === 'number' && v > 0) return 8;
    return base;
  }
  function attStyleRm(r, c, v) {
    var base = (c === 9 || c === 12) ? 3 : 2;
    if ((c === 7 || c === 8) && String(v).indexOf('缺卡') >= 0) return 8;
    if (c === 10 && String(v).indexOf('旷工') >= 0) return 8;
    if ((c === 15 || c === 16) && typeof v === 'number' && v !== 0) return 8;
    if (c >= 22 && c <= 24 && typeof v === 'number' && v > 0) return 8;
    return base;
  }
  function attStyleUni(r, c, v) {
    var base = (c === 8) ? 11 : (c === 9 || c === 12 || c === 14) ? 3 : 2;
    if ((c === 6 || c === 7) && String(v).indexOf('缺卡') >= 0) return 8;
    if (c === 10 && String(v).indexOf('旷工') >= 0) return 8;
    if ((c === 16 || c === 17) && typeof v === 'number' && v !== 0) return 8;
    if (c >= 18 && c <= 20 && typeof v === 'number' && v > 0) return 8;
    return base;
  }
  function statStyle(r, c, v, isTotal) {
    if (isTotal) return c === 13 ? 14 : 9;
    if (c >= 4 && c <= 13 && typeof v === 'number' && v > 0) return 8;
    return c === 13 ? 3 : 2;   // 扣款合计列（第 14 列，下标 13）保留两位小数
  }

  /* 计薪工时高亮规则（第 0~3 列是 项目/入职/离职/姓名，从第 4 列起是日期工时格）：
     - 0.1~0.9 → 淡黄FFF2CC底+加粗+两位小数（样式19）
     - 0 → 留空 + 亮黄FFFFCC底（样式21）；留空由 matrixSpec 的 blankZero 把值转成''
     - summaryCol>=0 时该列为"出勤天数汇总" → 加粗两位小数（样式22） */
  function matrixStyleFn(rows, highlight, blankZero, summaryCol) {
    if (!highlight && !blankZero && summaryCol == null) return null;
    return function (r, cc) {
      if (cc < 4) return 2;
      if (summaryCol != null && cc === summaryCol) return 22;
      var v = rows[r] && rows[r][cc];
      if (blankZero && v === 0) return 21;
      if (highlight && typeof v === 'number' && v >= 0.1 && v <= 0.9) return 19;
      return 2;
    };
  }
  /* opt: {highlight, tall, blankZero, summary}
     highlight=0.1-0.9淡黄加粗 | blankZero=0值留空填亮黄 | summary=末尾加"出勤天数汇总"列(各日期求和保留2位)
     tall=数据行高0.83cm/表头1.06cm/表头顶端居中/开筛选 */
  function matrixSpec(name, days, rows, opt) {
    opt = opt || {};
    var header = MATRIX_HEAD4.concat(days);
    var outRows = rows;
    var summaryCol = null;
    if (opt.summary) {
      summaryCol = 4 + days.length;
      header = header.concat(['出勤天数汇总']);
      outRows = rows.map(function (row) {
        var sum = 0;
        for (var i = 4; i < 4 + days.length; i++) {
          var v = row[i];
          if (typeof v === 'number') sum += v;
        }
        return row.concat([round2(sum)]);
      });
    }
    if (opt.blankZero) {
      outRows = outRows.map(function (row) {
        return row.map(function (v, i) {
          // 汇总列不参与留空；日期工时格里的 0 转成空串
          return (i >= 4 && (summaryCol == null || i < summaryCol) && v === 0) ? '' : v;
        });
      });
    }
    var spec = {
      name: name, header: header, rows: outRows,
      headerStyle: opt.tall ? 20 : 7,
      dataStyle: 2, freeze: { x: 4, y: 1 },
      cols: [18, 12, 12, 10].concat(days.map(function () { return 5; }))
        .concat(opt.summary ? [14] : [])
    };
    spec.styleFn = matrixStyleFn(rows, opt.highlight, opt.blankZero, summaryCol);
    if (opt.tall) { spec.rowHeight = 23.5; spec.headerHeight = 30; spec.autoFilter = true; }
    return spec;
  }

  /* 蝶城结果 → xlsx spec 数组 */
  /* 迟到早退统计表 spec（外包文件与机师岗文件共用） */
  function statSpec(c, statData, rawCount) {
    var statHead = SUMMARY_STAT_HEAD.slice(1);              // 去掉蝶城列
    if (statData.length) {
      var t = { l5: 0, l15: 0, l120: 0, e5: 0, e15: 0, e120: 0, lt: 0, et: 0, tt: 0, ded: 0 };
      statData.forEach(function (r) {
        var d = statNums(r);
        t.l5 += d[0]; t.l15 += d[1]; t.l120 += d[2]; t.e5 += d[3]; t.e15 += d[4]; t.e120 += d[5];
        t.lt += d[6]; t.et += d[7]; t.tt += d[8]; t.ded += d[9];
      });
      statData.push(['', '', '', '合计', t.l5, t.l15, t.l120, t.e5, t.e15, t.e120, t.lt, t.et, t.tt, round2(t.ded)]);
    }
    return {
      name: '迟到早退统计', header: statHead, rows: statData, headerStyle: 7,
      headerStyleFn: function (cc) { return (cc === 5 || cc === 8) ? 10 : 7; },
      styleFn: function (r, cc) {
        if (r === statData.length - 1 && statData.length > rawCount) return cc === 13 ? 14 : 9;   // 合计行
        return statStyle(r, cc, statData[r][cc], false);
      },
      freeze: { x: 0, y: 1 }, autoFilter: true, rowHeight: 24.75,
      cols: [10, 12, 14, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 10]
    };
  }

  /* 一个蝶城拆成两份文件：外包（SHR+差异+统计+矩阵，矩阵为两类人员汇总）、机师岗（RM员工考勤+机师岗迟到早退明细/统计） */
  function cityFileSpecs(c) {
    var files = [];
    if (c.shrRows.length) {
      var specs = [];
      specs.push({
        name: 'SHR员工考勤', header: SHR_HEAD, rows: c.shrRows, headerStyle: 7,
        styleFn: function (r, cc) { return attStyleShr(r, cc, c.shrRows[r][cc]); },
        freeze: { x: 4, y: 1 }, autoFilter: true, rowHeight: 24.75,
        cols: [10, 12, 10, 11, 16, 13, 8, 8, 8, 9, 13, 8, 8, 8, 8, 8, 9, 9, 8, 8, 8]
      });
      specs.push({
        name: '考勤差异明细', header: UNI_HEAD, rows: c.diffRows, headerStyle: 7,
        styleFn: function (r, cc) { return attStyleUni(r, cc, c.diffRows[r][cc]); },
        freeze: { x: 4, y: 1 }, autoFilter: true, rowHeight: 24.75,
        cols: [10, 12, 10, 11, 16, 13, 8, 8, 8, 8, 9, 13, 8, 8, 8, 7, 9, 9, 8, 8, 8, 9, 9, 9]
      });
      specs.push({
        name: '迟到早退明细', header: UNI_HEAD, rows: c.detailRows, headerStyle: 7,
        styleFn: function (r, cc) { return attStyleUni(r, cc, c.detailRows[r][cc]); },
        freeze: { x: 4, y: 1 }, autoFilter: true, rowHeight: 24.75,
        cols: [10, 12, 10, 11, 16, 13, 8, 8, 8, 8, 9, 13, 8, 8, 8, 7, 9, 9, 8, 8, 8, 9, 9, 9]
      });
      specs.push(statSpec(c, c.statRows.slice(), c.statRows.length));
      specs.push(matrixSpec('计薪工时', c.matrixDays, c.payMatrix, { highlight: true, blankZero: true, summary: true, tall: true }));
      specs.push(matrixSpec('理论工时', c.matrixDays, c.theoryMatrix));
      specs.push(matrixSpec('打卡工时', c.matrixDays, [], { tall: true }));
      specs.push(matrixSpec('工时比对', c.matrixDays, []));
      files.push({ name: c.fileNameShr, sheets: specs });
    }
    if (c.rmRows.length) {
      var rmSpecs = [{
        name: 'RM员工考勤', header: RM_HEAD, rows: c.rmRows, headerStyle: 7,
        styleFn: function (r, cc) { return attStyleRm(r, cc, c.rmRows[r][cc]); },
        freeze: { x: 4, y: 1 }, autoFilter: true, rowHeight: 24.75,
        cols: [10, 9, 9, 11, 16, 12, 13, 8, 8, 8, 9, 13, 8, 8, 7, 9, 9, 9, 9, 9, 8, 8, 8, 8, 8]
      }];
      if (c.detailRowsRm.length) {
        rmSpecs.push({
          name: '迟到早退明细', header: UNI_HEAD, rows: c.detailRowsRm, headerStyle: 7,
          styleFn: function (r, cc) { return attStyleUni(r, cc, c.detailRowsRm[r][cc]); },
          freeze: { x: 4, y: 1 }, autoFilter: true, rowHeight: 24.75,
          cols: [10, 12, 10, 11, 16, 13, 8, 8, 8, 8, 9, 13, 8, 8, 8, 7, 9, 9, 8, 8, 8, 9, 9, 9]
        });
      }
      rmSpecs.push(statSpec(c, c.statRowsRm.slice(), c.statRowsRm.length));
      /* 机师岗文件：补四张工时矩阵（仅 RM 人员），取数规则与外包一致 */
      rmSpecs.push(matrixSpec('计薪工时', c.matrixDays, c.rmPayMatrix, { highlight: true, blankZero: true, summary: true, tall: true }));
      rmSpecs.push(matrixSpec('理论工时', c.matrixDays, c.rmTheoryMatrix));
      rmSpecs.push(matrixSpec('打卡工时', c.matrixDays, c.rmPunchMatrix, { tall: true }));
      rmSpecs.push(matrixSpec('工时比对', c.matrixDays, []));
      files.push({ name: c.fileNameRm, sheets: rmSpecs });
    }
    return files;
  }

  function summarySpecs(s, cities) {
    var specs = [];
    var statData = s.statRows.slice();
    if (statData.length) {
      var t = { l5: 0, l15: 0, l120: 0, e5: 0, e15: 0, e120: 0, lt: 0, et: 0, tt: 0, ded: 0 };
      statData.forEach(function (r) {
        var d = statNums(r);   // 汇总行前面多一列蝶城，必须走助手换算
        t.l5 += d[0]; t.l15 += d[1]; t.l120 += d[2]; t.e5 += d[3]; t.e15 += d[4]; t.e120 += d[5];
        t.lt += d[6]; t.et += d[7]; t.tt += d[8]; t.ded += d[9];
      });
      statData.push(['', '', '', '合计', '', t.l5, t.l15, t.l120, t.e5, t.e15, t.e120, t.lt, t.et, t.tt, round2(t.ded)]);
    }
    specs.push({
      name: '迟到早退统计汇总', header: SUMMARY_STAT_HEAD, rows: statData, headerStyle: 7,
      styleFn: function (r, cc) {
        // 汇总表 15 列：蝶城(0) 姓名(1) 编号(2) 项目(3) 岗位(4) 数值(5..14)，扣款合计在 14
        if (r === statData.length - 1 && statData.length > s.statRows.length) return cc === 14 ? 14 : 9;   // 合计行
        if (cc >= 5 && cc <= 14 && typeof statData[r][cc] === 'number' && statData[r][cc] > 0) return 8;
        return cc === 14 ? 3 : 2;
      },
      freeze: { x: 0, y: 1 }, autoFilter: true, rowHeight: 24.75,
      cols: [12, 10, 12, 14, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 10]
    });
    specs.push({
      name: '迟到早退明细汇总', header: SUMMARY_DETAIL_HEAD, rows: s.detailRows, headerStyle: 7,
      styleFn: function (r, cc) { return attStyleUni(r, cc - 1, s.detailRows[r][cc]); },
      freeze: { x: 5, y: 1 }, autoFilter: true, rowHeight: 24.75,
      cols: [12].concat([10, 12, 10, 11, 16, 13, 8, 8, 8, 8, 9, 13, 8, 8, 8, 7, 9, 9, 8, 8, 8, 9, 9, 9])
    });
    specs.push({
      name: '在职无出勤', header: ['蝶城', '姓名', '员工编号', '项目', '入职日期', '离职日期'],
      rows: s.noAttendance, headerStyle: 7, dataStyle: 2, freeze: { x: 0, y: 1 }, autoFilter: true,
      cols: [12, 10, 12, 16, 11, 11]
    });

    /* 各蝶城差异数据汇总：差异总数 / 迟到总数 / 早退总数 / 迟到早退扣款总数 */
    if (cities && cities.length) {
      var cityRows = [], T = { diff: 0, late: 0, early: 0, ded: 0 };
      cities.forEach(function (c) {
        var late = 0, early = 0, ded = 0;
        (c.statRows || []).forEach(function (r) {
          var d = statNums(r);
          late += d[6]; early += d[7]; ded += d[9];
        });
        var diff = (c.diffRows || []).length;
        cityRows.push([c.city, diff, late, early, round2(ded)]);
        T.diff += diff; T.late += late; T.early += early; T.ded += ded;
      });
      cityRows.push(['合计', T.diff, T.late, T.early, round2(T.ded)]);
      specs.push({
        name: '各蝶城差异汇总',
        header: ['蝶城', '差异总数', '迟到总数', '早退总数', '迟到早退扣款总数'],
        rows: cityRows, headerStyle: 7, dataStyle: 2,
        styleFn: function (r, cc) {
          if (r === cityRows.length - 1) return cc === 4 ? 14 : 9;   // 合计行
          return cc >= 1 && cc <= 4 && typeof cityRows[r][cc] === 'number' && cityRows[r][cc] > 0 ? 3 : 2;
        },
        freeze: { x: 0, y: 1 }, autoFilter: true, rowHeight: 24.75,
        cols: [18, 12, 12, 12, 20]
      });
    }
    return specs;
  }

  return {
    DEFAULT_CONFIG: DEFAULT_CONFIG,
    norm: norm, parseHM: parseHM, parseDateKey: parseDateKey, parseShift: parseShift,
    detectCity: detectCity, calcLate: calcLate, calcEarly: calcEarly, deduction: deduction,
    ROSTER_FIELDS: ROSTER_FIELDS,
    guessRosterMap: guessRosterMap,
    parseRoster: parseRoster,
    parseAttFile: parseAttFile,
    processAll: processAll,
    matrixStyleFn: matrixStyleFn,
    isDiff: isDiff,
    splitTokens: splitTokens,
    makeExcludeFilter: makeExcludeFilter,
    inferFileCityFor: function (parsed, rosterRows) {
      return inferFileCity(parsed, buildRosterIndex(rosterRows || []));
    },
    cityFileSpecs: cityFileSpecs,
    summarySpecs: summarySpecs,
    statDed: statDed,
    matchRoster: matchRoster
  };
});
