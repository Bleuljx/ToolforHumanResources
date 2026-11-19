/* ============================================================
 * idle-core.js — 出勤异常核对核心（超7天未出勤）
 * 三类输入：
 *   ① 考勤统计结果导出：SHRid/姓名/蝶城/项目/应服务工时/计薪天数/月份…
 *   ② 人员花名册 / 员工信息维护：员工编码/姓名/组织/项目/职位/入司日期/离职日期…
 *   ③ 上期反馈结果（本工具上期导出的表）：异常原因/请假起止/请假天数/是否已报离/备注
 * 口径（与参考工具一致）：
 *   无出勤天数 = 本期应出勤天数（日历天数）− 计薪天数，按在职区间封顶
 *   门槛：无出勤 ≥ N 天（默认 7）或 实出勤 < N 天
 *   上期反馈：已报离/离职 → 剔除备查；请假 → 扣区间与本核算期重叠天数
 *   本期离职人员 → 假异常，剔除备查
 *   员工表在职但考勤表无工号 → 估算（在职日历天数），单独子表
 * ============================================================ */
window.IdleCore = (function () {
  'use strict';

  /* ---------------- 基础工具 ---------------- */
  function normHeader(s) {
    return String(s == null ? '' : s)
      .replace(/[\s\u3000]/g, '')
      .replace(/[（(][^)）]*[)）]/g, '')
      .replace(/[：:*＊]/g, '');
  }
  function num(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    if (v == null) return 0;
    var s = String(v).replace(/[,，\s\u3000]/g, '').replace(/[^\d.\-]/g, '');
    var n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }
  function toDate(v) {
    if (v instanceof Date && !isNaN(v.getTime())) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
    if (typeof v === 'number' && v > 20000 && v < 80000) {
      var d = window.XLSX && window.XLSX.SSF ? window.XLSX.SSF.parse_date_code(v) : null;
      if (d) return new Date(d.y, d.m - 1, d.d);
    }
    if (typeof v === 'string') {
      var s = v.trim().replace(/[年月]/g, '-').replace(/日/g, '').replace(/\//g, '-');
      var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
      if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    }
    return null;
  }
  function r4(n) { return Math.round(n * 10000) / 10000; }
  function fmtDay(n) {
    if (n == null || !isFinite(n)) return '—';
    var v = Math.round(n * 100) / 100;
    return (Math.abs(v - Math.round(v)) < 1e-9) ? String(Math.round(v)) : String(v);
  }
  function fmtDate(d) {
    if (!d) return '';
    var m = d.getMonth() + 1, dd = d.getDate();
    return d.getFullYear() + '-' + (m < 10 ? '0' : '') + m + '-' + (dd < 10 ? '0' : '') + dd;
  }
  function daysBetween(a, b) {
    if (!a || !b) return 0;
    var da = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
    var db = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
    var n = Math.round((db - da) / 86400000) + 1;
    return n > 0 ? n : 0;
  }
  function monthBounds(ym) {
    var m = String(ym || '').match(/(\d{4})\D+(\d{1,2})/);
    if (!m) {
      var now = new Date();
      return { start: new Date(now.getFullYear(), now.getMonth(), 1), end: new Date(now.getFullYear(), now.getMonth() + 1, 0), label: now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') };
    }
    var y = +m[1], mo = +m[2];
    return { start: new Date(y, mo - 1, 1), end: new Date(y, mo, 0), label: y + '-' + String(mo < 10 ? '0' : '') + mo };
  }

  /* ---------------- 列识别 ---------------- */
  var EMPJOIN = ['入职日期', '入司日期', '入司时间'];
  var EMPORG = ['行政组织', '所属部门', '部门'];

  var FIELD_EMP = [
    { k: 'id', label: '员工编码', req: true, aliases: ['员工编码', '工号', 'SHRid'] },
    { k: 'name', label: '姓名', aliases: ['员工姓名', '人员姓名', '人员', '姓名'] },
    { k: 'dept', label: '组织 / 调度中心', aliases: ['最新所属组织名称', '所属行政组织名称', '行政组织名称', '所属部门', '组织名称', '部门', '行政组织'] },
    { k: 'proj', label: '项目', aliases: ['员工项目', '战图项目名称', '所属项目', '项目'] },
    { k: 'pos', label: '职位', aliases: ['最新职位名称', '职位', '岗位'] },
    { k: 'join', label: '入职日期', aliases: ['入职日期', '本次入职日期', '入司日期', '入司时间'] },
    { k: 'end', label: '离职日期', aliases: ['离职日期'] },
    { k: 'status', label: '用工关系状态', aliases: ['用工关系状态', '用工关系', '状态'] },
    { k: 'sup', label: '供应商', aliases: ['供应商名称', '所属供应商', '供应商'] },
    { k: 'city', label: '蝶城（若有）', aliases: ['蝶城', '所属蝶城'] }
  ];
  var FIELD_ATT = [
    { k: 'id', label: 'SHRid / 工号', req: true, aliases: ['SHRid', 'SHRID', '工号', '员工编码'] },
    { k: 'name', label: '姓名', aliases: ['姓名', '员工姓名'] },
    { k: 'city', label: '蝶城', aliases: ['蝶城', '所属蝶城'] },
    { k: 'proj', label: '项目', aliases: ['项目', '员工项目', '所属项目'] },
    { k: 'should', label: '应服务工时', req: true, aliases: ['应服务工时', '应出勤工时'] },
    { k: 'pay', label: '计薪天数', req: true, aliases: ['计薪天数', '出勤天数'] },
    { k: 'actual', label: '实际服务工时', aliases: ['实际服务工时', '实出勤工时'] },
    { k: 'month', label: '月份', aliases: ['月份', '考勤月份'] },
    { k: 'nature', label: '员工性质', aliases: ['员工性质'] },
    { k: 'sup', label: '供应商', aliases: ['供应商'] }
  ];
  var FIELD_PREV = [
    { k: 'id', label: '员工编码', aliases: ['员工编码', 'SHRid', 'SHRID', '工号'] },
    { k: 'name', label: '员工姓名', aliases: ['员工姓名', '姓名'] },
    { k: 'city', label: '蝶城', aliases: ['蝶城', '城市'] },
    { k: 'reason', label: '异常原因', aliases: ['异常原因', '原因', '反馈情况', '异常类型'] },
    { k: 'ls', label: '请假开始时间', aliases: ['请假开始时间', '请假开始', '开始时间'] },
    { k: 'le', label: '请假结束时间', aliases: ['请假结束时间', '请假结束', '结束时间'] },
    { k: 'ld', label: '请假天数', aliases: ['请假天数', '请假合计', '合计天数'] },
    { k: 'rep', label: '是否已报离', aliases: ['是否已报离', '已报离', '报离'] },
    { k: 'note', label: '备注', aliases: ['备注', '说明'] }
  ];

  var SOFT = {
    id: ['员工编码', '人员编号', '工号', 'SHRid', '编码', '编号'],
    name: ['员工姓名', '人员姓名', '姓名', '人员', '名字'],
    dept: ['所属行政组织名称', '最新所属组织名称', '所属中心', '组织名称', '调度中心', '所属部门', '部门', '中心'],
    proj: ['战图项目名称', '员工项目', '楼盘', '项目'],
    pos: ['最新职位名称', '职位名称', '岗级', '岗位', '职位'],
    join: ['本次入职日期', '入职日期', '到岗日', '入司日期', '入司时间'],
    end: ['离职日期', '离职时间', '离开日期'],
    status: ['用工关系状态', '用工关系', '在职状态', '状态'],
    sup: ['供应商名称', '所属供应商', '供应商'],
    city: ['蝶城', '所属蝶城', '城市'],
    should: ['应服务工时', '应出勤工时'],
    actual: ['实际服务工时', '实出勤工时'],
    pay: ['计薪天数', '出勤天数'],
    month: ['考勤月份', '月份'],
    reason: ['异常原因', '原因', '反馈情况'],
    ls: ['请假开始时间', '请假开始'],
    le: ['请假结束时间', '请假结束'],
    ld: ['请假天数', '请假合计'],
    rep: ['是否已报离', '已报离', '报离'],
    note: ['备注', '说明']
  };
  function guessCol(headers, key) {
    var kws = SOFT[key] || [];
    var best = -1, bs = -1e9;
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i]; if (!h) continue;
      for (var j = 0; j < kws.length; j++) {
        var kw = kws[j], sc = 0;
        if (h === kw) sc = 100;
        else if (h.slice(-kw.length) === kw) sc = 90;
        else if (h.indexOf(kw) >= 0) sc = 80;
        if (!sc) continue;
        sc -= h.length * 0.5;
        if (sc > bs) { bs = sc; best = i; }
      }
    }
    return best;
  }
  function colIdx(headers, aliases) {
    for (var a = 0; a < aliases.length; a++) {
      var al = normHeader(aliases[a]);
      for (var i = 0; i < headers.length; i++) if (headers[i] === al) return i;
    }
    for (var a2 = 0; a2 < aliases.length; a2++) {
      var al2 = normHeader(aliases[a2]);
      for (var j = 0; j < headers.length; j++) if (headers[j].indexOf(al2) >= 0) return j;
    }
    return -1;
  }
  function buildCols(fields, headers) {
    var c = {};
    fields.forEach(function (f) {
      var v = colIdx(headers, f.aliases);
      c[f.k] = v >= 0 ? v : guessCol(headers, f.k);
    });
    return c;
  }
  function scoreRow(headers, fields, reqKeys) {
    var c = buildCols(fields, headers);
    var hit = 0;
    fields.forEach(function (f) { if (c[f.k] >= 0) hit++; });
    var missReq = (reqKeys || []).filter(function (k) { return c[k] < 0; }).length;
    return { score: hit - missReq * 5, cols: c, missReq: missReq };
  }

  /* ---------------- 文件解析（wb 为 SheetJS workbook，kind: att/emp/prev） ---------------- */
  function parseFile(wb, kind, fileName) {
    var FIELDS = kind === 'att' ? FIELD_ATT : kind === 'emp' ? FIELD_EMP : FIELD_PREV;
    var REQ = kind === 'att' ? ['id', 'pay', 'should'] : kind === 'emp' ? ['id'] : ['name'];
    var best = null;
    (wb.SheetNames || []).forEach(function (sn) {
      var aoa;
      try {
        aoa = window.XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null, blankrows: false });
      } catch (e) { return; }
      var limit = Math.min(aoa.length, 12);
      for (var i = 0; i < limit; i++) {
        var row = aoa[i] || [];
        if (row.filter(function (x) { return x != null && String(x).trim() !== ''; }).length < 3) continue;
        var headers = row.map(normHeader);
        var sc = scoreRow(headers, FIELDS, REQ);
        if (sc.missReq > 0) continue;
        if (!best || sc.score > best.score) best = { aoa: aoa, hIdx: i, headers: headers, cols: sc.cols, sheet: sn, score: sc.score };
      }
    });
    if (!best) return { error: '没有找到匹配的表头（需要列：' + FIELDS.filter(function (f) { return f.req; }).map(function (f) { return f.label; }).join('、') + '）。请确认这份表是不是放在了正确的上传区' };
    var out = { sheet: best.sheet, headers: best.headers, cols: best.cols };
    if (kind === 'att') { parseAttRows(best, out); if (out.error) return out; }
    else if (kind === 'emp') { parseEmpRows(best, out); if (out.error) return out; }
    else parsePrevRows(best, out);
    out.fileName = fileName || '';
    return out;
  }

  function parseAttRows(best, out) {
    var c = best.cols, aoa = best.aoa;
    var rows = [];
    for (var i = best.hIdx + 1; i < aoa.length; i++) {
      var row = aoa[i]; if (!row) continue;
      var id = String(row[c.id] == null ? '' : row[c.id]).trim();
      var name = c.name >= 0 ? String(row[c.name] == null ? '' : row[c.name]).trim() : '';
      if (!id && !name) continue;
      rows.push({
        city: c.city >= 0 ? String(row[c.city] == null ? '' : row[c.city]).trim() : '',
        project: c.proj >= 0 ? String(row[c.proj] == null ? '' : row[c.proj]).trim() : '',
        month: c.month >= 0 ? String(row[c.month] == null ? '' : row[c.month]).trim() : '',
        name: name,
        id: id.toUpperCase(),
        should: num(row[c.should]),
        actual: c.actual >= 0 ? num(row[c.actual]) : 0,
        pay: num(row[c.pay]),
        nature: c.nature >= 0 ? String(row[c.nature] == null ? '' : row[c.nature]).trim() : '',
        sup: c.sup >= 0 ? String(row[c.sup] == null ? '' : row[c.sup]).trim() : ''
      });
    }
    out.rows = rows;
  }

  function parseEmpRows(best, out) {
    var c = best.cols, aoa = best.aoa;
    var warns = [];
    if (c.join < 0) warns.push('没找到「入司日期 / 入职日期」列，本期入职、离职人员不会按在职天数封顶');
    if (c.dept < 0 && c.proj < 0) warns.push('没找到「所属行政组织名称 / 员工项目」列，单蝶城导入时无法按范围匹配');
    if (c.pos < 0) warns.push('没找到「职位」列，排除名单里填的职位（如机师岗）不生效');
    var map = new Map();
    for (var i = best.hIdx + 1; i < aoa.length; i++) {
      var row = aoa[i]; if (!row) continue;
      var id = String(row[c.id] == null ? '' : row[c.id]).trim().toUpperCase();
      if (!id || id === 'NULL') continue;
      var rec = {
        id: id,
        city: c.city >= 0 ? String(row[c.city] == null ? '' : row[c.city]).trim() : '',
        name: c.name >= 0 ? String(row[c.name] == null ? '' : row[c.name]).trim() : '',
        dept: c.dept >= 0 ? String(row[c.dept] == null ? '' : row[c.dept]).trim() : '',
        project: c.proj >= 0 ? String(row[c.proj] == null ? '' : row[c.proj]).trim() : '',
        sup: c.sup >= 0 ? String(row[c.sup] == null ? '' : row[c.sup]).trim() : '',
        join: c.join >= 0 ? toDate(row[c.join]) : null,
        end: c.end >= 0 ? toDate(row[c.end]) : null,
        pos: c.pos >= 0 ? String(row[c.pos] == null ? '' : row[c.pos]).trim() : '',
        status: c.status >= 0 ? String(row[c.status] == null ? '' : row[c.status]).trim() : ''
      };
      var old = map.get(id);
      if (!old) { map.set(id, rec); continue; }
      // 同一工号多行（重新入职等情况）：优先取在职（无离职日期）的，其次取入职日期最新的
      if (!old.end && rec.end) continue;
      if (old.end && !rec.end) { map.set(id, rec); continue; }
      if (old.join && rec.join && rec.join > old.join) map.set(id, rec);
    }
    out.rows = Array.from(map.values());
    out.warnings = warns;
  }

  function parsePrevRows(best, out) {
    var c = best.cols, aoa = best.aoa;
    var ids = new Set(), byId = {}, byName = {}, n = 0;
    for (var i = best.hIdx + 1; i < aoa.length; i++) {
      var row = aoa[i]; if (!row) continue;
      var rec = {
        id: c.id >= 0 ? String(row[c.id] == null ? '' : row[c.id]).trim().toUpperCase() : '',
        name: c.name >= 0 ? String(row[c.name] == null ? '' : row[c.name]).trim() : '',
        city: c.city >= 0 ? String(row[c.city] == null ? '' : row[c.city]).trim() : '',
        reason: c.reason >= 0 ? String(row[c.reason] == null ? '' : row[c.reason]).trim() : '',
        leaveFrom: c.ls >= 0 ? toDate(row[c.ls]) : null,
        leaveTo: c.le >= 0 ? toDate(row[c.le]) : null,
        leaveDays: c.ld >= 0 ? num(row[c.ld]) : 0,
        reported: c.rep >= 0 ? String(row[c.rep] == null ? '' : row[c.rep]).trim() : '',
        note: c.note >= 0 ? String(row[c.note] == null ? '' : row[c.note]).trim() : ''
      };
      if (!rec.id && !rec.name) continue;
      n++;
      if (rec.id) { ids.add(rec.id); if (!byId[rec.id]) byId[rec.id] = rec; }
      if (rec.name) {
        var k1 = rec.name + '|' + rec.city;
        if (!byName[k1]) byName[k1] = rec;
        if (!byName[rec.name]) byName[rec.name] = rec;
      }
    }
    out.ids = Array.from(ids);
    out.byId = byId;
    out.byName = byName;
    out.count = n;
  }

  /* ---------------- 核心计算 ---------------- */
  function compute(cfg) {
    var att = cfg.att || [], empList = cfg.emp || [];
    var prevIds = new Set(cfg.prevIds || []);
    var prevById = cfg.prevById || {}, prevByName = cfg.prevByName || {};
    var thr = num(cfg.thr) || 7;
    var mode = cfg.mode || 'gap';
    var scopeOn = cfg.scopeOn !== false;
    var leaveOutOn = cfg.leaveOut !== false;
    var showNoRec = cfg.showNoRec !== false;
    var exclKeys = String(cfg.excl || '').split(/[,，、;；]+/).map(function (s) { return s.trim(); }).filter(Boolean);

    var res = {
      anomalies: [], noRecord: [], leftOut: [], reportedOut: [],
      stats: { dedup: 0, filteredOut: 0, outScope: 0, leftOut: 0, exclN: 0, reportedOut: 0, deducted: 0, cleared: 0 },
      warnings: []
    };
    if (!att.length) { res.ran = false; return res; }

    // 排除名单：匹配 工号（精确，忽略大小写）/ 姓名（精确）/ 职位（包含）
    function isExcluded(emp, name, id) {
      if (!exclKeys.length) return false;
      var sid = String(id == null ? '' : id).trim().toUpperCase();
      var sname = String(name == null ? '' : name).trim();
      var spos = emp ? String(emp.pos || '').trim() : '';
      for (var i = 0; i < exclKeys.length; i++) {
        var k = exclKeys[i]; if (!k) continue;
        if (sid && sid === k.toUpperCase()) return true;
        if (sname && sname === k) return true;
        if (spos && spos.indexOf(k) >= 0) return true;
      }
      return false;
    }

    // 期次：考勤表「月份」列众数
    var mCount = {};
    att.forEach(function (r) { if (r.month) mCount[r.month] = (mCount[r.month] || 0) + 1; });
    var monthKey = Object.keys(mCount).sort(function (a, b) { return mCount[b] - mCount[a]; })[0] || '';
    var mb = monthBounds(monthKey);
    res.month = mb.label;
    var monthFirst = new Date(mb.start.getFullYear(), mb.start.getMonth(), mb.start.getDate());
    var monthLast = new Date(mb.end.getFullYear(), mb.end.getMonth(), mb.end.getDate());

    // 按工号合并
    var agg = new Map();
    att.forEach(function (r) {
      var key = r.id || ('N:' + r.name + '|' + r.project);
      var o = agg.get(key);
      if (!o) {
        o = { id: r.id, name: r.name, citySet: [], projects: [], should: 0, actual: 0, pay: 0, nature: r.nature, sups: [] };
        agg.set(key, o);
      }
      if (r.city && o.citySet.indexOf(r.city) < 0) o.citySet.push(r.city);
      if (r.project && o.projects.indexOf(r.project) < 0) o.projects.push(r.project);
      if (r.nature && !o.nature) o.nature = r.nature;
      o.should += r.should; o.actual += r.actual; o.pay += r.pay;
      if (!o.name && r.name) o.name = r.name;
    });

    var empMap = new Map();
    empList.forEach(function (e) { if (!empMap.has(e.id)) empMap.set(e.id, e); });

    // 这份数据实际统计到哪天：计薪天数的最高值 = 数据覆盖的日历天数
    var maxPay = 0;
    agg.forEach(function (o) { if (o.pay > maxPay) maxPay = o.pay; });
    res.maxPay = maxPay;
    var de = new Date(monthFirst.getTime());
    de.setDate(de.getDate() + Math.ceil(maxPay) - 1);
    if (de > monthLast) de = new Date(monthLast.getTime());
    if (de < monthFirst) de = new Date(monthFirst.getTime());
    res.dataEnd = de;

    // 核算周期：默认 起始=月初、结束=数据截止日
    var periodStart = cfg.periodStart ? new Date(cfg.periodStart.getFullYear(), cfg.periodStart.getMonth(), cfg.periodStart.getDate()) : new Date(monthFirst.getTime());
    var periodEnd = cfg.periodEnd ? new Date(cfg.periodEnd.getFullYear(), cfg.periodEnd.getMonth(), cfg.periodEnd.getDate()) : new Date(de.getTime());
    if (periodStart < monthFirst) periodStart = new Date(monthFirst.getTime());
    if (periodStart > monthLast) periodStart = new Date(monthLast.getTime());
    if (periodEnd > monthLast) periodEnd = new Date(monthLast.getTime());
    if (periodEnd < periodStart) periodEnd = new Date(periodStart.getTime());
    res.periodStart = periodStart; res.periodEnd = periodEnd;
    var spanDays = daysBetween(periodStart, periodEnd);
    res.spanDays = spanDays;

    // 项目 → 蝶城 映射（供无考勤记录表用）
    var projCity = {};
    att.forEach(function (r) {
      if (!r.project || !r.city) return;
      projCity[r.project] = projCity[r.project] || {};
      projCity[r.project][r.city] = (projCity[r.project][r.city] || 0) + 1;
    });
    function cityOfProject(p) {
      var o = projCity[p]; if (!o) return '';
      var ks = Object.keys(o).sort(function (a, b) { return o[b] - o[a]; });
      return ks[0] || '';
    }

    var attCities = {}, attProjs = {};
    att.forEach(function (r) { if (r.city) attCities[r.city] = 1; if (r.project) attProjs[r.project] = 1; });
    var cityNames = Object.keys(attCities);
    var projNames = Object.keys(attProjs);
    res.scope = { cities: cityNames, projCount: projNames.length };

    // 名字归一：去掉 调度中心 / 蝶城外项目 / 蝶城 等后缀
    function coreName(s) {
      return String(s == null ? '' : s)
        .replace(/[\s\u3000]/g, '')
        .replace(/(蝶城外项目|调度中心|蝶城|蝶外|分公司|项目部|项目)/g, '');
    }
    var cityCores = cityNames.map(coreName);
    function cityFor(e) {
      if (e.city) return e.city;
      if (e.project) {
        var c0 = cityOfProject(e.project);
        if (c0) return c0;
        for (var i = 0; i < projNames.length; i++) {
          if (projNames[i].indexOf(e.project) >= 0 || e.project.indexOf(projNames[i]) >= 0) {
            var c1 = cityOfProject(projNames[i]);
            if (c1) return c1;
          }
        }
      }
      var dc = coreName(e.dept);
      if (dc) {
        var bestC = '', bl = 0;
        for (var k = 0; k < cityCores.length; k++) {
          var cc = cityCores[k];
          if (!cc) continue;
          if (dc.indexOf(cc) >= 0 || cc.indexOf(dc) >= 0) {
            if (cc.length > bl) { bl = cc.length; bestC = cityNames[k]; }
          }
        }
        if (bestC) return bestC;
      }
      return '';
    }
    function inScope(e) {
      if (!scopeOn) return true;
      if (!cityNames.length) return true;
      if (e.city) return !!attCities[e.city];
      if (e.project && attProjs[e.project]) return true;
      return !!cityFor(e);
    }

    res.peopleCount = agg.size;
    var cover = {};

    // 上期反馈查找：工号 → 姓名+蝶城 → 姓名
    function prevOf(id, name, cities) {
      var sid = String(id || '').trim().toUpperCase();
      if (sid && prevById[sid]) return prevById[sid];
      var nm = String(name || '').trim();
      if (!nm) return null;
      var cs = Array.isArray(cities) ? cities : (cities ? [cities] : []);
      for (var i = 0; i < cs.length; i++) {
        var hit = prevByName[nm + '|' + String(cs[i]).trim()];
        if (hit) return hit;
      }
      return prevByName[nm] || null;
    }
    // 上期反馈应用：已报离/离职 → 剔除；请假 → 扣重叠天数
    function applyPrev(rec, gap) {
      if (!rec) return { action: 'keep', deduct: 0, prev: null };
      var isLeft = /离职|辞职|离场/.test(rec.reason || '');
      if (rec.reported === '是' || isLeft) {
        return { action: 'drop', deduct: 0, prev: rec, why: rec.reported === '是' ? '上期已报离' : ('上期原因：' + rec.reason) };
      }
      var deduct = 0;
      if (rec.leaveFrom && rec.leaveTo) {
        var from = rec.leaveFrom > periodStart ? rec.leaveFrom : periodStart;
        var to = rec.leaveTo < periodEnd ? rec.leaveTo : periodEnd;
        if (to >= from) deduct = daysBetween(from, to);
        if (deduct > 0 && rec.leaveDays > 0 && rec.leaveDays < deduct) deduct = rec.leaveDays;
      } else if (rec.leaveDays > 0) {
        deduct = rec.leaveDays;
      }
      if (deduct > gap) deduct = gap;
      if (deduct < 0) deduct = 0;
      return { action: 'keep', deduct: Math.round(deduct * 100) / 100, prev: rec };
    }

    agg.forEach(function (o) {
      var ckey = o.citySet.length ? o.citySet[0] : '未匹配';
      if (!cover[ckey]) cover[ckey] = { imp: 0, norec: 0, ano: 0, gap: 0 };
      var cv = cover[ckey];
      cv.imp++;

      var emp = o.id ? empMap.get(o.id) : null;
      if (emp) {
        if (emp.join && emp.join > periodEnd) { res.stats.filteredOut++; return; }
        if (emp.end && emp.end < periodStart) { res.stats.filteredOut++; return; }
        if (scopeOn && emp.city && cityNames.length && !attCities[emp.city]) { res.stats.outScope++; return; }
      }
      var skipExcl = isExcluded(emp, o.name, o.id);
      if (emp) {
        // 本期离职：离职后无需出勤、考勤常被清档 → 算出来的是假异常，剔除备查
        if (leaveOutOn && emp.end && emp.end <= periodEnd) {
          if (skipExcl) return;
          var lf = (emp.join && emp.join > periodStart) ? emp.join : periodStart;
          var lShould = Math.min(spanDays, daysBetween(lf, emp.end));
          var lGap = Math.round((lShould - o.pay) * 100) / 100;
          if (lGap < 0) lGap = 0;
          if (lGap >= thr - 1e-6) {
            res.stats.leftOut++;
            res.leftOut.push({
              city: o.citySet.length ? o.citySet.join(' / ') : '未匹配',
              project: o.projects.length ? o.projects.join(' / ') : '未匹配',
              id: o.id || '', name: o.name || '',
              gap: lGap, pay: Math.round(o.pay * 100) / 100, should: lShould,
              joinTxt: emp.join ? fmtDate(emp.join) : '', endTxt: fmtDate(emp.end),
              why: '本期离职'
            });
          }
          return;
        }
      }
      // 应出勤 = 本期日历天数；本期入职/离职的只算实际在职天数
      var payDays = Math.round(o.pay * 100) / 100;
      var shouldDays = spanDays;
      var capDays = 0, capped = false;
      var pFrom = periodStart, pTo = periodEnd;
      if (emp) {
        if (emp.join && emp.join > periodStart) pFrom = emp.join;
        if (emp.end && emp.end < periodEnd) pTo = emp.end;
      }
      if (pFrom > periodStart || pTo < periodEnd) {
        capDays = daysBetween(pFrom, pTo);
        if (capDays > 0 && shouldDays > capDays) { shouldDays = capDays; capped = true; }
      }
      var gap = Math.round((shouldDays - o.pay) * 100) / 100;
      if (gap < 0) gap = 0;

      var hit;
      if (mode === 'low') hit = (payDays < thr - 1e-6) && (gap > 1e-6);
      else hit = (gap >= thr - 1e-6);
      if (!hit) return;
      if (skipExcl) { res.stats.exclN++; return; }

      var pv = prevOf(o.id, o.name, o.citySet);
      var ap = applyPrev(pv, gap);
      var cityTxt = o.citySet.length ? o.citySet.join(' / ') : '未匹配';
      var projTxt = o.projects.length ? o.projects.join(' / ') : '未匹配';
      if (ap.action === 'drop') {
        res.stats.reportedOut++;
        res.reportedOut.push({
          city: cityTxt, project: projTxt, id: o.id || '', name: o.name || '',
          gap: gap, pay: payDays, should: shouldDays, why: ap.why,
          joinTxt: (emp && emp.join) ? fmtDate(emp.join) : '',
          endTxt: (emp && emp.end) ? fmtDate(emp.end) : ''
        });
        return;
      }
      var deduct = (mode === 'low') ? 0 : (ap.deduct || 0);
      var netGap = Math.round((gap - deduct) * 100) / 100;
      if (netGap < 0) netGap = 0;
      if (netGap < thr - 1e-6) {
        if (deduct > 0) res.stats.cleared++;
        return;
      }
      if (deduct > 0) res.stats.deducted++;
      cv.ano++; cv.gap += netGap;
      res.anomalies.push({
        city: cityTxt, project: projTxt, id: o.id || '', name: o.name || '',
        gap: netGap, rawGap: gap, deducted: deduct, pay: payDays, should: shouldDays,
        prev: pv || null,
        reason: pv ? (pv.reason || '') : '',
        leaveFrom: pv && pv.leaveFrom ? fmtDate(pv.leaveFrom) : '',
        leaveTo: pv && pv.leaveTo ? fmtDate(pv.leaveTo) : '',
        leaveDays: pv ? (pv.leaveDays || 0) : 0,
        reported: pv ? (pv.reported || '') : '',
        note: pv ? (pv.note || '') : '',
        joinTxt: (emp && emp.join) ? fmtDate(emp.join) : '',
        endTxt: (emp && emp.end) ? fmtDate(emp.end) : ''
      });
    });

    // 无考勤记录：员工表本期在职，但考勤表没有该工号（天数按在职日历天数估算）
    var attIds = new Set();
    att.forEach(function (r) { if (r.id) attIds.add(r.id); });
    if (showNoRec && empList.length) {
      empList.forEach(function (e) {
        if (attIds.has(e.id)) return;
        if (e.join && e.join > periodEnd) return;
        if (e.end && e.end < periodStart) return;
        if (!inScope(e)) { res.stats.outScope++; return; }
        if (leaveOutOn && e.end && e.end <= periodEnd) {
          var lf2 = (e.join && e.join > periodStart) ? e.join : periodStart;
          var lGap2 = r4(daysBetween(lf2, e.end));
          if (lGap2 >= thr - 1e-6) {
            res.stats.leftOut++;
            res.leftOut.push({
              city: cityFor(e) || '未匹配',
              project: e.project || e.dept || '未匹配',
              id: e.id, name: e.name,
              gap: lGap2, pay: 0, should: lGap2, est: true,
              joinTxt: e.join ? fmtDate(e.join) : '', endTxt: fmtDate(e.end),
              why: '本期离职'
            });
          }
          return;
        }
        var from = (e.join && e.join > periodStart) ? e.join : periodStart;
        var to = (e.end && e.end < periodEnd) ? e.end : periodEnd;
        if (to < from) return;
        var est = r4(daysBetween(from, to));
        if (est < thr - 1e-6) return;
        if (isExcluded(e, e.name, e.id)) { res.stats.exclN++; return; }
        var city = cityFor(e) || '未匹配';
        var proj = e.project || e.dept || '未匹配';
        var pv2 = prevOf(e.id, e.name, [city]);
        var ap2 = applyPrev(pv2, est);
        if (ap2.action === 'drop') {
          res.stats.reportedOut++;
          res.reportedOut.push({
            city: city, project: proj, id: e.id || '', name: e.name || '',
            gap: est, pay: 0, should: est, why: ap2.why, est: true,
            joinTxt: e.join ? fmtDate(e.join) : '', endTxt: e.end ? fmtDate(e.end) : ''
          });
          return;
        }
        var deduct2 = ap2.deduct || 0;
        var net2 = Math.round((est - deduct2) * 100) / 100;
        if (net2 < 0) net2 = 0;
        if (net2 < thr - 1e-6) {
          if (deduct2 > 0) res.stats.cleared++;
          return;
        }
        if (deduct2 > 0) res.stats.deducted++;
        if (!cover[city]) cover[city] = { imp: 0, norec: 0, ano: 0, gap: 0 };
        cover[city].norec++;
        res.noRecord.push({
          city: city, project: proj, id: e.id, name: e.name,
          gap: net2, rawGap: est, deducted: deduct2, pay: 0, should: est, est: true,
          status: e.status || '',
          reason: pv2 ? (pv2.reason || '') : '',
          leaveFrom: pv2 && pv2.leaveFrom ? fmtDate(pv2.leaveFrom) : '',
          leaveTo: pv2 && pv2.leaveTo ? fmtDate(pv2.leaveTo) : '',
          leaveDays: pv2 ? (pv2.leaveDays || 0) : 0,
          reported: pv2 ? (pv2.reported || '') : '',
          note: pv2 ? (pv2.note || '') : '',
          joinTxt: e.join ? fmtDate(e.join) : '', endTxt: e.end ? fmtDate(e.end) : ''
        });
      });
    }

    res.cover = cover;
    res.ran = true;
    sortRows(res.anomalies); sortRows(res.noRecord); sortRows(res.leftOut); sortRows(res.reportedOut);
    return res;
  }

  function sortRows(rows) {
    rows.sort(function (a, b) {
      var c = String(a.city).localeCompare(String(b.city), 'zh-Hans-CN');
      if (c) return c;
      var p = String(a.project).localeCompare(String(b.project), 'zh-Hans-CN');
      if (p) return p;
      if (b.gap !== a.gap) return b.gap - a.gap;
      return String(a.name).localeCompare(String(b.name), 'zh-Hans-CN');
    });
  }

  return {
    parseFile: parseFile,
    compute: compute,
    fmtDay: fmtDay,
    fmtDate: fmtDate,
    parseDate: toDate,
    daysBetween: daysBetween,
    monthBounds: monthBounds
  };
})();
