/* ============================================================
 * ren-core.js — 文件重命名核心
 * 输入：
 *   人员花名册（必填）：调度中心（蝶城/组织）/ 员工项目 / 员工编码 / 员工姓名
 *   资料文件：按模块上传（多文件或 zip 解包），模块名即窗口名
 * 匹配规则（用户确认）：
 *   ① 文件名按分隔符切段，姓名段 / 编码段 / 姓名子串 → 候选员工
 *   ② 同名同姓时用文件名里的项目段（先比员工项目，再比调度中心）消歧
 *   ③ 都定不下来 → 进问题清单
 * 改名模板（分段式）：可增删段（编码/姓名/模块名/项目/调度中心/日期/序号/固定文字），段间 - 连接，
 *   后缀独立输入框（留空用原文件后缀）；兼容旧花括号模板字符串
 *   目标文件名重复 → 相关文件全部跳过并进问题清单
 *   缺失检查（可选勾选模块）：花名册员工在勾选模块没有文件 → 问题清单
 * ============================================================ */
window.RenCore = (function () {
  'use strict';

  var MODULES = ['身份证正面', '身份证背面', '手持证件照', '入场登记表', '入场承诺书', '劳务合同协议', '投保证明', '推荐激励诚信书', '其他档案材料'];

  /** 特殊模块：投保证明只上传 1 个原始文件，按花名册全员复制重命名 */
  var DUP_MODULE = '投保证明';

  /* ---------------- 基础工具 ---------------- */
  function norm(s) {
    return String(s == null ? '' : s)
      .replace(/[\uFF01-\uFF5E]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
      .replace(/\s|\u3000/g, '')
      .toLowerCase();
  }
  function sanitize(v) {
    return String(v == null ? '' : v).replace(/[\\/:*?"<>|\r\n]/g, '').trim();
  }
  function pad2(n) { return n < 10 ? '0' + n : '' + n; }

  /* ---------------- 花名册解析 ---------------- */
  function normHeader(s) {
    return String(s == null ? '' : s)
      .replace(/[\s\u3000]/g, '')
      .replace(/[（(][^)）]*[)）]/g, '')
      .replace(/[：:*＊]/g, '');
  }

  var FIELD_ROSTER = [
    { k: 'id', label: '员工编码', req: true, aliases: ['员工编码', '工号', 'SHRid', '人员编码'] },
    { k: 'name', label: '员工姓名', req: true, aliases: ['员工姓名', '姓名'] },
    { k: 'center', label: '调度中心', aliases: ['调度中心', '所属调度中心', '蝶城', '所属蝶城', '组织', '所属组织', '行政组织', '所属部门', '部门'] },
    { k: 'project', label: '员工项目', aliases: ['员工项目', '所属项目', '战图项目名称', '项目'] }
  ];

  var SOFT = {
    id: ['员工编码', '人员编号', '工号', 'SHRid', '编码', '编号'],
    name: ['员工姓名', '人员姓名', '姓名', '名字'],
    center: ['调度中心', '所属行政组织名称', '所属蝶城', '组织名称', '蝶城', '所属部门', '部门', '中心'],
    project: ['员工项目', '战图项目名称', '所属项目', '项目']
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

  /** 第一步：找最佳工作表 + 表头行 + 自动猜字段映射（返回原始表头供页面改映射） */
  function parseRosterMeta(wb) {
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
        var cols = {}, hit = 0, missReq = 0;
        FIELD_ROSTER.forEach(function (f) {
          var v = colIdx(headers, f.aliases);
          if (v < 0) v = guessCol(headers, f.k);
          cols[f.k] = v;
          if (v >= 0) hit++;
          if (f.req && v < 0) missReq++;
        });
        if (missReq > 0) continue;
        if (!best || hit > best.hit) best = { aoa: aoa, hIdx: i, cols: cols, sheet: sn, hit: hit,
          headers: row.map(function (x) { return x == null ? '' : String(x).trim(); }) };
      }
    });
    if (!best) return { error: '没有找到匹配的表头（需要列：员工编码、员工姓名）。请确认这份表是不是人员花名册' };
    return { sheet: best.sheet, aoa: best.aoa, hIdx: best.hIdx, cols: best.cols, headers: best.headers };
  }

  /** 第二步：按映射列提取员工行（页面改映射后重新调用） */
  function extractRoster(meta, cols) {
    var c = cols, aoa = meta.aoa, rows = [], seen = {};
    for (var i = meta.hIdx + 1; i < aoa.length; i++) {
      var row = aoa[i]; if (!row) continue;
      var id = c.id >= 0 ? String(row[c.id] == null ? '' : row[c.id]).trim() : '';
      var name = c.name >= 0 ? String(row[c.name] == null ? '' : row[c.name]).trim() : '';
      if (!id || !name || id.toUpperCase() === 'NULL') continue;
      if (seen[id]) continue; // 同编码取第一行
      seen[id] = true;
      rows.push({
        seq: rows.length + 1,
        id: id,
        name: name,
        center: c.center >= 0 ? String(row[c.center] == null ? '' : row[c.center]).trim() : '',
        project: c.project >= 0 ? String(row[c.project] == null ? '' : row[c.project]).trim() : ''
      });
    }
    if (!rows.length) return { error: '花名册里没有读到有效员工行（需要工号 + 姓名两列都有值）' };
    return { sheet: meta.sheet, rows: rows };
  }

  function parseRoster(wb, fileName) {
    var meta = parseRosterMeta(wb);
    if (meta.error) return meta;
    var r = extractRoster(meta, meta.cols);
    if (r.error) return r;
    r.fileName = fileName || '';
    return r;
  }

  /* ---------------- 文件名 → 员工匹配 ---------------- */
  function splitSegs(base) {
    return base.split(/[-_—–·,，、;；\s（）()\[\]【】]+/).map(norm).filter(Boolean);
  }

  function matchOne(file, roster) {
    var base = String(file.name || '').replace(/\.[^.]+$/, '');
    var segs = splitSegs(base);
    var nbase = norm(base);

    // ① 按姓名：段相等优先，其次子串
    var cands = roster.filter(function (r) {
      var nn = norm(r.name);
      return nn && (segs.indexOf(nn) >= 0 || nbase.indexOf(nn) >= 0);
    });
    // ② 按编码：只认整段（避免 01 误中 101）
    if (!cands.length) {
      cands = roster.filter(function (r) {
        var ni = norm(r.id);
        return ni && segs.indexOf(ni) >= 0;
      });
    }
    if (!cands.length) {
      return { ok: false, reason: '未匹配到花名册员工（文件名里没有对得上的姓名或编码）', advice: '文件名需包含员工姓名或编码，如「张三.jpg」「R001.jpg」' };
    }

    // ③ 同名同姓消歧：先项目后调度中心
    if (cands.length > 1) {
      var byProj = cands.filter(function (r) {
        var np = norm(r.project);
        return np && (segs.indexOf(np) >= 0 || nbase.indexOf(np) >= 0);
      });
      var pool = byProj.length ? byProj : cands.filter(function (r) {
        var nc = norm(r.center);
        return nc && (segs.indexOf(nc) >= 0 || nbase.indexOf(nc) >= 0);
      });
      if (pool.length === 1) return { ok: true, emp: pool[0] };
      var names = cands.map(function (r) { return r.name + '(' + r.id + (r.project ? '·' + r.project : '') + ')'; }).join('、');
      return {
        ok: false,
        reason: '同名同姓无法区分，候选：' + names,
        advice: '文件名中加入员工编码或所属项目，如「' + cands[0].id + '-' + cands[0].name + '.jpg」'
      };
    }
    return { ok: true, emp: cands[0] };
  }

  function renderName(tpl, emp, module, ext, dateStr) {
    var v = {
      '编码': emp ? emp.id : '',
      '姓名': emp ? emp.name : '',
      '模块名': module || '',
      '项目': emp ? (emp.project || '') : '',
      '调度中心': emp ? (emp.center || '') : '',
      '日期': dateStr || '',
      '序号': emp ? pad2(emp.seq) : '',
      '后缀': ext || ''
    };
    var out = String(tpl || '');
    Object.keys(v).forEach(function (k) {
      out = out.split('{' + k + '}').join(sanitize(v[k]));
    });
    return out;
  }

  /** 分段模板渲染：segs=[{t:'var',v:'编码'},...|{t:'text',v:'文字'}]，段间 - 连接，空段跳过 */
  function segParts(segs, emp, module, dateStr) {
    var v = {
      '编码': emp ? emp.id : '',
      '姓名': emp ? emp.name : '',
      '模块名': module || '',
      '项目': emp ? (emp.project || '') : '',
      '调度中心': emp ? (emp.center || '') : '',
      '日期': dateStr || '',
      '序号': emp ? pad2(emp.seq) : ''
    };
    var parts = [];
    (segs || []).forEach(function (s) {
      var val = s && s.t === 'text' ? String(s.v || '') : (v[s && s.v] || '');
      val = sanitize(val);
      if (val) parts.push(val);
    });
    return parts;
  }

  /** 目标文件名：tplCfg 为字符串（旧花括号模板）或 {segs, ext, sep}（分段模板，sep 留空回退 -） */
  function buildTarget(tplCfg, emp, module, origExt, dateStr) {
    if (tplCfg && typeof tplCfg === 'object' && tplCfg.segs) {
      var parts = segParts(tplCfg.segs, emp, module, dateStr);
      // 固定尾段：工号-类型 用连接符连在一起、固定在最后（后缀之前），不可配置
      var tail = [sanitize(emp ? emp.id : ''), sanitize(module || '')].filter(Boolean);
      parts = parts.concat(tail);
      if (!parts.length) return '';
      var ext = String(tplCfg.ext || '').replace(/^\.+/, '').trim() || origExt || '';
      var sep = (tplCfg && typeof tplCfg.sep === 'string' && tplCfg.sep.trim()) ? tplCfg.sep.trim() : '-';
      return parts.join(sep) + (ext ? '.' + ext : '');
    }
    return renderName(tplCfg, emp, module, origExt, dateStr);
  }

  /**
   * files: [{module, name, data}]  roster: parseRoster().rows
   * 投保证明（DUP_MODULE）特殊：>1 个文件直接阻断；=1 个时按花名册全员复制重命名
   * 返回 { ok:[{file,target,emp,module,data}], fails:[{file,module,reason,advice}], empCount }
   *       或 { error } 表示阻断
   */
  function matchAll(files, roster, tplCfg, dateStr) {
    var all = files || [];
    var dupFiles = all.filter(function (f) { return f.module === DUP_MODULE; });
    var others = all.filter(function (f) { return f.module !== DUP_MODULE; });
    if (dupFiles.length > 1) {
      return { error: '「' + DUP_MODULE + '」只需上传 1 个原始文件，当前上传了 ' + dupFiles.length + ' 个，请清空该窗口多余文件后重新开始' };
    }
    var ok = [], fails = [];
    others.forEach(function (f) {
      var extM = String(f.name || '').match(/\.([^.]+)$/);
      var ext = extM ? extM[1] : '';
      var m = matchOne(f, roster);
      if (!m.ok) {
        fails.push({ file: f.name, module: f.module, reason: m.reason, advice: m.advice });
        return;
      }
      var target = buildTarget(tplCfg, m.emp, f.module, ext, dateStr);
      if (!target || target === '.' || !/[^.\s]/.test(target.replace(/\.\w+$/, ''))) {
        fails.push({ file: f.name, module: f.module, reason: '模板渲染结果为空', advice: '检查改名模板的分段配置，至少保留一个有效段' });
        return;
      }
      ok.push({ file: f.name, target: target, emp: m.emp, module: f.module, data: f.data });
    });

    // 投保证明：单人原始文件 × 花名册全员复制
    if (dupFiles.length === 1) {
      var df = dupFiles[0];
      var extM2 = String(df.name || '').match(/\.([^.]+)$/);
      var ext2 = extM2 ? extM2[1] : '';
      (roster || []).forEach(function (r) {
        var target = buildTarget(tplCfg, r, DUP_MODULE, ext2, dateStr);
        if (!target || target === '.' || !/[^.\s]/.test(target.replace(/\.\w+$/, ''))) {
          fails.push({ file: df.name + '（复制给 ' + r.name + '）', module: DUP_MODULE, reason: '模板渲染结果为空', advice: '检查改名模板的分段配置，至少保留一个有效段' });
          return;
        }
        ok.push({ file: df.name, target: target, emp: r, module: DUP_MODULE, data: df.data });
      });
    }

    // 目标文件名重复 → 相关文件全部跳过
    var byName = {};
    ok.forEach(function (x) { (byName[x.target] = byName[x.target] || []).push(x); });
    Object.keys(byName).forEach(function (t) {
      if (byName[t].length > 1) {
        byName[t].forEach(function (x) {
          fails.push({
            file: x.file, module: x.module,
            reason: '目标文件名重复：「' + t + '」共 ' + byName[t].length + ' 个文件',
            advice: x.module === DUP_MODULE ? '投保证明复制重名：花名册中存在重复的员工编码，请检查花名册数据' : '删除多余重复文件后重新上传'
          });
        });
        ok = ok.filter(function (x) { return byName[x.target].length === 1; });
      }
    });

    var empSet = {};
    ok.forEach(function (x) { empSet[x.emp.id] = true; });
    return { ok: ok, fails: fails, empCount: Object.keys(empSet).length };
  }

  /** 缺失检查：checkedModules 内，花名册员工在模块下没有成功重命名的文件 → 问题清单 */
  function findMissing(matchRes, roster, checkedModules) {
    var have = {};
    matchRes.ok.forEach(function (x) { have[x.emp.id + '|' + x.module] = true; });
    var out = [];
    roster.forEach(function (r) {
      (checkedModules || []).forEach(function (m) {
        if (!have[r.id + '|' + m]) {
          out.push({
            file: '—', module: m,
            reason: '员工「' + r.name + '(' + r.id + ')」缺少「' + m + '」',
            advice: '补传该员工的' + m + '资料（文件名需包含姓名或编码）'
          });
        }
      });
    });
    return out;
  }

  return {
    MODULES: MODULES,
    DUP_MODULE: DUP_MODULE,
    parseRoster: parseRoster,
    parseRosterMeta: parseRosterMeta,
    extractRoster: extractRoster,
    matchAll: matchAll,
    findMissing: findMissing,
    renderName: renderName,
    buildTarget: buildTarget,
    norm: norm,
    pad2: pad2
  };
})();
