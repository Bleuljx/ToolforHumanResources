/* ============================================================
 * xlsxw.js — 从零生成 xlsx（支持样式、冻结、筛选、列宽行高、下拉验证、隐藏表）
 * 为什么自己写：SheetJS 社区版写不出样式/数据验证，只能读取
 * ============================================================ */
(function (root, factory) {
  var api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.XlsxWriter = api;
})(typeof window !== 'undefined' ? window : globalThis, function () {
  'use strict';

  var NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  var RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

  function esc(s) {
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&apos;')
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  }

  function colName(n) { // 1 -> A
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  /* 样式表：
   * 0 默认 | 1 表头(蓝底白字加粗居中顶对齐) | 2 数据(细边框居中) | 3 数据+两位小数
   * 4 无边框(隐藏选项表) | 5 表头(无底色,仅加粗边框) | 6 左对齐数据
   * 7 考勤表头(浅蓝底99CCFF 宋体14加粗黑字) | 8 珊瑚红标红单元格 | 9 合计行(黄底加粗)
   * 10 黄表头单元格(FFFF00 宋体14加粗) | 11 数据+一位小数
   * 12 珊瑚红+两位小数 | 13 珊瑚红+一位小数 | 14 合计行+两位小数
   * 19 淡黄FFF2CC底+加粗+两位小数（计薪工时0.1-0.9高亮） | 20 浅蓝底宋体14加粗黑字+水平居中垂直靠上（矩阵表头）
   */
  function stylesXml() {
    return HEAD +
      '<styleSheet xmlns="' + NS + '">' +
      '<numFmts count="2"><numFmt numFmtId="164" formatCode="0.00"/><numFmt numFmtId="165" formatCode="0.0"/></numFmts>' +
      '<fonts count="7">' +
      '<font><sz val="11"/><color theme="1"/><name val="宋体"/><charset val="134"/></font>' +
      '<font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="宋体"/><charset val="134"/></font>' +
      '<font><b/><sz val="11"/><color theme="1"/><name val="宋体"/><charset val="134"/></font>' +
      '<font><b/><sz val="14"/><color theme="1"/><name val="宋体"/><charset val="134"/></font>' +
      '<font><sz val="14"/><color theme="1"/><name val="宋体"/><charset val="134"/></font>' +
      '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><charset val="134"/></font>' +
      '<font><b/><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/><charset val="134"/></font>' +
      '</fonts>' +
      '<fills count="8">' +
      '<fill><patternFill patternType="none"/></fill>' +
      '<fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FF0070C0"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FF99CCFF"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFFF9999"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFCC"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>' +
      '</fills>' +
      '<borders count="2">' +
      '<border><left/><right/><top/><bottom/><diagonal/></border>' +
      '<border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right>' +
      '<top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border>' +
      '</borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="21">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="top" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="left" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="2" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="3" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="164" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="165" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="164" fontId="2" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      // 15: Calibri 11 加粗黑字无边框（宽表表头）；16: 红字表头；17: 无边框居中数据；18: 无边框 0.00 数据
      '<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      // 19: 淡黄FFF2CC底+加粗+两位小数（计薪工时0.1-0.9高亮）
      '<xf numFmtId="164" fontId="2" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      // 20: 浅蓝底99CCFF 宋体14加粗黑字 水平居中垂直靠上（矩阵表头）
      '<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
      '<alignment horizontal="center" vertical="top" wrapText="1"/></xf>' +
      '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="0"/>' +
      '</styleSheet>';
  }

  function paneXml(s) {
    // s.freeze = {x, y}（冻结列数/行数）；s.freezeHeader 等价 {x:0,y:1}
    var f = s.freeze || (s.freezeHeader ? { x: 0, y: 1 } : null);
    if (!f || (!f.x && !f.y)) return '';
    var tl = colName(f.x + 1) + (f.y + 1);
    return '<pane' + (f.x ? ' xSplit="' + f.x + '"' : '') + (f.y ? ' ySplit="' + f.y + '"' : '') +
      ' topLeftCell="' + tl + '" activePane="bottomRight" state="frozen"/>' +
      '<selection pane="bottomRight" activeCell="' + tl + '" sqref="' + tl + '"/>';
  }

  function sheetXml(s) {
    var header = s.header || null;
    var rows = s.rows || [];
    var nCols = header ? header.length : 0;
    rows.forEach(function (r) { if (r.length > nCols) nCols = r.length; });
    (s.cols || []).forEach(function (c, i) { if (i + 1 > nCols) nCols = i + 1; });
    if (nCols < 1) nCols = 1;

    var totalRows = rows.length + (header ? 1 : 0);
    var ref = 'A1:' + colName(nCols) + Math.max(1, totalRows);

    var x = HEAD + '<worksheet xmlns="' + NS + '" xmlns:r="' + RNS + '">' +
      '<dimension ref="' + ref + '"/>' +
      '<sheetViews><sheetView workbookViewId="0">' +
      paneXml(s) +
      '</sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="16.5"/>';

    if (s.cols && s.cols.length) {
      x += '<cols>';
      s.cols.forEach(function (w, i) {
        if (w == null) return;
        x += '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>';
      });
      x += '</cols>';
    }

    x += '<sheetData>';
    var rowNo = 1;
    if (header) {
      x += '<row r="1"' + (s.headerHeight ? ' ht="' + s.headerHeight + '" customHeight="1"' : '') + '>';
      for (var c = 0; c < header.length; c++) {
        if (header[c] == null || header[c] === '') continue;
        var hStyle = s.headerStyle == null ? 1 : s.headerStyle;
        if (s.headerStyleFn) { var hv = s.headerStyleFn(c); if (hv != null) hStyle = hv; }
        x += cellXml(colName(c + 1) + '1', header[c], hStyle);
      }
      x += '</row>';
      rowNo = 2;
    }
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      x += '<row r="' + rowNo + '"' + (s.rowHeight ? ' ht="' + s.rowHeight + '" customHeight="1"' : '') + '>';
      for (var cc = 0; cc < nCols; cc++) {
        var v = row[cc];
        var st = s.dataStyle == null ? 2 : s.dataStyle;
        if (s.numFmtCols && s.numFmtCols.indexOf(cc) >= 0) st = 3;
        if (s.styleFn) { var ov = s.styleFn(r, cc); if (ov != null) st = ov; }
        // 空单元格也写入样式：保证整个数据区都有框线（无边框样式 0/4 除外）
        if (v == null || v === '') {
          if (st === 0 || st === 4) continue;
          x += cellXml(colName(cc + 1) + rowNo, '', st);
          continue;
        }
        x += cellXml(colName(cc + 1) + rowNo, v, st);
      }
      x += '</row>';
      rowNo++;
    }
    x += '</sheetData>';

    if (s.autoFilter && totalRows > 1) {
      x += '<autoFilter ref="A1:' + colName(nCols) + totalRows + '"/>';
    }
    if (s.validations && s.validations.length) {
      x += '<dataValidations count="' + s.validations.length + '">';
      s.validations.forEach(function (v) {
        x += '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="' + v.sqref + '">' +
          '<formula1>' + esc(v.formula) + '</formula1></dataValidation>';
      });
      x += '</dataValidations>';
    }
    x += '</worksheet>';
    return x;
  }

  function cellXml(ref, v, styleId) {
    var sAttr = ' s="' + styleId + '"';
    // 空单元格：只带样式（框线），不写值
    if (v === '' || v == null) return '<c r="' + ref + '"' + sAttr + '/>';
    if (typeof v === 'number' && isFinite(v)) {
      return '<c r="' + ref + '"' + sAttr + '><v>' + v + '</v></c>';
    }
    if (typeof v === 'boolean') {
      return '<c r="' + ref + '"' + sAttr + ' t="b"><v>' + (v ? 1 : 0) + '</v></c>';
    }
    var txt = String(v);
    return '<c r="' + ref + '"' + sAttr + ' t="inlineStr"><is><t xml:space="preserve">' + esc(txt) + '</t></is></c>';
  }

  function build(spec) {
    var sheets = spec.sheets || [];
    var files = [];

    var ct = HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
    sheets.forEach(function (s, i) {
      ct += '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    });
    ct += '</Types>';

    files.push({ name: '[Content_Types].xml', data: ct });
    files.push({
      name: '_rels/.rels',
      data: HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>'
    });

    var wbXml = HEAD + '<workbook xmlns="' + NS + '" xmlns:r="' + RNS + '"><sheets>';
    var rels = HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    sheets.forEach(function (s, i) {
      wbXml += '<sheet name="' + esc(s.name) + '" sheetId="' + (i + 1) + '"' +
        (s.hidden ? ' state="hidden"' : '') + ' r:id="rId' + (i + 1) + '"/>';
      rels += '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>';
    });
    wbXml += '</sheets></workbook>';
    rels += '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
    rels += '</Relationships>';

    files.push({ name: 'xl/workbook.xml', data: wbXml });
    files.push({ name: 'xl/_rels/workbook.xml.rels', data: rels });
    files.push({ name: 'xl/styles.xml', data: stylesXml() });
    sheets.forEach(function (s, i) {
      files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: sheetXml(s) });
    });

    return (spec.zipImpl || (typeof ZipUtil !== 'undefined' ? ZipUtil : null)).zip(files);
  }

  return { build: build, colName: colName };
});
