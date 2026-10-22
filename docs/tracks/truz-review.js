/* TRUZ YG검토본 빌더 — 회계팀 제출용
 * IPX 원본 엑셀 패키지를 그대로 둔 채(서식·DashBoard 피벗·차트·숨김 시트 유지) 아래 시트만 붙인다. 양식 기준: 26.2Q YG검토본.
 *   자재       원본 상세 시트를 같은 열 구성(정산월~처리 일시 + 환율·KRW·멤버)으로 다시 씀. KRW = 수수료×환율, 멤버 = VLOOKUP(제품 명, 멤버!A:C)
 *   PIVOT      피벗①(사업구분·Royalty% × 월, 로열티 합계) + 카테고리 표·분기별 누적 로열티 + 피벗②(월·사업구분·채널·Royalty%) + 정리본
 *   PIVOT 멤버 피벗(필터 채널·제작방식 / 행 멤버 / 열 국가 / KRW 합계) + 정리본
 *   멤버       자재명 / 캐릭터 / 멤버
 * 피벗은 계산 결과를 미리 채워 두고 refreshOnLoad 로 엑셀이 열 때 원본 범위에서 다시 계산한다. 회사 데이터는 이 파일에 없다. */
(function (root) {
  var NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  var NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
  var CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';
  var ACCT = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-';
  var X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
  var FILL_DOWN = '<extLst><ext uri="{2946ED86-A175-432a-8AC1-64E0C546D7DE}" xmlns:x14="' + X14 + '"><x14:pivotField fillDownLabels="1"/></ext></extLst>';
  var PT_EXT = '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>' +
    '<extLst><ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="' + X14 + '"><x14:pivotTableDefinition hideValuesRow="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/></ext></extLst>';
  var PT_ATTR = 'applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="값" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0"';
  var MON = { Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6, Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12 };

  var esc = function (s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); };
  var attrName = function (s) { return esc(String(s).replace(/\r\n/g, '\n')).replace(/\n/g, '_x000a_'); };
  var N = function (s) { return String(s == null ? '' : s).replace(/\s+/g, ''); };
  var col = function (i) { var s = ''; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  // 엑셀 피벗 기본 정렬과 같은 순서: 기호 → 숫자 → 영문(대소문자 무시) → 한글
  var rank = function (ch) { var c = ch.charCodeAt(0); return /[0-9]/.test(ch) ? 1 : /[a-z]/i.test(ch) ? 2 : c >= 0xAC00 && c <= 0xD7A3 ? 3 : c > 127 ? 4 : 0; };
  var ko = function (a, b) {
    a = String(a); b = String(b);
    for (var i = 0; i < Math.min(a.length, b.length); i++) {
      var x = a[i], y = b[i]; if (x === y) continue;
      var rx = rank(x), ry = rank(y); if (rx !== ry) return rx - ry;
      if (rx === 4) { var lc = x.localeCompare(y, 'ko'); if (lc) return lc; }
      var lx = x.toLowerCase(), ly = y.toLowerCase(); if (lx !== ly) return lx < ly ? -1 : 1;
    }
    return a.length - b.length;
  };
  var uniq = function (a) { return a.filter(function (v, i) { return a.indexOf(v) === i; }); };

  // ---------- 스타일: 원본 styles.xml 끝에 필요한 서식만 덧붙인다 ----------
  function Styles(xml) { this.xml = xml; this.memo = {}; this.font = {}; this.fill = {}; this.border = null; this.nf = {}; }
  Styles.prototype.count = function (tag) { var m = this.xml.match(new RegExp('<' + tag + ' count="(\\d+)"')); return m ? +m[1] : 0; };
  Styles.prototype.append = function (tag, item) {
    var n = this.count(tag);
    this.xml = this.xml.replace(new RegExp('<' + tag + ' count="\\d+"'), '<' + tag + ' count="' + (n + 1) + '"').replace('</' + tag + '>', item + '</' + tag + '>');
    return n;
  };
  Styles.prototype.numFmt = function (code) {
    var builtin = { General: 0, '0%': 9, '0.00%': 10, '#,##0': 3 };
    if (code in builtin) return builtin[code];
    if (this.nf[code] != null) return this.nf[code];
    if (this.xml.indexOf('<numFmts') < 0) this.xml = this.xml.replace(/(<styleSheet[^>]*>)/, '$1<numFmts count="0"></numFmts>');
    var ids = (this.xml.match(/numFmtId="(\d+)"/g) || []).map(function (s) { return +s.replace(/\D/g, ''); });
    var id = Math.max.apply(null, [175].concat(ids)) + 1;
    this.append('numFmts', '<numFmt numFmtId="' + id + '" formatCode="' + esc(code) + '"/>');
    return (this.nf[code] = id);
  };
  Styles.prototype.fontId = function (bold) {
    if (!bold) return 0;
    if (this.font.b != null) return this.font.b;
    var first = this.xml.match(/<fonts[^>]*>\s*(<font>[\s\S]*?<\/font>|<font\/>)/)[1];
    var b = first === '<font/>' ? '<font><b/></font>' : first.replace('<font>', '<font><b/>');
    return (this.font.b = this.append('fonts', b));
  };
  Styles.prototype.fillId = function (key) {
    if (!key) return 0;
    if (this.fill[key] != null) return this.fill[key];
    var fg = key === 'gray' ? '<fgColor theme="0" tint="-4.9989318521683403E-2"/>' : key === 'blue' ? '<fgColor theme="4" tint="0.79998168889431442"/>' :
      key === 'green' ? '<fgColor theme="5" tint="0.79998168889431442"/>' : '<fgColor rgb="FFFFFF00"/>';
    return (this.fill[key] = this.append('fills', '<fill><patternFill patternType="solid">' + fg + '<bgColor indexed="64"/></patternFill></fill>'));
  };
  Styles.prototype.borderId = function (on) {
    if (!on) return 0;
    if (this.border != null) return this.border;
    var s = function (t) { return '<' + t + ' style="thin"><color indexed="64"/></' + t + '>'; };
    return (this.border = this.append('borders', '<border>' + s('left') + s('right') + s('top') + s('bottom') + '<diagonal/></border>'));
  };
  // o: { b, fill: gray|blue|green|yellow, nf, bd, h, v }
  Styles.prototype.xf = function (o) {
    var key = JSON.stringify(o);
    if (this.memo[key] != null) return this.memo[key];
    var nf = this.numFmt(o.nf || 'General'), fo = this.fontId(o.b), fi = this.fillId(o.fill), bd = this.borderId(o.bd);
    var al = o.h || o.v ? '<alignment' + (o.h ? ' horizontal="' + o.h + '"' : '') + (o.v ? ' vertical="' + o.v + '"' : '') + '/>' : '';
    var xf = '<xf numFmtId="' + nf + '" fontId="' + fo + '" fillId="' + fi + '" borderId="' + bd + '" xfId="0"' + (nf ? ' applyNumberFormat="1"' : '') + (fo ? ' applyFont="1"' : '') +
      (fi ? ' applyFill="1"' : '') + (bd ? ' applyBorder="1"' : '') + (al ? ' applyAlignment="1">' + al + '</xf>' : '/>');
    return (this.memo[key] = this.append('cellXfs', xf));
  };

  // ---------- 시트 XML ----------
  function Sheet() { this.cells = {}; this.cols = []; this.merges = []; this.autoFilter = null; }
  // v: 숫자·문자열·null, opt: { f: 수식, s: 스타일 번호, err: true(오류 값) }
  Sheet.prototype.set = function (c, r, v, s, f, err) { this.cells[r] = this.cells[r] || {}; this.cells[r][c] = { v: v, s: s, f: f, err: err }; };
  Sheet.prototype.xml = function () {
    var rows = Object.keys(this.cells).map(Number).sort(function (a, b) { return a - b; }), maxC = 0, out = [];
    rows.forEach(function (r) {
      var cs = this.cells[r], keys = Object.keys(cs).map(Number).sort(function (a, b) { return a - b; });
      maxC = Math.max(maxC, keys[keys.length - 1] || 0);
      out.push('<row r="' + r + '">' + keys.map(function (c) {
        var x = cs[c], ref = col(c) + r, s = x.s ? ' s="' + x.s + '"' : '', f = x.f ? '<f>' + esc(x.f) + '</f>' : '';
        if (x.err) return '<c r="' + ref + '"' + s + ' t="e">' + f + '<v>' + esc(x.v) + '</v></c>';
        if (x.v == null || x.v === '') return f ? '<c r="' + ref + '"' + s + ' t="str">' + f + '<v></v></c>' : '<c r="' + ref + '"' + s + '/>';
        if (typeof x.v === 'number' && isFinite(x.v)) return '<c r="' + ref + '"' + s + '>' + f + '<v>' + x.v + '</v></c>';
        if (f) return '<c r="' + ref + '"' + s + ' t="str">' + f + '<v>' + esc(x.v) + '</v></c>';
        return '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + esc(String(x.v).replace(/\r\n/g, '\n')) + '</t></is></c>';
      }).join('') + '</row>');
    }, this);
    var dim = rows.length ? 'A1:' + col(maxC) + rows[rows.length - 1] : 'A1';
    var cols = this.cols.length ? '<cols>' + this.cols.map(function (w, i) { return w ? '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>' : ''; }).join('') + '</cols>' : '';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="' + NS_MAIN + '" xmlns:r="' + NS_R + '"><dimension ref="' + dim + '"/>' +
      '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="16.5"/>' + cols + '<sheetData>' + out.join('') + '</sheetData>' +
      (this.autoFilter ? '<autoFilter ref="' + this.autoFilter + '"/>' : '') +
      (this.merges.length ? '<mergeCells count="' + this.merges.length + '">' + this.merges.map(function (m) { return '<mergeCell ref="' + m + '"/>'; }).join('') + '</mergeCells>' : '') +
      '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
  };

  // ---------- 피벗 캐시 (원본 범위 + 레코드) ----------
  // fields: [{ name, vals: [...], order: [...] (축 필드만: 항목 순서) }]
  function itemXml(v) {
    if (v && v.err) return '<e v="' + esc(v.err) + '"/>';
    if (v == null || v === '') return '<m/>';
    if (typeof v === 'number') return '<n v="' + v + '"/>';
    return '<s v="' + esc(String(v).replace(/\r\n/g, '\n')) + '"/>';
  }
  function typeAttrs(vals, counted) {
    var str = 0, num = 0, blank = 0, err = 0, int = 1, min = Infinity, max = -Infinity;
    vals.forEach(function (v) {
      if (v && v.err) err = 1; else if (v == null || v === '') blank = 1;
      else if (typeof v === 'number') { num = 1; if (v % 1) int = 0; min = Math.min(min, v); max = Math.max(max, v); } else str = 1;
    });
    var a = '';
    if (!str && !blank && !err && num) a += ' containsSemiMixedTypes="0" containsString="0"';
    else if (!str && num) a += ' containsString="0"';
    if (blank) a += ' containsBlank="1"';
    if (num && (str || err)) a += ' containsMixedTypes="1"';
    if (num) a += ' containsNumber="1"' + (int ? ' containsInteger="1"' : '') + ' minValue="' + min + '" maxValue="' + max + '"';
    if (counted != null) a = ' count="' + counted + '"' + a;
    return a;
  }
  function cacheParts(sheetName, ref, fields, nrec) {
    var idx = fields.map(function (f) {
      if (!f.order) return null;
      var m = {}; f.order.forEach(function (v, i) { m[v && v.err ? '#' + v.err : String(v)] = i; }); return m;
    });
    var def = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="' + NS_MAIN + '" xmlns:r="' + NS_R + '" r:id="rId1" refreshOnLoad="1" refreshedBy="정산 데스크" refreshedDate="' +
      (25569 + Date.now() / 864e5).toFixed(6) + '" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="' + nrec + '">' +
      '<cacheSource type="worksheet"><worksheetSource ref="' + ref + '" sheet="' + esc(sheetName) + '"/></cacheSource><cacheFields count="' + fields.length + '">' +
      fields.map(function (f) {
        var nf = f.numFmtId ? ' numFmtId="' + f.numFmtId + '"' : ' numFmtId="0"';
        if (f.order) return '<cacheField name="' + attrName(f.name) + '"' + nf + '><sharedItems' + typeAttrs(f.order, f.order.length) + '>' + f.order.map(itemXml).join('') + '</sharedItems></cacheField>';
        return '<cacheField name="' + attrName(f.name) + '"' + nf + '><sharedItems' + typeAttrs(f.vals) + '/></cacheField>';
      }).join('') + '</cacheFields></pivotCacheDefinition>';
    var rec = [];
    for (var i = 0; i < nrec; i++) {
      rec.push('<r>' + fields.map(function (f, j) {
        var v = f.vals[i];
        if (idx[j]) return '<x v="' + idx[j][v && v.err ? '#' + v.err : String(v)] + '"/>';
        return itemXml(v);
      }).join('') + '</r>');
    }
    var records = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords xmlns="' + NS_MAIN + '" xmlns:r="' + NS_R + '" count="' + nrec + '">' + rec.join('') + '</pivotCacheRecords>';
    return { def: def, records: records };
  }
  // 축 필드 항목 목록 (hidden: 숨길 값 목록)
  function itemsXml(order, withDefault, hidden) {
    var it = order.map(function (v, i) { return '<item' + (hidden && hidden.indexOf(v) >= 0 ? ' h="1"' : '') + ' x="' + i + '"/>'; });
    if (withDefault) it.push('<item t="default"/>');
    return '<items count="' + it.length + '">' + it.join('') + '</items>';
  }

  // ---------- 원본 Data 시트 읽기 (피벗 원본 범위: Year ~ Royalty Total Amt.) ----------
  function dataRange(ws) {
    var R = XLSX.utils.decode_range(ws['!ref']), hr = -1, c0 = -1;
    for (var r = R.s.r; r <= Math.min(R.e.r, 40) && hr < 0; r++) for (var c = R.s.c; c <= R.e.c; c++) { var x = ws[XLSX.utils.encode_cell({ r: r, c: c })]; if (x && N(x.v) === 'Year') { hr = r; c0 = c; break; } }
    if (hr < 0) throw new Error('Data 시트 머리글(Year)을 찾지 못했습니다');
    var head = [], c1 = c0;
    for (var cc = c0; cc <= R.e.c; cc++) { var h = ws[XLSX.utils.encode_cell({ r: hr, c: cc })]; if (!h || h.v == null || h.v === '') break; head.push(String(h.v).replace(/\r\n/g, '\n')); c1 = cc; if (N(h.v) === 'RoyaltyTotalAmt.') break; }
    var rows = [], last = hr;
    for (var rr = hr + 1; rr <= R.e.r; rr++) {
      var y = ws[XLSX.utils.encode_cell({ r: rr, c: c0 })];
      if (!y || y.v == null || y.v === '') continue;
      var row = []; for (var k = c0; k <= c1; k++) { var v = ws[XLSX.utils.encode_cell({ r: rr, c: k })]; row.push(v ? v.v : null); }
      rows.push(row); last = rr;
    }
    return { head: head, rows: rows, ref: col(c0) + (hr + 1) + ':' + col(c1) + (last + 1) };
  }

  // ---------- 패키지 조작 ----------
  async function build(bytes, o) {
    var zip = await JSZip.loadAsync(bytes);
    var rd = function (p) { var f = zip.file(p); return f ? f.async('string') : Promise.resolve(null); };
    var wbXml = await rd('xl/workbook.xml'), wbRels = await rd('xl/_rels/workbook.xml.rels'), ctXml = await rd('[Content_Types].xml');
    var st = new Styles(await rd('xl/styles.xml'));
    var sheets = []; wbXml.replace(/<sheet [^>]*\/>/g, function (s) { sheets.push({ tag: s, name: s.match(/name="([^"]*)"/)[1], rid: s.match(/r:id="([^"]*)"/)[1] }); });
    var unx = function (s) { return s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'"); };
    sheets.forEach(function (s) { s.name = unx(s.name); });
    ['PIVOT', 'PIVOT 멤버', '멤버'].forEach(function (n) { if (sheets.some(function (s) { return s.name === n; })) throw new Error('이미 YG검토본입니다 — IPX 원본 파일을 올려 주세요'); });
    var detIdx = sheets.findIndex(function (s) { return s.name === o.p.detailSheet; });
    if (detIdx < 0) throw new Error('상세 시트를 찾지 못했습니다');
    var target = function (rid) { var m = wbRels.match(new RegExp('<Relationship [^>]*Id="' + rid + '"[^>]*>')); return 'xl/' + m[0].match(/Target="([^"]*)"/)[1].replace(/^\/xl\//, ''); };
    var maxN = function (re) { var n = 0; Object.keys(zip.files).forEach(function (p) { var m = p.match(re); if (m) n = Math.max(n, +m[1]); }); return n; };
    var rid = 0; var newRid = function () { rid++; return 'rIdSD' + rid; };
    var addCT = function (part, type) { ctXml = ctXml.replace('</Types>', '<Override PartName="/' + part + '" ContentType="' + CT + type + '+xml"/></Types>'); };
    var relsXml = function (list) { return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="' + NS_PKG + '">' + list.map(function (x) { return '<Relationship Id="' + x[0] + '" Type="' + NS_R + '/' + x[1] + '" Target="' + x[2] + '"/>'; }).join('') + '</Relationships>'; };
    var sheetNo = maxN(/^xl\/worksheets\/sheet(\d+)\.xml$/), ptNo = maxN(/^xl\/pivotTables\/pivotTable(\d+)\.xml$/), pcNo = maxN(/^xl\/pivotCache\/pivotCacheDefinition(\d+)\.xml$/);
    var ids = (wbXml.match(/<pivotCache cacheId="(\d+)"/g) || []).map(function (s) { return +s.replace(/\D/g, ''); });
    var cacheId = Math.max.apply(null, [0].concat(ids));
    var sheetIds = (wbXml.match(/sheetId="(\d+)"/g) || []).map(function (s) { return +s.replace(/\D/g, ''); }), sheetId = Math.max.apply(null, sheetIds);

    function addCache(parts) {
      pcNo++; cacheId++;
      var defP = 'xl/pivotCache/pivotCacheDefinition' + pcNo + '.xml', recP = 'xl/pivotCache/pivotCacheRecords' + pcNo + '.xml';
      zip.file(defP, parts.def); zip.file(recP, parts.records);
      zip.file('xl/pivotCache/_rels/pivotCacheDefinition' + pcNo + '.xml.rels', relsXml([['rId1', 'pivotCacheRecords', 'pivotCacheRecords' + pcNo + '.xml']]));
      addCT(defP, 'pivotCacheDefinition'); addCT(recP, 'pivotCacheRecords');
      var r = newRid(); wbRels = wbRels.replace('</Relationships>', '<Relationship Id="' + r + '" Type="' + NS_R + '/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition' + pcNo + '.xml"/></Relationships>');
      var tag = '<pivotCache cacheId="' + cacheId + '" r:id="' + r + '"/>';
      wbXml = wbXml.indexOf('</pivotCaches>') >= 0 ? wbXml.replace('</pivotCaches>', tag + '</pivotCaches>') : wbXml.replace(/(<\/calcPr>|<calcPr[^>]*\/>)/, '$1<pivotCaches>' + tag + '</pivotCaches>');
      return { id: cacheId, file: 'pivotCacheDefinition' + pcNo + '.xml' };
    }
    function addSheet(name, sh, pivots) {
      sheetNo++; sheetId++;
      var p = 'xl/worksheets/sheet' + sheetNo + '.xml';
      zip.file(p, sh.xml()); addCT(p, 'worksheet');
      var r = newRid();
      wbRels = wbRels.replace('</Relationships>', '<Relationship Id="' + r + '" Type="' + NS_R + '/worksheet" Target="worksheets/sheet' + sheetNo + '.xml"/></Relationships>');
      wbXml = wbXml.replace('</sheets>', '<sheet name="' + esc(name) + '" sheetId="' + sheetId + '" r:id="' + r + '"/></sheets>');
      if (pivots && pivots.length) {
        zip.file('xl/worksheets/_rels/sheet' + sheetNo + '.xml.rels', relsXml(pivots.map(function (pt, i) {
          ptNo++;
          var pp = 'xl/pivotTables/pivotTable' + ptNo + '.xml';
          zip.file(pp, pt.xml); addCT(pp, 'pivotTable');
          zip.file('xl/pivotTables/_rels/pivotTable' + ptNo + '.xml.rels', relsXml([['rId1', 'pivotCacheDefinition', '../pivotCache/' + pt.cache.file]]));
          return ['rId' + (i + 1), 'pivotTable', '../pivotTables/pivotTable' + ptNo + '.xml'];
        })));
      }
    }

    var p = o.p, r = o.r, rates = o.rates, map = o.map;
    var S = {
      hdr: st.xf({ b: true, bd: true }), hdrN: st.xf({ b: true, nf: '#,##0', bd: true }), val: st.xf({ nf: '#,##0_ ', bd: true }), note: st.xf({ nf: '#,##0_ ' }),
      ttlF: st.xf({ b: true, nf: ACCT, bd: true }), ttlRowL: st.xf({ b: true, fill: 'gray', bd: true }), ttlRow: st.xf({ b: true, fill: 'gray', nf: ACCT, bd: true }),
      cumL: st.xf({ nf: ACCT, bd: true }), cumV: st.xf({ nf: '#,##0', bd: true }), cumLB: st.xf({ b: true, nf: ACCT, bd: true }), cumVB: st.xf({ b: true, nf: '#,##0', bd: true }),
      cumTL: st.xf({ b: true, fill: 'gray' }), cumTV: st.xf({ b: true, fill: 'gray', nf: '#,##0' }),
      cpH: st.xf({ b: true, fill: 'blue', bd: true, v: 'center' }), cpMon: st.xf({ b: true, bd: true, v: 'center' }), cpBiz: st.xf({ b: true, bd: true, h: 'center', v: 'center' }),
      cpTxt: st.xf({ bd: true, v: 'center' }), cpPct: st.xf({ nf: '0%', bd: true, v: 'center' }), cpNum: st.xf({ nf: ACCT, bd: true, v: 'center' }),
      cpMonT: st.xf({ b: true, fill: 'blue', bd: true, v: 'center' }), cpB: st.xf({ b: true, bd: true, v: 'center' }), cpBN: st.xf({ b: true, nf: ACCT, bd: true, v: 'center' }),
      cpG: st.xf({ b: true, fill: 'green', bd: true, v: 'center' }), cpGN: st.xf({ b: true, fill: 'green', nf: ACCT, bd: true, v: 'center' }),
      pmMem: st.xf({ b: true, bd: true, h: 'left', v: 'center' }), pmT: st.xf({ b: true, fill: 'blue', bd: true, h: 'left', v: 'center' }), pmTN: st.xf({ b: true, fill: 'blue', nf: ACCT, bd: true, v: 'center' }),
      acct: st.xf({ nf: ACCT }), pct2: st.xf({ nf: '0.00%' }), yel: st.xf({ fill: 'yellow', v: 'center' }), vc: st.xf({ v: 'center' })
    };

    // ===== 자재 =====
    var COLS = ['정산월', '주문월', '국가', '채널', '채널 명', 'IP', '제작방식', '제품 명', '주문 수량', '거래통화', '실매출액', '세액', '즉시환급', '상품할인', '쿠폰할인', '정산대상 금액', '제외비', '수수료율', '수수료', '처리 일시', '환율', 'KRW', '멤버'];
    var dh = p.detailHeader.map(N), srcIx = COLS.map(function (h) { var i = dh.indexOf(N(h)); return i; });
    var memberOf = {}; // 멤버 시트(=VLOOKUP 대상) 기준 이름 → 멤버
    var names = uniq(p.detail.map(function (d) { return d.name; }));
    names.forEach(function (n) { if (map[N(n)]) memberOf[n] = map[N(n)]; });
    var mat = new Sheet(); mat.cols = [8.7, null, 9.3, null, 35, 14.5, 9.3, 55.2, 10.2, 9.3, null, 8, 9.3, 9.8, 9.3, 14.5, 9.2, 9.3, 10.2, 12, 9];
    COLS.forEach(function (h, c) { mat.set(c, 1, h); });
    var recs = [];
    p.detail.forEach(function (d, i) {
      var row = i + 2, rate = d.cur === 'KRW' ? 1 : rates[d.ym + '|' + d.cur];
      var vals = COLS.map(function (h, c) {
        if (h === '국가') return d.country;
        if (h === '채널') return d.channel;
        if (c >= 20) return null;
        return srcIx[c] >= 0 ? d.raw[srcIx[c]] : null;
      });
      vals.forEach(function (v, c) { if (c < 20 && v != null && v !== '') mat.set(c, row, v); });
      // B2B 행은 멤버 배분 대상이 아니라 KRW 0 (26.2Q YG검토본과 같은 처리) → PIVOT 멤버 필터 "(모두)" 그대로 OEM 합계
      var b2b = d.channel === 'B2B', krw = b2b ? 0 : rate == null || rate === '' ? null : (+d.fee || 0) * rate, mem = memberOf[d.name];
      if (rate != null && rate !== '') mat.set(20, row, +rate);
      if (b2b) mat.set(21, row, 0, S.acct); else mat.set(21, row, krw == null ? 0 : krw, S.acct, 'S' + row + '*U' + row);
      if (mem) mat.set(22, row, mem, 0, 'VLOOKUP(H' + row + ',멤버!A:C,3,0)'); else mat.set(22, row, '#N/A', 0, 'VLOOKUP(H' + row + ',멤버!A:C,3,0)', true);
      vals[20] = rate == null || rate === '' ? null : +rate; vals[21] = krw == null ? 0 : krw; vals[22] = mem || { err: '#N/A' };
      recs.push(vals);
    });
    var lastMat = p.detail.length + 1;
    mat.autoFilter = 'A1:W' + lastMat;
    var detPath = target(sheets[detIdx].rid);
    zip.file(detPath, mat.xml());
    // 시트 이름 → 자재, 필터 범위 갱신
    var oldName = sheets[detIdx].name;
    wbXml = wbXml.replace(sheets[detIdx].tag, sheets[detIdx].tag.replace(/name="[^"]*"/, 'name="자재"'));
    wbXml = wbXml.replace(new RegExp('(<definedName name="_xlnm._FilterDatabase" localSheetId="' + detIdx + '"[^>]*>)[^<]*'), function (m0, tag) { return tag + '자재!$A$1:$W$' + lastMat; });
    if (oldName !== '자재') wbXml = wbXml.split("'" + oldName + "'!").join('자재!').split(oldName + '!').join('자재!');

    // ===== PIVOT (Data 시트 피벗 2개 + 표) =====
    var dws = p.wb.Sheets[p.dataSheet], D = dataRange(dws), H = D.head.map(N);
    var fi = function (n) { var i = H.indexOf(n); if (i < 0) throw new Error('Data 시트에 ' + n + ' 열이 없습니다'); return i; };
    var iMon = fi('Month'), iBiz = fi('byBizCat.'), iCh = fi('byChannel'), iPct = fi('Royalty%'), iSal = H.indexOf('SalesTotalAmt.(SubjecttoRoyalty)'), iCost = H.indexOf('CostTotalAmt.'), iRoy = fi('RoyaltyTotalAmt.');
    var ymOf = function (row) { return +row[H.indexOf('Year')] * 100 + (MON[String(row[iMon]).slice(0, 3)] || 0); };
    var monOrder = uniq(D.rows.slice().sort(function (a, b) { return ymOf(a) - ymOf(b); }).map(function (x) { return x[iMon]; }));
    var bizOrder = uniq(D.rows.map(function (x) { return x[iBiz]; })).sort(ko), chOrder = uniq(D.rows.map(function (x) { return x[iCh]; })).sort(ko);
    var pctOrder = uniq(D.rows.map(function (x) { return x[iPct]; })).sort(function (a, b) { return a - b; });
    var dFields = D.head.map(function (h, j) {
      var f = { name: h, vals: D.rows.map(function (x) { return x[j]; }) };
      if (j === iMon) f.order = monOrder; if (j === iBiz) f.order = bizOrder; if (j === iCh) f.order = chOrder; if (j === iPct) f.order = pctOrder;
      return f;
    });
    var dCache = addCache(cacheParts(p.dataSheet, D.ref, dFields, D.rows.length));
    var pv = new Sheet(); pv.cols = [11.2, 13.3, 36.9, 11.5, 39.5, 23.4, 20.3, 19.8, 27.9, 14.6, 34.9, 15.7, 9.5, 23.4, 34.9, 9.5, 39.5, 20.3, 23.4];
    var sum = function (rows, j) { return rows.reduce(function (a, x) { return a + (typeof x[j] === 'number' ? x[j] : 0); }, 0); };
    var royHead = D.head[iRoy], salHead = D.head[iSal], costHead = D.head[iCost];
    // 피벗① A3
    pv.set(0, 3, '합계 : ' + royHead); pv.set(2, 3, '열 레이블');
    pv.set(0, 4, '행 레이블'); pv.set(1, 4, 'Royalty%');
    monOrder.forEach(function (m, k) { pv.set(2 + k, 4, m); }); pv.set(2 + monOrder.length, 4, '총합계');
    var rr = 5;
    bizOrder.forEach(function (b) {
      var first = true;
      pctOrder.forEach(function (pc) {
        var rows = D.rows.filter(function (x) { return x[iBiz] === b && x[iPct] === pc; }); if (!rows.length) return;
        if (first) pv.set(0, rr, b); first = false;
        pv.set(1, rr, pc, S.pct2);
        monOrder.forEach(function (m, k) { var rs = rows.filter(function (x) { return x[iMon] === m; }); if (rs.length) pv.set(2 + k, rr, sum(rs, iRoy), S.acct); });
        pv.set(2 + monOrder.length, rr, sum(rows, iRoy), S.acct); rr++;
      });
    });
    pv.set(0, rr, '총합계');
    monOrder.forEach(function (m, k) { pv.set(2 + k, rr, sum(D.rows.filter(function (x) { return x[iMon] === m; }), iRoy), S.acct); });
    pv.set(2 + monOrder.length, rr, sum(D.rows, iRoy), S.acct);
    var p1End = rr, p1Ref = 'A3:' + col(2 + monOrder.length) + p1End;
    var pt1Fields = D.head.map(function (h, j) {
      if (j === iMon) return '<pivotField axis="axisCol" showAll="0">' + itemsXml(monOrder, true) + '</pivotField>';
      if (j === iBiz) return '<pivotField name="by" axis="axisRow" outline="0" showAll="0" defaultSubtotal="0">' + itemsXml(bizOrder, false) + '</pivotField>';
      if (j === iPct) return '<pivotField axis="axisRow" numFmtId="10" showAll="0">' + itemsXml(pctOrder, true) + '</pivotField>';
      if (j === iRoy) return '<pivotField dataField="1" numFmtId="41" showAll="0"/>';
      if (j === iSal) return '<pivotField numFmtId="41" showAll="0"/>';
      return '<pivotField showAll="0"/>';
    });
    var pt1 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="' + NS_MAIN + '" name="피벗 테이블1" cacheId="' + dCache.id + '" ' + PT_ATTR + '>' +
      '<location ref="' + p1Ref + '" firstHeaderRow="1" firstDataRow="2" firstDataCol="2"/><pivotFields count="' + D.head.length + '">' + pt1Fields.join('') + '</pivotFields>' +
      '<rowFields count="2"><field x="' + iBiz + '"/><field x="' + iPct + '"/></rowFields><colFields count="1"><field x="' + iMon + '"/></colFields>' +
      '<dataFields count="1"><dataField name="합계 : ' + attrName(royHead) + '" fld="' + iRoy + '" baseField="0" baseItem="0" numFmtId="41"/></dataFields>' + PT_EXT + '</pivotTableDefinition>';

    // 카테고리 표 H3 (품의 표와 같은 월별 반올림 값) + 분기별 누적 로열티 H13
    var cats = o.cats, cms = o.months, mlab = cms.map(function (m) { return (+m.slice(5)) + '월'; });
    var H0 = 7, lastC = H0 + cms.length + 1;
    pv.set(H0, 3, '구분', S.hdr); mlab.forEach(function (l, k) { pv.set(H0 + 1 + k, 3, l, S.hdrN); }); pv.set(lastC, 3, 'TTL', S.hdrN);
    cats.forEach(function (c, i) {
      var row = 4 + i;
      pv.set(H0, row, c.label, S.hdr);
      cms.forEach(function (m, k) { pv.set(H0 + 1 + k, row, (r.catMonthR[c.key] || {})[m] || 0, S.val); });
      var tot = cms.reduce(function (a, m) { return a + ((r.catMonthR[c.key] || {})[m] || 0); }, 0);
      pv.set(lastC, row, tot, S.ttlF, 'SUM(' + col(H0 + 1) + row + ':' + col(lastC - 1) + row + ')');
      var q = Math.round(r.cat[c.key] || 0); if (q !== tot) pv.set(lastC + 1, row, q, S.note);
    });
    var tRow = 4 + cats.length, tableSum = 0;
    pv.set(H0, tRow, 'TTL', S.ttlRowL);
    for (var k = 0; k <= cms.length; k++) {
      var cc = H0 + 1 + k, v = 0;
      cats.forEach(function (c) { v += k < cms.length ? ((r.catMonthR[c.key] || {})[cms[k]] || 0) : cms.reduce(function (a, m) { return a + ((r.catMonthR[c.key] || {})[m] || 0); }, 0); });
      if (k === cms.length) tableSum = v;
      pv.set(cc, tRow, v, S.ttlRow, k < cms.length ? 'SUM(' + col(cc) + '4:' + col(cc) + (tRow - 1) + ')' : 'SUM(' + col(H0 + 1) + tRow + ':' + col(lastC - 1) + tRow + ')');
    }
    pv.set(lastC + 1, tRow, tableSum === r.supply ? '합은 맞음' : '공급가액 ' + r.supply.toLocaleString('ko-KR') + ' / 차이 ' + (r.supply - tableSum > 0 ? '+' : '') + (r.supply - tableSum) + '원');
    var cumTop = tRow + 3;
    pv.set(H0, cumTop, '구분', st.xf({ b: true, nf: ACCT, bd: true })); pv.set(H0 + 1, cumTop, '로열티', S.hdrN);
    o.hist.forEach(function (h, i) {
      var y = +h[0].slice(0, 4), q = +h[0].slice(5), lab = y < 2024 || (y === 2024 && q <= 2) ? y + '_Q' + q + '.' : y + '.Q' + q + '.', last = i === o.hist.length - 1;
      pv.set(H0, cumTop + 1 + i, lab, last ? S.cumLB : S.cumL); pv.set(H0 + 1, cumTop + 1 + i, h[1], last ? S.cumVB : S.cumV);
    });
    var cumEnd = cumTop + o.hist.length + 1;
    pv.set(H0, cumEnd, '총합계', S.cumTL);
    pv.set(H0 + 1, cumEnd, o.hist.reduce(function (a, h) { return a + h[1]; }, 0), S.cumTV, 'SUM(' + col(H0 + 1) + (cumTop + 1) + ':' + col(H0 + 1) + (cumEnd - 1) + ')');

    // 피벗② (월·사업구분·채널·Royalty%) + 오른쪽 정리본
    var p2Top = Math.max(p1End, cumEnd) + 4;
    pv.set(0, p2Top, '행 레이블'); pv.set(1, p2Top, 'by'); pv.set(2, p2Top, 'by C'); pv.set(3, p2Top, 'Royalty%');
    pv.set(4, p2Top, '합계 : ' + salHead); pv.set(5, p2Top, '합계 : ' + royHead); pv.set(6, p2Top, '개수 : ' + costHead);
    var cp = 12, cpTop = p2Top - 1;
    ['행 레이블', 'by', 'by C', 'Royalty%', '합계 : ' + salHead, '개수 : ' + costHead, '합계 : ' + royHead].forEach(function (t, k) { pv.set(cp + k, cpTop, t, S.cpH); });
    var row2 = p2Top + 1, rowC = cpTop + 1;
    var cnt = function (rows) { return rows.filter(function (x) { return x[iCost] != null && x[iCost] !== ''; }).length; };
    monOrder.forEach(function (m) {
      var mRows = D.rows.filter(function (x) { return x[iMon] === m; }), mStart = rowC;
      bizOrder.forEach(function (b) {
        var bStart = rowC;
        chOrder.forEach(function (ch) {
          pctOrder.forEach(function (pc) {
            var rs = mRows.filter(function (x) { return x[iBiz] === b && x[iCh] === ch && x[iPct] === pc; }); if (!rs.length) return;
            var c = cnt(rs);
            pv.set(0, row2, m); pv.set(1, row2, b); pv.set(2, row2, ch); pv.set(3, row2, pc, S.pct2);
            pv.set(4, row2, sum(rs, iSal), S.acct); pv.set(5, row2, sum(rs, iRoy), S.acct); if (c) pv.set(6, row2, c);
            pv.set(cp + 2, rowC, ch, S.cpTxt); pv.set(cp + 3, rowC, pc, S.cpPct); pv.set(cp + 4, rowC, sum(rs, iSal), S.cpNum);
            pv.set(cp + 5, rowC, c || null, S.cpNum); pv.set(cp + 6, rowC, sum(rs, iRoy), S.cpNum);
            row2++; rowC++;
          });
        });
        if (rowC > bStart) {
          pv.set(cp + 1, bStart, N(b) === 'Partnership' ? 'Sticker' : b, S.cpBiz);
          for (var z = bStart + 1; z < rowC; z++) pv.set(cp + 1, z, null, S.cpBiz);
          if (rowC - bStart > 1) pv.merges.push(col(cp + 1) + bStart + ':' + col(cp + 1) + (rowC - 1));
        }
      });
      pv.set(0, row2, m + ' 요약'); pv.set(4, row2, sum(mRows, iSal), S.acct); pv.set(5, row2, sum(mRows, iRoy), S.acct); pv.set(6, row2, cnt(mRows)); row2++;
      pv.set(cp, mStart, m, S.cpMon); for (var z2 = mStart + 1; z2 < rowC; z2++) pv.set(cp, z2, null, S.cpMon);
      if (rowC - mStart > 1) pv.merges.push(col(cp) + mStart + ':' + col(cp) + (rowC - 1));
      var mm = MON[String(m).slice(0, 3)] || m;
      pv.set(cp, rowC, mm + '월 TTL', S.cpMonT); [1, 2, 3].forEach(function (k) { pv.set(cp + k, rowC, null, S.cpB); });
      pv.set(cp + 4, rowC, sum(mRows, iSal), S.cpBN); pv.set(cp + 5, rowC, cnt(mRows), S.cpBN); pv.set(cp + 6, rowC, sum(mRows, iRoy), S.cpBN); rowC++;
    });
    pv.set(0, row2, '총합계'); pv.set(4, row2, sum(D.rows, iSal), S.acct); pv.set(5, row2, sum(D.rows, iRoy), S.acct); pv.set(6, row2, cnt(D.rows));
    pv.set(cp, rowC, '총합계', S.cpG); [1, 2, 3].forEach(function (k) { pv.set(cp + k, rowC, null, S.cpG); });
    pv.set(cp + 4, rowC, sum(D.rows, iSal), S.cpGN); pv.set(cp + 5, rowC, cnt(D.rows), S.cpGN); pv.set(cp + 6, rowC, sum(D.rows, iRoy), S.cpGN);
    var pt2Fields = D.head.map(function (h, j) {
      if (j === iMon) return '<pivotField axis="axisRow" outline="0" showAll="0">' + itemsXml(monOrder, true) + FILL_DOWN + '</pivotField>';
      if (j === iBiz) return '<pivotField name="by" axis="axisRow" outline="0" showAll="0" defaultSubtotal="0">' + itemsXml(bizOrder, false) + FILL_DOWN + '</pivotField>';
      if (j === iCh) return '<pivotField name="by C" axis="axisRow" outline="0" showAll="0" defaultSubtotal="0">' + itemsXml(chOrder, false) + FILL_DOWN + '</pivotField>';
      if (j === iPct) return '<pivotField axis="axisRow" numFmtId="10" outline="0" showAll="0">' + itemsXml(pctOrder, true) + FILL_DOWN + '</pivotField>';
      if (j === iSal || j === iRoy) return '<pivotField dataField="1" numFmtId="41" showAll="0"/>';
      if (j === iCost) return '<pivotField dataField="1" showAll="0"/>';
      return '<pivotField showAll="0"/>';
    });
    var pt2 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="' + NS_MAIN + '" name="피벗 테이블2" cacheId="' + dCache.id + '" ' + PT_ATTR + '>' +
      '<location ref="A' + p2Top + ':G' + row2 + '" firstHeaderRow="0" firstDataRow="1" firstDataCol="4"/><pivotFields count="' + D.head.length + '">' + pt2Fields.join('') + '</pivotFields>' +
      '<rowFields count="4"><field x="' + iMon + '"/><field x="' + iBiz + '"/><field x="' + iCh + '"/><field x="' + iPct + '"/></rowFields><colFields count="1"><field x="-2"/></colFields>' +
      '<dataFields count="3"><dataField name="합계 : ' + attrName(salHead) + '" fld="' + iSal + '" baseField="0" baseItem="0"/><dataField name="합계 : ' + attrName(royHead) + '" fld="' + iRoy + '" baseField="0" baseItem="0"/>' +
      '<dataField name="개수 : ' + attrName(costHead) + '" fld="' + iCost + '" subtotal="count" baseField="0" baseItem="0"/></dataFields>' + PT_EXT + '</pivotTableDefinition>';

    // ===== PIVOT 멤버 (자재 피벗) =====
    var memVals = recs.map(function (x) { return x[22]; });
    var memOrder = uniq(memVals.filter(function (v) { return !(v && v.err); })).sort(ko);
    if (memVals.some(function (v) { return v && v.err; })) memOrder.push({ err: '#N/A' });
    var ctyOrder = uniq(recs.map(function (x) { return x[2]; })).sort(ko), chanOrder = uniq(recs.map(function (x) { return x[3]; })).sort(ko).reverse();
    var makeOrder = uniq(recs.map(function (x) { return x[6]; })).sort(ko).reverse();
    var mFields = COLS.map(function (h, j) {
      var f = { name: h, vals: recs.map(function (x) { return x[j]; }) };
      if (j === 2) f.order = ctyOrder; if (j === 3) f.order = chanOrder; if (j === 6) f.order = makeOrder; if (j === 22) f.order = memOrder;
      return f;
    });
    var mCache = addCache(cacheParts('자재', 'A1:W' + lastMat, mFields, recs.length));
    // B2B 행에 금액이 있을 때만 필터에서 뺀다 (26.2Q까지는 B2B 행 금액 0 → 예전처럼 "(모두)")
    var hasB2B = recs.some(function (x) { return x[3] === 'B2B' && Math.abs(x[21] || 0) > 0.0001; });
    var oem = recs.filter(function (x) { return x[3] !== 'B2B'; });
    var ctyShown = ctyOrder.filter(function (c) { return oem.some(function (x) { return x[2] === c; }); });
    var memShown = memOrder.filter(function (m) { return oem.some(function (x) { return (x[22] && x[22].err ? '#N/A' : x[22]) === (m && m.err ? '#N/A' : m); }); });
    var pm = new Sheet(); pm.cols = [11.3, 11.2, 10.9, null, 8.3, 10.9, null, null, 11.3, 10.9, null, null, 8.3, 10.9];
    pm.set(0, 1, '채널'); pm.set(1, 1, hasB2B ? '(다중 항목)' : '(모두)'); pm.set(0, 2, '제작방식'); pm.set(1, 2, '(모두)');
    pm.set(0, 4, '합계 : KRW'); pm.set(1, 4, '열 레이블');
    var memKey = function (m) { return m && m.err ? '#N/A' : m; };
    var cell = function (m, c) { var rs = oem.filter(function (x) { return memKey(x[22]) === memKey(m) && (c == null || x[2] === c); }); return rs.length ? rs.reduce(function (a, x) { return a + (x[21] || 0); }, 0) : null; };
    var CI = 8;
    pm.set(0, 5, '행 레이블'); pm.set(CI, 5, '행 레이블', S.cpH);
    ctyShown.forEach(function (c, k) { pm.set(1 + k, 5, c); pm.set(CI + 1 + k, 5, c, S.cpH); });
    pm.set(1 + ctyShown.length, 5, '총합계'); pm.set(CI + 1 + ctyShown.length, 5, '총합계', S.cpH);
    memShown.forEach(function (m, i) {
      var row = 6 + i;
      pm.set(0, row, memKey(m)); pm.set(CI, row, memKey(m), S.pmMem);
      ctyShown.forEach(function (c, k) { var v = cell(m, c); if (v != null) pm.set(1 + k, row, v, S.acct); pm.set(CI + 1 + k, row, v, S.cpNum); });
      pm.set(1 + ctyShown.length, row, cell(m), S.acct); pm.set(CI + 1 + ctyShown.length, row, cell(m), S.cpNum);
    });
    var pmEnd = 6 + memShown.length;
    pm.set(0, pmEnd, '총합계'); pm.set(CI, pmEnd, '총합계', S.pmT);
    ctyShown.forEach(function (c, k) { var v = oem.filter(function (x) { return x[2] === c; }).reduce(function (a, x) { return a + (x[21] || 0); }, 0); pm.set(1 + k, pmEnd, v, S.acct); pm.set(CI + 1 + k, pmEnd, v, S.pmTN); });
    var allK = oem.reduce(function (a, x) { return a + (x[21] || 0); }, 0);
    pm.set(1 + ctyShown.length, pmEnd, allK, S.acct); pm.set(CI + 1 + ctyShown.length, pmEnd, allK, S.pmTN);
    var pt3Fields = COLS.map(function (h, j) {
      if (j === 2) return '<pivotField axis="axisCol" showAll="0">' + itemsXml(ctyOrder, true) + '</pivotField>';
      if (j === 3) return '<pivotField axis="axisPage"' + (hasB2B ? ' multipleItemSelectionAllowed="1"' : '') + ' showAll="0">' + itemsXml(chanOrder, true, hasB2B ? ['B2B'] : null) + '</pivotField>';
      if (j === 6) return '<pivotField axis="axisPage" showAll="0">' + itemsXml(makeOrder, true) + '</pivotField>';
      if (j === 21) return '<pivotField dataField="1" numFmtId="41" showAll="0"/>';
      if (j === 22) return '<pivotField axis="axisRow" showAll="0">' + itemsXml(memOrder, true) + '</pivotField>';
      return '<pivotField showAll="0"/>';
    });
    var pt3 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="' + NS_MAIN + '" name="피벗 테이블6" cacheId="' + mCache.id + '" ' + PT_ATTR + '>' +
      '<location ref="A4:' + col(1 + ctyShown.length) + pmEnd + '" firstHeaderRow="1" firstDataRow="2" firstDataCol="1" rowPageCount="2" colPageCount="1"/><pivotFields count="23">' + pt3Fields.join('') + '</pivotFields>' +
      '<rowFields count="1"><field x="22"/></rowFields><colFields count="1"><field x="2"/></colFields><pageFields count="2"><pageField fld="3" hier="-1"/><pageField fld="6" hier="-1"/></pageFields>' +
      '<dataFields count="1"><dataField name="합계 : KRW" fld="21" baseField="0" baseItem="0" numFmtId="41"/></dataFields>' + PT_EXT + '</pivotTableDefinition>';

    // ===== 멤버 (자재명 / 캐릭터 / 멤버) =====
    var mb = new Sheet(); mb.cols = [56.5, 9.4, 11.3];
    mb.set(0, 1, '자재명', S.yel); mb.set(1, 1, '캐릭터', S.yel); mb.set(2, 1, '멤버', S.yel);
    var listed = {}, list = [];
    names.forEach(function (n) { if (memberOf[n]) { list.push([n, memberOf[n]]); listed[N(n)] = 1; } });
    Object.keys(map).forEach(function (k) { if (!listed[k]) list.push([k, map[k]]); });
    list.sort(function (a, b) { return ko(a[0], b[0]); });
    list.forEach(function (x, i) { mb.set(0, i + 2, x[0], S.vc); mb.set(1, i + 2, o.charOf ? o.charOf(x[0], x[1]) || null : null, S.vc); mb.set(2, i + 2, x[1], S.vc); });

    addSheet('PIVOT', pv, [{ xml: pt1, cache: dCache }, { xml: pt2, cache: dCache }]);
    addSheet('PIVOT 멤버', pm, [{ xml: pt3, cache: mCache }]);
    addSheet('멤버', mb);

    // 수식(KRW·멤버·합계)은 열 때 다시 계산
    wbXml = /<calcPr[^>]*fullCalcOnLoad/.test(wbXml) ? wbXml : wbXml.replace(/<calcPr/, '<calcPr fullCalcOnLoad="1"');
    zip.file('xl/workbook.xml', wbXml); zip.file('xl/_rels/workbook.xml.rels', wbRels); zip.file('[Content_Types].xml', ctXml); zip.file('xl/styles.xml', st.xml);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  var API = { build: build };
  if (typeof module !== 'undefined' && module.exports) module.exports = API; else root.TRUZ_REVIEW = API;
})(typeof window !== 'undefined' ? window : this);
