/*
 * Minimal, dependency-free .xlsx reader.
 *
 * An .xlsx file is a ZIP of XML parts. This reads only what the report needs:
 * the shared-string table and the cell values of one worksheet. It never
 * evaluates formulas, macros or external links, and it never touches the
 * network. Works in any modern browser and in Node 18+ (for the tests).
 */
(function (BGI) {
  'use strict';

  var MAX_FILE_BYTES = 60 * 1024 * 1024;      // refuse inputs larger than 60 MB
  var MAX_INFLATED_BYTES = 300 * 1024 * 1024; // zip-bomb guard across all parts read
  var MAX_ROWS = 200000;

  function XlsxError(message) {
    var e = new Error(message);
    e.name = 'XlsxError';
    return e;
  }

  // ---------- ZIP ----------

  function readZipDirectory(bytes) {
    var view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    var eocd = -1;
    var min = Math.max(0, bytes.length - 65557);
    for (var i = bytes.length - 22; i >= min; i--) {
      if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw XlsxError('This is not an Excel .xlsx file (no ZIP directory found).');
    var count = view.getUint16(eocd + 10, true);
    var offset = view.getUint32(eocd + 16, true);
    var decoder = new TextDecoder('utf-8');
    var entries = {};
    for (var n = 0; n < count; n++) {
      if (offset + 46 > bytes.length || view.getUint32(offset, true) !== 0x02014b50) {
        throw XlsxError('The Excel file is damaged (bad ZIP directory).');
      }
      var method = view.getUint16(offset + 10, true);
      var compressedSize = view.getUint32(offset + 20, true);
      var size = view.getUint32(offset + 24, true);
      var nameLen = view.getUint16(offset + 28, true);
      var extraLen = view.getUint16(offset + 30, true);
      var commentLen = view.getUint16(offset + 32, true);
      var localOffset = view.getUint32(offset + 42, true);
      var name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLen));
      entries[name] = { name: name, method: method, compressedSize: compressedSize, size: size, localOffset: localOffset };
      offset += 46 + nameLen + extraLen + commentLen;
    }
    return { bytes: bytes, view: view, entries: entries, inflated: 0 };
  }

  async function inflateRaw(data) {
    var ds = new DecompressionStream('deflate-raw');
    var writer = ds.writable.getWriter();
    writer.write(data).catch(function () {});
    writer.close().catch(function () {});
    var reader = ds.readable.getReader();
    var chunks = [];
    var total = 0;
    for (;;) {
      var r = await reader.read();
      if (r.done) break;
      total += r.value.length;
      if (total > MAX_INFLATED_BYTES) throw XlsxError('The Excel file is too large to read safely.');
      chunks.push(r.value);
    }
    var out = new Uint8Array(total);
    var pos = 0;
    for (var i = 0; i < chunks.length; i++) { out.set(chunks[i], pos); pos += chunks[i].length; }
    return out;
  }

  async function readEntry(zip, name) {
    var e = zip.entries[name];
    if (!e) return null;
    var v = zip.view;
    if (v.getUint32(e.localOffset, true) !== 0x04034b50) throw XlsxError('The Excel file is damaged (bad ZIP entry).');
    var nameLen = v.getUint16(e.localOffset + 26, true);
    var extraLen = v.getUint16(e.localOffset + 28, true);
    var start = e.localOffset + 30 + nameLen + extraLen;
    var data = zip.bytes.subarray(start, start + e.compressedSize);
    var out;
    if (e.method === 0) out = data;
    else if (e.method === 8) out = await inflateRaw(data);
    else throw XlsxError('The Excel file uses a compression method this reader does not support.');
    zip.inflated += out.length;
    if (zip.inflated > MAX_INFLATED_BYTES) throw XlsxError('The Excel file is too large to read safely.');
    return new TextDecoder('utf-8').decode(out);
  }

  // ---------- XML helpers (regex based; the parts we read are machine-written) ----------

  var ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  function unescapeXml(s) {
    if (s.indexOf('&') < 0) return s;
    return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, function (m, g) {
      if (g[0] === '#') {
        var code = g[1] === 'x' ? parseInt(g.slice(2), 16) : parseInt(g.slice(1), 10);
        return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : '';
      }
      return ENTITIES[g];
    }).replace(/_x([0-9A-Fa-f]{4})_/g, function (m, h) { return String.fromCharCode(parseInt(h, 16)); });
  }

  function attr(tagAttrs, name) {
    var m = new RegExp('(?:^|\\s)' + name + '="([^"]*)"').exec(tagAttrs);
    return m ? unescapeXml(m[1]) : null;
  }

  function textRuns(xml) {
    // Concatenate every <t> inside the fragment, skipping phonetic runs.
    var cleaned = xml.replace(/<rPh[\s\S]*?<\/rPh>/g, '');
    var out = '';
    var re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
    var m;
    while ((m = re.exec(cleaned))) out += unescapeXml(m[1]);
    return out;
  }

  function parseSharedStrings(xml) {
    var list = [];
    if (!xml) return list;
    var re = /<si>([\s\S]*?)<\/si>|<si\/>/g;
    var m;
    while ((m = re.exec(xml))) list.push(m[1] ? textRuns(m[1]) : '');
    return list;
  }

  function colIndex(ref) {
    var n = 0;
    for (var i = 0; i < ref.length; i++) {
      var c = ref.charCodeAt(i);
      if (c < 65 || c > 90) break;
      n = n * 26 + (c - 64);
    }
    return n - 1;
  }

  function parseSheet(xml, shared) {
    var rows = [];
    var rowRe = /<row\b[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g;
    var rm;
    while ((rm = rowRe.exec(xml))) {
      if (rows.length >= MAX_ROWS) throw XlsxError('The sheet has more rows than this report supports.');
      var row = [];
      var body = rm[1];
      if (body) {
        var cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
        var cm;
        while ((cm = cellRe.exec(body))) {
          var attrs = cm[1];
          var inner = cm[2];
          var ref = attr(attrs, 'r');
          var idx = ref ? colIndex(ref) : row.length;
          if (idx < 0 || idx > 16383) continue;
          var type = attr(attrs, 't') || 'n';
          var value = null;
          if (inner) {
            var vm = /<v>([\s\S]*?)<\/v>/.exec(inner);
            if (type === 'inlineStr') value = textRuns(inner);
            else if (vm) {
              var raw = unescapeXml(vm[1]);
              if (type === 's') value = shared[parseInt(raw, 10)];
              else if (type === 'b') value = raw === '1';
              else if (type === 'str' || type === 'e') value = raw;
              else { var num = Number(raw); value = isFinite(num) ? num : raw; }
              if (value === undefined) value = null;
            }
          }
          row[idx] = value;
        }
      }
      rows.push(row);
    }
    return rows;
  }

  /**
   * Read one worksheet as an array of rows (arrays of cell values).
   * preferredSheet: sheet name to use when present; otherwise the first sheet.
   * Returns { sheetName, rows, sheetNames }.
   */
  async function readWorkbook(arrayBuffer, preferredSheet) {
    if (typeof DecompressionStream === 'undefined') {
      throw XlsxError('This browser is too old to open Excel files. Use a current version of Edge or Chrome.');
    }
    var bytes = arrayBuffer instanceof Uint8Array ? arrayBuffer : new Uint8Array(arrayBuffer);
    if (bytes.length > MAX_FILE_BYTES) throw XlsxError('That file is larger than 60 MB, which is not a normal export.');
    if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
      throw XlsxError('This is not an Excel .xlsx file. Export the report from Excel as .xlsx and try again.');
    }
    var zip = readZipDirectory(bytes);
    var wbXml = await readEntry(zip, 'xl/workbook.xml');
    var relsXml = await readEntry(zip, 'xl/_rels/workbook.xml.rels');
    if (!wbXml || !relsXml) throw XlsxError('This file is a ZIP but not an Excel workbook.');

    var rels = {};
    var relRe = /<Relationship\b([^>]*?)\/?>/g;
    var m;
    while ((m = relRe.exec(relsXml))) rels[attr(m[1], 'Id')] = attr(m[1], 'Target');

    var sheets = [];
    var sheetRe = /<sheet\b([^>]*?)\/?>/g;
    while ((m = sheetRe.exec(wbXml))) {
      var target = rels[attr(m[1], 'r:id')];
      if (!target) continue;
      target = target.replace(/^\/?(xl\/)?/, '');
      sheets.push({ name: attr(m[1], 'name'), path: 'xl/' + target });
    }
    if (!sheets.length) throw XlsxError('The workbook has no worksheets.');
    var chosen = sheets[0];
    for (var i = 0; i < sheets.length; i++) if (sheets[i].name === preferredSheet) chosen = sheets[i];

    var shared = parseSharedStrings(await readEntry(zip, 'xl/sharedStrings.xml'));
    var sheetXml = await readEntry(zip, chosen.path);
    if (!sheetXml) throw XlsxError('The worksheet "' + chosen.name + '" could not be read.');
    return {
      sheetName: chosen.name,
      sheetNames: sheets.map(function (s) { return s.name; }),
      rows: parseSheet(sheetXml, shared)
    };
  }

  BGI.xlsx = { readWorkbook: readWorkbook, _parseSheet: parseSheet, _parseSharedStrings: parseSharedStrings };
})(typeof globalThis !== 'undefined' ? (globalThis.BGI = globalThis.BGI || {}) : (window.BGI = window.BGI || {}));
