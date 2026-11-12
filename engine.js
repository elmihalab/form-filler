/*
 * مُعبِّئ الاستبيانات — المحرك
 * شعبة الشؤون العلمية في حلب
 *
 * هذا الملف لا يرسل أي شيء إلى الإنترنت: كل العمل يجري داخل المتصفح.
 * يعتمد على مكتبتين محفوظتين داخل مجلد lib:
 *   PizZip        لفتح ملفات الوورد والإكسل (وهي ملفات مضغوطة في حقيقتها)
 *   docxtemplater لملء الوسوم {…} داخل ملف الوورد
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(
      require('pizzip'),
      require('docxtemplater'),
      require('@xmldom/xmldom')
    );
  } else {
    root.Engine = factory(root.PizZip, root.docxtemplater, {
      DOMParser: root.DOMParser,
      XMLSerializer: root.XMLSerializer,
    });
  }
})(typeof self !== 'undefined' ? self : this, function (PizZip, Docxtemplater, xml) {
  'use strict';

  const VERSION = '1.1';
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

  // الوسوم الجاهزة التي يفهمها البرنامج دون أن تكون أعمدة في الإكسل
  const BUILTINS = {
    TODAY: 'اليوم',
    HIJRI: 'اليوم الهجري',
    SEQ: 'التسلسل',
    SURVEY: 'الاستبيان',
    ENTITY: 'الجهة',
    NAME: 'الاسم',
    REASONS: 'أسباب التوجيه',
    FIELDS: 'الحقول',
    Q: 'السؤال',
    A: 'الجواب',
  };

  /* ------------------------------------------------------------------ *
   *  أدوات نصية عامة
   * ------------------------------------------------------------------ */

  const AR_DIGITS = { '٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4', '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9',
    '۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4', '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9' };

  function toLatinDigits(s) {
    return String(s).replace(/[٠-٩۰-۹]/g, (d) => AR_DIGITS[d]);
  }

  /** توحيد النص للمقارنة: يحذف التشكيل والتطويل وعلامات الاتجاه، ويوحّد الهمزات والتاء والياء */
  function norm(s) {
    if (s === null || s === undefined) return '';
    return toLatinDigits(String(s))
      .replace(/[​-‏‪-‮⁦-⁩﻿]/g, '')
      .replace(/[ً-ٰٟـ]/g, '')
      .replace(/[أإآٱ]/g, 'ا')
      .replace(/ى/g, 'ي')
      .replace(/ة/g, 'ه')
      .replace(/[:：؟?!.،,;؛"'«»()\[\]{}\-–—_]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  }

  /** الاسم المقترح للوسم من عنوان العمود: يحذف النقطتين وعلامة الاستفهام من الطرفين */
  function defaultTag(header) {
    return String(header)
      .replace(/[​-‏‪-‮⁦-⁩﻿]/g, '')
      .replace(/[{}]/g, '')
      .replace(/^[\s:：؟?!.،,]+|[\s:：؟?!.،,]+$/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /** دالة بصمة سريعة (cyrb53) — تعطي رمزًا ثابتًا لكل نص */
  function hash(str) {
    let h1 = 0xdeadbeef, h2 = 0x41c6ce57;
    for (let i = 0; i < str.length; i++) {
      const ch = str.charCodeAt(i);
      h1 = Math.imul(h1 ^ ch, 2654435761);
      h2 = Math.imul(h2 ^ ch, 1597334677);
    }
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  function formatDate(d, withTime) {
    const s = pad2(d.getUTCDate()) + '/' + pad2(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
    if (!withTime) return s;
    return s + ' ' + pad2(d.getUTCHours()) + ':' + pad2(d.getUTCMinutes());
  }

  function formatTime(d) {
    return pad2(d.getUTCHours()) + ':' + pad2(d.getUTCMinutes());
  }

  /** تاريخ اليوم بالتقويم الهجري (أم القرى) مثل 21/05/1448 */
  function hijriToday() {
    const tries = ['en-u-ca-islamic-umalqura-nu-latn', 'en-u-ca-islamic-nu-latn'];
    for (const loc of tries) {
      try {
        const parts = new Intl.DateTimeFormat(loc, { day: 'numeric', month: 'numeric', year: 'numeric' }).formatToParts(new Date());
        const g = (t) => (parts.find((x) => x.type === t) || {}).value;
        if (g('year') && g('month') && g('day')) return pad2(+g('day')) + '/' + pad2(+g('month')) + '/' + toLatinDigits(g('year'));
      } catch (e) { /* نجرّب التالي */ }
    }
    return '';
  }

  function todayString() {
    const d = new Date();
    return pad2(d.getDate()) + '/' + pad2(d.getMonth() + 1) + '/' + d.getFullYear();
  }

  /* ------------------------------------------------------------------ *
   *  قراءة ملف الإكسل (xlsx) أو CSV
   * ------------------------------------------------------------------ */

  function parseXml(text) {
    return new xml.DOMParser().parseFromString(text, 'application/xml');
  }

  function byLocal(node, name) {
    return Array.from(node.getElementsByTagNameNS('*', name));
  }

  function childByLocal(node, name) {
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1 && n.localName === name) return n;
    }
    return null;
  }

  function decodeXlsxEscapes(s) {
    return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
  }

  function textOfSi(si) {
    // نص الخلية المشتركة: كل عناصر t ما عدا الموجودة داخل rPh (القراءة الصوتية)
    let out = '';
    byLocal(si, 't').forEach((t) => {
      let p = t.parentNode, skip = false;
      while (p && p !== si) { if (p.localName === 'rPh') { skip = true; break; } p = p.parentNode; }
      if (!skip) out += t.textContent;
    });
    return decodeXlsxEscapes(out);
  }

  const BUILTIN_DATE_FMTS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);
  const BUILTIN_TIME_FMTS = new Set([18, 19, 20, 21, 45, 46, 47]);

  function classifyFormat(id, code) {
    if (BUILTIN_TIME_FMTS.has(id)) return 'time';
    if (BUILTIN_DATE_FMTS.has(id)) return id === 22 ? 'datetime' : 'date';
    if (!code) return null;
    const c = code.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[(?!h\]|m\]|s\])[^\]]*\]/gi, '').toLowerCase();
    const hasDate = /[dy]/.test(c) || (/m/.test(c) && !/[hs]/.test(c));
    const hasTime = /[hs]/.test(c);
    if (hasDate && hasTime) return 'datetime';
    if (hasDate) return 'date';
    if (hasTime) return 'time';
    return null;
  }

  function serialToDate(serial, date1904) {
    const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
    return new Date(epoch + Math.round(serial * 86400000));
  }

  function colIndex(ref) {
    const m = /^([A-Z]+)/i.exec(ref || '');
    if (!m) return -1;
    let n = 0;
    for (const ch of m[1].toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  function cellToString(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') {
      if (Number.isInteger(v)) return String(v);
      return String(Math.round(v * 1e6) / 1e6);
    }
    if (typeof v === 'boolean') return v ? 'نعم' : 'لا';
    return String(v).replace(/\r\n?/g, '\n').trim();
  }

  function readXlsx(buffer) {
    let zip;
    try { zip = new PizZip(buffer); } catch (e) {
      throw new Error('تعذّر فتح ملف الإكسل. تأكد أنه بصيغة ‎.xlsx‎ (احفظه من إكسل بصيغة «مصنف Excel»).');
    }
    const wbFile = zip.file('xl/workbook.xml');
    if (!wbFile) throw new Error('هذا الملف ليس ملف إكسل حديثًا (‎.xlsx‎).');
    const wb = parseXml(wbFile.asText());
    const wbPr = byLocal(wb, 'workbookPr')[0];
    const date1904 = !!(wbPr && /^(1|true)$/i.test(wbPr.getAttribute('date1904') || ''));
    const sheets = byLocal(wb, 'sheet').map((s) => ({
      name: s.getAttribute('name'),
      rid: s.getAttributeNS(R, 'id') || s.getAttribute('r:id'),
    }));
    const rels = {};
    const relFile = zip.file('xl/_rels/workbook.xml.rels');
    if (relFile) {
      byLocal(parseXml(relFile.asText()), 'Relationship').forEach((r) => {
        let t = r.getAttribute('Target');
        t = t.startsWith('/') ? t.slice(1) : 'xl/' + t;
        rels[r.getAttribute('Id')] = t;
      });
    }
    // النصوص المشتركة
    const shared = [];
    const ssFile = zip.file('xl/sharedStrings.xml');
    if (ssFile) byLocal(parseXml(ssFile.asText()), 'si').forEach((si) => shared.push(textOfSi(si)));
    // تنسيقات الأرقام (لمعرفة خلايا التاريخ والوقت)
    const xfKinds = [];
    const stFile = zip.file('xl/styles.xml');
    if (stFile) {
      const st = parseXml(stFile.asText());
      const custom = {};
      byLocal(st, 'numFmt').forEach((f) => { custom[+f.getAttribute('numFmtId')] = f.getAttribute('formatCode'); });
      const cellXfs = byLocal(st, 'cellXfs')[0];
      if (cellXfs) {
        for (let n = cellXfs.firstChild; n; n = n.nextSibling) {
          if (n.nodeType === 1 && n.localName === 'xf') {
            const id = +(n.getAttribute('numFmtId') || 0);
            xfKinds.push(classifyFormat(id, custom[id]));
          }
        }
      }
    }
    const first = sheets[0];
    if (!first) throw new Error('لا توجد أوراق داخل ملف الإكسل.');
    const path = rels[first.rid] || 'xl/worksheets/sheet1.xml';
    const shFile = zip.file(path);
    if (!shFile) throw new Error('تعذّرت قراءة الورقة الأولى من ملف الإكسل.');
    const sh = parseXml(shFile.asText());
    const rows = [];
    byLocal(sh, 'row').forEach((row) => {
      const rIdx = +(row.getAttribute('r') || rows.length + 1) - 1;
      const arr = [];
      let lastCol = -1;
      for (let c = row.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1 || c.localName !== 'c') continue;
        let ci = colIndex(c.getAttribute('r'));
        if (ci < 0) ci = lastCol + 1;
        lastCol = ci;
        const t = c.getAttribute('t') || 'n';
        const vNode = childByLocal(c, 'v');
        const raw = vNode ? vNode.textContent : '';
        let val = '';
        if (t === 's') val = shared[+raw] || '';
        else if (t === 'inlineStr') { const is = childByLocal(c, 'is'); val = is ? textOfSi(is) : ''; }
        else if (t === 'str') val = decodeXlsxEscapes(raw);
        else if (t === 'b') val = raw === '1';
        else if (t === 'e') val = '';
        else if (t === 'd') { const d = new Date(raw); val = isNaN(d) ? raw : formatDate(new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes()))); }
        else if (raw !== '') {
          const num = +raw;
          const kind = xfKinds[+(c.getAttribute('s') || 0)];
          if (kind && !isNaN(num)) {
            const d = serialToDate(num, date1904);
            if (kind === 'time' || (kind !== 'date' && num < 1)) val = formatTime(d);
            else if (kind === 'datetime' && Math.abs(num % 1) > 1e-9) val = formatDate(d, true);
            else val = formatDate(d);
          } else val = num;
        }
        arr[ci] = cellToString(val);
      }
      rows[rIdx] = arr;
    });
    return { sheetName: first.name, rows: Array.from(rows, (r) => (r ? Array.from(r, (x) => x || '') : [])) };
  }

  function readCsv(text) {
    text = text.replace(/^﻿/, '');
    const delim = (text.split('\n')[0].match(/;/g) || []).length > (text.split('\n')[0].match(/,/g) || []).length ? ';' : ',';
    const rows = [];
    let row = [], cur = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; }
        else cur += ch;
      } else if (ch === '"') q = true;
      else if (ch === delim) { row.push(cur); cur = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(cur); rows.push(row); row = []; cur = '';
      } else cur += ch;
    }
    if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
    return { sheetName: 'CSV', rows: rows.map((r) => r.map(cellToString)) };
  }

  /**
   * يقرأ ملف الإكسل ويعيد: العناوين + الصفوف (كل صف كائن مفتاحه العنوان الموحَّد)
   */
  function readTable(buffer, fileName) {
    const isCsv = /\.csv$/i.test(fileName || '');
    let sheet;
    if (isCsv) {
      sheet = readCsv(new TextDecoder('utf-8').decode(new Uint8Array(buffer)));
    } else {
      const head = new Uint8Array(buffer.slice ? buffer.slice(0, 4) : buffer).subarray(0, 4);
      if (head[0] === 0xD0 && head[1] === 0xCF) {
        throw new Error('هذا ملف إكسل قديم (‎.xls‎). افتحه في إكسل ثم احفظه بصيغة ‎.xlsx‎ وأعد رفعه.');
      }
      sheet = readXlsx(buffer);
    }
    const all = sheet.rows;
    let h = all.findIndex((r) => r.filter((c) => String(c).trim() !== '').length >= 2);
    if (h < 0) throw new Error('لم أجد صف عناوين في ملف الإكسل.');
    const seen = {};
    const headers = all[h].map((c, i) => {
      let col = String(c).replace(/\s+/g, ' ').trim() || ('عمود ' + (i + 1));
      let key = norm(col) || ('col' + i);
      if (seen[key]) { seen[key]++; col += ' (' + seen[key] + ')'; key = norm(col); } else seen[key] = 1;
      return { col, key, index: i };
    });
    const rows = [];
    for (let r = h + 1; r < all.length; r++) {
      const src = all[r] || [];
      if (!src.some((c) => String(c).trim() !== '')) continue;
      const values = {};
      headers.forEach((hd) => { values[hd.key] = src[hd.index] !== undefined ? String(src[hd.index]) : ''; });
      rows.push({ seq: rows.length + 1, excelRow: r + 1, values });
    }
    return { sheetName: sheet.sheetName, headers, rows };
  }

  /* ------------------------------------------------------------------ *
   *  مطابقة الاستبيانات وبصمات الصفوف
   * ------------------------------------------------------------------ */

  function headerSimilarity(keysA, keysB) {
    const a = new Set(keysA), b = new Set(keysB);
    let inter = 0;
    a.forEach((k) => { if (b.has(k)) inter++; });
    const union = new Set([...a, ...b]).size;
    return union ? inter / union : 0;
  }

  /** بصمة الصف: تُحسب من أعمدة الاستبيان الأصلية فقط، فلا تتغير إن أضافت الوزارة عمودًا جديدًا */
  function rowFingerprint(profile, row) {
    const keys = profile.headers.map((h) => h.key);
    return hash(keys.map((k) => norm(row.values[k] || '')).join('\u0001'));
  }

  /* ------------------------------------------------------------------ *
   *  الحقول والقيم
   * ------------------------------------------------------------------ */

  function makeFields(headers, old) {
    const fields = {};
    headers.forEach((h) => {
      const prev = old && old[h.key];
      fields[h.key] = prev ? Object.assign({}, prev) : { tag: defaultTag(h.col), include: true, type: 'text' };
    });
    return fields;
  }

  function cleanItem(s) {
    return String(s)
      .replace(/^[\s\-–—•▪●○*·]+/, '')
      .replace(/^[(\[]?[0-9٠-٩]{1,2}[)\].\-–:]\s*/, '')
      .replace(/^[\s\-–—•▪●○*·]+/, '')
      .trim();
  }

  function splitList(value, mode) {
    const v = String(value || '').trim();
    if (!v) return [];
    let parts;
    if (mode === 'commas') parts = v.split(/[،,؛;\n]+/);
    else if (mode === 'numbers') {
      parts = v.split(/(?:^|\s|\n)(?=[(\[]?[0-9٠-٩]{1,2}[)\].\-–:]\s)/);
    } else parts = v.split(/\n+/);
    return parts.map(cleanItem).filter(Boolean);
  }

  function formatValue(value, field) {
    let v = String(value || '');
    if (field && field.type === 'date') {
      const m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(toLatinDigits(v));
      if (m) v = pad2(+m[3]) + '/' + pad2(+m[2]) + '/' + m[1];
      else v = v.replace(/\s+\d{1,2}:\d{2}(:\d{2})?\s*(AM|PM|ص|م)?$/i, '');
    }
    return v;
  }

  /* ------------------------------------------------------------------ *
   *  القواعد (متى يكون الطلب مكتملًا)
   * ------------------------------------------------------------------ */

  const OPS = {
    eq: { label: 'يساوي', fn: (v, x) => norm(v) === norm(x) },
    neq: { label: 'لا يساوي', fn: (v, x) => norm(v) !== norm(x) },
    contains: { label: 'يحتوي على', fn: (v, x) => norm(v).includes(norm(x)) },
    notcontains: { label: 'لا يحتوي على', fn: (v, x) => !norm(v).includes(norm(x)) },
    notempty: { label: 'غير فارغ', fn: (v) => norm(v) !== '' },
    empty: { label: 'فارغ', fn: (v) => norm(v) === '' },
  };

  function evaluateRow(profile, row) {
    const reasons = [];
    (profile.rules || []).forEach((rule) => {
      if (!rule || !rule.key || !OPS[rule.op]) return;
      const ok = OPS[rule.op].fn(row.values[rule.key] || '', rule.value || '');
      if (!ok) reasons.push(String(rule.message || '').trim() || ('استكمال: ' + rule.key));
    });
    return { complete: reasons.length === 0, reasons };
  }

  /* ------------------------------------------------------------------ *
   *  حلّ الوسوم {…}
   * ------------------------------------------------------------------ */

  function buildTagIndex(profile) {
    // كل وسم يمكن الوصول إليه بالاسم المختصر أو بعنوان العمود الكامل
    const idx = {};
    (profile.headers || []).forEach((h) => {
      const f = (profile.fields || {})[h.key] || {};
      const keys = [norm(f.tag || ''), h.key, norm(defaultTag(h.col))];
      keys.forEach((k) => { if (k && !idx[k]) idx[k] = h.key; });
    });
    return idx;
  }

  class RowCtx {
    constructor(profile, row, extra) {
      this.profile = profile;
      this.row = row;
      this.extra = extra || {};
      this.index = extra && extra.tagIndex ? extra.tagIndex : buildTagIndex(profile);
      this.maxSlot = (extra && extra.maxSlot) || {};
      const s = profile.settings || {};
      this.empty = s.emptyValue !== undefined ? s.emptyValue : '';
      this.markOn = s.markOn || '●';
      this.markOff = s.markOff !== undefined ? s.markOff : ' ';
    }

    fieldKey(name) { return this.index[norm(name)] || null; }

    fieldValue(key) {
      const field = (this.profile.fields || {})[key] || {};
      if (field.include === false) return '';
      return formatValue(this.row.values[key], field);
    }

    listItems(key) {
      const field = (this.profile.fields || {})[key] || {};
      if (field.include === false) return [];
      return splitList(this.row.values[key], field.split || 'lines');
    }

    builtin(name) {
      const n = norm(name);
      const p = this.profile;
      if (n === norm(BUILTINS.TODAY)) return todayString();
      if (n === norm(BUILTINS.HIJRI)) return hijriToday();
      if (n === norm(BUILTINS.SEQ)) return String(this.row.seq);
      if (n === norm(BUILTINS.SURVEY)) return p.name || '';
      if (n === norm(BUILTINS.ENTITY)) return (p.settings && p.settings.entity) || '';
      if (n === norm(BUILTINS.NAME)) return p.nameField ? this.fieldValue(p.nameField) : '';
      if (n === norm(BUILTINS.REASONS)) return this.extra.reasons || [];
      if (n === norm(BUILTINS.FIELDS)) return this.autoFields();
      return undefined;
    }

    autoFields() {
      const p = this.profile;
      return (p.headers || [])
        .filter((h) => (p.fields[h.key] || {}).include !== false)
        .map((h) => {
          const f = p.fields[h.key] || {};
          let v = this.fieldValue(h.key);
          if (f.type === 'list') v = this.listItems(h.key).map((x, i) => (i + 1) + '. ' + x).join('\n');
          return { [BUILTINS.Q]: f.label || defaultTag(h.col), [BUILTINS.A]: v === '' ? (this.empty || '—') : v };
        });
    }

    /** يعيد قيمة الوسم؛ isLoop = هل هو وسم شرطي/تكراري {#…} */
    resolve(tag, isLoop) {
      const raw = String(tag).trim();
      // ١) اسم عمود أو اسمه المختصر
      let key = this.fieldKey(raw);
      if (key) {
        const f = (this.profile.fields || {})[key] || {};
        if (f.type === 'list') {
          const items = this.listItems(key);
          if (isLoop) return items;
          return items.length ? items.join('\n') : this.empty;
        }
        const v = this.fieldValue(key);
        if (isLoop) return norm(v) !== '';
        return v === '' ? this.empty : v;
      }
      // ٢) وسم جاهز
      const b = this.builtin(raw);
      if (b !== undefined) {
        if (isLoop) return Array.isArray(b) ? b : norm(b) !== '';
        return Array.isArray(b) ? b.map((x) => (typeof x === 'string' ? x : '')).join('\n') : b;
      }
      // ٣) علامة اختيار:  {نوع النشاط=درس علمي}
      const eq = raw.indexOf('=');
      if (eq > 0) {
        const k = this.fieldKey(raw.slice(0, eq));
        if (k) {
          const want = norm(raw.slice(eq + 1));
          const val = this.row.values[k] || '';
          const hit = norm(val) === want || splitList(val, 'commas').some((x) => norm(x) === want);
          if (isLoop) return hit;
          return hit ? this.markOn : this.markOff;
        }
      }
      // ٤) عنصر من قائمة:  {الأهداف 1}
      const m = /^(.*?)\s*([0-9٠-٩]+)$/.exec(raw);
      if (m && m[1]) {
        const k = this.fieldKey(m[1]);
        if (k) {
          const n = +toLatinDigits(m[2]);
          const items = this.listItems(k);
          let v = items[n - 1] || '';
          // الخانة الأخيرة في القالب تجمع ما زاد عن عدد الخانات
          if (n === this.maxSlot[k] && items.length > n) v = items.slice(n - 1).join('، ');
          if (isLoop) return norm(v) !== '';
          return v === '' ? this.empty : v;
        }
      }
      return undefined;
    }
  }

  function makeParser() {
    return function parser(tag, meta) {
      const isLoop = !!(meta && meta.tag && meta.tag.module === 'loop');
      return {
        get(scope) {
          if (tag === '.') return scope;
          if (scope instanceof RowCtx) {
            const v = scope.resolve(tag, isLoop);
            return v === undefined ? (isLoop ? false : '') : v;
          }
          if (scope && typeof scope === 'object' && Object.prototype.hasOwnProperty.call(scope, tag)) return scope[tag];
          return undefined;
        },
      };
    };
  }

  /** يحدد هل الوسم معروف لهذا الاستبيان (للتنبيه قبل التوليد) */
  function describeTag(profile, tag) {
    const fake = new RowCtx(profile, { seq: 1, values: {} });
    const raw = String(tag).trim();
    if (raw === '.') return { ok: true, kind: 'item' };
    if (fake.fieldKey(raw)) return { ok: true, kind: 'field', key: fake.fieldKey(raw) };
    if (fake.builtin(raw) !== undefined || norm(raw) === norm(BUILTINS.Q) || norm(raw) === norm(BUILTINS.A)) return { ok: true, kind: 'builtin' };
    const eq = raw.indexOf('=');
    if (eq > 0 && fake.fieldKey(raw.slice(0, eq))) return { ok: true, kind: 'mark', key: fake.fieldKey(raw.slice(0, eq)) };
    const m = /^(.*?)\s*([0-9٠-٩]+)$/.exec(raw);
    if (m && m[1] && fake.fieldKey(m[1])) {
      const key = fake.fieldKey(m[1]);
      const isList = ((profile.fields || {})[key] || {}).type === 'list';
      return { ok: isList, kind: 'slot', key, n: +toLatinDigits(m[2]), warn: isList ? '' : 'نوع هذا الحقل ليس «قائمة»، اجعله قائمة ليعمل الترقيم.' };
    }
    if (m && m[1]) {
      const sg = suggest(profile, m[1]);
      if (sg) return { ok: false, kind: 'unknown', suggestion: sg + ' ' + toLatinDigits(m[2]), warn: 'واجعل نوع الحقل «قائمة»' };
    }
    return { ok: false, kind: 'unknown', suggestion: suggest(profile, raw) };
  }

  function rankFields(profile, raw) {
    const strip = (w) => w.replace(/^(ال|وال|بال)(?=..)/, '');
    const n = norm(raw);
    const out = [];
    (profile.headers || []).forEach((h) => {
      const tag = ((profile.fields || {})[h.key] || {}).tag || defaultTag(h.col);
      let best = 0;
      [norm(tag), h.key].forEach((cand) => {
        if (!cand || !n) return;
        const a = new Set(n.split(' ').map(strip)), b = new Set(cand.split(' ').map(strip));
        let inter = 0; a.forEach((w) => { if (b.has(w)) inter++; });
        const score = inter / Math.max(a.size, b.size) + (cand.includes(n) || n.includes(cand) ? 0.5 : 0);
        if (score > best) best = score;
      });
      if (best > 0) out.push({ key: h.key, tag, score: best });
    });
    return out.sort((x, y) => y.score - x.score);
  }

  function suggest(profile, raw) {
    const r = rankFields(profile, raw);
    return r.length && r[0].score >= 0.5 ? r[0].tag : null;
  }

  /**
   * يربط تلقائيًا وسوم القالب غير المعروفة بأقرب عمود في الإكسل إن كان التطابق واضحًا،
   * فيغيّر الاسم المختصر لذلك العمود إلى اسم الوسم. يعيد قائمة بما رُبط.
   */
  function autoMapTags(profile, tags) {
    const mapped = [];
    const used = new Set();
    tags.forEach((t) => { const d = describeTag(profile, t.tag); if (d.ok && d.key) used.add(d.key); });
    tags.forEach((t) => {
      if (describeTag(profile, t.tag).ok) return;
      let name = String(t.tag).trim();
      let slot = /^(.*?)\s*([0-9٠-٩]+)$/.exec(name);
      const base = slot && slot[1] ? slot[1] : name;
      const eq = base.indexOf('=');
      const field = eq > 0 ? base.slice(0, eq) : base;
      const r = rankFields(profile, field);
      if (!r.length || r[0].score < 0.9 || (r[1] && r[0].score - r[1].score < 0.2) || used.has(r[0].key)) return;
      const f = profile.fields[r[0].key];
      if (!f) return;
      f.tag = field.trim();
      if (slot && slot[1] && eq < 0) { f.type = 'list'; f.split = f.split || 'lines'; }
      used.add(r[0].key);
      mapped.push({ tag: f.tag, col: (profile.headers.find((h) => h.key === r[0].key) || {}).col });
    });
    return mapped;
  }

  /**
   * يبني حقول استبيان من وسوم قالب الوورد نفسه (للتعبئة اليدوية دون إكسل)
   */
  function fieldsFromTemplate(tags) {
    const builtinNames = new Set(Object.values(BUILTINS).map(norm));
    const order = [];
    const info = {};
    const add = (name) => {
      const key = norm(name);
      if (!key || builtinNames.has(key)) return null;
      if (!info[key]) { info[key] = { col: defaultTag(name), key, type: 'text', options: [] }; order.push(key); }
      return info[key];
    };
    tags.forEach((t) => {
      const raw = String(t.tag).trim();
      if (raw === '.') return;
      const eq = raw.indexOf('=');
      if (eq > 0) {
        const f = add(raw.slice(0, eq));
        const opt = raw.slice(eq + 1).trim();
        if (f) { f.type = 'choice'; if (opt && !f.options.includes(opt)) f.options.push(opt); }
        return;
      }
      const m = /^(.*?)\s*([0-9٠-٩]+)$/.exec(raw);
      if (m && m[1] && !builtinNames.has(norm(raw))) {
        const f = add(m[1]);
        if (f) f.type = 'list';
        return;
      }
      const f = add(raw);
      if (f && f.type === 'text' && /تاريخ/.test(raw)) f.type = 'date';
    });
    const headers = order.map((k) => ({ col: info[k].col, key: k }));
    const fields = {};
    order.forEach((k) => {
      const f = info[k];
      fields[k] = { tag: f.col, include: true, type: f.type };
      if (f.type === 'list') fields[k].split = 'lines';
      if (f.type === 'choice') fields[k].options = f.options;
    });
    return { headers, fields };
  }

  /* ------------------------------------------------------------------ *
   *  قوالب الوورد
   * ------------------------------------------------------------------ */

  const ERROR_TEXT = {
    unopened_tag: 'وسم مغلق بلا فتح — يوجد «}» دون «{» قبله',
    unclosed_tag: 'وسم مفتوح بلا إغلاق — يوجد «{» دون «}» بعده',
    unclosed_loop: 'قسم شرطي مفتوح بلا إغلاق — يوجد {#…} دون {/…}',
    unopened_loop: 'إغلاق قسم بلا فتح — يوجد {/…} دون {#…}',
    closing_tag_does_not_match_opening_tag: 'إغلاق القسم لا يطابق فتحه',
    duplicate_open_tag: 'تكرار «{{» — استعمل قوسًا واحدًا',
    duplicate_close_tag: 'تكرار «}}» — استعمل قوسًا واحدًا',
    raw_xml_tag_should_be_only_text_in_paragraph: 'وسم غير صالح في هذا الموضع',
    unbalanced_loop_tags: 'أقسام شرطية متداخلة بشكل غير صحيح',
  };

  function explainError(e) {
    const list = (e && e.properties && e.properties.errors) || [e];
    return list.slice(0, 8).map((er) => {
      const p = (er && er.properties) || {};
      const why = ERROR_TEXT[p.id] || (er && er.message) || 'خطأ غير معروف';
      const where = p.xtag || p.context || '';
      return why + (where ? '  ← «' + String(where).slice(0, 60) + '»' : '');
    });
  }

  function newDoc(buffer, profile) {
    const zip = new PizZip(buffer);
    return new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      parser: makeParser(),
      nullGetter: () => '',
    });
  }

  const TAG_PARTS = /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/;

  /** يقرأ قالب الوورد ويعيد قائمة الوسوم التي فيه، أو أخطاءه */
  function inspectTemplate(buffer) {
    let zip;
    try { zip = new PizZip(buffer); } catch (e) {
      return { ok: false, errors: ['تعذّر فتح ملف الوورد. تأكد أنه بصيغة ‎.docx‎.'], tags: [] };
    }
    if (!zip.file('word/document.xml')) return { ok: false, errors: ['هذا الملف ليس ملف وورد حديثًا (‎.docx‎).'], tags: [] };
    let doc;
    try {
      doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, parser: makeParser(), nullGetter: () => '' });
    } catch (e) {
      return { ok: false, errors: explainError(e), tags: [] };
    }
    const tags = [];
    const seen = new Set();
    Object.keys(zip.files).filter((p) => TAG_PARTS.test(p)).forEach((p) => {
      let text = '';
      try { text = doc.getFullText(p); } catch (e) { return; }
      const re = /\{([^{}]+)\}/g;
      let m;
      while ((m = re.exec(text))) {
        let t = m[1].trim();
        let kind = 'value';
        if (/^[#^]/.test(t)) { kind = 'section'; t = t.slice(1).trim(); }
        else if (/^\//.test(t)) continue;
        const k = kind + '|' + t;
        if (!seen.has(k)) { seen.add(k); tags.push({ tag: t, kind, part: p }); }
      }
    });
    return { ok: true, errors: [], tags };
  }

  function computeMaxSlots(profile, tags) {
    const max = {};
    tags.forEach((t) => {
      const d = describeTag(profile, t.tag);
      if (d.kind === 'slot') max[d.key] = Math.max(max[d.key] || 0, d.n);
    });
    return max;
  }

  /** يحفظ ملف الوورد دون «مدخلات المجلدات» داخل الملف المضغوط (بعض خدمات التحرير السحابي ترفضها) */
  function generateDocx(zip) {
    Object.keys(zip.files).forEach((n) => { if (zip.files[n].dir) delete zip.files[n]; });
    return zip.generate({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }

  /** يملأ القالب ببيانات صف واحد ويعيد ملف وورد (Uint8Array) */
  function renderOne(buffer, profile, row, extra) {
    const doc = newDoc(buffer, profile);
    const ctx = new RowCtx(profile, row, extra);
    doc.render(ctx);
    return generateDocx(doc.getZip());
  }

  /** نصّ اسم الملف من النمط، مثل:  {التسلسل} - {الاسم} */
  function fileNameFor(pattern, profile, row, extra) {
    const ctx = new RowCtx(profile, row, extra);
    let s = String(pattern || '{التسلسل} - {الاسم}').replace(/\{([^{}]+)\}/g, (_, t) => {
      const v = ctx.resolve(t, false);
      return v === undefined ? '' : String(v).split('\n')[0];
    });
    s = s.replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').replace(/^[\s.\-–]+|[\s.\-–]+$/g, '').trim();
    if (s.length > 110) s = s.slice(0, 110).trim();
    return s || ('طلب ' + row.seq);
  }

  /* ------------------------------------------------------------------ *
   *  دمج عدة ملفات وورد في ملف واحد (كل طلب في قسم مستقل يبدأ بصفحة جديدة)
   * ------------------------------------------------------------------ */

  function serialize(doc) {
    return new xml.XMLSerializer().serializeToString(doc);
  }

  function relsOf(zip, partPath) {
    const dir = partPath.replace(/[^/]+$/, '');
    const name = partPath.replace(/^.*\//, '');
    const relPath = dir + '_rels/' + name + '.rels';
    const f = zip.file(relPath);
    return { relPath, doc: f ? parseXml(f.asText()) : null };
  }

  function bodyParts(doc) {
    const body = doc.getElementsByTagNameNS(W, 'body')[0];
    const kids = [];
    let sect = null;
    for (let n = body.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      if (n.namespaceURI === W && n.localName === 'sectPr') sect = n;
      else kids.push(n);
    }
    return { body, kids, sect };
  }

  function addSectionBreak(doc, body, lastNode, sectClone, before) {
    // نجعل نهاية الطلب «فاصل مقطع — صفحة تالية» ليحتفظ كل طلب بترويسة صفحته الأولى
    Array.from(sectClone.childNodes).forEach((c) => {
      if (c.nodeType === 1 && c.localName === 'type') sectClone.removeChild(c);
    });
    if (lastNode && lastNode.namespaceURI === W && lastNode.localName === 'p') {
      let pPr = null;
      for (let c = lastNode.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 1 && c.localName === 'pPr') { pPr = c; break; }
      }
      if (!pPr) { pPr = doc.createElementNS(W, 'w:pPr'); lastNode.insertBefore(pPr, lastNode.firstChild); }
      Array.from(pPr.childNodes).forEach((c) => { if (c.nodeType === 1 && c.localName === 'sectPr') pPr.removeChild(c); });
      let change = null;
      for (let c = pPr.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 1 && c.localName === 'pPrChange') change = c;
      }
      if (change) pPr.insertBefore(sectClone, change); else pPr.appendChild(sectClone);
    } else {
      const p = doc.createElementNS(W, 'w:p');
      const pPr = doc.createElementNS(W, 'w:pPr');
      pPr.appendChild(sectClone);
      p.appendChild(pPr);
      body.insertBefore(p, before);
    }
  }

  function mergeDocs(list) {
    if (!list.length) throw new Error('لا توجد ملفات للدمج.');
    if (list.length === 1) return list[0];
    const baseZip = new PizZip(list[0]);
    const baseDoc = parseXml(baseZip.file('word/document.xml').asText());
    const base = bodyParts(baseDoc);
    const finalSect = base.sect;
    const baseRels = relsOf(baseZip, 'word/document.xml');
    const relRoot = baseRels.doc.documentElement;
    const ctDoc = parseXml(baseZip.file('[Content_Types].xml').asText());
    const numFile = baseZip.file('word/numbering.xml');
    const numDoc = numFile ? parseXml(numFile.asText()) : null;
    let nextNumId = 1;
    if (numDoc) byLocal(numDoc, 'num').forEach((n) => { nextNumId = Math.max(nextNumId, +n.getAttributeNS(W, 'numId') + 1); });
    let partCounter = 100;
    let relCounter = 1000;

    const relTarget = (relDoc, id) => {
      const r = byLocal(relDoc, 'Relationship').find((x) => x.getAttribute('Id') === id);
      return r ? r.getAttribute('Target') : null;
    };

    let prevLast = base.kids[base.kids.length - 1];
    let prevSect = finalSect ? finalSect.cloneNode(true) : null;

    for (let i = 1; i < list.length; i++) {
      const z = new PizZip(list[i]);
      const d = parseXml(z.file('word/document.xml').asText());
      const parts = bodyParts(d);
      const rels = relsOf(z, 'word/document.xml').doc;

      // ترويسات وتذييلات هذا الطلب: إن اختلفت عن الطلب الأول نضيفها نسخة مستقلة
      const mySect = parts.sect ? parts.sect.cloneNode(true) : (finalSect ? finalSect.cloneNode(true) : null);
      if (mySect && rels) {
        Array.from(mySect.childNodes).forEach((ref) => {
          if (ref.nodeType !== 1 || !/^(headerReference|footerReference)$/.test(ref.localName)) return;
          const id = ref.getAttributeNS(R, 'id');
          const target = relTarget(rels, id);
          if (!target) return;
          const path = 'word/' + target.replace(/^\/?word\//, '');
          const mine = z.file(path) ? z.file(path).asText() : null;
          const theirs = baseZip.file(path) ? baseZip.file(path).asText() : null;
          if (mine === null || mine === theirs) {
            // نفس المحتوى: نستعمل معرّف الطلب الأول لنفس الملف
            const baseRel = byLocal(baseRels.doc, 'Relationship').find((x) => x.getAttribute('Target') === target);
            if (baseRel) ref.setAttributeNS(R, 'r:id', baseRel.getAttribute('Id'));
            return;
          }
          const kind = ref.localName === 'headerReference' ? 'header' : 'footer';
          const newName = kind + (partCounter++) + '.xml';
          baseZip.file('word/' + newName, mine);
          const pr = relsOf(z, path);
          if (pr.doc) baseZip.file('word/_rels/' + newName + '.rels', serialize(pr.doc));
          const newId = 'rIdM' + (relCounter++);
          const rel = baseRels.doc.createElementNS(PKG_REL, 'Relationship');
          rel.setAttribute('Id', newId);
          rel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + kind);
          rel.setAttribute('Target', newName);
          relRoot.appendChild(rel);
          const ov = ctDoc.createElementNS(CT_NS, 'Override');
          ov.setAttribute('PartName', '/word/' + newName);
          ov.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.' + kind + '+xml');
          ctDoc.documentElement.appendChild(ov);
          ref.setAttributeNS(R, 'r:id', newId);
        });
      }

      // ترقيم مستقل لكل طلب (حتى لا تكمل القوائم المرقمة من الطلب السابق)
      const numMap = {};
      if (numDoc) {
        parts.kids.forEach((k) => byLocal(k, 'numId').forEach((nid) => {
          const old = nid.getAttributeNS(W, 'val');
          if (!old || old === '0') return;
          if (!numMap[old]) {
            const src = byLocal(numDoc, 'num').find((n) => n.getAttributeNS(W, 'numId') === old);
            if (!src) return;
            const absEl = byLocal(src, 'abstractNumId')[0];
            const absId = absEl ? absEl.getAttributeNS(W, 'val') : null;
            const abs = byLocal(numDoc, 'abstractNum').find((a) => a.getAttributeNS(W, 'abstractNumId') === absId);
            const num = numDoc.createElementNS(W, 'w:num');
            num.setAttributeNS(W, 'w:numId', String(nextNumId));
            const an = numDoc.createElementNS(W, 'w:abstractNumId');
            an.setAttributeNS(W, 'w:val', absId);
            num.appendChild(an);
            for (let lvl = 0; lvl < 9; lvl++) {
              const lv = abs ? byLocal(abs, 'lvl').find((l) => l.getAttributeNS(W, 'ilvl') === String(lvl)) : null;
              const st = lv ? byLocal(lv, 'start')[0] : null;
              const o = numDoc.createElementNS(W, 'w:lvlOverride');
              o.setAttributeNS(W, 'w:ilvl', String(lvl));
              const so = numDoc.createElementNS(W, 'w:startOverride');
              so.setAttributeNS(W, 'w:val', st ? st.getAttributeNS(W, 'val') : '1');
              o.appendChild(so);
              num.appendChild(o);
            }
            const cleanup = byLocal(numDoc, 'numIdMacAtCleanup')[0];
            if (cleanup) numDoc.documentElement.insertBefore(num, cleanup); else numDoc.documentElement.appendChild(num);
            numMap[old] = String(nextNumId++);
          }
          nid.setAttributeNS(W, 'w:val', numMap[old]);
        }));
      }

      // فاصل مقطع في آخر الطلب السابق
      if (prevSect) addSectionBreak(baseDoc, base.body, prevLast, prevSect, finalSect);

      let last = null;
      parts.kids.forEach((k) => {
        byLocal(k, 'bookmarkStart').concat(byLocal(k, 'bookmarkEnd')).forEach((b) => b.parentNode.removeChild(b));
        const imported = baseDoc.importNode(k, true);
        base.body.insertBefore(imported, finalSect);
        last = imported;
      });
      prevLast = last;
      prevSect = mySect ? baseDoc.importNode(mySect, true) : null;
    }

    // المقطع الأخير يأخذ خصائص الطلب الأخير
    if (finalSect && prevSect) {
      Array.from(prevSect.childNodes).forEach((c) => { if (c.nodeType === 1 && c.localName === 'type') prevSect.removeChild(c); });
      base.body.replaceChild(prevSect, finalSect);
    }

    // معرّفات الصور يجب أن تكون فريدة
    let pid = 1;
    byLocal(baseDoc, 'docPr').forEach((el) => el.setAttribute('id', String(pid++)));

    baseZip.file('word/document.xml', serialize(baseDoc));
    baseZip.file(baseRels.relPath, serialize(baseRels.doc));
    baseZip.file('[Content_Types].xml', serialize(ctDoc));
    if (numDoc) baseZip.file('word/numbering.xml', serialize(numDoc));
    return generateDocx(baseZip);
  }

  /* ------------------------------------------------------------------ *
   *  ملف مضغوط (ZIP) لعدة ملفات
   * ------------------------------------------------------------------ */

  function makeZip(files) {
    const zip = new PizZip();
    const used = new Set();
    files.forEach((f) => {
      const dot = f.name.lastIndexOf('.');
      const stem = dot > 0 ? f.name.slice(0, dot) : f.name;
      const ext = dot > 0 ? f.name.slice(dot) : '';
      let name = f.name;
      let n = 2;
      while (used.has(name)) name = stem + ' (' + (n++) + ')' + ext;
      used.add(name);
      zip.file(name, f.data, { binary: true });
    });
    return zip.generate({ type: 'uint8array', compression: 'DEFLATE' });
  }

  /* ------------------------------------------------------------------ *
   *  تصدير واستيراد إعدادات الاستبيان
   * ------------------------------------------------------------------ */

  function bufToB64(buf) {
    const bytes = new Uint8Array(buf);
    let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return typeof btoa === 'function' ? btoa(s) : Buffer.from(bytes).toString('base64');
  }

  function b64ToBuf(b64) {
    if (typeof atob !== 'function') return new Uint8Array(Buffer.from(b64, 'base64')).buffer;
    const s = atob(b64);
    const bytes = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) bytes[i] = s.charCodeAt(i);
    return bytes.buffer;
  }

  function exportProfile(profile) {
    const copy = JSON.parse(JSON.stringify(Object.assign({}, profile, { template: null, guideTemplate: null, lastExcel: null, processed: {}, manualRows: [] })));
    if (profile.template && profile.template.data) copy.template = { name: profile.template.name, b64: bufToB64(profile.template.data) };
    if (profile.guideTemplate && profile.guideTemplate.data) copy.guideTemplate = { name: profile.guideTemplate.name, b64: bufToB64(profile.guideTemplate.data) };
    return JSON.stringify({ format: 'mu3abbi-survey', version: VERSION, exportedAt: new Date().toISOString(), profile: copy });
  }

  function importProfile(text) {
    let obj;
    try { obj = JSON.parse(text); } catch (e) { throw new Error('ملف الاستيراد تالف أو ليس من هذا البرنامج.'); }
    if (!obj || obj.format !== 'mu3abbi-survey' || !obj.profile) throw new Error('هذا الملف ليس ملف استبيان مُصدَّرًا من مُعبِّئ الاستبيانات.');
    const p = obj.profile;
    if (p.template && p.template.b64) p.template = { name: p.template.name, data: b64ToBuf(p.template.b64) };
    if (p.guideTemplate && p.guideTemplate.b64) p.guideTemplate = { name: p.guideTemplate.name, data: b64ToBuf(p.guideTemplate.b64) };
    p.processed = p.processed || {};
    p.manualRows = p.manualRows || [];
    p.lastExcel = null;
    return p;
  }

  return {
    VERSION, BUILTINS, OPS,
    norm, defaultTag, hash, toLatinDigits,
    readTable, headerSimilarity, rowFingerprint,
    makeFields, splitList, evaluateRow, buildTagIndex,
    RowCtx, describeTag, inspectTemplate, computeMaxSlots,
    renderOne, fileNameFor, mergeDocs, makeZip,
    fieldsFromTemplate, autoMapTags, hijriToday,
    exportProfile, importProfile, bufToB64, b64ToBuf,
  };
});
