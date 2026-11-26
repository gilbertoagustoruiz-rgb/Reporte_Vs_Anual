/* Reporte Vs Anual — comparativo de ventas año contra año.
 * Todo se procesa en el navegador: los archivos no salen del equipo del usuario.
 * Formatos: .xlsx, .xls, .xlsm, .xlsb (vía SheetJS) y .csv/.txt (con detección de codificación y separador).
 */
(() => {
  'use strict';

  // ---------------------------------------------------------------- utilidades
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => [...el.querySelectorAll(s)];
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const norm = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();

  const MESES = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Set', 'Oct', 'Nov', 'Dic'];
  const MESES_LARGO = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Setiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  const MONTH_WORDS = {
    ene: 1, enero: 1, jan: 1, january: 1,
    feb: 2, febrero: 2, february: 2,
    mar: 3, marzo: 3, march: 3,
    abr: 4, abril: 4, apr: 4, april: 4,
    may: 5, mayo: 5,
    jun: 6, junio: 6, june: 6,
    jul: 7, julio: 7, july: 7,
    ago: 8, agosto: 8, aug: 8, august: 8,
    set: 9, sep: 9, sept: 9, setiembre: 9, septiembre: 9, september: 9,
    oct: 10, octubre: 10, october: 10,
    nov: 11, noviembre: 11, november: 11,
    dic: 12, diciembre: 12, dec: 12, december: 12,
  };

  const fmtMoney0 = new Intl.NumberFormat('es-PE', { style: 'currency', currency: 'PEN', maximumFractionDigits: 0 });
  const fmtNum0 = new Intl.NumberFormat('es-PE', { maximumFractionDigits: 0 });
  const fmtPct = new Intl.NumberFormat('es-PE', { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const money = (v) => fmtMoney0.format(v || 0);
  const moneyShort = (v) => { const a = Math.abs(v || 0), sg = v < 0 ? '-' : ''; return a >= 1e6 ? `${sg}S/ ${(a / 1e6).toFixed(a >= 1e8 ? 0 : 1)} M` : a >= 1e4 ? `${sg}S/ ${(a / 1e3).toFixed(0)} mil` : `${sg}S/ ${fmtNum0.format(a)}`; };
  const num = (v) => fmtNum0.format(v || 0);
  const pct = (v) => (v == null || !isFinite(v) ? '—' : (v > 0 ? '+' : '') + fmtPct.format(v));
  const signMoney = (v) => (v > 0 ? '+' : '') + money(v);

  const store = {
    get(k, d) { try { const v = localStorage.getItem('rva:' + k); return v == null ? d : JSON.parse(v); } catch { return d; } },
    set(k, v) { try { localStorage.setItem('rva:' + k, JSON.stringify(v)); } catch { /* sin almacenamiento */ } },
  };

  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(() => (t.hidden = true), 2600);
  }

  // ---------------------------------------------------------------- parseo de valores
  function parseAmount(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    if (typeof v === 'boolean' || v instanceof Date) return null;
    let s = String(v).trim();
    if (!s) return null;
    let neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    s = s.replace(/s\/\.?|pen|usd|\$|\s/gi, '');
    if (s.endsWith('-')) { neg = !neg; s = s.slice(0, -1); }
    if (s.startsWith('-')) { neg = !neg; s = s.slice(1); }
    if (!/^[\d.,]+$/.test(s)) return null;
    const hasC = s.includes(','), hasD = s.includes('.');
    if (hasC && hasD) {
      if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
      else s = s.replace(/,/g, '');
    } else if (hasC) {
      s = /^\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, '') : s.replace(',', '.');
    } else if (hasD && (s.match(/\./g) || []).length > 1) {
      s = s.replace(/\./g, '');
    }
    const n = parseFloat(s);
    return isFinite(n) ? (neg ? -n : n) : null;
  }

  function fixYear(y) { y = +y; return y < 100 ? 2000 + y : y; }
  function okYM(y, m, d = null) {
    if (!(y >= 1990 && y <= 2100 && m >= 1 && m <= 12)) return null;
    d = d == null ? null : +d;
    if (d != null && !(d >= 1 && d <= new Date(Date.UTC(y, m, 0)).getUTCDate())) d = null;
    return { y, m, d };
  }

  /** Semana del año: semanas de lunes a domingo; la semana 1 es la que contiene el 1 de enero. */
  function weekOf(y, m, d) {
    const doy = Math.round((Date.UTC(y, m - 1, d) - Date.UTC(y, 0, 1)) / 864e5) + 1;
    const wd = (new Date(Date.UTC(y, 0, 1)).getUTCDay() + 6) % 7;
    return Math.floor((doy - 1 + wd) / 7) + 1;
  }

  /** 1 = coberturado, 0 = no coberturado, null = sin dato. */
  function parseCobertura(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'number') return v === 1 ? 1 : v === 0 ? 0 : null;
    const n = norm(v);
    if (!n) return null;
    if (/^(no|n|sin|0|false|falso|descubiert)/.test(n)) return 0;
    if (/^(si|s|x|1|true|verdadero|ok|cubiert|con cobertura|cobertur|activ)/.test(n)) return 1;
    return null;
  }

  function parseWeek(v) {
    if (v == null || v === '') return null;
    const m = String(v).match(/(\d{1,2})\s*$/) || String(v).match(/(\d{1,2})/);
    const n = m ? +m[1] : null;
    return n >= 1 && n <= 54 ? n : null;
  }

  function monthWord(w) { return MONTH_WORDS[norm(w).replace(/\s/g, '')] || MONTH_WORDS[norm(w).slice(0, 3)] || null; }

  /** Devuelve {y, m} desde una fecha en cualquier formato habitual, o null. */
  function parseYM(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) {
      if (isNaN(v)) return null;
      const d = new Date(v.getTime() + 12 * 3600e3); // evita desfases de zona horaria en fechas de Excel
      return okYM(d.getFullYear(), d.getMonth() + 1, d.getDate());
    }
    if (typeof v === 'number') {
      if (v >= 199001 && v <= 210012 && v % 100 >= 1 && v % 100 <= 12) return okYM(Math.floor(v / 100), v % 100); // aaaamm
      if (v >= 19900101 && v <= 21001231) return okYM(Math.floor(v / 10000), Math.floor(v / 100) % 100, v % 100); // aaaammdd
      if (v > 20000 && v < 80000) { const p = XLSX.SSF.parse_date_code(v); return p ? okYM(p.y, p.m, p.d) : null; } // serial Excel
      return null;
    }
    const s = String(v).trim();
    let m;
    if ((m = s.match(/^(\d{4})[-/.](\d{1,2})(?:[-/.](\d{1,2}))?(?:[ T].*)?$/))) return okYM(+m[1], +m[2], m[3]);
    if ((m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})(?:[ T].*)?$/))) {
      let a = +m[1], b = +m[2];
      // Perú: día/mes/año. Si el segundo número no puede ser mes, se asume mes/día/año.
      const mo = b <= 12 ? b : a;
      return okYM(fixYear(m[3]), mo, b <= 12 ? a : b);
    }
    if ((m = s.match(/^(\d{1,2})[-/.](\d{4})$/))) return okYM(+m[2], +m[1]);
    if ((m = s.match(/^(\d{4})(\d{2})(\d{2})?$/))) return okYM(+m[1], +m[2], m[3]);
    if ((m = s.match(/^([a-zA-ZáéíóúÁÉÍÓÚ.]+)[\s\-/.]*(?:de[\s]+|del[\s]+)?(\d{2,4})$/))) {
      const mo = monthWord(m[1]); return mo ? okYM(fixYear(m[2]), mo) : null;
    }
    if ((m = s.match(/^(\d{4})[\s\-/.]*([a-zA-ZáéíóúÁÉÍÓÚ.]+)$/))) {
      const mo = monthWord(m[2]); return mo ? okYM(+m[1], mo) : null;
    }
    if ((m = s.match(/^(\d{1,2})\s+de\s+([a-zA-ZáéíóúÁÉÍÓÚ]+)(?:\s+(?:de|del)\s+(\d{2,4}))?$/i))) {
      const mo = monthWord(m[2]); return mo && m[3] ? okYM(fixYear(m[3]), mo, m[1]) : null;
    }
    return null;
  }

  function parseMonth(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date || (typeof v === 'number' && v > 12)) { const r = parseYM(v); return r ? r.m : null; }
    if (typeof v === 'number') return v >= 1 && v <= 12 ? Math.round(v) : null;
    const s = String(v).trim();
    if (/^\d{1,2}$/.test(s)) { const n = +s; return n >= 1 && n <= 12 ? n : null; }
    const r = parseYM(s); if (r) return r.m;
    const m = s.match(/^(\d{1,2})\s*[-.]?\s*([a-zA-Z]+)$/); // "01 - Enero"
    if (m) return +m[1] >= 1 && +m[1] <= 12 ? +m[1] : monthWord(m[2]);
    return monthWord(s);
  }

  function parseYear(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return parseYM(v)?.y ?? null;
    const n = typeof v === 'number' ? v : parseInt(String(v).replace(/\D/g, '').slice(0, 4), 10);
    if (!isFinite(n)) return null;
    const y = n < 100 ? 2000 + n : n;
    return y >= 1990 && y <= 2100 ? y : null;
  }

  // ---------------------------------------------------------------- detección de columnas
  const ROLES = [
    { key: 'venta', label: 'Venta (S/)', required: true },
    { key: 'fecha', label: 'Fecha' },
    { key: 'anio', label: 'Año' },
    { key: 'mes', label: 'Mes' },
    { key: 'semana', label: 'Semana' },
    { key: 'cadena', label: 'Cadena' },
    { key: 'tienda', label: 'Tienda / PDV' },
    { key: 'marca', label: 'Marca' },
    { key: 'unidades', label: 'Unidades' },
    { key: 'cobertura', label: 'Cobertura' },
    { key: 'coordinador', label: 'Coordinador' },
  ];

  const RULES = {
    venta: { pats: [/^(venta|ventas|venta s|venta soles|venta neta|venta total|venta valor|sell out|sellout|venta pen|monto venta|importe venta)$/, /\bventa|sell ?out|\bsoles\b/, /\b(monto|importe|valor|total|neto|facturacion)\b/], not: /cant|unid|\bund|qty|cuota|meta|objetivo|ppto|presupuesto|precio|costo|margen|%|porc|var|cod|fecha|mes|ano/ },
    fecha: { pats: [/^(fecha|fecha venta|fecha de venta|fec venta|date|dia)$/, /\bfecha|\bdate\b|\bfec\b/], not: /registro|carga|actualiz|creacion|nacim|ingreso|alta/ },
    anio: { pats: [/^(ano|anio|year|periodo ano|ano venta)$/, /\bano\b|\banio\b|\byear\b/], not: /mes|fecha/ },
    mes: { pats: [/^(mes|month|mes venta|nombre mes|mes nombre)$/, /\bmes\b|\bmonth\b|\bperiodo\b/], not: /fecha|ano|semana/ },
    semana: { pats: [/^(semana|sem|week|wk|nro semana|n semana|semana venta)$/, /\bsemana\b|\bweek\b|\bsem\b/], not: /fecha|dia/ },
    cadena: { pats: [/^(cadena|cadena comercial|retail|cliente|razon social|chain)$/, /\bcadena|\bretail\b|\bcliente\b|\bchain\b|\bcanal\b/], not: /cod|id\b|ruc/ },
    tienda: { pats: [/^(tienda|pdv|punto de venta|local|sucursal|store|nombre tienda|nombre pdv|nombre local|desc tienda|descripcion tienda|tienda nombre)$/, /\b(tienda|pdv|punto de venta|local|sucursal|store)\b/], not: /\bcod|codigo|\bid\b|cadena/ },
    marca: { pats: [/^(marca|brand)$/, /\bmarca\b|\bbrand\b/], not: /cod/ },
    coordinador: { pats: [/^(coordinador|coordinadora|coord|supervisor|supervisora|coordinador comercial|coordinador zonal|nombre coordinador)$/, /coordin|supervis/], not: /cod|dni|telef|cel|correo|mail/ },
    cobertura: { pats: [/^(cobertura|coberturado|coberturada|cobertura pdv|estado cobertura|tipo cobertura|cobertura tienda|cob)$/, /cobert/], not: /fecha|venta|monto/ },
    unidades: { pats: [/^(unidades|cantidad|und|qty|piezas)$/, /\bcant|\bunid|\bund\b|\bqty\b|\bpiezas\b/], not: /precio|venta s/ },
  };
  const FALLBACK = { tienda: /\b(tienda|pdv|local|sucursal|store)\b/ }; // si solo hay código de tienda

  function detectHeaderRow(rows) {
    let best = 0, bestScore = -1;
    const lim = Math.min(rows.length, 30);
    for (let i = 0; i < lim; i++) {
      const r = rows[i] || [];
      let score = 0, filled = 0;
      for (const c of r) {
        if (c == null || c === '') continue;
        filled++;
        if (typeof c !== 'string') continue;
        const n = norm(c);
        for (const k in RULES) if (RULES[k].pats.some((p) => p.test(n))) { score += 2; break; }
      }
      score += Math.min(filled, 6) * 0.1;
      if (score > bestScore) { bestScore = score; best = i; }
    }
    return best;
  }

  function numericShare(rows, col) {
    let n = 0, ok = 0;
    for (let i = 0; i < rows.length && n < 200; i++) {
      const v = rows[i]?.[col];
      if (v == null || v === '') continue;
      n++; if (parseAmount(v) != null) ok++;
    }
    return n ? ok / n : 0;
  }

  function detectMapping(headers, rows) {
    const used = new Set();
    const map = {};
    const hn = headers.map(norm);
    for (const role of ['venta', 'fecha', 'anio', 'semana', 'mes', 'cobertura', 'coordinador', 'cadena', 'tienda', 'marca', 'unidades']) {
      const r = RULES[role];
      let found = -1;
      for (const p of r.pats) {
        const cands = hn.map((h, i) => (h && !used.has(i) && p.test(h) && !r.not.test(h) ? i : -1)).filter((i) => i >= 0);
        if (role === 'venta') {
          const numCand = cands.find((i) => numericShare(rows, i) > 0.6);
          if (numCand != null) { found = numCand; break; }
        } else if (cands.length) { found = cands[0]; break; }
      }
      if (found < 0 && FALLBACK[role]) found = hn.findIndex((h, i) => h && !used.has(i) && FALLBACK[role].test(h));
      if (found >= 0) { map[role] = found; used.add(found); }
    }
    // Si "mes" tiene fechas completas y no hay "fecha", úsalo como fecha.
    if (map.fecha == null && map.mes != null && map.anio == null) {
      const sample = rows.slice(0, 50).map((r) => r?.[map.mes]).filter((v) => v != null && v !== '');
      if (sample.length && sample.filter((v) => parseYM(v)).length / sample.length > 0.6) { map.fecha = map.mes; delete map.mes; }
    }
    return map;
  }

  /** Columnas cuyo encabezado es un mes con año (formato ancho: Ene-25, 2025-01, fecha…). */
  function detectWideColumns(headers, mapping) {
    const used = new Set(Object.values(mapping));
    const cols = [];
    headers.forEach((h, i) => {
      if (used.has(i) || h == null || h === '') return;
      const ym = parseYM(h);
      if (ym) cols.push({ i, ...ym });
    });
    return cols;
  }

  // ---------------------------------------------------------------- lectura de archivos
  const state = {
    files: [],
    records: [], // ventas consolidadas: {cadena, tienda, marca, y, m, v, u}
    years: [],
    hasBrand: false,
    hasStore: false,
    sort: { key: 'b', dir: -1 },
    charts: {},
    view: 'datos',
  };
  let fileSeq = 0;

  async function readFile(file) {
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    const buf = await file.arrayBuffer();
    let wb;
    if (ext === 'csv' || ext === 'txt') {
      let text;
      try { text = new TextDecoder('utf-8', { fatal: true }).decode(buf); } catch { text = new TextDecoder('windows-1252').decode(buf); }
      text = text.replace(/^﻿/, '');
      const first = text.split(/\r?\n/).find((l) => l.trim()) || '';
      const counts = { ';': (first.match(/;/g) || []).length, ',': (first.match(/,/g) || []).length, '\t': (first.match(/\t/g) || []).length, '|': (first.match(/\|/g) || []).length };
      const FS = Object.entries(counts).sort((a, b) => b[1] - a[1])[0][0];
      wb = XLSX.read(text, { type: 'string', raw: true, FS });
    } else {
      wb = XLSX.read(buf, { type: 'array', cellDates: true, dense: true });
    }
    return wb;
  }

  function sheetRows(wb, name) {
    const ws = wb.Sheets[name];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
  }

  const ALL_SHEETS = '__all__';

  /** Prepara una hoja: encabezados, filas y columnas detectadas. */
  function preparePart(f, sheet, headerRowOpt, mappingOpt) {
    const all = sheetRows(f.wb, sheet);
    const headerRow = headerRowOpt ?? detectHeaderRow(all);
    const rawHeaders = all[headerRow] || [];
    const width = all.reduce((w, r) => Math.max(w, r ? r.length : 0), 0);
    const headers = Array.from({ length: width }, (_, i) => rawHeaders[i]);
    const headerText = headers.map((h, i) => (h == null || h === '' ? `Columna ${XLSX.utils.encode_col(i)}` : h instanceof Date ? `${MESES[h.getMonth()]}-${h.getFullYear()}` : String(h).trim()));
    const rows = all.slice(headerRow + 1);
    const sig = 'map:' + headerText.map(norm).join('|');
    const mapping = mappingOpt ?? store.get(sig, null) ?? detectMapping(headerText, rows);
    // Si la hoja se llama como un mes (ENERO, Feb-26…), ese mes sirve cuando la fila no trae fecha.
    const sheetYM = parseYM(sheet);
    const sheetMonth = sheetYM ? sheetYM.m : monthWord(String(sheet).trim());
    return { sheet, headerRow, headers, headerText, rows, mapping, sig, totalRows: all.length, sheetMonth, sheetYear: sheetYM ? sheetYM.y : null };
  }

  /** Hojas con datos de venta utilizables. */
  function usableSheets(f) {
    const out = [];
    for (const n of f.wb.SheetNames) {
      const p = preparePart(f, n);
      if (!p.rows.length) continue;
      const m = p.mapping;
      const hasDate = m.fecha != null || m.anio != null || m.mes != null || m.semana != null || p.sheetMonth;
      const wide = detectWideColumns(p.headers, m).length >= 2;
      if ((m.venta != null && hasDate) || wide) out.push(n);
    }
    return out;
  }

  function prepareFile(f, opts = {}) {
    f.fileYear = parseYear((f.name.match(/(19|20)\d{2}/) || [])[0]);
    if (opts.sheet !== undefined) f.sheetMode = opts.sheet;
    if (f.sheetMode === undefined) {
      const ok = usableSheets(f);
      f.usable = ok;
      f.sheetMode = ok.length >= 2 ? ALL_SHEETS : (ok[0] ?? f.wb.SheetNames[0]);
    }
    if (!f.usable) f.usable = usableSheets(f);
    const names = f.sheetMode === ALL_SHEETS ? f.usable : [f.sheetMode];
    const single = names.length === 1;
    f.parts = names.map((n) => preparePart(f, n, single ? opts.headerRow : undefined, single ? opts.mapping : undefined));
    const first = f.parts[0] || preparePart(f, f.wb.SheetNames[0]);
    Object.assign(f, { sheet: first.sheet, headerRow: first.headerRow, headers: first.headers, headerText: first.headerText, mapping: first.mapping, sig: first.sig, totalRows: first.totalRows });
    buildFileAgg(f);
  }

  function buildFileAgg(f) {
    const agg = new Map();
    const warnings = [];
    let used = 0, noDate = 0, noAmt = 0;
    const years = new Set();
    let total = 0;
    let maxStamp = 0, maxDate = null;
    const maxByYear = {};
    const monthsByYear = {};
    const partErrors = [];
    let anyWide = false;
    let covTotal = 0, uncovTotal = 0, unkTotal = 0;
    const unknownLabels = new Set();
    f.error = null;

    // Nombres unificados en mayúsculas y sin espacios dobles ("Tottus" y "TOTTUS" son la misma cadena).
    const label = (v, d) => { const s = v == null ? '' : v instanceof Date ? v.toLocaleDateString('es-PE') : String(v).replace(/\s+/g, ' ').trim().toUpperCase(); return s || d; };
    const dagg = new Map(), tdagg = new Map();
    const add = (cad, tie, mar, y, m, w, v, u, d = null, cob = null, coord = '') => {
      const tk = cad + '\u0001' + tie + '\u0001' + coord + '\u0001' + mar + '\u0001' + (cob == null ? '' : cob) + '\u0001' + y + '\u0001' + m + '\u0001' + (w || 0) + '\u0001' + (d || 0);
      const tc = tdagg.get(tk);
      if (tc) { tc.v += v; tc.u += u; } else tdagg.set(tk, { v, u });
      const dk = cad + '\u0001' + mar + '\u0001' + (cob == null ? '' : cob) + '\u0001' + y + '\u0001' + m + '\u0001' + (w || 0) + '\u0001' + (d || 0);
      const dc = dagg.get(dk);
      if (dc) { dc.v += v; dc.u += u; } else dagg.set(dk, { v, u });
      const key = cad + '\u0001' + tie + '\u0001' + mar + '\u0001' + y + '\u0001' + m + '\u0001' + (w || 0) + '\u0001' + (cob == null ? '' : cob);
      if (cob === 1) covTotal += v; else if (cob === 0) uncovTotal += v; else unkTotal += v;
      const cur = agg.get(key);
      if (cur) { cur.v += v; cur.u += u; } else agg.set(key, { v, u });
      years.add(y); total += v; used++;
      if (v) {
        (monthsByYear[y] || (monthsByYear[y] = new Set())).add(m);
        const stamp = y * 10000 + m * 100 + (d || 0);
        if (stamp > maxStamp) { maxStamp = stamp; maxDate = { y, m, d, w: w || 0 }; }
        if (!maxByYear[y] || stamp > maxByYear[y].stamp) maxByYear[y] = { stamp, m, d };
      }
    };
    const isTotal = (cad, tie) => /^(total|sub ?total|gran total)\b/.test(norm(tie)) || /^(total|sub ?total|gran total)\b/.test(norm(cad));

    for (const part of f.parts) {
      const { rows, mapping: mp } = part;
      const wide = mp.venta == null ? detectWideColumns(part.headers, mp) : [];
      const isWide = wide.length >= 2;
      anyWide = anyWide || isWide;
      const fallbackYear = part.sheetYear || f.fileYear;
      if (mp.venta == null && !isWide) { partErrors.push(`${part.sheet}: no encontré la columna de venta.`); continue; }
      if (!isWide && mp.fecha == null && mp.mes == null && mp.anio == null && mp.semana == null && !part.sheetMonth) { partErrors.push(`${part.sheet}: no encontré fecha, año ni mes.`); continue; }
      const val = (r, k) => (mp[k] == null ? null : r[mp[k]]);
      const cobOf = (r) => {
        if (mp.cobertura == null) return null;
        const raw = val(r, 'cobertura'), c = parseCobertura(raw);
        if (c == null && raw != null && raw !== '' && unknownLabels.size < 6) unknownLabels.add(String(raw).trim());
        return c;
      };
      for (const r of rows) {
        if (!r) continue;
        const cad = label(val(r, 'cadena'), '(Sin cadena)');
        const tie = label(val(r, 'tienda'), '(Sin tienda)');
        const mar = label(val(r, 'marca'), '');
        if (isWide) {
          if (mp.tienda != null && val(r, 'tienda') == null && mp.cadena != null && val(r, 'cadena') == null) continue;
          if (isTotal(cad, tie)) continue;
          for (const c of wide) {
            const v = parseAmount(r[c.i]);
            if (v == null || v === 0) continue;
            add(cad, tie, mar, c.y, c.m, 0, v, 0, null, cobOf(r), label(val(r, 'coordinador'), ''));
          }
          continue;
        }
        const v = parseAmount(val(r, 'venta'));
        if (v == null) { if (r.some((c) => c != null && c !== '')) noAmt++; continue; }
        let ym = null;
        if (mp.fecha != null) ym = parseYM(val(r, 'fecha'));
        if (!ym && mp.mes != null) {
          const m = parseMonth(val(r, 'mes'));
          const y = (mp.anio != null ? parseYear(val(r, 'anio')) : null) || parseYM(val(r, 'mes'))?.y || fallbackYear;
          if (y && m) ym = { y, m, d: null };
        }
        if (!ym && mp.semana != null) {
          const y = (mp.anio != null ? parseYear(val(r, 'anio')) : null) || fallbackYear, w = parseWeek(val(r, 'semana'));
          if (y && w) { const wd = (new Date(Date.UTC(y, 0, 1)).getUTCDay() + 6) % 7; const dt = new Date(Date.UTC(y, 0, 1 + (w - 1) * 7 - wd + 3)); ym = { y, m: dt.getUTCFullYear() === y ? dt.getUTCMonth() + 1 : (w < 10 ? 1 : 12), d: null }; }
        }
        if (!ym && part.sheetMonth) {
          const y = (mp.anio != null ? parseYear(val(r, 'anio')) : null) || fallbackYear;
          if (y) ym = { y, m: part.sheetMonth, d: null };
        }
        if (!ym) { noDate++; continue; }
        if (isTotal(cad, tie)) continue;
        const u = mp.unidades != null ? parseAmount(val(r, 'unidades')) || 0 : 0;
        let w = ym.d ? weekOf(ym.y, ym.m, ym.d) : 0;
        if (mp.semana != null) w = parseWeek(val(r, 'semana')) || w;
        add(cad, tie, mar, ym.y, ym.m, w, v, u, ym.d, cobOf(r), label(val(r, 'coordinador'), ''));
      }
    }
    if (!used) f.error = partErrors[0] || 'No pude leer ninguna fila con venta y fecha. Revisa la fila de encabezados y las columnas elegidas.';
    else partErrors.forEach((e) => warnings.push('Hoja omitida · ' + e));
    if (noDate) warnings.push(`${num(noDate)} filas sin fecha válida se omitieron.`);
    if (noAmt > 0 && noAmt > used * 0.05) warnings.push(`${num(noAmt)} filas sin monto de venta se omitieron.`);
    if (used && !anyWide && f.mapping.tienda == null) warnings.push('No hay columna de tienda: el detalle por tienda saldrá agrupado.');
    // Aviso clave: un año con un solo mes suele indicar que faltan hojas o que la columna de fecha no es la correcta.
    for (const y in monthsByYear) {
      if (monthsByYear[y].size === 1) {
        const m = [...monthsByYear[y]][0];
        warnings.push(`⚠ El año ${y} solo tiene datos de ${MESES_LARGO[m - 1].toLowerCase()}. Si el archivo trae más meses, revisa la hoja elegida o la columna de fecha/mes.`);
      }
    }
    const hasCob = f.parts.some((p) => p.mapping.cobertura != null);
    if (used && !hasCob) warnings.push('⚠ No encontré la columna de cobertura: estos datos no aparecerán cuando el reporte muestre "Solo coberturados". Elígela en la lista "Cobertura".');
    if (unknownLabels.size) warnings.push(`Valores de cobertura que no reconocí (cuentan como "sin dato"): ${[...unknownLabels].join(', ')}.`);
    if (used && !f.parts.some((p) => p.mapping.coordinador != null)) warnings.push('No encontré la columna de coordinador: la comparativa por coordinador saldrá vacía para este archivo. Elígela en la lista "Coordinador".');
    Object.assign(f, { agg, dagg, tdagg, warnings, used, years: [...years].sort(), total, maxDate, maxByYear, monthsByYear, wide: anyWide, hasCob, covTotal, uncovTotal, unkTotal });
  }

  function monthsText(set) {
    const ms = [...set].sort((a, b) => a - b);
    if (!ms.length) return '';
    const contiguous = ms[ms.length - 1] - ms[0] + 1 === ms.length;
    return contiguous ? (ms.length === 1 ? MESES[ms[0] - 1] : `${MESES[ms[0] - 1]}–${MESES[ms[ms.length - 1] - 1]}`) : ms.map((m) => MESES[m - 1]).join(', ');
  }

  // ---------------------------------------------------------------- servidor (base de datos)
  const API = new URLSearchParams(location.search).get('api')
    || (location.hostname.endsWith('onrender.com') ? 'https://reporte-vs-anual-api.onrender.com' : 'http://localhost:3000');
  const session = { token: store.get('token', null), user: store.get('user', null) };

  async function api(path, opts = {}) {
    const headers = { ...(opts.body ? { 'Content-Type': 'application/json' } : {}), ...(session.token ? { Authorization: 'Bearer ' + session.token } : {}) };
    let res;
    try { res = await fetch(API + path, { ...opts, headers, body: opts.body ? JSON.stringify(opts.body) : undefined }); }
    catch { throw new Error('No se pudo conectar con el servidor.'); }
    const data = await res.json().catch(() => ({}));
    if (res.status === 401 && session.token && path !== '/api/login') { setSession(null); }
    if (!res.ok) throw new Error(data.error || `Error ${res.status}`);
    return data;
  }

  function setSession(s) {
    const was = session.token;
    session.token = s ? s.token : null;
    session.user = s ? s.user : null;
    store.set('token', session.token); store.set('user', session.user);
    if (!s && was) location.reload(); // borra de la pantalla los datos del usuario anterior
    renderSession();
  }

  const perm = (k) => !!session.user?.perms?.[k];
  const canDatos = () => perm('subir') || perm('eliminar') || perm('gestionar_coord');

  function renderSession() {
    const u = session.user;
    // Sin sesión no se ve nada del reporte: solo la pantalla de ingreso.
    document.body.classList.toggle('locked', !u);
    $('#gate').hidden = !!u;
    $('#userBox').innerHTML = u
      ? `<span class="who">👤 ${esc(u.nombre)}</span><span class="perfil">${esc(u.perfil || '')}</span><button class="btn btn-ghost" id="btnPwd">Contraseña</button><button class="btn btn-ghost" id="btnLogout">Salir</button>`
      : '';
    $('.tab[data-view="datos"]').hidden = !canDatos();
    $('.tab[data-view="admin"]').hidden = !perm('admin_usuarios');
    $('#dlGroup').hidden = !perm('descargar');
    $('#btnExport').hidden = !perm('descargar');
    $('#uploadArea').hidden = !perm('subir');
    $('#loginHint').hidden = true;
    renderCoord();
    renderBase();
    if (u && ((state.view === 'datos' && !canDatos()) || (state.view === 'admin' && !perm('admin_usuarios')))) setView('resumen');
  }

  const ACC_LABEL = { coordinador: 'Coordinador', cadena: 'Cadena', tienda: 'Tienda' };
  const accText = (a) => (a.tipo === 'coordinador' ? `tiendas de ${a.valor}` : a.tipo === 'cadena' ? `toda la cadena ${a.valor}` : `tienda ${a.valor.split('|')[1]} (${a.valor.split('|')[0]})`);
  const accesosText = (list) => { const t = list.map(accText); return t.length <= 1 ? t.join('') : t.slice(0, -1).join(', ') + ' y ' + t[t.length - 1]; };

  function renderScope() {
    const bar = $('#scopeBar');
    const u = session.user;
    if (u && !perm('ver_todo')) {
      bar.hidden = false;
      const acc = u.accesos || [];
      bar.innerHTML = acc.length
        ? `Estás viendo: <b>${esc(accesosText(acc))}</b>.`
        : 'Tu usuario todavía no tiene accesos asignados, así que no hay tiendas para mostrar. Pide a un administrador que te los asigne.';
    } else bar.hidden = true;
  }

  let serverData = null, coordData = null;

  function renderCoord() {
    const box = $('#coordBox');
    if (!coordData) return;
    box.hidden = !coordData.coordinadores.length && !coordData.reemplazos.length;
    $('#tblCoord').innerHTML = `<thead><tr><th>Coordinador</th><th class="num">Tiendas</th><th class="num">Venta</th></tr></thead><tbody>${
      coordData.coordinadores.map((c) => `<tr><td>${esc(c.coordinador)}</td><td class="num">${num(c.tiendas)}</td><td class="num">${money(c.venta)}</td></tr>`).join('') || '<tr><td colspan="3" class="muted">Aún no hay coordinadores guardados: vuelve a subir los archivos.</td></tr>'}</tbody>`;
    const logged = perm('gestionar_coord');
    $('#tblRepl').innerHTML = `<thead><tr><th>Nombre en el archivo</th><th>Se muestra como</th>${logged ? '<th></th>' : ''}</tr></thead><tbody>${
      coordData.reemplazos.map((r) => `<tr><td>${esc(r.original)}</td><td><b>${esc(r.nuevo)}</b></td>${logged ? `<td><button class="btn btn-danger" data-del-repl="${esc(r.original)}">Quitar</button></td>` : ''}</tr>`).join('') || '<tr><td colspan="3" class="muted">Sin reemplazos.</td></tr>'}</tbody>`;
    $('#formRepl').hidden = !logged;
    $('#coordList').innerHTML = coordData.coordinadores.map((c) => `<option value="${esc(c.coordinador)}">`).join('');
  }
  async function loadServer() {
    if (!session.token) return;
    const st = $('#serverStatus');
    st.hidden = false; st.className = 'server-status';
    st.innerHTML = '<span class="spinner"></span> Cargando los datos de la base… (si el servidor estaba en reposo puede tardar hasta un minuto)';
    try {
      serverData = await api('/api/data');
      cmp.data = null; cmp.cache = null;
      const firstLoad = !state.loadedOnce; state.loadedOnce = true;
      api('/api/coordinadores').then((r) => { coordData = r; renderCoord(); fillCoordList(); }).catch(() => {});
      applyServerData(serverData);
      renderScope();
      st.hidden = true;
      if (firstLoad && (state.records.length || !canDatos()) && state.view === 'datos') setView('resumen');
    } catch (e) {
      st.className = 'server-status err';
      st.innerHTML = `${esc(e.message)} <button class="btn" id="btnRetry">Reintentar</button>`;
      $('#btnRetry').onclick = loadServer;
    }
  }

  function applyServerData(D) {
    const { dict } = D;
    const records = D.rows.map(([ci, ti, mi, y, m, w, v, u, cob]) => ({ cadena: dict.cadenas[ci], tienda: dict.tiendas[ti], marca: dict.marcas[mi], tipo: tipoDe(dict.marcas[mi]), y, m, w, v, u, cob: cob ?? null }));
    const years = [...new Set(records.filter((r) => r.v).map((r) => r.y))].sort((a, b) => a - b);
    const weekYears = new Set(records.filter((r) => r.w > 0 && r.v).map((r) => r.y));
    let cut = null;
    const lastInfo = D.anios.filter((a) => years.includes(a.anio)).pop();
    if (lastInfo) {
      const y = lastInfo.anio;
      // Corte de evaluación: siempre el último año, mes y semana con venta en la base.
      cut = {
        y, d: lastInfo.max_dia,
        m: Math.max(...records.filter((r) => r.y === y && r.v).map((r) => r.m)),
        w: Math.max(0, ...records.filter((r) => r.y === y && r.w > 0 && r.v).map((r) => r.w)),
      };
      if (cut.m !== lastInfo.max_mes) cut.d = null;
    }
    Object.assign(state, { records, years, cut, weekYears, hasBrand: records.some((r) => r.marca), hasCob: records.some((r) => r.cob != null) });
    syncFilterOptions();
    renderFiles();
    renderBase();
    renderAll();
  }

  function renderBase() {
    const D = serverData;
    if (!D) return;
    const logged = perm('eliminar');
    const monthsOf = new Map();
    for (const r of state.records) if (r.v) (monthsOf.get(r.y) || monthsOf.set(r.y, new Set()).get(r.y)).add(r.m);
    $('#baseYears').innerHTML = D.anios.length ? [...D.anios].reverse().map((a) => {
      const ms = monthsOf.get(a.anio) || new Set();
      const lastDay = new Date(Date.UTC(a.anio, a.max_mes, 0)).getUTCDate();
      const strip = MESES.map((m, i) => {
        const on = ms.has(i + 1), part = on && i + 1 === a.max_mes && a.max_dia && a.max_dia < lastDay;
        return `<span class="${part ? 'part' : on ? 'on' : ''}" title="${MESES_LARGO[i]}${on ? (part ? `: hasta el día ${a.max_dia}` : ': cargado') : ': sin datos'}">${m[0]}</span>`;
      }).join('');
      const tags = [
        ms.size === 12 ? '<span class="yc-tag">12 meses</span>' : `<span class="yc-tag ${ms.size <= 1 ? 'warn' : ''}">${ms.size} ${ms.size === 1 ? 'mes' : 'meses'}</span>`,
        a.con_cobertura ? `<span class="yc-tag">${a.total ? fmtPct.format(a.cubierto / a.total) : '—'} coberturado</span>` : '<span class="yc-tag warn">Sin cobertura</span>',
      ].join('');
      return `<article class="yc">
        <div class="yc-top"><span class="yc-year">${a.anio}</span><span class="yc-total">${moneyShort(a.total)}</span></div>
        <div class="yc-months" aria-label="Meses cargados de ${a.anio}">${strip}</div>
        <div class="yc-tags">${tags}</div>
        <div class="yc-meta">Último dato: <b>${a.max_dia ? `${a.max_dia} de ${MESES_LARGO[a.max_mes - 1].toLowerCase()}` : MESES_LARGO[a.max_mes - 1]}</b><br>Subido por <b>${esc(a.usuario || '—')}</b> · ${new Date(a.actualizado).toLocaleDateString('es-PE', { day: 'numeric', month: 'short' })}</div>
        ${logged ? `<button class="btn btn-danger yc-del" data-del-year="${a.anio}">Eliminar ${a.anio}</button>` : ''}
      </article>`;
    }).join('') : '<p class="muted">La base todavía está vacía. Inicia sesión y sube tus archivos.</p>';
    const rows = D.anios.map((a) => `<tr>
      <td><b>${a.anio}</b></td>
      <td>${a.max_dia ? `${String(a.max_dia).padStart(2, '0')}/${String(a.max_mes).padStart(2, '0')}/${a.anio}` : `${MESES_LARGO[a.max_mes - 1]} ${a.anio}`}</td>
      <td class="num">${num(a.filas)}</td>
      <td class="num">${money(a.total)}</td>
      <td class="num">${a.con_cobertura ? `${money(a.cubierto)} <span class="muted small">(${a.total ? fmtPct.format(a.cubierto / a.total) : '—'})</span>` : '<span class="down small">Sin columna de cobertura</span>'}</td>
      <td>${esc(a.usuario || '—')}</td>
      <td>${new Date(a.actualizado).toLocaleString('es-PE', { dateStyle: 'short', timeStyle: 'short' })}</td>
      <td class="muted small">${esc((a.archivos || []).join(', '))}</td>
      ${logged ? `<td><button class="btn btn-danger" data-del-year="${a.anio}">Eliminar</button></td>` : ''}
    </tr>`).join('');
    $('#tblBase').innerHTML = `<thead><tr><th>Año</th><th>Último dato</th><th class="num">Registros</th><th class="num">Venta</th><th class="num">Venta coberturada</th><th>Subido por</th><th>Actualizado</th><th>Archivos</th>${logged ? '<th></th>' : ''}</tr></thead>
      <tbody>${rows || `<tr><td colspan="9" class="muted">La base todavía está vacía.</td></tr>`}</tbody>`;
    const last = state.years[state.years.length - 1];
    const bm = new Map();
    for (const r of state.records) {
      if (!r.marca) continue;
      const o = bm.get(r.marca) || { total: 0, last: 0 };
      o.total += r.v; if (r.y === last) o.last += r.v;
      bm.set(r.marca, o);
    }
    const marcas = [...bm.entries()].sort((a, b) => (tipoDe(a[0]) === 'paraiso' ? 0 : 1) - (tipoDe(b[0]) === 'paraiso' ? 0 : 1) || b[1].total - a[1].total);
    $('#brandsBox').hidden = !marcas.length;
    $('#tblMarcas').innerHTML = `<thead><tr><th>Marca</th><th>Grupo</th><th class="num">Venta total</th><th class="num">Venta ${last ?? ''}</th></tr></thead><tbody>${
      marcas.map(([m, o]) => { const t = tipoDe(m); return `<tr><td>${esc(m)}</td><td><span class="pill ${t === 'paraiso' ? 'crece' : 'igual'}">${TIPO_LABEL[t]}</span></td><td class="num">${money(o.total)}</td><td class="num">${money(o.last)}</td></tr>`; }).join('')}</tbody>`;
    const noBrand = state.records.filter((r) => !r.marca && r.v).length;
    $('#brandsNote').textContent = `Grupo Paraíso: ${MARCAS_PARAISO.join(', ')}. Todas las demás marcas cuentan como venta de la competencia.` + (noBrand ? ` Hay ${num(noBrand)} registros sin marca: no entran en ninguno de los dos grupos.` : '');
    $('#tblCargas').innerHTML = `<thead><tr><th>Fecha</th><th>Usuario</th><th>Años</th><th class="num">Registros</th><th class="num">Venta</th><th>Archivos</th></tr></thead><tbody>${
      D.cargas.map((c) => `<tr><td>${new Date(c.creado_en).toLocaleString('es-PE', { dateStyle: 'short', timeStyle: 'short' })}</td><td>${esc(c.usuario || '—')}</td><td>${c.anios.join(', ')}</td><td class="num">${num(c.filas)}</td><td class="num">${money(c.total)}</td><td class="muted small">${esc(c.archivos.join(', '))}</td></tr>`).join('') || '<tr><td colspan="6" class="muted">Sin cargas.</td></tr>'}</tbody>`;
  }

  // Archivos por guardar (aún no están en la base).
  function stagingPayload() {
    const ok = state.files.filter((f) => !f.error);
    const agg = new Map(), dagg = new Map(), tdagg = new Map();
    const maxY = {};
    const merge = (dst, src) => { for (const [k, o] of src) { const c = dst.get(k); if (c) { c.v += o.v; c.u += o.u; } else dst.set(k, { v: o.v, u: o.u }); } };
    for (const f of ok) {
      merge(agg, f.agg); merge(dagg, f.dagg || new Map()); merge(tdagg, f.tdagg || new Map());
      for (const y in f.maxByYear) { const a = f.maxByYear[y]; if (!maxY[y] || a.stamp > maxY[y].stamp) maxY[y] = a; }
    }
    const byYear = {};
    const Y = (y) => byYear[y] || (byYear[y] = { rows: [], dias: [], tdias: [] });
    let total = 0, nRows = 0;
    for (const [k, o] of agg) {
      const [c, t, m, y, mo, w, cob] = k.split('\u0001');
      Y(y).rows.push([c, t, m, +y, +mo, +w, Math.round(o.v * 100) / 100, Math.round(o.u * 100) / 100, cob === '' ? null : +cob]);
      total += o.v; nRows++;
    }
    for (const [k, o] of dagg) {
      const [c, m, cob, y, mo, w, d] = k.split('\u0001');
      Y(y).dias.push([c, m, cob === '' ? null : +cob, +y, +mo, +w, +d, Math.round(o.v * 100) / 100, Math.round(o.u * 100) / 100]);
    }
    for (const [k, o] of tdagg) {
      const [c, t, co, m, cob, y, mo, w, d] = k.split('\u0001');
      Y(y).tdias.push([c, t, co, m, cob === '' ? null : +cob, +y, +mo, +w, +d, Math.round(o.v * 100) / 100, Math.round(o.u * 100) / 100]);
    }
    const anios = Object.keys(maxY).map((y) => ({ anio: +y, max_mes: maxY[y].m, max_dia: maxY[y].d })).sort((a, b) => a.anio - b.anio);
    return { archivos: ok.map((f) => f.name), anios, byYear, total, rows: { length: nRows } };
  }

  function rebuildAll() {
    renderFiles();
    const P = stagingPayload();
    const exist = new Set((serverData?.anios || []).map((a) => a.anio));
    const rep = P.anios.filter((a) => exist.has(a.anio)).map((a) => a.anio);
    $('#stagingInfo').innerHTML = P.anios.length
      ? `Se guardarán los años <b>${P.anios.map((a) => a.anio).join(', ')}</b> (${money(P.total)}).` + (rep.length ? ` Ya existen en la base y se reemplazarán: <b>${rep.join(', ')}</b>.` : '')
      : '';
    $('#btnSave').disabled = !P.rows.length;
  }

  async function saveToBase() {
    const P = stagingPayload();
    if (!P.rows.length) return;
    const exist = new Set((serverData?.anios || []).map((a) => a.anio));
    const rep = P.anios.filter((a) => exist.has(a.anio)).map((a) => a.anio);
    const single = [];
    for (const f of state.files.filter((x) => !x.error)) for (const y in f.monthsByYear || {}) if (f.monthsByYear[y].size === 1) single.push(`${y} (${f.name})`);
    if (single.length && !confirm(`Atención: estos años solo traen un mes de datos: ${single.join(', ')}.\n\nSi los archivos deberían tener todo el año, cancela y revisa la hoja o la columna de fecha. ¿Guardar de todos modos?`)) return;
    if (rep.length && !confirm(`Los años ${rep.join(', ')} ya están en la base y se reemplazarán por completo con estos archivos. ¿Continuar?`)) return;
    const btn = $('#btnSave');
    btn.disabled = true;
    try {
      let filas = 0;
      // Un envío por año: cada año se reemplaza por completo y la carga no se vuelve demasiado pesada.
      for (const a of P.anios) {
        btn.textContent = `Guardando ${a.anio}…`;
        const Y = P.byYear[a.anio] || { rows: [], dias: [], tdias: [] };
        const r = await api('/api/cargas', { method: 'POST', body: { archivos: P.archivos, anios: [a], rows: Y.rows, dias: Y.dias, tdias: Y.tdias } });
        filas += r.filas;
      }
      toast(`Guardado: años ${P.anios.map((a) => a.anio).join(', ')} · ${num(filas)} registros.`);
      state.files = [];
      rebuildAll();
      await loadServer();
      setView('anual');
    } catch (e) {
      toast(e.message);
    } finally {
      btn.textContent = 'Guardar en la base'; btn.disabled = false;
    }
  }

  async function deleteYear(y) {
    if (!confirm(`¿Eliminar todos los datos del año ${y} de la base? Esto no se puede deshacer.`)) return;
    try { await api('/api/anios/' + y, { method: 'DELETE' }); toast(`Año ${y} eliminado.`); await loadServer(); }
    catch (e) { toast(e.message); }
  }

  async function openUsers() {
    const dlg = $('#dlgUsers');
    $('[data-error]', dlg).textContent = '';
    try {
      const { usuarios } = await api('/api/usuarios');
      $('#tblUsers').innerHTML = `<thead><tr><th>DNI</th><th>Nombre</th><th>Rol</th><th></th></tr></thead><tbody>${usuarios.map((u) => `<tr><td>${esc(u.dni)}</td><td>${esc(u.nombre)}</td><td>${u.es_admin ? 'Administrador' : 'Usuario'}</td><td>${u.dni === session.user.dni ? '' : `<button class="btn btn-danger" data-del-user="${esc(u.dni)}">Eliminar</button>`}</td></tr>`).join('')}</tbody>`;
      if (!dlg.open) dlg.showModal();
    } catch (e) { toast(e.message); }
  }

  async function handleFiles(list) {
    const files = [...list];
    if (!files.length) return;
    const bad = files.filter((f) => !/\.(xlsx|xls|xlsb|xlsm|csv|txt)$/i.test(f.name));
    if (bad.length) toast(`Formato no soportado: ${bad.map((f) => f.name).join(', ')}`);
    $('#loading').hidden = false;
    for (const file of files) {
      if (bad.includes(file)) continue;
      if (state.files.some((f) => f.name === file.name && f.size === file.size)) { toast(`"${file.name}" ya estaba cargado.`); continue; }
      $('#loadingText').textContent = `Leyendo ${file.name} (${(file.size / 1048576).toFixed(1)} MB)…`;
      await new Promise((r) => setTimeout(r, 30));
      try {
        const wb = await readFile(file);
        const f = { id: ++fileSeq, name: file.name, size: file.size, wb };
        prepareFile(f);
        state.files.push(f);
      } catch (err) {
        console.error(err);
        state.files.push({ id: ++fileSeq, name: file.name, size: file.size, wb: null, parts: [], monthsByYear: {}, error: 'No se pudo leer el archivo: ' + (err.message || err), headers: [], headerText: [], mapping: {}, warnings: [], years: [], used: 0, total: 0 });
      }
    }
    $('#loading').hidden = true;
    rebuildAll();
  }

  // ---------------------------------------------------------------- panel de archivos
  function renderFiles() {
    const panel = $('#filesPanel');
    panel.hidden = !state.files.length;
    $$('[data-empty]').forEach((el) => (el.hidden = !!state.records.length));
    $$('[data-content]').forEach((el) => (el.hidden = !state.records.length));
    const wrap = $('#filesList');
    wrap.innerHTML = state.files.map((f) => {
      const ok = !f.error;
      const multi = f.usable && f.usable.length >= 2;
      const sheetSel = f.wb && f.wb.SheetNames.length > 1
        ? `<select data-act="sheet" data-id="${f.id}" title="Hoja">${multi ? `<option value="${ALL_SHEETS}" ${f.sheetMode === ALL_SHEETS ? 'selected' : ''}>Todas las hojas (${f.usable.length})</option>` : ''}${f.wb.SheetNames.map((n) => `<option value="${esc(n)}" ${f.sheetMode === n ? 'selected' : ''}>Hoja: ${esc(n)}</option>`).join('')}</select>` : '';
      const hdrSel = f.wb && f.sheetMode !== ALL_SHEETS
        ? `<select data-act="header" data-id="${f.id}" title="Fila de encabezados">${Array.from({ length: Math.min(30, f.totalRows || 1) }, (_, i) => `<option value="${i}" ${i === f.headerRow ? 'selected' : ''}>Encabezados en fila ${i + 1}</option>`).join('')}</select>` : '';
      const opts = (sel) => `<option value="">(ninguna)</option>` + f.headerText.map((h, i) => `<option value="${i}" ${sel === i ? 'selected' : ''}>${esc(h)}</option>`).join('');
      const map = f.wb ? `<div class="map-grid">${ROLES.map((r) => `<div><label>${r.label}</label><select data-act="map" data-role="${r.key}" data-id="${f.id}">${opts(f.mapping[r.key])}</select></div>`).join('')}</div>` : '';
      const info = ok
        ? `<span class="badge ok">✓ ${num(f.used)} registros</span>${Object.keys(f.monthsByYear || {}).sort().map((y) => `<span class="badge ${f.monthsByYear[y].size === 1 ? 'err' : ''}">${y}: ${monthsText(f.monthsByYear[y])} (${f.monthsByYear[y].size} ${f.monthsByYear[y].size === 1 ? 'mes' : 'meses'})</span>`).join('')}${f.sheetMode === ALL_SHEETS ? `<span class="badge">${f.parts.length} hojas</span>` : ''}${f.hasCob && f.total ? `<span class="badge ok">Coberturado: ${money(f.covTotal)} (${fmtPct.format(f.covTotal / f.total)})</span>` : ''}<span class="badge">${money(f.total)}</span>${f.wide ? '<span class="badge">Formato ancho</span>' : ''}`
        : `<span class="badge err">Revisar</span>`;
      const warns = [f.error, ...(f.warnings || [])].filter(Boolean).map((w) => `<div class="file-warn">${esc(w)}</div>`).join('');
      return `<div class="file-item">
        <div class="file-top"><span class="file-name">${esc(f.name)}</span>${info}<span class="spacer"></span>${sheetSel}${hdrSel}<button class="link" data-act="remove" data-id="${f.id}">Quitar</button></div>
        ${map}${warns}
      </div>`;
    }).join('');
  }

  function onFilesPanelChange(e) {
    const el = e.target.closest('[data-act]');
    if (!el) return;
    const f = state.files.find((x) => x.id === +el.dataset.id);
    if (!f) return;
    const act = el.dataset.act;
    if (act === 'sheet') prepareFile(f, { sheet: el.value });
    else if (act === 'header') prepareFile(f, { headerRow: +el.value });
    else if (act === 'map') {
      const role = el.dataset.role;
      const mapping = { ...f.mapping };
      if (el.value === '') delete mapping[role];
      else {
        const idx = +el.value;
        for (const k in mapping) if (mapping[k] === idx) delete mapping[k];
        mapping[role] = idx;
      }
      f.mapping = mapping;
      store.set(f.sig, mapping);
      // En modo "todas las hojas", el cambio se aplica a las hojas con los mismos encabezados.
      for (const p of f.parts) if (p.sig === f.sig) p.mapping = mapping;
      buildFileAgg(f);
    }
    rebuildAll();
  }

  // ---------------------------------------------------------------- filtros
  function syncFilterOptions() {
    const setOpts = (sel, items, val) => {
      sel.innerHTML = items.map(([v, t]) => `<option value="${esc(v)}">${esc(t)}</option>`).join('');
      if (items.some(([v]) => String(v) === String(val))) sel.value = val;
    };
    const ys = state.years;
    const prev = $('#fFrom').value;
    const auto = ys.find((y) => y >= 2022) ?? ys[0];
    const def = state.fromTouched && prev && ys.includes(+prev) ? +prev : auto;
    setOpts($('#fFrom'), ys.map((y) => [y, y]), def);

    const chains = [...new Set(state.records.map((r) => r.cadena))].sort((a, b) => a.localeCompare(b, 'es'));
    setOpts($('#fChain'), [['', 'Todas'], ...chains.map((c) => [c, c])], $('#fChain').value);
    const this_brand = $('#fBrand').value;
    const brands = [...new Set(state.records.map((r) => r.marca).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
    setOpts($('#fBrand'), [['', 'Todas'], ...brands.map((c) => [c, c])], $('#fBrand').value);
    $('#fBrandWrap').hidden = !brands.length;
    $('#fTipoWrap').hidden = !brands.length;
    $('#fCobWrap').hidden = !state.hasCob;
    $('#fCob option[value="sd"]').hidden = !state.records.some((r) => r.cob == null);
    // Marca: muestra a qué grupo pertenece cada una.
    $('#fBrand').innerHTML = [['', 'Todas'], ...brands.map((c) => [c, `${c} · ${tipoDe(c) === 'paraiso' ? 'Paraíso' : 'Competencia'}`])].map(([v, t]) => `<option value="${esc(v)}">${esc(t)}</option>`).join('');
    if (brands.includes(this_brand)) $('#fBrand').value = this_brand;
  }

  // ---------------------------------------------------------------- marcas: venta Paraíso vs competencia
  // Grupo Paraíso: estas marcas son "Venta Paraíso"; cualquier otra marca es "Venta Competencia".
  const MARCAS_PARAISO = ['PARAISO', 'VIVA-HOME', 'CASA JOVEN', 'KREA', 'KOMFORT'];
  const compact = (s) => norm(s).replace(/\s+/g, '');
  const PARAISO_SET = new Set(MARCAS_PARAISO.map(compact));
  const COB_LABEL = { si: 'Solo coberturados', no: 'Solo no coberturados', sd: 'Sin dato de cobertura', all: 'Todos' };
  const TIPO_LABEL = { paraiso: 'Venta Paraíso', competencia: 'Venta Competencia', sinmarca: 'Sin marca' };
  function tipoDe(marca) {
    const c = compact(marca);
    if (!c) return 'sinmarca';
    return PARAISO_SET.has(c) ? 'paraiso' : 'competencia';
  }

  // ---------------------------------------------------------------- colores por año
  // El último año va en azul y el anterior en naranja; cada año conserva su color en toda la app.
  const PALETTE = ['#3D6BFF', '#FF8A3D', '#14B88A', '#8B5CF6', '#E0A100', '#E8618C', '#0EA5E9', '#8A93A3'];
  const yearColor = (y) => {
    const i = [...state.years].reverse().indexOf(y);
    return i < 0 ? '#9aa1ad' : i < PALETTE.length ? PALETTE[i] : '#9aa1ad';
  };

  // ---------------------------------------------------------------- cálculos

  function cutInfo() {
    const c = state.cut;
    const hasW = !!c.w;
    return {
      ...c, hasW,
      ytdLabel: (y) => (hasW && state.weekYears.has(y) ? `Sem 1–${c.w}` : `Ene–${MESES[c.m - 1]}`),
      monthLabel: MESES_LARGO[c.m - 1],
      dateText: c.d ? `${String(c.d).padStart(2, '0')}/${String(c.m).padStart(2, '0')}/${c.y}` : `${MESES_LARGO[c.m - 1]} ${c.y}`,
    };
  }

  function currentFilters() {
    const C = cutInfo();
    const from = +$('#fFrom').value || state.years[0];
    const years = state.years.filter((y) => y >= from && y <= C.y);
    const chain = $('#fChain').value, brand = $('#fBrand').value, tipo = state.hasBrand ? $('#fTipo').value : '';
    const cob = state.hasCob ? $('#fCob').value : 'all';
    const cobOk = (r) => cob === 'all' || (cob === 'si' ? r.cob === 1 : cob === 'no' ? r.cob === 0 : r.cob == null);
    const recs = state.records.filter((r) => r.y >= from && r.y <= C.y && (!chain || r.cadena === chain) && (!brand || r.marca === brand) && (!tipo || r.tipo === tipo) && cobOk(r));
    // La participación Paraíso solo tiene sentido viendo todas las marcas.
    const showShare = state.hasBrand && !tipo && !brand;
    return { C, from, years, chain, brand, tipo, cob, showShare, recs };
  }

  const blank = () => ({ full: 0, ytd: 0, month: 0, week: 0 });

  function compute(F) {
    const { C } = F;
    const inYtd = (r) => (C.hasW && state.weekYears.has(r.y) && r.w > 0 ? r.w <= C.w : r.m <= C.m);
    const byYear = new Map(F.years.map((y) => [y, { ...blank(), u: 0, stores: new Set(), storesFull: new Set(), p: blank(), c: blank() }]));
    const byYM = new Map(), byYW = new Map();
    const byChain = new Map(), byStore = new Map();
    let maxW = 0;
    for (const r of F.recs) {
      const Y = byYear.get(r.y);
      if (!Y) continue;
      // Mes de corte parejo: en los demás años, el mes solo cuenta hasta la semana de corte (no el mes completo).
      const ytd = inYtd(r), mo = r.m === C.m && (r.y === C.y || !(C.hasW && state.weekYears.has(r.y) && r.w > 0) || r.w <= C.w), wk = C.hasW && r.w === C.w;
      const sk = r.cadena + '\u0001' + r.tienda;
      Y.full += r.v; if (r.v) Y.storesFull.add(sk);
      if (ytd) { Y.ytd += r.v; Y.u += r.u; if (r.v) Y.stores.add(sk); }
      if (mo) Y.month += r.v;
      if (wk) Y.week += r.v;
      if (r.tipo !== 'sinmarca') {
        const G = r.tipo === 'paraiso' ? Y.p : Y.c;
        G.full += r.v; if (ytd) G.ytd += r.v; if (mo) G.month += r.v; if (wk) G.week += r.v;
      }
      byYM.set(r.y * 100 + r.m, (byYM.get(r.y * 100 + r.m) || 0) + r.v);
      if (r.w > 0) { byYW.set(r.y * 100 + r.w, (byYW.get(r.y * 100 + r.w) || 0) + r.v); if (r.v) maxW = Math.max(maxW, r.w); }
      for (const [map, key, base] of [[byChain, r.cadena, { cadena: r.cadena }], [byStore, sk, { cadena: r.cadena, tienda: r.tienda }]]) {
        let o = map.get(key);
        if (!o) map.set(key, (o = { ...base, y: {}, yp: {} }));
        const oy = o.y[r.y] || (o.y[r.y] = blank());
        oy.full += r.v; if (ytd) oy.ytd += r.v; if (mo) oy.month += r.v; if (wk) oy.week += r.v;
        if (r.tipo === 'paraiso') {
          const op = o.yp[r.y] || (o.yp[r.y] = blank());
          op.full += r.v; if (ytd) op.ytd += r.v; if (mo) op.month += r.v; if (wk) op.week += r.v;
        }
      }
    }
    return { byYear, byYM, byYW, byChain, byStore, maxW };
  }

  const varPct = (b, a) => (a > 0 ? (b - a) / a : null);
  function statusOf(a, b) {
    if (a <= 0 && b > 0) return 'nueva';
    if (a > 0 && b <= 0) return 'perdida';
    if (b > a) return 'crece';
    if (b < a) return 'cae';
    return 'igual';
  }
  const STATUS_LABEL = { crece: '▲ Crece', cae: '▼ Cae', nueva: '● Nueva', perdida: '✕ Sin venta', igual: '= Igual' };

  // ---------------------------------------------------------------- tablas (datos reutilizados para copiar y exportar)
  // types: t texto · $ soles · % variación con signo · p porcentaje · n número
  function yearsAOA(F, R) {
    const { C } = F;
    const head = ['Año', 'Año completo', 'Var %', `Acumulado al corte`, 'Var S/', 'Var %', `Mes de corte (${MESES[C.m - 1]})`, 'Var %'];
    const types = ['t', '$', '%', '$', '$', '%', '$', '%'];
    if (C.hasW) { head.push(`Semana ${C.w}`, 'Var %'); types.push('$', '%'); }
    head.push('Tiendas con venta (acum.)', 'Período acumulado');
    types.push('n', 't');
    const rows = F.years.map((y) => {
      const Y = R.byYear.get(y), P = R.byYear.get(y - 1);
      const row = [String(y) + (y === C.y ? ' *' : ''), Y.full, P && y !== C.y ? varPct(Y.full, P.full) : null, Y.ytd, P ? Y.ytd - P.ytd : null, P ? varPct(Y.ytd, P.ytd) : null, Y.month, P ? varPct(Y.month, P.month) : null];
      if (C.hasW) row.push(state.weekYears.has(y) ? Y.week : null, P && state.weekYears.has(y) && state.weekYears.has(y - 1) ? varPct(Y.week, P.week) : null);
      row.push(Y.stores.size, C.ytdLabel(y));
      return row;
    });
    return { head, rows, foot: [], types, varCols: [2, 4, 5, 7, 9] };
  }

  function groupsAOA(F, R) {
    const { C } = F;
    const share = (G, A, k) => (G[k] + A[k] > 0 ? G[k] / (G[k] + A[k]) : null);
    const head = ['Año', 'Venta Paraíso (acum.)', 'Var %', 'Venta Competencia (acum.)', 'Var %', 'Part. Paraíso (acum.)', 'Var pp', `Part. Paraíso ${MESES[C.m - 1]}`];
    const types = ['t', '$', '%', '$', '%', 'p', 'pp', 'p'];
    if (C.hasW) { head.push(`Part. Paraíso sem ${C.w}`); types.push('p'); }
    head.push('Venta Paraíso (año completo)', 'Venta Competencia (año completo)', 'Part. Paraíso (año completo)');
    types.push('$', '$', 'p');
    const rows = F.years.map((y) => {
      const Y = R.byYear.get(y), P = R.byYear.get(y - 1);
      const sh = share(Y.p, Y.c, 'ytd'), shP = P ? share(P.p, P.c, 'ytd') : null;
      const row = [String(y) + (y === C.y ? ' *' : ''), Y.p.ytd, P ? varPct(Y.p.ytd, P.p.ytd) : null, Y.c.ytd, P ? varPct(Y.c.ytd, P.c.ytd) : null, sh, sh != null && shP != null ? sh - shP : null, share(Y.p, Y.c, 'month')];
      if (C.hasW) row.push(state.weekYears.has(y) ? share(Y.p, Y.c, 'week') : null);
      row.push(Y.p.full, Y.c.full, share(Y.p, Y.c, 'full'));
      return row;
    });
    return { head, rows, foot: [], types, varCols: [2, 4, 6] };
  }

  // ---------------------------------------------------------------- render
  let lastF = null, lastR = null;

  function cell(v, t) {
    if (v == null) return '—';
    if (t === '$') return money(v);
    if (t === '%') return pct(v);
    if (t === 'p') return fmtPct.format(v);
    if (t === 'pp') return (v > 0 ? '+' : '') + (v * 100).toFixed(1) + ' pp';
    if (t === 'n') return num(v);
    return esc(v);
  }
  const dirCls = (v) => (v > 0 ? 'up' : v < 0 ? 'down' : '');

  function renderTable(id, T, sortKind) {
    const sort = sortKind ? state.sort[sortKind] : null;
    const th = T.head.map((h, i) => {
      const k = T.keys?.[i];
      const arrow = sort && k === sort.key ? `<span class="arrow">${sort.dir > 0 ? '▲' : '▼'}</span>` : '';
      return `<th class="${T.types[i] === 't' ? '' : 'num'}" ${k ? `data-sort="${k}"` : ''}>${esc(h)}${arrow}</th>`;
    }).join('');
    const row = (r) => `<tr>${r.map((v, i) => {
      if (i === T.stCol && v) { const st = Object.keys(STATUS_LABEL).find((s) => STATUS_LABEL[s].endsWith(v)); return `<td><span class="pill ${st}">${STATUS_LABEL[st]}</span></td>`; }
      if (T.types[i] === 't') return `<td>${esc(v)}</td>`;
      if (T.bar === i && v != null && T.barMax) return `<td class="num dbar ${dirCls(v)}" style="--w:${Math.min(1, Math.abs(v) / T.barMax).toFixed(3)}">${cell(v, T.types[i])}</td>`;
      return `<td class="num ${T.varCols?.includes(i) ? dirCls(v) : ''}">${cell(v, T.types[i])}</td>`;
    }).join('')}</tr>`;
    const body = T.rows.length ? T.rows.map(row).join('') : `<tr><td colspan="${T.head.length}" class="muted">Sin datos para estos filtros.</td></tr>`;
    $('#' + id).innerHTML = `<thead><tr>${th}</tr></thead><tbody>${body}</tbody>${T.foot.length ? `<tfoot>${T.foot.map(row).join('')}</tfoot>` : ''}`;
  }

  function renderAll() {
    if (!state.records.length || !state.cut) {
      Object.values(state.charts).forEach((c) => c.destroy());
      state.charts = {};
      return;
    }
    const F = currentFilters();
    const R = compute(F);
    lastF = F; lastR = R;
    renderBanner(F);
    renderResumen(F, R);
    renderKpis(F, R);
    renderAnual(F, R);
    renderYearCharts(F, R);
    renderTable('tblYears', yearsAOA(F, R));
    $('#groupsCard').hidden = !F.showShare;
    if (F.showShare) { renderTable('tblGroups', groupsAOA(F, R)); renderGroupChart(F, R); }
    $('#noteYears').textContent = `* ${F.C.y} es el año en curso: su "Año completo" llega solo hasta el corte, por eso no se compara con el año anterior. El acumulado de cada año se calcula hasta el mismo punto (${F.C.hasW ? 'semana ' + F.C.w : MESES_LARGO[F.C.m - 1]}) para comparar en igualdad de condiciones.` + (F.C.hasW && F.years.some((y) => !state.weekYears.has(y)) ? ' Los años sin fecha diaria se acumulan hasta el mes de corte completo.' : '');
    renderCmp();
  }

  // Cuántos filtros cambiaron respecto de lo normal (se muestra en la pestaña "Filtros").
  function renderFilterCount(F) {
    let n = 0;
    if (F.tipo) n++;
    if (F.chain) n++;
    if (F.brand) n++;
    if (state.hasCob && F.cob !== 'si') n++;
    if (state.fromTouched && F.from !== (state.years.find((y) => y >= 2022) ?? state.years[0])) n++;
    const el = $('#filterCount');
    el.hidden = !n; el.textContent = n;
    $('#btnFilters').setAttribute('aria-label', n ? `Filtros (${n} activos)` : 'Filtros');
  }

  function setDrawer(open) {
    const d = $('#filterDrawer');
    d.classList.toggle('open', open);
    d.setAttribute('aria-hidden', String(!open));
    $('#drawerScrim').hidden = !open;
    $('#btnFilters').setAttribute('aria-expanded', String(open));
    if (open) $('#fFrom').focus(); else if (d.contains(document.activeElement)) $('#btnFilters').focus();
  }

  function renderBanner(F) {
    renderFilterCount(F);
    const { C } = F;
    const parts = [`<b>Año ${C.y}</b>`, `<b>${C.monthLabel}</b>`];
    if (C.hasW) parts.push(`<b>Semana ${C.w}</b>`);
    let t = `<span class="cut-tag">Corte de evaluación</span> ${parts.join(' · ')} <span class="muted">— último dato en la base: ${C.dateText}</span>`;
    t += `<span class="cut-years">Años evaluados: ${F.years.join(', ')}${F.cob && F.cob !== 'all' ? ' · <b>' + COB_LABEL[F.cob] + '</b>' : ''}${F.tipo ? ' · <b>' + TIPO_LABEL[F.tipo] + '</b>' : ''}${F.chain ? ' · Cadena: ' + esc(F.chain) : ''}${F.brand ? ' · Marca: ' + esc(F.brand) : ''}</span>`;
    $('#cutBanner').innerHTML = t;
  }

  function kpi(label, value, sub, cls = '') {
    return `<div class="kpi"><div class="k-label">${label}</div><div class="k-value ${cls}">${value}</div>${sub ? `<div class="k-sub">${sub}</div>` : ''}</div>`;
  }

  function renderKpis(F, R) {
    const { C } = F;
    const Y = R.byYear.get(C.y), P = R.byYear.get(C.y - 1);
    const vs = (b, a) => {
      if (!P) return '';
      const p = varPct(b, a), d = b - a;
      return `<span class="${dirCls(d)}">${d > 0 ? '▲' : d < 0 ? '▼' : ''} ${pct(p)} · ${signMoney(d)}</span> vs ${C.y - 1}`;
    };
    const best = F.years.reduce((b, y) => (R.byYear.get(y).ytd > (b ? R.byYear.get(b).ytd : -Infinity) ? y : b), null);
    const cards = [
      kpi(`Acumulado ${C.y} (${C.ytdLabel(C.y)})`, money(Y.ytd), vs(Y.ytd, P?.ytd || 0)),
      kpi(`Mes de corte · ${C.monthLabel} ${C.y}`, money(Y.month), vs(Y.month, P?.month || 0)),
    ];
    if (C.hasW) cards.push(kpi(`Semana ${C.w} · ${C.y}`, money(Y.week), state.weekYears.has(C.y - 1) ? vs(Y.week, P?.week || 0) : ''));
    cards.push(kpi('Tiendas con venta', num(Y.stores.size), P ? `${num(P.stores.size)} en ${C.y - 1} (mismo corte)` : ''));
    if (F.showShare) {
      const sh = (G, A) => (G.ytd + A.ytd > 0 ? G.ytd / (G.ytd + A.ytd) : null);
      const s1 = sh(Y.p, Y.c), s0 = P ? sh(P.p, P.c) : null;
      const dpp = s1 != null && s0 != null ? s1 - s0 : null;
      cards.push(kpi('Participación Paraíso (acum.)', s1 == null ? '—' : fmtPct.format(s1),
        dpp == null ? `Paraíso ${money(Y.p.ytd)}` : `<span class="${dirCls(dpp)}">${dpp > 0 ? '▲' : dpp < 0 ? '▼' : ''} ${(dpp > 0 ? '+' : '') + (dpp * 100).toFixed(1)} pp</span> vs ${C.y - 1}`));
    }
    if (best != null) cards.push(kpi('Mejor año al corte', String(best), money(R.byYear.get(best).ytd)));
    $('#kpis').innerHTML = cards.join('');
  }

  const GRID = '#eef0f3', TICK = '#7a8292';
  const moneyTick = (v) => (Math.abs(v) >= 1e6 ? 'S/ ' + (v / 1e6).toFixed(1) + ' M' : Math.abs(v) >= 1e3 ? 'S/ ' + Math.round(v / 1e3) + ' K' : 'S/ ' + v);

  function upsertChart(key, canvasId, config) {
    // Si el gráfico está en una pestaña oculta, se dibuja sin animación para que aparezca listo al mostrarla.
    if (!$('#' + canvasId).offsetParent) config.options = { ...config.options, animation: false };
    const ch = state.charts[key];
    if (ch) { ch.config.type = config.type; ch.data = config.data; ch.options = config.options; ch.update(); }
    else state.charts[key] = new Chart($('#' + canvasId), config);
  }

  function renderYearCharts(F, R) {
    const { C } = F;
    $('#ttlYtdChart').textContent = `Acumulado al corte por año (${C.hasW ? 'sem 1–' + C.w : 'Ene–' + MESES[C.m - 1]})`;
    const bar = (field, id, key) => upsertChart(key, id, {
      type: 'bar',
      data: { labels: F.years.map(String), datasets: [{ data: F.years.map((y) => R.byYear.get(y)[field]), backgroundColor: F.years.map(yearColor), borderRadius: 4, borderSkipped: 'bottom', barPercentage: 0.6 }] },
      options: {
        responsive: true, maintainAspectRatio: false, animation: { duration: 250 },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: (c) => { const y = F.years[c.dataIndex], P = R.byYear.get(y - 1); const v = c.parsed.y; return ` ${money(v)}` + (P ? `  (${pct(varPct(v, P[field]))} vs ${y - 1})` : ''); } } },
        },
        scales: { x: { grid: { display: false }, ticks: { color: TICK } }, y: { beginAtZero: true, grid: { color: GRID }, border: { display: false }, ticks: { color: TICK, callback: moneyTick } } },
      },
    });
    bar('ytd', 'chartYears', 'y1');
    bar('full', 'chartYearsFull', 'y2');
  }

  const COL_PARAISO = '#3D6BFF', COL_COMP = '#B4B9D3';
  function renderGroupChart(F, R) {
    $('#groupLegend').innerHTML = `<span><i class="swatch" style="background:${COL_PARAISO}"></i>Venta Paraíso</span><span><i class="swatch" style="background:${COL_COMP}"></i>Venta Competencia</span>`;
    const ds = (label, field, color) => ({ label, data: F.years.map((y) => R.byYear.get(y)[field].ytd), backgroundColor: color, borderRadius: 4, borderSkipped: 'bottom', categoryPercentage: 0.7, barPercentage: 0.9 });
    upsertChart('g', 'chartGroups', {
      type: 'bar',
      data: { labels: F.years.map(String), datasets: [ds('Venta Paraíso', 'p', COL_PARAISO), ds('Venta Competencia', 'c', COL_COMP)] },
      options: {
        responsive: true, maintainAspectRatio: false, animation: { duration: 250 },
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: {
            label: (c) => ` ${c.dataset.label}: ${money(c.parsed.y)}`,
            footer: (it) => { const Y = R.byYear.get(F.years[it[0].dataIndex]); const t = Y.p.ytd + Y.c.ytd; return t ? `Part. Paraíso: ${fmtPct.format(Y.p.ytd / t)}` : ''; },
          } },
        },
        scales: { x: { grid: { display: false }, ticks: { color: TICK } }, y: { beginAtZero: true, grid: { color: GRID }, border: { display: false }, ticks: { color: TICK, callback: moneyTick } } },
      },
    });
  }

  // ---------------------------------------------------------------- administración: usuarios y perfiles
  const PERM_INFO = {
    ver_todo: ['Ver todas las tiendas', 'Si no se marca, solo ve las tiendas del coordinador asignado al usuario.'],
    subir: ['Subir y reemplazar datos', 'Cargar archivos de Excel o CSV a la base.'],
    eliminar: ['Eliminar años de la base', 'Borrar todos los datos de un año.'],
    gestionar_coord: ['Cambiar nombres de coordinadores', 'Crear o quitar reemplazos de nombres.'],
    descargar: ['Descargar PDF y Excel', 'Bajar lo que tiene en pantalla.'],
    admin_usuarios: ['Gestionar usuarios y perfiles', 'Crear, editar y eliminar usuarios y perfiles.'],
  };
  const admin = { users: [], perfiles: [] };

  function fillCoordList() {
    $('#coordListAll').innerHTML = (coordData?.coordinadores || []).map((c) => `<option value="${esc(c.coordinador)}">`).join('');
  }

  async function loadAdmin() {
    if (!perm('admin_usuarios')) return;
    admin.catalog = null;
    try {
      const [u, p] = await Promise.all([api('/api/usuarios'), api('/api/perfiles')]);
      admin.users = u.usuarios; admin.perfiles = p.perfiles;
      renderAdmin();
    } catch (e) { toast(e.message); }
  }

  function renderAdmin() {
    const P = admin.perfiles;
    $('#tblAdminUsers').innerHTML = `<thead><tr><th>Nombre</th><th>DNI</th><th>Perfil</th><th>Ve</th><th></th></tr></thead><tbody>${
      admin.users.map((u) => `<tr><td><b>${esc(u.nombre)}</b></td><td>${esc(u.dni)}</td><td><span class="role-chip">${esc(u.perfil || 'Sin perfil')}</span></td>
        <td>${u.ver_todo ? 'Todas las tiendas' : u.accesos.length ? `<span class="acc-sum">${u.accesos.map((a) => `<span class="role-chip">${esc(accText(a))}</span>`).join('')}</span>` : '<span class="down">Sin accesos asignados</span>'}</td>
        <td class="num"><button class="btn" data-edit-user="${esc(u.dni)}">Editar</button> ${u.dni === session.user.dni ? '' : `<button class="btn btn-danger" data-del-user="${esc(u.dni)}">Eliminar</button>`}</td></tr>`).join('')}</tbody>`;
    $('#profileCards').innerHTML = P.map((p) => `<article class="pc">
      <div class="pc-head"><span class="pc-name">${esc(p.nombre)}</span><span class="pc-count">${p.usuarios} ${p.usuarios === 1 ? 'usuario' : 'usuarios'}</span></div>
      <ul class="pc-perms">${Object.keys(PERM_INFO).map((k) => `<li class="${p[k] ? 'on' : 'off'}">${PERM_INFO[k][0]}</li>`).join('')}</ul>
      <div class="pc-actions"><button class="btn" data-edit-profile="${p.id}">Editar</button>${p.sistema ? '' : `<button class="btn btn-danger" data-del-profile="${p.id}">Eliminar</button>`}</div>
    </article>`).join('');
  }

  function openUserDialog(u) {
    const d = $('#dlgUser'), f = $('#formUserEdit');
    f.reset(); $('[data-error]', f).textContent = '';
    $('[data-title]', d).textContent = u ? `Editar a ${u.nombre}` : 'Nuevo usuario';
    f.dni.value = u?.dni || ''; f.dni.readOnly = !!u;
    f.nombre.value = u?.nombre || '';
    f.perfil_id.innerHTML = admin.perfiles.map((p) => `<option value="${p.id}">${esc(p.nombre)}</option>`).join('');
    f.perfil_id.value = u?.perfil_id || (admin.perfiles.find((p) => p.nombre === 'Coordinador') || admin.perfiles[0]).id;
    admin.acc = (u?.accesos || []).map((a) => ({ ...a }));
    renderAccList();
    loadCatalog();
    f.password.required = !u;
    $('[data-pwhint]', d).textContent = u ? '(déjala vacía para no cambiarla)' : '(mín. 8)';
    const sync = () => { const p = admin.perfiles.find((x) => x.id === +f.perfil_id.value); $('[data-coord]', d).hidden = !!p?.ver_todo; };
    f.perfil_id.onchange = sync; sync();
    d.dataset.mode = u ? 'edit' : 'new';
    d.showModal();
  }

  // ---- editor de accesos (coordinador, cadena completa o tienda puntual)
  async function loadCatalog() {
    if (!admin.catalog) { try { admin.catalog = await api('/api/catalogo'); } catch { admin.catalog = { cadenas: [], tiendas: [], coordinadores: [] }; } }
    fillAccOptions();
  }
  function fillAccOptions() {
    const c = admin.catalog; if (!c) return;
    const t = $('#accTipo').value;
    const opts = t === 'coordinador' ? c.coordinadores : t === 'cadena' ? c.cadenas : c.tiendas.map((x) => x.replace('|', ' | '));
    $('#accOptions').innerHTML = opts.map((o) => `<option value="${esc(o)}">`).join('');
    $('#accValor').placeholder = t === 'coordinador' ? 'Ej. HERLES CORNELIO' : t === 'cadena' ? 'Ej. TOTTUS' : 'Ej. TOTTUS | JOCKEY';
  }
  function renderAccList() {
    $('#accList').innerHTML = admin.acc.length
      ? admin.acc.map((a, i) => `<li><span class="k ${a.tipo}">${ACC_LABEL[a.tipo]}</span><b>${esc(a.tipo === 'tienda' ? a.valor.replace('|', ' · ') : a.valor)}</b><button type="button" data-acc-del="${i}" aria-label="Quitar acceso">✕</button></li>`).join('')
      : '<li class="empty">Sin accesos: este usuario no verá ninguna tienda.</li>';
  }
  function addAcceso() {
    const tipo = $('#accTipo').value, raw = $('#accValor').value.trim(), err = $('#accError');
    err.textContent = '';
    if (!raw) return;
    let valor = raw.toUpperCase().replace(/\s+/g, ' ');
    if (tipo === 'tienda') {
      const parts = valor.split('|').map((x) => x.trim());
      if (parts.length !== 2 || !parts[0] || !parts[1]) { err.textContent = 'Para una tienda escribe la cadena y la tienda, así: TOTTUS | JOCKEY. Elígela de la lista para no equivocarte.'; return; }
      valor = parts.join('|');
    }
    const known = admin.catalog && (tipo === 'coordinador' ? admin.catalog.coordinadores : tipo === 'cadena' ? admin.catalog.cadenas : admin.catalog.tiendas);
    if (known && known.length && !known.includes(valor)) err.textContent = `Ojo: "${raw}" todavía no aparece en los datos. Se guarda igual y aplicará cuando aparezca.`;
    if (!admin.acc.some((a) => a.tipo === tipo && a.valor === valor)) admin.acc.push({ tipo, valor });
    $('#accValor').value = '';
    renderAccList();
  }

  function openProfileDialog(p) {
    const d = $('#dlgProfile'), f = $('#formProfile');
    f.reset(); $('[data-error]', f).textContent = '';
    $('[data-title]', d).textContent = p ? `Editar perfil ${p.nombre}` : 'Nuevo perfil';
    f.id.value = p?.id || ''; f.nombre.value = p?.nombre || ''; f.nombre.readOnly = !!p?.sistema;
    $('#permList').innerHTML = Object.entries(PERM_INFO).map(([k, [t, sub]]) => `<label><input type="checkbox" name="${k}" ${p?.[k] ? 'checked' : ''} ${p?.sistema && k === 'admin_usuarios' ? 'disabled checked' : ''}><span>${t}<small>${sub}</small></span></label>`).join('');
    d.showModal();
  }

  function initAdmin() {
    $('#accTipo').addEventListener('change', () => { $('#accValor').value = ''; fillAccOptions(); });
    $('#accAdd').addEventListener('click', addAcceso);
    $('#accValor').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); addAcceso(); } });
    $('#accList').addEventListener('click', (e) => { const b = e.target.closest('[data-acc-del]'); if (b) { admin.acc.splice(+b.dataset.accDel, 1); renderAccList(); } });
    $('#btnNewUser').addEventListener('click', () => openUserDialog(null));
    $('#btnNewProfile').addEventListener('click', () => openProfileDialog(null));
    $('#tblAdminUsers').addEventListener('click', async (e) => {
      const ed = e.target.closest('[data-edit-user]'), del = e.target.closest('[data-del-user]');
      if (ed) openUserDialog(admin.users.find((u) => u.dni === ed.dataset.editUser));
      if (del && confirm(`¿Eliminar el usuario con DNI ${del.dataset.delUser}? Ya no podrá ingresar.`)) {
        try { await api('/api/usuarios/' + encodeURIComponent(del.dataset.delUser), { method: 'DELETE' }); toast('Usuario eliminado.'); loadAdmin(); } catch (x) { toast(x.message); }
      }
    });
    $('#profileCards').addEventListener('click', async (e) => {
      const ed = e.target.closest('[data-edit-profile]'), del = e.target.closest('[data-del-profile]');
      if (ed) openProfileDialog(admin.perfiles.find((p) => p.id === +ed.dataset.editProfile));
      if (del && confirm('¿Eliminar este perfil?')) {
        try { await api('/api/perfiles/' + del.dataset.delProfile, { method: 'DELETE' }); toast('Perfil eliminado.'); loadAdmin(); } catch (x) { toast(x.message); }
      }
    });
    $('#formUserEdit').addEventListener('submit', async (e) => {
      e.preventDefault();
      const f = e.target, err = $('[data-error]', f), mode = $('#dlgUser').dataset.mode;
      const body = { dni: f.dni.value.trim(), nombre: f.nombre.value, perfil_id: +f.perfil_id.value, accesos: $('[data-coord]', f).hidden ? [] : admin.acc, password: f.password.value };
      try {
        await api(mode === 'edit' ? '/api/usuarios/' + encodeURIComponent(body.dni) : '/api/usuarios', { method: 'POST', body });
        $('#dlgUser').close(); toast(mode === 'edit' ? 'Usuario actualizado.' : 'Usuario creado.'); loadAdmin();
        if (body.dni === session.user.dni) api('/api/me').then((r) => { session.user = r.user; store.set('user', r.user); renderSession(); loadServer(); });
      } catch (x) { err.textContent = x.message; }
    });
    $('#formProfile').addEventListener('submit', async (e) => {
      e.preventDefault();
      const f = e.target, err = $('[data-error]', f);
      const body = { id: f.id.value ? +f.id.value : undefined, nombre: f.nombre.value };
      for (const k of Object.keys(PERM_INFO)) body[k] = f[k].checked;
      try {
        await api('/api/perfiles', { method: 'POST', body });
        $('#dlgProfile').close(); toast(body.id ? 'Perfil actualizado.' : 'Perfil creado.'); loadAdmin();
        api('/api/me').then((r) => { session.user = r.user; store.set('user', r.user); renderSession(); });
      } catch (x) { err.textContent = x.message; }
    });
  }

  // ---------------------------------------------------------------- por año: titular y línea de tiempo
  function renderAnual(F, R) {
    const C = F.C, Y = C.y, P = Y - 1;
    const ys = F.years;
    const ytd = (y) => R.byYear.get(y)?.ytd || 0;
    const best = ys.reduce((b, y) => (ytd(y) > ytd(b) ? y : b), ys[0]);
    const p = ys.includes(P) ? varPct(ytd(Y), ytd(P)) : null;
    $('#anualCut').textContent = C.d ? `${C.d} de ${C.monthLabel.toLowerCase()}` : C.monthLabel.toLowerCase();
    const delta = p == null ? '' : ` <span class="delta ${p >= 0 ? 'up' : 'down'}">${(p > 0 ? '+' : '') + (p * 100).toFixed(1)}%</span> <span class="vs">frente a ${P}</span>`;
    $('#anualHeadline').innerHTML = best === Y
      ? `${Y} es el mejor año al corte: <span class="amount">${moneyShort(ytd(Y))}</span>,${delta}`
      : `${Y} va <span class="amount">${moneyShort(ytd(Y))}</span> al corte,${delta}`;
    const first = ys[0];
    const pt = first !== Y ? varPct(ytd(Y), ytd(first)) : null;
    $('#anualSub').textContent = [
      best !== Y ? `El mejor año al corte sigue siendo ${best} (${moneyShort(ytd(best))}).` : null,
      pt != null ? `Desde ${first}, la venta al mismo corte ${pt >= 0 ? 'creció' : 'cayó'} ${pct(Math.abs(pt)).replace('+', '')} en ${ys.length - 1} ${ys.length - 1 === 1 ? 'año' : 'años'}.` : null,
    ].filter(Boolean).join(' ');
    const max = Math.max(1, ...ys.map(ytd));
    const showShare = state.hasBrand && !F.tipo && !F.brand;
    $('#anualTimeline').innerHTML = ys.map((y) => {
      const Yo = R.byYear.get(y), Po = R.byYear.get(y - 1);
      const d = Po ? varPct(Yo.ytd, Po.ytd) : null;
      const sh = Yo.p.ytd + Yo.c.ytd > 0 ? Yo.p.ytd / (Yo.p.ytd + Yo.c.ytd) : null;
      const badge = y === Y ? '<span class="tl-badge now">Año en curso</span>' : y === best ? '<span class="tl-badge best">Mejor año al corte</span>' : '<span class="tl-badge ghost" aria-hidden="true">&nbsp;</span>';
      return `<li class="tl ${y === Y ? 'now' : ''}" style="--tc:${yearColor(y)}">
        <span class="tl-year">${y}</span>${badge}
        <div class="tl-row"><span>Al corte</span><b>${moneyShort(Yo.ytd)}</b></div>
        <div class="tl-row"><span>vs año anterior</span><span class="tl-delta ${dirCls(d)}">${d == null ? '—' : pct(d)}</span></div>
        <div class="tl-row"><span>${y === Y ? 'En lo que va del año' : 'Año completo'}</span><b>${moneyShort(Yo.full)}</b></div>
        ${showShare && sh != null ? `<div class="tl-row"><span>Part. Paraíso</span><b>${fmtPct.format(sh)}</b></div>` : ''}
        <div class="tl-bar" aria-hidden="true"><i style="width:${(Yo.ytd / max) * 100}%"></i></div>
      </li>`;
    }).join('');
  }

  // ---------------------------------------------------------------- resumen (dashboard inicial)
  const CHAIN_PAL = ['#3D6BFF', '#FF8A3D', '#8B5CF6', '#0EA5E9', '#E0A100', '#1B1F4B', '#C2410C', '#0891B2', '#7C3AED', '#A16207'];
  const chainColor = (c) => {
    const all = [...new Set(state.records.map((r) => r.cadena))].sort((a, b) => a.localeCompare(b, 'es'));
    const i = all.indexOf(c);
    return i < 0 ? '#94A3B8' : CHAIN_PAL[i % CHAIN_PAL.length];
  };
  const ppTxt = (x) => (x == null ? '—' : `${x > 0 ? '+' : ''}${(x * 100).toFixed(1)} pp`);

  // Ir al detalle: abre la Comparativa ya filtrada en lo que se tocó.
  function goDetail({ months = null, weeks = null, days = null, by = null, chain = null }) {
    const C = lastF.C;
    cmp.init = true;
    cmp.months = new Set(months ?? []);
    cmp.weeks = new Set(weeks ?? []);
    cmp.days = new Set(days ?? []);
    if (by) $('#cmpBy').value = by;
    if (chain != null) { $('#fChain').value = chain; renderAll(); }
    const yB = C.y, yA = state.years.includes(C.y - 1) ? C.y - 1 : state.years[0];
    if ([...$('#cmpYearB').options].some((o) => +o.value === yB)) { $('#cmpYearB').value = yB; $('#cmpYearA').value = yA; }
    setView('comparativa');
  }

  function renderResumen(F, R) {
    const C = F.C, Y = C.y, P = Y - 1;
    const cur = R.byYear.get(Y), prev = R.byYear.get(P);
    $$('[data-sy]').forEach((e) => (e.textContent = Y));
    $$('[data-sp]').forEach((e) => (e.textContent = P));
    const cutTxt = C.d ? `${C.d} de ${C.monthLabel.toLowerCase()} de ${Y}` : `${C.monthLabel.toLowerCase()} de ${Y}`;
    $('#sumCutText').textContent = cutTxt;

    // Titular
    const p = prev ? varPct(cur.ytd, prev.ytd) : null;
    const cls = p == null ? '' : p >= 0 ? 'up' : 'down';
    $('#sumHeadline').innerHTML = p == null
      ? `En ${Y} vamos <span class="amount">${moneyShort(cur.ytd)}</span> de venta`
      : `${Y} va <span class="amount">${moneyShort(cur.ytd)}</span>, <span class="delta ${cls}">${(p > 0 ? '+' : '') + (p * 100).toFixed(1)}%</span> <span class="vs">frente al mismo corte de ${P}</span>`;
    const filt = [F.cob && F.cob !== 'all' && state.hasCob ? COB_LABEL[F.cob].toLowerCase() : null, F.tipo ? TIPO_LABEL[F.tipo].toLowerCase() : null, F.chain ? `cadena ${F.chain}` : null, F.brand ? `marca ${F.brand}` : null].filter(Boolean);
    $('#sumSub').textContent = `Acumulado de ${C.ytdLabel(Y).replace('Sem', 'la semana').replace('–', ' a la ')} en cada año, para comparar en igualdad de condiciones${filt.length ? '. Filtro: ' + filt.join(', ') : ''}.`;

    // Mini barras de años en el hero
    const ys = F.years;
    const maxY = Math.max(1, ...ys.map((y) => R.byYear.get(y).ytd));
    $('#sumYears').innerHTML = ys.map((y) => {
      const v = R.byYear.get(y).ytd;
      return `<div class="sy-col ${y === Y ? 'now' : ''}" title="${y}: ${money(v)}"><span class="sy-val">${moneyShort(v)}</span><div class="sy-bar" style="height:${Math.max(3, (v / maxY) * 100)}%"></div><span class="sy-year">${y}</span></div>`;
    }).join('');

    // Indicadores
    const foot = (a, b) => { const x = varPct(b, a); return x == null ? `sin dato de ${P}` : `<span class="${dirCls(x)}">${x > 0 ? '▲' : x < 0 ? '▼' : ''} ${pct(x)}</span> vs ${P}`; };
    const shareOf = (Yo) => (Yo && Yo.p.ytd + Yo.c.ytd > 0 ? Yo.p.ytd / (Yo.p.ytd + Yo.c.ytd) : null);
    const sh = shareOf(cur), shP = shareOf(prev);
    const kpis = [
      [`Acumulado ${Y}`, moneyShort(cur.ytd), prev ? foot(prev.ytd, cur.ytd) : ''],
      [`${C.monthLabel} ${Y} (al corte)`, moneyShort(cur.month), prev ? foot(prev.month, cur.month) : ''],
      C.hasW ? [`Semana ${C.w}`, moneyShort(cur.week), prev && state.weekYears.has(P) ? foot(prev.week, cur.week) : ''] : [`Año completo ${P}`, prev ? moneyShort(prev.full) : '—', ''],
      ['Participación Paraíso', sh == null ? '—' : fmtPct.format(sh), sh != null && shP != null ? `<span class="${dirCls(sh - shP)}">${ppTxt(sh - shP)}</span> vs ${P}` : ''],
      ['Tiendas con venta', num(cur.stores.size), prev ? `${num(prev.stores.size)} en ${P}` : ''],
    ];
    $('#sumKpis').innerHTML = kpis.map(([l, v, f]) => `<div class="hk"><div class="hk-label">${l}</div><div class="hk-value">${v}</div><div class="hk-foot">${f}</div></div>`).join('');

    renderSumMonths(F, R);
    renderSumShare(F, R);
    renderSumRanks(F, R);
    renderSumDonut(F, R, sh, shP);
    renderSumWeeks(F, R);
    renderSumHeat(F);
  }

  function renderSumMonths(F, R) {
    const C = F.C, Y = C.y, P = Y - 1;
    const has = (y, m) => R.byYM.has(y * 100 + m);
    const vals = (y, last) => MESES.map((_, i) => (i + 1 > last ? null : R.byYM.get(y * 100 + i + 1) || 0));
    const cur = vals(Y, C.m), prev = state.years.includes(P) ? vals(P, 12) : null;
    const partial = !!(C.d && C.d < new Date(Date.UTC(Y, C.m, 0)).getUTCDate());
    $('#sumMonthNote').textContent = partial ? ` ${C.monthLabel} de ${Y} va al día ${C.d} (línea punteada).` : '';
    let best = null;
    if (prev) for (let m = 1; m <= (partial ? C.m - 1 : C.m); m++) { const x = varPct(cur[m - 1], prev[m - 1]); if (x != null && has(Y, m) && (!best || x > best.x)) best = { m, x }; }
    $('#sumMonthTitle').textContent = best ? `${MESES_LARGO[best.m - 1]} fue el mejor mes frente a ${P} (${pct(best.x)})` : 'Mes a mes';
    upsertChart('sm', 'chartSumMonths', {
      type: 'line',
      data: { labels: MESES, datasets: [
        ...(prev ? [{ label: String(P), data: prev, borderColor: yearColor(P), backgroundColor: yearColor(P), borderWidth: 2, borderDash: [6, 5], pointRadius: 3, tension: 0.3 }] : []),
        { label: String(Y), data: cur, borderColor: yearColor(Y), backgroundColor: 'rgba(61,107,255,.12)', fill: prev ? '-1' : 'origin', borderWidth: 3, pointRadius: 4, pointHoverRadius: 7, tension: 0.3,
          segment: { borderDash: (c) => (partial && c.p1DataIndex === C.m - 1 ? [5, 5] : undefined) },
          pointBackgroundColor: cur.map((_, i) => (partial && i === C.m - 1 ? '#fff' : yearColor(Y))) },
      ] },
      options: {
        responsive: true, maintainAspectRatio: false, animation: { duration: 300 },
        interaction: { mode: 'index', intersect: false },
        onClick: (_e, els) => { if (els.length) goDetail({ months: [els[0].index + 1], by: 'cadena' }); },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
        plugins: {
          legend: { display: true, position: 'top', align: 'end', labels: { usePointStyle: true, boxWidth: 8, font: { weight: 600 } } },
          tooltip: { callbacks: { label: (c) => ` ${c.dataset.label}: ${money(c.parsed.y)}${partial && c.dataset.label === String(Y) && c.dataIndex === C.m - 1 ? ` (al día ${C.d})` : ''}`, footer: (it) => { const i = it[0].dataIndex; if (partial && i === C.m - 1) return 'Mes en curso: no se compara con el mes completo'; const x = prev ? varPct(cur[i], prev[i]) : null; return x == null || cur[i] == null ? '' : `Var: ${pct(x)}`; } } },
        },
        scales: { x: { grid: { display: false } }, y: { beginAtZero: true, grid: { color: '#EEF0F6' }, border: { display: false }, ticks: { callback: (v) => moneyShort(v) } } },
      },
    });
  }

  function renderSumShare(F, R) {
    const C = F.C, Y = C.y, P = Y - 1;
    const chains = [...R.byChain.values()].map((o) => ({ c: o.cadena, b: o.y[Y]?.ytd || 0, a: o.y[P]?.ytd || 0 })).filter((x) => x.a || x.b);
    const TB = chains.reduce((s, x) => s + x.b, 0), TA = chains.reduce((s, x) => s + x.a, 0);
    chains.sort((x, y) => y.b - x.b);
    const top = chains.slice(0, 8), rest = chains.slice(8);
    const items = [...top.map((x) => ({ ...x, color: chainColor(x.c) }))];
    if (rest.length) items.push({ c: `Otras (${rest.length})`, b: rest.reduce((s, x) => s + x.b, 0), a: rest.reduce((s, x) => s + x.a, 0), color: '#CBD5E1', other: true });
    const lead = items[0];
    $('#sumShareTitle').textContent = lead && TB ? `${lead.c} concentra el ${fmtPct.format(lead.b / TB)} de la venta` : 'Cómo se reparte la venta';
    const bar = (field, tot) => items.map((x) => {
      const w = tot ? (x[field] / tot) * 100 : 0;
      return w <= 0 ? '' : `<button class="share-seg" style="width:${w}%;background:${x.color}" ${x.other ? 'disabled' : `data-chain="${esc(x.c)}"`} title="${esc(x.c)}: ${fmtPct.format(x[field] / tot)} · ${money(x[field])}" aria-label="${esc(x.c)} ${fmtPct.format(x[field] / tot)}"></button>`;
    }).join('');
    const rows = [];
    if (TA) rows.push(`<div class="share-row"><b>${P}</b><div class="share-bar">${bar('a', TA)}</div></div>`);
    rows.push(`<div class="share-row"><b>${Y}</b><div class="share-bar">${bar('b', TB)}</div></div>`);
    const list = items.map((x) => {
      const sb = TB ? x.b / TB : 0, sa = TA ? x.a / TA : null;
      const d = sa == null ? null : sb - sa, vp = varPct(x.b, x.a);
      return `<li><span class="sw" style="background:${x.color}"></span>${x.other ? `<span class="nm">${esc(x.c)}</span>` : `<button class="nm" data-chain="${esc(x.c)}">${esc(x.c)}</button>`}<span class="sv">${moneyShort(x.b)}</span><span class="vp ${dirCls(vp)}">${vp == null ? 'nueva' : pct(vp)}</span><span class="sh">${fmtPct.format(sb)}</span><span class="pp ${dirCls(d)}">${ppTxt(d)}</span></li>`;
    }).join('');
    $('#sumShare').innerHTML = rows.join('') + `<div class="share-head"><span></span><span>Cadena</span><span>Venta ${Y}</span><span>Var %</span><span>Part.</span><span>vs ${P}</span></div><ul class="share-list">${list}</ul>`;
  }

  function renderSumRanks(F, R) {
    const C = F.C, Y = C.y, P = Y - 1;
    const st = [...R.byStore.values()].map((o) => ({ o, b: o.y[Y]?.ytd || 0, a: o.y[P]?.ytd || 0 })).map((x) => ({ ...x, d: x.b - x.a })).filter((x) => x.a || x.b);
    const up = st.filter((x) => x.d > 0).sort((a, b) => b.d - a.d).slice(0, 8);
    const down = st.filter((x) => x.d < 0).sort((a, b) => a.d - b.d).slice(0, 8);
    const max = Math.max(1, ...up.map((x) => x.d), ...down.map((x) => -x.d));
    const li = (x, i) => `<li><button data-chain="${esc(x.o.cadena)}" title="Ver las tiendas de ${esc(x.o.cadena)}"><span class="fill" style="--w:${(Math.abs(x.d) / max) * 100}%"></span><span class="pos">${i + 1}</span><span class="who"><b>${esc(x.o.tienda)}</b><small>${esc(x.o.cadena)}</small></span><span class="amt">${x.d > 0 ? '+' : ''}${moneyShort(x.d)}<small>${x.a ? pct(varPct(x.b, x.a)) : 'nueva'}</small></span></button></li>`;
    $('#sumUp').innerHTML = up.map(li).join('') || '<li class="empty">Ninguna tienda crece con estos filtros.</li>';
    $('#sumDown').innerHTML = down.map(li).join('') || '<li class="empty">Ninguna tienda cae con estos filtros.</li>';
  }

  function renderSumDonut(F, R, sh, shP) {
    const card = $('#sumShareParaisoCard');
    const Yo = R.byYear.get(F.C.y), Po = R.byYear.get(F.C.y - 1);
    const show = state.hasBrand && !F.tipo && !F.brand && Yo.p.ytd + Yo.c.ytd > 0;
    card.hidden = !show;
    if (!show) return;
    $('#sumParTitle').textContent = shP == null ? `Paraíso tiene el ${fmtPct.format(sh)} de la venta` : sh >= shP ? `Paraíso gana ${ppTxt(sh - shP).replace('+', '')} de participación` : `Paraíso pierde ${ppTxt(sh - shP).replace('-', '')} de participación`;
    $('#sumDonutCenter').innerHTML = `<b>${fmtPct.format(sh)}</b><span>Paraíso</span>`;
    const pv = (a, b) => (a ? pct(varPct(b, a)) : '—');
    $('#sumDonutLegend').innerHTML = `
      <li><span class="sw" style="background:${COL_PARAISO}"></span><div><b>Venta Paraíso · ${moneyShort(Yo.p.ytd)}</b><small>${Po ? `${pv(Po.p.ytd, Yo.p.ytd)} vs ${F.C.y - 1}` : ''} · PARAISO, VIVA-HOME, CASA JOVEN, KREA, KOMFORT</small></div></li>
      <li><span class="sw" style="background:${COL_COMP}"></span><div><b>Venta Competencia · ${moneyShort(Yo.c.ytd)}</b><small>${Po ? `${pv(Po.c.ytd, Yo.c.ytd)} vs ${F.C.y - 1}` : ''} · todas las demás marcas</small></div></li>`;
    upsertChart('sd', 'chartSumDonut', {
      type: 'doughnut',
      data: { labels: ['Venta Paraíso', 'Venta Competencia'], datasets: [{ data: [Yo.p.ytd, Yo.c.ytd], backgroundColor: [COL_PARAISO, COL_COMP], borderColor: '#fff', borderWidth: 3, hoverOffset: 6 }] },
      options: { responsive: true, maintainAspectRatio: false, cutout: '72%', animation: { duration: 300 }, plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => ` ${c.label}: ${money(c.parsed)}` } } } },
    });
  }

  function renderSumWeeks(F, R) {
    const C = F.C, Y = C.y, P = Y - 1;
    const card = $('#chartSumWeeks').closest('.panel');
    const ok = C.hasW && state.weekYears.has(Y) && state.weekYears.has(P);
    card.hidden = !ok;
    if (!ok) return;
    const weeks = []; for (let w = Math.max(1, C.w - 7); w <= C.w; w++) weeks.push(w);
    const v = weeks.map((w) => varPct(R.byYW.get(Y * 100 + w) || 0, R.byYW.get(P * 100 + w) || 0));
    const last4 = weeks.slice(-4);
    const s4b = last4.reduce((s, w) => s + (R.byYW.get(Y * 100 + w) || 0), 0), s4a = last4.reduce((s, w) => s + (R.byYW.get(P * 100 + w) || 0), 0);
    const p4 = varPct(s4b, s4a);
    $('#sumWeeksTitle').textContent = p4 == null ? 'Últimas semanas' : `Las últimas 4 semanas van ${pct(p4)} frente a ${P}`;
    const m = Math.max(0.05, ...v.filter((x) => x != null).map(Math.abs)) * 1.2;
    upsertChart('sw', 'chartSumWeeks', {
      type: 'bar',
      data: { labels: weeks.map((w) => `Sem ${w}`), datasets: [{ data: v.map((x) => (x == null ? null : x * 100)), backgroundColor: v.map((x) => (x == null ? '#E2E6F0' : x >= 0 ? '#14B88A' : '#F0475A')), borderRadius: 6, borderSkipped: false, barPercentage: 0.62 }] },
      options: {
        responsive: true, maintainAspectRatio: false, animation: { duration: 300 },
        onClick: (_e, els) => { if (els.length) goDetail({ weeks: [weeks[els[0].index]], by: 'cadena' }); },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
        layout: { padding: { top: 18 } },
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => ` ${pct(v[c.dataIndex])}` } }, wkLabels: { v } },
        scales: { x: { grid: { display: false } }, y: { min: -m * 100, max: m * 100, grid: { color: (c) => (c.tick.value === 0 ? '#161A3A' : '#EEF0F6') }, border: { display: false }, ticks: { callback: (x) => (x > 0 ? '+' : '') + Math.round(x) + '%' } } },
      },
      plugins: [{
        id: 'wkLabels',
        afterDatasetsDraw(chart, _a, o) {
          const { ctx } = chart; ctx.save();
          ctx.font = '700 11.5px "Plus Jakarta Sans", sans-serif'; ctx.textAlign = 'center';
          chart.getDatasetMeta(0).data.forEach((bar, i) => { const x = o.v[i]; if (x == null) return; ctx.fillStyle = x >= 0 ? '#0E8F6B' : '#D42E43'; ctx.fillText(pct(x), bar.x, x >= 0 ? bar.y - 6 : bar.y + 14); });
          ctx.restore();
        },
      }],
    });
  }

  // Calendario de calor: cada día del año actual vs el mismo día de la semana del año anterior (364 días antes).
  async function renderSumHeat(F) {
    const box = $('#sumHeat');
    let data;
    try { data = await loadCmp(); } catch { box.innerHTML = '<p class="muted">No se pudo cargar el detalle diario.</p>'; return; }
    const C = F.C, Y = C.y, P = Y - 1;
    const cobOk = (r) => F.cob === 'all' || !state.hasCob || (F.cob === 'si' ? r.cob === 1 : F.cob === 'no' ? r.cob === 0 : r.cob == null);
    const day = new Map();
    for (const r of data) {
      if ((r.y !== Y && r.y !== P) || !(r.d > 0) || !cobOk(r) || (F.tipo && r.tipo !== F.tipo) || (F.chain && r.cadena !== F.chain) || (F.brand && r.marca !== F.brand)) continue;
      const k = Date.UTC(r.y, r.m - 1, r.d);
      day.set(k, (day.get(k) || 0) + r.v);
    }
    if (!day.size) { box.closest('.panel').hidden = true; return; }
    box.closest('.panel').hidden = false;
    const start = Date.UTC(Y, 0, 1), end = Date.UTC(Y, C.m - 1, C.d || new Date(Date.UTC(Y, C.m, 0)).getUTCDate());
    const color = (x) => {
      if (x == null) return '#E6E8F0';
      const t = Math.min(1, Math.abs(x) / 0.4);
      const mix = (a, b) => a.map((v, i) => Math.round(v + (b[i] - v) * t));
      const base = [230, 232, 240];
      const c = x >= 0 ? mix(base, [20, 184, 138]) : mix(base, [240, 71, 90]);
      return `rgb(${c.join(',')})`;
    };
    const wd0 = (new Date(start).getUTCDay() + 6) % 7;
    let html = '', col = 0, better = 0, counted = 0;
    const cells = [];
    for (let i = 0; i < wd0; i++) cells.push('<span class="hx"></span>');
    const monthAt = {};
    for (let t = start, idx = wd0; t <= end; t += 864e5, idx++) {
      const d = new Date(t), b = day.get(t) || 0, a = day.get(t - 364 * 864e5);
      const x = a ? varPct(b, a) : null;
      if (x != null) { counted++; if (x > 0) better++; }
      const c = Math.floor(idx / 7);
      if (d.getUTCDate() === 1) monthAt[c] = MESES[d.getUTCMonth()];
      const label = `${d.getUTCDate()} ${MESES[d.getUTCMonth()].toLowerCase()}`;
      cells.push(b || a
        ? `<button class="hd" style="--c:${color(x)}" data-t="${t}" data-b="${b}" data-a="${a ?? ''}" aria-label="${label}: ${money(b)}${x == null ? '' : ', ' + pct(x)}"></button>`
        : `<span class="hd none" title="${label}: sin venta"></span>`);
      col = c;
    }
    // Columnas: fila 1 = mes, filas 2..8 = lunes..domingo (con una primera columna de iniciales)
    html += '<span class="hm"></span>' + ['L', 'M', 'X', 'J', 'V', 'S', 'D'].map((d) => `<span class="hw">${d}</span>`).join('');
    for (let c = 0; c <= col; c++) {
      html += `<span class="hm">${monthAt[c] || ''}</span>`;
      html += cells.slice(c * 7, c * 7 + 7).join('') + '<span class="hx"></span>'.repeat(Math.max(0, 7 - cells.slice(c * 7, c * 7 + 7).length));
    }
    box.innerHTML = html;
    $('#sumHeatTitle').textContent = counted ? `${better} de ${counted} días vendieron más que el mismo día de ${P}` : 'El año día por día';
  }

  // ---------------------------------------------------------------- comparativa (segmentadores Mes / Semana / Día)
  const cmp = { data: null, loading: null, months: null, weeks: new Set(), days: new Set(), init: false };

  async function loadCmp() {
    if (cmp.data) return cmp.data;
    if (!cmp.loading) {
      cmp.loading = api('/api/dias').then((D) => {
        cmp.data = D.rows.map(([ci, mi, cob, y, m, w, d, v, u]) => ({ cadena: D.dict.cadenas[ci], marca: D.dict.marcas[mi], tipo: tipoDe(D.dict.marcas[mi]), cob: cob ?? null, y, m, w, d, v, u }));
        return cmp.data;
      }).finally(() => { cmp.loading = null; });
    }
    return cmp.loading;
  }

  async function renderCmp() {
    if (state.view !== 'comparativa' || !lastF) return;
    $('#cmpStatus').hidden = false;
    $('#cmpStatus').textContent = 'Cargando el detalle diario…';
    let data;
    try { data = await loadCmp(); } catch (e) { $('#cmpStatus').textContent = e.message; return; }
    $('#cmpStatus').hidden = true;
    const F = lastF, C = F.C;
    $('#cmpEmpty').hidden = !!data.length;
    $('#cmpContent').hidden = !data.length;
    if (!data.length) return;

    // Años a comparar (por defecto el año del corte contra el anterior).
    const years = [...new Set(data.map((r) => r.y))].sort((a, b) => a - b);
    const selB = $('#cmpYearB'), selA = $('#cmpYearA');
    const yearsKey = years.join(',');
    if (cmp.yearsKey !== yearsKey) {
      const keepB = cmp.init && years.includes(+selB.value) ? +selB.value : null, keepA = cmp.init && years.includes(+selA.value) ? +selA.value : null;
      selB.innerHTML = years.map((y) => `<option>${y}</option>`).join(''); selB.value = keepB ?? (years.includes(C.y) ? C.y : years[years.length - 1]);
      selA.innerHTML = years.map((y) => `<option>${y}</option>`).join(''); selA.value = keepA ?? (years.includes(+selB.value - 1) ? +selB.value - 1 : years[0]);
      cmp.yearsKey = yearsKey;
    }
    if (!cmp.init) {
      // Por defecto: el mes de corte, hasta el día de corte (comparación pareja con el año anterior).
      cmp.months = new Set([C.m]);
      cmp.days = new Set(C.d ? Array.from({ length: C.d }, (_, i) => i + 1) : []);
      cmp.init = true;
    }
    const yB = +selB.value, yA = +selA.value;

    // Filtros generales (cobertura, tipo de venta, cadena, marca).
    const cobOk = (r) => F.cob === 'all' || !state.hasCob || (F.cob === 'si' ? r.cob === 1 : F.cob === 'no' ? r.cob === 0 : r.cob == null);
    const baseAll = data.filter((r) => cobOk(r) && (!F.tipo || r.tipo === F.tipo) && (!F.chain || r.cadena === F.chain) && (!F.brand || r.marca === F.brand));
    const base = baseAll.filter((r) => r.y === yA || r.y === yB);

    // Segmentadores: opciones disponibles en cascada (Mes → Semana → Día), como en Excel.
    const monthsAvail = new Set(base.map((r) => r.m));
    const inM = (r) => !cmp.months.size || cmp.months.has(r.m);
    const weeksAvail = new Set(base.filter((r) => inM(r) && r.w > 0).map((r) => r.w));
    const inW = (r) => !cmp.weeks.size || cmp.weeks.has(r.w);
    const daysAvail = new Set(base.filter((r) => inM(r) && inW(r) && r.d > 0).map((r) => r.d));
    const chips = (el, values, sel, label, avail) => {
      el.innerHTML = values.map((v) => `<button class="chip ${sel.has(v) ? 'on' : ''} ${avail.has(v) ? '' : 'off'}" data-v="${v}">${label(v)}</button>`).join('');
    };
    chips($('#slMes'), Array.from({ length: 12 }, (_, i) => i + 1).filter((m) => monthsAvail.has(m) || cmp.months.has(m)), cmp.months, (m) => MESES_LARGO[m - 1], monthsAvail);
    chips($('#slSem'), [...weeksAvail, ...cmp.weeks].filter((v, i, a) => a.indexOf(v) === i).sort((a, b) => a - b), cmp.weeks, (w) => `Sem-${w}`, weeksAvail);
    chips($('#slDia'), Array.from({ length: 31 }, (_, i) => i + 1), cmp.days, (d) => d, daysAvail);
    const noDay = !base.some((r) => r.d > 0);
    $('#slDiaWrap').hidden = noDay; $('#slSemWrap').hidden = !base.some((r) => r.w > 0);

    const inD = (r) => !cmp.days.size || cmp.days.has(r.d);
    const sel = base.filter((r) => inM(r) && inW(r) && inD(r));
    {
      const tot = { a: 0, b: 0, pa: 0, pb: 0 };
      for (const r of sel) { if (r.tipo === 'sinmarca') continue; if (r.y === yA) { tot.a += r.v; if (r.tipo === 'paraiso') tot.pa += r.v; } else { tot.b += r.v; if (r.tipo === 'paraiso') tot.pb += r.v; } }
      cmp.shareP = { a: tot.a ? tot.pa / tot.a : null, b: tot.b ? tot.pb / tot.b : null };
    }
    const by = $('#cmpBy').value;
    const LABELS = { cadena: ['Cadena'], marca: ['Marca'], tipo: ['Tipo de venta'], tienda: ['Cadena', 'Tienda', 'Coordinador'], coordinador: ['Coordinador'] };

    // Filas de trabajo {cols, y, w, v, u}: cadena/marca/grupo salen del detalle diario ya cargado;
    // tienda y coordinador se piden al servidor (son muchas filas para traerlas todas).
    let work;
    if (by === 'tienda' || by === 'coordinador') {
      const marcas = F.tipo || F.brand ? [...new Set(data.map((r) => r.marca))].filter((m) => (!F.tipo || tipoDe(m) === F.tipo) && (!F.brand || m === F.brand)) : null;
      const body = { anios: [yA, yB], meses: [...cmp.months], semanas: [...cmp.weeks], dias: [...cmp.days], cob: state.hasCob ? F.cob : 'all', marcas, cadena: F.chain || '', por: by };
      const key = JSON.stringify(body);
      const token = (cmp.token = (cmp.token || 0) + 1);
      let resp = cmp.cache?.key === key ? cmp.cache.resp : null;
      if (!resp) {
        $('#cmpStatus').hidden = false; $('#cmpStatus').textContent = `Calculando la comparativa por ${by}…`;
        try { resp = await api('/api/comparativa', { method: 'POST', body }); } catch (e) { $('#cmpStatus').textContent = e.message; return; }
        if (token !== cmp.token) return; // llegó una consulta más nueva
        cmp.cache = { key, resp };
        $('#cmpStatus').hidden = true;
      }
      $('#cmpNoStore').hidden = resp.disponible;
      if (by === 'tienda') {
        const coordOf = new Map();
        for (const r of resp.rows) { const k = r.cadena + '|' + r.tienda; if (r.coordinador && (!coordOf.has(k) || r.anio === yB)) coordOf.set(k, r.coordinador); }
        work = resp.rows.map((r) => ({ cols: [r.cadena, r.tienda, coordOf.get(r.cadena + '|' + r.tienda) || '—'], y: r.anio, w: r.semana, v: r.v, u: r.u }));
      } else {
        work = resp.rows.map((r) => ({ cols: [r.coordinador || '(Sin coordinador)'], y: r.anio, w: r.semana, v: r.v, u: r.u }));
      }
    } else {
      $('#cmpNoStore').hidden = true;
      const keyOf = (r) => (by === 'marca' ? r.marca || '(Sin marca)' : by === 'tipo' ? TIPO_LABEL[r.tipo] : r.cadena);
      work = sel.map((r) => ({ cols: [keyOf(r)], y: r.y, w: r.w, v: r.v, u: r.u }));
    }

    const groups = new Map(), wk = new Map();
    for (const r of work) {
      const k = r.cols.join('\u0001');
      const g = groups.get(k) || { cols: r.cols, vA: 0, vB: 0, uA: 0, uB: 0 };
      if (r.y === yA) { g.vA += r.v; g.uA += r.u; } else { g.vB += r.v; g.uB += r.u; }
      groups.set(k, g);
      if (r.w > 0) {
        const gw = wk.get(k) || { cols: r.cols, w: {} };
        const c = gw.w[r.w] || (gw.w[r.w] = { vA: 0, vB: 0, uA: 0, uB: 0 });
        if (r.y === yA) { c.vA += r.v; c.uA += r.u; } else { c.vB += r.v; c.uB += r.u; }
        wk.set(k, gw);
      }
    }
    const byCols = (a, b) => a.cols.join(' ').localeCompare(b.cols.join(' '), 'es');
    const list = [...groups.values()].sort(byCols);
    cmp.last = { yA, yB, list, by, labels: LABELS[by] };
    const hasUnits = work.some((r) => r.u);
    renderTable('tblCmpSoles', cmpAOA('v'));
    $('#cmpUndCard').hidden = !hasUnits;
    if (hasUnits) renderTable('tblCmpUnd', cmpAOA('u'));
    $('#cmpByTitle').textContent = LABELS[by][by === 'tienda' ? 1 : 0].toLowerCase();

    // Detalle por semana: columnas Sem-N × año, como la tabla dinámica de Excel.
    const weeks = (cmp.weeks.size ? [...cmp.weeks] : [...new Set(work.filter((r) => r.w > 0).map((r) => r.w))]).sort((a, b) => a - b);
    cmp.wk = { weeks, list: [...wk.values()].sort(byCols), yA, yB, by, labels: LABELS[by] };
    renderAnalysis(hasUnits);
    $('#cmpWeekCard').hidden = !weeks.length;
    $('#cmpWeekUndCard').hidden = !weeks.length || !hasUnits;
    if (weeks.length) {
      renderWeekTable('tblCmpWeek', 'v');
      if (hasUnits) renderWeekTable('tblCmpWeekUnd', 'u');
      // Gráfico: barras por semana para los últimos 3 años hasta el año comparado.
      const chartYears = [...new Set(baseAll.map((r) => r.y))].filter((y) => y <= yB).sort((a, b) => a - b).slice(-3);
      if (!chartYears.includes(yA)) chartYears.unshift(yA);
      const inMD = (r) => inM(r) && inD(r);
      const sumBy = (field) => { const m = new Map(); for (const r of baseAll) { if (!(r.w > 0) || !weeks.includes(r.w) || !inMD(r)) continue; const k = r.y * 100 + r.w; m.set(k, (m.get(k) || 0) + r[field]); } return m; };
      renderWeekChart('chartCmpWeek', 'cw', weeks, chartYears, sumBy('v'), yA, yB, money);
      if (hasUnits) renderWeekChart('chartCmpWeekUnd', 'cwu', weeks, chartYears, sumBy('u'), yA, yB, num);
    }

    applyUnitPanes(hasUnits, weeks.length > 0);
    const ms = [...cmp.months].sort((a, b) => a - b).map((m) => MESES_LARGO[m - 1]);
    $('#cmpTitle').textContent = cmp.months.size === 1 && [...cmp.months][0] === C.m && yB === C.y ? 'mes actual' : (ms.length ? ms.join(' y ').toLowerCase() : 'año completo');
    const dsel = [...cmp.days].sort((a, b) => a - b);
    const dTxt = !dsel.length ? 'todos los días' : dsel.length > 2 && dsel[dsel.length - 1] - dsel[0] + 1 === dsel.length ? `días ${dsel[0]} al ${dsel[dsel.length - 1]}` : `días ${dsel.join(', ')}`;
    $('#cmpNote').innerHTML = `Comparando <b>${yB}</b> contra <b>${yA}</b> · ${ms.length ? ms.join(', ') : 'todos los meses'} · ${cmp.weeks.size ? 'semanas ' + [...cmp.weeks].sort((a, b) => a - b).join(', ') : 'todas las semanas'} · ${dTxt}${noDay ? ' · <span class="down">los archivos no traen fecha diaria</span>' : ''}`;
  }

  const unitSel = { detail: 'v', weeks: 'v' };
  function applyUnitPanes(hasUnits, hasWeeks) {
    for (const seg of $$('.seg[data-seg]')) {
      const k = seg.dataset.seg;
      if (!hasUnits) unitSel[k] = 'v';
      seg.querySelectorAll('button').forEach((b) => { b.classList.toggle('on', b.dataset.unit === unitSel[k]); b.disabled = b.dataset.unit === 'u' && !hasUnits; b.setAttribute('aria-selected', String(b.dataset.unit === unitSel[k])); });
    }
    $('#cmpSolesCard').hidden = unitSel.detail !== 'v';
    $('#cmpUndCard').hidden = unitSel.detail !== 'u';
    $('#cmpNoWeeks').hidden = hasWeeks;
    $('#cmpWeekCard').hidden = !hasWeeks || unitSel.weeks !== 'v';
    $('#cmpWeekUndCard').hidden = !hasWeeks || unitSel.weeks !== 'u';
    $$('.seg[data-seg="weeks"]').forEach((s) => (s.hidden = !hasWeeks));
  }

  function weekAOA(field) {
    const { weeks, list, yA, yB } = cmp.wk;
    const unit = field === 'v' ? 'Soles' : 'Und';
    const A = field === 'v' ? 'vA' : 'uA', B = field === 'v' ? 'vB' : 'uB';
    const t = field === 'v' ? '$' : 'n';
    const head = [...cmp.wk.labels];
    const types = head.map(() => 't');
    for (const w of weeks) { head.push(`Sem-${w} ${yA} ${unit}`, `Sem-${w} ${yB} ${unit}`, `Sem-${w} Var %`); types.push(t, t, '%'); }
    head.push(`Total ${yA} ${unit}`, `Total ${yB} ${unit}`, 'Total Var %'); types.push(t, t, '%');
    const nl = cmp.wk.labels.length;
    const rowOf = (cols, cells) => {
      const r = [...cols]; let ta = 0, tb = 0;
      for (const w of weeks) { const c = cells(w); r.push(c.a, c.b, varPct(c.b, c.a)); ta += c.a; tb += c.b; }
      r.push(ta, tb, varPct(tb, ta));
      return r;
    };
    const rows = list.map((g) => rowOf(g.cols, (w) => ({ a: g.w[w]?.[A] || 0, b: g.w[w]?.[B] || 0 }))).filter((r) => r.slice(nl).some((v, i) => types[i + nl] !== '%' && v));
    const foot = [rowOf(['Total general', ...Array(nl - 1).fill('')], (w) => ({ a: list.reduce((s, g) => s + (g.w[w]?.[A] || 0), 0), b: list.reduce((s, g) => s + (g.w[w]?.[B] || 0), 0) }))];
    return { head, rows, foot, types, weeks, yA, yB, unit, nl };
  }

  function renderWeekTable(id, field) {
    const T = weekAOA(field);
    const groups = [...T.weeks.map((w) => `Sem-${w}`), 'Total'];
    const th1 = `<tr>${T.head.slice(0, T.nl).map((h) => `<th rowspan="2">${esc(h)}</th>`).join('')}${groups.map((g) => `<th colspan="3" class="grp">${g}</th>`).join('')}</tr>`;
    const th2 = `<tr>${groups.map(() => `<th class="num">${T.yA}<br>${T.unit}</th><th class="num">${T.yB}<br>${T.unit}</th><th class="num">Var%</th>`).join('')}</tr>`;
    const row = (r) => `<tr>${r.map((v, i) => (i < T.nl ? `<td>${esc(v)}</td>` : `<td class="num ${T.types[i] === '%' ? dirCls(v) : ''} ${(i - T.nl) % 3 === 0 ? 'sep' : ''}">${cell(v, T.types[i])}</td>`)).join('')}</tr>`;
    $('#' + id).innerHTML = `<thead>${th1}${th2}</thead><tbody>${T.rows.map(row).join('')}</tbody><tfoot>${T.foot.map(row).join('')}</tfoot>`;
  }

  // Etiquetas: valor vertical dentro de cada barra y Var % (último año vs anterior) sobre cada semana.
  const barLabels = {
    id: 'barLabels',
    afterDatasetsDraw(chart, _a, opts) {
      const { ctx } = chart;
      ctx.save();
      chart.data.datasets.forEach((ds, di) => {
        chart.getDatasetMeta(di).data.forEach((bar, i) => {
          const v = ds.data[i];
          if (!v || bar.height < 46) return;
          ctx.save();
          ctx.translate(bar.x, bar.base - 6);
          ctx.rotate(-Math.PI / 2);
          ctx.fillStyle = '#fff'; ctx.font = '600 11px Segoe UI, system-ui, sans-serif'; ctx.textBaseline = 'middle';
          ctx.fillText(opts.fmt(v), 0, 0);
          ctx.restore();
        });
      });
      (opts.vars || []).forEach((p, i) => {
        if (p == null) return;
        const metas = chart.data.datasets.map((_, di) => chart.getDatasetMeta(di).data[i]).filter(Boolean);
        const top = Math.min(...metas.map((m) => m.y));
        const x = metas.reduce((s, m) => s + m.x, 0) / metas.length;
        const txt = (p > 0 ? '+' : '') + (p * 100).toFixed(1) + '%';
        ctx.font = '700 12px Segoe UI, system-ui, sans-serif';
        const w = ctx.measureText(txt).width + 12;
        ctx.fillStyle = p >= 0 ? '#0f7b3d' : '#c62828';
        ctx.beginPath(); ctx.roundRect(x - w / 2, top - 26, w, 20, 5); ctx.fill();
        ctx.fillStyle = '#fff'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        ctx.fillText(txt, x, top - 16);
      });
      ctx.restore();
    },
  };

  function renderWeekChart(canvasId, key, weeks, years, sums, yA, yB, fmt) {
    const legend = years.map((y) => `<span><i class="swatch" style="background:${yearColor(y)}"></i>${y}</span>`).join('') + `<span><i class="swatch" style="background:#0f7b3d"></i>Var % ${yB} vs ${yA}</span>`;
    $('#' + canvasId).closest('.panel, .card').querySelector('[data-legend]').innerHTML = legend;
    const vars = weeks.map((w) => varPct(sums.get(yB * 100 + w) || 0, sums.get(yA * 100 + w) || 0));
    const short = (v) => (fmt === money ? 'S/ ' + fmtNum0.format(v) : fmtNum0.format(v));
    upsertChart(key, canvasId, {
      type: 'bar',
      data: { labels: weeks.map((w) => `Sem-${w}`), datasets: years.map((y) => ({ label: String(y), data: weeks.map((w) => sums.get(y * 100 + w) || 0), backgroundColor: yearColor(y), borderRadius: 4, borderSkipped: 'bottom', categoryPercentage: 0.75, barPercentage: 0.9 })) },
      options: {
        responsive: true, maintainAspectRatio: false, animation: { duration: 250 },
        layout: { padding: { top: 30 } },
        interaction: { mode: 'index', intersect: false },
        plugins: { legend: { display: false }, barLabels: { fmt: short, vars }, tooltip: { callbacks: { label: (c) => ` ${c.dataset.label}: ${fmt(c.parsed.y)}`, footer: (it) => { const p = vars[it[0].dataIndex]; return p == null ? '' : `Var % ${yB} vs ${yA}: ${pct(p)}`; } } } },
        scales: { x: { grid: { display: false }, ticks: { color: TICK } }, y: { beginAtZero: true, grid: { color: GRID }, border: { display: false }, ticks: { color: TICK, callback: fmt === money ? moneyTick : (v) => fmtNum0.format(v) } } },
      },
      plugins: [barLabels],
    });
  }

  function cmpAOA(field) {
    const { yA, yB, list, labels } = cmp.last;
    const unit = field === 'v' ? 'Soles' : 'Und';
    const A = field === 'v' ? 'vA' : 'uA', B = field === 'v' ? 'vB' : 'uB';
    const t = field === 'v' ? '$' : 'n';
    const nl = labels.length;
    const head = [...labels, `${yA} ${unit}`, `${yB} ${unit}`, field === 'v' ? 'Var S/' : 'Var Und', 'Var %'];
    const rows = list.filter((g) => g[A] || g[B]).map((g) => [...g.cols, g[A], g[B], g[B] - g[A], varPct(g[B], g[A])]);
    const sa = rows.reduce((s, r) => s + r[nl], 0), sb = rows.reduce((s, r) => s + r[nl + 1], 0);
    const barMax = Math.max(0, ...rows.map((r) => Math.abs(r[nl + 3] ?? 0)).filter(Number.isFinite));
    return { head, rows, foot: [['Total general', ...Array(nl - 1).fill(''), sa, sb, sb - sa, varPct(sb, sa)]], types: [...labels.map(() => 't'), t, t, t, '%'], varCols: [nl + 2, nl + 3], bar: nl + 3, barMax };
  }

  // ---------------------------------------------------------------- análisis del Vs (texto automático)
  function renderAnalysis(hasUnits) {
    const { yA, yB, list, by, labels } = cmp.last;
    const name = (g) => (by === 'tienda' ? `${g.cols[1]} (${g.cols[0]})` : g.cols[0]);
    const what = { cadena: 'cadenas', marca: 'marcas', tipo: 'grupos', tienda: 'tiendas', coordinador: 'coordinadores' }[by];
    const A = list.reduce((s, g) => s + g.vA, 0), B = list.reduce((s, g) => s + g.vB, 0);
    const uA = list.reduce((s, g) => s + g.uA, 0), uB = list.reduce((s, g) => s + g.uB, 0);
    const out = [];
    if (!A && !B) { $('#cmpAnalysis').innerHTML = '<li>No hay ventas para la selección actual.</li>'; cmp.analysis = []; return; }
    const d = B - A, p = varPct(B, A);
    const sign = (x) => (x > 0 ? 'crece' : x < 0 ? 'cae' : 'se mantiene');
    out.push(`<b>Resultado general:</b> la venta ${sign(d)} ${p == null ? '' : `<span class="${dirCls(d)}">${pct(p)}</span>`} — de ${money(A)} en ${yA} a ${money(B)} en ${yB} (${signMoney(d)}).`);
    if (hasUnits && (uA || uB)) {
      const pu = varPct(uB, uA);
      const tA = uA ? A / uA : null, tB = uB ? B / uB : null;
      out.push(`<b>Unidades:</b> ${num(uA)} → ${num(uB)} (${pct(pu)}).` + (tA && tB ? ` El precio promedio por unidad pasa de ${money(tA)} a ${money(tB)} (${pct(varPct(tB, tA))}), así que ${Math.abs(varPct(tB, tA)) > Math.abs(pu ?? 0) ? 'la variación se explica más por precio/mix que por volumen' : 'la variación se explica más por volumen que por precio'}.` : ''));
    }
    const moved = list.map((g) => ({ g, d: g.vB - g.vA })).filter((x) => x.d);
    const up = moved.filter((x) => x.d > 0).sort((a, b) => b.d - a.d), down = moved.filter((x) => x.d < 0).sort((a, b) => a.d - b.d);
    const active = list.filter((g) => g.vA || g.vB).length;
    const pl = (n, one, many) => `${n} ${n === 1 ? one : many}`;
    out.push(`<b>${active} ${what}</b> con venta: ${pl(up.length, 'crece', 'crecen')} y ${pl(down.length, 'cae', 'caen')}.`);
    const share = (x) => {
      if (!d) return '';
      const r = Math.abs(x.d / d);
      return r > 1 ? `: por sí solo mueve más que la variación neta total (${signMoney(d)}), que se compensa con lo que va en sentido contrario` : ` el ${fmtPct.format(r)} de la variación neta`;
    };
    const fmtList = (arr) => arr.slice(0, 3).map((x) => `${esc(name(x.g))} ${signMoney(x.d)}${x.g.vA ? ` (${pct(varPct(x.g.vB, x.g.vA))})` : ' (nueva)'}`).join('; ');
    if (up.length) out.push(`<b>Lo que más empuja hacia arriba:</b> ${fmtList(up)}.` + (d > 0 ? ` ${esc(name(up[0].g))} explica${share(up[0])}.` : ''));
    if (down.length) out.push(`<b>Lo que más resta:</b> ${fmtList(down)}.` + (d < 0 ? ` ${esc(name(down[0].g))} explica${share(down[0])}.` : ''));
    if (by === 'tienda') {
      const nuevas = list.filter((g) => !g.vA && g.vB).length, sinVenta = list.filter((g) => g.vA && !g.vB).length;
      if (nuevas || sinVenta) out.push(`<b>Tiendas:</b> ${nuevas} con venta solo en ${yB} y ${sinVenta} que vendieron en ${yA} pero no en ${yB}.`);
    }
    if (list.length > 1 && B) {
      const top = [...list].sort((a, b) => b.vB - a.vB)[0];
      out.push(`<b>Mayor peso:</b> ${esc(name(top))} concentra el ${fmtPct.format(top.vB / B)} de la venta de ${yB}.`);
    }
    // Participación Paraíso en la misma selección (si se están viendo todas las marcas).
    if (state.hasBrand && !lastF.tipo && !lastF.brand && cmp.shareP) {
      const { a, b } = cmp.shareP;
      if (a != null && b != null) out.push(`<b>Participación Paraíso:</b> ${fmtPct.format(a)} → ${fmtPct.format(b)} (<span class="${dirCls(b - a)}">${(b - a > 0 ? '+' : '') + ((b - a) * 100).toFixed(1)} pp</span>).`);
    }
    if (cmp.wk?.weeks.length > 1) {
      const ws = cmp.wk.weeks.map((w) => { const a = cmp.wk.list.reduce((s, g) => s + (g.w[w]?.vA || 0), 0), b = cmp.wk.list.reduce((s, g) => s + (g.w[w]?.vB || 0), 0); return { w, p: varPct(b, a) }; }).filter((x) => x.p != null);
      if (ws.length > 1) {
        const best = ws.reduce((m, x) => (x.p > m.p ? x : m)), worst = ws.reduce((m, x) => (x.p < m.p ? x : m));
        out.push(`<b>Semanas:</b> la mejor fue Sem-${best.w} (${pct(best.p)}) y la más floja Sem-${worst.w} (${pct(worst.p)}).`);
      }
    }
    cmp.analysis = out.map((h) => h.replace(/<[^>]+>/g, ''));
    $('#cmpAnalysis').innerHTML = out.map((h) => `<li>${h}</li>`).join('');
    renderHero({ A, B, uA, uB, d, p, hasUnits });
    renderInsights({ up, down, list, B, d, uA, uB, A, hasUnits, name, what });
    renderBridge({ list, A, B, name, what });
    renderDiverge({ list, name });
  }

  // ---------------------------------------------------------------- presentación (hero, tarjetas, gráficos)
  const WHAT_ONE = { cadena: 'cadena', marca: 'marca', tipo: 'grupo', tienda: 'tienda', coordinador: 'coordinador' };

  // Frase natural del período: "Del 1 al 21 de noviembre de 2026", "En octubre y noviembre de 2026"…
  function periodLabel(year) {
    const ms = [...cmp.months].sort((a, b) => a - b).map((m) => MESES_LARGO[m - 1].toLowerCase());
    const ds = [...cmp.days].sort((a, b) => a - b);
    const mtxt = ms.length ? (ms.length === 1 ? ms[0] : ms.slice(0, -1).join(', ') + ' y ' + ms[ms.length - 1]) : null;
    const contiguous = ds.length > 1 && ds[ds.length - 1] - ds[0] + 1 === ds.length;
    let t;
    if (mtxt && contiguous) t = `Del ${ds[0]} al ${ds[ds.length - 1]} de ${mtxt} de ${year}`;
    else if (mtxt && ds.length === 1) t = `El ${ds[0]} de ${mtxt} de ${year}`;
    else if (mtxt && ds.length) t = `En ${ds.length} días de ${mtxt} de ${year}`;
    else if (mtxt) t = `En ${mtxt} de ${year}`;
    else t = `En ${year}`;
    if (cmp.weeks.size) t += ` (semanas ${[...cmp.weeks].sort((a, b) => a - b).join(', ')})`;
    return t;
  }

  function renderHero({ A, B, uA, uB, d, p, hasUnits }) {
    const { yA, yB } = cmp.last;
    const per = periodLabel(yB);
    const cls = d >= 0 ? 'up' : 'down';
    $('#cmpHeadline').innerHTML = `${esc(per)} vendimos <span class="amount">${moneyShort(B)}</span>, <span class="delta ${cls}">${p == null ? 'sin base' : (p > 0 ? '+' : '') + (p * 100).toFixed(1) + '%'}</span> <span class="vs">frente a ${yA}</span>`;
    const pu = varPct(uB, uA);
    const tA = uA ? A / uA : null, tB = uB ? B / uB : null;
    const foot = (x, extra = '') => (x == null ? extra : `<span class="${dirCls(x)}">${x > 0 ? '▲' : x < 0 ? '▼' : ''} ${pct(x)}</span> vs ${yA}${extra}`);
    const tiles = [
      [`Venta ${yB}`, moneyShort(B), `${moneyShort(A)} en ${yA}`],
      ['Diferencia', `${d > 0 ? '+' : ''}${moneyShort(d)}`, foot(p)],
    ];
    if (hasUnits && (uA || uB)) {
      tiles.push([`Unidades ${yB}`, num(uB), foot(pu)]);
      tiles.push(['Precio promedio', tB ? money(tB) : '—', foot(tA && tB ? varPct(tB, tA) : null)]);
    } else {
      const sp = cmp.shareP;
      tiles.push(['Participación Paraíso', sp?.b != null ? fmtPct.format(sp.b) : '—', sp?.a != null && sp?.b != null ? `<span class="${dirCls(sp.b - sp.a)}">${(sp.b - sp.a > 0 ? '+' : '') + ((sp.b - sp.a) * 100).toFixed(1)} pp</span> vs ${yA}` : '']);
      tiles.push([`${cmp.last.list.filter((g) => g.vB).length} con venta`, `${cmp.last.list.filter((g) => g.vB > g.vA).length} crecen`, `${cmp.last.list.filter((g) => g.vB < g.vA).length} caen`]);
    }
    $('#cmpKpis').innerHTML = tiles.map(([l, v, f]) => `<div class="hk"><div class="hk-label">${l}</div><div class="hk-value">${v}</div><div class="hk-foot">${f}</div></div>`).join('');
    $$('[data-ya]').forEach((e) => (e.textContent = yA));
    $$('[data-yb]').forEach((e) => (e.textContent = yB));
    $$('[data-what]').forEach((e) => (e.textContent = WHAT_ONE[cmp.last.by]));
  }

  function renderInsights({ up, down, list, B, d, uA, uB, A, hasUnits, name, what }) {
    const card = (kind, icon, title, nm, value, note) => `<article class="ins ${kind}"><div class="ins-top"><span class="ins-icon" aria-hidden="true">${icon}</span>${title}</div><div class="ins-name">${nm}</div><div class="ins-value">${value}</div><div class="ins-note">${note}</div></article>`;
    const cards = [];
    const contrib = (x) => (d && Math.abs(x.d / d) <= 1 ? `explica el ${fmtPct.format(Math.abs(x.d / d))} de la diferencia` : 'mueve más que la diferencia total');
    cards.push(up.length
      ? card('up', '▲', 'Lo que más suma', esc(name(up[0].g)), `+${moneyShort(up[0].d)}`, up[0].g.vA ? `${pct(varPct(up[0].g.vB, up[0].g.vA))} · ${contrib(up[0])}` : `Nueva en ${cmp.last.yB}`)
      : card('up', '▲', 'Lo que más suma', 'Nadie crece', '—', 'Todos bajan o se mantienen'));
    cards.push(down.length
      ? card('down', '▼', 'Lo que más resta', esc(name(down[0].g)), moneyShort(down[0].d), down[0].g.vB ? `${pct(varPct(down[0].g.vB, down[0].g.vA))} · ${contrib(down[0])}` : `Sin venta en ${cmp.last.yB}`)
      : card('down', '▼', 'Lo que más resta', 'Nadie cae', '—', 'Todos suben o se mantienen'));
    const top = [...list].sort((a, b) => b.vB - a.vB)[0];
    if (top && B) cards.push(card('weight', '★', 'Mayor peso', esc(name(top)), fmtPct.format(top.vB / B), `de la venta de ${cmp.last.yB} · ${moneyShort(top.vB)}`));
    if (hasUnits && uA && uB) {
      const pu = varPct(uB, uA), pp = varPct(B / uB, A / uA);
      const byPrice = Math.abs(pp) > Math.abs(pu);
      cards.push(card('mix', '⚖', 'Qué explica la diferencia', byPrice ? 'Más el precio que el volumen' : 'Más el volumen que el precio', `${pct(byPrice ? pp : pu)}`, `Unidades ${pct(pu)} · precio promedio ${pct(pp)}`));
    } else if (cmp.shareP?.a != null && cmp.shareP?.b != null) {
      const dd = cmp.shareP.b - cmp.shareP.a;
      cards.push(card('mix', '⚖', 'Participación Paraíso', `${fmtPct.format(cmp.shareP.b)} de la venta`, `${dd > 0 ? '+' : ''}${(dd * 100).toFixed(1)} pp`, `frente a ${fmtPct.format(cmp.shareP.a)} en ${cmp.last.yA}`));
    }
    $('#cmpInsights').innerHTML = cards.join('');
  }

  // Puente: venta del año base → aportes de cada grupo → venta del año comparado.
  function renderBridge({ list, A, B, name }) {
    const { yA, yB } = cmp.last;
    const moved = list.map((g) => ({ label: name(g), d: g.vB - g.vA })).filter((x) => x.d);
    moved.sort((a, b) => Math.abs(b.d) - Math.abs(a.d));
    const top = moved.slice(0, 7);
    const rest = moved.slice(7).reduce((s, x) => s + x.d, 0);
    if (Math.abs(rest) > 0.5) top.push({ label: `Otros (${moved.length - 7})`, d: rest });
    top.sort((a, b) => b.d - a.d);
    const labels = [String(yA), ...top.map((x) => x.label), String(yB)];
    let cum = A;
    const bars = [[0, A]], colors = [yearColor(yA)], deltas = [null];
    for (const x of top) { bars.push([cum, cum + x.d]); colors.push(x.d >= 0 ? '#14B88A' : '#F0475A'); deltas.push(x.d); cum += x.d; }
    bars.push([0, B]); colors.push(yearColor(yB)); deltas.push(null);
    const lo = Math.min(...bars.slice(1, -1).flat(), A, B), hi = Math.max(...bars.flat());
    const pad = (hi - lo) * 0.25 || hi * 0.05;
    const short = (s) => (s.length > 16 ? s.slice(0, 15) + '…' : s);
    upsertChart('bridge', 'chartBridge', {
      type: 'bar',
      data: { labels: labels.map(short), datasets: [{ data: bars, backgroundColor: colors, borderRadius: 6, borderSkipped: false, barPercentage: 0.72 }] },
      options: {
        responsive: true, maintainAspectRatio: false, animation: { duration: 300 },
        layout: { padding: { top: 24 } },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { title: (it) => labels[it[0].dataIndex], label: (c) => { const dv = deltas[c.dataIndex]; return dv == null ? ` ${money(bars[c.dataIndex][1])}` : ` ${signMoney(dv)}`; } } },
          bridgeLabels: { deltas, bars },
        },
        scales: {
          x: { grid: { display: false }, ticks: { maxRotation: 40, minRotation: 0, autoSkip: false, font: { size: 11 } } },
          y: { min: Math.max(0, lo - pad), grid: { color: '#EEF0F6' }, border: { display: false }, ticks: { callback: (v) => moneyShort(v) } },
        },
      },
      plugins: [bridgeLabels],
    });
  }
  const bridgeLabels = {
    id: 'bridgeLabels',
    afterDatasetsDraw(chart, _a, o) {
      const { ctx } = chart;
      ctx.save();
      ctx.font = '700 11px "Plus Jakarta Sans", system-ui, sans-serif'; ctx.textAlign = 'center';
      chart.getDatasetMeta(0).data.forEach((bar, i) => {
        const dv = o.deltas[i];
        const txt = dv == null ? moneyShort(o.bars[i][1]) : (dv > 0 ? '+' : '') + moneyShort(dv);
        const top = Math.min(bar.y, bar.base);
        ctx.fillStyle = dv == null ? '#161A3A' : dv >= 0 ? '#0E8F6B' : '#D42E43';
        ctx.fillText(txt, bar.x, top - 6);
      });
      ctx.restore();
    },
  };

  // Variación % simétrica alrededor de cero (verde a la derecha, rosa a la izquierda).
  function renderDiverge({ list, name }) {
    let items = list.filter((g) => g.vA > 0).map((g) => ({ label: name(g), p: varPct(g.vB, g.vA), d: g.vB - g.vA }));
    items.sort((a, b) => b.p - a.p);
    if (items.length > 12) items = [...items.slice(0, 6), ...items.slice(-6)];
    const m = Math.max(0.05, ...items.map((x) => Math.abs(x.p))) * 1.18;
    const short = (s) => (s.length > 22 ? s.slice(0, 21) + '…' : s);
    upsertChart('diverge', 'chartDiverge', {
      type: 'bar',
      data: { labels: items.map((x) => short(x.label)), datasets: [{ data: items.map((x) => x.p * 100), backgroundColor: items.map((x) => (x.p >= 0 ? '#14B88A' : '#F0475A')), borderRadius: 6, borderSkipped: false, barPercentage: 0.7 }] },
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false, animation: { duration: 300 },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { title: (it) => items[it[0].dataIndex].label, label: (c) => ` ${pct(items[c.dataIndex].p)} · ${signMoney(items[c.dataIndex].d)}` } },
          divLabels: { items },
        },
        scales: {
          x: { min: -m * 100, max: m * 100, grid: { color: (c) => (c.tick.value === 0 ? '#161A3A' : '#EEF0F6'), lineWidth: (c) => (c.tick.value === 0 ? 1.5 : 1) }, border: { display: false }, ticks: { callback: (v) => (v > 0 ? '+' : '') + Math.round(v) + '%' } },
          y: { grid: { display: false }, ticks: { autoSkip: false, font: { size: 11.5 } } },
        },
      },
      plugins: [divLabels],
    });
  }
  const divLabels = {
    id: 'divLabels',
    afterDatasetsDraw(chart, _a, o) {
      const { ctx } = chart;
      ctx.save();
      ctx.font = '700 11px "Plus Jakarta Sans", system-ui, sans-serif'; ctx.textBaseline = 'middle';
      chart.getDatasetMeta(0).data.forEach((bar, i) => {
        const p = o.items[i].p, txt = pct(p);
        const inside = Math.abs(bar.x - bar.base) > ctx.measureText(txt).width + 16;
        if (inside) { ctx.fillStyle = '#fff'; ctx.textAlign = p >= 0 ? 'right' : 'left'; ctx.fillText(txt, bar.x + (p >= 0 ? -8 : 8), bar.y); }
        else { ctx.fillStyle = p >= 0 ? '#0E8F6B' : '#D42E43'; ctx.textAlign = p >= 0 ? 'left' : 'right'; ctx.fillText(txt, bar.x + (p >= 0 ? 6 : -6), bar.y); }
      });
      ctx.restore();
    },
  };

  // ---------------------------------------------------------------- descargas con el formato del reporte (PDF y Excel)
  const loaded = {};
  const loadScript = (src) => loaded[src] || (loaded[src] = new Promise((ok, fail) => {
    const el = document.createElement('script'); el.src = src; el.onload = ok; el.onerror = () => fail(new Error('No se pudo cargar ' + src)); document.head.appendChild(el);
  }));
  const VIEW_NAME = { resumen: 'Resumen', comparativa: 'Comparativa', anual: 'Por año', datos: 'Datos', admin: 'Administración' };
  const stamp = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
  // Texto del corte y filtros para encabezados de descargas.
  function cutLine(short) {
    const F = lastF, C = F.C;
    const corte = `Corte: ${C.dateText}${C.hasW ? ` (semana ${C.w})` : ''}`;
    if (short) return corte;
    const filt = [state.hasCob && F.cob !== 'all' ? COB_LABEL[F.cob] : null, F.tipo ? TIPO_LABEL[F.tipo] : null, F.chain ? `Cadena ${F.chain}` : null, F.brand ? `Marca ${F.brand}` : null].filter(Boolean);
    return [corte, `Años ${F.years.join(', ')}`, ...filt].join(' · ');
  }
  const scopeText = () => (!perm('ver_todo') && session.user?.accesos?.length ? ` · Alcance: ${accesosText(session.user.accesos)}` : '');

  // Tablas del Resumen (se arman con los mismos cálculos que la pantalla).
  function resumenAOAs(F, R) {
    const C = F.C, Y = C.y, P = Y - 1, out = [];
    const cur = R.byYear.get(Y), prev = R.byYear.get(P);
    const kp = [['Acumulado al corte', prev?.ytd ?? null, cur.ytd], [`${C.monthLabel} (al corte)`, prev?.month ?? null, cur.month]];
    if (C.hasW) kp.push([`Semana ${C.w}`, prev?.week ?? null, cur.week]);
    out.push({ name: 'Resumen', title: 'Indicadores clave', T: { head: ['Indicador', String(P), String(Y), 'Var S/', 'Var %'], rows: kp.map(([l, a, b]) => [l, a, b, a == null ? null : b - a, varPct(b, a)]), foot: [], types: ['t', '$', '$', '$', '%'] } });
    const months = [];
    for (let m = 1; m <= 12; m++) { const a = R.byYM.get(P * 100 + m) || 0, b = m > C.m ? null : R.byYM.get(Y * 100 + m) || 0; months.push([MESES_LARGO[m - 1] + (m === C.m ? ' (al corte)' : ''), a, b, b == null ? null : b - a, b == null ? null : varPct(b, a)]); }
    out.push({ name: 'Mes a mes', title: `Venta mes a mes ${Y} vs ${P}`, T: { head: ['Mes', String(P), String(Y), 'Var S/', 'Var %'], rows: months, foot: [], types: ['t', '$', '$', '$', '%'] } });
    const ch = [...R.byChain.values()].map((o) => ({ c: o.cadena, a: o.y[P]?.ytd || 0, b: o.y[Y]?.ytd || 0 })).filter((x) => x.a || x.b).sort((x, y) => y.b - x.b);
    const TA = ch.reduce((t, x) => t + x.a, 0), TB = ch.reduce((t, x) => t + x.b, 0);
    out.push({ name: 'Cadenas', title: `Reparto por cadena (acumulado al corte)`, T: { head: ['Cadena', String(P), String(Y), 'Var %', `Part. ${P}`, `Part. ${Y}`], rows: ch.map((x) => [x.c, x.a, x.b, varPct(x.b, x.a), TA ? x.a / TA : null, TB ? x.b / TB : null]), foot: [['Total', TA, TB, varPct(TB, TA), TA ? 1 : null, TB ? 1 : null]], types: ['t', '$', '$', '%', 'p', 'p'] } });
    const st = [...R.byStore.values()].map((o) => ({ o, a: o.y[P]?.ytd || 0, b: o.y[Y]?.ytd || 0 })).filter((x) => x.a || x.b).map((x) => [x.o.cadena, x.o.tienda, x.a, x.b, x.b - x.a, varPct(x.b, x.a)]).sort((x, y) => y[4] - x[4]);
    out.push({ name: 'Tiendas', title: 'Tiendas ordenadas por diferencia (acumulado al corte)', T: { head: ['Cadena', 'Tienda', String(P), String(Y), 'Var S/', 'Var %'], rows: st, foot: [], types: ['t', 't', '$', '$', '$', '%'] } });
    return out;
  }

  function sheetsForView() {
    const F = lastF, R = lastR, v = state.view;
    if (v === 'resumen') return resumenAOAs(F, R);
    if (v === 'anual') {
      const out = [{ name: 'Por año', title: 'Evaluación por año', T: yearsAOA(F, R) }];
      if (F.showShare) out.push({ name: 'Paraíso vs Competencia', title: 'Paraíso vs Competencia por año', T: groupsAOA(F, R) });
      return out;
    }
    if (v === 'comparativa' && cmp.last) {
      const hasU = cmp.last.list.some((g) => g.uA || g.uB);
      const out = [];
      if (cmp.analysis?.length) out.push({ name: 'Análisis', title: 'Análisis del Vs', lines: cmp.analysis });
      out.push({ name: 'Soles', title: `Comparativa en soles por ${cmp.last.labels.join(' / ').toLowerCase()}`, T: cmpAOA('v') });
      if (hasU) out.push({ name: 'Unidades', title: 'Comparativa en unidades', T: cmpAOA('u') });
      if (cmp.wk?.weeks.length) { out.push({ name: 'Semanas soles', title: 'Soles por semana', T: weekAOA('v') }); if (hasU) out.push({ name: 'Semanas und', title: 'Unidades por semana', T: weekAOA('u') }); }
      return out;
    }
    return [];
  }

  async function exportXlsxStyled() {
    if (!lastR) return toast('Todavía no hay datos para descargar.');
    const sheets = sheetsForView();
    if (!sheets.length) return toast('Esta pestaña no tiene tablas para descargar.');
    toast('Preparando el Excel…');
    await loadScript('vendor/exceljs.min.js');
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Reporte Vs Anual';
    const NIGHT = 'FF1B1F4B', ZEBRA = 'FFF8F9FD', GOOD = 'FF0E8F6B', BAD = 'FFD42E43';
    const sub = `${VIEW_NAME[state.view]} · ${cutLine()}${scopeText()}`;
    for (const sh of sheets) {
      const ws = wb.addWorksheet(sh.name.slice(0, 31), { views: [{ showGridLines: false }] });
      ws.getCell('A1').value = sh.title; ws.getCell('A1').font = { name: 'Calibri', size: 16, bold: true, color: { argb: NIGHT } };
      ws.getCell('A2').value = sub; ws.getCell('A2').font = { size: 10, color: { argb: 'FF7C829E' } };
      if (sh.lines) { sh.lines.forEach((t, i) => { const c = ws.getCell(4 + i, 1); c.value = '• ' + t; c.alignment = { wrapText: true, vertical: 'top' }; }); ws.getColumn(1).width = 140; continue; }
      const T = sh.T, hr = 4;
      const fmt = (t) => (t === '$' ? '"S/" #,##0.00' : t === '%' ? '+0.0%;-0.0%;0.0%' : t === 'p' ? '0.0%' : t === 'pp' ? '+0.0%;-0.0%' : t === 'n' ? '#,##0' : '@');
      const header = ws.getRow(hr); header.values = T.head; header.height = 30;
      header.eachCell((c) => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NIGHT } }; c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }; });
      const put = (arr, rowIdx, kind) => {
        const row = ws.getRow(rowIdx); row.values = arr.map((v) => (v == null ? '' : v));
        row.eachCell({ includeEmpty: true }, (c, col) => {
          const t = T.types[col - 1];
          if (t && t !== 't') { c.numFmt = fmt(t); c.alignment = { horizontal: 'right' }; }
          if (kind === 'foot') { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NIGHT } }; c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; }
          else if (kind === 'zebra') c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ZEBRA } };
          if (kind !== 'foot' && (t === '%' || t === 'pp') && typeof c.value === 'number') c.font = { bold: true, color: { argb: c.value >= 0 ? GOOD : BAD } };
          c.border = { bottom: { style: 'thin', color: { argb: 'FFE2E6F0' } } };
        });
      };
      T.rows.forEach((r, i) => put(r, hr + 1 + i, i % 2 ? 'zebra' : 'row'));
      (T.foot || []).forEach((r, i) => put(r, hr + 1 + T.rows.length + i, 'foot'));
      T.head.forEach((h, i) => { ws.getColumn(i + 1).width = T.types[i] === 't' ? Math.max(14, Math.min(34, Math.max(String(h).length, ...T.rows.slice(0, 200).map((r) => String(r[i] ?? '').length)) + 2)) : Math.max(13, Math.min(22, String(h).length + 2)); });
      ws.views = [{ showGridLines: false, state: 'frozen', ySplit: hr }];
      if (T.rows.length) ws.autoFilter = { from: { row: hr, column: 1 }, to: { row: hr + T.rows.length, column: T.head.length } };
    }
    const buf = await wb.xlsx.writeBuffer();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    a.download = `Reporte_Vs_Anual_${VIEW_NAME[state.view].replace(/\s/g, '_')}_${stamp()}.xlsx`;
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  // PDF: captura la pestaña tal como se ve y la reparte en hojas A4 horizontales sin cortar los bloques.
  async function exportPdf() {
    const node = $('#view-' + state.view);
    if (!node || !lastR) return toast('Todavía no hay datos para descargar.');
    toast('Preparando el PDF…');
    await Promise.all([loadScript('vendor/html2canvas.min.js'), loadScript('vendor/jspdf.umd.min.js')]);
    document.body.classList.add('pdf-mode');
    // Las secciones plegadas se abren para la captura (y así el PDF incluye el análisis completo).
    const closed = [...node.querySelectorAll('details:not([open])')];
    closed.forEach((d) => (d.open = true));
    window.scrollTo(0, 0);
    await new Promise((r) => setTimeout(r, 120));
    try {
      const scale = 2;
      const canvas = await html2canvas(node, { scale, backgroundColor: '#F4F6FB', useCORS: true, logging: false, windowWidth: document.documentElement.clientWidth });
      const top0 = node.getBoundingClientRect().top + window.scrollY;
      // Puntos de corte: inicio de cada bloque; si el bloque va justo después de un título de sección, se corta antes del título.
      const breaks = [...node.querySelectorAll('.hero, .period, .story, .panel, .insights, .pair, .timeline, .year-cards')]
        .map((el) => { const blk = el.parentElement?.closest('.pair, .insights') || el; const prev = blk.previousElementSibling; return prev && prev.matches('.story-title, .story-head, .subnav') ? prev : blk; })
        .map((el) => Math.round((el.getBoundingClientRect().top + window.scrollY - top0) * scale)).filter((y) => y > 0).sort((a, b) => a - b);
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
      const PW = 297, PH = 210, M = 8, HEAD = 12;
      const imgW = PW - 2 * M, pxPerMm = canvas.width / imgW, maxH = Math.floor((PH - M - HEAD - 6) * pxPerMm);
      let y = 0, page = 0;
      const title = `Reporte Vs Anual · ${VIEW_NAME[state.view]}${scopeText()}`;
      const cut = cutLine(true);
      while (y < canvas.height - 4) {
        let end = Math.min(canvas.height, y + maxH);
        if (end < canvas.height) { const b = breaks.filter((v) => v > y + maxH * 0.35 && v <= end).pop(); if (b) end = b - 6 * scale; }
        const h = end - y;
        const part = document.createElement('canvas'); part.width = canvas.width; part.height = h;
        part.getContext('2d').drawImage(canvas, 0, y, canvas.width, h, 0, 0, canvas.width, h);
        if (page) pdf.addPage();
        pdf.setFillColor(27, 31, 75); pdf.rect(0, 0, PW, HEAD, 'F');
        pdf.setTextColor(255, 255, 255); pdf.setFont('helvetica', 'bold'); pdf.setFontSize(11); pdf.text(title, M, 8);
        pdf.setFont('helvetica', 'normal'); pdf.setFontSize(8); pdf.text(`${cut}  |  ${session.user?.nombre || ''}  |  ${new Date().toLocaleDateString('es-PE')}`, PW - M, 8, { align: 'right' });
        pdf.addImage(part.toDataURL('image/jpeg', 0.92), 'JPEG', M, HEAD + 4, imgW, h / pxPerMm);
        pdf.setTextColor(124, 130, 158); pdf.setFontSize(8); pdf.text(`Página ${page + 1}`, PW - M, PH - 4, { align: 'right' });
        y = end; page++;
      }
      pdf.save(`Reporte_Vs_Anual_${VIEW_NAME[state.view].replace(/\s/g, '_')}_${stamp()}.pdf`);
    } finally {
      closed.forEach((d) => (d.open = false));
      document.body.classList.remove('pdf-mode');
    }
  }

  // ---------------------------------------------------------------- copiar / exportar
  function aoaFor(id) {
    const F = lastF, R = lastR;
    return { cmpSoles: () => cmpAOA('v'), cmpUnd: () => cmpAOA('u'), cmpWeek: () => weekAOA('v'), cmpWeekUnd: () => weekAOA('u'), groups: () => groupsAOA(F, R), years: () => yearsAOA(F, R) }[id]();
  }
  async function copyTable(id) {
    const T = aoaFor(id);
    const fmt = (v, t) => (v == null ? '' : t === 'pp' ? (v * 100).toFixed(1) + ' pp' : t === '%' || t === 'p' ? (v * 100).toFixed(1) + '%' : t === '$' || t === 'n' ? String(Math.round(v * 100) / 100) : String(v));
    const lines = [T.head, ...T.rows, ...T.foot].map((r) => r.map((v, i) => fmt(v, T.types[i])).join('\t'));
    try { await navigator.clipboard.writeText(lines.join('\n')); toast('Tabla copiada: pégala en Excel.'); }
    catch { toast('No se pudo copiar en este navegador.'); }
  }

  function sheetFrom(T, title) {
    const aoa = [[title], [], T.head, ...T.rows, ...T.foot];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let r = 3; r <= range.e.r; r++) {
      T.types.forEach((t, c) => {
        const ce = ws[XLSX.utils.encode_cell({ r, c })];
        if (!ce || ce.t !== 'n') return;
        if (t === '$') ce.v = Math.round(ce.v * 100) / 100;
        ce.z = t === '%' || t === 'p' ? '0.0%' : t === 'pp' ? '+0.0%;-0.0%' : t === '$' ? '#,##0.00' : '#,##0';
      });
    }
    ws['!cols'] = T.head.map((h, i) => ({ wch: T.types[i] === 't' && i < 2 ? 28 : Math.max(13, String(h).length + 2) }));
    if (T.rows.length) ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 2, c: 0 }, e: { r: 2 + T.rows.length, c: T.head.length - 1 } }) };
    return ws;
  }

  function exportExcel() {
    if (!lastR) return;
    const F = lastF, R = lastR, C = F.C;
    const corte = `Corte: ${C.y} · ${C.monthLabel}${C.hasW ? ' · Semana ' + C.w : ''} (último dato ${C.dateText})`;
    const filt = `${corte}${F.cob && F.cob !== 'all' ? ' · ' + COB_LABEL[F.cob] : ''}${F.tipo ? ' · ' + TIPO_LABEL[F.tipo] : ''}${F.chain ? ' · Cadena: ' + F.chain : ''}${F.brand ? ' · Marca: ' + F.brand : ''}`;
    const wb = XLSX.utils.book_new();
    if (cmp.analysis?.length) {
      const ws = XLSX.utils.aoa_to_sheet([['Análisis del Vs'], [$('#cmpNote').textContent], [], ...cmp.analysis.map((t) => [t])]);
      ws['!cols'] = [{ wch: 140 }];
      XLSX.utils.book_append_sheet(wb, ws, 'Análisis');
    }
    if (cmp.last && cmp.last.list.length) {
      XLSX.utils.book_append_sheet(wb, sheetFrom(cmpAOA('v'), 'Comparativa (soles) · ' + $('#cmpNote').textContent), 'Comparativa soles');
      if (cmp.last.list.some((g) => g.uA || g.uB)) XLSX.utils.book_append_sheet(wb, sheetFrom(cmpAOA('u'), 'Comparativa (unidades) · ' + $('#cmpNote').textContent), 'Comparativa und');
      if (cmp.wk && cmp.wk.weeks.length) {
        XLSX.utils.book_append_sheet(wb, sheetFrom(weekAOA('v'), 'Comparativa por semana (soles) · ' + $('#cmpNote').textContent), 'Semanas soles');
        if (cmp.last.list.some((g) => g.uA || g.uB)) XLSX.utils.book_append_sheet(wb, sheetFrom(weekAOA('u'), 'Comparativa por semana (unidades) · ' + $('#cmpNote').textContent), 'Semanas und');
      }
    }
    XLSX.utils.book_append_sheet(wb, sheetFrom(yearsAOA(F, R), 'Evaluación por año · ' + filt), 'Por año');
    if (F.showShare) XLSX.utils.book_append_sheet(wb, sheetFrom(groupsAOA(F, R), 'Venta Paraíso vs Competencia · ' + filt), 'Paraíso vs Competencia');
    XLSX.writeFile(wb, `Reporte_Vs_Anual_${F.years[0]}-${C.y}${F.chain ? '_' + F.chain.replace(/[^\w]+/g, '_') : ''}.xlsx`);
  }

  // ---------------------------------------------------------------- eventos
  function setView(v) {
    state.view = v;
    $$('.tab').forEach((t) => t.classList.toggle('active', t.dataset.view === v));
    $$('.view').forEach((s) => s.classList.toggle('active', s.id === 'view-' + v));
    $('#analysis').hidden = v === 'datos' || v === 'admin';
    $$('[data-analysis]').forEach((s) => s.toggleAttribute('data-off', !state.records.length));
    if (v === 'comparativa') renderCmp();
    if (v === 'admin') loadAdmin();
    requestAnimationFrame(() => requestAnimationFrame(() => Object.values(state.charts).forEach((c) => c.resize())));
    window.scrollTo({ top: 0 });
  }

  function init() {
    if (window.Chart) { Chart.defaults.font.family = '"Plus Jakarta Sans", "Segoe UI", system-ui, sans-serif'; Chart.defaults.color = '#4A5072'; }
    if (!window.XLSX) { document.body.innerHTML = '<p style="padding:40px">No se pudo cargar el lector de Excel. Revisa tu conexión y recarga la página.</p>'; return; }
    $$('.tab').forEach((t) => t.addEventListener('click', () => setView(t.dataset.view)));
    $('#btnGoCompare').addEventListener('click', () => setView('resumen'));
    $('#fileInput').addEventListener('change', (e) => { handleFiles(e.target.files).then(() => { e.target.value = ''; }); });
    $('#btnSave').addEventListener('click', saveToBase);
    $('#formRepl').addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      try {
        const r = await api('/api/reemplazos', { method: 'POST', body: { original: fd.get('original'), nuevo: fd.get('nuevo') } });
        toast(`Reemplazo guardado (${num(r.filas)} registros actualizados).`); e.target.reset(); cmp.cache = null;
        coordData = await api('/api/coordinadores'); renderCoord(); renderCmp();
      } catch (x) { toast(x.message); }
    });
    $('#tblRepl').addEventListener('click', async (e) => {
      const b = e.target.closest('[data-del-repl]'); if (!b) return;
      if (!confirm(`¿Quitar el reemplazo de "${b.dataset.delRepl}"? Lo ya guardado conserva el nombre nuevo; el reemplazo deja de aplicarse en las próximas cargas.`)) return;
      try { await api('/api/reemplazos/' + encodeURIComponent(b.dataset.delRepl), { method: 'DELETE' }); coordData = await api('/api/coordinadores'); renderCoord(); }
      catch (x) { toast(x.message); }
    });
    $('#btnReload').addEventListener('click', loadServer);
    for (const id of ['#tblBase', '#baseYears']) $(id).addEventListener('click', (e) => { const b = e.target.closest('[data-del-year]'); if (b) deleteYear(+b.dataset.delYear); });

    // Sesión y usuarios
    document.addEventListener('click', (e) => {
      if (e.target.closest('[data-open-login]')) { const d = $('#dlgLogin'); $('[data-error]', d).textContent = ''; d.showModal(); }
      if (e.target.closest('[data-close]')) e.target.closest('dialog').close();
      if (e.target.closest('#btnLogout')) { setSession(null); toast('Sesión cerrada.'); }
      if (e.target.closest('#btnUsers')) openUsers();
      if (e.target.closest('#btnPwd')) { const d = $('#dlgPassword'); d.querySelector('form').reset(); $('[data-error]', d).textContent = ''; d.showModal(); }
    });
    $('#formLogin').addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target), err = $('[data-error]', e.target);
      err.textContent = '';
      try {
        const r = await api('/api/login', { method: 'POST', body: { dni: fd.get('dni'), password: fd.get('password') } });
        setSession(r); e.target.reset(); $('#dlgLogin').close(); toast(`Hola, ${r.user.nombre}.`); loadServer();
      } catch (x) { err.textContent = x.message; }
    });
    $('#formPassword').addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target), err = $('[data-error]', e.target);
      try { await api('/api/password', { method: 'POST', body: { actual: fd.get('actual'), nueva: fd.get('nueva') } }); $('#dlgPassword').close(); toast('Contraseña actualizada.'); }
      catch (x) { err.textContent = x.message; }
    });
    $('#formUser').addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target), err = $('[data-error]', e.target);
      err.textContent = '';
      try { await api('/api/usuarios', { method: 'POST', body: { dni: fd.get('dni'), nombre: fd.get('nombre'), password: fd.get('password'), admin: !!fd.get('admin') } }); e.target.reset(); toast('Usuario creado.'); openUsers(); }
      catch (x) { err.textContent = x.message; }
    });
    $('#tblUsers').addEventListener('click', async (e) => {
      const b = e.target.closest('[data-del-user]');
      if (!b || !confirm(`¿Eliminar el usuario con DNI ${b.dataset.delUser}?`)) return;
      try { await api('/api/usuarios/' + encodeURIComponent(b.dataset.delUser), { method: 'DELETE' }); openUsers(); }
      catch (x) { toast(x.message); }
    });

    const dz = $('#dropzone');
    ['dragenter', 'dragover'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); if (ev === 'drop' || e.target === document.documentElement) dz.classList.remove('over'); }));
    document.addEventListener('drop', (e) => { if (e.dataTransfer?.files?.length && session.user) { setView('datos'); handleFiles(e.dataTransfer.files); } });

    $('#filesList').addEventListener('change', onFilesPanelChange);
    $('#filesList').addEventListener('click', (e) => {
      const el = e.target.closest('[data-act="remove"]');
      if (!el) return;
      state.files = state.files.filter((f) => f.id !== +el.dataset.id);
      rebuildAll();
    });
    $('#btnClear').addEventListener('click', () => { state.files = []; rebuildAll(); });

    $('#fFrom').addEventListener('change', () => { state.fromTouched = true; });
    ['#fFrom', '#fCob', '#fTipo', '#fChain', '#fBrand'].forEach((s) => $(s).addEventListener('change', renderAll));
    $$('[data-copy]').forEach((b) => b.addEventListener('click', () => copyTable(b.dataset.copy)));
    $('#btnExport').addEventListener('click', () => (perm('descargar') ? exportXlsxStyled().catch((e) => toast(e.message)) : toast('Tu perfil no puede descargar.')));
    $('#btnXlsx').addEventListener('click', () => exportXlsxStyled().catch((e) => toast(e.message)));
    $('#btnPdf').addEventListener('click', () => exportPdf().catch((e) => { document.body.classList.remove('pdf-mode'); toast(e.message); }));
    $('#sumShare').addEventListener('click', (e) => { const b = e.target.closest('[data-chain]'); if (b) goDetail({ months: [lastF.C.m], days: lastF.C.d ? Array.from({ length: lastF.C.d }, (_, i) => i + 1) : [], by: 'tienda', chain: b.dataset.chain }); });
    for (const id of ['#sumUp', '#sumDown']) $(id).addEventListener('click', (e) => { const b = e.target.closest('[data-chain]'); if (b) goDetail({ months: [lastF.C.m], days: lastF.C.d ? Array.from({ length: lastF.C.d }, (_, i) => i + 1) : [], by: 'tienda', chain: b.dataset.chain }); });
    const heat = $('#sumHeat'), tip = $('#heatTip');
    heat.addEventListener('click', (e) => { const b = e.target.closest('.hd[data-t]'); if (!b) return; const d = new Date(+b.dataset.t); goDetail({ months: [d.getUTCMonth() + 1], days: [d.getUTCDate()], by: 'cadena' }); });
    heat.addEventListener('mousemove', (e) => {
      const b = e.target.closest('.hd[data-t]');
      if (!b) { tip.hidden = true; return; }
      const d = new Date(+b.dataset.t), bv = +b.dataset.b, av = b.dataset.a === '' ? null : +b.dataset.a;
      const dias = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
      tip.innerHTML = `<b>${dias[d.getUTCDay()]} ${d.getUTCDate()} de ${MESES_LARGO[d.getUTCMonth()].toLowerCase()}</b><br>${money(bv)}${av == null ? '<br>Sin dato del año anterior' : `<br>${lastF.C.y - 1}: ${money(av)} · <b>${pct(varPct(bv, av))}</b>`}`;
      tip.hidden = false;
      tip.style.left = Math.min(window.innerWidth - 240, e.clientX + 14) + 'px'; tip.style.top = e.clientY + 14 + 'px';
    });
    heat.addEventListener('mouseleave', () => { tip.hidden = true; });
    $('#btnFilters').addEventListener('click', () => setDrawer(!$('#filterDrawer').classList.contains('open')));
    $('#btnFiltersClose').addEventListener('click', () => setDrawer(false));
    $('#drawerScrim').addEventListener('click', () => setDrawer(false));
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && $('#filterDrawer').classList.contains('open')) setDrawer(false); });
    $('#btnFiltersReset').addEventListener('click', () => {
      $('#fTipo').value = ''; $('#fChain').value = ''; $('#fBrand').value = ''; $('#fCob').value = 'si';
      state.fromTouched = false; syncFilterOptions(); renderAll();
    });
    const toggle = (set, v) => (set.has(v) ? set.delete(v) : set.add(v));
    for (const [id, key] of [['#slMes', 'months'], ['#slSem', 'weeks'], ['#slDia', 'days']]) {
      $(id).addEventListener('click', (e) => {
        const b = e.target.closest('.chip'); if (!b) return;
        const v = +b.dataset.v;
        // Clic normal: selecciona solo ese valor; Ctrl/Cmd + clic: suma o quita (como los segmentadores de Excel).
        if (e.ctrlKey || e.metaKey || $('#cmpMulti').checked) toggle(cmp[key], v);
        else if (cmp[key].size === 1 && cmp[key].has(v)) cmp[key].clear();
        else { cmp[key].clear(); cmp[key].add(v); }
        renderCmp();
      });
    }
    $$('[data-clear]').forEach((b) => b.addEventListener('click', () => { cmp[b.dataset.clear].clear(); renderCmp(); }));
    $$('.seg[data-seg]').forEach((seg) => seg.addEventListener('click', (e) => {
      const b = e.target.closest('button[data-unit]'); if (!b || b.disabled) return;
      unitSel[seg.dataset.seg] = b.dataset.unit;
      renderCmp().then(() => Object.values(state.charts).forEach((c) => c.resize()));
    }));
    $('#cmpToCut').addEventListener('click', () => {
      const C = lastF.C; cmp.months = new Set([C.m]); cmp.weeks.clear();
      cmp.days = new Set(C.d ? Array.from({ length: C.d }, (_, i) => i + 1) : []); renderCmp();
    });
    ['#cmpYearA', '#cmpYearB', '#cmpBy'].forEach((s) => $(s).addEventListener('change', renderCmp));
    renderFiles();
    renderSession();
    // Comprueba que la sesión guardada siga siendo válida y trae los permisos vigentes.
    if (session.token) api('/api/me').then((r) => { session.user = r.user; store.set('user', r.user); renderSession(); loadServer(); }).catch(() => { setSession(null); });
    $('#formGate').addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target), err = $('[data-error]', e.target);
      err.textContent = '';
      const btn = e.target.querySelector('button'); btn.disabled = true; btn.textContent = 'Ingresando…';
      try {
        const r = await api('/api/login', { method: 'POST', body: { dni: fd.get('dni'), password: fd.get('password') } });
        setSession(r); e.target.reset(); loadServer();
      } catch (x) { err.textContent = x.message; } finally { btn.disabled = false; btn.textContent = 'Ingresar'; }
    });
    initAdmin();
  }

  // Exponer funciones de parseo para pruebas automáticas.
  window.__rva = { parseAmount, parseYM, parseMonth, parseYear, weekOf, detectMapping, detectHeaderRow, state };
  document.addEventListener('DOMContentLoaded', init);
})();
