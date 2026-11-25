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
    for (const role of ['venta', 'fecha', 'anio', 'semana', 'mes', 'cobertura', 'cadena', 'tienda', 'marca', 'unidades']) {
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

    const label = (v, d) => { const s = v == null ? '' : v instanceof Date ? v.toLocaleDateString('es-PE') : String(v).trim(); return s || d; };
    const dagg = new Map();
    const add = (cad, tie, mar, y, m, w, v, u, d = null, cob = null) => {
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
            add(cad, tie, mar, c.y, c.m, 0, v, 0, null, cobOf(r));
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
        add(cad, tie, mar, ym.y, ym.m, w, v, u, ym.d, cobOf(r));
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
    Object.assign(f, { agg, dagg, warnings, used, years: [...years].sort(), total, maxDate, maxByYear, monthsByYear, wide: anyWide, hasCob, covTotal, uncovTotal, unkTotal });
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
    session.token = s ? s.token : null;
    session.user = s ? s.user : null;
    store.set('token', session.token); store.set('user', session.user);
    renderSession();
  }

  function renderSession() {
    const u = session.user;
    $('#userBox').innerHTML = u
      ? `<span class="who">👤 ${esc(u.nombre)}</span>${u.admin ? '<button class="btn btn-ghost" id="btnUsers">Usuarios</button>' : ''}<button class="btn btn-ghost" id="btnPwd">Contraseña</button><button class="btn btn-ghost" id="btnLogout">Salir</button>`
      : `<button class="btn btn-primary" data-open-login>Iniciar sesión</button>`;
    $('#uploadArea').hidden = !u;
    $('#loginHint').hidden = !!u;
    renderBase();
  }

  let serverData = null;
  async function loadServer() {
    const st = $('#serverStatus');
    st.hidden = false; st.className = 'server-status';
    st.innerHTML = '<span class="spinner"></span> Cargando los datos de la base… (si el servidor estaba en reposo puede tardar hasta un minuto)';
    try {
      serverData = await api('/api/data');
      cmp.data = null;
      applyServerData(serverData);
      st.hidden = true;
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
    const logged = !!session.user;
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
    const agg = new Map(), dagg = new Map();
    const maxY = {};
    const merge = (dst, src) => { for (const [k, o] of src) { const c = dst.get(k); if (c) { c.v += o.v; c.u += o.u; } else dst.set(k, { v: o.v, u: o.u }); } };
    for (const f of ok) {
      merge(agg, f.agg); merge(dagg, f.dagg || new Map());
      for (const y in f.maxByYear) { const a = f.maxByYear[y]; if (!maxY[y] || a.stamp > maxY[y].stamp) maxY[y] = a; }
    }
    const byYear = {};
    const Y = (y) => byYear[y] || (byYear[y] = { rows: [], dias: [] });
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
        const Y = P.byYear[a.anio] || { rows: [], dias: [] };
        const r = await api('/api/cargas', { method: 'POST', body: { archivos: P.archivos, anios: [a], rows: Y.rows, dias: Y.dias } });
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
    const def = prev && ys.includes(+prev) ? +prev : (ys.find((y) => y >= 2022) ?? ys[0]);
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
  const PALETTE = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#6250d6', '#e34948'];
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
      const ytd = inYtd(r), mo = r.m === C.m, wk = C.hasW && r.w === C.w;
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
    renderKpis(F, R);
    renderYearCharts(F, R);
    renderTable('tblYears', yearsAOA(F, R));
    $('#groupsCard').hidden = !F.showShare;
    if (F.showShare) { renderTable('tblGroups', groupsAOA(F, R)); renderGroupChart(F, R); }
    $('#noteYears').textContent = `* ${F.C.y} es el año en curso: su "Año completo" llega solo hasta el corte, por eso no se compara con el año anterior. El acumulado de cada año se calcula hasta el mismo punto (${F.C.hasW ? 'semana ' + F.C.w : MESES_LARGO[F.C.m - 1]}) para comparar en igualdad de condiciones.` + (F.C.hasW && F.years.some((y) => !state.weekYears.has(y)) ? ' Los años sin fecha diaria se acumulan hasta el mes de corte completo.' : '');
    renderCmp();
  }

  function renderBanner(F) {
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

  const COL_PARAISO = '#1baf7a', COL_COMP = '#8a93a3';
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
    if (!cmp.init) {
      selB.innerHTML = years.map((y) => `<option>${y}</option>`).join(''); selB.value = years.includes(C.y) ? C.y : years[years.length - 1];
      selA.innerHTML = years.map((y) => `<option>${y}</option>`).join(''); selA.value = years.includes(+selB.value - 1) ? +selB.value - 1 : years[0];
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
    const by = $('#cmpBy').value;
    const keyOf = (r) => (by === 'marca' ? r.marca || '(Sin marca)' : by === 'tipo' ? TIPO_LABEL[r.tipo] : r.cadena);
    const groups = new Map();
    for (const r of sel) {
      const k = keyOf(r);
      const g = groups.get(k) || { k, vA: 0, vB: 0, uA: 0, uB: 0 };
      if (r.y === yA) { g.vA += r.v; g.uA += r.u; } else { g.vB += r.v; g.uB += r.u; }
      groups.set(k, g);
    }
    const list = [...groups.values()].sort((a, b) => a.k.localeCompare(b.k, 'es'));
    cmp.last = { yA, yB, list, by };
    const hasUnits = sel.some((r) => r.u);
    renderTable('tblCmpSoles', cmpAOA('v'));
    $('#cmpUndCard').hidden = !hasUnits;
    if (hasUnits) renderTable('tblCmpUnd', cmpAOA('u'));

    // Detalle por semana: columnas Sem-N × año, como la tabla dinámica de Excel.
    const weeks = (cmp.weeks.size ? [...cmp.weeks] : [...new Set(sel.filter((r) => r.w > 0).map((r) => r.w))]).sort((a, b) => a - b);
    const wk = new Map();
    for (const r of sel) {
      if (!(r.w > 0)) continue;
      const k = keyOf(r);
      const g = wk.get(k) || { k, w: {} };
      const c = g.w[r.w] || (g.w[r.w] = { vA: 0, vB: 0, uA: 0, uB: 0 });
      if (r.y === yA) { c.vA += r.v; c.uA += r.u; } else { c.vB += r.v; c.uB += r.u; }
      wk.set(k, g);
    }
    cmp.wk = { weeks, list: [...wk.values()].sort((a, b) => a.k.localeCompare(b.k, 'es')), yA, yB, by };
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

    const ms = [...cmp.months].sort((a, b) => a - b).map((m) => MESES_LARGO[m - 1]);
    $('#cmpTitle').textContent = cmp.months.size === 1 && [...cmp.months][0] === C.m && yB === C.y ? 'MES ACTUAL' : (ms.join(' + ') || 'TODO EL AÑO').toUpperCase();
    const dsel = [...cmp.days].sort((a, b) => a - b);
    const dTxt = !dsel.length ? 'todos los días' : dsel.length > 2 && dsel[dsel.length - 1] - dsel[0] + 1 === dsel.length ? `días ${dsel[0]} al ${dsel[dsel.length - 1]}` : `días ${dsel.join(', ')}`;
    $('#cmpNote').innerHTML = `Comparando <b>${yB}</b> contra <b>${yA}</b> · ${ms.length ? ms.join(', ') : 'todos los meses'} · ${cmp.weeks.size ? 'semanas ' + [...cmp.weeks].sort((a, b) => a - b).join(', ') : 'todas las semanas'} · ${dTxt}${noDay ? ' · <span class="down">los archivos no traen fecha diaria</span>' : ''}`;
  }

  function weekAOA(field) {
    const { weeks, list, yA, yB, by } = cmp.wk;
    const unit = field === 'v' ? 'Soles' : 'Und';
    const A = field === 'v' ? 'vA' : 'uA', B = field === 'v' ? 'vB' : 'uB';
    const t = field === 'v' ? '$' : 'n';
    const head = [by === 'marca' ? 'Marca' : by === 'tipo' ? 'Tipo de venta' : 'Cadena'];
    const types = ['t'];
    for (const w of weeks) { head.push(`Sem-${w} ${yA} ${unit}`, `Sem-${w} ${yB} ${unit}`, `Sem-${w} Var %`); types.push(t, t, '%'); }
    head.push(`Total ${yA} ${unit}`, `Total ${yB} ${unit}`, 'Total Var %'); types.push(t, t, '%');
    const rowOf = (label, cells) => {
      const r = [label]; let ta = 0, tb = 0;
      for (const w of weeks) { const c = cells(w); r.push(c.a, c.b, varPct(c.b, c.a)); ta += c.a; tb += c.b; }
      r.push(ta, tb, varPct(tb, ta));
      return r;
    };
    const rows = list.map((g) => rowOf(g.k, (w) => ({ a: g.w[w]?.[A] || 0, b: g.w[w]?.[B] || 0 }))).filter((r) => r.slice(1).some((v, i) => types[i + 1] !== '%' && v));
    const foot = [rowOf('Total general', (w) => ({ a: list.reduce((s, g) => s + (g.w[w]?.[A] || 0), 0), b: list.reduce((s, g) => s + (g.w[w]?.[B] || 0), 0) }))];
    return { head, rows, foot, types, weeks, yA, yB, unit };
  }

  function renderWeekTable(id, field) {
    const T = weekAOA(field);
    const groups = [...T.weeks.map((w) => `Sem-${w}`), 'Total'];
    const th1 = `<tr><th rowspan="2">${esc(T.head[0])}</th>${groups.map((g) => `<th colspan="3" class="grp">${g}</th>`).join('')}</tr>`;
    const th2 = `<tr>${groups.map(() => `<th class="num">${T.yA}<br>${T.unit}</th><th class="num">${T.yB}<br>${T.unit}</th><th class="num">Var%</th>`).join('')}</tr>`;
    const row = (r) => `<tr>${r.map((v, i) => (i === 0 ? `<td>${esc(v)}</td>` : `<td class="num ${T.types[i] === '%' ? dirCls(v) : ''} ${(i - 1) % 3 === 0 ? 'sep' : ''}">${cell(v, T.types[i])}</td>`)).join('')}</tr>`;
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
    $('#' + canvasId).closest('.card').querySelector('[data-legend]').innerHTML = legend;
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
    const { yA, yB, list, by } = cmp.last;
    const unit = field === 'v' ? 'Soles' : 'Und';
    const A = field === 'v' ? 'vA' : 'uA', B = field === 'v' ? 'vB' : 'uB';
    const head = [by === 'marca' ? 'Marca' : by === 'tipo' ? 'Tipo de venta' : 'Cadena', `${yA} ${unit}`, `${yB} ${unit}`, 'Var S/'.replace('S/', field === 'v' ? 'S/' : 'Und'), 'Var %'];
    const t = field === 'v' ? '$' : 'n';
    const rows = list.filter((g) => g[A] || g[B]).map((g) => [g.k, g[A], g[B], g[B] - g[A], varPct(g[B], g[A])]);
    const sa = rows.reduce((s, r) => s + r[1], 0), sb = rows.reduce((s, r) => s + r[2], 0);
    return { head, rows, foot: [['Total general', sa, sb, sb - sa, varPct(sb, sa)]], types: ['t', t, t, t, '%'], varCols: [3, 4] };
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
    $('#analysis').hidden = v === 'datos';
    $$('[data-analysis]').forEach((s) => s.toggleAttribute('data-off', !state.records.length));
    if (v === 'comparativa') renderCmp();
    Object.values(state.charts).forEach((c) => c.resize());
    window.scrollTo({ top: 0 });
  }

  function init() {
    if (!window.XLSX) { document.body.innerHTML = '<p style="padding:40px">No se pudo cargar el lector de Excel. Revisa tu conexión y recarga la página.</p>'; return; }
    $$('.tab').forEach((t) => t.addEventListener('click', () => setView(t.dataset.view)));
    $('#btnGoCompare').addEventListener('click', () => setView('anual'));
    $('#fileInput').addEventListener('change', (e) => { handleFiles(e.target.files).then(() => { e.target.value = ''; }); });
    $('#btnSave').addEventListener('click', saveToBase);
    $('#btnReload').addEventListener('click', loadServer);
    $('#tblBase').addEventListener('click', (e) => { const b = e.target.closest('[data-del-year]'); if (b) deleteYear(+b.dataset.delYear); });

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
        setSession(r); e.target.reset(); $('#dlgLogin').close(); toast(`Hola, ${r.user.nombre}.`);
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

    ['#fFrom', '#fCob', '#fTipo', '#fChain', '#fBrand'].forEach((s) => $(s).addEventListener('change', renderAll));
    $$('[data-copy]').forEach((b) => b.addEventListener('click', () => copyTable(b.dataset.copy)));
    $('#btnExport').addEventListener('click', exportExcel);
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
    $('#cmpToCut').addEventListener('click', () => {
      const C = lastF.C; cmp.months = new Set([C.m]); cmp.weeks.clear();
      cmp.days = new Set(C.d ? Array.from({ length: C.d }, (_, i) => i + 1) : []); renderCmp();
    });
    ['#cmpYearA', '#cmpYearB', '#cmpBy'].forEach((s) => $(s).addEventListener('change', renderCmp));
    renderFiles();
    renderSession();
    loadServer();
    // Comprueba que la sesión guardada siga siendo válida.
    if (session.token) api('/api/me').then((r) => { session.user = r.user; store.set('user', r.user); renderSession(); }).catch(() => {});
  }

  // Exponer funciones de parseo para pruebas automáticas.
  window.__rva = { parseAmount, parseYM, parseMonth, parseYear, weekOf, detectMapping, detectHeaderRow, state };
  document.addEventListener('DOMContentLoaded', init);
})();
