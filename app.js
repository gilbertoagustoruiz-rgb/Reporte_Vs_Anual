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
  function okYM(y, m) { return y >= 1990 && y <= 2100 && m >= 1 && m <= 12 ? { y, m } : null; }

  function monthWord(w) { return MONTH_WORDS[norm(w).replace(/\s/g, '')] || MONTH_WORDS[norm(w).slice(0, 3)] || null; }

  /** Devuelve {y, m} desde una fecha en cualquier formato habitual, o null. */
  function parseYM(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) {
      if (isNaN(v)) return null;
      const d = new Date(v.getTime() + 12 * 3600e3); // evita desfases de zona horaria en fechas de Excel
      return okYM(d.getFullYear(), d.getMonth() + 1);
    }
    if (typeof v === 'number') {
      if (v >= 199001 && v <= 210012 && v % 100 >= 1 && v % 100 <= 12) return okYM(Math.floor(v / 100), v % 100); // aaaamm
      if (v >= 19900101 && v <= 21001231) return okYM(Math.floor(v / 10000), Math.floor(v / 100) % 100); // aaaammdd
      if (v > 20000 && v < 80000) { const p = XLSX.SSF.parse_date_code(v); return p ? okYM(p.y, p.m) : null; } // serial Excel
      return null;
    }
    const s = String(v).trim();
    let m;
    if ((m = s.match(/^(\d{4})[-/.](\d{1,2})(?:[-/.](\d{1,2}))?(?:[ T].*)?$/))) return okYM(+m[1], +m[2]);
    if ((m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})(?:[ T].*)?$/))) {
      let a = +m[1], b = +m[2];
      // Perú: día/mes/año. Si el segundo número no puede ser mes, se asume mes/día/año.
      const mo = b <= 12 ? b : a;
      return okYM(fixYear(m[3]), mo);
    }
    if ((m = s.match(/^(\d{1,2})[-/.](\d{4})$/))) return okYM(+m[2], +m[1]);
    if ((m = s.match(/^(\d{4})(\d{2})(\d{2})?$/))) return okYM(+m[1], +m[2]);
    if ((m = s.match(/^([a-zA-ZáéíóúÁÉÍÓÚ.]+)[\s\-/.]*(?:de[\s]+|del[\s]+)?(\d{2,4})$/))) {
      const mo = monthWord(m[1]); return mo ? okYM(fixYear(m[2]), mo) : null;
    }
    if ((m = s.match(/^(\d{4})[\s\-/.]*([a-zA-ZáéíóúÁÉÍÓÚ.]+)$/))) {
      const mo = monthWord(m[2]); return mo ? okYM(+m[1], mo) : null;
    }
    if ((m = s.match(/^(\d{1,2})\s+de\s+([a-zA-ZáéíóúÁÉÍÓÚ]+)(?:\s+(?:de|del)\s+(\d{2,4}))?$/i))) {
      const mo = monthWord(m[2]); return mo && m[3] ? okYM(fixYear(m[3]), mo) : null;
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
    { key: 'cadena', label: 'Cadena' },
    { key: 'tienda', label: 'Tienda / PDV' },
    { key: 'marca', label: 'Marca' },
    { key: 'unidades', label: 'Unidades' },
  ];

  const RULES = {
    venta: { pats: [/^(venta|ventas|venta s|venta soles|venta neta|venta total|venta valor|sell out|sellout|venta pen|monto venta|importe venta)$/, /\bventa|sell ?out|\bsoles\b/, /\b(monto|importe|valor|total|neto|facturacion)\b/], not: /cant|unid|\bund|qty|cuota|meta|objetivo|ppto|presupuesto|precio|costo|margen|%|porc|var|cod|fecha|mes|ano/ },
    fecha: { pats: [/^(fecha|fecha venta|fecha de venta|fec venta|date|dia)$/, /\bfecha|\bdate\b|\bfec\b/], not: /registro|carga|actualiz|creacion|nacim|ingreso|alta/ },
    anio: { pats: [/^(ano|anio|year|periodo ano|ano venta)$/, /\bano\b|\banio\b|\byear\b/], not: /mes|fecha/ },
    mes: { pats: [/^(mes|month|mes venta|nombre mes|mes nombre)$/, /\bmes\b|\bmonth\b|\bperiodo\b/], not: /fecha|ano/ },
    cadena: { pats: [/^(cadena|cadena comercial|retail|cliente|razon social|chain)$/, /\bcadena|\bretail\b|\bcliente\b|\bchain\b|\bcanal\b/], not: /cod|id\b|ruc/ },
    tienda: { pats: [/^(tienda|pdv|punto de venta|local|sucursal|store|nombre tienda|nombre pdv|nombre local|desc tienda|descripcion tienda|tienda nombre)$/, /\b(tienda|pdv|punto de venta|local|sucursal|store)\b/], not: /\bcod|codigo|\bid\b|cadena/ },
    marca: { pats: [/^(marca|brand)$/, /\bmarca\b|\bbrand\b/], not: /cod/ },
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
    for (const role of ['venta', 'fecha', 'anio', 'mes', 'cadena', 'tienda', 'marca', 'unidades']) {
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

  function pickSheet(wb) {
    let best = wb.SheetNames[0], bestScore = -1;
    for (const n of wb.SheetNames) {
      const rows = sheetRows(wb, n);
      if (!rows.length) continue;
      const h = rows[detectHeaderRow(rows)] || [];
      const m = detectMapping(h.map((x) => (x == null ? '' : String(x))), rows);
      const score = Object.keys(m).length * 1000 + Math.min(rows.length, 999);
      if (score > bestScore) { bestScore = score; best = n; }
    }
    return best;
  }

  function prepareFile(f, opts = {}) {
    const sheet = opts.sheet ?? f.sheet ?? pickSheet(f.wb);
    const all = sheetRows(f.wb, sheet);
    const headerRow = opts.headerRow ?? detectHeaderRow(all);
    const rawHeaders = all[headerRow] || [];
    const width = all.reduce((w, r) => Math.max(w, r ? r.length : 0), 0);
    const headers = Array.from({ length: width }, (_, i) => rawHeaders[i]);
    const headerText = headers.map((h, i) => (h == null || h === '' ? `Columna ${XLSX.utils.encode_col(i)}` : h instanceof Date ? `${MESES[h.getMonth()]}-${h.getFullYear()}` : String(h).trim()));
    const rows = all.slice(headerRow + 1);
    const sig = 'map:' + headerText.map(norm).join('|');
    let mapping = opts.mapping ?? store.get(sig, null) ?? detectMapping(headerText, rows);
    Object.assign(f, { sheet, headerRow, headers, headerText, rows, mapping, sig, totalRows: all.length });
    buildFileAgg(f);
  }

  function buildFileAgg(f) {
    const { rows, mapping: mp } = f;
    const agg = new Map();
    const warnings = [];
    let used = 0, noDate = 0, noAmt = 0;
    const years = new Set();
    let total = 0;
    const wide = mp.venta == null ? detectWideColumns(f.headers, mp) : [];
    f.wide = wide.length >= 2;
    f.error = null;

    if (mp.venta == null && !f.wide) {
      f.error = 'No encontré la columna de venta. Elígela en la lista "Venta (S/)".';
    } else if (!f.wide && mp.fecha == null && mp.mes == null && mp.anio == null) {
      f.error = 'No encontré la fecha ni el año/mes. Elige la columna "Fecha" o las columnas "Año" y "Mes".';
    }

    if (!f.error) {
      const val = (r, k) => (mp[k] == null ? null : r[mp[k]]);
      const label = (v, d) => { const s = v == null ? '' : v instanceof Date ? v.toLocaleDateString('es-PE') : String(v).trim(); return s || d; };
      const add = (cad, tie, mar, y, m, v, u) => {
        const key = cad + '\u0001' + tie + '\u0001' + mar + '\u0001' + y + '\u0001' + m;
        const cur = agg.get(key);
        if (cur) { cur.v += v; cur.u += u; } else agg.set(key, { v, u });
        years.add(y); total += v; used++;
      };
      for (const r of rows) {
        if (!r) continue;
        const cad = label(val(r, 'cadena'), '(Sin cadena)');
        const tie = label(val(r, 'tienda'), '(Sin tienda)');
        const mar = label(val(r, 'marca'), '');
        if (f.wide) {
          if (mp.tienda != null && val(r, 'tienda') == null && mp.cadena != null && val(r, 'cadena') == null) continue;
          const tieN = norm(tie);
          if (/^(total|sub ?total|gran total)\b/.test(tieN) || /^(total|sub ?total|gran total)\b/.test(norm(cad))) continue;
          for (const c of wide) {
            const v = parseAmount(r[c.i]);
            if (v == null || v === 0) continue;
            add(cad, tie, mar, c.y, c.m, v, 0);
          }
          continue;
        }
        const v = parseAmount(val(r, 'venta'));
        if (v == null) { if (r.some((c) => c != null && c !== '')) noAmt++; continue; }
        let ym = null;
        if (mp.fecha != null) ym = parseYM(val(r, 'fecha'));
        if (!ym && mp.anio != null) {
          const y = parseYear(val(r, 'anio'));
          const m = mp.mes != null ? parseMonth(val(r, 'mes')) : null;
          if (y && m) ym = { y, m };
        }
        if (!ym && mp.mes != null) ym = parseYM(val(r, 'mes'));
        if (!ym) { noDate++; continue; }
        if (/^(total|sub ?total|gran total)\b/.test(norm(tie)) || /^(total|sub ?total|gran total)\b/.test(norm(cad))) continue;
        const u = mp.unidades != null ? parseAmount(val(r, 'unidades')) || 0 : 0;
        add(cad, tie, mar, ym.y, ym.m, v, u);
      }
      if (!used) f.error = 'No pude leer ninguna fila con venta y fecha. Revisa la fila de encabezados y las columnas elegidas.';
      if (noDate) warnings.push(`${num(noDate)} filas sin fecha válida se omitieron.`);
      if (noAmt > 0 && noAmt > used * 0.05) warnings.push(`${num(noAmt)} filas sin monto de venta se omitieron.`);
      if (!f.wide && mp.tienda == null) warnings.push('No hay columna de tienda: el detalle por tienda saldrá agrupado.');
      if (f.wide && mp.anio == null && mp.mes == null && mp.fecha == null) warnings.push(`Formato ancho: ${wide.length} columnas de mes detectadas como ventas.`);
    }
    Object.assign(f, { agg, warnings, used, years: [...years].sort(), total });
  }

  function rebuildAll() {
    const agg = new Map();
    let hasBrand = false, hasStore = false;
    for (const f of state.files) {
      if (f.error) continue;
      if (f.mapping.marca != null) hasBrand = true;
      if (f.mapping.tienda != null) hasStore = true;
      for (const [k, o] of f.agg) {
        const cur = agg.get(k);
        if (cur) { cur.v += o.v; cur.u += o.u; } else agg.set(k, { v: o.v, u: o.u });
      }
    }
    const years = new Set();
    const records = [];
    for (const [k, o] of agg) {
      const [cadena, tienda, marca, y, m] = k.split('\u0001');
      const rec = { cadena, tienda, marca, y: +y, m: +m, v: o.v, u: o.u };
      records.push(rec);
      if (o.v) years.add(rec.y);
    }
    Object.assign(state, { records, years: [...years].sort((a, b) => a - b), hasBrand, hasStore });
    syncFilterOptions();
    renderFiles();
    renderAll();
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
        state.files.push({ id: ++fileSeq, name: file.name, size: file.size, wb: null, error: 'No se pudo leer el archivo: ' + (err.message || err), headers: [], headerText: [], mapping: {}, warnings: [], years: [], used: 0, total: 0 });
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
      const sheetSel = f.wb && f.wb.SheetNames.length > 1
        ? `<select data-act="sheet" data-id="${f.id}" title="Hoja">${f.wb.SheetNames.map((n) => `<option ${n === f.sheet ? 'selected' : ''}>${esc(n)}</option>`).join('')}</select>` : '';
      const hdrSel = f.wb
        ? `<select data-act="header" data-id="${f.id}" title="Fila de encabezados">${Array.from({ length: Math.min(30, f.totalRows || 1) }, (_, i) => `<option value="${i}" ${i === f.headerRow ? 'selected' : ''}>Encabezados en fila ${i + 1}</option>`).join('')}</select>` : '';
      const opts = (sel) => `<option value="">(ninguna)</option>` + f.headerText.map((h, i) => `<option value="${i}" ${sel === i ? 'selected' : ''}>${esc(h)}</option>`).join('');
      const map = f.wb ? `<div class="map-grid">${ROLES.map((r) => `<div><label>${r.label}</label><select data-act="map" data-role="${r.key}" data-id="${f.id}">${opts(f.mapping[r.key])}</select></div>`).join('')}</div>` : '';
      const info = ok
        ? `<span class="badge ok">✓ ${num(f.used)} registros</span><span class="badge">Años: ${f.years.join(', ') || '—'}</span><span class="badge">${money(f.total)}</span>${f.wide ? '<span class="badge">Formato ancho</span>' : ''}`
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
    if (act === 'sheet') prepareFile(f, { sheet: el.value, headerRow: undefined });
    else if (act === 'header') prepareFile(f, { sheet: f.sheet, headerRow: +el.value, mapping: undefined });
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
      buildFileAgg(f);
    }
    rebuildAll();
  }

  // ---------------------------------------------------------------- filtros
  function syncFilterOptions() {
    const ys = state.years;
    const setOpts = (sel, items, val) => {
      sel.innerHTML = items.map(([v, t]) => `<option value="${esc(v)}">${esc(t)}</option>`).join('');
      if (items.some(([v]) => String(v) === String(val))) sel.value = val;
    };
    const yA = $('#fYearA'), yB = $('#fYearB');
    const prevB = +yB.value, prevA = +yA.value;
    const defB = ys.includes(prevB) ? prevB : ys[ys.length - 1];
    const defA = ys.includes(prevA) && prevA !== defB ? prevA : ys.filter((y) => y < defB).pop() ?? ys[0];
    setOpts(yA, ys.map((y) => [y, y]), defA);
    setOpts(yB, ys.map((y) => [y, y]), defB);

    const chains = [...new Set(state.records.map((r) => r.cadena))].sort((a, b) => a.localeCompare(b, 'es'));
    setOpts($('#fChain'), [['', 'Todas'], ...chains.map((c) => [c, c])], $('#fChain').value);
    const brands = [...new Set(state.records.map((r) => r.marca).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
    setOpts($('#fBrand'), [['', 'Todas'], ...brands.map((c) => [c, c])], $('#fBrand').value);
    $('#fBrandWrap').hidden = !brands.length;

    if (!$('#fMonthFrom').options.length) {
      setOpts($('#fMonthFrom'), MESES_LARGO.map((m, i) => [i + 1, m]), 1);
      setOpts($('#fMonthTo'), MESES_LARGO.map((m, i) => [i + 1, m]), 12);
    }
  }

  function currentFilters() {
    const yearA = +$('#fYearA').value, yearB = +$('#fYearB').value;
    const chain = $('#fChain').value, brand = $('#fBrand').value;
    const base = state.records.filter((r) => (!chain || r.cadena === chain) && (!brand || r.marca === brand));
    const period = $('#fPeriod').value;
    let months;
    if (period === 'full') months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
    else if (period === 'custom') {
      const a = +$('#fMonthFrom').value, b = +$('#fMonthTo').value;
      const lo = Math.min(a, b), hi = Math.max(a, b);
      months = Array.from({ length: hi - lo + 1 }, (_, i) => lo + i);
    } else {
      const lastB = Math.max(0, ...base.filter((r) => r.y === yearB && r.v).map((r) => r.m));
      const lastA = Math.max(0, ...base.filter((r) => r.y === yearA && r.v).map((r) => r.m));
      const last = lastB || lastA || 12;
      months = Array.from({ length: last }, (_, i) => i + 1);
    }
    const mset = new Set(months);
    const recs = base.filter((r) => (r.y === yearA || r.y === yearB) && mset.has(r.m));
    return { yearA, yearB, chain, brand, period, months, recs };
  }

  function periodText(F) {
    const span = F.months.length === 12 ? 'Ene–Dic' : F.months.length === 1 ? MESES[F.months[0] - 1] : `${MESES[F.months[0] - 1]}–${MESES[F.months[F.months.length - 1] - 1]}`;
    let t = `Comparando <b>${span} ${F.yearB}</b> contra <b>${span} ${F.yearA}</b>`;
    if (F.chain) t += ` · Cadena: <b>${esc(F.chain)}</b>`;
    if (F.brand) t += ` · Marca: <b>${esc(F.brand)}</b>`;
    if (F.yearA === F.yearB) t += ' · <span class="down">Elige dos años distintos</span>';
    return t;
  }

  // ---------------------------------------------------------------- cálculos
  function statusOf(a, b) {
    if (a <= 0 && b > 0) return 'nueva';
    if (a > 0 && b <= 0) return 'perdida';
    if (b > a) return 'crece';
    if (b < a) return 'cae';
    return 'igual';
  }
  const STATUS_LABEL = { crece: '▲ Crece', cae: '▼ Cae', nueva: '● Nueva', perdida: '✕ Sin venta', igual: '= Igual' };

  function compute(F) {
    const byMonth = new Map(F.months.map((m) => [m, { a: 0, b: 0 }]));
    const byChain = new Map();
    const byStore = new Map();
    let A = 0, B = 0, uA = 0, uB = 0;
    for (const r of F.recs) {
      const isB = r.y === F.yearB;
      const isA = r.y === F.yearA;
      const mo = byMonth.get(r.m);
      let c = byChain.get(r.cadena);
      if (!c) byChain.set(r.cadena, (c = { cadena: r.cadena, a: 0, b: 0 }));
      const sk = r.cadena + '\u0001' + r.tienda;
      let s = byStore.get(sk);
      if (!s) byStore.set(sk, (s = { cadena: r.cadena, tienda: r.tienda, a: 0, b: 0, mb: {} }));
      if (isB) { B += r.v; uB += r.u; mo.b += r.v; c.b += r.v; s.b += r.v; s.mb[r.m] = (s.mb[r.m] || 0) + r.v; }
      if (isA && F.yearA !== F.yearB) { A += r.v; uA += r.u; mo.a += r.v; c.a += r.v; s.a += r.v; }
      else if (isA) { A += r.v; mo.a += r.v; c.a += r.v; s.a += r.v; }
    }
    const stores = [...byStore.values()].map((s) => {
      const st = statusOf(s.a, s.b);
      return { ...s, d: s.b - s.a, p: s.a > 0 ? (s.b - s.a) / s.a : null, st };
    });
    const chains = [...byChain.values()].map((c) => {
      const ss = stores.filter((s) => s.cadena === c.cadena);
      return {
        ...c, d: c.b - c.a, p: c.a > 0 ? (c.b - c.a) / c.a : null,
        nA: ss.filter((s) => s.a > 0).length, nB: ss.filter((s) => s.b > 0).length,
        up: ss.filter((s) => s.st === 'crece' || s.st === 'nueva').length,
        down: ss.filter((s) => s.st === 'cae' || s.st === 'perdida').length,
      };
    }).sort((x, y) => y.b - x.b || y.a - x.a);
    const count = (st) => stores.filter((s) => s.st === st).length;
    return {
      A, B, uA, uB, d: B - A, p: A > 0 ? (B - A) / A : null,
      months: F.months.map((m) => ({ m, ...byMonth.get(m) })),
      chains, stores,
      nA: stores.filter((s) => s.a > 0).length, nB: stores.filter((s) => s.b > 0).length,
      crece: count('crece'), cae: count('cae'), nueva: count('nueva'), perdida: count('perdida'),
    };
  }

  // ---------------------------------------------------------------- render
  let lastF = null, lastR = null;

  function renderAll() {
    if (!state.records.length) {
      Object.values(state.charts).forEach((c) => c.destroy());
      state.charts = {};
      return;
    }
    $('#fRangeWrap').hidden = $('#fPeriod').value !== 'custom';
    const F = currentFilters();
    const R = compute(F);
    lastF = F; lastR = R;
    $('#periodNote').innerHTML = periodText(F);
    $('#periodNote2').innerHTML = periodText(F);
    renderKpis(F, R);
    renderCharts(F, R);
    renderChainTable(F, R);
    renderMonthTable(F, R);
    renderStoreTable();
  }

  function kpi(label, value, sub, cls = '', sw = '') {
    return `<div class="kpi"><div class="k-label">${sw}${label}</div><div class="k-value ${cls}">${value}</div>${sub ? `<div class="k-sub">${sub}</div>` : ''}</div>`;
  }

  function renderKpis(F, R) {
    const sw = (c) => `<span class="swatch" style="background:var(${c})"></span>`;
    const dir = R.d > 0 ? 'up' : R.d < 0 ? 'down' : '';
    const arrow = R.d > 0 ? '▲ ' : R.d < 0 ? '▼ ' : '';
    $('#kpis').innerHTML = [
      kpi(`Venta ${F.yearA}`, money(R.A), R.uA ? `${num(R.uA)} unidades` : `${num(R.nA)} tiendas con venta`, '', sw('--series-a')),
      kpi(`Venta ${F.yearB}`, money(R.B), R.uB ? `${num(R.uB)} unidades` : `${num(R.nB)} tiendas con venta`, '', sw('--series-b')),
      kpi('Variación S/', arrow + signMoney(R.d), `${F.yearB} vs ${F.yearA}`, dir),
      kpi('Variación %', arrow + pct(R.p), R.A > 0 ? 'sobre el año base' : 'sin venta en el año base', dir),
      kpi('Tiendas que crecen', num(R.crece + R.nueva), `${num(R.nueva)} nuevas`, 'up'),
      kpi('Tiendas que caen', num(R.cae + R.perdida), `${num(R.perdida)} sin venta en ${F.yearB}`, 'down'),
    ].join('');
  }

  function cssVar(n) { return getComputedStyle(document.documentElement).getPropertyValue(n).trim(); }

  function renderCharts(F, R) {
    const colA = cssVar('--series-a'), colB = cssVar('--series-b');
    $('#chartLegend').innerHTML = `<span><i class="swatch" style="background:${colA}"></i>${F.yearA}</span><span><i class="swatch" style="background:${colB}"></i>${F.yearB}</span>`;
    const grid = '#eef0f3', tick = '#7a8292';
    const moneyTick = (v) => (Math.abs(v) >= 1e6 ? 'S/ ' + (v / 1e6).toFixed(1) + ' M' : Math.abs(v) >= 1e3 ? 'S/ ' + Math.round(v / 1e3) + ' K' : 'S/ ' + v);
    const common = {
      responsive: true, maintainAspectRatio: false, animation: { duration: 250 },
      plugins: { legend: { display: false } },
    };

    const mData = {
      labels: R.months.map((x) => MESES[x.m - 1]),
      datasets: [
        { label: String(F.yearA), data: R.months.map((x) => x.a), backgroundColor: colA, borderRadius: 4, borderSkipped: 'bottom', categoryPercentage: 0.7, barPercentage: 0.9 },
        { label: String(F.yearB), data: R.months.map((x) => x.b), backgroundColor: colB, borderRadius: 4, borderSkipped: 'bottom', categoryPercentage: 0.7, barPercentage: 0.9 },
      ],
    };
    const mOpts = {
      ...common,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (c) => ` ${c.dataset.label}: ${money(c.parsed.y)}`,
            footer: (items) => {
              const a = items.find((i) => i.datasetIndex === 0)?.parsed.y || 0;
              const b = items.find((i) => i.datasetIndex === 1)?.parsed.y || 0;
              return a > 0 ? `Var: ${signMoney(b - a)} (${pct((b - a) / a)})` : '';
            },
          },
        },
      },
      scales: {
        x: { grid: { display: false }, ticks: { color: tick } },
        y: { beginAtZero: true, grid: { color: grid }, border: { display: false }, ticks: { color: tick, callback: moneyTick } },
      },
    };
    if (state.charts.m) { state.charts.m.data = mData; state.charts.m.options = mOpts; state.charts.m.update(); }
    else state.charts.m = new Chart($('#chartMonthly'), { type: 'bar', data: mData, options: mOpts });

    const top = R.chains.filter((c) => c.a > 0);
    const good = cssVar('--good'), bad = cssVar('--bad');
    const cData = {
      labels: top.map((c) => (c.cadena.length > 22 ? c.cadena.slice(0, 21) + '…' : c.cadena)),
      datasets: [{ data: top.map((c) => +(c.p * 100).toFixed(1)), backgroundColor: top.map((c) => (c.p >= 0 ? good : bad)), borderRadius: 4, barPercentage: 0.75 }],
    };
    const cOpts = {
      ...common,
      indexAxis: 'y',
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { title: (i) => top[i[0].dataIndex].cadena, label: (c) => { const x = top[c.dataIndex]; return ` ${pct(x.p)} · ${signMoney(x.d)}  (${money(x.a)} → ${money(x.b)})`; } } },
      },
      scales: {
        x: { grid: { color: grid }, border: { display: false }, ticks: { color: tick, callback: (v) => v + '%' } },
        y: { grid: { display: false }, ticks: { color: '#4c5361', autoSkip: false } },
      },
    };
    $('#chartChains').parentElement.style.height = Math.max(220, top.length * 26 + 60) + 'px';
    if (state.charts.c) { state.charts.c.data = cData; state.charts.c.options = cOpts; state.charts.c.update(); }
    else state.charts.c = new Chart($('#chartChains'), { type: 'bar', data: cData, options: cOpts });
  }

  const tdNum = (v, cls = '') => `<td class="num ${cls}">${v}</td>`;
  const dirCls = (d) => (d > 0 ? 'up' : d < 0 ? 'down' : '');

  // Datos tabulares (se reutilizan para copiar y exportar)
  function chainAOA(F, R) {
    const head = ['Cadena', `Venta ${F.yearA}`, `Venta ${F.yearB}`, 'Var S/', 'Var %', `Tiendas ${F.yearA}`, `Tiendas ${F.yearB}`, 'Crecen', 'Caen', `Part. ${F.yearB}`];
    const rows = R.chains.map((c) => [c.cadena, c.a, c.b, c.d, c.p, c.nA, c.nB, c.up, c.down, R.B ? c.b / R.B : null]);
    const tot = ['Total', R.A, R.B, R.d, R.p, R.nA, R.nB, R.crece + R.nueva, R.cae + R.perdida, R.B ? 1 : null];
    return { head, rows, tot, types: ['t', '$', '$', '$', '%', 'n', 'n', 'n', 'n', 'p'] };
  }
  function monthAOA(F, R) {
    const head = ['Mes', `Venta ${F.yearA}`, `Venta ${F.yearB}`, 'Var S/', 'Var %'];
    const rows = R.months.map((x) => [MESES_LARGO[x.m - 1], x.a, x.b, x.b - x.a, x.a > 0 ? (x.b - x.a) / x.a : null]);
    return { head, rows, tot: ['Total', R.A, R.B, R.d, R.p], types: ['t', '$', '$', '$', '%'] };
  }
  function storeRows() {
    const F = lastF, R = lastR;
    const q = norm($('#fSearch').value);
    const st = $('#fStatus').value;
    let rows = R.stores.filter((s) => (s.a || s.b));
    if (q) rows = rows.filter((s) => norm(s.tienda).includes(q) || norm(s.cadena).includes(q));
    if (st) rows = rows.filter((s) => s.st === st);
    const { key, dir } = state.sort;
    rows.sort((x, y) => {
      let a = x[key], b = y[key];
      if (key.startsWith('m')) { a = x.mb[+key.slice(1)] || 0; b = y.mb[+key.slice(1)] || 0; }
      if (a == null) return 1; if (b == null) return -1;
      return (typeof a === 'string' ? a.localeCompare(b, 'es') : a - b) * dir;
    });
    return { F, R, rows };
  }
  function storeAOA() {
    const { F, rows } = storeRows();
    const withMonths = $('#fMonthCols').value === 'yes';
    const head = ['Cadena', 'Tienda', `Venta ${F.yearA}`, `Venta ${F.yearB}`, 'Var S/', 'Var %', 'Estado', ...(withMonths ? F.months.map((m) => `${MESES[m - 1]} ${F.yearB}`) : [])];
    const keys = ['cadena', 'tienda', 'a', 'b', 'd', 'p', 'st', ...(withMonths ? F.months.map((m) => 'm' + m) : [])];
    const types = ['t', 't', '$', '$', '$', '%', 't', ...(withMonths ? F.months.map(() => '$') : [])];
    const data = rows.map((s) => [s.cadena, s.tienda, s.a, s.b, s.d, s.p, STATUS_LABEL[s.st].replace(/^\S+\s/, ''), ...(withMonths ? F.months.map((m) => s.mb[m] || 0) : [])]);
    const sum = (i) => rows.reduce((t, s) => t + (s[i] || 0), 0);
    const A = sum('a'), B = sum('b');
    const tot = ['Total', `${num(rows.length)} tiendas`, A, B, B - A, A > 0 ? (B - A) / A : null, '', ...(withMonths ? F.months.map((m) => rows.reduce((t, s) => t + (s.mb[m] || 0), 0)) : [])];
    return { head, rows: data, tot, types, keys, raw: rows };
  }

  function cell(v, t) {
    if (t === '$') return money(v);
    if (t === '%') return pct(v);
    if (t === 'p') return v == null ? '—' : fmtPct.format(v);
    if (t === 'n') return num(v);
    return esc(v);
  }

  function renderSimpleTable(id, T, varIdx = [3, 4]) {
    const th = T.head.map((h, i) => `<th class="${T.types[i] === 't' ? '' : 'num'}">${esc(h)}</th>`).join('');
    const body = T.rows.map((r) => `<tr>${r.map((v, i) => (T.types[i] === 't' ? `<td>${esc(v)}</td>` : tdNum(cell(v, T.types[i]), varIdx.includes(i) ? dirCls(r[3]) : ''))).join('')}</tr>`).join('');
    const foot = `<tr>${T.tot.map((v, i) => (T.types[i] === 't' ? `<td>${esc(v)}</td>` : tdNum(cell(v, T.types[i]), varIdx.includes(i) ? dirCls(T.tot[3]) : ''))).join('')}</tr>`;
    $('#' + id).innerHTML = `<thead><tr>${th}</tr></thead><tbody>${body}</tbody><tfoot>${foot}</tfoot>`;
  }
  function renderChainTable(F, R) { renderSimpleTable('tblChains', chainAOA(F, R)); }
  function renderMonthTable(F, R) { renderSimpleTable('tblMonths', monthAOA(F, R)); }

  function renderStoreTable() {
    if (!lastR) return;
    const T = storeAOA();
    const { key, dir } = state.sort;
    const th = T.head.map((h, i) => `<th class="${T.types[i] === 't' ? '' : 'num'}" data-sort="${T.keys[i]}">${esc(h)}${T.keys[i] === key ? `<span class="arrow">${dir > 0 ? '▲' : '▼'}</span>` : ''}</th>`).join('');
    const shown = T.raw;
    const body = shown.map((s, ri) => {
      const r = T.rows[ri];
      return `<tr>${r.map((v, i) => {
        const k = T.keys[i];
        if (k === 'st') return `<td><span class="pill ${s.st}">${STATUS_LABEL[s.st]}</span></td>`;
        if (T.types[i] === 't') return `<td>${esc(v)}</td>`;
        const cls = (k === 'd' || k === 'p') ? dirCls(s.d) : '';
        return tdNum(cell(v, T.types[i]), cls);
      }).join('')}</tr>`;
    }).join('');
    const foot = `<tr>${T.tot.map((v, i) => (T.types[i] === 't' ? `<td>${esc(v)}</td>` : tdNum(cell(v, T.types[i]), (T.keys[i] === 'd' || T.keys[i] === 'p') ? dirCls(T.tot[4]) : ''))).join('')}</tr>`;
    $('#tblStores').innerHTML = `<thead><tr>${th}</tr></thead><tbody>${body || `<tr><td colspan="${T.head.length}" class="muted">Sin tiendas para estos filtros.</td></tr>`}</tbody><tfoot>${foot}</tfoot>`;
    $('#storeCount').textContent = `· ${num(T.raw.length)}`;
  }

  // ---------------------------------------------------------------- copiar / exportar
  function aoaFor(id) {
    if (id === 'tblChains') return chainAOA(lastF, lastR);
    if (id === 'tblMonths') return monthAOA(lastF, lastR);
    return storeAOA();
  }
  async function copyTable(id) {
    const T = aoaFor(id);
    const fmt = (v, t) => (v == null ? '' : t === '%' || t === 'p' ? (v * 100).toFixed(1) + '%' : t === '$' || t === 'n' ? String(Math.round(v * 100) / 100) : String(v));
    const lines = [T.head, ...T.rows, T.tot].map((r) => r.map((v, i) => fmt(v, T.types[i])).join('\t'));
    try { await navigator.clipboard.writeText(lines.join('\n')); toast('Tabla copiada: pégala en Excel.'); }
    catch { toast('No se pudo copiar en este navegador.'); }
  }

  function sheetFrom(T, title) {
    const aoa = [[title], [], T.head, ...T.rows, T.tot];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = 3; R <= range.e.r; R++) {
      T.types.forEach((t, C) => {
        const c = ws[XLSX.utils.encode_cell({ r: R, c: C })];
        if (!c || c.t !== 'n') return;
        if (t === '$') c.v = Math.round(c.v * 100) / 100;
        c.z = t === '%' || t === 'p' ? '0.0%' : t === '$' ? '#,##0.00' : '#,##0';
      });
    }
    ws['!cols'] = T.head.map((h, i) => ({ wch: i < 2 && T.types[i] === 't' ? 28 : Math.max(12, String(h).length + 2) }));
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 2, c: 0 }, e: { r: 2 + T.rows.length, c: T.head.length - 1 } }) };
    return ws;
  }

  function exportExcel() {
    if (!lastR) return;
    const F = lastF, R = lastR;
    const plain = periodText(F).replace(/<[^>]+>/g, '');
    const wb = XLSX.utils.book_new();
    const k = [
      ['Reporte Vs Anual'], [plain], [],
      ['Indicador', String(F.yearA), String(F.yearB), 'Var S/', 'Var %'],
      ['Venta (S/)', R.A, R.B, R.d, R.p],
      ['Tiendas con venta', R.nA, R.nB, R.nB - R.nA, R.nA ? (R.nB - R.nA) / R.nA : null],
      [], ['Tiendas que crecen', R.crece], ['Tiendas que caen', R.cae], ['Tiendas nuevas', R.nueva], [`Sin venta en ${F.yearB}`, R.perdida],
    ];
    const ws0 = XLSX.utils.aoa_to_sheet(k);
    ['B5', 'C5', 'D5'].forEach((a) => ws0[a] && (ws0[a].z = '#,##0.00'));
    ['E5', 'E6'].forEach((a) => ws0[a] && (ws0[a].z = '0.0%'));
    ws0['!cols'] = [{ wch: 26 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, ws0, 'Resumen');
    XLSX.utils.book_append_sheet(wb, sheetFrom(chainAOA(F, R), 'Resumen por cadena · ' + plain), 'Por cadena');
    XLSX.utils.book_append_sheet(wb, sheetFrom(monthAOA(F, R), 'Detalle mensual · ' + plain), 'Mensual');
    XLSX.utils.book_append_sheet(wb, sheetFrom(storeAOA(), 'Detalle por tienda · ' + plain), 'Por tienda');
    XLSX.writeFile(wb, `Reporte_Vs_Anual_${F.yearB}_vs_${F.yearA}${F.chain ? '_' + F.chain.replace(/[^\w]+/g, '_') : ''}.xlsx`);
  }

  // ---------------------------------------------------------------- eventos
  function setView(v) {
    state.view = v;
    $$('.tab').forEach((t) => t.classList.toggle('active', t.dataset.view === v));
    $$('.view').forEach((s) => s.classList.toggle('active', s.id === 'view-' + v));
    if (v !== 'datos') Object.values(state.charts).forEach((c) => c.resize());
    window.scrollTo({ top: 0 });
  }

  function init() {
    if (!window.XLSX) { document.body.innerHTML = '<p style="padding:40px">No se pudo cargar el lector de Excel. Revisa tu conexión y recarga la página.</p>'; return; }
    $$('.tab').forEach((t) => t.addEventListener('click', () => setView(t.dataset.view)));
    $('#btnGoCompare').addEventListener('click', () => setView('comparativo'));
    ['#fileInput', '#fileInputTop'].forEach((s) => $(s).addEventListener('change', (e) => { handleFiles(e.target.files).then(() => { e.target.value = ''; }); if (s === '#fileInputTop') setView('datos'); }));

    const dz = $('#dropzone');
    ['dragenter', 'dragover'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); if (ev === 'drop' || e.target === document.documentElement) dz.classList.remove('over'); }));
    document.addEventListener('drop', (e) => { if (e.dataTransfer?.files?.length) { setView('datos'); handleFiles(e.dataTransfer.files); } });

    $('#filesList').addEventListener('change', onFilesPanelChange);
    $('#filesList').addEventListener('click', (e) => {
      const el = e.target.closest('[data-act="remove"]');
      if (!el) return;
      state.files = state.files.filter((f) => f.id !== +el.dataset.id);
      rebuildAll();
    });
    $('#btnClear').addEventListener('click', () => { state.files = []; rebuildAll(); });

    ['#fYearA', '#fYearB', '#fPeriod', '#fMonthFrom', '#fMonthTo', '#fChain', '#fBrand'].forEach((s) => $(s).addEventListener('change', () => { renderAll(); }));
    let tSearch;
    $('#fSearch').addEventListener('input', () => { clearTimeout(tSearch); tSearch = setTimeout(() => { renderStoreTable(); }, 150); });
    ['#fStatus', '#fMonthCols'].forEach((s) => $(s).addEventListener('change', () => { renderStoreTable(); }));
    $('#tblStores').addEventListener('click', (e) => {
      const th = e.target.closest('th[data-sort]');
      if (!th) return;
      const k = th.dataset.sort;
      state.sort = state.sort.key === k ? { key: k, dir: -state.sort.dir } : { key: k, dir: ['cadena', 'tienda', 'st'].includes(k) ? 1 : -1 };
      renderStoreTable();
    });
    $$('[data-copy]').forEach((b) => b.addEventListener('click', () => copyTable(b.dataset.copy)));
    $('#btnExport').addEventListener('click', exportExcel);
    renderFiles();
  }

  // Exponer funciones de parseo para pruebas automáticas.
  window.__rva = { parseAmount, parseYM, parseMonth, parseYear, detectMapping, detectHeaderRow, state };
  document.addEventListener('DOMContentLoaded', init);
})();
