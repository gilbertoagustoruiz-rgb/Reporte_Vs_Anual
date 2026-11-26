/* API del Reporte Vs Anual.
 * Guarda las ventas (ya agregadas por cadena/tienda/marca/año/mes/semana) en PostgreSQL,
 * en el esquema "reporte_vs_anual", para que todas las personas vean el mismo reporte.
 *
 * Variables de entorno:
 *   DATABASE_URL     Cadena de conexión de PostgreSQL (obligatoria)
 *   SESSION_SECRET   Clave para firmar las sesiones (obligatoria)
 *   ADMIN_DNI        DNI del primer administrador (se crea al iniciar si no existe)
 *   ADMIN_PASSWORD   Contraseña inicial de ese administrador (mín. 8 caracteres)
 *   ADMIN_NOMBRE     Nombre del administrador (opcional)
 *   ALLOWED_ORIGINS  Orígenes permitidos para CORS, separados por coma
 */
'use strict';

const express = require('express');
const compression = require('compression');
const cors = require('cors');
const crypto = require('crypto');
const { Pool } = require('pg');

const PORT = process.env.PORT || 3000;
const SCHEMA = 'reporte_vs_anual';
const SECRET = process.env.SESSION_SECRET;
if (!process.env.DATABASE_URL) { console.error('Falta DATABASE_URL'); process.exit(1); }
if (!SECRET || SECRET.length < 16) { console.error('Falta SESSION_SECRET (mínimo 16 caracteres)'); process.exit(1); }

const isLocalDb = /localhost|127\.0\.0\.1|\.internal|^postgres(ql)?:\/\/[^@]+@dpg-[^.]+(\/|:|$)/.test(process.env.DATABASE_URL);
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: isLocalDb ? false : { rejectUnauthorized: false },
  max: 5,
});

// ---------------------------------------------------------------- esquema
async function migrate() {
  await pool.query(`
    CREATE SCHEMA IF NOT EXISTS ${SCHEMA};
    CREATE TABLE IF NOT EXISTS ${SCHEMA}.usuarios (
      dni         text PRIMARY KEY,
      nombre      text NOT NULL,
      hash        text NOT NULL,
      es_admin    boolean NOT NULL DEFAULT false,
      creado_en   timestamptz NOT NULL DEFAULT now()
    );
    CREATE TABLE IF NOT EXISTS ${SCHEMA}.cargas (
      id          bigserial PRIMARY KEY,
      usuario_dni text,
      usuario     text,
      archivos    text[] NOT NULL,
      anios       int[] NOT NULL,
      filas       int NOT NULL,
      total       numeric(16,2) NOT NULL,
      creado_en   timestamptz NOT NULL DEFAULT now()
    );
    CREATE TABLE IF NOT EXISTS ${SCHEMA}.anios (
      anio        int PRIMARY KEY,
      max_mes     int NOT NULL,
      max_dia     int,
      carga_id    bigint REFERENCES ${SCHEMA}.cargas(id) ON DELETE SET NULL,
      actualizado timestamptz NOT NULL DEFAULT now()
    );
    CREATE TABLE IF NOT EXISTS ${SCHEMA}.ventas (
      anio      int NOT NULL,
      mes       smallint NOT NULL,
      semana    smallint NOT NULL DEFAULT 0,
      cadena    text NOT NULL,
      tienda    text NOT NULL,
      marca     text NOT NULL DEFAULT '',
      venta     numeric(16,2) NOT NULL,
      unidades  numeric(16,2) NOT NULL DEFAULT 0
    );
    CREATE INDEX IF NOT EXISTS ventas_anio_idx ON ${SCHEMA}.ventas (anio);
    -- Cobertura: 1 = coberturado, 0 = no coberturado, NULL = el archivo no traía el dato.
    ALTER TABLE ${SCHEMA}.ventas ADD COLUMN IF NOT EXISTS cobertura smallint;
    -- Detalle diario por cadena/marca/cobertura para la vista "Comparativa" (segmentadores Mes/Semana/Día).
    CREATE TABLE IF NOT EXISTS ${SCHEMA}.ventas_dia (
      anio      int NOT NULL,
      mes       smallint NOT NULL,
      semana    smallint NOT NULL DEFAULT 0,
      dia       smallint NOT NULL DEFAULT 0,
      cadena    text NOT NULL,
      marca     text NOT NULL DEFAULT '',
      cobertura smallint,
      venta     numeric(16,2) NOT NULL,
      unidades  numeric(16,2) NOT NULL DEFAULT 0
    );
    CREATE INDEX IF NOT EXISTS ventas_dia_anio_idx ON ${SCHEMA}.ventas_dia (anio);
    -- Detalle diario por tienda y coordinador, para la comparativa por tienda y por coordinador (se calcula en el servidor).
    CREATE TABLE IF NOT EXISTS ${SCHEMA}.ventas_tienda_dia (
      anio        int NOT NULL,
      mes         smallint NOT NULL,
      semana      smallint NOT NULL DEFAULT 0,
      dia         smallint NOT NULL DEFAULT 0,
      cadena      text NOT NULL,
      tienda      text NOT NULL,
      coordinador text NOT NULL DEFAULT '',
      marca       text NOT NULL DEFAULT '',
      cobertura   smallint,
      venta       numeric(16,2) NOT NULL,
      unidades    numeric(16,2) NOT NULL DEFAULT 0
    );
    CREATE INDEX IF NOT EXISTS ventas_tienda_dia_anio_mes_idx ON ${SCHEMA}.ventas_tienda_dia (anio, mes);
    -- Nombres unificados en mayúsculas (los archivos traían "Tottus" y "TOTTUS", etc.).
    UPDATE ${SCHEMA}.ventas SET cadena = upper(cadena), tienda = upper(tienda), marca = upper(marca)
      WHERE cadena <> upper(cadena) OR tienda <> upper(tienda) OR marca <> upper(marca);
    UPDATE ${SCHEMA}.ventas_dia SET cadena = upper(cadena), marca = upper(marca)
      WHERE cadena <> upper(cadena) OR marca <> upper(marca);
  `);
  const existed = await pool.query(`SELECT to_regclass('${SCHEMA}.reemplazos') IS NOT NULL AS ok`);
  await pool.query(`CREATE TABLE IF NOT EXISTS ${SCHEMA}.reemplazos (
      tipo      text NOT NULL,
      original  text NOT NULL,
      nuevo     text NOT NULL,
      creado_en timestamptz NOT NULL DEFAULT now(),
      PRIMARY KEY (tipo, original)
    )`);
  if (!existed.rows[0].ok) {
    await pool.query(`INSERT INTO ${SCHEMA}.reemplazos (tipo, original, nuevo) VALUES ('coordinador', 'DIEGO AMIAS', 'HERLES CORNELIO') ON CONFLICT DO NOTHING`);
  }
  await pool.query(`
    CREATE TABLE IF NOT EXISTS ${SCHEMA}.perfiles (
      id              serial PRIMARY KEY,
      nombre          text NOT NULL UNIQUE,
      ver_todo        boolean NOT NULL DEFAULT false,
      subir           boolean NOT NULL DEFAULT false,
      eliminar        boolean NOT NULL DEFAULT false,
      gestionar_coord boolean NOT NULL DEFAULT false,
      descargar       boolean NOT NULL DEFAULT false,
      admin_usuarios  boolean NOT NULL DEFAULT false,
      sistema         boolean NOT NULL DEFAULT false,
      creado_en       timestamptz NOT NULL DEFAULT now()
    );
    ALTER TABLE ${SCHEMA}.usuarios ADD COLUMN IF NOT EXISTS perfil_id int REFERENCES ${SCHEMA}.perfiles(id);
    ALTER TABLE ${SCHEMA}.usuarios ADD COLUMN IF NOT EXISTS coordinador text NOT NULL DEFAULT '';
    CREATE INDEX IF NOT EXISTS ventas_tienda_dia_coord_idx ON ${SCHEMA}.ventas_tienda_dia (coordinador);
  `);
  // Perfiles iniciales (solo se crean si no existen; luego se editan desde la app).
  const seed = [
    ['Administrador', true, true, true, true, true, true, true],
    ['Analista', true, true, true, true, true, false, false],
    ['Coordinador', false, false, false, false, true, false, false],
    ['Gerencia', true, false, false, false, true, false, false],
  ];
  for (const [nombre, verTodo, subir, eliminar, coord, descargar, adminU, sistema] of seed) {
    await pool.query(`INSERT INTO ${SCHEMA}.perfiles (nombre, ver_todo, subir, eliminar, gestionar_coord, descargar, admin_usuarios, sistema)
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8) ON CONFLICT (nombre) DO NOTHING`, [nombre, verTodo, subir, eliminar, coord, descargar, adminU, sistema]);
  }
  // Usuarios que existían antes de los perfiles: administradores → Administrador; el resto → Gerencia.
  await pool.query(`UPDATE ${SCHEMA}.usuarios u SET perfil_id = (SELECT id FROM ${SCHEMA}.perfiles WHERE nombre = CASE WHEN u.es_admin THEN 'Administrador' ELSE 'Gerencia' END) WHERE perfil_id IS NULL`);
  const { ADMIN_DNI, ADMIN_PASSWORD, ADMIN_NOMBRE } = process.env;
  if (ADMIN_DNI && ADMIN_PASSWORD) {
    if (ADMIN_PASSWORD.length < 8) console.warn('ADMIN_PASSWORD debe tener al menos 8 caracteres: no se creó el administrador.');
    else {
      const r = await pool.query(`INSERT INTO ${SCHEMA}.usuarios (dni, nombre, hash, es_admin, perfil_id) VALUES ($1, $2, $3, true, (SELECT id FROM ${SCHEMA}.perfiles WHERE nombre = 'Administrador')) ON CONFLICT (dni) DO NOTHING`,
        [ADMIN_DNI.trim(), (ADMIN_NOMBRE || 'Administrador').trim(), hashPassword(ADMIN_PASSWORD)]);
      if (r.rowCount) console.log('Administrador inicial creado:', ADMIN_DNI);
    }
  }
}

// ---------------------------------------------------------------- textos
const clean = (v, n = 200) => String(v ?? '').replace(/\s+/g, ' ').trim().toUpperCase().slice(0, n);
async function coordMap(client = pool) {
  const r = await client.query(`SELECT original, nuevo FROM ${SCHEMA}.reemplazos WHERE tipo = 'coordinador'`);
  return new Map(r.rows.map((x) => [x.original, x.nuevo]));
}

// ---------------------------------------------------------------- contraseñas y sesiones
function hashPassword(pw) {
  const salt = crypto.randomBytes(16);
  const hash = crypto.scryptSync(pw, salt, 64);
  return salt.toString('hex') + ':' + hash.toString('hex');
}
function checkPassword(pw, stored) {
  const [s, h] = String(stored).split(':');
  if (!s || !h) return false;
  const calc = crypto.scryptSync(pw, Buffer.from(s, 'hex'), 64);
  const exp = Buffer.from(h, 'hex');
  return exp.length === calc.length && crypto.timingSafeEqual(exp, calc);
}
const b64 = (s) => Buffer.from(s).toString('base64url');
const PERMS = ['ver_todo', 'subir', 'eliminar', 'gestionar_coord', 'descargar', 'admin_usuarios'];
const USER_SQL = `SELECT u.dni, u.nombre, u.hash, u.coordinador, p.id AS perfil_id, p.nombre AS perfil, ${PERMS.map((k) => `coalesce(p.${k}, false) AS ${k}`).join(', ')}
  FROM ${SCHEMA}.usuarios u LEFT JOIN ${SCHEMA}.perfiles p ON p.id = u.perfil_id WHERE u.dni = $1`;
const publicUser = (u) => ({ dni: u.dni, nombre: u.nombre, coordinador: u.coordinador || '', perfil: u.perfil || 'Sin perfil', perfil_id: u.perfil_id, perms: Object.fromEntries(PERMS.map((k) => [k, !!u[k]])), admin: !!u.admin_usuarios });
function signToken(user) {
  const payload = b64(JSON.stringify({ dni: user.dni, exp: Date.now() + 7 * 864e5 }));
  const sig = crypto.createHmac('sha256', SECRET).update(payload).digest('base64url');
  return payload + '.' + sig;
}
function readToken(tok) {
  const [payload, sig] = String(tok || '').split('.');
  if (!payload || !sig) return null;
  const exp = crypto.createHmac('sha256', SECRET).update(payload).digest('base64url');
  if (sig.length !== exp.length || !crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(exp))) return null;
  try {
    const p = JSON.parse(Buffer.from(payload, 'base64url').toString());
    return p.exp > Date.now() ? p : null;
  } catch { return null; }
}
async function auth(req, res, next) {
  const p = readToken((req.headers.authorization || '').replace(/^Bearer\s+/i, ''));
  if (!p) return res.status(401).json({ error: 'Inicia sesión para continuar.' });
  // El perfil y los permisos se leen en cada pedido: un cambio de perfil rige de inmediato.
  const r = await pool.query(USER_SQL, [p.dni]);
  if (!r.rowCount) return res.status(401).json({ error: 'Tu usuario ya no existe.' });
  req.user = publicUser(r.rows[0]);
  next();
}
const can = (perm, msg) => (req, res, next) => (req.user?.perms[perm] ? next() : res.status(403).json({ error: msg || 'Tu perfil no tiene permiso para hacer esto.' }));
const adminOnly = can('admin_usuarios', 'Solo un administrador puede gestionar usuarios y perfiles.');
// Alcance de datos: perfiles sin "ver todo" solo ven las tiendas del coordinador asignado.
const scopeOf = (u) => (u.perms.ver_todo ? null : (u.coordinador || '##SIN-COORDINADOR-ASIGNADO##'));

// Límite simple de intentos de inicio de sesión por IP.
const attempts = new Map();
function loginLimiter(req, res, next) {
  const ip = req.ip, now = Date.now();
  const a = (attempts.get(ip) || []).filter((t) => now - t < 10 * 60e3);
  if (a.length >= 10) return res.status(429).json({ error: 'Demasiados intentos. Espera unos minutos.' });
  a.push(now); attempts.set(ip, a);
  next();
}

// ---------------------------------------------------------------- app
const app = express();
app.set('trust proxy', 1);
const origins = (process.env.ALLOWED_ORIGINS || 'https://reporte-vs-anual.onrender.com,http://localhost:8765,http://localhost:8000').split(',').map((s) => s.trim()).filter(Boolean);
app.use(cors({ origin: (o, cb) => cb(null, !o || origins.includes(o)), methods: ['GET', 'POST', 'DELETE'], allowedHeaders: ['Content-Type', 'Authorization'] }));
app.use(compression());
app.use(express.json({ limit: '60mb' }));
const wrap = (fn) => (req, res, next) => fn(req, res, next).catch(next);

app.get('/api/health', wrap(async (req, res) => { await pool.query('SELECT 1'); res.json({ ok: true }); }));

// Datos para el reporte (requiere sesión; un coordinador solo recibe las tiendas que le corresponden).
app.get('/api/data', wrap(auth), wrap(async (req, res) => {
  const scope = scopeOf(req.user);
  const [ventas, anios, cargas] = await Promise.all([
    scope == null
      ? pool.query(`SELECT cadena, tienda, marca, anio, mes, semana, venta::float8 AS v, unidades::float8 AS u, cobertura FROM ${SCHEMA}.ventas`)
      : pool.query(`WITH s AS (SELECT DISTINCT anio, cadena, tienda FROM ${SCHEMA}.ventas_tienda_dia WHERE coordinador = $1)
          SELECT v.cadena, v.tienda, v.marca, v.anio, v.mes, v.semana, v.venta::float8 AS v, v.unidades::float8 AS u, v.cobertura
          FROM ${SCHEMA}.ventas v JOIN s USING (anio, cadena, tienda)`, [scope]),
    pool.query(`SELECT a.anio, a.max_mes, a.max_dia, a.actualizado, c.usuario, c.archivos,
                       (SELECT count(*) FROM ${SCHEMA}.ventas v WHERE v.anio = a.anio)::int AS filas,
                       (SELECT coalesce(sum(venta), 0)::float8 FROM ${SCHEMA}.ventas v WHERE v.anio = a.anio) AS total,
                       (SELECT coalesce(sum(venta), 0)::float8 FROM ${SCHEMA}.ventas v WHERE v.anio = a.anio AND v.cobertura = 1) AS cubierto,
                       (SELECT count(*) > 0 FROM ${SCHEMA}.ventas v WHERE v.anio = a.anio AND v.cobertura IS NOT NULL) AS con_cobertura
                FROM ${SCHEMA}.anios a LEFT JOIN ${SCHEMA}.cargas c ON c.id = a.carga_id ORDER BY a.anio`),
    pool.query(`SELECT id, usuario, archivos, anios, filas, total::float8 AS total, creado_en FROM ${SCHEMA}.cargas ORDER BY id DESC LIMIT 20`),
  ]);
  // Formato compacto: diccionarios de textos + filas con índices.
  const dict = { cadenas: [], tiendas: [], marcas: [] };
  const idx = { cadenas: new Map(), tiendas: new Map(), marcas: new Map() };
  const id = (k, v) => { let i = idx[k].get(v); if (i == null) { i = dict[k].length; dict[k].push(v); idx[k].set(v, i); } return i; };
  const rows = ventas.rows.map((r) => [id('cadenas', r.cadena), id('tiendas', r.tienda), id('marcas', r.marca), r.anio, r.mes, r.semana, r.v, r.u, r.cobertura]);
  res.set('Cache-Control', 'no-store');
  // Quien no ve todo no recibe totales ni el historial de cargas de la base completa.
  const aniosOut = scope == null ? anios.rows : anios.rows.map((a) => ({ anio: a.anio, max_mes: a.max_mes, max_dia: a.max_dia, actualizado: a.actualizado }));
  res.json({ anios: aniosOut, cargas: scope == null ? cargas.rows : [], dict, rows, alcance: scope == null ? null : req.user.coordinador });
}));

// Detalle diario para la Comparativa y el Resumen (con el mismo alcance que /api/data).
app.get('/api/dias', wrap(auth), wrap(async (req, res) => {
  const scope = scopeOf(req.user);
  const r = scope == null
    ? await pool.query(`SELECT cadena, marca, cobertura, anio, mes, semana, dia, venta::float8 AS v, unidades::float8 AS u FROM ${SCHEMA}.ventas_dia`)
    : await pool.query(`SELECT cadena, marca, cobertura, anio, mes, semana, dia, sum(venta)::float8 AS v, sum(unidades)::float8 AS u
        FROM ${SCHEMA}.ventas_tienda_dia WHERE coordinador = $1 GROUP BY cadena, marca, cobertura, anio, mes, semana, dia`, [scope]);
  const dict = { cadenas: [], marcas: [] };
  const idx = { cadenas: new Map(), marcas: new Map() };
  const id = (k, v) => { let i = idx[k].get(v); if (i == null) { i = dict[k].length; dict[k].push(v); idx[k].set(v, i); } return i; };
  const rows = r.rows.map((x) => [id('cadenas', x.cadena), id('marcas', x.marca), x.cobertura, x.anio, x.mes, x.semana, x.dia, x.v, x.u]);
  res.set('Cache-Control', 'no-store');
  res.json({ dict, rows });
}));

// Comparativa por tienda o por coordinador (se agrega en el servidor por volumen).
app.post('/api/comparativa', wrap(auth), wrap(async (req, res) => {
  const b = req.body || {};
  const ints = (a) => (Array.isArray(a) ? a.map(Number).filter(Number.isFinite) : []);
  const anios = ints(b.anios).slice(0, 6);
  if (!anios.length) return res.status(400).json({ error: 'Faltan los años.' });
  const where = ['anio = ANY($1::int[])'];
  const params = [anios];
  const add = (sql, v) => { params.push(v); where.push(sql.replace('$?', '$' + params.length)); };
  if (ints(b.meses).length) add('mes = ANY($?::smallint[])', ints(b.meses));
  if (ints(b.semanas).length) add('semana = ANY($?::smallint[])', ints(b.semanas));
  if (ints(b.dias).length) add('dia = ANY($?::smallint[])', ints(b.dias));
  if (b.cob === 'si') where.push('cobertura = 1'); else if (b.cob === 'no') where.push('cobertura = 0'); else if (b.cob === 'sd') where.push('cobertura IS NULL');
  if (Array.isArray(b.marcas)) add('marca = ANY($?::text[])', b.marcas.map((m) => clean(m, 120)));
  if (b.cadena) add('cadena = $?', clean(b.cadena));
  const scope = scopeOf(req.user);
  if (scope != null) add('coordinador = $?', scope);
  const dims = b.por === 'coordinador' ? ['coordinador'] : ['cadena', 'tienda'];
  const coordCol = b.por === 'coordinador' ? '' : `, (array_agg(coordinador ORDER BY mes DESC, dia DESC) FILTER (WHERE coordinador <> ''))[1] AS coordinador`;
  const r = await pool.query(
    `SELECT ${dims.join(', ')}${coordCol}, anio, semana, sum(venta)::float8 AS v, sum(unidades)::float8 AS u
     FROM ${SCHEMA}.ventas_tienda_dia WHERE ${where.join(' AND ')}
     GROUP BY ${dims.join(', ')}, anio, semana`, params);
  const cnt = await pool.query(`SELECT count(*)::int AS n FROM ${SCHEMA}.ventas_tienda_dia WHERE anio = ANY($1::int[])`, [anios]);
  res.json({ rows: r.rows, disponible: cnt.rows[0].n > 0 });
}));

app.get('/api/coordinadores', wrap(auth), wrap(async (req, res) => {
  const scope = scopeOf(req.user);
  const r = await pool.query(`SELECT coordinador, count(DISTINCT tienda)::int AS tiendas, round(sum(venta))::float8 AS venta FROM ${SCHEMA}.ventas_tienda_dia
    WHERE coordinador <> '' AND ($1::text IS NULL OR coordinador = $1) GROUP BY coordinador ORDER BY coordinador`, [scope]);
  const rep = scope == null ? await pool.query(`SELECT original, nuevo FROM ${SCHEMA}.reemplazos WHERE tipo = 'coordinador' ORDER BY original`) : { rows: [] };
  res.json({ coordinadores: r.rows, reemplazos: rep.rows });
}));

// Reemplazar el nombre de un coordinador: se aplica a lo ya guardado y a las próximas cargas.
app.post('/api/reemplazos', wrap(auth), can('gestionar_coord'), wrap(async (req, res) => {
  const original = clean(req.body?.original, 120), nuevo = clean(req.body?.nuevo, 120);
  if (!original || !nuevo) return res.status(400).json({ error: 'Escribe el nombre actual y el nuevo.' });
  if (original === nuevo) return res.status(400).json({ error: 'El nombre nuevo es igual al actual.' });
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    await client.query(`INSERT INTO ${SCHEMA}.reemplazos (tipo, original, nuevo) VALUES ('coordinador', $1, $2) ON CONFLICT (tipo, original) DO UPDATE SET nuevo = EXCLUDED.nuevo`, [original, nuevo]);
    // Si algún reemplazo apuntaba al nombre original, ahora apunta al nuevo (evita cadenas A→B→C).
    await client.query(`UPDATE ${SCHEMA}.reemplazos SET nuevo = $2 WHERE tipo = 'coordinador' AND nuevo = $1`, [original, nuevo]);
    const u = await client.query(`UPDATE ${SCHEMA}.ventas_tienda_dia SET coordinador = $2 WHERE coordinador = $1`, [original, nuevo]);
    await client.query('COMMIT');
    res.json({ ok: true, filas: u.rowCount });
  } catch (e) { await client.query('ROLLBACK'); throw e; } finally { client.release(); }
}));
app.delete('/api/reemplazos/:original', wrap(auth), can('gestionar_coord'), wrap(async (req, res) => {
  await pool.query(`DELETE FROM ${SCHEMA}.reemplazos WHERE tipo = 'coordinador' AND original = $1`, [clean(req.params.original, 120)]);
  res.json({ ok: true });
}));

app.post('/api/login', loginLimiter, wrap(async (req, res) => {
  const dni = String(req.body?.dni || '').trim(), password = String(req.body?.password || '');
  const r = await pool.query(USER_SQL, [dni]);
  if (!r.rowCount || !checkPassword(password, r.rows[0].hash)) return res.status(401).json({ error: 'DNI o contraseña incorrectos.' });
  const u = r.rows[0];
  res.json({ token: signToken(u), user: publicUser(u) });
}));

app.get('/api/me', wrap(auth), (req, res) => res.json({ user: req.user }));

// Guardar archivos: reemplaza por completo los años que vienen en la carga.
app.post('/api/cargas', wrap(auth), can('subir', 'Tu perfil no puede subir datos.'), wrap(async (req, res) => {
  const { archivos, anios, rows, dias, tdias } = req.body || {};
  if (!Array.isArray(rows) || !rows.length) return res.status(400).json({ error: 'No llegaron datos para guardar.' });
  if (!Array.isArray(anios) || !anios.length) return res.status(400).json({ error: 'No llegaron los años de la carga.' });
  const yearSet = new Set(anios.map((a) => +a.anio));
  const cleanRows = [];
  for (const r of rows) {
    if (!Array.isArray(r) || r.length < 8) return res.status(400).json({ error: 'Formato de datos no válido.' });
    const [cad, tie, mar, y, m, w, v, u, cob] = r;
    if (!yearSet.has(+y) || !(m >= 1 && m <= 12) || !(w >= 0 && w <= 54) || !isFinite(v)) return res.status(400).json({ error: `Fila no válida: ${JSON.stringify(r).slice(0, 120)}` });
    cleanRows.push([clean(cad), clean(tie), clean(mar, 120), +y, +m, +w, Math.round(v * 100) / 100, Math.round((+u || 0) * 100) / 100, cob === 1 || cob === 0 ? cob : null]);
  }
  const cleanDias = [];
  for (const r of Array.isArray(dias) ? dias : []) {
    if (!Array.isArray(r) || r.length < 9) return res.status(400).json({ error: 'Formato de datos diarios no válido.' });
    const [cad, mar, cob, y, m, w, d, v, u] = r;
    if (!yearSet.has(+y) || !(m >= 1 && m <= 12) || !(w >= 0 && w <= 54) || !(d >= 0 && d <= 31) || !isFinite(v)) return res.status(400).json({ error: `Fila diaria no válida: ${JSON.stringify(r).slice(0, 120)}` });
    cleanDias.push([clean(cad), clean(mar, 120), cob === 1 || cob === 0 ? cob : null, +y, +m, +w, +d, Math.round(v * 100) / 100, Math.round((+u || 0) * 100) / 100]);
  }
  const cmap = await coordMap();
  const cleanTD = [];
  for (const r of Array.isArray(tdias) ? tdias : []) {
    if (!Array.isArray(r) || r.length < 11) return res.status(400).json({ error: 'Formato de datos por tienda no válido.' });
    const [cad, tie, coord, mar, cob, y, m, w, d, v, u] = r;
    if (!yearSet.has(+y) || !(m >= 1 && m <= 12) || !(w >= 0 && w <= 54) || !(d >= 0 && d <= 31) || !isFinite(v)) return res.status(400).json({ error: `Fila por tienda no válida: ${JSON.stringify(r).slice(0, 120)}` });
    const c = clean(coord, 120);
    cleanTD.push([clean(cad), clean(tie), cmap.get(c) || c, clean(mar, 120), cob === 1 || cob === 0 ? cob : null, +y, +m, +w, +d, Math.round(v * 100) / 100, Math.round((+u || 0) * 100) / 100]);
  }
  const total = cleanRows.reduce((t, r) => t + r[6], 0);
  const years = [...yearSet].sort();
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const c = await client.query(`INSERT INTO ${SCHEMA}.cargas (usuario_dni, usuario, archivos, anios, filas, total) VALUES ($1, $2, $3, $4, $5, $6) RETURNING id`,
      [req.user.dni, req.user.nombre, (archivos || []).map(String).slice(0, 50), years, cleanRows.length, total]);
    const cargaId = c.rows[0].id;
    await client.query(`DELETE FROM ${SCHEMA}.ventas WHERE anio = ANY($1::int[])`, [years]);
    for (let i = 0; i < cleanRows.length; i += 5000) {
      const chunk = cleanRows.slice(i, i + 5000);
      const col = (k) => chunk.map((r) => r[k]);
      await client.query(
        `INSERT INTO ${SCHEMA}.ventas (cadena, tienda, marca, anio, mes, semana, venta, unidades, cobertura)
         SELECT * FROM unnest($1::text[], $2::text[], $3::text[], $4::int[], $5::smallint[], $6::smallint[], $7::numeric[], $8::numeric[], $9::smallint[])`,
        [col(0), col(1), col(2), col(3), col(4), col(5), col(6), col(7), col(8)]);
    }
    await client.query(`DELETE FROM ${SCHEMA}.ventas_dia WHERE anio = ANY($1::int[])`, [years]);
    for (let i = 0; i < cleanDias.length; i += 5000) {
      const chunk = cleanDias.slice(i, i + 5000);
      const col = (k) => chunk.map((r) => r[k]);
      await client.query(
        `INSERT INTO ${SCHEMA}.ventas_dia (cadena, marca, cobertura, anio, mes, semana, dia, venta, unidades)
         SELECT * FROM unnest($1::text[], $2::text[], $3::smallint[], $4::int[], $5::smallint[], $6::smallint[], $7::smallint[], $8::numeric[], $9::numeric[])`,
        [col(0), col(1), col(2), col(3), col(4), col(5), col(6), col(7), col(8)]);
    }
    await client.query(`DELETE FROM ${SCHEMA}.ventas_tienda_dia WHERE anio = ANY($1::int[])`, [years]);
    for (let i = 0; i < cleanTD.length; i += 5000) {
      const chunk = cleanTD.slice(i, i + 5000);
      const col = (k) => chunk.map((r) => r[k]);
      await client.query(
        `INSERT INTO ${SCHEMA}.ventas_tienda_dia (cadena, tienda, coordinador, marca, cobertura, anio, mes, semana, dia, venta, unidades)
         SELECT * FROM unnest($1::text[], $2::text[], $3::text[], $4::text[], $5::smallint[], $6::int[], $7::smallint[], $8::smallint[], $9::smallint[], $10::numeric[], $11::numeric[])`,
        [col(0), col(1), col(2), col(3), col(4), col(5), col(6), col(7), col(8), col(9), col(10)]);
    }
    for (const a of anios) {
      await client.query(
        `INSERT INTO ${SCHEMA}.anios (anio, max_mes, max_dia, carga_id, actualizado) VALUES ($1, $2, $3, $4, now())
         ON CONFLICT (anio) DO UPDATE SET max_mes = EXCLUDED.max_mes, max_dia = EXCLUDED.max_dia, carga_id = EXCLUDED.carga_id, actualizado = now()`,
        [+a.anio, +a.max_mes, a.max_dia == null ? null : +a.max_dia, cargaId]);
    }
    await client.query('COMMIT');
    res.json({ ok: true, carga: cargaId, anios: years, filas: cleanRows.length, dias: cleanDias.length, tiendasDia: cleanTD.length, total });
  } catch (e) {
    await client.query('ROLLBACK');
    throw e;
  } finally {
    client.release();
  }
}));

app.delete('/api/anios/:anio', wrap(auth), can('eliminar', 'Tu perfil no puede eliminar datos.'), wrap(async (req, res) => {
  const y = +req.params.anio;
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const d = await client.query(`DELETE FROM ${SCHEMA}.ventas WHERE anio = $1`, [y]);
    await client.query(`DELETE FROM ${SCHEMA}.ventas_dia WHERE anio = $1`, [y]);
    await client.query(`DELETE FROM ${SCHEMA}.ventas_tienda_dia WHERE anio = $1`, [y]);
    await client.query(`DELETE FROM ${SCHEMA}.anios WHERE anio = $1`, [y]);
    await client.query('COMMIT');
    res.json({ ok: true, filas: d.rowCount });
  } catch (e) { await client.query('ROLLBACK'); throw e; } finally { client.release(); }
}));

// ---------------------------------------------------------------- usuarios
const perfilExists = async (id) => (await pool.query(`SELECT 1 FROM ${SCHEMA}.perfiles WHERE id = $1`, [id])).rowCount > 0;
const adminCount = async (exceptDni) => (await pool.query(`SELECT count(*)::int AS n FROM ${SCHEMA}.usuarios u JOIN ${SCHEMA}.perfiles p ON p.id = u.perfil_id WHERE p.admin_usuarios AND u.dni <> $1`, [exceptDni || ''])).rows[0].n;

app.get('/api/usuarios', wrap(auth), adminOnly, wrap(async (req, res) => {
  const r = await pool.query(`SELECT u.dni, u.nombre, u.coordinador, u.creado_en, p.id AS perfil_id, p.nombre AS perfil, coalesce(p.ver_todo, false) AS ver_todo
    FROM ${SCHEMA}.usuarios u LEFT JOIN ${SCHEMA}.perfiles p ON p.id = u.perfil_id ORDER BY u.nombre`);
  res.json({ usuarios: r.rows });
}));
app.post('/api/usuarios', wrap(auth), adminOnly, wrap(async (req, res) => {
  const dni = String(req.body?.dni || '').trim(), nombre = String(req.body?.nombre || '').trim(), pw = String(req.body?.password || '');
  const perfilId = +req.body?.perfil_id, coord = clean(req.body?.coordinador, 120);
  if (!/^\d{8,12}$/.test(dni)) return res.status(400).json({ error: 'El DNI debe tener entre 8 y 12 dígitos.' });
  if (!nombre) return res.status(400).json({ error: 'Escribe el nombre.' });
  if (pw.length < 8) return res.status(400).json({ error: 'La contraseña debe tener al menos 8 caracteres.' });
  if (!(await perfilExists(perfilId))) return res.status(400).json({ error: 'Elige un perfil.' });
  const r = await pool.query(`INSERT INTO ${SCHEMA}.usuarios (dni, nombre, hash, es_admin, perfil_id, coordinador)
    VALUES ($1, $2, $3, (SELECT admin_usuarios FROM ${SCHEMA}.perfiles WHERE id = $4), $4, $5) ON CONFLICT (dni) DO NOTHING`, [dni, nombre, hashPassword(pw), perfilId, coord]);
  if (!r.rowCount) return res.status(409).json({ error: 'Ya existe un usuario con ese DNI.' });
  res.json({ ok: true });
}));
// Editar un usuario: nombre, perfil, coordinador asignado y, si se escribe, una contraseña nueva.
app.post('/api/usuarios/:dni', wrap(auth), adminOnly, wrap(async (req, res) => {
  const dni = req.params.dni;
  const cur = await pool.query(`SELECT dni FROM ${SCHEMA}.usuarios WHERE dni = $1`, [dni]);
  if (!cur.rowCount) return res.status(404).json({ error: 'No existe ese usuario.' });
  const nombre = String(req.body?.nombre || '').trim(), perfilId = +req.body?.perfil_id, coord = clean(req.body?.coordinador, 120), pw = String(req.body?.password || '');
  if (!nombre) return res.status(400).json({ error: 'Escribe el nombre.' });
  if (!(await perfilExists(perfilId))) return res.status(400).json({ error: 'Elige un perfil.' });
  if (pw && pw.length < 8) return res.status(400).json({ error: 'La contraseña nueva debe tener al menos 8 caracteres.' });
  const newIsAdmin = (await pool.query(`SELECT admin_usuarios FROM ${SCHEMA}.perfiles WHERE id = $1`, [perfilId])).rows[0].admin_usuarios;
  if (!newIsAdmin && (await adminCount(dni)) === 0) return res.status(400).json({ error: 'Debe quedar al menos un usuario que pueda administrar usuarios y perfiles.' });
  await pool.query(`UPDATE ${SCHEMA}.usuarios SET nombre = $2, perfil_id = $3, coordinador = $4, es_admin = $5${pw ? ', hash = $6' : ''} WHERE dni = $1`,
    pw ? [dni, nombre, perfilId, coord, newIsAdmin, hashPassword(pw)] : [dni, nombre, perfilId, coord, newIsAdmin]);
  res.json({ ok: true });
}));
app.delete('/api/usuarios/:dni', wrap(auth), adminOnly, wrap(async (req, res) => {
  if (req.params.dni === req.user.dni) return res.status(400).json({ error: 'No puedes eliminar tu propio usuario.' });
  await pool.query(`DELETE FROM ${SCHEMA}.usuarios WHERE dni = $1`, [req.params.dni]);
  res.json({ ok: true });
}));

app.get('/api/perfiles', wrap(auth), adminOnly, wrap(async (req, res) => {
  const r = await pool.query(`SELECT p.*, (SELECT count(*)::int FROM ${SCHEMA}.usuarios u WHERE u.perfil_id = p.id) AS usuarios FROM ${SCHEMA}.perfiles p ORDER BY p.sistema DESC, p.id`);
  res.json({ perfiles: r.rows, permisos: PERMS });
}));
// Crear o editar un perfil (con id = editar).
app.post('/api/perfiles', wrap(auth), adminOnly, wrap(async (req, res) => {
  const b = req.body || {};
  const nombre = String(b.nombre || '').replace(/\s+/g, ' ').trim().slice(0, 60);
  if (!nombre) return res.status(400).json({ error: 'Escribe el nombre del perfil.' });
  const vals = PERMS.map((k) => !!b[k]);
  if (b.id) {
    const cur = await pool.query(`SELECT * FROM ${SCHEMA}.perfiles WHERE id = $1`, [+b.id]);
    if (!cur.rowCount) return res.status(404).json({ error: 'No existe ese perfil.' });
    if (cur.rows[0].sistema) vals[PERMS.indexOf('admin_usuarios')] = true; // el perfil Administrador siempre administra
    if (cur.rows[0].admin_usuarios && !vals[PERMS.indexOf('admin_usuarios')]) {
      const others = (await pool.query(`SELECT count(*)::int AS n FROM ${SCHEMA}.usuarios u JOIN ${SCHEMA}.perfiles p ON p.id = u.perfil_id WHERE p.admin_usuarios AND p.id <> $1`, [+b.id])).rows[0].n;
      if (!others) return res.status(400).json({ error: 'Debe quedar al menos un perfil con usuarios que pueda administrar.' });
    }
    try {
      await pool.query(`UPDATE ${SCHEMA}.perfiles SET nombre = $2, ${PERMS.map((k, i) => `${k} = $${i + 3}`).join(', ')} WHERE id = $1`, [+b.id, cur.rows[0].sistema ? cur.rows[0].nombre : nombre, ...vals]);
    } catch (e) { if (e.code === '23505') return res.status(409).json({ error: 'Ya existe un perfil con ese nombre.' }); throw e; }
    await pool.query(`UPDATE ${SCHEMA}.usuarios SET es_admin = $2 WHERE perfil_id = $1`, [+b.id, vals[PERMS.indexOf('admin_usuarios')]]);
    return res.json({ ok: true });
  }
  try {
    const r = await pool.query(`INSERT INTO ${SCHEMA}.perfiles (nombre, ${PERMS.join(', ')}) VALUES ($1, ${PERMS.map((_, i) => '$' + (i + 2)).join(', ')}) RETURNING id`, [nombre, ...vals]);
    res.json({ ok: true, id: r.rows[0].id });
  } catch (e) { if (e.code === '23505') return res.status(409).json({ error: 'Ya existe un perfil con ese nombre.' }); throw e; }
}));
app.delete('/api/perfiles/:id', wrap(auth), adminOnly, wrap(async (req, res) => {
  const cur = await pool.query(`SELECT sistema, (SELECT count(*)::int FROM ${SCHEMA}.usuarios u WHERE u.perfil_id = p.id) AS n FROM ${SCHEMA}.perfiles p WHERE id = $1`, [+req.params.id]);
  if (!cur.rowCount) return res.status(404).json({ error: 'No existe ese perfil.' });
  if (cur.rows[0].sistema) return res.status(400).json({ error: 'El perfil Administrador no se puede eliminar.' });
  if (cur.rows[0].n) return res.status(400).json({ error: `Este perfil tiene ${cur.rows[0].n} usuario(s). Cámbiales el perfil antes de eliminarlo.` });
  await pool.query(`DELETE FROM ${SCHEMA}.perfiles WHERE id = $1`, [+req.params.id]);
  res.json({ ok: true });
}));

app.post('/api/password', wrap(auth), wrap(async (req, res) => {
  const actual = String(req.body?.actual || ''), nueva = String(req.body?.nueva || '');
  const r = await pool.query(`SELECT hash FROM ${SCHEMA}.usuarios WHERE dni = $1`, [req.user.dni]);
  if (!checkPassword(actual, r.rows[0].hash)) return res.status(400).json({ error: 'La contraseña actual no es correcta.' });
  if (nueva.length < 8) return res.status(400).json({ error: 'La nueva contraseña debe tener al menos 8 caracteres.' });
  await pool.query(`UPDATE ${SCHEMA}.usuarios SET hash = $2 WHERE dni = $1`, [req.user.dni, hashPassword(nueva)]);
  res.json({ ok: true });
}));

app.use((err, req, res, next) => {
  console.error(err);
  if (err.type === 'entity.too.large') return res.status(413).json({ error: 'La carga es demasiado grande. Sube menos archivos a la vez.' });
  res.status(500).json({ error: 'Error del servidor: ' + (err.message || 'desconocido') });
});

migrate()
  .then(() => app.listen(PORT, () => console.log(`API Reporte Vs Anual escuchando en el puerto ${PORT}`)))
  .catch((e) => { console.error('No se pudo preparar la base de datos:', e); process.exit(1); });
