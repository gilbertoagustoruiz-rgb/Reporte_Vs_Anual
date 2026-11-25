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
  `);
  const { ADMIN_DNI, ADMIN_PASSWORD, ADMIN_NOMBRE } = process.env;
  if (ADMIN_DNI && ADMIN_PASSWORD) {
    if (ADMIN_PASSWORD.length < 8) console.warn('ADMIN_PASSWORD debe tener al menos 8 caracteres: no se creó el administrador.');
    else {
      const r = await pool.query(`INSERT INTO ${SCHEMA}.usuarios (dni, nombre, hash, es_admin) VALUES ($1, $2, $3, true) ON CONFLICT (dni) DO NOTHING`,
        [ADMIN_DNI.trim(), (ADMIN_NOMBRE || 'Administrador').trim(), hashPassword(ADMIN_PASSWORD)]);
      if (r.rowCount) console.log('Administrador inicial creado:', ADMIN_DNI);
    }
  }
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
function signToken(user) {
  const payload = b64(JSON.stringify({ dni: user.dni, nombre: user.nombre, admin: user.es_admin, exp: Date.now() + 7 * 864e5 }));
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
  // Confirma que el usuario sigue existiendo (por si lo eliminaron).
  const r = await pool.query(`SELECT dni, nombre, es_admin FROM ${SCHEMA}.usuarios WHERE dni = $1`, [p.dni]);
  if (!r.rowCount) return res.status(401).json({ error: 'Tu usuario ya no existe.' });
  req.user = r.rows[0];
  next();
}
const adminOnly = (req, res, next) => (req.user?.es_admin ? next() : res.status(403).json({ error: 'Solo un administrador puede hacer esto.' }));

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

// Datos para el reporte (público: cualquiera con el enlace puede ver).
app.get('/api/data', wrap(async (req, res) => {
  const [ventas, anios, cargas] = await Promise.all([
    pool.query(`SELECT cadena, tienda, marca, anio, mes, semana, venta::float8 AS v, unidades::float8 AS u, cobertura FROM ${SCHEMA}.ventas`),
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
  res.json({ anios: anios.rows, cargas: cargas.rows, dict, rows });
}));

// Detalle diario para la vista Comparativa (público, como el resto del reporte).
app.get('/api/dias', wrap(async (req, res) => {
  const r = await pool.query(`SELECT cadena, marca, cobertura, anio, mes, semana, dia, venta::float8 AS v, unidades::float8 AS u FROM ${SCHEMA}.ventas_dia`);
  const dict = { cadenas: [], marcas: [] };
  const idx = { cadenas: new Map(), marcas: new Map() };
  const id = (k, v) => { let i = idx[k].get(v); if (i == null) { i = dict[k].length; dict[k].push(v); idx[k].set(v, i); } return i; };
  const rows = r.rows.map((x) => [id('cadenas', x.cadena), id('marcas', x.marca), x.cobertura, x.anio, x.mes, x.semana, x.dia, x.v, x.u]);
  res.set('Cache-Control', 'no-store');
  res.json({ dict, rows });
}));

app.post('/api/login', loginLimiter, wrap(async (req, res) => {
  const dni = String(req.body?.dni || '').trim(), password = String(req.body?.password || '');
  const r = await pool.query(`SELECT dni, nombre, hash, es_admin FROM ${SCHEMA}.usuarios WHERE dni = $1`, [dni]);
  if (!r.rowCount || !checkPassword(password, r.rows[0].hash)) return res.status(401).json({ error: 'DNI o contraseña incorrectos.' });
  const u = r.rows[0];
  res.json({ token: signToken(u), user: { dni: u.dni, nombre: u.nombre, admin: u.es_admin } });
}));

app.get('/api/me', wrap(auth), (req, res) => res.json({ user: { dni: req.user.dni, nombre: req.user.nombre, admin: req.user.es_admin } }));

// Guardar archivos: reemplaza por completo los años que vienen en la carga.
app.post('/api/cargas', wrap(auth), wrap(async (req, res) => {
  const { archivos, anios, rows, dias } = req.body || {};
  if (!Array.isArray(rows) || !rows.length) return res.status(400).json({ error: 'No llegaron datos para guardar.' });
  if (!Array.isArray(anios) || !anios.length) return res.status(400).json({ error: 'No llegaron los años de la carga.' });
  const yearSet = new Set(anios.map((a) => +a.anio));
  const clean = [];
  for (const r of rows) {
    if (!Array.isArray(r) || r.length < 8) return res.status(400).json({ error: 'Formato de datos no válido.' });
    const [cad, tie, mar, y, m, w, v, u, cob] = r;
    if (!yearSet.has(+y) || !(m >= 1 && m <= 12) || !(w >= 0 && w <= 54) || !isFinite(v)) return res.status(400).json({ error: `Fila no válida: ${JSON.stringify(r).slice(0, 120)}` });
    clean.push([String(cad).slice(0, 200), String(tie).slice(0, 200), String(mar ?? '').slice(0, 120), +y, +m, +w, Math.round(v * 100) / 100, Math.round((+u || 0) * 100) / 100, cob === 1 || cob === 0 ? cob : null]);
  }
  const cleanDias = [];
  for (const r of Array.isArray(dias) ? dias : []) {
    if (!Array.isArray(r) || r.length < 9) return res.status(400).json({ error: 'Formato de datos diarios no válido.' });
    const [cad, mar, cob, y, m, w, d, v, u] = r;
    if (!yearSet.has(+y) || !(m >= 1 && m <= 12) || !(w >= 0 && w <= 54) || !(d >= 0 && d <= 31) || !isFinite(v)) return res.status(400).json({ error: `Fila diaria no válida: ${JSON.stringify(r).slice(0, 120)}` });
    cleanDias.push([String(cad).slice(0, 200), String(mar ?? '').slice(0, 120), cob === 1 || cob === 0 ? cob : null, +y, +m, +w, +d, Math.round(v * 100) / 100, Math.round((+u || 0) * 100) / 100]);
  }
  const total = clean.reduce((t, r) => t + r[6], 0);
  const years = [...yearSet].sort();
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const c = await client.query(`INSERT INTO ${SCHEMA}.cargas (usuario_dni, usuario, archivos, anios, filas, total) VALUES ($1, $2, $3, $4, $5, $6) RETURNING id`,
      [req.user.dni, req.user.nombre, (archivos || []).map(String).slice(0, 50), years, clean.length, total]);
    const cargaId = c.rows[0].id;
    await client.query(`DELETE FROM ${SCHEMA}.ventas WHERE anio = ANY($1::int[])`, [years]);
    for (let i = 0; i < clean.length; i += 5000) {
      const chunk = clean.slice(i, i + 5000);
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
    for (const a of anios) {
      await client.query(
        `INSERT INTO ${SCHEMA}.anios (anio, max_mes, max_dia, carga_id, actualizado) VALUES ($1, $2, $3, $4, now())
         ON CONFLICT (anio) DO UPDATE SET max_mes = EXCLUDED.max_mes, max_dia = EXCLUDED.max_dia, carga_id = EXCLUDED.carga_id, actualizado = now()`,
        [+a.anio, +a.max_mes, a.max_dia == null ? null : +a.max_dia, cargaId]);
    }
    await client.query('COMMIT');
    res.json({ ok: true, carga: cargaId, anios: years, filas: clean.length, dias: cleanDias.length, total });
  } catch (e) {
    await client.query('ROLLBACK');
    throw e;
  } finally {
    client.release();
  }
}));

app.delete('/api/anios/:anio', wrap(auth), wrap(async (req, res) => {
  const y = +req.params.anio;
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const d = await client.query(`DELETE FROM ${SCHEMA}.ventas WHERE anio = $1`, [y]);
    await client.query(`DELETE FROM ${SCHEMA}.ventas_dia WHERE anio = $1`, [y]);
    await client.query(`DELETE FROM ${SCHEMA}.anios WHERE anio = $1`, [y]);
    await client.query('COMMIT');
    res.json({ ok: true, filas: d.rowCount });
  } catch (e) { await client.query('ROLLBACK'); throw e; } finally { client.release(); }
}));

// ---------------------------------------------------------------- usuarios
app.get('/api/usuarios', wrap(auth), adminOnly, wrap(async (req, res) => {
  const r = await pool.query(`SELECT dni, nombre, es_admin, creado_en FROM ${SCHEMA}.usuarios ORDER BY nombre`);
  res.json({ usuarios: r.rows });
}));
app.post('/api/usuarios', wrap(auth), adminOnly, wrap(async (req, res) => {
  const dni = String(req.body?.dni || '').trim(), nombre = String(req.body?.nombre || '').trim(), pw = String(req.body?.password || '');
  if (!/^\d{8,12}$/.test(dni)) return res.status(400).json({ error: 'El DNI debe tener entre 8 y 12 dígitos.' });
  if (!nombre) return res.status(400).json({ error: 'Escribe el nombre.' });
  if (pw.length < 8) return res.status(400).json({ error: 'La contraseña debe tener al menos 8 caracteres.' });
  const r = await pool.query(`INSERT INTO ${SCHEMA}.usuarios (dni, nombre, hash, es_admin) VALUES ($1, $2, $3, $4) ON CONFLICT (dni) DO NOTHING`, [dni, nombre, hashPassword(pw), !!req.body?.admin]);
  if (!r.rowCount) return res.status(409).json({ error: 'Ya existe un usuario con ese DNI.' });
  res.json({ ok: true });
}));
app.delete('/api/usuarios/:dni', wrap(auth), adminOnly, wrap(async (req, res) => {
  if (req.params.dni === req.user.dni) return res.status(400).json({ error: 'No puedes eliminar tu propio usuario.' });
  await pool.query(`DELETE FROM ${SCHEMA}.usuarios WHERE dni = $1`, [req.params.dni]);
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
