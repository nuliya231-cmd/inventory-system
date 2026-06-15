const express = require('express');
const session = require('express-session');
const multer = require('multer');
const bcrypt = require('bcryptjs');
const path = require('path');
const { Pool } = require('pg');
const { v4: uuidv4 } = require('uuid');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;

// Multer - memory storage (no local files)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ok = /\.(jpg|jpeg|png|gif|webp|bmp)$/i.test(file.originalname);
    cb(null, ok);
  }
});

// === Database (PostgreSQL) ===
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false,
});

async function query(sql, params = []) {
  const { rows } = await pool.query(sql, params);
  return rows;
}

async function queryOne(sql, params = []) {
  const rows = await query(sql, params);
  return rows.length > 0 ? rows[0] : null;
}

async function run(sql, params = []) {
  await pool.query(sql, params);
}

// === Init DB schema ===
async function initDB() {
  await run(`
    CREATE TABLE IF NOT EXISTS users (
      id TEXT PRIMARY KEY,
      phone TEXT UNIQUE NOT NULL,
      password TEXT NOT NULL,
      name TEXT NOT NULL,
      role TEXT NOT NULL DEFAULT 'bd',
      created_at TIMESTAMPTZ DEFAULT NOW()
    )
  `);
  await run(`
    CREATE TABLE IF NOT EXISTS product_groups (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      image_data TEXT,
      image_ext TEXT DEFAULT 'jpeg',
      created_at TIMESTAMPTZ DEFAULT NOW()
    )
  `);
  await run(`
    CREATE TABLE IF NOT EXISTS product_variants (
      id TEXT PRIMARY KEY,
      group_id TEXT NOT NULL REFERENCES product_groups(id) ON DELETE CASCADE,
      name TEXT NOT NULL,
      stock INTEGER NOT NULL DEFAULT 0,
      occupied INTEGER NOT NULL DEFAULT 0,
      price NUMERIC(10,2) NOT NULL DEFAULT 0,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )
  `);
  await run(`
    CREATE TABLE IF NOT EXISTS records (
      id TEXT PRIMARY KEY,
      user_id TEXT NOT NULL,
      bd_name TEXT NOT NULL,
      bd_id TEXT DEFAULT '',
      cinema_special_id TEXT DEFAULT '',
      cinema_name TEXT DEFAULT '',
      product_group_id TEXT NOT NULL,
      product_group_name TEXT NOT NULL,
      variant_id TEXT NOT NULL,
      variant_name TEXT NOT NULL,
      qty INTEGER NOT NULL,
      price NUMERIC(10,2) NOT NULL,
      subtotal NUMERIC(10,2) NOT NULL,
      receiver_name TEXT NOT NULL,
      receiver_phone TEXT NOT NULL,
      province TEXT DEFAULT '',
      city TEXT DEFAULT '',
      district TEXT DEFAULT '',
      address TEXT NOT NULL,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )
  `);
  await run(`
    CREATE TABLE IF NOT EXISTS payment_images (
      id TEXT PRIMARY KEY,
      record_id TEXT NOT NULL REFERENCES records(id) ON DELETE CASCADE,
      data TEXT NOT NULL,
      ext TEXT DEFAULT 'jpeg',
      original_name TEXT DEFAULT '',
      created_at TIMESTAMPTZ DEFAULT NOW()
    )
  `);

  // Seed admin
  const adminExists = await queryOne("SELECT id FROM users WHERE phone = 'admin'");
  if (!adminExists) {
    await run(`INSERT INTO users (id, phone, password, name, role) VALUES ($1, 'admin', $2, '管理员', 'admin')`, [uuidv4(), bcrypt.hashSync('admin123', 10)]);
  }

  // Seed demo products
  const prodCount = await queryOne("SELECT COUNT(*) as c FROM product_groups");
  if (!prodCount || parseInt(prodCount.c) === 0) {
    const g1 = uuidv4(), g2 = uuidv4();
    await run(`INSERT INTO product_groups (id, name) VALUES ($1, '春节档电影周边礼盒')`, [g1]);
    await run(`INSERT INTO product_groups (id, name) VALUES ($1, '情人节限定套装')`, [g2]);
    await run(`INSERT INTO product_variants (id, group_id, name, stock, price) VALUES ($1, $2, 'A款-经典红', 100, 68)`, [uuidv4(), g1]);
    await run(`INSERT INTO product_variants (id, group_id, name, stock, price) VALUES ($1, $2, 'B款-雅致金', 50, 88)`, [uuidv4(), g1]);
    await run(`INSERT INTO product_variants (id, group_id, name, stock, price) VALUES ($1, $2, 'C款-潮流黑', 80, 78)`, [uuidv4(), g1]);
    await run(`INSERT INTO product_variants (id, group_id, name, stock, price) VALUES ($1, $2, '玫瑰礼盒', 60, 128)`, [uuidv4(), g2]);
    await run(`INSERT INTO product_variants (id, group_id, name, stock, price) VALUES ($1, $2, '巧克力礼盒', 40, 98)`, [uuidv4(), g2]);
  }
}

// === Middleware ===
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// Session store
let sessionStore;
try {
  const pgSession = require('connect-pg-simple');
  sessionStore = new (pgSession(session))({ pool, createTableIfMissing: true });
  console.log('Session store (PG) initialized.');
} catch(e) {
  console.warn('Session store init failed, using memory store:', e.message);
  sessionStore = new session.MemoryStore();
}

app.use(session({
  store: sessionStore,
  secret: process.env.SESSION_SECRET || 'inv-mgmt-secret-2026',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 7 * 24 * 60 * 60 * 1000 }
}));

app.use(express.static('public'));

// Auth helpers
function auth(req, res, next) {
  if (!req.session.user) return res.status(401).json({ error: '请先登录' });
  next();
}
function adminOnly(req, res, next) {
  if (!req.session.user || req.session.user.role !== 'admin') return res.status(403).json({ error: '无权限访问' });
  next();
}

// Helper: convert buffer to base64
function bufToBase64(buf) {
  return buf.toString('base64');
}

// Helper: get extension from filename
function getExt(filename) {
  const ext = path.extname(filename || '').slice(1).toLowerCase();
  if (ext === 'jpg') return 'jpeg';
  if (['jpeg', 'png', 'gif', 'webp', 'bmp'].includes(ext)) return ext;
  return 'jpeg';
}

// ====== IMAGE SERVING ======
app.get('/api/images/payment/:id', async (req, res) => {
  try {
    const img = await queryOne('SELECT data, ext FROM payment_images WHERE id = $1', [req.params.id]);
    if (!img || !img.data) return res.status(404).send('Not found');
    const buf = Buffer.from(img.data, 'base64');
    const contentType = `image/${img.ext || 'jpeg'}`;
    res.set('Content-Type', contentType);
    res.set('Cache-Control', 'public, max-age=86400');
    res.send(buf);
  } catch (e) { res.status(500).send('Error'); }
});

app.get('/api/images/product/:id', async (req, res) => {
  try {
    const p = await queryOne('SELECT image_data, image_ext FROM product_groups WHERE id = $1', [req.params.id]);
    if (!p || !p.image_data) return res.status(404).send('Not found');
    const buf = Buffer.from(p.image_data, 'base64');
    const contentType = `image/${p.image_ext || 'jpeg'}`;
    res.set('Content-Type', contentType);
    res.set('Cache-Control', 'public, max-age=86400');
    res.send(buf);
  } catch (e) { res.status(500).send('Error'); }
});

// ====== AUTH ROUTES ======
app.post('/api/login', async (req, res) => {
  try {
    const { phone, password } = req.body;
    console.log('[Login attempt] phone:', phone);
    if (!phone || !password) return res.status(400).json({ error: '请输入手机号和密码' });
    const user = await queryOne('SELECT * FROM users WHERE phone = $1', [phone]);
    console.log('[Login] user found:', !!user);
    if (!user) return res.status(401).json({ error: '手机号或密码错误' });
    const match = bcrypt.compareSync(password, user.password);
    console.log('[Login] password match:', match);
    if (!match) return res.status(401).json({ error: '手机号或密码错误' });
    req.session.user = { id: user.id, phone: user.phone, name: user.name, role: user.role };
    res.json({ user: req.session.user });
  } catch (e) {
    console.error('[Login error]', e.message, e.stack);
    res.status(500).json({ error: '登录失败: ' + e.message });
  }
});

app.post('/api/logout', (req, res) => { req.session.destroy(); res.json({ ok: true }); });

app.get('/api/me', (req, res) => {
  if (!req.session.user) return res.status(401).json({ error: '未登录' });
  res.json({ user: req.session.user });
});

// Debug endpoint - remove in production
app.get('/api/debug', async (req, res) => {
  try {
    const userCount = await queryOne('SELECT COUNT(*) as c FROM users');
    const users = await query('SELECT id, phone, name, role FROM users LIMIT 5');
    const tables = await query(`SELECT table_name FROM information_schema.tables WHERE table_schema='public'`);
    res.json({ userCount, users, tables: tables.map(t => t.table_name) });
  } catch(e) {
    res.json({ error: e.message });
  }
});

app.put('/api/me/password', auth, async (req, res) => {
  try {
    const { oldPassword, newPassword } = req.body;
    if (!oldPassword || !newPassword) return res.status(400).json({ error: '请输入旧密码和新密码' });
    const user = await queryOne('SELECT * FROM users WHERE id = $1', [req.session.user.id]);
    if (!bcrypt.compareSync(oldPassword, user.password)) return res.status(400).json({ error: '旧密码错误' });
    if (newPassword.length < 4) return res.status(400).json({ error: '新密码至少4位' });
    await run('UPDATE users SET password = $1 WHERE id = $2', [bcrypt.hashSync(newPassword, 10), req.session.user.id]);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: '修改失败' }); }
});

// ====== USER MANAGEMENT (Admin) ======
app.get('/api/users', adminOnly, async (req, res) => {
  try {
    const users = await query('SELECT id, phone, name, role, created_at FROM users ORDER BY created_at DESC');
    res.json({ users });
  } catch (e) { res.status(500).json({ error: '查询失败' }); }
});

app.post('/api/users', adminOnly, async (req, res) => {
  try {
    const { phone, name, role, password } = req.body;
    if (!phone || !name || !role || !password) return res.status(400).json({ error: '请填写所有字段' });
    const exists = await queryOne('SELECT id FROM users WHERE phone = $1', [phone]);
    if (exists) return res.status(400).json({ error: '该手机号已存在' });
    const id = uuidv4();
    await run('INSERT INTO users (id, phone, password, name, role) VALUES ($1, $2, $3, $4, $5)', [id, phone, bcrypt.hashSync(password, 10), name, role]);
    res.json({ ok: true, user: { id, phone, name, role } });
  } catch (e) { res.status(500).json({ error: '创建失败' }); }
});

app.post('/api/users/batch', adminOnly, async (req, res) => {
  try {
    const { users } = req.body;
    if (!Array.isArray(users) || users.length === 0) return res.status(400).json({ error: '无数据' });
    let added = 0, skipped = 0;
    for (const u of users) {
      if (!u.phone || !u.name || !u.password) { skipped++; continue; }
      const exists = await queryOne('SELECT id FROM users WHERE phone = $1', [u.phone]);
      if (exists) { skipped++; continue; }
      await run('INSERT INTO users (id, phone, password, name, role) VALUES ($1, $2, $3, $4, $5)', [uuidv4(), u.phone, bcrypt.hashSync(u.password, 10), u.name, u.role || 'bd']);
      added++;
    }
    res.json({ ok: true, added, skipped });
  } catch (e) { res.status(500).json({ error: '批量创建失败' }); }
});

app.put('/api/users/:id', adminOnly, async (req, res) => {
  try {
    const { name, role, password } = req.body;
    const user = await queryOne('SELECT * FROM users WHERE id = $1', [req.params.id]);
    if (!user) return res.status(404).json({ error: '用户不存在' });
    if (password) {
      await run('UPDATE users SET name = $1, role = $2, password = $3 WHERE id = $4', [name, role, bcrypt.hashSync(password, 10), req.params.id]);
    } else {
      await run('UPDATE users SET name = $1, role = $2 WHERE id = $3', [name, role, req.params.id]);
    }
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: '修改失败' }); }
});

app.delete('/api/users/:id', adminOnly, async (req, res) => {
  try {
    const user = await queryOne('SELECT * FROM users WHERE id = $1', [req.params.id]);
    if (!user) return res.status(404).json({ error: '用户不存在' });
    const adminCount = await queryOne("SELECT COUNT(*) as c FROM users WHERE role='admin'");
    if (user.role === 'admin' && parseInt(adminCount.c) <= 1) return res.status(400).json({ error: '不能删除最后一个管理员' });
    await run('DELETE FROM users WHERE id = $1', [req.params.id]);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: '删除失败' }); }
});

// ====== PRODUCT MANAGEMENT ======
app.get('/api/products', async (req, res) => {
  try {
    const groups = await query('SELECT id, name, image_ext, created_at FROM product_groups ORDER BY created_at DESC');
    const variants = await query('SELECT * FROM product_variants ORDER BY created_at ASC');
    const products = groups.map(g => ({
      ...g,
      image: g.image_ext ? `/api/images/product/${g.id}` : null,
      variants: variants.filter(v => v.group_id === g.id)
    }));
    res.json({ products });
  } catch (e) { res.status(500).json({ error: '查询失败' }); }
});

app.post('/api/products', adminOnly, upload.single('image'), async (req, res) => {
  try {
    const { name } = req.body;
    if (!name) return res.status(400).json({ error: '请输入商品组名称' });
    const id = uuidv4();
    let imageData = null, imageExt = null;
    if (req.file) {
      imageData = bufToBase64(req.file.buffer);
      imageExt = getExt(req.file.originalname);
    }
    await run('INSERT INTO product_groups (id, name, image_data, image_ext) VALUES ($1, $2, $3, $4)', [id, name, imageData, imageExt]);
    res.json({ ok: true, product: { id, name, image: imageExt ? `/api/images/product/${id}` : null, variants: [] } });
  } catch (e) { res.status(500).json({ error: '创建失败' }); }
});

app.put('/api/products/:id', adminOnly, upload.single('image'), async (req, res) => {
  try {
    const { name, removeImage } = req.body;
    const p = await queryOne('SELECT * FROM product_groups WHERE id = $1', [req.params.id]);
    if (!p) return res.status(404).json({ error: '商品不存在' });
    let imageData = p.image_data, imageExt = p.image_ext;
    if (req.file) {
      imageData = bufToBase64(req.file.buffer);
      imageExt = getExt(req.file.originalname);
    } else if (removeImage === '1') {
      imageData = null;
      imageExt = null;
    }
    await run('UPDATE product_groups SET name = $1, image_data = $2, image_ext = $3 WHERE id = $4', [name, imageData, imageExt, req.params.id]);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: '修改失败' }); }
});

app.delete('/api/products/:id', adminOnly, async (req, res) => {
  try {
    const p = await queryOne('SELECT * FROM product_groups WHERE id = $1', [req.params.id]);
    if (!p) return res.status(404).json({ error: '商品不存在' });
    // Cascade deletes will handle payment_images, records, variants
    await run('DELETE FROM product_groups WHERE id = $1', [req.params.id]);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: '删除失败' }); }
});

// ====== VARIANT MANAGEMENT ======
app.post('/api/products/:groupId/variants', adminOnly, async (req, res) => {
  try {
    const { name, stock, price } = req.body;
    if (!name) return res.status(400).json({ error: '请输入规格名称' });
    const g = await queryOne('SELECT id FROM product_groups WHERE id = $1', [req.params.groupId]);
    if (!g) return res.status(404).json({ error: '商品组不存在' });
    const id = uuidv4();
    await run('INSERT INTO product_variants (id, group_id, name, stock, price) VALUES ($1, $2, $3, $4, $5)', [id, req.params.groupId, name, stock || 0, price || 0]);
    res.json({ ok: true, variant: { id, group_id: req.params.groupId, name, stock: stock || 0, occupied: 0, price: price || 0 } });
  } catch (e) { res.status(500).json({ error: '创建失败' }); }
});

app.put('/api/products/:groupId/variants/:variantId', adminOnly, async (req, res) => {
  try {
    const { name, stock, price } = req.body;
    const v = await queryOne('SELECT * FROM product_variants WHERE id = $1 AND group_id = $2', [req.params.variantId, req.params.groupId]);
    if (!v) return res.status(404).json({ error: '规格不存在' });
    await run('UPDATE product_variants SET name = $1, stock = $2, price = $3 WHERE id = $4', [name, stock !== undefined ? stock : v.stock, price !== undefined ? price : v.price, req.params.variantId]);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: '修改失败' }); }
});

app.delete('/api/products/:groupId/variants/:variantId', adminOnly, async (req, res) => {
  try {
    const v = await queryOne('SELECT * FROM product_variants WHERE id = $1 AND group_id = $2', [req.params.variantId, req.params.groupId]);
    if (!v) return res.status(404).json({ error: '规格不存在' });
    // Cascade deletes will handle payment_images and records
    await run('DELETE FROM product_variants WHERE id = $1', [req.params.variantId]);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: '删除失败' }); }
});

// ====== RECORDS ======
async function parseRecordImages(recs) {
  const result = [];
  for (const r of recs) {
    const imgs = await query('SELECT id, ext, original_name FROM payment_images WHERE record_id = $1', [r.id]);
    result.push({ ...r, images: imgs.map(i => ({ id: i.id, file_path: `/api/images/payment/${i.id}`, ext: i.ext, original_name: i.original_name })) });
  }
  return result;
}

app.get('/api/records', auth, async (req, res) => {
  try {
    let recs;
    if (req.session.user.role === 'admin') {
      recs = await query('SELECT * FROM records ORDER BY created_at DESC');
    } else {
      recs = await query('SELECT * FROM records WHERE user_id = $1 ORDER BY created_at DESC', [req.session.user.id]);
    }
    recs = await parseRecordImages(recs);
    res.json({ records: recs });
  } catch (e) { res.status(500).json({ error: '查询失败' }); }
});

app.post('/api/records', auth, upload.array('paymentImages', 10), async (req, res) => {
  try {
    const { cinemaSpecialId, cinemaName, variantId, qty, price, receiverName, receiverPhone, province, city, district, address, bdId } = req.body;
    if (!variantId) return res.status(400).json({ error: '请选择商品规格' });
    if (!qty || qty < 1) return res.status(400).json({ error: '数量至少为1' });
    if (!receiverName) return res.status(400).json({ error: '请填写收件人姓名' });
    if (!receiverPhone) return res.status(400).json({ error: '请填写收件人手机号' });
    if (!address) return res.status(400).json({ error: '请填写收件人地址' });
    if (!req.files || req.files.length === 0) return res.status(400).json({ error: '请上传付款截图' });

    const variant = await queryOne('SELECT v.*, pg.name as group_name, pg.id as group_id FROM product_variants v JOIN product_groups pg ON v.group_id = pg.id WHERE v.id = $1', [variantId]);
    if (!variant) return res.status(400).json({ error: '商品规格不存在' });
    const remain = variant.stock - variant.occupied;
    if (qty > remain) return res.status(400).json({ error: `库存不足，最多可占用 ${remain} 件` });

    const actualPrice = parseFloat(price) || parseFloat(variant.price);
    const subtotal = qty * actualPrice;
    const id = uuidv4();

    await run(`INSERT INTO records (id, user_id, bd_name, bd_id, cinema_special_id, cinema_name, product_group_id, product_group_name, variant_id, variant_name, qty, price, subtotal, receiver_name, receiver_phone, province, city, district, address) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18)`,
      [id, req.session.user.id, req.session.user.name, bdId || '', cinemaSpecialId || '', cinemaName || '', variant.group_id, variant.group_name, variant.id, variant.name, qty, actualPrice, subtotal, receiverName, receiverPhone, province || '', city || '', district || '', address]);

    for (const file of req.files) {
      const imgId = uuidv4();
      const b64 = bufToBase64(file.buffer);
      const ext = getExt(file.originalname);
      await run('INSERT INTO payment_images (id, record_id, data, ext, original_name) VALUES ($1, $2, $3, $4, $5)', [imgId, id, b64, ext, file.originalname]);
    }

    await run('UPDATE product_variants SET occupied = occupied + $1 WHERE id = $2', [qty, variantId]);
    res.json({ ok: true, recordId: id });
  } catch (e) { console.error('Record error:', e); res.status(500).json({ error: '提交失败' }); }
});

app.delete('/api/records/:id', adminOnly, async (req, res) => {
  try {
    const rec = await queryOne('SELECT * FROM records WHERE id = $1', [req.params.id]);
    if (!rec) return res.status(404).json({ error: '记录不存在' });
    await run('UPDATE product_variants SET occupied = GREATEST(0, occupied - $1) WHERE id = $2', [rec.qty, rec.variant_id]);
    // Cascade delete will handle payment_images
    await run('DELETE FROM records WHERE id = $1', [req.params.id]);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: '删除失败' }); }
});

// ====== DASHBOARD ======
app.get('/api/dashboard', adminOnly, async (req, res) => {
  try {
    const today = new Date().toISOString().slice(0, 10);
    const todayRecs = await query("SELECT * FROM records WHERE created_at >= $1", [today + ' 00:00:00']);
    const totalRecs = await query('SELECT * FROM records');
    const variants = await query('SELECT v.*, pg.name as group_name FROM product_variants v JOIN product_groups pg ON v.group_id = pg.id');

    const trend = [];
    for (let i = 6; i >= 0; i--) {
      const d = new Date(); d.setDate(d.getDate() - i);
      const ds = d.toISOString().slice(0, 10);
      const dayRecs = await query("SELECT qty, subtotal FROM records WHERE created_at >= $1 AND created_at < $2", [ds + ' 00:00:00', ds + ' 23:59:59']);
      trend.push({ date: ds, qty: dayRecs.reduce((s, r) => s + parseInt(r.qty), 0), amount: dayRecs.reduce((s, r) => s + parseFloat(r.subtotal), 0), count: dayRecs.length });
    }

    const bdRanking = await query("SELECT bd_name, SUM(qty) as total_qty, SUM(subtotal) as total_amount, COUNT(*) as count FROM records WHERE created_at >= $1 GROUP BY bd_name ORDER BY total_qty DESC LIMIT 10", [today + ' 00:00:00']);

    res.json({
      today: { occupied: todayRecs.reduce((s, r) => s + parseInt(r.qty), 0), amount: todayRecs.reduce((s, r) => s + parseFloat(r.subtotal), 0), count: todayRecs.length },
      total: { occupied: totalRecs.reduce((s, r) => s + parseInt(r.qty), 0), amount: totalRecs.reduce((s, r) => s + parseFloat(r.subtotal), 0), count: totalRecs.length },
      inventory: variants.map(v => ({ name: v.group_name + ' - ' + v.name, stock: parseInt(v.stock), occupied: parseInt(v.occupied), remain: parseInt(v.stock) - parseInt(v.occupied), rate: parseInt(v.stock) > 0 ? Math.round(parseInt(v.occupied) / parseInt(v.stock) * 100) : 0 })),
      trend,
      bdRanking: bdRanking.map(r => ({ ...r, total_qty: parseInt(r.total_qty), total_amount: parseFloat(r.total_amount), count: parseInt(r.count) }))
    });
  } catch (e) { console.error('Dashboard error:', e); res.status(500).json({ error: '查询失败' }); }
});

// ====== EXPORT WITH IMAGES ======
app.get('/api/export', adminOnly, async (req, res) => {
  try {
    const { dateFrom, dateTo, bdName, productId } = req.query;
    let recs = await query('SELECT * FROM records ORDER BY created_at ASC');
    if (dateFrom) recs = recs.filter(r => r.created_at && r.created_at.toISOString().slice(0, 10) >= dateFrom);
    if (dateTo) recs = recs.filter(r => r.created_at && r.created_at.toISOString().slice(0, 10) <= dateTo);
    if (bdName) recs = recs.filter(r => r.bd_name && r.bd_name.includes(bdName));
    if (productId) recs = recs.filter(r => r.product_group_id === productId);
    if (recs.length === 0) return res.status(400).json({ error: '无数据可导出' });

    // Get all images for these records
    const recordIds = recs.map(r => r.id);
    const allImages = await query(`SELECT * FROM payment_images WHERE record_id = ANY($1)`, [recordIds]);

    let maxImgs = 0;
    const imagesByRecord = {};
    for (const img of allImages) {
      if (!imagesByRecord[img.record_id]) imagesByRecord[img.record_id] = [];
      imagesByRecord[img.record_id].push(img);
      if (imagesByRecord[img.record_id].length > maxImgs) maxImgs = imagesByRecord[img.record_id].length;
    }

    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('库存占用数据');

    const headers = ['提交时间', 'BD姓名', 'BD工号', '影城专资ID', '影城名称', '收件人姓名', '收件人手机号/电话', '省', '市', '区', '收件人地址', '商品名称', '货品规格/种类', '货品数量', '货品单价(¥)', '金额小计(¥)'];
    for (let i = 0; i < Math.max(maxImgs, 1); i++) headers.push('付款截图' + (maxImgs > 1 ? (i + 1) : ''));

    ws.columns = headers.map((h, idx) => ({
      header: h,
      width: idx >= 16 ? 22 : h.includes('地址') ? 30 : h.includes('手机') ? 16 : 15
    }));

    ws.getRow(1).height = 25;
    ws.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    for (let i = 0; i < recs.length; i++) {
      const r = recs[i];
      const createdAt = r.created_at ? new Date(r.created_at).toLocaleString('zh-CN') : '';
      const row = ws.addRow([
        createdAt, r.bd_name, r.bd_id || '',
        r.cinema_special_id || '', r.cinema_name || '',
        r.receiver_name, r.receiver_phone,
        r.province || '', r.city || '', r.district || '', r.address,
        r.product_group_name, r.variant_name,
        r.qty, r.price, r.subtotal,
      ]);
      row.height = 75;

      const imgs = imagesByRecord[r.id] || [];
      if (imgs.length > 0) {
        for (let j = 0; j < imgs.length; j++) {
          try {
            const imgData = imgs[j];
            const buf = Buffer.from(imgData.data, 'base64');
            const ext = (imgData.ext || 'jpeg').replace('jpg', 'jpeg');
            const imageId = workbook.addImage({
              buffer: buf,
              extension: ext
            });
            ws.addImage(imageId, {
              tl: { col: 16 + j, row: i + 1 },
              ext: { width: 120, height: 65 }
            });
          } catch (e) { console.error('Image embed error:', e.message); }
        }
      }
    }

    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: recs.length + 1, column: headers.length } };

    const fileName = `inventory-export-${new Date().toISOString().slice(0, 10)}.xlsx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(fileName)}"`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error('Export error:', e);
    res.status(500).json({ error: '导出失败: ' + e.message });
  }
});

// SPA fallback
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// === Start ===
(async () => {
  try {
    await initDB();
    console.log('Database initialized.');
  } catch (e) {
    console.error('DB init error:', e.message);
  }
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n=================================`);
    console.log(`  库存管理系统已启动！`);
    console.log(`  访问地址: http://localhost:${PORT}`);
    console.log(`  默认管理员: 手机号 admin / 密码 admin123`);
    console.log(`=================================\n`);
  });
})();
