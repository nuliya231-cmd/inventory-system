const express = require('express');
const session = require('express-session');
const multer = require('multer');
const bcrypt = require('bcryptjs');
const path = require('path');
const fs = require('fs');
const initSqlJs = require('sql.js');
const { v4: uuidv4 } = require('uuid');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;

// === Ensure directories ===
fs.mkdirSync('./uploads', { recursive: true });
fs.mkdirSync('./data', { recursive: true });

// === Database (sql.js) ===
let db = null;
let SQL = null;

function loadDB() {
  const dbPath = './data/inventory.db';
  if (fs.existsSync(dbPath)) {
    const buffer = fs.readFileSync(dbPath);
    db = new SQL.Database(buffer);
  } else {
    db = new SQL.Database();
  }
}

function saveDB() {
  const data = db.export();
  fs.writeFileSync('./data/inventory.db', Buffer.from(data));
}

// Helper: run SQL, no result
function dbRun(sql, params = []) {
  // Replace named params with positional
  // sql.js uses ? for params
  db.run(sql, params);
  saveDB();
}

// Helper: get single row
function dbGet(sql, params = []) {
  const stmt = db.prepare(sql);
  const result = stmt.get(params);
  stmt.free();
  if (!result) return null;
  // Convert array to object using column names
  const columns = db.exec("PRAGMA table_info(" + getTableName(sql) + ")");
  return result;
}

// Actually, sql.js returns results differently.
// Let me use a proper wrapper.

// === Proper DB wrapper for sql.js ===
// sql.js query helpers
function dbExec(sql) {
  // For CREATE/INSERT/UPDATE/DELETE
  db.run(sql);
  saveDB();
}

function dbSelect(sql, params = []) {
  // Returns array of objects
  const results = db.exec(sql);
  if (!results || results.length === 0) return [];
  const { columns, values } = results[0];
  return values.map(row => {
    const obj = {};
    columns.forEach((col, i) => obj[col] = row[i]);
    return obj;
  });
}

function dbSelectOne(sql, params = []) {
  const results = db.exec(sql);
  if (!results || results.length === 0 || results[0].values.length === 0) return null;
  const { columns, values } = results[0];
  const obj = {};
  columns.forEach((col, i) => obj[col] = values[0][i]);
  return obj;
}

function dbRunWithLastId(sql, params = []) {
  db.run(sql, params);
  const lastId = dbSelectOne("SELECT last_insert_rowid() as id");
  saveDB();
  return lastId ? lastId.id : null;
}

// For INSERT with returning id
function dbInsert(sql, params = []) {
  db.run(sql, params);
  const result = dbSelectOne("SELECT last_insert_rowid() as lastId");
  saveDB();
  return result ? result.lastId : null;
}

// Init DB schema
function initDBSchema() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS users (
      id TEXT PRIMARY KEY,
      phone TEXT UNIQUE NOT NULL,
      password TEXT NOT NULL,
      name TEXT NOT NULL,
      role TEXT NOT NULL DEFAULT 'bd',
      created_at TEXT DEFAULT (datetime('now', 'localtime'))
    );
    CREATE TABLE IF NOT EXISTS product_groups (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      image TEXT,
      created_at TEXT DEFAULT (datetime('now', 'localtime'))
    );
    CREATE TABLE IF NOT EXISTS product_variants (
      id TEXT PRIMARY KEY,
      group_id TEXT NOT NULL,
      name TEXT NOT NULL,
      stock INTEGER NOT NULL DEFAULT 0,
      occupied INTEGER NOT NULL DEFAULT 0,
      price REAL NOT NULL DEFAULT 0,
      created_at TEXT DEFAULT (datetime('now', 'localtime')),
      FOREIGN KEY (group_id) REFERENCES product_groups(id) ON DELETE CASCADE
    );
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
      price REAL NOT NULL,
      subtotal REAL NOT NULL,
      receiver_name TEXT NOT NULL,
      receiver_phone TEXT NOT NULL,
      province TEXT DEFAULT '',
      city TEXT DEFAULT '',
      district TEXT DEFAULT '',
      address TEXT NOT NULL,
      created_at TEXT DEFAULT (datetime('now', 'localtime'))
    );
    CREATE TABLE IF NOT EXISTS payment_images (
      id TEXT PRIMARY KEY,
      record_id TEXT NOT NULL,
      file_path TEXT NOT NULL,
      original_name TEXT DEFAULT '',
      created_at TEXT DEFAULT (datetime('now', 'localtime')),
      FOREIGN KEY (record_id) REFERENCES records(id) ON DELETE CASCADE
    );
  `);
  saveDB();
}

// === Initialize ===
(async () => {
  SQL = await initSqlJs();
  loadDB();
  initDBSchema();

  // Seed admin
  const adminExists = dbSelectOne("SELECT id FROM users WHERE phone = 'admin'");
  if (!adminExists) {
    db.exec(`INSERT INTO users (id, phone, password, name, role) VALUES ('${uuidv4()}', 'admin', '${bcrypt.hashSync('admin123', 10)}', '管理员', 'admin')`);
    saveDB();
  }

  // Seed demo data
  const prodCount = dbSelectOne("SELECT COUNT(*) as c FROM product_groups");
  if (!prodCount || prodCount.c === 0) {
    const g1 = uuidv4(), g2 = uuidv4();
    db.exec(`INSERT INTO product_groups (id, name) VALUES ('${g1}', '春节档电影周边礼盒')`);
    db.exec(`INSERT INTO product_groups (id, name) VALUES ('${g2}', '情人节限定套装')`);
    db.exec(`INSERT INTO product_variants (id, group_id, name, stock, price) VALUES ('${uuidv4()}', '${g1}', 'A款-经典红', 100, 68)`);
    db.exec(`INSERT INTO product_variants (id, group_id, name, stock, price) VALUES ('${uuidv4()}', '${g1}', 'B款-雅致金', 50, 88)`);
    db.exec(`INSERT INTO product_variants (id, group_id, name, stock, price) VALUES ('${uuidv4()}', '${g1}', 'C款-潮流黑', 80, 78)`);
    db.exec(`INSERT INTO product_variants (id, group_id, name, stock, price) VALUES ('${uuidv4()}', '${g2}', '玫瑰礼盒', 60, 128)`);
    db.exec(`INSERT INTO product_variants (id, group_id, name, stock, price) VALUES ('${uuidv4()}', '${g2}', '巧克力礼盒', 40, 98)`);
    saveDB();
  }
})();

// === Middleware ===
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(session({
  secret: 'inv-mgmt-secret-' + Date.now(),
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 7 * 24 * 60 * 60 * 1000 }
}));
app.use(express.static('public'));
app.use('/uploads', express.static('uploads'));

// Multer
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, './uploads'),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname) || '.jpg';
    cb(null, Date.now() + '-' + uuidv4().slice(0, 8) + ext);
  }
});
const upload = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ok = /\.(jpg|jpeg|png|gif|webp|bmp)$/i.test(file.originalname);
    cb(null, ok);
  }
});

// Auth helpers
function auth(req, res, next) {
  if (!req.session.user) return res.status(401).json({ error: '请先登录' });
  next();
}
function adminOnly(req, res, next) {
  if (!req.session.user || req.session.user.role !== 'admin') return res.status(403).json({ error: '无权限访问' });
  next();
}

// ====== AUTH ROUTES ======
app.post('/api/login', (req, res) => {
  const { phone, password } = req.body;
  if (!phone || !password) return res.status(400).json({ error: '请输入手机号和密码' });
  const user = dbSelectOne(`SELECT * FROM users WHERE phone = '${phone.replace(/'/g, "''")}'`);
  if (!user || !bcrypt.compareSync(password, user.password)) return res.status(401).json({ error: '手机号或密码错误' });
  req.session.user = { id: user.id, phone: user.phone, name: user.name, role: user.role };
  res.json({ user: req.session.user });
});

app.post('/api/logout', (req, res) => { req.session.destroy(); res.json({ ok: true }); });

app.get('/api/me', (req, res) => {
  if (!req.session.user) return res.status(401).json({ error: '未登录' });
  res.json({ user: req.session.user });
});

app.put('/api/me/password', auth, (req, res) => {
  const { oldPassword, newPassword } = req.body;
  if (!oldPassword || !newPassword) return res.status(400).json({ error: '请输入旧密码和新密码' });
  const user = dbSelectOne(`SELECT * FROM users WHERE id = '${req.session.user.id}'`);
  if (!bcrypt.compareSync(oldPassword, user.password)) return res.status(400).json({ error: '旧密码错误' });
  if (newPassword.length < 4) return res.status(400).json({ error: '新密码至少4位' });
  dbExec(`UPDATE users SET password = '${bcrypt.hashSync(newPassword, 10)}' WHERE id = '${req.session.user.id}'`);
  res.json({ ok: true });
});

// ====== USER MANAGEMENT (Admin) ======
app.get('/api/users', adminOnly, (req, res) => {
  const users = dbSelect(`SELECT id, phone, name, role, created_at FROM users ORDER BY created_at DESC`);
  res.json({ users });
});

app.post('/api/users', adminOnly, (req, res) => {
  const { phone, name, role, password } = req.body;
  if (!phone || !name || !role || !password) return res.status(400).json({ error: '请填写所有字段' });
  const exists = dbSelectOne(`SELECT id FROM users WHERE phone = '${phone.replace(/'/g, "''")}'`);
  if (exists) return res.status(400).json({ error: '该手机号已存在' });
  const id = uuidv4();
  dbExec(`INSERT INTO users (id, phone, password, name, role) VALUES ('${id}', '${phone.replace(/'/g, "''")}', '${bcrypt.hashSync(password, 10)}', '${name.replace(/'/g, "''")}', '${role}')`);
  res.json({ ok: true, user: { id, phone, name, role } });
});

app.post('/api/users/batch', adminOnly, (req, res) => {
  const { users } = req.body;
  if (!Array.isArray(users) || users.length === 0) return res.status(400).json({ error: '无数据' });
  let added = 0, skipped = 0;
  for (const u of users) {
    if (!u.phone || !u.name || !u.password) { skipped++; continue; }
    const exists = dbSelectOne(`SELECT id FROM users WHERE phone = '${u.phone.replace(/'/g, "''")}'`);
    if (exists) { skipped++; continue; }
    dbExec(`INSERT INTO users (id, phone, password, name, role) VALUES ('${uuidv4()}', '${u.phone.replace(/'/g, "''")}', '${bcrypt.hashSync(u.password, 10)}', '${u.name.replace(/'/g, "''")}', '${u.role || 'bd'}')`);
    added++;
  }
  saveDB();
  res.json({ ok: true, added, skipped });
});

app.put('/api/users/:id', adminOnly, (req, res) => {
  const { name, role, password } = req.body;
  const user = dbSelectOne(`SELECT * FROM users WHERE id = '${req.params.id}'`);
  if (!user) return res.status(404).json({ error: '用户不存在' });
  let sql = `UPDATE users SET name = '${name.replace(/'/g, "''")}', role = '${role}'`;
  if (password) sql += `, password = '${bcrypt.hashSync(password, 10)}'`;
  sql += ` WHERE id = '${req.params.id}'`;
  dbExec(sql);
  res.json({ ok: true });
});

app.delete('/api/users/:id', adminOnly, (req, res) => {
  const user = dbSelectOne(`SELECT * FROM users WHERE id = '${req.params.id}'`);
  if (!user) return res.status(404).json({ error: '用户不存在' });
  const adminCount = dbSelectOne("SELECT COUNT(*) as c FROM users WHERE role='admin'");
  if (user.role === 'admin' && adminCount.c <= 1) return res.status(400).json({ error: '不能删除最后一个管理员' });
  dbExec(`DELETE FROM users WHERE id = '${req.params.id}'`);
  res.json({ ok: true });
});

// ====== PRODUCT MANAGEMENT ======
app.get('/api/products', (req, res) => {
  const groups = dbSelect('SELECT * FROM product_groups ORDER BY created_at DESC');
  const variants = dbSelect('SELECT * FROM product_variants ORDER BY created_at ASC');
  const products = groups.map(g => ({ ...g, variants: variants.filter(v => v.group_id === g.id) }));
  res.json({ products });
});

app.post('/api/products', adminOnly, upload.single('image'), (req, res) => {
  const { name } = req.body;
  if (!name) return res.status(400).json({ error: '请输入商品组名称' });
  const id = uuidv4();
  const image = req.file ? `/uploads/${req.file.filename}` : null;
  dbExec(`INSERT INTO product_groups (id, name, image) VALUES ('${id}', '${name.replace(/'/g, "''")}', ${image ? "'" + image + "'" : 'NULL'})`);
  res.json({ ok: true, product: { id, name, image, variants: [] } });
});

app.put('/api/products/:id', adminOnly, upload.single('image'), (req, res) => {
  const { name, removeImage } = req.body;
  const p = dbSelectOne(`SELECT * FROM product_groups WHERE id = '${req.params.id}'`);
  if (!p) return res.status(404).json({ error: '商品不存在' });
  let image = p.image;
  if (req.file) {
    image = `/uploads/${req.file.filename}`;
  } else if (removeImage === '1') {
    if (p.image) try { fs.unlinkSync('.' + p.image); } catch (e) { }
    image = null;
  }
  dbExec(`UPDATE product_groups SET name = '${name.replace(/'/g, "''")}', image = ${image ? "'" + image + "'" : 'NULL'} WHERE id = '${req.params.id}'`);
  res.json({ ok: true });
});

app.delete('/api/products/:id', adminOnly, (req, res) => {
  const p = dbSelectOne(`SELECT * FROM product_groups WHERE id = '${req.params.id}'`);
  if (!p) return res.status(404).json({ error: '商品不存在' });
  const variants = dbSelect(`SELECT id FROM product_variants WHERE group_id = '${req.params.id}'`);
  for (const v of variants) {
    const recs = dbSelect(`SELECT id FROM records WHERE variant_id = '${v.id}'`);
    for (const r of recs) {
      const imgs = dbSelect(`SELECT file_path FROM payment_images WHERE record_id = '${r.id}'`);
      for (const img of imgs) { try { fs.unlinkSync('.' + img.file_path); } catch (e) { } }
      dbExec(`DELETE FROM payment_images WHERE record_id = '${r.id}'`);
    }
    dbExec(`DELETE FROM records WHERE variant_id = '${v.id}'`);
  }
  dbExec(`DELETE FROM product_variants WHERE group_id = '${req.params.id}'`);
  if (p.image) { try { fs.unlinkSync('.' + p.image); } catch (e) { } }
  dbExec(`DELETE FROM product_groups WHERE id = '${req.params.id}'`);
  res.json({ ok: true });
});

// ====== VARIANT MANAGEMENT ======
app.post('/api/products/:groupId/variants', adminOnly, (req, res) => {
  const { name, stock, price } = req.body;
  if (!name) return res.status(400).json({ error: '请输入规格名称' });
  const g = dbSelectOne(`SELECT id FROM product_groups WHERE id = '${req.params.groupId}'`);
  if (!g) return res.status(404).json({ error: '商品组不存在' });
  const id = uuidv4();
  dbExec(`INSERT INTO product_variants (id, group_id, name, stock, price) VALUES ('${id}', '${req.params.groupId}', '${name.replace(/'/g, "''")}', ${stock || 0}, ${price || 0})`);
  res.json({ ok: true, variant: { id, group_id: req.params.groupId, name, stock: stock || 0, occupied: 0, price: price || 0 } });
});

app.put('/api/products/:groupId/variants/:variantId', adminOnly, (req, res) => {
  const { name, stock, price } = req.body;
  const v = dbSelectOne(`SELECT * FROM product_variants WHERE id = '${req.params.variantId}' AND group_id = '${req.params.groupId}'`);
  if (!v) return res.status(404).json({ error: '规格不存在' });
  dbExec(`UPDATE product_variants SET name = '${name.replace(/'/g, "''")}', stock = ${stock !== undefined ? stock : v.stock}, price = ${price !== undefined ? price : v.price} WHERE id = '${req.params.variantId}'`);
  res.json({ ok: true });
});

app.delete('/api/products/:groupId/variants/:variantId', adminOnly, (req, res) => {
  const v = dbSelectOne(`SELECT * FROM product_variants WHERE id = '${req.params.variantId}' AND group_id = '${req.params.groupId}'`);
  if (!v) return res.status(404).json({ error: '规格不存在' });
  const recs = dbSelect(`SELECT id FROM records WHERE variant_id = '${req.params.variantId}'`);
  for (const r of recs) {
    const imgs = dbSelect(`SELECT file_path FROM payment_images WHERE record_id = '${r.id}'`);
    for (const img of imgs) { try { fs.unlinkSync('.' + img.file_path); } catch (e) { } }
    dbExec(`DELETE FROM payment_images WHERE record_id = '${r.id}'`);
  }
  dbExec(`DELETE FROM records WHERE variant_id = '${req.params.variantId}'`);
  dbExec(`DELETE FROM product_variants WHERE id = '${req.params.variantId}'`);
  res.json({ ok: true });
});

// ====== RECORDS ======
function parseRecordImages(recs) {
  return recs.map(r => {
    const imgs = dbSelect(`SELECT file_path, original_name FROM payment_images WHERE record_id = '${r.id}'`);
    return { ...r, images: imgs };
  });
}

app.get('/api/records', auth, (req, res) => {
  let recs;
  if (req.session.user.role === 'admin') {
    recs = dbSelect(`SELECT * FROM records ORDER BY created_at DESC`);
  } else {
    recs = dbSelect(`SELECT * FROM records WHERE user_id = '${req.session.user.id}' ORDER BY created_at DESC`);
  }
  recs = parseRecordImages(recs);
  res.json({ records: recs });
});

app.post('/api/records', auth, upload.array('paymentImages', 10), (req, res) => {
  const { cinemaSpecialId, cinemaName, variantId, qty, price, receiverName, receiverPhone, province, city, district, address } = req.body;
  if (!variantId) return res.status(400).json({ error: '请选择商品规格' });
  if (!qty || qty < 1) return res.status(400).json({ error: '数量至少为1' });
  if (!receiverName) return res.status(400).json({ error: '请填写收件人姓名' });
  if (!receiverPhone) return res.status(400).json({ error: '请填写收件人手机号' });
  if (!address) return res.status(400).json({ error: '请填写收件人地址' });
  if (!req.files || req.files.length === 0) return res.status(400).json({ error: '请上传付款截图' });

  const variant = dbSelectOne(`SELECT v.*, pg.name as group_name, pg.id as group_id FROM product_variants v JOIN product_groups pg ON v.group_id = pg.id WHERE v.id = '${variantId}'`);
  if (!variant) return res.status(400).json({ error: '商品规格不存在' });
  const remain = variant.stock - variant.occupied;
  if (qty > remain) return res.status(400).json({ error: `库存不足，最多可占用 ${remain} 件` });

  const actualPrice = parseFloat(price) || variant.price;
  const subtotal = qty * actualPrice;
  const id = uuidv4();

  dbExec(`INSERT INTO records (id, user_id, bd_name, bd_id, cinema_special_id, cinema_name, product_group_id, product_group_name, variant_id, variant_name, qty, price, subtotal, receiver_name, receiver_phone, province, city, district, address) VALUES (
    '${id}', '${req.session.user.id}', '${req.session.user.name.replace(/'/g, "''")}', '${req.body.bdId ? req.body.bdId.replace(/'/g, "''") : ''}',
    '${cinemaSpecialId ? cinemaSpecialId.replace(/'/g, "''") : ''}', '${cinemaName ? cinemaName.replace(/'/g, "''") : ''}',
    '${variant.group_id}', '${variant.group_name.replace(/'/g, "''")}', '${variant.id}', '${variant.name.replace(/'/g, "''")}',
    ${qty}, ${actualPrice}, ${subtotal},
    '${receiverName.replace(/'/g, "''")}', '${receiverPhone.replace(/'/g, "''")}',
    '${province ? province.replace(/'/g, "''") : ''}', '${city ? city.replace(/'/g, "''") : ''}', '${district ? district.replace(/'/g, "''") : ''}',
    '${address.replace(/'/g, "''")}'
  )`);

  for (const file of req.files) {
    dbExec(`INSERT INTO payment_images (id, record_id, file_path, original_name) VALUES ('${uuidv4()}', '${id}', '/uploads/${file.filename}', '${file.originalname.replace(/'/g, "''")}')`);
  }

  dbExec(`UPDATE product_variants SET occupied = occupied + ${qty} WHERE id = '${variantId}'`);
  res.json({ ok: true, recordId: id });
});

app.delete('/api/records/:id', adminOnly, (req, res) => {
  const rec = dbSelectOne(`SELECT * FROM records WHERE id = '${req.params.id}'`);
  if (!rec) return res.status(404).json({ error: '记录不存在' });
  dbExec(`UPDATE product_variants SET occupied = MAX(0, occupied - ${rec.qty}) WHERE id = '${rec.variant_id}'`);
  const imgs = dbSelect(`SELECT file_path FROM payment_images WHERE record_id = '${req.params.id}'`);
  for (const img of imgs) { try { fs.unlinkSync('.' + img.file_path); } catch (e) { } }
  dbExec(`DELETE FROM payment_images WHERE record_id = '${req.params.id}'`);
  dbExec(`DELETE FROM records WHERE id = '${req.params.id}'`);
  res.json({ ok: true });
});

// ====== DASHBOARD ======
app.get('/api/dashboard', adminOnly, (req, res) => {
  const today = new Date().toISOString().slice(0, 10);
  const todayRecs = dbSelect(`SELECT * FROM records WHERE created_at LIKE '${today}%'`);
  const totalRecs = dbSelect(`SELECT * FROM records`);
  const variants = dbSelect(`SELECT v.*, pg.name as group_name FROM product_variants v JOIN product_groups pg ON v.group_id = pg.id`);

  const trend = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date(); d.setDate(d.getDate() - i);
    const ds = d.toISOString().slice(0, 10);
    const dayRecs = dbSelect(`SELECT qty, subtotal FROM records WHERE created_at LIKE '${ds}%'`);
    trend.push({ date: ds, qty: dayRecs.reduce((s, r) => s + r.qty, 0), amount: dayRecs.reduce((s, r) => s + r.subtotal, 0), count: dayRecs.length });
  }

  const bdRanking = dbSelect(`SELECT bd_name, SUM(qty) as total_qty, SUM(subtotal) as total_amount, COUNT(*) as count FROM records WHERE created_at LIKE '${today}%' GROUP BY bd_name ORDER BY total_qty DESC LIMIT 10`);

  res.json({
    today: { occupied: todayRecs.reduce((s, r) => s + r.qty, 0), amount: todayRecs.reduce((s, r) => s + r.subtotal, 0), count: todayRecs.length },
    total: { occupied: totalRecs.reduce((s, r) => s + r.qty, 0), amount: totalRecs.reduce((s, r) => s + r.subtotal, 0), count: totalRecs.length },
    inventory: variants.map(v => ({ name: v.group_name + ' - ' + v.name, stock: v.stock, occupied: v.occupied, remain: v.stock - v.occupied, rate: v.stock > 0 ? Math.round(v.occupied / v.stock * 100) : 0 })),
    trend,
    bdRanking
  });
});

// ====== EXPORT WITH IMAGES ======
app.get('/api/export', adminOnly, async (req, res) => {
  try {
    const { dateFrom, dateTo, bdName, productId } = req.query;
    let recs = dbSelect(`SELECT * FROM records ORDER BY created_at ASC`);
    if (dateFrom) recs = recs.filter(r => r.created_at && r.created_at.slice(0, 10) >= dateFrom);
    if (dateTo) recs = recs.filter(r => r.created_at && r.created_at.slice(0, 10) <= dateTo);
    if (bdName) recs = recs.filter(r => r.bd_name && r.bd_name.includes(bdName));
    if (productId) recs = recs.filter(r => r.product_group_id === productId);
    if (recs.length === 0) return res.status(400).json({ error: '无数据可导出' });

    // Attach images
    recs = parseRecordImages(recs);

    let maxImgs = 0;
    recs.forEach(r => { if ((r.images || []).length > maxImgs) maxImgs = r.images.length; });

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
      const row = ws.addRow([
        r.created_at ? new Date(r.created_at + 'Z').toLocaleString('zh-CN') : '',
        r.bd_name, r.bd_id || '',
        r.cinema_special_id || '', r.cinema_name || '',
        r.receiver_name, r.receiver_phone,
        r.province || '', r.city || '', r.district || '', r.address,
        r.product_group_name, r.variant_name,
        r.qty, r.price, r.subtotal,
      ]);
      row.height = 75;

      if (r.images && r.images.length > 0) {
        for (let j = 0; j < r.images.length; j++) {
          try {
            const fullPath = '.' + r.images[j].file_path;
            if (fs.existsSync(fullPath)) {
              const ext = path.extname(fullPath).slice(1).toLowerCase();
              const imageId = workbook.addImage({
                buffer: fs.readFileSync(fullPath),
                extension: ext === 'jpg' ? 'jpeg' : ext
              });
              ws.addImage(imageId, {
                tl: { col: 16 + j, row: i + 1 },
                ext: { width: 120, height: 65 }
              });
            }
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

app.listen(PORT, '0.0.0.0', () => {
  console.log(`\n=================================`);
  console.log(`  库存管理系统已启动！`);
  console.log(`  访问地址: http://localhost:${PORT}`);
  console.log(`  默认管理员: 手机号 admin / 密码 admin123`);
  console.log(`=================================\n`);
});
