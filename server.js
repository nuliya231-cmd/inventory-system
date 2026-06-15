const express = require('express');
const session = require('express-session');
const multer  = require('multer');
const bcrypt = require('bcryptjs');
const path    = require('path');
const { Pool } = require('pg');
const { v4: uuidv4 } = require('uuid');
const ExcelJS = require('exceljs');

const app   = express();
const PORT  = parseInt(process.env.PORT) || 3000;

// ---- Multer: 内存存储（图片存数据库） ----
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    cb(null, /\.(jpg|jpeg|png|gif|webp|bmp)$/i.test(file.originalname));
  },
});

// ============================================================
//  数据库连接
// ============================================================
const DATABASE_URL = process.env.DATABASE_URL || '';
console.log('[DB] DATABASE_URL 长度:', DATABASE_URL.length);
console.log('[DB] DATABASE_URL 前缀:', DATABASE_URL ? DATABASE_URL.substring(0, 30) : '(空)');

let pool;
try {
  pool = new Pool({
    connectionString: DATABASE_URL,
    ssl: DATABASE_URL.includes('supabase') ? { rejectUnauthorized: false } : false,
    max: 10,
    idleTimeoutMillis: 30000,
    connectionTimeoutMillis: 10000,
  });
} catch (e) {
  console.error('[DB] 创建Pool失败:', e.message);
}

async function query(sql, params = []) {
  const { rows } = await pool.query(sql, params);
  return rows;
}
async function queryOne(sql, params = []) {
  const rows = await query(sql, params);
  return rows[0] || null;
}
async function run(sql, params = []) {
  await pool.query(sql, params);
}

// ============================================================
//  初始化数据库（建表 + 种子数据）
// ============================================================
async function initDB() {
  // 测试连接
  try {
    await pool.query('SELECT NOW()');
    console.log('[DB] 连接成功！');
  } catch (e) {
    console.error('[DB] 连接失败！', e.message);
    throw e;
  }

  await run(`
    CREATE TABLE IF NOT EXISTS users (
      id          TEXT PRIMARY KEY,
      phone       TEXT UNIQUE NOT NULL,
      password    TEXT NOT NULL,
      name        TEXT NOT NULL,
      role        TEXT NOT NULL DEFAULT 'bd',
      created_at  TIMESTAMPTZ DEFAULT NOW()
    )`);
  console.log('[DB] 表 users ✓');

  await run(`
    CREATE TABLE IF NOT EXISTS product_groups (
      id          TEXT PRIMARY KEY,
      name        TEXT NOT NULL,
      image_data  TEXT,
      image_ext   TEXT DEFAULT 'jpeg',
      created_at  TIMESTAMPTZ DEFAULT NOW()
    )`);
  console.log('[DB] 表 product_groups ✓');

  await run(`
    CREATE TABLE IF NOT EXISTS product_variants (
      id         TEXT PRIMARY KEY,
      group_id   TEXT NOT NULL REFERENCES product_groups(id) ON DELETE CASCADE,
      name       TEXT NOT NULL,
      stock      INTEGER NOT NULL DEFAULT 0,
      occupied   INTEGER NOT NULL DEFAULT 0,
      price      NUMERIC(10,2) NOT NULL DEFAULT 0,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )`);
  console.log('[DB] 表 product_variants ✓');

  await run(`
    CREATE TABLE IF NOT EXISTS records (
      id                TEXT PRIMARY KEY,
      user_id           TEXT NOT NULL,
      bd_name           TEXT NOT NULL,
      bd_id             TEXT DEFAULT '',
      cinema_special_id TEXT DEFAULT '',
      cinema_name       TEXT DEFAULT '',
      product_group_id  TEXT NOT NULL,
      product_group_name TEXT NOT NULL,
      variant_id        TEXT NOT NULL,
      variant_name      TEXT NOT NULL,
      qty               INTEGER NOT NULL,
      price             NUMERIC(10,2) NOT NULL,
      subtotal          NUMERIC(10,2) NOT NULL,
      receiver_name     TEXT NOT NULL,
      receiver_phone    TEXT NOT NULL,
      province          TEXT DEFAULT '',
      city              TEXT DEFAULT '',
      district          TEXT DEFAULT '',
      address           TEXT NOT NULL,
      created_at        TIMESTAMPTZ DEFAULT NOW()
    )`);
  console.log('[DB] 表 records ✓');

  await run(`
    CREATE TABLE IF NOT EXISTS payment_images (
      id            TEXT PRIMARY KEY,
      record_id     TEXT NOT NULL REFERENCES records(id) ON DELETE CASCADE,
      data          TEXT NOT NULL,
      ext           TEXT DEFAULT 'jpeg',
      original_name TEXT DEFAULT '',
      created_at    TIMESTAMPTZ DEFAULT NOW()
    )`);
  console.log('[DB] 表 payment_images ✓');

  // --- 种子：管理员账号 ---
  const admin = await queryOne("SELECT id FROM users WHERE phone = 'admin'");
  if (!admin) {
    await run(
      "INSERT INTO users (id, phone, password, name, role) VALUES ($1,'admin',$2,'管理员','admin')",
      [uuidv4(), bcrypt.hashSync('admin123', 10)]
    );
    console.log('[DB] 种子管理员 admin/admin123 ✓');
  } else {
    console.log('[DB] 管理员已存在 ✓');
  }

  // --- 种子：示例商品 ---
  const cnt = await queryOne("SELECT COUNT(*) as c FROM product_groups");
  if (!cnt || Number(cnt.c) === 0) {
    const g1 = uuidv4(), g2 = uuidv4();
    await run(`INSERT INTO product_groups (id,name) VALUES($1,'春节档电影周边礼盒')`, [g1]);
    await run(`INSERT INTO product_groups (id,name) VALUES($1,'情人节限定套装')`, [g2]);
    await run(`INSERT INTO product_variants (id,group_id,name,stock,price) VALUES($1,$2,'A款-经典红',100,68)`, [uuidv4(), g1]);
    await run(`INSERT INTO product_variants (id,group_id,name,stock,price) VALUES($1,$2,'B款-雅致金',50,88)`,  [uuidv4(), g1]);
    await run(`INSERT INTO product_variants (id,group_id,name,stock,price) VALUES($1,$2,'C款-潮流黑',80,78)`,  [uuidv4(), g1]);
    await run(`INSERT INTO product_variants (id,group_id,name,stock,price) VALUES($1,$2,'玫瑰礼盒',60,128)`, [uuidv4(), g2]);
    await run(`INSERT INTO product_variants (id,group_id,name,stock,price) VALUES($1,$2,'巧克力礼盒',40,98)`, [uuidv4(), g2]);
    console.log('[DB] 示例商品已创建 ✓');
  }
}

// ============================================================
//  中间件
// ============================================================
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// Session（内存存储，不依赖额外模块）
app.use(session({
  secret: process.env.SESSION_SECRET || 'inv-mgmt-secret-2026-prod',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 7 * 24 * 60 * 60 * 1000, sameSite: 'lax' },
}));

app.use(express.static('public'));

// ---- 权限守卫 ----
function auth(req, res, next) {
  if (!req.session.user) return res.status(401).json({ error: '请先登录' });
  next();
}
function adminOnly(req, res, next) {
  if (!req.session.user || req.session.user.role !== 'admin')
    return res.status(403).json({ error: '无权限访问' });
  next();
}

// ---- 工具函数 ----
function bufToB64(buf) { return buf.toString('base64'); }
function getExt(fn) {
  const e = path.extname(fn || '').slice(1).toLowerCase();
  if (['jpeg','png','gif','webp','bmp'].includes(e)) return e === 'jpg' ? 'jpeg' : e;
  return 'jpeg';
}

// ============================================================
//  图片接口
// ============================================================
app.get('/api/images/payment/:id', async (req, res) => {
  try {
    const img = await queryOne('SELECT data,ext FROM payment_images WHERE id=$1', [req.params.id]);
    if (!img?.data) return res.status(404).send('Not found');
    res.set('Content-Type', 'image/' + (img.ext||'jpeg'));
    res.set('Cache-Control', 'public,max-age=86400');
    res.send(Buffer.from(img.data, 'base64'));
  } catch { res.status(500).send('Error'); }
});

app.get('/api/images/product/:id', async (req, res) => {
  try {
    const p = await queryOne('SELECT image_data,image_ext FROM product_groups WHERE id=$1', [req.params.id]);
    if (!p?.image_data) return res.status(404).send('Not found');
    res.set('Content-Type', 'image/' + (p.image_ext||'jpeg'));
    res.set('Cache-Control', 'public,max-age=86400');
    res.send(Buffer.from(p.image_data, 'base64'));
  } catch { res.status(500).send('Error'); }
});

// ============================================================
//  认证
// ============================================================
app.post('/api/login', async (req, res) => {
  try {
    const { phone, password } = req.body;
    if (!phone || !password) return res.status(400).json({ error: '请输入手机号和密码' });

    console.log('[Login] 尝试登录 phone=' + phone);
    const user = await queryOne('SELECT * FROM users WHERE phone=$1', [phone]);
    console.log('[Login] 查询结果 userExists=', !!user);

    if (!user) return res.status(401).json({ error: '手机号或密码错误' });
    if (!bcrypt.compareSync(password, user.password))
      return res.status(401).json({ error: '手机号或密码错误' });

    req.session.user = {
      id: user.id, phone: user.phone, name: user.name, role: user.role
    };
    console.log('[Login] 成功 role=' + user.role);
    res.json({ user: req.session.user });
  } catch (e) {
    console.error('[Login] 异常:', e.message, e.stack);
    res.status(500).json({ error: '登录失败:' + e.message });
  }
});

app.post('/api/logout', (req, res) => { req.session.destroy(() => res.json({ ok: true })); });

app.get('/api/me', (req, res) => {
  if (!req.session.user) return res.status(401).json({ error: '未登录' });
  res.json({ user: req.session.user });
});

// 健康检查 / 调试
app.get('/api/health', async (req, res) => {
  try {
    const dbTest = await pool.query('SELECT NOW() as time');
    const uc = await queryOne('SELECT COUNT(*)::int as c FROM users');
    res.json({
      status: 'ok',
      dbTime: dbTest.rows[0].time,
      userCount: uc.c,
      env: {
        hasDbUrl: !!DATABASE_URL,
        hasSessionSecret: !!process.env.SESSION_SECRET,
        port: PORT,
      },
    });
  } catch (e) {
    res.status(500).json({ status: 'error', message: e.message });
  }
});

app.put('/api/me/password', auth, async (req, res) => {
  try {
    const { oldPassword, newPassword } = req.body;
    if (!oldPassword || !newPassword) return res.status(400).json({ error: '请填写完整' });
    const u = await queryOne('SELECT * FROM users WHERE id=$1', [req.session.user.id]);
    if (!bcrypt.compareSync(oldPassword, u.password)) return res.status(400).json({ error: '旧密码错误' });
    if (newPassword.length < 4) return res.status(400).json({ error: '新密码至少4位' });
    await run('UPDATE users SET password=$1 WHERE id=$2', [bcrypt.hashSync(newPassword, 10), req.session.user.id]);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: '修改失败' }); }
});

// ============================================================
//  用户管理（管理员）
// ============================================================
app.get('/api/users', adminOnly, async (_req, res) => {
  try {
    const rows = await query('SELECT id,phone,name,role,created_at FROM users ORDER BY created_at DESC');
    res.json({ users: rows });
  } catch (e) { res.status(500).json({ error: '查询失败' }); }
});

app.post('/api/users', adminOnly, async (req, res) => {
  try {
    const { phone, name, role, password } = req.body;
    if (!phone || !name || !role || !password) return res.status(400).json({ error: '请填完整' });
    if (await queryOne('SELECT id FROM users WHERE phone=$1', [phone]))
      return res.status(400).json({ error: '该手机号已存在' });
    const id = uuidv4();
    await run('INSERT INTO users(id,phone,password,name,role) VALUES($1,$2,$3,$4,$5)',
      [id, phone, bcrypt.hashSync(password,10), name, role]);
    res.json({ ok:true, user:{ id, phone, name, role } });
  } catch (e) { res.status(500).json({ error: '创建失败' }); }
});

app.post('/api/users/batch', adminOnly, async (req, res) => {
  try {
    const { users:list } = req.body;
    if (!Array.isArray(list)||list.length===0) return res.status(400).json({ error: '无数据' });
    let added=0, skip=0;
    for (const u of list) {
      if(!u.phone||!u.name||!u.password){skip++;continue;}
      if(await queryOne('SELECT id FROM users WHERE phone=$1',[u.phone])){skip++;continue;}
      await run('INSERT INTO users(id,phone,password,name,role) VALUES($1,$2,$3,$4,$5)',
        [uuidv4(), u.phone, bcrypt.hashSync(u.password,10), u.name, u.role||'bd']);
      added++;
    }
    res.json({ ok:true, added, skipped:skip });
  } catch (e) { res.status(500).json({ error: '批量创建失败:'+e.message }); }
});

app.put('/api/users/:id', adminOnly, async (req, res) => {
  try {
    const { name, role, password } = req.body;
    if(password)
      await run('UPDATE users SET name=$1,role=$2,password=$3 WHERE id=$4',[name,role,bcrypt.hashSync(password,10),req.params.id]);
    else
      await run('UPDATE users SET name=$1,role=$2 WHERE id=$3',[name,role,req.params.id]);
    res.json({ ok:true });
  } catch (e) { res.status(500).json({ error: '修改失败' }); }
});

app.delete('/api/users/:id', adminOnly, async (req, res) => {
  try {
    const u = await queryOne('SELECT * FROM users WHERE id=$1',[req.params.id]);
    if(!u) return res.status(404).json({ error:'不存在' });
    if(u.role==='admin'&&Number((await queryOne("SELECT COUNT(*) as c FROM users WHERE role='admin'")).c)<=1)
      return res.status(400).json({ error:'不能删除最后管理员' });
    await run('DELETE FROM users WHERE id=$1',[req.params.id]);
    res.json({ ok:true });
  } catch (e) { res.status(500).json({ error: '删除失败' }); }
});

// ============================================================
//  商品管理
// ============================================================
app.get('/api/products', async (_req, res) => {
  try {
    const groups   = await query('SELECT id,name,image_ext,created_at FROM product_groups ORDER BY created_at DESC');
    const variants = await query('SELECT * FROM product_variants ORDER BY created_at ASC');
    res.json({
      products: groups.map(g => ({
        ...g,
        image: g.image_ext?`/api/images/product/${g.id}`:null,
        variants: variants.filter(v=>v.group_id===g.id),
      }))
    });
  } catch (e) { console.error('[Products]',e.message); res.status(500).json({ error: '查询失败:'+e.message }); }
});

app.post('/api/products', adminOnly, upload.single('image'), async (req, res) => {
  try {
    const { name } = req.body;
    if(!name) return res.status(400).json({ error:'名称不能为空' });
    const id=uuidv4();
    let imgData=null,imgExt=null;
    if(req.file){imgData=bufToB64(req.file.buffer);imgExt=getExt(req.file.originalname);}
    await run('INSERT INTO product_groups(id,name,image_data,image_ext) VALUES($1,$2,$3,$4)',[id,name,imgData,imgExt]);
    res.json({ ok:true,product:{ id,name,image:imgExt?`/api/images/product/${id}`:null,variants:[] } });
  } catch (e) { res.status(500).json({ error:'创建失败' }); }
});

app.put('/api/products/:id', adminOnly, upload.single('image'), async (req, res) => {
  try {
    const { name, removeImage } = req.body;
    const p = await queryOne('SELECT * FROM product_groups WHERE id=$1',[req.params.id]);
    if(!p) return res.status(404).json({ error:'不存在' });
    let d=p.image_data,e=p.image_ext;
    if(req.file){d=bufToB64(req.file.buffer);e=getExt(req.file.originalname);} else if(removeImage==='1'){d=null;e=null;}
    await run('UPDATE product_groups SET name=$1,image_data=$2,image_ext=$3 WHERE id=$4',[name,d,e,req.params.id]);
    res.json({ ok:true });
  } catch (er) { res.status(500).json({ error:'修改失败' }); }
});

app.delete('/api/products/:id', adminOnly, async (req, res) => {
  try {
    if(!(await queryOne('SELECT id FROM product_groups WHERE id=$1',[req.params.id]))) return res.status(404).json({ error:'不存在' });
    await run('DELETE FROM product_groups WHERE id=$1',[req.params.id]);
    res.json({ ok:true });
  } catch (e) { res.status(500).json({ error:'删除失败' }); }
});

// 规格管理
app.post('/api/products/:gid/variants', adminOnly, async (req, res) => {
  try {
    const { name, stock, price } = req.body;
    if(!name) return res.status(400).json({ error:'规格名不能为空' });
    if(!(await queryOne('SELECT id FROM product_groups WHERE id=$1',[req.params.gid])))
      return res.status(404).json({ error:'商品组不存在' });
    const id=uuidv4();
    await run('INSERT INTO product_variants(id,group_id,name,stock,price) VALUES($1,$2,$3,$4,$5)',[id,req.params.gid,name,stock||0,price||0]);
    res.json({ok:true,variant:{id,group_id:req.params.gid,name,stock:stock||0,occupied:0,price:price||0}});
  } catch (e) { res.status(500).json({ error:'创建失败' }); }
});

app.put('/api/products/:gid/variants/:vid', adminOnly, async (req, res) => {
  try {
    const { name, stock, price } = req.body;
    if(!(await queryOne('SELECT * FROM product_variants WHERE id=$1 AND group_id=$2',[req.params.vid,req.params.gid])))
      return res.status(404).json({ error:'不存在' });
    await run('UPDATE product_variants SET name=$1,stock=$2,price=$3 WHERE id=$4',[name,stock,price,req.params.vid]);
    res.json({ok:true});
  } catch (e) { res.status(500).json({ error:'修改失败' }); }
});

app.delete('/api/products/:gid/variants/:vid', adminOnly, async (req, res) => {
  try {
    if(!(await queryOne('SELECT id FROM product_variants WHERE id=$1 AND group_id=$2',[req.params.vid,req.params.gid])))
      return res.status(404).json({ error:'不存在' });
    await run('DELETE FROM product_variants WHERE id=$1',[req.params.vid]);
    res.json({ok:true});
  } catch (e) { res.status(500).json({ error:'删除失败' }); }
});

// ============================================================
//  占用记录
// ============================================================
async function attachImages(recs) {
  const out=[];
  for(const r of recs){
    const imgs = await query('SELECT id,ext,original_name FROM payment_images WHERE record_id=$1',[r.id]);
    out.push({...r,images:imgs.map(i=>({id:i.id,file_path:`/api/images/payment/${i.id}`,ext:i.ext}))});
  }
  return out;
}

app.get('/api/records', auth, async (req, res) => {
  try {
    let recs = req.session.user.role==='admin'
      ? await query('SELECT * FROM records ORDER BY created_at DESC')
      : await query('SELECT * FROM records WHERE user_id=$1 ORDER BY created_at DESC',[req.session.user.id]);
    res.json({ records:await attachImages(recs) });
  } catch (e) { res.status(500).json({ error:'查询失败' }); }
});

app.post('/api/records', auth, upload.array('paymentImages',10), async (req, res) => {
  try {
    const { cinemaSpecialId,cinemaName,variantId,qty,price,receiverName,receiverPhone,province,city,district,address,bdId } = req.body;

    if(!variantId) return res.status(400).json({ error:'请选择商品规格' });
    const q=parseInt(qty); if(!q||q<1) return res.status(400).json({ error:'数量>=1' });
    if(!receiverName) return res.status(400).json({ error:'收件人姓名必填' });
    if(!receiverPhone) return res.status(400).json({ error:'手机号必填' });
    if(!address) return res.status(400).json({ error:'地址必填' });
    if(!req.files?.length) return res.status(400).json({ error:'请上传付款截图' });

    const variant = await queryOne(
      'SELECT v.*,pg.name AS group_name,pg.id AS group_id FROM product_variants v JOIN product_groups pg ON v.group_id=pg.id WHERE v.id=$1',
      [variantId]
    );
    if(!variant) return res.status(400).json({ error:'商品规格不存在' });
    const remain = Number(variant.stock)-Number(variant.occupied);
    if(q > remain) return res.status(400).json({ error:`库存不足，最多${remain}件` });

    const p = parseFloat(price)||parseFloat(variant.price);
    const recId = uuidv4();

    await run(
      `INSERT INTO records(id,user_id,bd_name,bd_id,cinema_special_id,cinema_name,product_group_id,product_group_name,variant_id,variant_name,qty,price,subtotal,receiver_name,receiver_phone,province,city,district,address)
       VALUES($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19)`,
      [recId,req.session.user.id,req.session.user.name,bdId||'',cinemaSpecialId||'',cinemaName||'',
       variant.group_id,variant.group_name,variant.id,variant.name,q,p,q*p,
       receiverName,receiverPhone,province||'',city||'',district||'',address]);

    for(const f of req.files){
      await run('INSERT INTO payment_images(id,record_id,data,ext,original_name) VALUES($1,$2,$3,$4,$5)',
        [uuidv4(),recId,bufToB64(f.buffer),getExt(f.originalname),f.originalname]);
    }

    await run('UPDATE product_variants SET occupied=occupied+$1 WHERE id=$2',[q,variantId]);
    res.json({ok:true,recordId:recId});
  } catch (e) { console.error('[Record create]',e); res.status(500).json({ error:'提交失败:'+e.message }); }
});

app.delete('/api/records/:id', adminOnly, async (req, res) => {
  try {
    const r = await queryOne('SELECT * FROM records WHERE id=$1',[req.params.id]);
    if(!r) return res.status(404).json({ error:'不存在' });
    await run('UPDATE product_variants SET occupied=GREATEST(0,occupied-$1) WHERE id=$2',[r.qty,r.variant_id]);
    await run('DELETE FROM records WHERE id=$1',[req.params.id]);
    res.json({ok:true});
  } catch (e) { res.status(500).json({ error:'删除失败' }); }
});

// ============================================================
//  看板
// ============================================================
app.get('/api/dashboard', adminOnly, async (_req, res) => {
  try {
    const todayStr = new Date().toISOString().slice(0,10);
    const allRecs  = await query('SELECT * FROM records');
    const todayRecs = allRecs.filter(r=>r.created_at&&r.created_at.slice(0,10)===todayStr);
    const variants = await query('SELECT v.*,pg.name AS group_name FROM product_variants v JOIN product_groups pg ON v.group_id=pg.id');

    const trend=[];
    for(let i=6;i>=0;i--){
      const d=new Date();d.setDate(d.getDate()-i);
      const ds=d.toISOString().slice(0,10);
      const dr=allRecs.filter(r=>r.created_at&&r.created_at.slice(0,10)===ds);
      trend.push({date:ds,qty:dr.reduce((s,r)=>s+Number(r.qty),0),
                  amount:dr.reduce((s,r)=>s+Number(r.subtotal),0),count:dr.length});
    }

    const bdRank = await query(
      "SELECT bd_name,SUM(qty)::int AS tq,SUM(subtotal)::float AS ta,COUNT(*) AS c FROM records WHERE created_at>=$1 GROUP BY bd_name ORDER BY tq DESC LIMIT 10",
      [todayStr+'T00:00:00']
    );

    res.json({
      today:{ occupied:todayRecs.reduce((s,r)=>s+Number(r.qty),0),
              amount:todayRecs.reduce((s,r)=>s+Number(r.subtotal),0),count:todayRecs.length },
      total:{ occupied:allRecs.reduce((s,r)=>s+Number(r.qty),0),
              amount:allRecs.reduce((s,r)=>s+Number(r.subtotal),0),count:allRecs.length },
      inventory: variants.map(v=>{
        const s=Number(v.stock),o=Number(v.occupied);
        return { name:v.group_name+' - '+v.name, stock:s, occupied:o, remain:s-o, rate:s>0?Math.round(o/s*100):0 };
      }),
      trend,
      bdRanking:bdRank.map(r=>({...r,total_qty:r.tq,total_amount:r.ta,count:r.c})),
    });
  } catch (e) { console.error('[Dash]',e); res.status(500).json({ error:'查询失败' }); }
});

// ============================================================
//  导出 Excel（含真实截图）
// ============================================================
app.get('/api/export', adminOnly, async (req, res) => {
  try {
    let recs = await query('SELECT * FROM records ORDER BY created_at ASC');
    const { dateFrom,dateTo,bdName,productId } = req.query;
    if(dateFrom) recs=recs.filter(r=>r.created_at&&new Date(r.created_at)>=new Date(dateFrom));
    if(dateTo)   recs=recs.filter(r=>r.created_at&&new Date(r.created_at)<=new Date(dateTo)+'T23:59:59');
    if(bdName)   recs=recs.filter(r=>r.bd_name&&r.bd_name.includes(bdName));
    if(productId)recs=recs.filter(r=>r.product_group_id===productId);
    if(recs.length===0) return res.status(400).json({ error:'无数据可导出' });

    const ridList = recs.map(r=>r.id);
    const allImgs = await query('SELECT * FROM payment_images WHERE record_id=ANY($1)',[ridList]);

    let maxImgs=0; const imgsBy={};
    for(const im of allImgs){
      if(!imgsBy[im.record_id]) imgsBy[im.record_id]=[];
      imgsBy[im.record_id].push(im);
      if(imgsBy[im.record_id].length>maxImgs) maxImgs=imgsBy[im.record_id].length;
    }

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('库存占用数据');
    const hdr=['提交时间','BD姓名','BD工号','影城专资ID','影城名称','收件人姓名','收件人手机号','省','市','区','收件人地址','商品名称','货品规格','货品数量','货品单价(¥)','金额小计(¥)'];
    for(let i=0;i<Math.max(maxImgs,1);i++) hdr.push('付款截图'+(maxImgs>1?' '+(i+1):''));

    ws.columns = hdr.map((h,i)=>({header:h,width:i>=16?22:h.includes('地址')?30:h.includes('手机')?16:14}));
    ws.getRow(1).height=25;
    ws.getRow(1).eachCell(c=>{c.font={bold:true,color:{argb:'FFFFFFFF'},size:11};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF4472C4'}};c.alignment={vertical:'middle',horizontal:'center'};});

    for(let i=0;i<recs.length;i++){
      const r=recs[i];
      const row = ws.addRow([
        r.created_at?new Date(r.created_at).toLocaleString('zh-CN'):'',r.bd_name,r.bd_id||'',
        r.cinema_special_id||'',r.cinema_name||'',
        r.receiver_name,r.receiver_phone,r.province||'',r.city||'',r.district||'',r.address,
        r.product_group_name,r.variant_name,r.qty,r.price,r.subtotal,
      ]);
      row.height=75;
      const imgs=(imgsBy[r.id]||[]);
      for(let j=0;j<imgs.length;j++){
        try{
          const buf=Buffer.from(imgs[j].data,'base64');
          const ext=(imgs[j].ext||'jpeg').replace('jpg','jpeg');
          const iid=wb.addImage({buffer:buf,extension:ext});
          ws.addImage(iid,{tl:{col:16+j,row:i+1},ext:{width:120,height:65}});
        }catch(_){}
      }
    }
    ws.autoFilter={from:{row:1,column:1},to:{row:recs.length+1,column:hdr.length}};
    const fn=`库存导出_${new Date().toISOString().slice(0,10)}.xlsx`;
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',`attachment; filename="${encodeURIComponent(fn)}"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (e) { console.error('[Export]',e); res.status(500).json({ error:'导出失败:'+e.message }); }
});

// SPA fallback
app.get('*', (_req, res) => { res.sendFile(path.join(__dirname,'public','index.html')); });

// ============================================================
//  启动
// ============================================================
(async () => {
  try {
    await initDB();
    console.log('\n✅ 数据库初始化完成\n');
  } catch (e) {
    console.error('\n❌ 数据库初始化失败！', e.message, '\n');
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`========================================`);
    console.log(`  🚀 库存管理系统已启动`);
    console.log(`  端口: ${PORT}`);
    console.log(`  管理员: admin / admin123`);
    console.log(`  健康检查: GET /api/health`);
    console.log(`========================================\n`);
  });
})();
