/*******************************************************
 * V&P TRADING — 销售 / 送货 / 司机抽成 系统
 * Google Apps Script 后端
 *
 * 需要的分页：ORDERS, SET_PRICE, BRANCH, SALESMAN,
 *            DRIVER_RULE, DRIVER, USERS, CONFIG
 *******************************************************/

var SS = SpreadsheetApp.getActiveSpreadsheet();
var TZ = Session.getScriptTimeZone() || 'Asia/Kuala_Lumpur';

/* ---------- 入口 ---------- */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('V&P TRADING')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- 工具 ---------- */
function sh_(name) {
  var s = SS.getSheetByName(name);
  if (!s) throw new Error('找不到分页：' + name + '（请确认已导入数据表）');
  return s;
}

/** 读整张表 → 物件阵列。自动跳过表头上方的说明行。 */
function readTable_(name) {
  var sheet = sh_(name);
  var v = sheet.getDataRange().getValues();
  if (!v.length) return { head: [], rows: [], headRow: 1, sheet: sheet };
  var hr = 0;
  for (var i = 0; i < Math.min(v.length, 5); i++) {
    var nonEmpty = v[i].filter(function (x) { return x !== '' && x !== null; }).length;
    if (nonEmpty >= 2) { hr = i; break; }
  }
  var head = v[hr].map(function (h) { return String(h).trim(); });
  var rows = [];
  for (var r = hr + 1; r < v.length; r++) {
    if (v[r].every(function (x) { return x === '' || x === null; })) continue;
    var o = {};
    for (var c = 0; c < head.length; c++) if (head[c]) o[head[c]] = v[r][c];
    o.__row = r + 1;
    rows.push(o);
  }
  return { head: head, rows: rows, headRow: hr + 1, sheet: sheet };
}

function toNum_(x) { var n = parseFloat(x); return isNaN(n) ? 0 : n; }
function up_(x) { return String(x == null ? '' : x).trim().toUpperCase(); }

function config_() {
  var t = readTable_('CONFIG'), o = {};
  t.rows.forEach(function (r) { o[up_(r.KEY)] = r.VALUE; });
  return o;
}

/* ---------- 登入 ---------- */
function login(pin) {
  var t = readTable_('USERS');
  var hit = null;
  t.rows.forEach(function (r) {
    if (String(r.PIN).trim() === String(pin).trim() && up_(r.ACTIVE) !== 'NO') hit = r;
  });
  if (!hit) return { ok: false, msg: 'PIN 不正确' };
  return { ok: true, name: String(hit.NAME).trim(), role: String(hit.ROLE).trim().toLowerCase() };
}

/** 登入 + 载入资料一次搞定（少一次往返，快一倍） */
function loginAndBoot(pin) {
  var r = login(pin);
  if (!r.ok) return r;
  return { ok: true, user: { name: r.name, role: r.role }, boot: bootstrap() };
}

/** 确保 ORDERS 有这些栏位，没有就补上 */
function ensureCols_(name, cols) {
  var t = readTable_(name);
  var head = t.head.slice(), added = 0;
  cols.forEach(function (c) { if (head.indexOf(c) < 0) { head.push(c); added++; } });
  if (!added) return t;
  t.sheet.getRange(t.headRow, 1, 1, head.length).setValues([head])
    .setBackground('#1C1C1E').setFontColor('#E8C86A').setFontWeight('bold')
    .setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center');
  return readTable_(name);
}

function isVoid_(r) { return up_(r.STATUS) === 'VOID'; }


/* ---------- 并发保护 ----------
   锁只包住「写」那一刻。读资料请在呼叫 withLock_ 之前做好 ——
   把读放进锁里会让一次收款抓着锁五到八秒，别人就撞到 Lock timeout。
--------------------------------- */
function friendlyErr_(e) {
  var m = String((e && e.message) || e || '');
  if (/lock/i.test(m)) return '有人同时在存资料，请等两秒再按一次';
  if (/timed?\s*out|deadline|exceeded maximum execution/i.test(m)) return '网络有点慢，请再试一次';
  if (/permission|authoriz/i.test(m)) return '权限不足，请用 admin 重新登入再试';
  return m || '存档失败，请再试一次';
}

/** 试着拿锁，拿不到回传 false（不抛例外）。 */
function tryLock_(lock, ms) {
  try { lock.waitLock(ms); return true; } catch (e) { return false; }
}

/** 拿锁 → 跑 fn → 放锁。拿不到会自动重试两次。 */
function withLock_(fn) {
  var lock = LockService.getScriptLock(), got = false;
  for (var i = 0; i < 3 && !got; i++) {
    try { lock.waitLock(i === 0 ? 8000 : 5000); got = true; }
    catch (e) { if (i < 2) Utilities.sleep(400 + i * 500); }
  }
  if (!got) return { ok: false, msg: '有人同时在存资料，请等两秒再按一次' };
  try { return fn(); }
  catch (e) { return { ok: false, msg: friendlyErr_(e) }; }
  finally { try { lock.releaseLock(); } catch (e2) { } }
}

/** 锁内用：读单行 → 核对单号没被别人动过 → 套用修改 → 整行写回。一读一写。 */
function patchRow_(sheet, head, row, orderId, patch) {
  var rng = sheet.getRange(row, 1, 1, head.length);
  var v = rng.getValues()[0];
  var idc = head.indexOf('ORDER_ID');
  if (idc >= 0 && orderId != null &&
      String(v[idc]).trim() !== String(orderId).trim())
    return { ok: false, msg: '这笔订单刚刚被别人改过，请重新整理后再试一次' };
  var touched = false;
  Object.keys(patch).forEach(function (k) {
    var c = head.indexOf(k);
    if (c >= 0) { v[c] = patch[k]; touched = true; }
  });
  if (touched) rng.setValues([v]);
  return { ok: true };
}

/* ---------- 付款状态 ----------
 * 客户的原始写法：OP = Online Pay、PC = Pay Cash，两者都代表「已付」；空白 = 还没付。
 * 新模型把「状态」和「方式」分开：
 *   PAY_STATUS = PAID / UNPAID
 *   PAY_METHOD = OP / PC
 * 下面两个函式同时看得懂新旧写法，所以就算还没跑迁移，金额也不会算错。
 */
function payPaid_(r) {
  var s = up_(r.PAY_STATUS);
  if (s === 'PAID' || s === 'OP' || s === 'PC') return true;
  return false;   // UNPAID、空白、PENDING 一律当未付
}
function payMethod_(r) {
  var m = up_(r.PAY_METHOD);
  if (m === 'OP' || m === 'PC') return m;
  var s = up_(r.PAY_STATUS);
  return (s === 'OP' || s === 'PC') ? s : '';
}

/* ---------- 帐号管理（只有 admin 能用） ---------- */
/** 每个动作都用 admin 自己的 PIN 重新验证一次，前端绕不过 */
function adminAuth_(pin) {
  var t = readTable_('USERS'), hit = null;
  t.rows.forEach(function (r) {
    if (String(r.PIN).trim() === String(pin).trim() && up_(r.ACTIVE) !== 'NO') hit = r;
  });
  if (!hit || String(hit.ROLE).trim().toLowerCase() !== 'admin') return null;
  return { row: hit, t: t };
}
function validPin_(p) { return /^\d{4}$/.test(String(p == null ? '' : p).trim()); }

function adminListUsers(adminPin) {
  var a = adminAuth_(adminPin);
  if (!a) return { ok: false, msg: '只有 admin 可以管理帐号' };
  return {
    ok: true,
    users: a.t.rows.filter(function (r) { return r.NAME; }).map(function (r) {
      return {
        name: String(r.NAME).trim(), pin: String(r.PIN).trim(),
        role: String(r.ROLE).trim().toLowerCase(), active: up_(r.ACTIVE) !== 'NO'
      };
    })
  };
}

function adminSetPin(adminPin, target, newPin) {
  var lock = LockService.getScriptLock();
  try {
    if (!tryLock_(lock, 15000)) return { ok: false, msg: '有人同时在存资料，请等两秒再按一次' };
    var a = adminAuth_(adminPin);
    if (!a) return { ok: false, msg: '只有 admin 可以改密码' };
    if (!validPin_(newPin)) return { ok: false, msg: 'PIN 必须是 4 位数字' };
    var t = a.t, hit = null, dup = false;
    t.rows.forEach(function (r) {
      if (up_(r.NAME) === up_(target)) hit = r;
      else if (String(r.PIN).trim() === String(newPin).trim()) dup = true;
    });
    if (!hit) return { ok: false, msg: '找不到使用者 ' + target };
    if (dup) return { ok: false, msg: '这个 PIN 已经有人在用，换一个' };
    var c = t.head.indexOf('PIN') + 1;
    if (c <= 0) return { ok: false, msg: 'USERS 分页找不到 PIN 栏' };
    t.sheet.getRange(hit.__row, c).setValue(String(newPin).trim());
    return { ok: true };
  } catch (e) { return { ok: false, msg: friendlyErr_(e) }; } finally { try { lock.releaseLock(); } catch (e2) { } }
}

/** 改帐号名字（例如 UNCLE → TAN） */
function adminRenameUser(adminPin, target, newName) {
  var lock = LockService.getScriptLock();
  try {
    if (!tryLock_(lock, 15000)) return { ok: false, msg: '有人同时在存资料，请等两秒再按一次' };
    var a = adminAuth_(adminPin);
    if (!a) return { ok: false, msg: '只有 admin 可以改名字' };
    newName = String(newName == null ? '' : newName).trim();
    if (!newName) return { ok: false, msg: '名字不可以空白' };
    if (newName.length > 20) return { ok: false, msg: '名字太长了（最多 20 个字）' };
    var t = a.t, hit = null, dup = false;
    t.rows.forEach(function (r) {
      if (up_(r.NAME) === up_(target)) hit = r;
      else if (up_(r.NAME) === up_(newName)) dup = true;
    });
    if (!hit) return { ok: false, msg: '找不到使用者 ' + target };
    if (dup) return { ok: false, msg: '「' + newName + '」这个名字已经有人在用' };
    var c = t.head.indexOf('NAME') + 1;
    if (c <= 0) return { ok: false, msg: 'USERS 分页找不到 NAME 栏' };
    t.sheet.getRange(hit.__row, c).setValue(newName);
    return { ok: true, name: newName };
  } catch (e) { return { ok: false, msg: friendlyErr_(e) }; } finally { try { lock.releaseLock(); } catch (e2) { } }
}

function adminAddUser(adminPin, name, pin, role) {
  var lock = LockService.getScriptLock();
  try {
    if (!tryLock_(lock, 15000)) return { ok: false, msg: '有人同时在存资料，请等两秒再按一次' };
    var a = adminAuth_(adminPin);
    if (!a) return { ok: false, msg: '只有 admin 可以新增帐号' };
    name = String(name || '').trim();
    role = String(role || '').trim().toLowerCase();
    if (!name) return { ok: false, msg: '请填名字' };
    if (!validPin_(pin)) return { ok: false, msg: 'PIN 必须是 4 位数字' };
    if (['admin', 'partner', 'driver'].indexOf(role) < 0) return { ok: false, msg: '身分只能是 admin / partner / driver' };
    var t = a.t, bad = null;
    t.rows.forEach(function (r) {
      if (up_(r.NAME) === up_(name)) bad = '这个名字已经存在';
      if (String(r.PIN).trim() === String(pin).trim()) bad = '这个 PIN 已经有人在用';
    });
    if (bad) return { ok: false, msg: bad };
    t.sheet.appendRow(t.head.map(function (h) {
      return h === 'NAME' ? name : h === 'PIN' ? String(pin).trim()
        : h === 'ROLE' ? role : h === 'ACTIVE' ? 'YES' : '';
    }));
    return { ok: true };
  } catch (e) { return { ok: false, msg: friendlyErr_(e) }; } finally { try { lock.releaseLock(); } catch (e2) { } }
}

/* ---------- 一次性资料升级：旧的 OP/PC 付款写法 → 新模型 ---------- *
 * 旧：PAY_STATUS = OP / PC / 空白，PAY_DATE = "02.01.26 (HL Bank)"
 * 新：PAY_STATUS = PAID / UNPAID
 *     PAY_METHOD = OP / PC
 *     PAY_DATE   = 2026-01-02
 *     PAY_NOTE   = HL Bank
 * 可以重复执行，已经是新格式的行会自动略过。
 */
function migratePayment(adminPin) {
  var lock = LockService.getScriptLock();
  try {
    if (!tryLock_(lock, 120000)) return { ok: false, msg: '有人同时在存资料，请等两秒再按一次' };
    if (adminPin !== undefined && !adminAuth_(adminPin))
      return { ok: false, msg: '只有 admin 可以执行资料升级' };

    var t = ensureCols_('ORDERS', ['PAY_METHOD', 'PAY_NOTE']);
    var sheet = t.sheet, head = t.head;
    var cS = head.indexOf('PAY_STATUS'), cM = head.indexOf('PAY_METHOD'),
        cD = head.indexOf('PAY_DATE'), cN = head.indexOf('PAY_NOTE'),
        cO = head.indexOf('DATE'), cSrc = head.indexOf('SOURCE');
    if (cS < 0 || cM < 0 || cD < 0 || cN < 0) return { ok: false, msg: 'ORDERS 缺少付款栏位' };

    var first = t.headRow + 1, n = t.rows.length;
    if (!n) return { ok: false, msg: 'ORDERS 没有资料' };

    var rng = sheet.getRange(first, 1, n, head.length);
    var v = rng.getValues();
    var stat = { paid: 0, unpaid: 0, op: 0, pc: 0, dated: 0, noted: 0, skipped: 0, fixed: 0, noMethod: 0 };

    for (var i = 0; i < n; i++) {
      var s = String(v[i][cS] == null ? '' : v[i][cS]).trim().toUpperCase();
      if (s === 'PAID' || s === 'UNPAID') { stat.skipped++; continue; }   // 已是新格式

      var raw = String(v[i][cD] == null ? '' : v[i][cD]).trim();
      var src = cSrc >= 0 ? String(v[i][cSrc] == null ? '' : v[i][cSrc]) : '';

      // 从 "02.01.26 (HL Bank)" / "TNG EWALLET 07.01.26" / "19.05" 抽出日期与备注
      var iso = '', note = '';
      var m = raw.match(/(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{2,4})/);
      if (m) {
        var yy = m[3].length === 2 ? '20' + m[3] : m[3];
        iso = yy + '-' + pad2_(m[2]) + '-' + pad2_(m[1]);
        note = raw.replace(m[0], ' ');
      } else {
        // 只写了「日.月」没写年份 → 用这笔订单本身的年份
        var m2 = raw.match(/(\d{1,2})[.\-\/](\d{1,2})(?!\d)/);
        if (m2) {
          var od = cO >= 0 ? String(v[i][cO] || '') : '';
          var oy = (od.match(/^(\d{4})/) || [])[1] ||
                   String(new Date().getFullYear());
          iso = oy + '-' + pad2_(m2[2]) + '-' + pad2_(m2[1]);
          note = raw.replace(m2[0], ' ');
        } else { note = raw; }
      }
      note = note.replace(/[()（）]/g, ' ').replace(/\band\b/gi, ' ').replace(/\s+/g, ' ').trim();

      // ── 判定收了没有 ──
      // OP / PC 只是「钱怎么进来的」，不代表收到了。真正的付款记录 = 有收款日期。
      // 旧资料里 1010 笔写了方式的，1010 笔都有日期 —— 因为他们是钱到手才填那一行。
      // 所以：抓不到日期 = 还没收到钱。
      var method = (s === 'OP' || s === 'PC') ? s : '';
      var paid = !!iso;
      if (src.indexOf('APP') === 0) paid = !!iso || !!method;   // App 下的单另有把关，不动它

      // 本来看起来像已收（写了方式）、但其实没有日期的，就是这次要救回来的
      if (!paid && (method || s)) stat.fixed++;
      if (!paid) { method = ''; iso = ''; note = ''; }
      // 有日期但没写 OP/PC 的：方式就留空白，不替客户猜。他自己去 App 按一下补上。
      if (paid && !method) stat.noMethod++;

      v[i][cS] = paid ? 'PAID' : 'UNPAID';
      v[i][cM] = paid ? method : '';
      v[i][cD] = paid ? iso : '';
      v[i][cN] = paid ? note : '';

      if (paid) { stat.paid++; if (method === 'OP') stat.op++; if (method === 'PC') stat.pc++; if (iso) stat.dated++; if (note) stat.noted++; }
      else stat.unpaid++;
    }

    rng.setValues(v);
    clearBootCache_();
    return {
      ok: true, total: n, stat: stat,
      msg: '升级完成：共 ' + n + ' 笔｜已付 ' + stat.paid + '（OP ' + stat.op + ' · PC ' + stat.pc +
           '）｜未付 ' + stat.unpaid + '（其中 ' + stat.fixed + ' 笔是修正回来的）｜已是新格式略过 ' + stat.skipped
    };
  } catch (e) { return { ok: false, msg: friendlyErr_(e) }; } finally { try { lock.releaseLock(); } catch (e2) { } }
}

/* ---------- 司机设定（限 admin） ---------- */
function adminGetDriver(adminPin) {
  var a = adminAuth_(adminPin);
  if (!a) return { ok: false, msg: '只有 admin 可以看司机设定' };
  var r = readTable_('DRIVER').rows[0] || {};
  return {
    ok: true, driver: {
      name: String(r.DRIVER_NAME || ''),
      phone: String(r.PHONE || '').replace(/[^0-9]/g, ''),
      allowance: toNum_(r.ALLOWANCE_PER_MONTH)
    }
  };
}

function adminSetDriver(adminPin, p) {
  var lock = LockService.getScriptLock();
  try {
    if (!tryLock_(lock, 15000)) return { ok: false, msg: '有人同时在存资料，请等两秒再按一次' };
    var a = adminAuth_(adminPin);
    if (!a) return { ok: false, msg: '只有 admin 可以改司机设定' };
    var name = String(p.name == null ? '' : p.name).trim();
    var phone = String(p.phone == null ? '' : p.phone).replace(/[^0-9]/g, '');
    var allow = toNum_(p.allowance);
    if (!name) return { ok: false, msg: '请填司机名字' };
    // 本地格式自动转国际格式：012-345 6789 → 60123456789
    if (phone && phone.charAt(0) === '0') phone = '60' + phone.slice(1);
    if (phone && (phone.length < 10 || phone.length > 15))
      return { ok: false, msg: '电话号码看起来不完整。直接打本地号码就行（例如 0123456789），系统会自动转成 60123456789。' };
    if (allow < 0) return { ok: false, msg: 'Allowance 不能是负数' };

    var t = readTable_('DRIVER');
    var row = t.rows[0];
    if (!row) {
      t.sheet.appendRow(t.head.map(function (h) {
        return h === 'DRIVER_NAME' ? name : h === 'PHONE' ? phone
          : h === 'ALLOWANCE_PER_MONTH' ? allow : h === 'ACTIVE' ? 'YES' : '';
      }));
    } else {
      var col = function (c) { return t.head.indexOf(c) + 1; };
      if (col('DRIVER_NAME') > 0) t.sheet.getRange(row.__row, col('DRIVER_NAME')).setValue(name);
      // 电话强制存成文字，避免 Sheets 把它当数字（会掉开头的 0 或变科学记号）
      if (col('PHONE') > 0) t.sheet.getRange(row.__row, col('PHONE')).setNumberFormat('@').setValue(phone);
      if (col('ALLOWANCE_PER_MONTH') > 0) t.sheet.getRange(row.__row, col('ALLOWANCE_PER_MONTH')).setValue(allow);
    }
    clearBootCache_();
    return { ok: true, phone: phone };
  } catch (e) { return { ok: false, msg: friendlyErr_(e) }; } finally { try { lock.releaseLock(); } catch (e2) { } }
}

function adminSetActive(adminPin, target, active) {
  var a = adminAuth_(adminPin);
  if (!a) return { ok: false, msg: '只有 admin 可以停用帐号' };
  var t = a.t, hit = null;
  t.rows.forEach(function (r) { if (up_(r.NAME) === up_(target)) hit = r; });
  if (!hit) return { ok: false, msg: '找不到使用者' };
  if (up_(hit.NAME) === up_(a.row.NAME) && !active) return { ok: false, msg: '不能停用自己' };
  var c = t.head.indexOf('ACTIVE') + 1;
  if (c <= 0) return { ok: false, msg: 'USERS 分页找不到 ACTIVE 栏' };
  t.sheet.getRange(hit.__row, c).setValue(active ? 'YES' : 'NO');
  return { ok: true };
}

/* ---------- 启动资料 ---------- */
function bootstrap() {
  var cache = CacheService.getScriptCache();
  try { var hit = cache.get('boot'); if (hit) return JSON.parse(hit); } catch (e) { }
  var o = buildBoot_();
  try { cache.put('boot', JSON.stringify(o), 900); } catch (e) { }
  return o;
}
function clearBootCache_() { try { CacheService.getScriptCache().remove('boot'); } catch (e) { } }

function buildBoot_() {
  var cfg = config_();
  var sp = readTable_('SET_PRICE').rows.filter(function (r) { return up_(r.ACTIVE) !== 'NO' && r.SET_TYPE; });
  var setTypes = {};
  sp.forEach(function (r) {
    var k = String(r.SET_TYPE).trim();
    if (!setTypes[k]) setTypes[k] = [];
    setTypes[k].push({ price: toNum_(r.UNIT_PRICE), profit: toNum_(r.PROFIT_PER_SET),
                       desc: String(r['说明'] || r.DESC || '').trim() });
  });
  Object.keys(setTypes).forEach(function (k) {
    setTypes[k].sort(function (a, b) { return a.price - b.price; });
  });

  var sm = readTable_('SALESMAN').rows;
  var people = {};
  sm.forEach(function (r) {
    var n = up_(r.SALESMAN); if (!n) return;
    if (!people[n]) people[n] = [];
    people[n].push({
      branch: String(r.BRANCH).trim(), brand: String(r.BRAND || '').trim(),
      state: String(r.STATE || '').trim(), region: String(r.REGION || '').trim(),
      n: toNum_(r['历史笔数']), primary: String(r['主要分行'] || '').indexOf('★') >= 0
    });
  });
  Object.keys(people).forEach(function (n) {
    people[n].sort(function (a, b) { return (b.primary ? 1 : 0) - (a.primary ? 1 : 0) || b.n - a.n; });
  });

  var branches = readTable_('BRANCH').rows
    .filter(function (r) { return up_(r.ACTIVE) !== 'NO' && r.BRANCH; })
    .map(function (r) {
      return {
        branch: String(r.BRANCH).trim(), brand: String(r.BRAND || '').trim(),
        state: String(r.STATE || '').trim(), region: String(r.REGION || '').trim()
      };
    });

  var drv = readTable_('DRIVER').rows.filter(function (r) { return up_(r.ACTIVE) !== 'NO'; })[0] || {};

  return {
    company: cfg.COMPANY || 'V&P TRADING',
    currency: cfg.CURRENCY || 'RM',
    payOptions: String(cfg.PAY_STATUS_OPTIONS || 'OP,PC,PAID,PENDING').split(',').map(function (s) { return s.trim(); }),
    defaultPay: cfg.DEFAULT_PAY_STATUS || 'OP',
    driver: { name: String(drv.DRIVER_NAME || ''), phone: String(drv.PHONE || '').replace(/[^0-9]/g, ''), allowance: toNum_(drv.ALLOWANCE_PER_MONTH) },
    setTypes: setTypes,
    people: people,
    branches: branches
  };
}

/* ---------- 计算：司机抽成 ---------- */
function calcDriverFee_(region, state, setType, qty, price) {
  var rules = readTable_('DRIVER_RULE').rows;
  var R = up_(region), S = up_(state);
  var rule = null;
  rules.forEach(function (r) {
    if (up_(r.REGION) === R && up_(r.STATE) === S) rule = r;
  });
  if (!rule) rules.forEach(function (r) { if (!rule && up_(r.REGION) === R) rule = r; });
  if (!rule) return 0;

  var type = up_(rule.RULE_TYPE);
  var isUmbrella = up_(setType).indexOf('UMBRELLA') >= 0;
  if (type === 'NONE') return 0;
  if (type === 'PERCENT') {
    var pct = toNum_(isUmbrella ? rule.UMBRELLA_RATE : rule.RATE);
    if (pct > 1) pct = pct / 100;
    return Math.round(qty * price * pct * 100) / 100;
  }
  // PER_SET
  var rate = toNum_(isUmbrella ? rule.UMBRELLA_RATE : rule.RATE);
  return Math.round(qty * rate * 100) / 100;
}

/** 前端即时预览用（不写入） */
function quote(p) {
  var qty = toNum_(p.qty), price = toNum_(p.price), profitPerSet = toNum_(p.profitPerSet);
  var total = Math.round(qty * price * 100) / 100;
  var mine = Math.round(qty * profitPerSet * 100) / 100;
  var fee = calcDriverFee_(p.region, p.state, p.setType, qty, price);
  return { total: total, myIncome: mine, driverFee: fee };
}

/* ---------- 写入订单 ---------- */
function submitOrder(p) {
  var qty = toNum_(p.qty), price = toNum_(p.price);
  if (!p.salesman) return { ok: false, msg: '请选择销售员' };
  if (!p.branch) return { ok: false, msg: '请选择 BRANCH' };
  if (!p.setType) return { ok: false, msg: '请选择 SET' };
  if (qty <= 0) return { ok: false, msg: '数量必须大于 0' };

  var t = ensureCols_('ORDERS', ['DRV_COLOR', 'DRV_GRADE', 'DRV_NOTE']);
  var sheet = t.sheet, head = t.head;

  var total = (p.totalOverride !== '' && p.totalOverride != null)
    ? toNum_(p.totalOverride) : Math.round(qty * price * 100) / 100;
  var mine = (p.myIncomeOverride !== '' && p.myIncomeOverride != null)
    ? toNum_(p.myIncomeOverride) : Math.round(qty * toNum_(p.profitPerSet) * 100) / 100;
  var fee = (p.driverFeeOverride !== '' && p.driverFeeOverride != null)
    ? toNum_(p.driverFeeOverride) : calcDriverFee_(p.region, p.state, p.setType, qty, price);

  var maxN = 0;
  t.rows.forEach(function (r) {
    var m = String(r.ORDER_ID || '').match(/VP(\d+)/);
    if (m) maxN = Math.max(maxN, parseInt(m[1], 10));
  });

  var d = p.date ? new Date(p.date) : new Date();
  var paid = up_(p.payStatus) === 'PAID';
  var mk = function (id) {
    return {
      ORDER_ID: id,
      DATE: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'),
      MONTH: d.getMonth() + 1,
      REGION: p.region, STATE: p.state, BRANCH: p.branch, BRAND: p.brand,
      SALESMAN: up_(p.salesman), SET_TYPE: p.setType,
      UNIT_PRICE: price, QTY: qty,
      TOTAL_INCOME: total, MY_INCOME: mine, DRIVER_FEE: fee,
      PAY_STATUS: paid ? 'PAID' : 'UNPAID',
      PAY_METHOD: paid ? up_(p.payMethod) : '',
      PAY_DATE: paid ? (p.payDate || Utilities.formatDate(d, TZ, 'yyyy-MM-dd')) : '',
      PAY_NOTE: paid ? String(p.payNote || '').trim() : '',
      DRV_COLOR: String(p.drvColor || '').trim(),
      DRV_GRADE: String(p.drvGrade || '').trim(),
      DRV_NOTE: String(p.drvNote || '').trim(),
      DELIVERY_STATUS: 'PENDING', NOTE: p.note || '',
      SOURCE: 'APP·' + (p.by || '')
    };
  };

  var rec = null;
  var r = withLock_(function () {
    // 拿到锁才定单号 —— 免得两个人同时下单撞到同一号
    var last = sheet.getLastRow();
    if (last >= 3) {
      var idc = head.indexOf('ORDER_ID') + 1;
      if (idc > 0) {
        var m2 = String(sheet.getRange(last, idc).getValue() || '').match(/VP(\d+)/);
        if (m2) maxN = Math.max(maxN, parseInt(m2[1], 10));
      }
    }
    rec = mk('VP' + String(maxN + 1).padStart(5, '0'));
    sheet.appendRow(head.map(function (h) { return rec.hasOwnProperty(h) ? rec[h] : ''; }));
    return { ok: true };
  });
  if (!r.ok) return r;
  return { ok: true, order: rec, wa: buildWaText_([rec]) };
}


/* ---------- WhatsApp 文字 ---------- */
/** 单卖雨伞、袋子这类单品用 pcs；goodie bag 之类套装用 set。TAN 的习惯。 */
function unitOf_(setType) {
  var t = up_(setType);
  if (t.indexOf('+') >= 0) return 'set';          // UMBRELLA + BAG 算一组
  if (/UMBRELLA|BAG/.test(t) && !/ITEM/.test(t)) return 'pcs';
  return 'set';
}

/** 送货单上那句中文：颜色 + 款式 + 价目表说明，能凑几个凑几个 */
function goodsDesc_(o) {
  var bits = [];
  var c = String(o.DRV_COLOR || '').trim();
  var g = String(o.DRV_GRADE || '').trim();
  if (c) bits.push(c);
  if (g) bits.push(g);
  if (!bits.length) {
    var d = String(o.DRV_DESC || '').trim();     // 下单时从价目表带过来的
    if (d) bits.push(d);
  }
  return bits.length ? '（' + bits.join('') + '）' : '';
}

function buildWaText_(orders) {
  var cfg = config_();
  var co = cfg.COMPANY || 'V&P TRADING';
  var L = ['*' + co + ' — 送货单*',
           Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy'), ''];

  var byBranch = {}, order = [];
  orders.forEach(function (o) {
    var k = String(o.BRANCH || '-');
    if (!byBranch[k]) { byBranch[k] = []; order.push(k); }
    byBranch[k].push(o);
  });

  order.forEach(function (b) {
    L.push('*' + b + '*');
    byBranch[b].forEach(function (o) {
      var qty = toNum_(o.QTY), price = toNum_(o.UNIT_PRICE), tot = toNum_(o.TOTAL_INCOME);
      L.push(String(o.SALESMAN || '') + ' — ' + String(o.SET_TYPE || '') + goodsDesc_(o));
      L.push('RM ' + price + ' × ' + qty + ' ' + unitOf_(o.SET_TYPE) +
             ' = RM ' + (Math.round(tot * 100) / 100));
      var note = String(o.DRV_NOTE || '').trim();
      if (note) L.push('＊ ' + note);
      L.push(payPaid_(o) ? '已收' : '未收钱 RM ' + (Math.round(tot * 100) / 100));
      L.push('');
    });
  });

  var totalSets = orders.reduce(function (a, o) { return a + toNum_(o.QTY); }, 0);
  var totalFee = orders.reduce(function (a, o) { return a + toNum_(o.DRIVER_FEE); }, 0);
  var totalAmt = orders.reduce(function (a, o) { return a + toNum_(o.TOTAL_INCOME); }, 0);
  L.push('―――');
  L.push('共 ' + orders.length + ' 单 · ' + totalSets + ' 件 · RM ' + (Math.round(totalAmt * 100) / 100));
  L.push('你的抽成：RM ' + totalFee.toFixed(2));
  return L.join('\n');
}


/** 指定订单 ID 生成通知文字 */
function notifyDriver(ids) {
  var t = readTable_('ORDERS');
  var set = {}; (ids || []).forEach(function (i) { set[String(i)] = 1; });
  var picked = t.rows.filter(function (r) { return set[String(r.ORDER_ID)] && !isVoid_(r); });
  if (!picked.length) return { ok: false, msg: '没有选中的订单' };
  var drv = readTable_('DRIVER').rows[0] || {};
  var phone = String(drv.PHONE || '').replace(/[^0-9]/g, '');
  return { ok: true, phone: phone, text: buildWaText_(picked), count: picked.length };
}

/** 标记已通知 / 已送 */
function markDelivered(ids) {
  var t = readTable_('ORDERS');
  var col = t.head.indexOf('DELIVERY_STATUS') + 1;
  if (col <= 0) return { ok: false, msg: '找不到 DELIVERY_STATUS 栏' };
  var want = {}; (ids || []).forEach(function (i) { want[String(i)] = 1; });
  var rows = t.rows.filter(function (r) { return want[String(r.ORDER_ID)]; });
  if (!rows.length) return { ok: false, msg: '没有选中的订单' };

  // 连续的列合并成一个区块写，减少往返
  rows.sort(function (a, b) { return a.__row - b.__row; });
  var blocks = [], cur = null;
  rows.forEach(function (r) {
    if (cur && r.__row === cur.from + cur.n) { cur.n++; }
    else { cur = { from: r.__row, n: 1 }; blocks.push(cur); }
  });
  var res = withLock_(function () {
    blocks.forEach(function (b) {
      var vals = [];
      for (var i = 0; i < b.n; i++) vals.push(['DELIVERED']);
      t.sheet.getRange(b.from, col, b.n, 1).setValues(vals);
    });
    return { ok: true };
  });
  if (!res.ok) return res;
  return { ok: true, n: rows.length };
}


/* ---------- 作废 / 编辑订单 ---------- */

/** 依 SET_PRICE 查每 set 利润；查不到回 null */
function profitFor_(setType, price) {
  var hit = null;
  readTable_('SET_PRICE').rows.forEach(function (r) {
    if (String(r.SET_TYPE).trim() === String(setType).trim() && toNum_(r.UNIT_PRICE) === toNum_(price)) hit = r;
  });
  return hit ? toNum_(hit.PROFIT_PER_SET) : null;
}

function findOrder_(t, id) {
  var hit = null;
  t.rows.forEach(function (r) { if (String(r.ORDER_ID) === String(id)) hit = r; });
  return hit;
}

/** 作废（不真删，保留稽核轨迹） */
function voidOrder(id, by) {
  var t = ensureCols_('ORDERS', ['STATUS', 'VOID_BY', 'VOID_AT']);
  var hit = findOrder_(t, id);
  if (!hit) return { ok: false, msg: '找不到订单' };
  var stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  return withLock_(function () {
    return patchRow_(t.sheet, t.head, hit.__row, hit.ORDER_ID, {
      STATUS: 'VOID', VOID_BY: by || '', VOID_AT: stamp
    });
  });
}

function unvoidOrder(id) {
  var t = ensureCols_('ORDERS', ['STATUS', 'VOID_BY', 'VOID_AT']);
  var hit = findOrder_(t, id);
  if (!hit) return { ok: false, msg: '找不到订单' };
  return withLock_(function () {
    return patchRow_(t.sheet, t.head, hit.__row, hit.ORDER_ID, { STATUS: 'ACTIVE' });
  });
}

function updateOrder(p) {
  var t = readTable_('ORDERS');
  var hit = findOrder_(t, p.id);
  if (!hit) return { ok: false, msg: '找不到订单' };
  var qty = toNum_(p.qty), price = toNum_(p.price);
  if (qty <= 0) return { ok: false, msg: '数量必须大于 0' };
  if (price <= 0) return { ok: false, msg: '售价必须大于 0' };
  var setType = p.setType || String(hit.SET_TYPE);

  var total = (p.totalOverride !== '' && p.totalOverride != null)
    ? toNum_(p.totalOverride) : Math.round(qty * price * 100) / 100;
  var pps = profitFor_(setType, price);
  if (pps === null) {
    var oq = toNum_(hit.QTY);
    pps = oq > 0 ? toNum_(hit.MY_INCOME) / oq : 0;
  }
  var mine = (p.myIncomeOverride !== '' && p.myIncomeOverride != null)
    ? toNum_(p.myIncomeOverride) : Math.round(pps * qty * 100) / 100;
  var fee = (p.driverFeeOverride !== '' && p.driverFeeOverride != null)
    ? toNum_(p.driverFeeOverride)
    : calcDriverFee_(hit.REGION, hit.STATE, setType, qty, price);

  var patch = {
    SET_TYPE: setType, UNIT_PRICE: price, QTY: qty,
    TOTAL_INCOME: total, MY_INCOME: mine, DRIVER_FEE: fee
  };
  if (p.payStatus) patch.PAY_STATUS = p.payStatus;
  if (p.note != null) patch.NOTE = p.note;

  var r = withLock_(function () {
    return patchRow_(t.sheet, t.head, hit.__row, hit.ORDER_ID, patch);
  });
  if (!r.ok) return r;
  return { ok: true, total: total, mine: mine, fee: fee };
}

function markPaid(p) {
  // 读、找、验证 —— 全部在锁外面做
  var t = ensureCols_('ORDERS', ['PAY_METHOD', 'PAY_NOTE']);
  var hit = findOrder_(t, p && p.id);
  if (!hit) return { ok: false, msg: '找不到订单' };
  var method = up_(p && p.method);
  if (method !== 'OP' && method !== 'PC') return { ok: false, msg: '请选收款方式（OP 转账 或 PC 现金）' };
  var d = String((p && p.date) || '').trim() || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return { ok: false, msg: '日期格式不对' };
  var note = String((p && p.note) == null ? '' : p.note).trim();

  return withLock_(function () {
    return patchRow_(t.sheet, t.head, hit.__row, hit.ORDER_ID, {
      PAY_STATUS: 'PAID', PAY_METHOD: method, PAY_DATE: d, PAY_NOTE: note
    });
  });
}

function markUnpaid(id) {
  var t = ensureCols_('ORDERS', ['PAY_METHOD', 'PAY_NOTE']);
  var hit = findOrder_(t, id);
  if (!hit) return { ok: false, msg: '找不到订单' };
  return withLock_(function () {
    return patchRow_(t.sheet, t.head, hit.__row, hit.ORDER_ID, {
      PAY_STATUS: 'UNPAID', PAY_METHOD: '', PAY_DATE: '', PAY_NOTE: ''
    });
  });
}

/** 未收款清单：依分行分组，附拖欠天数 */
function getUnpaid(opt) {
  opt = opt || {};
  var t = readTable_('ORDERS');
  var today = new Date();
  var groups = {}, tot = { n: 0, amount: 0 };
  t.rows.forEach(function (r) {
    if (isVoid_(r) || payPaid_(r)) return;
    var inc = toNum_(r.TOTAL_INCOME);
    var b = String(r.BRANCH || '-');
    var ds = fmtDate_(r.DATE);
    var age = 0;
    var m = ds.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) age = Math.max(0, Math.round((today - new Date(+m[1], +m[2] - 1, +m[3])) / 86400000));
    groups[b] = groups[b] || { branch: b, n: 0, amount: 0, oldest: 0, items: [] };
    groups[b].n++; groups[b].amount += inc;
    groups[b].oldest = Math.max(groups[b].oldest, age);
    groups[b].items.push({
      id: String(r.ORDER_ID), date: ds, salesman: String(r.SALESMAN),
      setType: String(r.SET_TYPE), qty: toNum_(r.QTY), total: inc, age: age, month: toNum_(r.MONTH)
    });
    tot.n++; tot.amount += inc;
  });
  var list = Object.keys(groups).map(function (k) { return groups[k]; })
    .sort(function (a, b) { return b.amount - a.amount; });
  list.forEach(function (g) { g.items.sort(function (a, b) { return b.age - a.age; }); });
  return { total: tot, branches: list };
}

/* ---------- 司机月结单 ----------
   照客户原本手写表的格式出：
   · 上面地区（PER_SET）：BIL / DATE / BRANCH / SALESMAN / SET / 抽成 / SET&DETAIL
   · 下面地区（PERCENT）：多一栏 PRICE，因为抽成是 售价 × 10%
   刻意不含「我的利润」，司机看得到也没关系。
--------------------------------- */
function statMonthName_(m) {
  return ['', 'JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY',
    'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'][m] || ('MONTH ' + m);
}
/** 2026-06-13 → 13.06.26 */
function shortDate_(iso) {
  var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? (m[3] + '.' + m[2] + '.' + m[1].slice(2)) : String(iso || '');
}
/** 那一区是按 % 抽还是按 set 抽 */
function ruleTypeOf_(rules, region, state) {
  var R = up_(region), S = up_(state), rule = null;
  rules.forEach(function (r) { if (up_(r.REGION) === R && up_(r.STATE) === S) rule = r; });
  if (!rule) rules.forEach(function (r) { if (!rule && up_(r.REGION) === R) rule = r; });
  return rule ? up_(rule.RULE_TYPE) : 'NONE';
}

function getStatement(opt) {
  opt = opt || {};
  var month = toNum_(opt.month);
  if (!(month >= 1 && month <= 12)) return { ok: false, msg: '请选择月份' };

  var rules = readTable_('DRIVER_RULE').rows;
  var t = readTable_('ORDERS');
  var dv = readTable_('DRIVER').rows[0] || {};
  var secs = {}, cash = [], cashTotal = 0, feeTotal = 0, setTotal = 0, n = 0;

  t.rows.forEach(function (r) {
    if (isVoid_(r)) return;
    if (toNum_(r.MONTH) !== month) return;
    var fee = toNum_(r.DRIVER_FEE);
    var qty = toNum_(r.QTY), price = toNum_(r.UNIT_PRICE);

    // 司机代收的现金（PC）→ 结帐时要扣掉。
    // 只算司机自己送的单（fee > 0）；南马是老板亲自送，钱没有经过司机的手。
    if (fee && payMethod_(r) === 'PC') {
      var amt = toNum_(r.TOTAL_INCOME);
      cash.push({ branch: String(r.BRANCH || ''), salesman: String(r.SALESMAN || ''), amount: amt, date: fmtDate_(r.DATE) });
      cashTotal += amt;
    }
    if (!fee) return;   // 不算抽成的（南马自送）不上月结单

    var type = ruleTypeOf_(rules, r.REGION, r.STATE);
    var key = type === 'PERCENT' ? 'PERCENT' : 'PER_SET';
    secs[key] = secs[key] || { type: key, rows: [], fee: 0, sets: 0 };
    secs[key].rows.push({
      date: shortDate_(fmtDate_(r.DATE)), iso: fmtDate_(r.DATE),
      branch: String(r.BRANCH || ''), salesman: String(r.SALESMAN || ''),
      qty: qty, price: price, fee: fee, setType: String(r.SET_TYPE || '')
    });
    secs[key].fee += fee; secs[key].sets += qty;
    feeTotal += fee; setTotal += qty; n++;
  });

  var order = ['PER_SET', 'PERCENT'];
  var sections = [];
  order.forEach(function (k) {
    if (!secs[k]) return;
    var s = secs[k];
    s.rows.sort(function (a, b) { return a.iso < b.iso ? -1 : a.iso > b.iso ? 1 : 0; });
    s.rows.forEach(function (r, i) { r.bil = i + 1; });
    s.fee = Math.round(s.fee * 100) / 100;
    s.title = k === 'PER_SET' ? '上面地区（每 set 计）' : '下面地区（售价 × %）';
    s.showPrice = (k === 'PERCENT');
    sections.push(s);
  });

  var allowance = toNum_(dv.ALLOWANCE_PER_MONTH);
  feeTotal = Math.round(feeTotal * 100) / 100;
  cashTotal = Math.round(cashTotal * 100) / 100;
  var subtotal = Math.round((feeTotal + allowance) * 100) / 100;

  return {
    ok: true, month: month, monthName: statMonthName_(month),
    driver: String(dv.DRIVER_NAME || ''), company: config_().COMPANY_NAME || config_().COMPANY || 'V&P TRADING',
    sections: sections, orders: n, sets: setTotal,
    fee: feeTotal, allowance: allowance, subtotal: subtotal,
    cash: cash, cashTotal: cashTotal,
    payable: Math.round((subtotal - cashTotal) * 100) / 100,
    generated: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Kuala_Lumpur', 'yyyy-MM-dd HH:mm')
  };
}

/* ---------- 月报（给持有人自己存一份 Excel） ----------
   栏位排法照客户原本那份月份分页，后面才补上收款方式 / 日期 / 备注。
   含「我的利润」，所以这份档案不要转发给司机。
------------------------------------------------------- */
function getMonthlyReport(opt) {
  opt = opt || {};
  var month = toNum_(opt.month);
  if (!(month >= 1 && month <= 12)) return { ok: false, msg: '请选择月份' };

  var t = readTable_('ORDERS');
  var dv = readTable_('DRIVER').rows[0] || {};
  var rules = readTable_('DRIVER_RULE').rows;

  var rows = [], byBranch = {}, bySet = {}, byMonth = {};
  var blank_ = function () {
    return { orders: 0, sets: 0, income: 0, mine: 0, fee: 0,
             paid: 0, paidN: 0, unpaid: 0, unpaidN: 0, op: 0, pc: 0, cash: 0 };
  };
  var add_ = function (o, v) {
    o.orders++; o.sets += v.qty; o.income += v.inc; o.mine += v.mine; o.fee += v.fee;
    if (v.paid) {
      o.paid += v.inc; o.paidN++;
      if (v.method === 'OP') o.op += v.inc;
      if (v.method === 'PC') { o.pc += v.inc; if (v.fee) o.cash += v.inc; }
    } else { o.unpaid += v.inc; o.unpaidN++; }
  };
  var tot = blank_(), ytd = blank_();

  t.rows.forEach(function (r) {
    if (isVoid_(r)) return;
    var m = toNum_(r.MONTH);
    if (m > month) return;                      // 只算到这个月为止

    var qty = toNum_(r.QTY), price = toNum_(r.UNIT_PRICE),
        inc = toNum_(r.TOTAL_INCOME), mine = toNum_(r.MY_INCOME), fee = toNum_(r.DRIVER_FEE);
    var paid = payPaid_(r), method = payMethod_(r);
    var v = { qty: qty, inc: inc, mine: mine, fee: fee, paid: paid, method: method };

    add_(ytd, v);                                // 1 月到本月的累计
    byMonth[m] = byMonth[m] || blank_();
    add_(byMonth[m], v);

    if (m !== month) return;                     // 以下只做本月的明细

    var b = String(r.BRANCH || '-'), st = String(r.SET_TYPE || '-');
    rows.push({
      date: fmtDate_(r.DATE), state: String(r.STATE || ''), branch: b,
      salesman: String(r.SALESMAN || ''), sets: qty, price: price,
      income: inc, mine: mine, setDetail: st, fee: fee,
      method: method, payDate: String(r.PAY_DATE || ''),
      payNote: String(r.PAY_NOTE || ''), paid: paid,
      delivery: String(r.DELIVERY_STATUS || ''), note: String(r.NOTE || '')
    });
    add_(tot, v);

    byBranch[b] = byBranch[b] || { name: b, orders: 0, sets: 0, income: 0, mine: 0, fee: 0, unpaid: 0 };
    byBranch[b].orders++; byBranch[b].sets += qty; byBranch[b].income += inc;
    byBranch[b].mine += mine; byBranch[b].fee += fee; if (!paid) byBranch[b].unpaid += inc;

    bySet[st] = bySet[st] || { name: st, sets: 0, income: 0, mine: 0 };
    bySet[st].sets += qty; bySet[st].income += inc; bySet[st].mine += mine;
  });

  if (!tot.orders) return { ok: false, msg: statMonthName_(month) + '没有订单' };

  rows.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  rows.forEach(function (r, i) { r.bil = i + 1; });

  var K = function (o) { return Object.keys(o).map(function (k) { return o[k]; }); };
  var r2 = function (n) { return Math.round(n * 100) / 100; };
  var MONEY = ['income', 'mine', 'fee', 'paid', 'unpaid', 'op', 'pc', 'cash'];
  var round_ = function (o) { MONEY.forEach(function (k) { o[k] = r2(o[k]); }); return o; };
  round_(tot); round_(ytd);

  // 每月走势（含逐月累计）
  var run = 0, runMine = 0, runFee = 0;
  var months = [];
  for (var mi = 1; mi <= month; mi++) {
    var b2 = byMonth[mi]; if (!b2) continue;
    round_(b2);
    run += b2.income; runMine += b2.mine; runFee += b2.fee;
    months.push({
      m: mi, name: statMonthName_(mi), orders: b2.orders, sets: b2.sets,
      income: b2.income, mine: b2.mine, fee: b2.fee,
      paid: b2.paid, unpaid: b2.unpaid,
      cumIncome: r2(run), cumMine: r2(runMine), cumFee: r2(runFee)
    });
  }

  var allowance = toNum_(dv.ALLOWANCE_PER_MONTH);
  var nMonths = months.length || 1;
  return {
    ok: true, month: month, monthName: statMonthName_(month),
    company: config_().COMPANY || 'V&P TRADING',
    driver: String(dv.DRIVER_NAME || ''),
    rows: rows, total: tot,
    ytd: ytd, months: months, nMonths: nMonths,
    ytdDriverPayable: r2(ytd.fee + allowance * nMonths - ytd.cash),
    allowance: allowance,
    driverPayable: r2(tot.fee + allowance - tot.cash),
    branches: K(byBranch).sort(function (a, b) { return b.income - a.income; }),
    sets: K(bySet).sort(function (a, b) { return b.income - a.income; }),
    statement: getStatement({ month: month }),
    generated: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Kuala_Lumpur', 'yyyy-MM-dd HH:mm')
  };
}

/* ---------- 报表 ---------- */
function getDashboard(opt) {
  opt = opt || {};
  var t = readTable_('ORDERS');
  var drv = readTable_('DRIVER').rows[0] || {};
  var allowance = toNum_(drv.ALLOWANCE_PER_MONTH);

  var byMonth = {}, byBranch = {}, byRegion = {}, bySet = {};
  var tot = { orders: 0, sets: 0, income: 0, mine: 0, fee: 0, unpaid: 0, unpaidN: 0 };

  t.rows.forEach(function (r) {
    if (isVoid_(r)) return;
    if (opt.month && toNum_(r.MONTH) !== toNum_(opt.month)) return;
    if (opt.region && up_(r.REGION) !== up_(opt.region)) return;
    var inc = toNum_(r.TOTAL_INCOME), mine = toNum_(r.MY_INCOME), fee = toNum_(r.DRIVER_FEE), q = toNum_(r.QTY);
    tot.orders++; tot.sets += q; tot.income += inc; tot.mine += mine; tot.fee += fee;
    if (!payPaid_(r)) { tot.unpaid += inc; tot.unpaidN = (tot.unpaidN || 0) + 1; }

    var m = toNum_(r.MONTH);
    byMonth[m] = byMonth[m] || { m: m, income: 0, mine: 0, fee: 0, orders: 0 };
    byMonth[m].income += inc; byMonth[m].mine += mine; byMonth[m].fee += fee; byMonth[m].orders++;

    var b = String(r.BRANCH);
    byBranch[b] = byBranch[b] || { name: b, income: 0, mine: 0, orders: 0 };
    byBranch[b].income += inc; byBranch[b].mine += mine; byBranch[b].orders++;

    var rg = String(r.REGION || '-');
    byRegion[rg] = byRegion[rg] || { name: rg, income: 0, mine: 0, fee: 0 };
    byRegion[rg].income += inc; byRegion[rg].mine += mine; byRegion[rg].fee += fee;

    var st = String(r.SET_TYPE || '-');
    bySet[st] = bySet[st] || { name: st, sets: 0, income: 0 };
    bySet[st].sets += q; bySet[st].income += inc;
  });

  var months = Object.keys(byMonth).map(function (k) { return byMonth[k]; }).sort(function (a, b) { return a.m - b.m; });
  var branches = Object.keys(byBranch).map(function (k) { return byBranch[k]; })
    .sort(function (a, b) { return b.income - a.income; }).slice(0, 12);
  var sets = Object.keys(bySet).map(function (k) { return bySet[k]; })
    .sort(function (a, b) { return b.income - a.income; });
  var regions = Object.keys(byRegion).map(function (k) { return byRegion[k]; });

  var nMonths = months.length || 1;
  return {
    total: tot, months: months, branches: branches, sets: sets, regions: regions,
    driverPayable: Math.round((tot.fee + allowance * (opt.month ? 1 : nMonths)) * 100) / 100,
    allowance: allowance, nMonths: nMonths
  };
}

/* ---------- 订单清单 ---------- */
function getOrders(opt) {
  opt = opt || {};
  var t = readTable_('ORDERS');
  var q = up_(opt.q || '');
  var out = [];
  for (var i = t.rows.length - 1; i >= 0; i--) {
    var r = t.rows[i];
    if (isVoid_(r) && !opt.includeVoid) continue;
    if (opt.month && toNum_(r.MONTH) !== toNum_(opt.month)) continue;
    if (opt.pending && up_(r.DELIVERY_STATUS) === 'DELIVERED') continue;
    if (opt.pay === 'UNPAID' && payPaid_(r)) continue;
    if (opt.pay === 'PAID' && !payPaid_(r)) continue;
    if (q && (up_(r.SALESMAN).indexOf(q) < 0 && up_(r.BRANCH).indexOf(q) < 0 && up_(r.ORDER_ID).indexOf(q) < 0)) continue;
    out.push({
      id: String(r.ORDER_ID), date: fmtDate_(r.DATE), month: toNum_(r.MONTH),
      branch: String(r.BRANCH), salesman: String(r.SALESMAN), setType: String(r.SET_TYPE),
      price: toNum_(r.UNIT_PRICE), qty: toNum_(r.QTY), total: toNum_(r.TOTAL_INCOME),
      mine: toNum_(r.MY_INCOME), fee: toNum_(r.DRIVER_FEE),
      paid: payPaid_(r), method: payMethod_(r), payDate: fmtDate_(r.PAY_DATE),
      payNote: String(r.PAY_NOTE || ''),
      delivery: String(r.DELIVERY_STATUS || ''), region: String(r.REGION || ''),
      state: String(r.STATE || ''), note: String(r.NOTE || ''), status: up_(r.STATUS) || 'ACTIVE',
      drvColor: String(r.DRV_COLOR || ''), drvGrade: String(r.DRV_GRADE || ''),
      drvNote: String(r.DRV_NOTE || '')
    });
    if (out.length >= (opt.limit || 300)) break;
  }
  return out;
}

function fmtDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v);
}

/* ---------- 新客户 / 新分行 ---------- */
/* ---------- 改分行 ----------
   分行换了，州属 / 区域 / 品牌 / 司机抽成都要跟着换。
   例：PERODUA KAJANG PRIMA（雪州，RM2.50/set）→ PERODUA RASAH JAYA（森州，10%）
       5 set × RM 35：抽成从 RM 12.50 变成 RM 17.50。
   所以一定要先给使用者看清楚再存。
------------------------------- */
function branchInfo_(name) {
  var hit = null;
  readTable_('BRANCH').rows.forEach(function (r) {
    if (up_(r.BRANCH) === up_(name)) hit = r;
  });
  if (!hit) return null;
  return {
    branch: String(hit.BRANCH).trim(), brand: String(hit.BRAND || '').trim(),
    state: String(hit.STATE || '').trim(), region: String(hit.REGION || '').trim()
  };
}

/** 先算给使用者看：改了之后抽成会变多少 */
function previewBranch(p) {
  var t = readTable_('ORDERS');
  var hit = findOrder_(t, p && p.id);
  if (!hit) return { ok: false, msg: '找不到订单' };
  var to = branchInfo_(p && p.branch);
  if (!to) return { ok: false, msg: '名单里没有这间分行：' + (p && p.branch) };

  var qty = toNum_(hit.QTY), price = toNum_(hit.UNIT_PRICE), st = String(hit.SET_TYPE);
  var newFee = calcDriverFee_(to.region, to.state, st, qty, price);
  var oldFee = toNum_(hit.DRIVER_FEE);
  return {
    ok: true,
    salesman: String(hit.SALESMAN || ''),
    from: {
      branch: String(hit.BRANCH || ''), state: String(hit.STATE || ''),
      region: String(hit.REGION || ''), fee: oldFee
    },
    to: { branch: to.branch, state: to.state, region: to.region, brand: to.brand, fee: newFee },
    feeChanged: Math.abs(newFee - oldFee) > 0.005,
    diff: Math.round((newFee - oldFee) * 100) / 100
  };
}

/** 真的改。alsoSalesman = true 时，顺便把这个销售员的主分行也换过去 */
function changeBranch(p) {
  var t = readTable_('ORDERS');
  var hit = findOrder_(t, p && p.id);
  if (!hit) return { ok: false, msg: '找不到订单' };
  var to = branchInfo_(p && p.branch);
  if (!to) return { ok: false, msg: '名单里没有这间分行：' + (p && p.branch) };
  if (up_(hit.BRANCH) === up_(to.branch)) return { ok: false, msg: '本来就是这间分行' };

  var qty = toNum_(hit.QTY), price = toNum_(hit.UNIT_PRICE), st = String(hit.SET_TYPE);
  var newFee = calcDriverFee_(to.region, to.state, st, qty, price);
  var oldFee = toNum_(hit.DRIVER_FEE);
  var salesman = up_(hit.SALESMAN);

  var r = withLock_(function () {
    return patchRow_(t.sheet, t.head, hit.__row, hit.ORDER_ID, {
      BRANCH: to.branch, BRAND: to.brand, STATE: to.state,
      REGION: to.region, DRIVER_FEE: newFee
    });
  });
  if (!r.ok) return r;

  var moved = false;
  if (p && p.alsoSalesman && salesman) moved = moveSalesman_(salesman, to);

  clearBootCache_();
  return {
    ok: true, oldFee: oldFee, newFee: newFee,
    diff: Math.round((newFee - oldFee) * 100) / 100,
    salesman: salesman, salesmanMoved: moved
  };
}

/** 把销售员的「主要分行」★ 移到新分行；名单里没有就补一行 */
function moveSalesman_(salesman, to) {
  try {
    var t = readTable_('SALESMAN');
    var star = t.head.indexOf('主要分行') + 1;
    if (star <= 0) return false;
    var mine = t.rows.filter(function (r) { return up_(r.SALESMAN) === up_(salesman); });
    var target = null;
    mine.forEach(function (r) { if (up_(r.BRANCH) === up_(to.branch)) target = r; });

    return withLock_(function () {
      mine.forEach(function (r) {
        if (String(r['主要分行'] || '').trim()) t.sheet.getRange(r.__row, star).setValue('');
      });
      if (target) {
        t.sheet.getRange(target.__row, star).setValue('★');
      } else {
        t.sheet.appendRow(t.head.map(function (h) {
          return h === 'SALESMAN' ? up_(salesman) : h === 'BRANCH' ? to.branch
            : h === 'BRAND' ? to.brand : h === 'STATE' ? to.state : h === 'REGION' ? to.region
              : h === '历史笔数' ? 0 : h === '主要分行' ? '★' : '';
        }));
      }
      return { ok: true };
    }).ok === true;
  } catch (e) { return false; }
}

function addSalesman(p) {
  try {
    var bt = readTable_('BRANCH');
    var exists = bt.rows.some(function (r) { return up_(r.BRANCH) === up_(p.branch); });
    if (!exists) {
      var bl = bt.head.map(function (h) {
        return h === 'BRANCH' ? String(p.branch).trim() : h === 'BRAND' ? p.brand
          : h === 'STATE' ? p.state : h === 'REGION' ? p.region
            : h === 'ACTIVE' ? 'YES' : h === '历史笔数' ? 0 : '';
      });
      bt.sheet.appendRow(bl);
    }
    var st = readTable_('SALESMAN');
    var dup = st.rows.some(function (r) {
      return up_(r.SALESMAN) === up_(p.salesman) && up_(r.BRANCH) === up_(p.branch);
    });
    if (!dup) {
      var sl = st.head.map(function (h) {
        return h === 'SALESMAN' ? up_(p.salesman) : h === 'BRANCH' ? String(p.branch).trim()
          : h === 'BRAND' ? p.brand : h === 'STATE' ? p.state : h === 'REGION' ? p.region
            : h === '历史笔数' ? 0 : h === '主要分行' ? '★' : '';
      });
      st.sheet.appendRow(sl);
    }
    clearBootCache_();
    return { ok: true };
  } catch (e) { return { ok: false, msg: String(e.message || e) }; }
}

function addSetPrice(p) {
  try {
    var t = readTable_('SET_PRICE');
    var line = t.head.map(function (h) {
      return h === 'SET_TYPE' ? String(p.setType).trim().toUpperCase()
        : h === 'UNIT_PRICE' ? toNum_(p.price)
          : h === 'PROFIT_PER_SET' ? toNum_(p.profit)
            : h === 'ACTIVE' ? 'YES' : h === '历史笔数' ? 0 : '';
    });
    t.sheet.appendRow(line);
    clearBootCache_();
    return { ok: true };
  } catch (e) { return { ok: false, msg: String(e.message || e) }; }
}

/*******************************************************
 * V&P TRADING — 一次性安装器
 * 直接读取现有的 January…December 分页，
 * 自动建立 ORDERS / SET_PRICE / BRANCH / SALESMAN /
 * DRIVER_RULE / DRIVER / USERS / CONFIG。
 *
 * 用法：在编辑器上方函式下拉选 INSTALL，按 Run。
 * 旧分页完全不动。可重复执行（会覆盖上述 8 张表）。
 *******************************************************/

var MONTH_SHEETS = [
  ['January', 1], ['February', 2], ['March', 3], ['April', 4],
  ['May', 5], ['Jun', 6], ['July', 7], ['August', 8],
  ['September', 9], ['October', 10], ['November', 11], ['December', 12]
];

/* 确定的拼写错误 → 合并 */
var BRANCH_FIX = {
  'HONDA NILAI IMPISN': 'HONDA NILAI IMPIAN',
  'MITSUMITSHI SKUDAI': 'MITSUBISHI SKUDAI',
  'MITUBUSHI BALAKONG': 'MITSUBISHI BALAKONG',
  'PERODAU KOTA MASAI': 'PERODUA KOTA MASAI',
  'PROOTN KESANG': 'PROTON KESANG',
  'JEACOO TMN DAYA': 'JAECOO TMN DAYA',
  'JEACOO TMN GAYA': 'JAECOO TMN DAYA',
  'I CAUR PELANGGIT': 'ICAUR PELANGGIT',
  'PERODUA SEREMBAN 2': 'PERODUA SEREMBAN2'
};

var BRAND_LIST = ['PERODUA', 'PROTON', 'HONDA', 'CHERY', 'JAECOO', 'JETOUR', 'MITSUBISHI', 'EMAS', 'ICAUR'];

function INSTALL() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var t0 = new Date();
  var rows = readHistory_(ss);
  if (!rows.length) throw new Error('读不到任何历史订单，请确认分页名字（January / February / …）');

  var setPrice = derivePrices_(rows);
  var lookup = {};
  setPrice.forEach(function (s) { lookup[s[0] + '|' + s[1]] = s[2]; });

  rows.forEach(function (r) { r.fee = driverFee_(r.region, r.state, r.detail, r.qty, r.price); });

  var keptN = writeOrders_(ss, rows);
  writeSheet_(ss, 'SET_PRICE', ['SET_TYPE', 'UNIT_PRICE', 'PROFIT_PER_SET', '说明', '历史笔数', 'ACTIVE'],
    setPrice.map(function (r) { return [r[0], r[1], r[2], SET_DESC_[up_(r[0]) + '|' + r[1]] || '', r[3], r[4]]; }),
    '● 价目表（可改）：App 依 SET_TYPE + UNIT_PRICE 查出每 set 利润。「说明」会印在给司机的送货单上（例：普通雨伞 / 好的雨伞）。改这里 App 立刻生效。');
  writeSheet_(ss, 'BRANCH', ['BRANCH', 'BRAND', 'STATE', 'REGION', '历史笔数', 'ACTIVE'], buildBranch_(rows),
    '● Branch 名单（已统一写法）。新增分行直接加一行。');
  writeSheet_(ss, 'SALESMAN', ['SALESMAN', 'BRANCH', 'BRAND', 'STATE', 'REGION', '历史笔数', '最后下单月', '主要分行'],
    buildSalesman_(rows), '● 销售员 → 分行。同名在多间分行时 App 会让你选（★ 为最常下单那间）。');

  writeSheet_(ss, 'DRIVER_RULE', ['REGION', 'STATE', 'RULE_TYPE', 'RATE', 'UMBRELLA_RATE', '说明'], [
    ['CENTRAL', 'Selangor', 'PER_SET', 2.5, 1.0, '每 set RM2.50；雨伞类每 set RM1.00'],
    ['CENTRAL', 'NS', 'PERCENT', 0.10, 0.10, '数量 × 售价 × 10%'],
    ['SOUTH', 'Johor', 'NONE', 0, 0, '老板亲自送货，无抽成'],
    ['SOUTH', 'Melaka', 'NONE', 0, 0, '老板亲自送货，无抽成']
  ], '● 司机抽成规则（可改）。已用 1–7 月 Uncle 表 1000+ 笔验证。');

  writeSheetKeep_(ss, 'DRIVER', ['DRIVER_NAME', 'PHONE', 'ALLOWANCE_PER_MONTH', 'ACTIVE'],
    [['TAN', '60122356648', 200, 'YES']],
    '● 司机资料。PHONE 用国际格式（60 开头，不要 +、空格或 -）。');

  writeSheetKeep_(ss, 'USERS', ['NAME', 'PIN', 'ROLE', 'ACTIVE', '说明'], [
    ['ADMIN', '1234', 'admin', 'YES', '看全部：下单、订单、报表、利润、司机抽成'],
    ['PARTNER1', '2345', 'partner', 'YES', '下单、订单、报表（看不到利润与司机抽成）'],
    ['TAN', '9999', 'driver', 'YES', '司机：只看送货单与自己的抽成，看不到售价与利润']
  ], '● 登入 PIN —— 请自行改掉！ROLE 只能填 admin / partner / driver。');

  writeSheetKeep_(ss, 'CONFIG', ['KEY', 'VALUE', '说明'], [
    ['COMPANY', 'V&P TRADING', '公司名，显示在 App 顶部'],
    ['CURRENCY', 'RM', ''],
    ['DEFAULT_PAY_STATUS', 'OP', ''],
    ['PAY_STATUS_OPTIONS', 'OP,PC,PAID,PENDING', '下单时付款状态选项（逗号分隔）'],
    ['DRIVER_PHONE', '', '（已停用，改由 DRIVER 分页控制）']
  ], '● 系统设定。');

  try { CacheService.getScriptCache().remove('boot'); } catch (e) { }
  var sec = Math.round((new Date() - t0) / 1000);
  var msg = '安装完成 ✓\n\n历史订单：' + rows.length + ' 笔\n'
    + (keptN ? 'App 新增单（已保留）：' + keptN + ' 笔\n' : '')
    + '价目表：' + setPrice.length + ' 组\n'
    + 'Branch：' + buildBranch_(rows).length + ' 间\n销售员：' + buildSalesman_(rows).length + ' 组\n\n耗时 ' + sec + ' 秒';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { }
  return msg;
}

/* ---------- 读历史 ---------- */
function readHistory_(ss) {
  var out = [];
  MONTH_SHEETS.forEach(function (m) {
    var sheet = ss.getSheetByName(m[0]);
    if (!sheet) return;
    var v = sheet.getDataRange().getValues();
    var region = 'CENTRAL';
    for (var i = 0; i < v.length; i++) {
      var r = v[i];
      var filled = r.filter(function (x) { return x !== '' && x !== null; });
      var j = r.join(' ').toUpperCase();
      if (filled.length <= 2 && j.indexOf('MELAKA') >= 0 && j.indexOf('JOHOR') >= 0) { region = 'SOUTH'; continue; }
      if (filled.length <= 2 && j.indexOf('SEREMBAN') >= 0 && j.indexOf('KL') >= 0) { region = 'CENTRAL'; continue; }
      if (String(r[0]).trim() === 'BIL') continue;
      if (!(r[3] && r[4] && typeof r[5] === 'number')) continue;
      var bn = normBranch_(r[3]);
      out.push({
        month: m[1], date: fmtD_(r[1], m[1]), region: region,
        state: String(r[2] || '').trim(), branch: bn.name, note: bn.note, brand: brandOf_(bn.name),
        salesman: String(r[4]).trim().toUpperCase().replace(/\s+/g, ' '),
        qty: n_(r[5]), price: n_(r[6]), total: n_(r[7]), mine: n_(r[8]),
        detail: normDetail_(r[9]),
        pay: String(r[10] || 'OP').trim().toUpperCase(),
        payDate: r[11] ? String(r[11]).trim() : '',
        src: '历史·' + m[0]
      });
    }
  });
  return out;
}

function n_(x) { var v = parseFloat(x); return isNaN(v) ? 0 : Math.round(v * 100) / 100; }

function fmtD_(v, month) {
  if (Object.prototype.toString.call(v) === '[object Date]')
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{2,4})/);
  if (m) {
    var y = m[3].length === 2 ? '20' + m[3] : m[3];
    return y + '-' + pad2_(m[2]) + '-' + pad2_(m[1]);
  }
  return '2026-' + pad2_(month) + '-01';
}
function pad2_(x) { x = String(parseInt(x, 10)); return x.length < 2 ? '0' + x : x; }

function normBranch_(b) {
  var s = String(b).trim().toUpperCase().replace(/\s+/g, ' ');
  s = s.replace(/MITSUMISHI|MITSHUBISHI|MITSUBHISI/g, 'MITSUBISHI').replace(/\bMIT\b/g, 'MITSUBISHI');
  s = s.replace(/PERORDUA|PEODUA|PERDUA/g, 'PERODUA').replace(/PROTOB|PORTON/g, 'PROTON');
  s = s.replace(/\bE\s*-?\s*MAS\b/g, 'EMAS').replace(/EMAS PROTON/g, 'EMAS').replace(/IMPIAM/g, 'IMPIAN');
  s = s.replace(/\s*\(\s*/g, '(').replace(/\s*\)\s*/g, ')').trim();
  if (BRANCH_FIX[s]) return { name: BRANCH_FIX[s], note: '' };
  var p = s.match(/^(.*?)\((.*)\)$/);
  if (p) {
    var base = p[1].trim();
    return { name: BRANCH_FIX[base] || base, note: p[2].trim() };
  }
  return { name: s, note: '' };
}

function brandOf_(b) {
  for (var i = 0; i < BRAND_LIST.length; i++) if (b.indexOf(BRAND_LIST[i]) === 0) return BRAND_LIST[i];
  return 'OTHER';
}

function normDetail_(d) {
  if (!d) return '';
  var s = String(d).trim().toUpperCase().replace(/\s+/g, ' ');
  s = s.replace(/TEMS/g, 'ITEMS').replace(/ITEMA/g, 'ITEMS').replace(/IITEMS/g, 'ITEMS').replace(/ITEMSS/g, 'ITEMS');
  var m = s.match(/^(\d+)\s*ITEMS?\b/);
  if (m) {
    var rest = s.slice(m[0].length).replace(/^[\s-]+/, '');
    return m[1] + ' ITEMS' + (rest ? ' - ' + rest : '');
  }
  if (s.indexOf('UMBRELLA') >= 0 || s.indexOf('UMBRELA') >= 0) {
    if (s.indexOf('BAG') >= 0) return 'UMBRELLA + BAG';
    if (s.indexOf('PREMIUM') >= 0) return 'PREMIUM UMBRELLA';
    if (s.indexOf('BLACK') >= 0) return 'BLACK UMBRELLA';
    if (s.indexOf('RED') >= 0) return 'RED UMBRELLA';
    if (s.indexOf('ICARS') >= 0) return 'ICARS UMBRELLA';
    return 'UMBRELLA';
  }
  if (s === 'BAG' || s === 'BAGS') return 'BAG';
  return s;
}

/* ---------- 司机抽成 ---------- */
function driverFee_(region, state, detail, qty, price) {
  if (String(region).toUpperCase() === 'SOUTH') return 0;
  var s = String(state).toUpperCase();
  if (s === 'NS' || s === 'N9' || s.indexOf('NEGERI') >= 0) return Math.round(qty * price * 0.10 * 100) / 100;
  var umb = String(detail).toUpperCase().indexOf('UMBRELLA') >= 0;
  return Math.round(qty * (umb ? 1.0 : 2.5) * 100) / 100;
}

/* ---------- 从历史推导价目表 ---------- */
function derivePrices_(rows) {
  var agg = {};
  rows.forEach(function (r) {
    if (!r.qty || !r.price || !r.mine) return;
    var k = r.detail + '|' + r.price;
    var pps = Math.round(r.mine / r.qty * 100) / 100;
    agg[k] = agg[k] || {};
    agg[k][pps] = (agg[k][pps] || 0) + 1;
  });
  var out = [];
  Object.keys(agg).forEach(function (k) {
    var c = agg[k], total = 0, best = null, bn = 0;
    Object.keys(c).forEach(function (p) { total += c[p]; if (c[p] > bn) { bn = c[p]; best = parseFloat(p); } });
    if (total >= 3 && bn / total >= 0.6) {
      var parts = k.split('|');
      out.push([parts[0], parseFloat(parts[1]), best, total, 'YES']);
    }
  });
  out.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] - b[1]; });
  return out;
}

/* ---------- 建资料表 ---------- */
function buildBranch_(rows) {
  var m = {};
  rows.forEach(function (r) {
    m[r.branch] = m[r.branch] || { brand: r.brand, state: r.state, region: r.region, n: 0 };
    m[r.branch].n++;
  });
  return Object.keys(m).sort().map(function (b) {
    return [b, m[b].brand, m[b].state, m[b].region, m[b].n, 'YES'];
  });
}

function buildSalesman_(rows) {
  var m = {};
  rows.forEach(function (r) {
    var k = r.salesman + '|' + r.branch;
    m[k] = m[k] || { s: r.salesman, b: r.branch, brand: r.brand, state: r.state, region: r.region, n: 0, last: 0 };
    m[k].n++; m[k].last = Math.max(m[k].last, r.month);
  });
  var byName = {};
  Object.keys(m).forEach(function (k) {
    var o = m[k];
    byName[o.s] = byName[o.s] || [];
    byName[o.s].push(o);
  });
  var out = [];
  Object.keys(byName).sort().forEach(function (name) {
    var list = byName[name].sort(function (a, b) { return b.n - a.n; });
    list.forEach(function (o, i) {
      out.push([o.s, o.b, o.brand, o.state, o.region, o.n, pad2_(o.last), i === 0 ? '★' : '']);
    });
  });
  return out;
}

/* ---------- 写表 ---------- */
/* 客户自己讲过的：同一个 SET 不同售价 = 不同货。这句会印给司机看。 */
var SET_DESC_ = {
  '9 ITEMS|35': '普通雨伞',
  '9 ITEMS|36': '普通雨伞',
  '9 ITEMS|40': '好的雨伞',
  '10 ITEMS|50': '好的雨伞',
  'UMBRELLA|12': '普通雨伞',
  'UMBRELLA|20': '好的雨伞',
  'UMBRELLA|26': '客制 logo 雨伞'
};

var ORDER_HEAD = ['ORDER_ID', 'DATE', 'MONTH', 'REGION', 'STATE', 'BRANCH', 'BRAND', 'SALESMAN',
  'SET_TYPE', 'UNIT_PRICE', 'QTY', 'TOTAL_INCOME', 'MY_INCOME', 'DRIVER_FEE',
  'PAY_STATUS', 'PAY_DATE', 'DRV_COLOR', 'DRV_GRADE', 'DRV_NOTE',
  'DELIVERY_STATUS', 'NOTE', 'SOURCE',
  'STATUS', 'VOID_BY', 'VOID_AT'];

/** 读出现有 ORDERS 里由 App 新增的单（SOURCE 以 APP 开头），重跑 INSTALL 时不能弄丢 */
function keepAppRows_(ss) {
  var sh = ss.getSheetByName('ORDERS');
  if (!sh) return [];
  var v = sh.getDataRange().getValues();
  var hr = -1;
  for (var i = 0; i < Math.min(v.length, 5); i++) {
    if (String(v[i][0]).trim() === 'ORDER_ID') { hr = i; break; }
  }
  if (hr < 0) return [];
  var head = v[hr].map(function (x) { return String(x).trim(); });
  var si = head.indexOf('SOURCE');
  if (si < 0) return [];
  var out = [];
  for (var r = hr + 1; r < v.length; r++) {
    if (String(v[r][si]).indexOf('APP') !== 0) continue;
    var o = {};
    for (var c = 0; c < head.length; c++) if (head[c]) o[head[c]] = v[r][c];
    out.push(ORDER_HEAD.map(function (h) { return o.hasOwnProperty(h) ? o[h] : ''; }));
  }
  return out;
}

function writeOrders_(ss, rows) {
  var kept = keepAppRows_(ss);
  rows.sort(function (a, b) { return a.month - b.month; });
  var data = rows.map(function (r, i) {
    return ['VP' + zero5_(i + 1), r.date, r.month, r.region, r.state, r.branch, r.brand, r.salesman,
      r.detail, r.price, r.qty, r.total, r.mine, r.fee, r.pay || 'UNPAID', r.payDate,
      '', '', '', 'DELIVERED', r.note, r.src, 'ACTIVE', '', ''];
  });
  // App 新增的单接在历史後面，重新编号避免撞号
  kept.forEach(function (row, i) { row[0] = 'VP' + zero5_(data.length + i + 1); data.push(row); });
  writeSheet_(ss, 'ORDERS', ORDER_HEAD, data,
    '● 主表：所有订单。App 会自动往下 append。旧的月份分页原封不动保留。重跑 INSTALL 不会弄丢 App 新增的单。');
  return kept.length;
}
function zero5_(n) { n = String(n); while (n.length < 5) n = '0' + n; return n; }

function writeSheet_(ss, name, head, data, note) {
  var sh = ss.getSheetByName(name);
  if (sh) sh.clear(); else sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, 1).setValue(note || '').setFontSize(9).setFontColor('#888888').setFontStyle('italic');
  sh.getRange(2, 1, 1, head.length).setValues([head])
    .setBackground('#1C1C1E').setFontColor('#E8C86A').setFontWeight('bold')
    .setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center');
  if (data.length) {
    sh.getRange(3, 1, data.length, head.length).setValues(data).setFontFamily('Arial').setFontSize(10);
  }
  sh.setFrozenRows(2);
  for (var c = 1; c <= head.length; c++) sh.autoResizeColumn(c);
  return sh;
}

/** 已存在就不覆盖（保护使用者改过的 PIN / 电话 / 设定） */
function writeSheetKeep_(ss, name, head, data, note) {
  if (ss.getSheetByName(name)) return;
  writeSheet_(ss, name, head, data, note);
}

/* ---------- 对账报告（选用） ---------- */
function RECONCILE() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = readHistory_(ss);
  var sp = derivePrices_(rows), look = {};
  sp.forEach(function (s) { look[s[0] + '|' + s[1]] = s[2]; });
  var ok = 0, bad = 0, none = 0, diffs = [];
  rows.forEach(function (r) {
    var p = look[r.detail + '|' + r.price];
    if (p === undefined) { none++; return; }
    var calc = Math.round(p * r.qty * 100) / 100;
    if (Math.abs(calc - r.mine) < 0.01) ok++;
    else { bad++; diffs.push([r.month, r.branch, r.salesman, r.detail, r.price, r.qty, r.total, r.mine, calc, Math.round((r.mine - calc) * 100) / 100]); }
  });
  writeSheet_(ss, '审核_利润差异',
    ['月份', 'BRANCH', 'SALESMAN', 'SET_TYPE', '售价', '数量', 'TOTAL', '原表My income', '按价目表算', '差额'], diffs,
    '● 这些旧记录的 My income 跟价目表算出来对不上（多数是特别谈的价）。仅供参考，不影响 App 使用。');
  var msg = '对账：相符 ' + ok + ' / 不符 ' + bad + ' / 无对照 ' + none + '（共 ' + rows.length + '）';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { }
  return msg;
}
