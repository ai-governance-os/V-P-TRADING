/** 由 GitHub Actions 在部署时填入 commit SHA；用来验证线上真的更新了 */
var BUILD_ID = '__BUILD_ID__';

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

/**
 * 只读最後 n 笔资料（表头照样抓）。
 * 用在「只要最新几百笔」的地方，免得为了 300 笔把一千多列整张搬回来。
 * truncated = true 表示上面还有没读到的资料，呼叫端要自己决定要不要补读整张。
 */
function readTailTable_(name, n) {
  var sheet = sh_(name);
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  var empty = { head: [], rows: [], headRow: 1, sheet: sheet, truncated: false };
  if (!lastRow || !lastCol) return empty;

  // 表头可能不在第 1 列（上面有说明行），跟 readTable_ 用同一套判断
  var top = sheet.getRange(1, 1, Math.min(5, lastRow), lastCol).getValues();
  var hr = -1, head = null;
  for (var i = 0; i < top.length; i++) {
    var nonEmpty = top[i].filter(function (x) { return x !== '' && x !== null; }).length;
    if (nonEmpty >= 2) { hr = i; head = top[i].map(function (h) { return String(h).trim(); }); break; }
  }
  if (!head) return empty;

  var firstData = hr + 2;                 // 表头的下一列（1-based）
  var total = lastRow - firstData + 1;
  if (total <= 0) return { head: head, rows: [], headRow: hr + 1, sheet: sheet, truncated: false };

  var take = Math.min(n, total);
  var startRow = lastRow - take + 1;
  var v = sheet.getRange(startRow, 1, take, lastCol).getValues();

  var rows = [];
  for (var r = 0; r < v.length; r++) {
    if (v[r].every(function (x) { return x === '' || x === null; })) continue;
    var o = {};
    for (var c = 0; c < head.length; c++) if (head[c]) o[head[c]] = v[r][c];
    o.__row = startRow + r;
    rows.push(o);
  }
  return { head: head, rows: rows, headRow: hr + 1, sheet: sheet, truncated: take < total };
}

/**
 * 数一数「有手动选过发票品牌」的订单笔数。
 * 用来决定下单页那段新手提醒还要不要显示 —— 学会了就自己退场。
 * 只读 INV_BRAND 那一栏，不整张搬（这支会在 bootstrap 里跑）。
 */
/**
 * 新手提醒还要不要显示。
 *
 * 「学会了」是不可逆的 —— 用过 15 笔就永远不必再数。
 * 所以一旦达标就把结论记进 PropertiesService，之後直接回 false，
 * 不再扫那一千多行订单。bootstrap 每次存设定都会重算，这里不能省着不做。
 */
function brandHintOn_() {
  try {
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty('BRAND_HINT_DONE') === '1') return false;
    if (invBrandUsedCount_() >= 15) {
      props.setProperty('BRAND_HINT_DONE', '1');
      return false;
    }
    return true;
  } catch (e) { return true; }
}

function invBrandUsedCount_() {
  try {
    var sheet = sh_('ORDERS');
    var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    if (lastRow < 2 || !lastCol) return 0;

    var top = sheet.getRange(1, 1, Math.min(5, lastRow), lastCol).getValues();
    var hr = -1, head = null;
    for (var i = 0; i < top.length; i++) {
      var ne = top[i].filter(function (x) { return x !== '' && x !== null; }).length;
      if (ne >= 2) { hr = i; head = top[i].map(function (h) { return String(h).trim(); }); break; }
    }
    if (!head) return 0;

    var c = head.indexOf('INV_BRAND');
    if (c < 0) return 0;
    var first = hr + 2, n = lastRow - first + 1;
    if (n <= 0) return 0;

    var v = sheet.getRange(first, c + 1, n, 1).getValues();
    var used = 0;
    for (var r = 0; r < v.length; r++) if (String(v[r][0]).trim()) used++;
    return used;
  } catch (e) { return 0; }
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
    var stat = { paid: 0, unpaid: 0, op: 0, pc: 0, dated: 0, noted: 0, skipped: 0, fixed: 0, noMethod: 0, backfilled: 0 };

    for (var i = 0; i < n; i++) {
      var s = String(v[i][cS] == null ? '' : v[i][cS]).trim().toUpperCase();
      if (s === 'PAID' || s === 'UNPAID') {
        stat.skipped++;
        // 旧版下单程式留下的：写了 PAID 却没有收款日期。用订单日期补上，
        // 免得司机月结单与月报那一栏空白。补过就不会再补，重跑依然安全。
        if (s === 'PAID' && !String(v[i][cD] == null ? '' : v[i][cD]).trim() && cO >= 0) {
          var od0 = v[i][cO];
          var iso0 = Object.prototype.toString.call(od0) === '[object Date]'
            ? Utilities.formatDate(od0, TZ, 'yyyy-MM-dd')
            : (String(od0 || '').match(/^\d{4}-\d{2}-\d{2}/) || [''])[0];
          if (iso0) { v[i][cD] = iso0; stat.backfilled++; }
        }
        continue;
      }

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
           '）｜未付 ' + stat.unpaid + '（其中 ' + stat.fixed + ' 笔是修正回来的）｜已是新格式略过 ' + stat.skipped +
           (stat.backfilled ? '｜补上收款日期 ' + stat.backfilled + ' 笔' : '')
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
    var pr = toNum_(r.UNIT_PRICE);
    setTypes[k].push({
      price: pr, profit: toNum_(r.PROFIT_PER_SET),
      // 表里填了就以表为准；没填就用内建的（客户自己讲过的那几组）
      desc: String(r['说明'] || r.DESC || '').trim() || (DESC_FALLBACK_[k + '|' + pr] || '')
    });
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
    build: (typeof BUILD_ID === 'string' ? BUILD_ID : ''),
    brands: invBrandList_().slice().sort(),
    // 新手提醒：手动选过 15 笔就当他们学会了，之後不再显示（不靠浏览器存档）
    showBrandHint: brandHintOn_(),
    setTypes: setTypes,
    people: people,
    branches: branches
  };
}

/* 客户在确认文件里自己写过的：同一个 SET 不同售价 = 不同货。
   SET_PRICE 分页的「说明」栏填了就盖过这里，这只是还没填之前的预设。 */
var DESC_FALLBACK_ = {
  '10 ITEMS|50': '好的雨伞',
  '9 ITEMS|35': '普通雨伞',
  '8 ITEMS|55': '好的雨伞',
  '9 ITEMS|40': '好的雨伞',
  '9 ITEMS|38': '普通雨伞',
  '10 ITEMS|55': '好的雨伞',
  'UMBRELLA|12': '普通雨伞',
  'UMBRELLA|13': '普通雨伞',
  '10 ITEMS|45': '普通雨伞',
  'UMBRELLA|20': '好的雨伞',
  'UMBRELLA|26': '普通雨伞',
  'UMBRELLA + BAG|16': '普通雨伞',
  'UMBRELLA|14': '普通雨伞',
  'UMBRELLA|25': '好的雨伞',
  '9 ITEMS|36': '普通雨伞',
  '8 ITEMS|26': '普通雨伞',
  '9 ITEMS|12': '普通雨伞',
  'BLACK UMBRELLA|14': '普通雨伞',
  'RED UMBRELLA|12': '普通雨伞',
  '8 ITEMS|60': '好的雨伞',
  'PREMIUM UMBRELLA|20': '好的雨伞',
  '8 ITEMS|50': '好的雨伞',
  'UMBRELLA|11': '普通雨伞',
  '10 ITEMS|40': '好的雨伞'
};

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

  var t = ensureCols_('ORDERS', ['DRV_COLOR', 'DRV_GRADE', 'DRV_NOTE', 'INVOICE_TO', 'INV_BRAND']);
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
      // 发票开给销售员本人还是他的公司 —— 客户说两种都有
      INVOICE_TO: up_(p.invoiceTo) === 'COMPANY' ? 'COMPANY' : 'SA',
      // 发票上印哪个牌子。留空 = 跟分行一样
      INV_BRAND: up_(p.invBrand || ''),
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
  var t = ensureCols_('ORDERS', ['INV_BRAND']);
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
  if (p.invBrand != null) patch.INV_BRAND = up_(p.invBrand);

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
  var lim = opt.limit || 300;
  var q = up_(opt.q || '');

  // 订单越来越多（已经一千多笔），整张表搬回来要十几秒，手机等不及会把分页丢掉。
  // 迴圈本来就从最新往回跑、满 lim 就停 —— 所以「没有任何筛选」时只读尾巴就够。
  //
  // 有筛选（月份／收款／未送／搜寻）就直接读整张：
  // 那些情况筛完通常凑不满 lim，走尾巴只会变成「读了尾巴又读整张」，比原本更慢。
  var filtered = !!(q || opt.month || opt.pay || opt.pending);
  if (filtered) return pickOrders_(readTable_('ORDERS'), opt, q, lim);

  // 只会跳过作废单（很少），抓 lim + 60 列绰绰有余
  var t = readTailTable_('ORDERS', lim + 60);
  var out = pickOrders_(t, opt, q, lim);
  // 万一真的凑不满而且上面还有资料 → 补读整张，结果跟以前完全一样
  if (out.length < lim && t.truncated) {
    out = pickOrders_(readTable_('ORDERS'), opt, q, lim);
  }
  return out;
}

function pickOrders_(t, opt, q, lim) {
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
      drvNote: String(r.DRV_NOTE || ''), invBrand: String(r.INV_BRAND || '')
    });
    if (out.length >= lim) break;
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

/* ---------- 合并分行 ----------
   同一间店有两种写法（PERODUA NILAI / PERODUA NILAI IMPIAN）时，
   把全部订单、销售员名单都搬到正式名称，旧写法从 BRANCH 名单移除。
   抽成会重算 —— 万一两个写法的州属不一样。
------------------------------- */
/* ───────── 分行改名护栏 ───────── */

/** 统一写法：转大写、去头尾空白、中间多个空格并成一个、括号前後不留空 */
function normBranchName_(x) {
  return String(x == null ? '' : x)
    .replace(/\u3000/g, ' ')          // 全形空格
    .trim().toUpperCase()
    .replace(/\s+/g, ' ')
    .replace(/\s*\(\s*/g, '(').replace(/\s*\)\s*/g, ')')
    .trim();
}

/** 两个字串差几个字（Levenshtein）。只用来抓打错字，不必快。 */
function editDist_(a, b) {
  a = String(a); b = String(b);
  if (a === b) return 0;
  var m = a.length, n = b.length;
  if (!m) return n;
  if (!n) return m;
  var prev = [], cur = [], i, j;
  for (j = 0; j <= n; j++) prev[j] = j;
  for (i = 1; i <= m; i++) {
    cur[0] = i;
    for (j = 1; j <= n; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1,
                        prev[j - 1] + (a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1));
    }
    for (j = 0; j <= n; j++) prev[j] = cur[j];
  }
  return prev[n];
}

/** 新名字跟现有分行像不像？像的话回传那几间，让人先看一眼再决定 */
function similarBranches_(name, excludeName) {
  var target = normBranchName_(name), ex = normBranchName_(excludeName);
  if (target.length < 4) return [];
  var out = [];
  readTable_('BRANCH').rows.forEach(function (r) {
    var b = normBranchName_(r.BRANCH);
    if (!b || b === target || b === ex) return;
    // 只差 1~2 个字，或差别只在空格 → 很可能是打错
    var d = editDist_(target, b);
    if (d <= 2 || target.replace(/\s/g, '') === b.replace(/\s/g, '')) {
      out.push({ branch: r.BRANCH, dist: d, n: toNum_(r['历史笔数']) });
    }
  });
  return out.sort(function (a, b) { return a.dist - b.dist; }).slice(0, 3);
}

/** 资料维护纪录（分行 / 销售员 / 价目表共用一本），出事查得到也改得回 */
function logChange_(rec) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('维护纪录');
    var head = ['时间', '类别', '动作', '从', '到', '影响订单', '影响名单', '抽成变动', '操作人'];
    if (!sh) {
      sh = ss.insertSheet('维护纪录');
      sh.getRange(1, 1, 1, head.length)
        .setValues([['● 资料维护纪录（只进不出，可用来还原）', '', '', '', '', '', '', '', '']]);
      sh.getRange(2, 1, 1, head.length).setValues([head]).setFontWeight('bold');
      sh.setFrozenRows(2);
    }
    sh.appendRow([
      Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'),
      rec.kind || '', rec.action || '', rec.from || '', rec.to || '',
      rec.orders == null ? '' : rec.orders,
      rec.rows == null ? '' : rec.rows,
      rec.feeDiff == null ? '' : rec.feeDiff, rec.by || ''
    ]);
  } catch (e) { /* 纪录写不进去也不该挡住主流程 */ }
}
function logBranchChange_(rec) {
  logChange_({ kind: '分行', action: rec.action, from: rec.from, to: rec.to,
               orders: rec.orders, rows: rec.salesmen, feeDiff: rec.feeDiff, by: rec.by });
}

function previewMergeBranch(p) {
  // 先统一写法：大写、去掉多余空格。「PERODUA  SEREMBAN 2」和「perodua seremban 2」
  // 在这里就变成同一个字串，不会因为多打一个空格生出第二间分行。
  var from = normBranchName_(p && p.from), to = normBranchName_(p && p.to);
  if (!from || !to) return { ok: false, msg: '请选两间分行' };
  if (from === to) return { ok: false, msg: '两个是同一个名字' };
  if (to.length < 3) return { ok: false, msg: '名字太短，请打完整的分行名' };

  var toInfo = branchInfo_(to), fromInfo = branchInfo_(from), rename = false, near = [];
  if (!toInfo) {
    // 新名字不在名单里 → 当作「改名」，州属沿用旧的，品牌照新名字重算
    if (!fromInfo) return { ok: false, msg: '名单里没有：' + from + '，也没有：' + to };
    var nb = brandOf_(to);
    toInfo = {
      branch: to,
      brand: (nb && nb !== 'OTHER') ? nb : fromInfo.brand,
      state: fromInfo.state, region: fromInfo.region
    };
    rename = true;
    // 打错字防护：新名字跟现有某间很像但不一样，先问一句
    near = similarBranches_(to, from);
  }

  var t = readTable_('ORDERS');
  var n = 0, feeDiff = 0;
  t.rows.forEach(function (r) {
    if (up_(r.BRANCH) !== up_(from) || isVoid_(r)) return;
    n++;
    var nf = calcDriverFee_(toInfo.region, toInfo.state, String(r.SET_TYPE), toNum_(r.QTY), toNum_(r.UNIT_PRICE));
    feeDiff += nf - toNum_(r.DRIVER_FEE);
  });
  var sm = readTable_('SALESMAN').rows.filter(function (r) { return up_(r.BRANCH) === up_(from); }).length;
  if (!n && !sm && !branchInfo_(from))
    return { ok: false, msg: '名单里没有「' + from + '」，也没有任何订单用这个写法' };
  return {
    ok: true, from: from, to: toInfo.branch, orders: n, salesmen: sm, rename: rename,
    feeDiff: Math.round(feeDiff * 100) / 100,
    fromInfo: fromInfo, toInfo: toInfo,
    near: near,
    brandChanged: !!(fromInfo && toInfo.brand && up_(fromInfo.brand) !== up_(toInfo.brand)),
    brandFrom: fromInfo ? fromInfo.brand : '', brandTo: toInfo.brand
  };
}

function mergeBranch(p) {
  var pv = previewMergeBranch(p);
  if (!pv.ok) return pv;
  var from = pv.from, toInfo = pv.toInfo;

  var t = readTable_('ORDERS');
  var cB = t.head.indexOf('BRANCH'), cBr = t.head.indexOf('BRAND'),
      cS = t.head.indexOf('STATE'), cR = t.head.indexOf('REGION'),
      cF = t.head.indexOf('DRIVER_FEE');
  var hits = t.rows.filter(function (r) { return up_(r.BRANCH) === up_(from); });

  var res = withLock_(function () {
    hits.forEach(function (r) {
      var rng = t.sheet.getRange(r.__row, 1, 1, t.head.length);
      var v = rng.getValues()[0];
      if (cB >= 0) v[cB] = toInfo.branch;
      if (cBr >= 0 && toInfo.brand) v[cBr] = toInfo.brand;
      if (cS >= 0) v[cS] = toInfo.state;
      if (cR >= 0) v[cR] = toInfo.region;
      if (cF >= 0 && !isVoid_(r))
        v[cF] = calcDriverFee_(toInfo.region, toInfo.state, String(r.SET_TYPE), toNum_(r.QTY), toNum_(r.UNIT_PRICE));
      rng.setValues([v]);
    });

    // 销售员名单：旧写法改成新的；同一人已经有新分行那一行就把旧的清空
    var st = readTable_('SALESMAN');
    var sB = st.head.indexOf('BRANCH') + 1;
    var have = {};
    st.rows.forEach(function (r) {
      if (up_(r.BRANCH) === up_(toInfo.branch)) have[up_(r.SALESMAN)] = 1;
    });
    var kill = [];
    st.rows.forEach(function (r) {
      if (up_(r.BRANCH) !== up_(from)) return;
      if (have[up_(r.SALESMAN)]) kill.push(r.__row);
      else if (sB > 0) st.sheet.getRange(r.__row, sB).setValue(toInfo.branch);
    });
    kill.sort(function (a, b) { return b - a; }).forEach(function (row) { st.sheet.deleteRow(row); });

    // BRANCH 名单：改名就改字，合并就把旧的删掉
    var bt = readTable_('BRANCH');
    var mine2 = bt.rows.filter(function (r) { return up_(r.BRANCH) === up_(from); });
    if (pv.rename) {
      var bc = bt.head.indexOf('BRANCH') + 1;
      if (bc > 0) mine2.forEach(function (r) { bt.sheet.getRange(r.__row, bc).setValue(toInfo.branch); });
    } else {
      mine2.map(function (r) { return r.__row; })
        .sort(function (a, b) { return b - a; })
        .forEach(function (row) { bt.sheet.deleteRow(row); });
    }

    return { ok: true };
  });
  if (!res.ok) return res;

  clearBootCache_();
  logBranchChange_({
    action: pv.rename ? '改名' : '合并', from: from, to: toInfo.branch,
    orders: hits.length, salesmen: pv.salesmen, feeDiff: pv.feeDiff, by: (p && p.by) || ''
  });
  return { ok: true, from: from, to: toInfo.branch, orders: hits.length,
           salesmen: pv.salesmen, feeDiff: pv.feeDiff, rename: pv.rename };
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

/* ═════════ 销售员改名 / 合并 ═════════ */

/** 销售员名字统一写法：大写、去掉多余空格 */
function normName_(x) {
  return String(x == null ? '' : x).replace(/　/g, ' ').trim().toUpperCase().replace(/\s+/g, ' ');
}

function previewRenameSalesman(p) {
  var from = normName_(p && p.from), to = normName_(p && p.to);
  var branch = p && p.branch ? up_(p.branch) : '';
  if (!from || !to) return { ok: false, msg: '请填两个名字' };
  if (from === to) return { ok: false, msg: '两个是同一个名字' };
  if (to.length < 2) return { ok: false, msg: '名字太短' };

  var st = readTable_('SALESMAN');
  var fromRows = st.rows.filter(function (r) {
    return normName_(r.SALESMAN) === from && (!branch || up_(r.BRANCH) === branch);
  });
  var toRows = st.rows.filter(function (r) { return normName_(r.SALESMAN) === to; });

  var t = readTable_('ORDERS');
  var n = 0, branches = {};
  t.rows.forEach(function (r) {
    if (normName_(r.SALESMAN) !== from || isVoid_(r)) return;
    branches[String(r.BRANCH || '')] = 1;
    if (branch && up_(r.BRANCH) !== branch) return;
    n++;
  });

  if (!n && !fromRows.length)
    return { ok: false, msg: branch
      ? '「' + from + '」在 ' + branch + ' 没有单，也没有名单资料'
      : '名单里没有「' + from + '」，也没有任何订单用这个名字' };

  var near = [];
  if (!toRows.length) {
    var seen = {};
    st.rows.forEach(function (r) {
      var nm = normName_(r.SALESMAN);
      if (!nm || nm === to || nm === from || seen[nm]) return;
      seen[nm] = 1;
      if (editDist_(to, nm) <= 2) near.push(nm);
    });
    near = near.slice(0, 3);
  }

  return {
    ok: true, from: from, to: to, orders: n,
    rows: fromRows.length,
    merge: toRows.length > 0,
    branches: Object.keys(branches).sort(),
    branch: branch,
    near: near
  };
}

function renameSalesman(p) {
  var pv = previewRenameSalesman(p);
  if (!pv.ok) return pv;
  var from = pv.from, to = pv.to, branch = pv.branch;

  var res = withLock_(function () {
    var t = readTable_('ORDERS');
    var c = t.head.indexOf('SALESMAN') + 1;
    var cb = t.head.indexOf('BRANCH') + 1;
    if (c > 0) {
      t.rows.forEach(function (r) {
        if (normName_(r.SALESMAN) !== from) return;
        if (branch && (cb <= 0 || up_(r.BRANCH) !== branch)) return;
        t.sheet.getRange(r.__row, c).setValue(to);
      });
    }
    var st = readTable_('SALESMAN');
    var sc = st.head.indexOf('SALESMAN') + 1;
    var scb = st.head.indexOf('BRANCH') + 1;
    var have = {};
    st.rows.forEach(function (r) {
      if (normName_(r.SALESMAN) === to) have[up_(r.BRANCH)] = 1;
    });
    var kill = [];
    st.rows.forEach(function (r) {
      if (normName_(r.SALESMAN) !== from) return;
      if (branch && (scb <= 0 || up_(r.BRANCH) !== branch)) return;
      if (have[up_(r.BRANCH)]) kill.push(r.__row);
      else if (sc > 0) st.sheet.getRange(r.__row, sc).setValue(to);
    });
    kill.sort(function (a, b) { return b - a; }).forEach(function (row) { st.sheet.deleteRow(row); });
    return { ok: true };
  });
  if (!res.ok) return res;

  clearBootCache_();
  logChange_({ kind: '销售员', action: pv.merge ? '合并' : (branch ? '拆分改名' : '改名'), from: from, to: to,
               orders: pv.orders, rows: pv.rows, branch: branch || '', by: (p && p.by) || '' });
  return { ok: true, from: from, to: to, orders: pv.orders, rows: pv.rows, merge: pv.merge, branch: branch || '' };
}

/* ═════════ 价目表维护 ═════════ */

/** 给维护介面看的完整价目表（含已停用的） */
function listSetPrice() {
  var t = ensureCols_('SET_PRICE', ['说明', '英文品名']);
  var used = {};
  readTable_('ORDERS').rows.forEach(function (r) {
    if (isVoid_(r)) return;
    var k = up_(r.SET_TYPE) + '|' + toNum_(r.UNIT_PRICE);
    used[k] = (used[k] || 0) + 1;
  });
  var out = [];
  t.rows.forEach(function (r) {
    if (!r.SET_TYPE) return;
    var st = String(r.SET_TYPE).trim(), pr = toNum_(r.UNIT_PRICE);
    out.push({
      setType: st, price: pr, profit: toNum_(r.PROFIT_PER_SET),
      desc: String(r['说明'] || '').trim() || (DESC_FALLBACK_[up_(st) + '|' + pr] || ''),
      invName: String(r['英文品名'] || '').trim() ||
               invNameOf_(st, String(r['说明'] || '').trim() || (DESC_FALLBACK_[up_(st) + '|' + pr] || '')),
      active: up_(r.ACTIVE) !== 'NO',
      used: used[up_(st) + '|' + pr] || 0
    });
  });
  out.sort(function (a, b) {
    return a.setType === b.setType ? a.price - b.price : (a.setType < b.setType ? -1 : 1);
  });
  return { ok: true, list: out };
}

/** 停用 / 恢复。不做真删除 —— 历史订单要查得回价目。
 * 价目表偶尔会有同个 SET_TYPE+价钱重复好几行的历史资料（例如汇入时重复），
 * 之前这里只认「最后找到的那一行」，重复资料时按了停用会看起来「没反应」——
 * 其实是改到了另一行。现在改成同样价钱的全部一起处理，不会再卡住。 */
function toggleSetPrice(p) {
  var st = up_(p && p.setType), pr = toNum_(p && p.price);
  var want = !!(p && p.active);
  if (!st || pr <= 0) return { ok: false, msg: '请选一组价目' };

  var t = ensureCols_('SET_PRICE', ['说明']);
  var hits = [];
  t.rows.forEach(function (r) {
    if (up_(r.SET_TYPE) === st && toNum_(r.UNIT_PRICE) === pr) hits.push(r);
  });
  if (!hits.length) return { ok: false, msg: '价目表里找不到这一组' };

  var c = t.head.indexOf('ACTIVE') + 1;
  if (c <= 0) return { ok: false, msg: 'SET_PRICE 缺少 ACTIVE 栏位' };

  var r = withLock_(function () {
    hits.forEach(function (hit) {
      t.sheet.getRange(hit.__row, c).setValue(want ? 'YES' : 'NO');
    });
    return { ok: true };
  });
  if (!r.ok) return r;

  clearBootCache_();
  logChange_({ kind: '价目表', action: want ? '恢复' : '停用',
               from: st + ' RM ' + pr + (hits.length > 1 ? '（' + hits.length + ' 行重复一起改）' : ''),
               to: '', by: (p && p.by) || '' });
  return { ok: true, setType: st, price: pr, active: want, rows: hits.length };
}

/** 新增一组价目。已存在就当作「改」——顺便把停用的恢复回来。 */
function saveSetPrice(p) {
  var st = up_(p && p.setType).replace(/\s+/g, ' ').trim();
  var pr = toNum_(p && p.price), pf = toNum_(p && p.profit);
  var desc = String((p && p.desc) || '').trim();
  if (!st) return { ok: false, msg: '请填 SET 类型' };
  if (pr <= 0) return { ok: false, msg: '售价要大于 0' };
  if (pf < 0) return { ok: false, msg: '每 set 利润不能是负的' };
  if (pf > pr) return { ok: false, msg: '利润比售价还高，请再确认' };

  var t = ensureCols_('SET_PRICE', ['说明', '英文品名']);
  var hit = null;
  t.rows.forEach(function (r) {
    if (up_(r.SET_TYPE) === st && toNum_(r.UNIT_PRICE) === pr) hit = r;
  });

  var cP = t.head.indexOf('PROFIT_PER_SET') + 1,
      cD = t.head.indexOf('说明') + 1,
      cN = t.head.indexOf('英文品名') + 1,
      cA = t.head.indexOf('ACTIVE') + 1;
  var inv = String((p && p.invName) || '').trim();

  var r = withLock_(function () {
    if (hit) {
      if (cP > 0) t.sheet.getRange(hit.__row, cP).setValue(pf);
      if (cD > 0 && desc) t.sheet.getRange(hit.__row, cD).setValue(desc);
      if (cN > 0 && inv) t.sheet.getRange(hit.__row, cN).setValue(inv);
      if (cA > 0) t.sheet.getRange(hit.__row, cA).setValue('YES');
    } else {
      t.sheet.appendRow(t.head.map(function (h) {
        return h === 'SET_TYPE' ? st : h === 'UNIT_PRICE' ? pr
          : h === 'PROFIT_PER_SET' ? pf : h === '说明' ? desc
            : h === '英文品名' ? (inv || invNameOf_(st, desc))
              : h === 'ACTIVE' ? 'YES' : h === '历史笔数' ? 0 : '';
      }));
    }
    return { ok: true };
  });
  if (!r.ok) return r;

  clearBootCache_();
  logChange_({ kind: '价目表', action: hit ? '修改' : '新增',
               from: st + ' RM ' + pr, to: '利润 RM ' + pf + (desc ? ' · ' + desc : ''),
               by: (p && p.by) || '' });
  return { ok: true, updated: !!hit, setType: st, price: pr, profit: pf, desc: desc };
}

/* ═════════════════════════════════════════════════════════
 * 一次性：客户第二轮回覆确认後的分行整并（2026-07-30）
 * 在 Apps Script 编辑器选这个函式按 Run。
 * 每一步先用线上真实资料算一次，对不上预期就立刻中止。
 * 跑完把这个函式删掉。
 * ═════════════════════════════════════════════════════════ */
function RUN_MERGE_20260730() {
  var log = [], abort = null;
  function say(s) { log.push(s); }

  function snap() {
    var t = readTable_('ORDERS'), n = 0, inc = 0, mine = 0, fee = 0;
    t.rows.forEach(function (r) {
      if (isVoid_(r)) return;
      n++; inc += toNum_(r.TOTAL_INCOME); mine += toNum_(r.MY_INCOME); fee += toNum_(r.DRIVER_FEE);
    });
    return { n: n, inc: Math.round(inc * 100) / 100,
             mine: Math.round(mine * 100) / 100, fee: Math.round(fee * 100) / 100 };
  }

  var before = snap();
  say('【开始前】' + before.n + ' 笔｜营业额 RM ' + before.inc +
      '｜利润 RM ' + before.mine + '｜司机抽成 RM ' + before.fee);

  // [从, 到, 预期笔数, 预期是改名?]
  // EMAS → EMAS PUTRAJAYA 不在名单里：底下还有两笔柔佛单（VP00401 / VP00742）
  // 归属未定，合了会把柔佛营业额算进雪兰莪，也会给司机他没送过的抽成。
  var PLAN = [
    ['PROTON TAMAN MIDAS',      'PROTON TAMAN MIDAH', 1, true ],
    ['PROTON MIDAH',            'PROTON TAMAN MIDAH', 2, false],
    ['PERODUA NILAI IMPIAN',    'PERODUA NILAI',      7, false],
    ['JAECOO SEREMBAN',         'JAECOO SEREMBAN 2',  2, false],
    ['PROTON IOI PUTRAJAYA',    'PROTON PUTRAJAYA',   1, false],
    ['PROTON PUTRAJAYA - EMAS', 'EMAS PUTRAJAYA',     1, false]
  ];

  for (var i = 0; i < PLAN.length && !abort; i++) {
    var from = PLAN[i][0], to = PLAN[i][1], expN = PLAN[i][2], expRename = PLAN[i][3];
    var tag = '[' + (i + 1) + '] ' + from + ' → ' + to;

    var p = previewMergeBranch({ from: from, to: to });
    if (!p.ok) { abort = tag + ' 预览失败：' + p.msg; break; }
    if (p.orders !== expN) { abort = tag + ' 中止：预期 ' + expN + ' 笔，实际 ' + p.orders + ' 笔'; break; }
    if (Math.abs(p.feeDiff) > 0.005) { abort = tag + ' 中止：抽成会变动 RM ' + p.feeDiff + '，预期 0'; break; }
    if (!!p.rename !== expRename) { abort = tag + ' 中止：预期' + (expRename ? '改名' : '合并') + '，实际是' + (p.rename ? '改名' : '合并'); break; }

    var r = mergeBranch({ from: from, to: to, by: 'MERGE-20260730' });
    if (!r.ok) { abort = tag + ' 执行失败：' + r.msg; break; }
    say(tag + '｜' + (r.rename ? '改名' : '合并') + ' ' + r.orders + ' 笔｜销售员 ' + r.salesmen + ' 行｜抽成变动 RM ' + r.feeDiff);
  }

  var after = snap();
  say('【结束後】' + after.n + ' 笔｜营业额 RM ' + after.inc +
      '｜利润 RM ' + after.mine + '｜司机抽成 RM ' + after.fee);

  var same = before.n === after.n &&
             Math.abs(before.inc - after.inc) < 0.005 &&
             Math.abs(before.mine - after.mine) < 0.005 &&
             Math.abs(before.fee - after.fee) < 0.005;
  say(same ? '✓ 对帐通过：笔数、营业额、利润、司机抽成四项完全没变'
           : '✗ 对帐不符！笔数 ' + (after.n - before.n) +
             '｜营业额 ' + Math.round((after.inc - before.inc) * 100) / 100 +
             '｜利润 ' + Math.round((after.mine - before.mine) * 100) / 100 +
             '｜抽成 ' + Math.round((after.fee - before.fee) * 100) / 100);

  if (abort) say('■ 已中止：' + abort + '（中止前的步骤已生效，之後的都没做）');

  var left = readTable_('ORDERS').rows.filter(function (r) {
    return up_(r.BRANCH) === 'EMAS' && !isVoid_(r);
  });
  if (left.length) {
    say('── 没有动、留给客户决定的 ──');
    left.forEach(function (r) {
      say('　' + r.ORDER_ID + '｜' + r.DATE + '｜' + r.STATE + '/' + r.REGION + '｜' +
          r.SALESMAN + '｜' + r.SET_TYPE + ' ' + r.QTY + '×' + r.UNIT_PRICE + '｜RM ' + r.TOTAL_INCOME);
    });
    say('　→ 两笔柔佛单要先各自归位，「EMAS」才能并去 EMAS PUTRAJAYA。');
  }

  var out = log.join('\n');
  Logger.log(out);
  return out;
}

/* ═════════════════════════════════════════════════════════
 * 发票（Invoice）
 * 客户第三轮回覆：一个月一张、每笔一行、全英文、不写收款状况、
 * 抬头有时是销售员本人有时是公司、要印地址与银行户口。
 * ═════════════════════════════════════════════════════════ */

var VP_LOGO_ = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEIAAABCCAIAAABsNpe/AAAT6ElEQVR42s1beVxTV74/997sKyELgRDCpk+WCqjY2iriVkWsiihUrXWpVrGj1nm1b6abM51Op7XbzJtX7TZVxNHBpbVWi0txRFmsAhYxIgiiAUISkkAge3LveX8ceycf1lCd13c+fAjce7bv+f5+v/M7v98JBiEED1woioIQ4jiOYdiIlSGEFEUBAIKsH0zBHgQGPRv6SV9fn9ls7uoyG40Gn8/PYjF9Pj9J+r1en0wmk8vlKlWEQqH4efgfPgyKoujZ22y9HR0dJpPJbDY7HA4GgwEA4PP5TCaTz+d5vT6fz+t0unp6emw2m8ViJUm/SCROSBj36KOTIyMjUSckST4ImFHDoAG4XK7Gxkaz2ez1+rxer81mMxoNHR363t7eri4zhgG/3w8hhBAymSylUqlSRcTHx8fGxoSEhPT29jU1NdXW1pIkNWXKo9nZC0JCxA8CZhQwIIRoAJutt7Hxltls8ft9XV3mpqYmna4NxzGpVBoREaFQKGQyGZPJAAADAHR3d9+9e/fu3bvNzS137tyx2+1isfiRRx6ZMSNzwoQ0p9NVWlp682bD+PGPrF//XGxs7M8DEywMhMHtdl+7ds1ms3E43I6Ojps3b/L5/PT09OTkpPDw8BE7cTgc9+7dKy+vOHfu3JUrV61Wa2JiYn5+XnJyUk1N7YULZWPGxL/88o6oqKh+cvsQYNAkNDU1tbW1SSQSvb5Tp9OpVKqMjGkSiaSf/Rmq9Ftgu91eXl5x8OCh06dLKArm5i7JyMior6+/dOlSdvaCHTv+k8lkkiRJEMRDY8Ptdjc03GIyGQRB1NXVKRRh06dnoAFIksR+KsFQigqGYfRKd3Z2FhUd+OSTTzs7O/Pz86dNm3by5LcGg/Hdd/+UkZFBUVQwnY8Mw2az6XS68PBwq9VqNJomTEjj8/kPblsQezSenp6evXv37dq1y+Pxrl//HEEQX3319bPPrnr11VeCEjA4bPH5fFar1efz2Wy2rq4u9JAkSfjwCkVRtE3T6XQbNxYwGOwpU574wx/+mJiYvHLlMy6Xa8RBR4BBN0b7FPr97yiBYE6ePBUTExcSErphw8ZHHkmZP3+B3W4ffnQwTL8jPnnohSRJBKajoyMnJ5fF4syaNScubsyiRTlOp5OiqKHmAIYn4RcpCInP59+27UUWizNxYrpUKl+7dh1ibFAkQ7Lh8Xg6OzuHWYB/Ny1o3Dff/AOfL4qPH8vnC99++080yBFguFwuj8dDkuS1az+63e5fkBOKopBQ7Nz5O4FAFBMTFxoqu3jx4qDCAvop2e3btyGEFRUVbW3t/zf6EIzev/DCr0JCQiMiItPTH7XZbDRX/WGgp7W1tQ6Ho6WlpaKi8hfXkEBO3G73vHnzFQqlWCzZufP3A0UL0Bh0urbKyiqfz1dYuN/n8yEMVHAlyAn9vB7QTJqbW2Jj4yMiIiMjo65fr6dFrj+M06fP+P3+48eP37hx4/8JFf1s1+effxESEhoaKluzZm2/GeLII2hububz+X19fUajKSkpCW3+yF4hZnw+n3dAQebP7XZ7vd4RvQ+6q8Ae6Cd+v394t5IkyXXr1kyfnoFh2JkzZ6urq3Ecp51RBvqoqanNyppXXHx4xoxMurHT6Vy1anVLSwtFUSKRCMMwCCl0igAAYhju8Xi8Xg9Fwfz8vNdee3UYzwdC+MILW8rLy9lstlAoQH4chgEMw+x2u93uePPN3+fn55EkyWAwBvppyDvEcfyNN95YuHBRX19fYeH+SZMm/cshhBB2dhq+/vprvV6PDHOgA3Lnzp3z58+/9NIOqVQulcoVCqVcHiaXh4WFhYvFkrS0CXv2fHL1arXZbB5RMHQ6XUlJyfr1G0SiELlcqVAoFQqlQCBat+65774r6enpCRShQbUFiVZBwWY+Xzh+fCry8VBNACEsLT3f2NhYWLj/n//851Ba8d13JRERkVFR0fQPUrWfIeivvPKaRCKNixsTGir7+OPdP02RbG5uvnz5Mr1ZDQSD7Gx9fb1arRGJQg4c+DuNDQcAmExGgUB461bjY4891i/SgRr7fL6srHnPPLOyu7sbwzAIodfrFYvFKpWKJEkEOxjP3Ov1UhQ1d+6TOI5brda8vGWbNxf4/f633vrj+PHjp02bvmDBwrlzsz777DOz2UIQBIZhJEkGagiEMDk5ef78LLvdce7cOSRvAADcYrEIhcLW1lYOh83hcAKboUoEQRAEQVHUpk0bJRIJ0kUmk2mxWO7du4sGC+bUgbrCMKypqdHtdoeFhb3yym8BAO+99/7Onb9TqSJjY2O5XG5VVdWWLS/Onj1n1673rFYrQRD3TWrAsSInZ/G8eU9aLBaDwYCwgerqmurqml273vv66+NDeSy0pG3duk0kCtFoYjSaGIFA9MEHHw7TZKhOli9fiWHEb37zCoSwoaFBoVCmpKQ5HA6fz6fVavfs+WTq1AweTyAUhowfn1pUdGCgt9rT09PU1PT00ytOnDiBJoAbDAY2m339en1Kyvh+EtVPJAAAq1Y9w+FwkMhyOJxvvjnh8Xjur0dwkaGWlpbKygqFQrFsWS4AoKjogMlkyM3N4fF4OE4kJiZu2rSxpOTUyy/v4HK5JpNp8+YX1qxZ293dHWhehUJhSEhIXFzs5cs/3J/b0aPHKioqc3KW+Hy+YPyChQsXhYbKECEhIaEnT54KkhBU54MPPsQwRm7uUoqizGZzYmJyeLiqsbEJVUB6iOrv3r1HIpHGxcULBKLMzJkGg4G2SyRJulyu4uLiZcvy7rOBYaCtTcdkshgMxvChDbSca9asDtSfwsL9tJ4Nr98EQbjd7uPHj7NYrGXLlmEY9s03JxoaGp58cs7YsWMoiiIIAsdxBoOBZlZQsGnJkiVms0WpVNbW1j777Bqn00nrCYfDSUhI6Ooyo+gJDgBmMBjlchktOcNspRDCefPmjRs3Dp2PhUJheXl5ff2NQMaHifaWlpbW1NQmJydlZWWRJHn48BE2m71y5cp+QyObASHcvn0bn893u90ymay8vPz99z9Ac0CrJpcrHA6HyWQCAOC9vb2dnQYU4RoeBoZhJElxudzly592Op04jhMEYbfb9+3bBwAYXjvQwIcO/cPlci1YsEAkElZWVlZUVEyZMmX69AzEVb8lAwCMGzcuMTHB7Xb7/X6JRPK3v33Z2nqXVsXQUAmbze7p6QEA4GazxWQyhYSEBCMbBIEDAPLz8yMjIz0eD0VRQqHw5MlTen0nQQyp6EgaGxsby8ouqlSq/PxlAIAjR47Z7X3Llz/NYrEGMonYIAhCrY7y+XwYhjGZTKvVWlxcTHPLZDJ5PJ7L5QIA4N3d1r6+Xjr4N6LtpyhKqQybP3++3W7HMIzFYur1nUeOHKF7H8rKHTlytKOjY/r06QkJCW1t7SdPnhwzZmxOzuJhzGPgK4qi2Gz2hQsXEDxatCwWCwAA9/tJu93OYrFGFSxbtWoln89H3gGXyykuPux2uwe1vGjUvr6+Eye+5XA4eXlLAQAnTpxobW3Jz8+TyWQo7jjUQLTvTFEUk8nU6dosFgviCi0rk8kEAOAEQTgcDiQtwRSkzampqRkZGX19fQBgXC5Xq9WWlJQgrgZV7rNnz12/fj01NWXWrFk+n+/o0WMSSeiKFSuGkWQcx/1+v053jzahGIY5HA673fEvHwTHxGIxAAAXCPj37t3z+fyjilsCANatW0MzjuN4YWERyhsNKuWHDx/2ej0LFy7k8/kXLlyoqKiYO3deQsK4QZvQWaimpqbbt5s5HA5NMpvN5nDYtHiTJMnhcAAA+Jgx8Xp9p9vtDh4GEp6ZM2empIx3Op1oW62qqqyuru7nzCHl1mpvXrx4SaOJzstbBgA4dKjY7/c988yK4dUJw7CjR4/19fWh9BWO4z6fLyIiQi6XI1S9vb04jiuVSgAAnp4+WSwWt7a2jmhw+yk6k8lcvny52+1BBxq327NvX/+t8CflPmIwGGbPnh0bG9va2lpSUpKZOWPWrFkomD2wf7Sjtbd3FBUdEAqFdIbR4/FMmzaVIAjknhoMBhaLJRaHAABwjSYqOTmpsbExGIPbz4Dk5CyOilJ7vV4IoUAgOHPmTHt7O70VIuXu7u4+ceJboVC4dGkusldGo3HMmDEcDgc5zvfPPT/B9vv9yBBt377dYrEwmUz0lqIoLpebm7uEnkNHR0dEhArHMZIkcQzDpk3LqK+vR+5A8ISQJCmXy3NyFtvtfTiOs1hMo9F46NA/aBIQmNOnz9y8eXPSpEmZmdNdLtexY1+FhYUdPXo0O3tBbe015IDQy4dhGIPBsNvtGzY8f+5cqUgkQiLKZDK7u7sXLVqUmppKH5Xr6q5PnDjh/s4LISwrK5PLw+7cuTOqgAiytrdu3VKp1JGRUVFR0WFh4WlpE3t7ewMjN4sWLcZxxl/+8t8QwuPHv8Ew/Le/ffXs2XPR0XFcLr+gYHNVVZXNZkMd6vX6Awf+PnVqhlgs0WhiIiOjIiOjoqNjZbKwxMSkjo4OpNZokhs3Fly7dg01BBBCu90eGxtfVHQgMDof/Plh1arVaFSNJkYoFO/fX4RCKhDC6uoaqVQRHz+2vb2doqilS/OEQnF9/Q0IYVdX1/btvxYKxRwOLy1tYnb2gjlz5iYmJovFErk8TKOJUas1Gk10TEysRCKNiYn74YcrgbFdk8m0aVMBOk5CCO+nfVetWr148ZLRRjtR29LSUolEig7oUql87tx5Pp8PvXrppR0AYFu3boMQarVasViydOkyiqJob1yr1RYUvKBQKAmCyeUKFAqlRhMTHR2rVmvCw1USiUwgEM2YMauuro5eNdTz8ePH0Tke/QtQj1999ZVcHmY0GvtF44KJBXq93jlz5kqlchpJWdlFiqL0en1CQlJoqKyysgpCuHPn7wHAv/32WzpcQDPf0HDro4/+vGjR4oSEJJVKLZXKZbKwsWPH5eTk7t9fhIj1+8lAEXjrrT+icPN9NtCHxWIJD1e9+uprozqU0pX37SsUCETR0bEaTYxYLFm9eg2E8JNPPiUI5lNPLfT7/T09tsTE5NTUCejMQHNOCzoqJlPXjRs3qqouX7lyRa/XD8y3oD8aGxvff/+DwOeA/qegYLNAILx58+aoFB3psdVqTUlJUyojoqKikcbfvt2ck7MExxmff/4FhPDgwUMAYB9++NGgy4SSTIPmt/o9RxP74ou/abXawWH8+GOdQCDKycklySGTOsMQ8uabbyFC1GqNUhmRnb0gOjp23LhEk8lEkuT8+QvCw1VogYdZIyTS6DQ7VNauvb39yy/39usHBKJcv34DAPhf//o/KAc7Knt1586d6OgYRAVSUD5ftGPHf0EIr1y5ymZzt2598QEj3KjtkSNHdLq2ftYIBK5obW1tWFi4QqFE4cPgkaCazz+/EYV/0DaiVEbU1NRCCLdv/zWbzb16tfpBYKBJt7S0lJaWDuwH9MO6Y8fLPJ5g0qTJ9fX1o1X3srKLUqkiKipao4mRSKQ5ObkQQoPBoFZrsrKyHzCNiEL3paWl9F4xJAyKooxGY3z82KlTM9asWXvjhjYYJHa7/cKFMiTTWVnZoaGymJg4kSjk4MFDEMJPP/0MAFBUVDTaRRlIRXNzs8ViGXRzAwOV9eDBQzyeYMeOlwsKNtfW1g4jXag7o9GYmTkTWaRDh/4hFktUKnVqahpK0mVmzlSrNUaj8QEziW63G0Xdg0ooIyS/+tVWoVC0e/eebdu2l5ScHiZYT5fm5haKohwOx5QpT3A4PLQFXbp0icvlL1mS++DZ0OGvSYBBTZ7T6Zw9e45MJv/++9LXX3/js88+DwQ51O0F9PbDDz9iMFhNTU0Qwg0bnmcwWGvXPjcqgzHQfvT29g5fBwxl18xm8+OPT1UolDU1NV9++eXWrdvu3dMFpquHSk92dXWdOnUKQqjX62Ni4ng8QXz82NbW1mAoHbg6fr/farWOqFTDXbawWCyzZ8/hcHhnz54rL6946qmFAVmVESZEUZTb7fnhhx+2bNmmVmsSE5PPnz8/Ks8A8UDnhH8ODBqJ0+l49tnVKAthMpm2bNmWlZWNfOahwAz09tva2t55591Jk9KzsrK12pvDe5/0K71ebzAYgmQPBHMLadeu95hMzowZMxsbG7//vvTJJ+c+99z65uYWGsxQ0wr0/FwuV2Hh/paWlqE0lSQpkqQghL29fWVlZa2trcFLIAiS36qqy2lpExkM9uuvv9Hc3PznP/9l8uTHXnzx18hbDsQTDD8DjQqN9urVq0ePHqP1MEgYQd01RKEKt9u9e/eet99+hyDwbdu2PvHE42fOnL148WJiYuKKFSsyM6cHcysPnbwHvXRYV1dXVVUVEaHKypqHIgnBhziCvcBKH+R1Ot3HH+/Zu3cvjuP5+XkTJky4detWbW0tk8maPDl95swZY8f+B4/H5fF4gTGewFhEv6CO2Wyprq6uqamRSCTZ2dkaTVTgddN/y3ViFD0BAJhMpr179xYVHejo0KekpDz++BQej2cwGNrbO6xWS0FBwfLlTw/Vj9/vb2tra25u/vHHuoaGBoqiJk6cmJ09H90lHtX1259/uTsQjN/vv3TpUnHx4YqKSpfLFRUVpVJFiMXi9PR0tVotEPBJkjKbzU6nw+FwtLW1d3V1dXV1eb0+DANSqSw9PT09fWJSUhIKhAd5V/VhXrUPBIPwaLXa6uoarVbb3t7R22vzeLw4jrNYLDabRRCERCJRq9UqlSomJiYuLlatVgfmZR7wLuyDfvEBQkhREAAY5N3lgcHmgaryC8Dod5k38Msl9NLSsU06I/EQv31Cl/8FfLd+Ax7Gd/MAAAAASUVORK5CYII=';

var VP_INFO_ = {
  name: 'V & P TRADING',
  ssm: '202403282822 (AS0486263-A)',
  addr: '397, JLN LAMAN DELFINA 3/4, NILAI IMPIAN,\n71800 NILAI, NEGERI SEMBILAN.',
  tel: '010-797 1699',
  email: 'serom1699@gmail.com',
  bank: 'HONG LEONG BANK',
  bankName: 'V & P TRADING',
  bankAcc: '33300265188'
};

/** 中文等级 → 发票上的英文等级 */
function gradeEn_(desc) {
  var d = String(desc || '');
  if (d.indexOf('好') >= 0) return 'PREMIUM';
  if (d.indexOf('普通') >= 0) return 'NORMAL';
  return '';
}

/**
 * 由 SET 名 + 中文说明推出发票上的英文品名（不含品牌）。
 * 规则来自客户回传的样板：PROTON NORMAL GOODIES BAGS PACKAGE 10 ITEMS
 *   有 ITEMS  → {等级} GOODIES BAGS PACKAGE (N ITEMS)
 *   雨伞类    → {等级} UMBRELLA
 * 颜色不进发票 —— 颜色不影响价钱，是给司机拿货用的，
 * 同一样东西在帐上每次都该叫同一个名字。
 */
function invNameOf_(setType, desc) {
  var st = up_(setType).trim(), g = gradeEn_(desc);
  // 价目表还没填「普通/好的」的，一律当 NORMAL。
  // 宁可全部一致，也不要同一份发票上有的写等级有的不写 —— 那才会被会计问。
  if (!g) g = 'NORMAL';
  var m = st.match(/(\d+)\s*ITEMS/);
  if (m) return (g ? g + ' ' : '') + 'GOODIES BAGS PACKAGE (' + m[1] + ' ITEMS)';
  if (/UMBRELLA/.test(st)) {
    var tail = /\+\s*BAG/.test(st) ? 'UMBRELLA + BAG' : 'UMBRELLA';
    // SET 名字本身已经写了等级就不重复加（例如 PREMIUM UMBRELLA）
    if (st.indexOf('PREMIUM') >= 0) return 'PREMIUM ' + tail;
    return (g ? g + ' ' : '') + tail;
  }
  if (st === 'BAG' || st === 'BAGS') return 'BAG';
  return st;   // 没见过的自己写，之後可在价目表手改
}

/** 价目表里那一栏（可手改）；没填就用规则推。t 同样可以从外面传进来 */
function invNameFor_(setType, price, t) {
  t = t || ensureCols_('SET_PRICE', ['说明', '英文品名']);
  var hit = null;
  t.rows.forEach(function (r) {
    if (up_(r.SET_TYPE) === up_(setType) && toNum_(r.UNIT_PRICE) === toNum_(price)) hit = r;
  });
  var manual = hit ? String(hit['英文品名'] || '').trim() : '';
  if (manual) return manual;
  var desc = hit ? String(hit['说明'] || '').trim() : '';
  if (!desc) desc = DESC_FALLBACK_[up_(setType) + '|' + toNum_(price)] || '';
  return invNameOf_(setType, desc);
}

/** 把价目表里空白的「英文品名」一次填好（可重复跑） */
function fillInvNames() {
  var t = ensureCols_('SET_PRICE', ['说明', '英文品名']);
  var c = t.head.indexOf('英文品名') + 1;
  if (c <= 0) return { ok: false, msg: 'SET_PRICE 缺少「英文品名」栏' };
  var n = 0;
  t.rows.forEach(function (r) {
    if (!r.SET_TYPE) return;
    if (String(r['英文品名'] || '').trim()) return;
    var desc = String(r['说明'] || '').trim() ||
               (DESC_FALLBACK_[up_(r.SET_TYPE) + '|' + toNum_(r.UNIT_PRICE)] || '');
    t.sheet.getRange(r.__row, c).setValue(invNameOf_(r.SET_TYPE, desc));
    n++;
  });
  clearBootCache_();
  return { ok: true, filled: n };
}

/**
 * 发票上该印哪个牌子。
 * 同一个人可能一次买了不同牌子的货，而分行只代表他在哪里上班，
 * 所以分行的 BRAND 不一定等於货的牌子 —— 客户是把牌子写在备注里的。
 * 备注有写就以备注为准，没写才用分行的品牌。
 */
var INV_BRANDS_FALLBACK_ = ['PERODUA', 'PROTON', 'HONDA', 'CHERY', 'JAECOO',
                            'JETOUR', 'MITSUBISHI', 'EMAS', 'ICAUR'];
/** 认得的牌子直接从 BRANCH 名单长出来 ——
    他们以後开一间新品牌的分行，备注写那个牌子就自动认得，不必找我改程式。 */
function invBrandList_() {
  var out = {};
  INV_BRANDS_FALLBACK_.forEach(function (b) { out[b] = 1; });
  try {
    readTable_('BRANCH').rows.forEach(function (r) {
      var b = up_(r.BRAND);
      if (b && b !== 'OTHER') out[b] = 1;
    });
  } catch (e) { }
  // 长的排前面，免得 EMAS 抢在 E.MAS PUTRAJAYA 之类前面误判
  return Object.keys(out).sort(function (a, b) { return b.length - a.length; });
}
function invBrandOf_(r, list) {
  // 下单时直接选的最优先；其次是备注里写的（旧资料都是这样）；最後才是分行的品牌
  var picked = up_(r.INV_BRAND || '');
  if (picked) return picked;
  list = list || invBrandList_();

  // 备注优先。找不到再看「特别交代」——
  // 那格本来是写给司机的，但品牌选项以前收在「进阶」里，客户看不到，
  // 就顺手把品牌打在这格（Jaecoo Goodies bag 之类）。这里一并认，旧单不必重打。
  var texts = [up_(r.NOTE || ''), up_(r.DRV_NOTE || '')];
  for (var t = 0; t < texts.length; t++) {
    if (!texts[t]) continue;
    for (var i = 0; i < list.length; i++) {
      if (texts[t].indexOf(list[i]) >= 0) return list[i];
    }
  }

  var b = String(r.BRAND || '').trim();
  return (b && b !== 'OTHER') ? b : '';
}

/** 一个客户 = 销售员 ＋ 开给谁（本人 / 公司）。同月同客户 = 一张发票。 */
function custKey_(o, perOrder) {
  var base = up_(o.SALESMAN) + '|' + (up_(o.INVOICE_TO) === 'COMPANY' ? 'COMPANY' : 'SA');
  // 逐笔开单的客户：键後面挂订单号，一笔就是一张发票。
  // 这样 invNoFor_ 完全不用改 —— 它只比对 CUST_KEY 字串，
  // 「同月同客户重印不变号」「作废不重用号码」两条保证自动继承。
  return perOrder ? base + '#' + String(o.ORDER_ID || '').trim() : base;
}

/** 从扩充键拆出订单号；月结的键没有 '#'，回传空字串。 */
function keyOrderId_(key) {
  var i = String(key).indexOf('#');
  return i < 0 ? '' : String(key).slice(i + 1).trim();
}

/**
 * 这个销售员是「月结」还是「逐笔」。
 * 存在 SALESMAN 分页的「发票方式」栏，留空 = 月结（现状，绝大多数）。
 */
function perOrderSet_(t) {
  t = t || ensureCols_('SALESMAN', ['发票抬头', '公司名', '地址', '电话', '发票方式']);
  var set = {};
  t.rows.forEach(function (r) {
    var v = String(r['发票方式'] || '').trim();
    // 键是「销售员＋分行」—— 同一个人可能这间分行要逐笔、另一间照样月结
    if (/逐笔|逐筆|PER.?ORDER|EACH/i.test(v)) set[poKey_(r.SALESMAN, r.BRANCH)] = 1;
  });
  return set;
}
function poKey_(salesman, branch) { return up_(salesman) + '@' + up_(branch); }

/** 销售员那几栏发票资料。
    t 可以从外面传进来 —— 不传的话每呼叫一次就重读整张 SALESMAN 表，
    一个月三十个客户就是三十次全表读取，那正是开发票清单要跑 30 秒的原因。

    branch：这张发票是哪一间分行的生意。一个人跑几间分行，SALESMAN 表里
    就有几行；只比对名字会永远拿到最前面那一行，发票上就印错车行。
    传了分行就优先拿那一行，没传 / 找不到才退回第一行（旧行为）。 */
function billTo_(salesman, mode, t, branch) {
  t = t || ensureCols_('SALESMAN', ['发票抬头', '公司名', '地址', '电话']);
  var hit = null, exact = null;
  t.rows.forEach(function (r) {
    if (up_(r.SALESMAN) !== up_(salesman)) return;
    if (!hit) hit = r;
    if (branch && !exact && up_(r.BRANCH) === up_(branch)) exact = r;
  });
  if (exact) hit = exact;
  var company = hit ? String(hit['公司名'] || '').trim() : '';
  var title = hit ? String(hit['发票抬头'] || '').trim() : '';
  var nm = (mode === 'COMPANY' ? (company || title) : (title || String(salesman))) || String(salesman);
  // 抬头和公司名填一样时只印一次，不要在发票上连印两行同样的名字
  var co = (mode === 'COMPANY') ? '' : company;
  if (co && up_(co) === up_(nm)) co = '';
  return {
    name: nm,
    company: co,
    addr: hit ? String(hit['地址'] || '').trim() : '',
    tel: hit ? String(hit['电话'] || '').trim() : '',
    // 订单说是哪一间就印哪一间 —— 名单里查不到那一行也照印，订单才是事实
    branch: String(branch || (hit ? hit.BRANCH : '') || '').trim()
  };
}

/** 发票号码：INV-YYMM-NNN，存在 INVOICE 分页，重印不变 */
function invSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('INVOICE');
  var head = ['INV_NO', 'YM', 'CUST_KEY', 'SALESMAN', 'BILL_MODE', 'BILL_NAME',
              'ORDERS', 'AMOUNT', 'ISSUED_AT', 'ISSUED_BY',
              'STATUS', 'VOID_AT', 'VOID_BY', 'VOID_REASON'];
  if (!sh) {
    sh = ss.insertSheet('INVOICE');
    sh.getRange(1, 1, 1, head.length)
      .setValues([['● 已开出的发票号码。重印同一张，号码不变。'
        + '开错了用「作废」—— 号码保留、标记作废，不要删除，帐上的序号才不会断。',
        '', '', '', '', '', '', '', '', '', '', '', '', '']]);
    sh.getRange(2, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(2);
  }
  return ensureCols_('INVOICE', ['STATUS', 'VOID_AT', 'VOID_BY', 'VOID_REASON']);
}
function invVoided_(r) { return up_(r.STATUS) === 'VOID'; }

function invNoFor_(ym, key, meta) {
  var t = invSheet_();
  var hit = null;
  t.rows.forEach(function (r) {
    // 作废掉的不算 —— 同一个客户同一个月可以重开，但会拿到新号码
    if (invVoided_(r)) return;
    if (String(r.YM) === String(ym) && up_(r.CUST_KEY) === up_(key)) hit = r;
  });
  if (hit) return { no: String(hit.INV_NO), isNew: false, row: hit.__row };

  // 找最大号时「连作废的一起算」，号码才永远不会重复使用
  var max = 0;
  t.rows.forEach(function (r) {
    if (String(r.YM) !== String(ym)) return;
    var m = String(r.INV_NO || '').match(/-(\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  var no = 'INV-' + ym + '-' + String(max + 1).padStart(3, '0');
  t.sheet.appendRow(t.head.map(function (h) {
    return h === 'INV_NO' ? no : h === 'YM' ? ym : h === 'CUST_KEY' ? key
      : h === 'SALESMAN' ? (meta.salesman || '') : h === 'BILL_MODE' ? (meta.mode || '')
        : h === 'BILL_NAME' ? (meta.billName || '') : h === 'ORDERS' ? (meta.n || 0)
          : h === 'AMOUNT' ? (meta.amount || 0)
            : h === 'ISSUED_AT' ? Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm')
              : h === 'ISSUED_BY' ? (meta.by || '') : '';
  }));
  return { no: no, isNew: true };
}

/** 某个月有哪些客户可以开发票 */
function listInvoiceMonth(p) {
  var ym = String((p && p.ym) || '').trim();          // 例 '2607'
  if (!/^\d{4}$/.test(ym)) return { ok: false, msg: '请选月份' };
  var yy = '20' + ym.slice(0, 2), mm = parseInt(ym.slice(2), 10);

  var t = ensureCols_('ORDERS', ['INVOICE_TO', 'INV_BRAND']);
  var sp = ensureCols_('SET_PRICE', ['说明', '英文品名']);
  var smT = ensureCols_('SALESMAN', ['发票抬头', '公司名', '地址', '电话', '发票方式']);   // 只读一次
  var perOrder = perOrderSet_(smT);
  var graded = {};
  sp.rows.forEach(function (r) {
    if (!r.SET_TYPE) return;
    var k = up_(r.SET_TYPE) + '|' + toNum_(r.UNIT_PRICE);
    if (String(r['说明'] || '').trim() || String(r['英文品名'] || '').trim()) graded[k] = 1;
  });
  // 先看这个月已经开出去哪些发票 —— 决定谁能走逐笔要用到
  var issued = {}, voided = {};
  invSheet_().rows.forEach(function (r) {
    if (String(r.YM) !== ym) return;
    if (invVoided_(r)) { voided[up_(r.CUST_KEY)] = (voided[up_(r.CUST_KEY)] || 0) + 1; return; }
    issued[up_(r.CUST_KEY)] = String(r.INV_NO);
  });

  var needMap = {};
  var g = {};
  t.rows.forEach(function (r) {
    if (isVoid_(r)) return;
    // DATE 在表里是日期物件不是字串，一定要经过 fmtDate_
    var d = fmtDate_(r.DATE);
    if (d.slice(0, 4) !== yy || parseInt(String(r.MONTH), 10) !== mm) return;
    var gk = up_(r.SET_TYPE) + '|' + toNum_(r.UNIT_PRICE);
    if (!graded[gk] && !DESC_FALLBACK_[gk] && !/^BAGS?$/.test(up_(r.SET_TYPE)))
      needMap[String(r.SET_TYPE).trim() + ' RM ' + toNum_(r.UNIT_PRICE)] = 1;
    var mode = up_(r.INVOICE_TO) === 'COMPANY' ? 'COMPANY' : 'SA';
    // 这个月如果已经用「月结」开过发票了，就维持月结 ——
    // 不然同一笔生意会再拿一个新号码，变成开了两张。已经开出去的不动。
    var po = !!perOrder[poKey_(r.SALESMAN, r.BRANCH)] && !issued[up_(custKey_(r))];
    var key = custKey_(r, po);
    if (!g[key]) {
      var b = billTo_(r.SALESMAN, mode, smT, r.BRANCH);
      g[key] = { key: key, salesman: String(r.SALESMAN || ''), mode: mode,
                 billName: b.name, brs: {},
                 hasAddr: !!b.addr, n: 0, amount: 0,
                 perOrder: po,
                 orderId: po ? String(r.ORDER_ID || '').trim() : '',
                 date: po ? fmtDate_(r.DATE) : '' };
    }
    // 月结可能跨分行 —— 清单上全部列出来，不要只显示碰巧第一笔那间
    var bn = String(r.BRANCH || '').trim();
    if (bn) g[key].brs[bn] = 1;
    g[key].n++; g[key].amount += toNum_(r.TOTAL_INCOME);
  });

  var need = Object.keys(needMap).sort();

  var list = Object.keys(g).map(function (k) {
    g[k].amount = Math.round(g[k].amount * 100) / 100;
    g[k].invNo = issued[up_(k)] || '';
    g[k].voided = voided[up_(k)] || 0;
    var brs = Object.keys(g[k].brs);
    g[k].branch = brs.join(' / ');
    g[k].nBranch = brs.length;
    delete g[k].brs;
    return g[k];
  }).sort(function (a, b) { return b.amount - a.amount; });

  return { ok: true, ym: ym, list: list, needGrade: need };
}

/** 组一张发票的完整资料（不配号码，纯预览用 preview=true） */
function getInvoice(p) {
  var ym = String((p && p.ym) || '').trim();
  var key = String((p && p.key) || '').trim();
  if (!/^\d{4}$/.test(ym) || !key) return { ok: false, msg: '参数不对' };
  var yy = '20' + ym.slice(0, 2), mm = parseInt(ym.slice(2), 10);
  var wantOrder = keyOrderId_(key);                       // 逐笔才有值
  var baseKey = wantOrder ? key.slice(0, key.indexOf('#')) : key;
  var mode = baseKey.split('|')[1] === 'COMPANY' ? 'COMPANY' : 'SA';
  var salesman = baseKey.split('|')[0];

  var t = ensureCols_('ORDERS', ['INVOICE_TO', 'INV_BRAND']);
  var spT = ensureCols_('SET_PRICE', ['说明', '英文品名']);                    // 只读一次
  var brandList = invBrandList_();                                            // 只算一次
  var smT = ensureCols_('SALESMAN', ['发票抬头', '公司名', '地址', '电话', '发票方式']);   // 只读一次
  var perOrder = perOrderSet_(smT);
  var lines = [], sub = 0, qty = 0, gross = 0;
  t.rows.forEach(function (r) {
    if (isVoid_(r) || custKey_(r) !== baseKey) return;
    // 逐笔：只收这一笔订单
    if (wantOrder && String(r.ORDER_ID || '').trim() !== wantOrder) return;
    // DATE 在表里是日期物件不是字串，一定要经过 fmtDate_
    var d = fmtDate_(r.DATE);
    if (d.slice(0, 4) !== yy || parseInt(String(r.MONTH), 10) !== mm) return;
    var q = toNum_(r.QTY), pr = toNum_(r.UNIT_PRICE), tot = toNum_(r.TOTAL_INCOME);
    var brand = invBrandOf_(r, brandList);
    var nm = invNameFor_(r.SET_TYPE, pr, spT);
    lines.push({
      date: shortDate_(d),
      iso: d,
      branch: String(r.BRANCH || '').trim(),
      desc: (brand ? brand + ' ' : '') + nm,
      qty: q, unit: unitOf_(r.SET_TYPE), price: pr,
      amount: Math.round(q * pr * 100) / 100,
      total: tot
    });
    qty += q; gross += q * pr; sub += tot;
  });
  if (!lines.length) return { ok: false, msg: '这个月这位客户没有订单' };

  lines.sort(function (a, b) { return a.iso < b.iso ? -1 : a.iso > b.iso ? 1 : 0; });  // 用 ISO 排，别用 dd.mm.yy

  // 月结可能一个人跨几间分行。段落顺序照「各分行第一笔的日期」，
  // 整张单读起来还是照时间走，不会变成按字母跳。
  var brOrder = {}, brList = [];
  lines.forEach(function (l) {
    if (l.branch && brOrder[l.branch] === undefined) {
      brOrder[l.branch] = brList.length; brList.push(l.branch);
    }
  });
  if (brList.length > 1) {
    lines.forEach(function (l, i) { l.__i = i; });                  // 同分行内保持日期序
    lines.sort(function (a, b) {
      var x = brOrder[a.branch] === undefined ? 9999 : brOrder[a.branch];
      var y = brOrder[b.branch] === undefined ? 9999 : brOrder[b.branch];
      return x !== y ? x - y : a.__i - b.__i;
    });
    lines.forEach(function (l) { delete l.__i; });
  }

  gross = Math.round(gross * 100) / 100;
  sub = Math.round(sub * 100) / 100;
  var disc = Math.round((gross - sub) * 100) / 100;

  // 单一分行：抬头那块印这间分行。跨分行：抬头不印，改成内文分段列出来。
  var b = billTo_(salesman, mode, smT, brList.length === 1 ? brList[0] : '');
  if (brList.length > 1) b.branch = '';
  return {
    ok: true, ym: ym, key: key, mode: mode, salesman: salesman,
    billTo: b, lines: lines, qty: qty, branches: brList,
    gross: gross, discount: disc, total: sub,
    company: VP_INFO_,
    dateStr: shortDate_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'))
  };
}

function esc_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
/** 长地址会把表格撑到超出纸宽，右边的金额栏就被挤掉。
    旧版渲染器只在空白处断行，所以逗号 / 斜线後面补一个空格。 */
function wrapAddr_(s) {
  return esc_(s).replace(/([,\/])(?=\S)/g, '$1 ').replace(/\n/g, '<br>');
}
function money_(n) {
  var v = Math.round(toNum_(n) * 100) / 100;
  var s = v.toFixed(2), p = s.split('.');
  return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
}

/**
 * 发票 HTML。刻意用表格与 inline style —— Apps Script 转 PDF 用的是
 * 旧版渲染器，不支援 flex / grid，但表格排版它处理得很好。
 */
function invoiceHtml_(iv, invNo) {
  var C = iv.company, b = iv.billTo;

  /* 一个人一个月跑了几间分行时，明细照分行分段，每段一个小计。
     只有跨分行才分段 —— 单一分行（绝大多数）渲染出来跟以前一模一样。
     小计加的是 Amount 那一栏（数量 × 单价），客户拿计算机加那一栏必须对得上；
     折扣不分摊进各分行，分摊会产生分不尽的角位，反而被质疑。 */
  var multi = (iv.branches || []).length > 1;
  var cell = 'padding:6px 4px;border-bottom:1px solid #e4e2dd;font-size:10pt';
  var brHead = function (br) {
    return '<tr><td colspan="6" style="padding:12px 4px 4px;font-size:9.5pt;font-weight:bold;'
      + 'letter-spacing:.5px;border-bottom:1px solid #cfccc5">' + esc_(br || '—') + '</td></tr>';
  };
  var brSub = function (br, amt) {
    return '<tr>'
      + '<td colspan="5" align="right" style="padding:5px 8px 2px;font-size:9.5pt;color:#555">'
      + esc_(br || '—') + ' 小计</td>'
      + '<td align="right" style="padding:5px 4px 2px;font-size:9.5pt;font-weight:bold">'
      + money_(amt) + '</td></tr>';
  };
  var rows = '', curBr = null, brSum = 0;
  iv.lines.forEach(function (l, i) {
    if (multi && l.branch !== curBr) {
      if (curBr !== null) rows += brSub(curBr, brSum);
      curBr = l.branch; brSum = 0;
      rows += brHead(l.branch);
    }
    brSum += l.amount;
    rows += '<tr>'
      + '<td style="' + cell + '">' + (i + 1) + '.</td>'          // 编号连续跑，查帐的人数行数才对得上
      + '<td style="' + cell + '">' + esc_(l.date) + '</td>'
      + '<td style="' + cell + '">' + esc_(l.desc) + '</td>'
      + '<td align="right" style="' + cell + '">' + l.qty + ' ' + l.unit + '</td>'
      + '<td align="right" style="' + cell + '">' + money_(l.price) + '</td>'
      + '<td align="right" style="' + cell + '">' + money_(l.amount) + '</td>'
      + '</tr>';
  });
  if (multi && curBr !== null) rows += brSub(curBr, brSum);

  var totRow = function (label, val, bold) {
    return '<tr>'
      + '<td align="right" style="padding:4px 8px;font-size:' + (bold ? '12pt' : '10pt') + ';'
      + (bold ? 'font-weight:bold;border-top:1.5px solid #333;' : '') + '">' + esc_(label) + '</td>'
      + '<td align="right" style="padding:4px 4px;font-size:' + (bold ? '12pt' : '10pt') + ';width:110px;'
      + (bold ? 'font-weight:bold;border-top:1.5px solid #333;' : '') + '">' + val + '</td>'
      + '</tr>';
  };

  return '<html><head><meta charset="utf-8"></head>'
    + '<body style="font-family:Arial,Helvetica,sans-serif;color:#1c1c1e;margin:0;padding:26px 30px">'

    /* ── 抬头 ── */
    + '<div style="text-align:center;margin-bottom:6px">'
    + '<img src="' + VP_LOGO_ + '"><br>'   // 图已是 66px，不指定宽高：旧渲染器只会裁不会缩
    + '<div style="font-size:20pt;font-weight:bold;letter-spacing:2px">' + esc_(C.name) + '</div>'
    + '<div style="font-size:8.5pt;color:#555;margin-top:3px">' + esc_(C.ssm) + '</div>'
    + '<div style="font-size:9pt;color:#333;margin-top:4px">' + esc_(C.addr).replace(/\n/g, '<br>') + '</div>'
    + '<div style="font-size:9pt;color:#333;margin-top:2px">Tel: ' + esc_(C.tel) + ' &nbsp;·&nbsp; ' + esc_(C.email) + '</div>'
    + '</div>'
    + '<hr style="border:none;border-top:2px solid #1c1c1e;margin:12px 0 14px">'

    /* ── 客户 ＋ 发票资讯 ── */
    + '<table width="100%" cellpadding="0" cellspacing="0"><tr>'
    + '<td valign="top" width="52%">'
    + '<div style="font-size:8.5pt;color:#777;text-decoration:underline;margin-bottom:4px">Customer Details</div>'
    + '<div style="font-size:11pt;font-weight:bold">' + esc_(b.name) + '</div>'
    + (b.company ? '<div style="font-size:10pt">' + esc_(b.company) + '</div>' : '')
    + (b.branch ? '<div style="font-size:9.5pt;color:#555">' + esc_(b.branch) + '</div>' : '')
    + (b.addr ? '<div style="font-size:9pt;color:#333;margin-top:3px">' + wrapAddr_(b.addr) + '</div>' : '')
    + (b.tel ? '<div style="font-size:9.5pt;color:#333">Tel: ' + esc_(b.tel) + '</div>' : '')
    + '</td>'
    + '<td valign="top" width="48%">'
    + '<div style="font-size:17pt;font-weight:bold;letter-spacing:1px;margin-bottom:8px">INVOICE</div>'
    + '<table cellpadding="0" cellspacing="0" style="font-size:9.5pt">'
    + '<tr><td style="color:#777;padding-right:8px">Invoice No</td><td style="font-weight:bold">' + esc_(invNo) + '</td></tr>'
    + '<tr><td style="color:#777;padding-right:8px">Date</td><td>' + esc_(iv.dateStr) + '</td></tr>'
    + '</table></td></tr></table>'

    /* ── 明细 ── */
    + '<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:16px;table-layout:fixed">'
    + '<tr style="background:#f4f2ec">'
    + '<th align="left" style="padding:7px 4px;font-size:9pt;width:26px">#</th>'
    + '<th align="left" style="padding:7px 4px;font-size:9pt;width:76px">Date</th>'
    + '<th align="left" style="padding:7px 4px;font-size:9pt">Description</th>'
    + '<th align="right" style="padding:7px 4px;font-size:9pt;width:64px">Qty</th>'
    + '<th align="right" style="padding:7px 4px;font-size:9pt;width:66px">Price</th>'
    + '<th align="right" style="padding:7px 4px;font-size:9pt;width:82px">Amount</th>'
    + '</tr>' + rows + '</table>'

    /* ── 合计 ── */
    + '<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:10px"><tr>'
    + '<td width="55%" valign="top">'
    + '<div style="font-size:9pt;color:#555">Total Qty: <b>' + iv.qty + '</b></div>'
    + '</td><td width="45%">'
    + '<table width="100%" cellpadding="0" cellspacing="0">'
    + totRow('Sub-Total (RM)', money_(iv.gross))
    + (iv.discount > 0.005 ? totRow('Discount (RM)', '-' + money_(iv.discount)) : '')
    + totRow('Grand Total (RM)', money_(iv.total), true)
    + '</table></td></tr></table>'

    /* ── 银行 ＋ 结尾 ── */
    + '<div style="margin-top:26px;padding-top:10px;border-top:1px solid #d8d6d0">'
    + '<div style="font-size:8.5pt;color:#777;margin-bottom:3px">BANK DETAIL</div>'
    + '<div style="font-size:10pt">' + esc_(C.bank) + ' &nbsp;·&nbsp; ' + esc_(C.bankName) + ' &nbsp;·&nbsp; <b>' + esc_(C.bankAcc) + '</b></div>'
    + '</div>'
    + '<div style="text-align:center;margin-top:28px;font-size:10pt;color:#555">Thank you for your business.</div>'
    + '</body></html>';
}

/** 产生发票 PDF，回传 base64 让前端下载。会配一个号码并记下来。 */
function makeInvoicePdf(p) {
  var iv = getInvoice(p);
  if (!iv.ok) return iv;

  var r = withLock_(function () {
    return { ok: true, inv: invNoFor_(iv.ym, iv.key, {
      salesman: iv.salesman, mode: iv.mode, billName: iv.billTo.name,
      n: iv.lines.length, amount: iv.total, by: (p && p.by) || ''
    }) };
  });
  if (!r.ok) return r;

  var invNo = r.inv.no;
  try {
    var html = invoiceHtml_(iv, invNo);
    var pdf = Utilities.newBlob(html, MimeType.HTML, invNo + '.html').getAs(MimeType.PDF);
    return {
      ok: true, invNo: invNo, isNew: r.inv.isNew,
      filename: invNo + ' ' + iv.billTo.name.replace(/[\\/:*?"<>|]/g, '') + '.pdf',
      b64: Utilities.base64Encode(pdf.getBytes()),
      total: iv.total, lines: iv.lines.length
    };
  } catch (e) {
    return { ok: false, msg: '产生 PDF 失败：' + friendlyErr_(e), invNo: invNo };
  }
}

/** 填销售员的发票资料（抬头 / 公司 / 地址 / 电话）
    一个人跑几间分行，是可以各自留不同抬头/地址的（例如车行地址 vs 个人地址）——
    预设只写选中那一间；p.sameAll 才广播到全部分行（旧行为，一样可以选）。
    发票方式一律只跟着分行走，跟 sameAll 无关。 */
function saveBillTo(p) {
  var name = normName_(p && p.salesman);
  if (!name) return { ok: false, msg: '请选销售员' };
  var branch = String(p.branch || '').trim();
  var t = ensureCols_('SALESMAN', ['发票抬头', '公司名', '地址', '电话', '发票方式']);
  var rows = t.rows.filter(function (r) { return normName_(r.SALESMAN) === name; });
  if (!rows.length) return { ok: false, msg: '名单里没有这个人' };

  var map = { '发票抬头': (p.title || '').trim(), '公司名': (p.company || '').trim(),
              '地址': (p.addr || '').trim(), '电话': (p.tel || '').trim() };
  // 发票方式只写「这一间分行」那行 —— 同一个人别间分行不受影响
  var brRows = branch
    ? rows.filter(function (r) { return up_(r.BRANCH) === up_(branch); })
    : [];
  // 只有一间分行的人没有「哪一间」的问题，永远写那一行
  var sameAll = rows.length <= 1 || !!p.sameAll;
  var target = sameAll ? rows : brRows;
  if (!target.length) return { ok: false, msg: '请先选要改哪一间分行' };

  var res = withLock_(function () {
    Object.keys(map).forEach(function (k) {
      var c = t.head.indexOf(k) + 1;
      if (c > 0) target.forEach(function (r) { t.sheet.getRange(r.__row, c).setValue(map[k]); });
    });
    if (p.perOrder !== undefined && brRows.length) {
      var cm = t.head.indexOf('发票方式') + 1;
      if (cm > 0) brRows.forEach(function (r) {
        t.sheet.getRange(r.__row, cm).setValue(p.perOrder ? '逐笔' : '');
      });
    }
    return { ok: true };
  });
  if (!res.ok) return res;
  clearBootCache_();
  return { ok: true, salesman: name };
}

/**
 * 客户名单：分行、销售员、地址、电话，全部从 SALESMAN 表现读。
 *
 * 不放进 bootstrap —— 三百多个地址会让每次开 app 都多背几十 KB。
 * 改成点开名单才呼叫，读一次三百行很快，而且他们改完地址马上就看得到。
 */
function getNameList() {
  var t = ensureCols_('SALESMAN', ['发票抬头', '公司名', '地址', '电话', '发票方式']);
  var out = [];
  t.rows.forEach(function (r) {
    var n = String(r.SALESMAN || '').trim();
    if (!n) return;
    out.push({
      name: n,
      branch: String(r.BRANCH || '').trim(),
      brand: String(r.BRAND || '').trim(),
      state: String(r.STATE || '').trim(),
      region: String(r.REGION || '').trim(),
      n: toNum_(r['历史笔数']),
      primary: String(r['主要分行'] || '').indexOf('★') >= 0,
      title: String(r['发票抬头'] || '').trim(),
      company: String(r['公司名'] || '').trim(),
      addr: String(r['地址'] || '').trim(),
      tel: String(r['电话'] || '').trim(),
      perOrder: /逐笔|逐筆|PER.?ORDER|EACH/i.test(String(r['发票方式'] || '').trim())
    });
  });
  return { ok: true, rows: out };
}

function getBillTo(p) {
  var name = normName_(p && p.salesman);
  var branch = String((p && p.branch) || '').trim();
  var t = ensureCols_('SALESMAN', ['发票抬头', '公司名', '地址', '电话', '发票方式']);
  var hit = null, brHit = null, branches = [];
  var isPo = function (r) {
    return /逐笔|逐筆|PER.?ORDER|EACH/i.test(String(r['发票方式'] || '').trim());
  };
  t.rows.forEach(function (r) {
    if (normName_(r.SALESMAN) !== name) return;
    if (!hit) hit = r;
    // 抬头／地址是「分行」的 —— 车行地址跟个人地址可以不一样，
    // 每间分行带自己的一份，前端切分行的时候才知道要换哪四个栏位
    if (branch && !brHit && up_(r.BRANCH) === up_(branch)) brHit = r;
    var bn = String(r.BRANCH || '').trim();
    if (bn) branches.push({
      branch: bn, perOrder: isPo(r),
      title: String(r['发票抬头'] || ''), company: String(r['公司名'] || ''),
      addr: String(r['地址'] || ''), tel: String(r['电话'] || '')
    });
  });
  if (!hit) return { ok: false, msg: '名单里没有这个人' };
  // 跨分行的人，前端要让他们自己挑要改哪一间 —— 不能从卡片猜，会改错间
  var row = brHit || (branches.length === 1 ? hit : null);
  var src = row || hit;
  return { ok: true, salesman: name,
    branches: branches,
    branch: row ? String(row.BRANCH || '') : '',
    perOrder: row ? isPo(row) : false,
    title: String(src['发票抬头'] || ''), company: String(src['公司名'] || ''),
    addr: String(src['地址'] || ''), tel: String(src['电话'] || '') };
}

/**
 * 作废一张发票号码。
 * 不删除 —— 号码留在纪录里标记作废，序号才不会断。
 * 查帐的人看到 001 作废、002 正常，是合理的；看到 001 凭空不见，会问。
 */
function voidInvoice(p) {
  var no = String((p && p.invNo) || '').trim();
  var reason = String((p && p.reason) || '').trim();
  if (!no) return { ok: false, msg: '请选一张发票' };
  if (reason.length < 2) return { ok: false, msg: '请写一句作废原因（查帐时要看的）' };

  var t = invSheet_();
  var hit = null;
  t.rows.forEach(function (r) { if (up_(r.INV_NO) === up_(no)) hit = r; });
  if (!hit) return { ok: false, msg: '找不到这张发票' };
  if (invVoided_(hit)) return { ok: false, msg: '这张已经作废过了' };

  var cS = t.head.indexOf('STATUS') + 1,
      cA = t.head.indexOf('VOID_AT') + 1,
      cB = t.head.indexOf('VOID_BY') + 1,
      cR = t.head.indexOf('VOID_REASON') + 1;
  if (cS <= 0) return { ok: false, msg: 'INVOICE 表缺少 STATUS 栏' };

  var r = withLock_(function () {
    t.sheet.getRange(hit.__row, cS).setValue('VOID');
    if (cA > 0) t.sheet.getRange(hit.__row, cA).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'));
    if (cB > 0) t.sheet.getRange(hit.__row, cB).setValue((p && p.by) || '');
    if (cR > 0) t.sheet.getRange(hit.__row, cR).setValue(reason);
    return { ok: true };
  });
  if (!r.ok) return r;

  logChange_({ kind: '发票', action: '作废', from: no, to: reason,
               by: (p && p.by) || '' });
  return { ok: true, invNo: no };
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
  '10 ITEMS|50': '好的雨伞',
  '9 ITEMS|35': '普通雨伞',
  '8 ITEMS|55': '好的雨伞',
  '9 ITEMS|40': '好的雨伞',
  '9 ITEMS|38': '普通雨伞',
  '10 ITEMS|55': '好的雨伞',
  'UMBRELLA|12': '普通雨伞',
  'UMBRELLA|13': '普通雨伞',
  '10 ITEMS|45': '普通雨伞',
  'UMBRELLA|20': '好的雨伞',
  'UMBRELLA|26': '普通雨伞',
  'UMBRELLA + BAG|16': '普通雨伞',
  'UMBRELLA|14': '普通雨伞',
  'UMBRELLA|25': '好的雨伞',
  '9 ITEMS|36': '普通雨伞',
  '8 ITEMS|26': '普通雨伞',
  '9 ITEMS|12': '普通雨伞',
  'BLACK UMBRELLA|14': '普通雨伞',
  'RED UMBRELLA|12': '普通雨伞',
  '8 ITEMS|60': '好的雨伞',
  'PREMIUM UMBRELLA|20': '好的雨伞',
  '8 ITEMS|50': '好的雨伞',
  'UMBRELLA|11': '普通雨伞',
  '10 ITEMS|40': '好的雨伞'
};

var ORDER_HEAD = ['ORDER_ID', 'DATE', 'MONTH', 'REGION', 'STATE', 'BRANCH', 'BRAND', 'SALESMAN',
  'SET_TYPE', 'UNIT_PRICE', 'QTY', 'TOTAL_INCOME', 'MY_INCOME', 'DRIVER_FEE',
  'PAY_STATUS', 'PAY_DATE', 'DRV_COLOR', 'DRV_GRADE', 'DRV_NOTE', 'INVOICE_TO', 'INV_BRAND',
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
