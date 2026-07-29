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
