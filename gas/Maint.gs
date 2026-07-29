/*******************************************************
 * 一次性维护脚本 —— 2026-07-29
 * 客户第二轮回覆确认後的资料整理。
 * 走的是跟 App 按钮完全一样的函式，差别只在：
 * 每一步先算预期值，对不上就立刻中止，绝不硬做。
 * 跑完可以整个档案删掉。
 *******************************************************/

function RUN_MAINT_20260729() {
  var log = [], abort = null;
  function say(s) { log.push(s); }

  function snapshot() {
    var t = readTable_('ORDERS'), inc = 0, fee = 0, mine = 0, n = 0;
    t.rows.forEach(function (r) {
      if (isVoid_(r)) return;
      n++; inc += toNum_(r.TOTAL_INCOME); fee += toNum_(r.DRIVER_FEE); mine += toNum_(r.MY_INCOME);
    });
    return { n: n, inc: Math.round(inc * 100) / 100,
             fee: Math.round(fee * 100) / 100, mine: Math.round(mine * 100) / 100 };
  }

  var before = snapshot();
  say('【开始前】订单 ' + before.n + ' 笔｜营业额 RM ' + before.inc +
      '｜利润 RM ' + before.mine + '｜司机抽成 RM ' + before.fee);

  /* ── 步骤 1：升级付款资料（补收款日期） ── */
  var m = migratePayment();
  if (!m.ok) { abort = '步骤1 升级失败：' + m.msg; }
  else say('[1] 升级付款资料 → ' + m.msg);

  /* ── 步骤 2：VP00401 改去 EMAS PERLING ── */
  if (!abort) {
    var pv = previewBranch({ id: 'VP00401', branch: 'EMAS PERLING' });
    if (!pv.ok) abort = '步骤2 预览失败：' + pv.msg;
    else if (up_(pv.from.branch) === 'EMAS PERLING') { say('[2] VP00401 已经是 EMAS PERLING，跳过'); pv = null; }
    else if (up_(pv.from.branch) !== 'EMAS') abort = '步骤2 中止：VP00401 现在的分行是「' + pv.from.branch + '」，既不是 EMAS 也不是 EMAS PERLING，资料跟预期不符';
    else if (pv && Math.abs(pv.diff) > 0.005) abort = '步骤2 中止：抽成会变动 RM ' + pv.diff + '，预期应为 0';
    else if (pv) {
      var cb = changeBranch({ id: 'VP00401', branch: 'EMAS PERLING', alsoSalesman: true });
      if (!cb.ok) abort = '步骤2 执行失败：' + cb.msg;
      else say('[2] VP00401 ' + pv.from.branch + ' → EMAS PERLING｜抽成 ' + cb.oldFee +
               ' → ' + cb.newFee + '（不变）｜销售员 ' + cb.salesman +
               (cb.salesmanMoved ? ' 名单已跟着搬' : ' 名单未动'));
    }
  }

  /* ── 步骤 3：七组合并 / 改名 ── */
  // [从, 到, 预期笔数, 是不是改名]
  // EMAS → EMAS PUTRAJAYA 不在这里：底下还有两笔柔佛单归属未定，
  // 合并会把柔佛营业额算进雪兰莪，也会给司机他没送过的抽成。留给客户自己处理。
  var PLAN = [
    ['PROTON TAMAN MIDAS',      'PROTON TAMAN MIDAH', 1, true],
    ['PROTON MIDAH',            'PROTON TAMAN MIDAH', 2, false],
    ['PERODUA NILAI IMPIAN',    'PERODUA NILAI',      7, false],
    ['JAECOO SEREMBAN',         'JAECOO SEREMBAN 2',  2, false],
    ['PROTON IOI PUTRAJAYA',    'PROTON PUTRAJAYA',   1, false],
    ['PROTON PUTRAJAYA - EMAS', 'EMAS PUTRAJAYA',     1, false]
  ];

  for (var i = 0; i < PLAN.length && !abort; i++) {
    var from = PLAN[i][0], to = PLAN[i][1], expN = PLAN[i][2], expRename = PLAN[i][3];
    var tag = '[3.' + (i + 1) + '] ' + from + ' → ' + to;

    var p = previewMergeBranch({ from: from, to: to });
    if (!p.ok) { abort = tag + ' 预览失败：' + p.msg; break; }
    if (p.orders !== expN) { abort = tag + ' 中止：预期 ' + expN + ' 笔，实际 ' + p.orders + ' 笔'; break; }
    if (Math.abs(p.feeDiff) > 0.005) { abort = tag + ' 中止：司机抽成会变动 RM ' + p.feeDiff + '，预期应为 0'; break; }
    if (!!p.rename !== expRename) { abort = tag + ' 中止：预期' + (expRename ? '改名' : '合并') + '，实际是' + (p.rename ? '改名' : '合并'); break; }

    var r = mergeBranch({ from: from, to: to, by: 'MAINT-20260729' });
    if (!r.ok) { abort = tag + ' 执行失败：' + r.msg; break; }
    say(tag + '｜' + (r.rename ? '改名' : '合并') + ' ' + r.orders + ' 笔｜销售员 ' +
        r.salesmen + ' 行｜抽成变动 RM ' + r.feeDiff);
  }

  /* ── 收尾对帐 ── */
  var after = snapshot();
  say('【结束後】订单 ' + after.n + ' 笔｜营业额 RM ' + after.inc +
      '｜利润 RM ' + after.mine + '｜司机抽成 RM ' + after.fee);

  var same = (before.n === after.n) &&
             Math.abs(before.inc - after.inc) < 0.005 &&
             Math.abs(before.mine - after.mine) < 0.005 &&
             Math.abs(before.fee - after.fee) < 0.005;
  say(same ? '✓ 对帐通过：笔数、营业额、利润、司机抽成四项完全没变'
           : '✗ 对帐不符！差额 → 笔数 ' + (after.n - before.n) +
             '｜营业额 ' + Math.round((after.inc - before.inc) * 100) / 100 +
             '｜利润 ' + Math.round((after.mine - before.mine) * 100) / 100 +
             '｜抽成 ' + Math.round((after.fee - before.fee) * 100) / 100);

  if (abort) say('■ 已中止：' + abort + '（中止之前的步骤已经生效，之後的都没做）');

  // 留给客户手动处理的部分，列出来
  var left = readTable_('ORDERS').rows.filter(function (r) {
    return up_(r.BRANCH) === 'EMAS' && !isVoid_(r);
  });
  if (left.length) {
    say('── 以下留给客户自己处理（没有动） ──');
    left.forEach(function (r) {
      say('　' + r.ORDER_ID + '｜' + r.DATE + '｜' + r.STATE + '/' + r.REGION + '｜' +
          r.SALESMAN + '｜' + r.SET_TYPE + ' ' + r.QTY + '×' + r.UNIT_PRICE +
          '｜RM ' + r.TOTAL_INCOME);
    });
    say('　→ 柔佛那几笔要先各自归到正确分行，「EMAS」才可以并去 EMAS PUTRAJAYA。');
  }

  var out = log.join('\n');
  Logger.log(out);
  return out;
}
