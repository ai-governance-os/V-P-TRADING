# 订单 & 抽佣记录 App — 项目交接说明

给客户做的一个小工具：录单自动算钱、查找记录、月度汇总、一键 WhatsApp 通知司机。不用 Supabase，全部跑在 Google Sheet + Apps Script 上，Drive 生态内就能多人共用。

## 这是什么

客户原本用 Google Drive 里的表格手动记订单、算 uncle jun（送货员）的抽佣，常常要一个个查、一个个输入，很烦。我们做了一个网页 App 代替手动操作。

**核心设计:以「名字」为中心**（v2 重做）。客户的真实烦恼是——记一单时还要先想这个 salesman 属于哪个 branch、哪个品牌、这个 set 卖多少钱。新版做到:**只要输名字,其它全自动带出**。

- 录单：**打名字**(例如打 "sya" 就弹出 SYAWAN)→ 选中后自动带出 branch / 品牌 / state / region → 点一下 set 类型(8/9/10 items / umbrella)自动填入该 branch 的默认价格(可改)→ 填数量,total 和司机抽佣实时算好 → 保存 + 一键 WhatsApp。找不到的名字可当场到「管理」页新增。
- 首页仪表盘：本月总收入 / 本月给司机抽佣 / 本月订单数 大数字卡片,加最近订单列表,一进 App 就看到经营概况。
- 查找：关键字(名字或 branch)+ 月份 + 地区 筛选。
- 月度汇总：自动统计每月 total income 和 uncle jun 总抽佣,按 branch 细分。
- 管理：随时新增 salesman / branch,改抽佣规则,改 branch 默认价格,填司机 WhatsApp 号码。
- WhatsApp 一键通知：录单页按一下,自动打开 WhatsApp、详情已填好,司机按发送即可。

界面改成手机 App 风格:蓝色顶栏、底部 5 个 tab、圆角卡片、头像式列表。

抽佣规则：

| 类别 | 算法 | 核对状态 |
|---|---|---|
| KL / Selangor（标准 set） | SET × RM2.50 | ✅ 已用 Uncle 表核对 |
| Umbrella（不分地区） | SET × RM1 | ✅ 已用 Uncle 表核对 |
| Seremban / NS（标准 set） | SET × Price × 10% | ⚠️ README 原值,部署前请再核对 |
| Johor（标准 set） | 暂按 SET × RM2.50 | ⚠️ Excel 无对应数据,部署前必须与客户确认 |

规则全部可在 App「管理」页里改数字,不用动代码。

## 目前进度

代码已经写完(v2 名字优先版),抽佣算法已用真实数据逐条验证(KL RM2.50/set、umbrella RM1/set、NS 10%、Johor 默认 都跑通)。**已把客户全年 Excel 的真实名单灌进种子数据:285 位 salesman、59 个 branch、3 个地区(KL/Selangor、Seremban/NS、Johor),外加 100 多个 branch 的默认价格**——不再是旧版手打的 24 个。

**还没有实际部署到 Google Sheet 上**,客户现在还不能打开链接 —— 部署要在能登入客户 Google 账号的电脑上完成。

## 交付的文件（都在这个文件夹里）

- `Code.gs` — 后台逻辑：存数据、自动算钱、查找、月度汇总、新增 branch/salesman、抽佣规则。
- `index.html` — 手机 / 电脑都能开的网页界面。
- `部署指南.docx` — 详细部署步骤（新建 Sheet → 粘贴代码 → 初始化 → 部署成网页 App → 分享链接 → 设置司机号码）。

## 接下来要做什么（换电脑后按顺序来）

1. 打开 `部署指南.docx`，跟着五个步骤做：
   - 新建一个 Google Sheet
   - 打开「扩展程序 → Apps Script」，把 `Code.gs` 和 `index.html` 的内容分别粘贴进去（index.html 要新建一个 HTML 文件命名为 `index`）
   - 运行一次 `setupSpreadsheet` 函数做初始化（会自动建好 Orders / Config / Rates 三个分页，并且把客户现有的 branch / salesman 名单和抽佣规则灌进去）
   - 部署成网页应用（Execute as: 我；访问权限建议选 Google 账号持有者）
   - 把生成的网址发给 2-3 个合伙人，并在「管理」分页填入 uncle jun 的 WhatsApp 号码
2. 部署完之后，实际点开 App 走一遍流程：录一笔单、查找、看月度汇总、按一次「通知司机」确认 WhatsApp 内容对不对。
3. 如果客户现有的 branch / salesman 名单还有遗漏（旧表格里可能不止我们看过的那几个），部署后直接在「管理」页面补上就行，不用改代码。
4. 确认没问题后，把网址和用法教给客户和另外 2 个合伙人。

## 待客户确认或可能要调整的点

- 部署时权限选「Google 账号持有者」还是「知道链接的任何人」，看客户对安全性的要求。
- 如果以后要真正做到「连按发送都不用」（全自动推送 WhatsApp），需要申请 WhatsApp Business API + 模板审核，目前用的是免费的 wa.me 一键方式（差一步按发送）。
- App 支持「新增」和「删除」订单(删除在查找页,每条右侧 🗑️,有确认框)。目前还没有「修改已有订单」——要改还是回 Google Sheet 的 Orders 分页手动改,如客户需要可后续加。
