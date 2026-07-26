# gas/ — 线上 App 的唯一真实来源

这个资料夹里的档案 **就是** 线上跑的 V&P TRADING App。改这里 → push 到 `main` → GitHub Actions 自动推上 Google Apps Script 并更新部署，网址不变。

| 档案 | 内容 |
|---|---|
| `Code.js` | 后端：登入、下单、报表、作废/编辑、帐号管理、一次性安装器 `INSTALL()` |
| `Index.html` | 整个 App 前端（HTML + CSS + JS 单档） |
| `appsscript.json` | Apps Script manifest（时区 Asia/Singapore、V8、Web App 设定） |
| `.clasp.json` | 指向哪个 Apps Script 专案（scriptId） |

## ⚠️ repo 里其他同名档案是旧的，别改

根目录的 `Code.gs`、`ImportHistory.gs`、`apps-script/`、`web/src/` 是**之前一次失败的尝试**，跟线上完全无关。改它们不会有任何效果。以这个 `gas/` 资料夹为准。

`web/public/` 则是短网址 `v-p-trading.vercel.app` 的包装页（一个全萤幕 iframe），跟 App 逻辑无关。

## 本机怎么改

```bash
npm install -g @google/clasp@2.4.2
clasp login                    # 只需一次
cd gas
clasp pull                     # 先拉，确认跟线上一致（git diff 应该是空的）
# ...改 Code.js / Index.html...
git commit -am "改了什么" && git push
# Actions 会自动 push + deploy
```

**永远先 `clasp pull` 再改。** 如果有人直接在 Apps Script 网页编辑器改过，不先 pull 就 push 会把他的改动盖掉。

## 紧急回滚

Apps Script 编辑器 → `Deploy` → `Manage deployments` → 铅笔 → Version 选旧的版本 → Deploy。
或者 git revert 那个 commit 再 push，Actions 会自动部署回去。
