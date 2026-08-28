# open-sheet

<img src=".github/assets/banner-light.png" alt="open-sheet — 一張畫在試算表網格上的依賴圖：兩格匯入一格" width="100%">

[![npm](https://img.shields.io/npm/v/@open-sheet/core?style=flat&label=%40open-sheet%2Fcore)](https://www.npmjs.com/package/@open-sheet/core)
[![CI](https://github.com/lianghsun/open-sheet/actions/workflows/ci.yml/badge.svg)](https://github.com/lianghsun/open-sheet/actions/workflows/ci.yml)
[![open-sheet.dev](https://img.shields.io/badge/open--sheet.dev-0b1020?style=flat&logo=cloudflare&logoColor=white)](https://open-sheet.dev)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat)](https://opensource.org/licenses/MIT)

[English](README.md) · **繁體中文**

**[open-sheet.dev](https://open-sheet.dev)** —— 它是什麼，以及你交出去的那份檔案能做什麼。

**專為 agent 打造的試算表框架。** 用自然語言描述你要的模型 — 你的 coding agent 負責寫 React，open-sheet 負責儲存格位址、公式參考、重新計算與匯出。

如果 [open-slide](https://github.com/1weiho/open-slide) 是 agent 的 Google Slides、[open-doc](https://github.com/simonliu-ai-product/open-doc) 是 Google Docs，那 open-sheet 就是 **Google Sheets**。同樣的想法、不同的媒材 — 但媒材改變了問題的本質。

> 簡報是像素。文件是印在紙上的像素。**試算表是一張依賴圖。**
> 一份烤成靜態數字的活頁簿，只是一張表格的照片。試算表的價值在於：收到的人可以改一個假設，然後看它重新算出來。

```bash
npx @open-sheet/cli init my-sheets
```

## 「這不就是 Claude for Excel 在做的事嗎?」

不同的工具,不同的工作。

> **Claude for Excel 讓一個人更快地做出一份試算表。**
> **open-sheet 讓那份試算表不需要有人再做第二次。**

| | Claude for Excel | open-sheet |
| --- | --- | --- |
| 輸入 | 一份已經存在的活頁簿 | 一份 `.tsx` 原始碼 |
| 誰在迴圈裡 | 人,每一次 | 人**一次**,在 review 的時候 |
| 產出 | 那份檔案,改過了 | 同一份檔案,每次重建都一樣 |
| 要 review 一個改動 | 一份二進位 diff | 一次 code review |
| 要做 500 次 | 500 個 session | 一個迴圈 |
| 需要 | Excel、帳號、付費方案 | Node。MIT。 |

在試算表**裡面**工作的助理,寫進儲存格的仍然是 `=SUM(B2:B13)`。那就是一個手寫的
位址 — 由模型手寫的,但仍然是手寫的。插入一列它就錯了,而且不會有任何東西告訴你。
這不是模型夠不夠聰明的問題:**A1 是那個媒材唯一的語言。**

open-sheet 寫的是 `ref('pl').column('revenue')`,在編譯期才解析。這不是「比較會寫
公式」— 是根本不寫位址。

而且算不出來的時候,它會說算不出來。`#NOT_EVALUATED`,絕不給一個看起來合理的數字。
任何直接往活的儲存格裡寫東西的工具,不管有沒有把握,產出的都是一個「長得像數字的
東西」—— 那是代價最高、又最看不出來的一種錯。

**Claude for Excel 贏的地方,而且差距不小。** 它寫進去的也是活公式 — 這裡的差別
從來不是「我們的會重算、它們的不會」。一份不是你產生的活頁簿 — 我們根本讀不了。它用 Excel 自己的引擎,所以它的數字天生就是 Excel 的數字,而我們得靠跨引擎
比對去掙來這件事。五百多個函式對我們的一百多個。樞紐分析表。「這個變異是什麼造成
的?」— 那是分析師對一個活模型的問題,不是編譯器的問題。

如果你手上有一份試算表想要幫忙,用那個。這裡處理的是另一件事:當試算表是**管線的
產出**而不是有人打開來的文件 — 每月的董事會資料包、五百張請款單、資料一進來就重新
產生的模型。

兩者是可以接起來的,而且順序是這樣:open-sheet 產出活公式,正是為了讓收到的人可以
打開它、對它提問。

*沒有人因為 Word 有 AI 就說 LaTeX 不該存在。*

## 為什麼要做

Agent 很會寫分析，但很不會做試算表，而原因非常具體：**`=SUM(B2:B13)`**。

儲存格位址是 agent 唯一抓不住的東西。它會把標題列算錯。它會忘記資料變多之後合計列已經往下移了。它寫了一個指向 `B7` 的參考，然後在上面插入一季 — 於是整個檔案裡的公式全部悄悄錯掉。不是壞掉，是**算錯**，而那比壞掉更糟。

所以 open-sheet 把位址收走了。你永遠不會寫到位址：

```tsx
col('grossProfit', {
  header: '毛利',
  formula: (r) => sub(r.cell('revenue'), r.cell('cogs')),
})
```

框架擁有每一個座標。在資料陣列裡加一列，所有參考自動重新解析。**一份 open-sheet 原始檔裡，不會出現任何一個 A1 位址。**

## 三者的譜系

三個框架各自吸收了「該媒材裡 agent 最做不好的那件苦工」：

| | open-slide | open-doc | **open-sheet** |
| --- | --- | --- | --- |
| 媒材 | 1920 × 1080 畫布 | A4 紙張 | 儲存格網格 |
| 吸收掉的苦工 | 縮放、導覽、簡報模式 | 分頁、目錄、頁碼 | **儲存格位址、公式參考、重新計算** |
| 你拿到的 | 一份能上台的簡報 | 一份印得出來的 PDF | **一個財務長可以改的模型** |

## 特色

### 🔢 參考，而不是位址

`r.cell('revenue')` · `r.prev().cell('revenue')` · `ref('pl').column('revenue')` · `ref('pl').total('revenue')` · `ref('assumptions').get('growth')`

這些要等到最後、等版面決定好每個東西的位置之後，才會被解析成 A1。純量假設還會額外輸出成 **Excel defined name**，所以匯出的活頁簿讀起來是 `=B4*growth` — 打開它的人看得懂，不是只有機器看得懂。

### 📐 自動排版

`<Stack>` 與 `<Row>` 會把區塊擺上網格且互不碰撞。你描述**順序**，框架決定**座標**。這是 open-sheet 對應 open-doc `flow()` 的那一塊。

### 🧮 一棵公式樹，兩個後端

每個公式都是一棵 expression AST，有兩個消費者：`serialize()` 把 Excel 公式字串寫進 `.xlsx`，`evaluate()` 算出 viewer 要顯示的數字。兩者在 CI 裡互相驗證 — 做法是**用 LibreOffice 重新計算匯出的活頁簿再比對差異**。這同時證明了兩件事：匯出的確實是活公式，而且我們的求值後端跟真正的試算表引擎意見一致。

算不出來的公式會顯示 `#NOT_EVALUATED`，絕不編一個數字出來。

### 📤 活的 `.xlsx`，以及 `.csv` / `.html` / `.pdf`

`.xlsx` 裡放的是**公式，不是烤好的值** — 含數字格式、defined names、凍結窗格、條件式格式。同一份樣式模型同時餵給 Excel writer 與 HTML/PDF renderer，所以印出來的報表跟活頁簿長得一樣。

### 🤖 為 agent 而生

Scaffolder 會一起帶上 skills：`/create-sheet`、`/sheet-authoring`、`/current-sheet`、`/apply-comments`。MCP server（`open-sheet dev --mcp`）讓任何 agent framework 都能驅動它。Inspect 模式讓你點一個儲存格就看到它的原始碼行號、解析後的公式、以及計算結果 — 或是留一則筆記給你的 agent。

## 目前狀態

**早期開發中。** 尚未發佈到 npm。進度請追蹤 [milestones](https://github.com/lianghsun/open-sheet/milestones)。

上面描述的都已經做完並有測試 —— 編譯器、參考系統與公式引擎、viewer 與 dev server、
skills、MCP server、inspect 模式、themes、design panel、原生圖表，以及四種匯出格式。
144 個測試，其中兩個會實際驅動真正的試算表應用程式。

已發佈 `0.1.0`：

```bash
npx @open-sheet/cli init my-sheets
cd my-sheets && npm install && npm run dev
```

要證明的那個瞬間，現在端到端成立：`apps/demo` 匯出的活頁簿裡，淨利欄位就是
`=F6*(1-taxRate)`，而測試會去改 `taxRate` 並斷言整欄跟著動 —— 做算術的是
LibreOffice，不是 open-sheet。


發佈前想先自己測：[TESTING.md](TESTING.md)。

## Repo 結構

pnpm + Turbo monorepo。

| 路徑 | 說明 |
| --- | --- |
| `packages/core` | `@open-sheet/core` — 編譯器、排版、參考系統、公式引擎、viewer、Vite plugin，以及 `open-sheet` CLI。 |
| `packages/cli` | `@open-sheet/cli` — `npx @open-sheet/cli init` 腳手架與專案範本。 |
| `packages/mcp` | `@open-sheet/mcp` — 走 Streamable HTTP 的 MCP server。 |
| `apps/demo` | 透過 `workspace:*` 使用 `@open-sheet/core` 的範例工作區。自用 dogfood 目標。 |

## 開發

```bash
pnpm install
pnpm dev        # 用本地的 @open-sheet/core 跑 demo
pnpm build      # 建置所有套件
pnpm typecheck  # 跨整張圖跑 tsc
pnpm check      # biome（格式化 + lint + 整理 import）
pnpm test       # vitest
```

## 致謝

open-sheet 沿著 [@1weiho](https://github.com/1weiho) 的 [open-slide](https://github.com/1weiho/open-slide) 與 [@simonliu-ai-product](https://github.com/simonliu-ai-product) 的 [open-doc](https://github.com/simonliu-ai-product/open-doc) 而來 — virtual-module 內容探索、腳手架、以及 skills-as-documentation 的做法都是他們的。這是第三個媒材。

## 授權

MIT
