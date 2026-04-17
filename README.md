# 日股復盤儀表板

每日自動產生東証 17 業種板塊地圖 + 證據瓦片(放量突破/漲跌幅 TOP3/底部放量),嵌入 WordPress 用。

## 結構

```
generate_jp.py        # 主腳本:抓 yfinance 資料 → 算指標 → 輸出 HTML
sectors.json          # 東証 17 業種 × 代表個股清單(可擴充)
jp_dashboard.html     # 產生的成品(GitHub Pages 服務)
.github/workflows/    # 每日 06:30 UTC (= 15:30 JST) 自動跑
requirements.txt
```

## 本地跑

```bash
pip install -r requirements.txt
python generate_jp.py
open jp_dashboard.html
```

約 30-60 秒(yfinance 批次下載 ~340 檔)。

## 板塊計算規則

- **板塊漲跌** = 該業種成分股「等權重」當日漲幅平均
- **放量突破** = 成交量 ≥ 20 日均量 2 倍 **且** 收盤創 20 日新高
- **底部放量** = 放量(同上)**且** 價位距 120 日低點 10% 以內
- **漲/跌幅 TOP3** = 純當日漲跌幅排序

## 擴充成分股

直接編輯 `sectors.json`,格式:
```json
"業種名": [
  {"code": "1234", "name": "公司名"}
]
```
`code` 為 TSE 4 位數,腳本會自動加 `.T` 後綴。指數類用 `^` 前綴(如 `^N225`)或 ETF 代號(如 `1306.T`)。

## WordPress 嵌入(自動)

每天 workflow 跑完會把 `jp_dashboard_embed.html`(iframe srcdoc 隔離版)透過 REST API 推到 WP page。

**第一次設定:**
1. WP Admin → Users → Profile → 滑到底 Application Passwords
   - Name: `jp-recap` → Add New → 複製 24 位密碼
2. GitHub repo → Settings → Secrets and variables → Actions → New repository secret
   - `WP_URL` = `https://tetsu811.com`
   - `WP_USER` = WP 用戶名
   - `WP_PASS` = 上面複製的 24 位
3. 手動跑一次 workflow(Actions → Daily JP Market Recap → Run workflow)
   - 會建立一個 **draft** page,印出 page id
4. 把 page id 設為 secret:
   - `WP_PAGE_ID` = 上一步的 id
5. 之後每天會 update 同一頁。預設 `publish`,要改 draft 就在 Variables 加 `WP_STATUS=draft`

**手動本地推:**
```bash
WP_URL=https://tetsu811.com WP_USER=xxx WP_PASS='xxxx xxxx ...' python wp_publish.py
```

**WPCode 替代方案:**
若想用 WPCode plugin,把 `jp_dashboard_embed.html` 內容貼到一個新 HTML snippet,用 shortcode 在文章插入。但這樣每天要手動貼 — 不如走 REST API 自動。

## 待辦/可改進

- [ ] 業績歸因瓦片(串 TDnet 決算速報)
- [ ] 新聞歸因瓦片(日経 RSS / Reuters JP 關鍵字頻次)
- [ ] 板塊權重從等權改成市值加權
- [ ] 自動從 JPX 月度檔抓 TOPIX 500 全名單(取代手動 sectors.json)
- [ ] 板塊歷史趨勢(過去 N 日板塊輪動)
