# 日股復盤儀表板

GitHub Pages 每日產生東証 17 業種板塊地圖，WordPress 以固定 iframe 嵌入。更新不需要本機開機或 WordPress 密碼。

## 行情與更新

- Yahoo Finance 提供日線歷史。個別股票缺少應有交易日或歷史不足時，改抓株探（Kabutan）的公開日線表；再失敗則嘗試 Yahoo Japan 公開歷史表。實測株探在 GitHub 主機可能回覆 405，因此雲端備援必須同時具備 Yahoo Japan 路徑，不能只依賴株探。
- 以東京交易所交易日曆判斷最近完整收盤日，收盤後保留 20 分鐘結算時間。週末與日本休市日不要求不存在的行情。
- 每檔必須有最近 21 個交易日，日經指數必須有最近 2 日；價格與成交量須有效。所有有效成分股與指數都通過才產生新圖表。不同來源價格基準不一致時，不拼接舊歷史。
- `market_cache.json.gz` 保留實際日期的歷史資料，不補造價格、不將舊日期改成新日期。
- JPX 官方下市清單會自動核對；已生效的下市股票移出計算，尚未生效的保留。下市紀錄保存於 `listing_status.json`，不自動把舊公司的行情接到新代碼。
- 每日 UTC 07:23、09:23、12:23、15:23、23:23 排程（台灣 15:23、17:23、20:23、23:23、翌日 07:23），每次失敗最多嘗試 3 次。GitHub 排程可能延遲，並非即時串流。
- 更新失敗仍發布 `health.json`，保留最後有效圖表並標示異常；工作流程保持失敗狀態供 GitHub 通知。頁面每 5 分鐘檢查更新狀態，超過 30 小時未核對也會警示。
- `health.json` 可檢查應有交易日、每檔實際日期、來源、失敗原因及排除清單。來源計數包含指數／ETF。

## 本地執行與驗證

```bash
pip install -r requirements.txt
python -m pytest -q
python generate_jp.py
```

可用 `JP_FORCE_BACKUP='7203.T,6758.T,^N225' python generate_jp.py` 驗證備援，或在 GitHub Actions 手動執行勾選 `force_backup`。強制備援若未實際成功，也會視為失敗。

完整更新時間取決於行情來源、市值快取與新聞服務。市值快取每 30 日更新；AI 摘要需有效 `ANTHROPIC_API_KEY`，與行情更新分開處理。

## 計算與維護

主要板塊漲跌使用市值加權，並提供等權對照。放量突破為成交量達前 20 日均量 2 倍且收盤突破前 20 日高點。低檔放量使用可取得的歷史低點，歷史長度依来源而異。資料為盤後日線，不是盤中即時報價。

`sectors.json` 管理成分股；`market_data.py` 處理來源、日期與上市狀態；`generate_jp.py` 產生儀表板；`.github/workflows/daily.yml` 負責測試、更新及部署。WordPress 固定 iframe 指向 https://tetsu811.github.io/jp-daily-recap/ 。
