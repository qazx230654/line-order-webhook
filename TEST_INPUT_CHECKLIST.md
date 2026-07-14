# LINE 滷味訂單助手測試表

測試日期：____________　測試人員：____________　Cloud Run 修訂版本：____________　Apps Script 部署版本：____________

## 測試前準備

- 先備份 `Orders`、`Customers`、`PriceList`、`ItemAlias`、`UnmatchedItems` 工作表（若存在）。
- 確認 Cloud Run 已設定 `GAS_URL`、`GAS_SHARED_SECRET`、`LINE_CHANNEL_ACCESS_TOKEN`、`LINE_CHANNEL_SECRET`。
- 確認 Apps Script 已設定 `WEBHOOK_SHARED_SECRET`、`OPENAI_API_KEY`。
- 測試修改功能時，請在建立原訂單後 20 分鐘內執行。
- 測試缺貨與未定價後，記得恢復 `PriceList`。

狀態填寫建議：`PASS`、`FAIL`、`BLOCKED`、`NOT RUN`

## A. 部署與安全

| ID | 測試項目 | 操作／輸入 | 預期結果 | 實際結果 | 狀態 | 備註 |
|---|---|---|---|---|---|---|
| A01 | Cloud Run 健康檢查 | 瀏覽器開啟 Cloud Run URL | 顯示 `OK` |  | NOT RUN |  |
| A02 | LINE Webhook Verify | LINE Developers → Messaging API → Verify | 驗證成功；Cloud Run 回傳 200 |  | NOT RUN |  |
| A03 | 錯誤簽章 | 使用錯誤的 `x-line-signature` POST Cloud Run | 回傳 `401 INVALID_SIGNATURE`，不新增訂單 |  | NOT RUN |  |
| A04 | Apps Script 錯誤 secret | 使用錯誤 `webhookSecret` 直接 POST Apps Script | 回傳 `{"error":"UNAUTHORIZED"}`，不新增訂單 |  | NOT RUN |  |

## B. 基本訂單輸入

| ID | 測試項目 | LINE 輸入 | 預期 LINE 回覆 | Orders／其他檢查 | 實際結果 | 狀態 | 備註 |
|---|---|---|---|---|---|---|---|
| B01 | 兩字商品 | `雞排` | 收到訂單確認，不可被忽略 | 建立一張訂單；商品為雞排 |  | NOT RUN |  |
| B02 | 一般多品項 | `高麗菜一份\n杏鮑菇兩份\n微辣\n18:00取餐` | 顯示兩項商品、微辣、預計時間及單號 | 同一 `systemOrderId` 寫入兩列；狀態待製作 |  | NOT RUN |  |
| B03 | 顧客資料 | `王小姐\n雞胸肉一份\n電話0912345678\n18:30取餐` | 顯示訂單成立 | customer、phone、pickup_time 正確 |  | NOT RUN |  |
| B04 | 半份數量 | `高麗菜半份\n百頁豆腐一份半` | 數量分別顯示 0.5、1.5 | quantity 分別為 0.5、1.5；總價正確 |  | NOT RUN |  |
| B05 | 分袋訂單 | `第一袋：高麗菜一份，微辣\n第二袋：雞胸肉一份，不辣，不酸菜` | 顯示兩個分組及各自備註 | group_name、note 分開保存 |  | NOT RUN |  |
| B06 | 商品別名 | 使用 `ItemAlias` 內的一個別名，例如：`杏包菇一份` | 回覆正規商品名稱 | Orders 商品名稱為別名對應的正式名稱 |  | NOT RUN | 依實際 ItemAlias 調整輸入 |
| B07 | 直接金額格式 | `百頁豆腐20` | 顯示品項且金額計算為 20 元 | quantity 與 `unit_price` 符合直接金額規則 |  | NOT RUN |  |
| B08 | 無指定取餐時間 | `高麗菜一份` | 訂單成立並顯示系統預估時間 | I 欄 pickup_time 寫入 `yyyy/MM/dd HH:mm` 預估值，不可空白 |  | NOT RUN |  |
| B09 | 空格與全半形正規化 | 將正式名稱加入不規則空格或使用全形字元 | 回覆 `PriceList` 內的正式名稱 | Orders 保存正式名稱 |  | NOT RUN |  |
| B10 | 常見字形正規化 | 正式名稱為「大豆干」時輸入 `大豆乾一份` | 回覆正式名稱「大豆干」 | 不新增 UnmatchedItems |  | NOT RUN |  |
| B11 | AI 限定菜單對應 | 輸入未登記於 ItemAlias、但可明確對應菜單的俗稱 | AI 只能回覆 PriceList 內的正式名稱 | Orders 保存正式名稱；價格可正常計算 |  | NOT RUN |  |
| B12 | 未知商品保守處理 | `神秘丸一份`（確認菜單及別名均不存在） | 訂單不中斷並顯示未設定價格，不可擅自換成其他品項 | Orders 保存原名稱且 N 欄為 `UNPRICED`；UnmatchedItems 新增 pending 記錄 |  | NOT RUN |  |
| B13 | 核准未知名稱 | 將 B12 的 suggested_name 改成有效正式名稱並將 status 設為 `approved`，再送相同內容 | 回覆核准後的正式名稱 | 不再新增未知記錄；Orders 使用正式名稱及其價格 |  | NOT RUN |  |
| B14 | 品項數量估時 | 在沒有其他待製作訂單時輸入 `高麗菜半份\n豆干兩份` | 系統等待時間採 10 分鐘保底 | 3 項 × 0.5 分鐘 + 1.5 分鐘加熱打包 = 3 分鐘，低於 10 分鐘所以採 10 分鐘 |  | NOT RUN |  |
| B15 | 大量品項估時 | 準備待製作內容與本次訂單合計超過 20 個品項單位 | 預計取餐時間依總品項單位增加 | 每單位 0.5 分鐘，再加 1.5 分鐘加熱打包，最終分鐘數無條件進位 |  | NOT RUN |  |
| B16 | 人工別名優先 | PriceList 同時有王子麵與蒸煮麵，ItemAlias 設定 A 欄科學麵、B 欄王子麵，再輸入 `科學麵一份` | 回覆王子麵，不可被 AI 猜成蒸煮麵 | Orders item_name 為王子麵；raw_name 保留科學麵 |  | NOT RUN |  |

## C. 價格與缺貨

| ID | 測試項目 | 前置設定 | LINE 輸入 | 預期結果 | Sheet 檢查 | 實際結果 | 狀態 |
|---|---|---|---|---|---|---|---|
| C01 | 成交價快照 | 確認高麗菜有價格 | `高麗菜兩份` | 回覆金額正確 | N 欄 `unit_price` 保存下單當下單價 |  | NOT RUN |
| C02 | 歷史價格不變 | C01 完成後修改 PriceList 的高麗菜價格 | 重新整理 Dashboard | C01 訂單營收不應改變 | Dashboard 使用 N 欄成交價 |  | NOT RUN |
| C03 | 未設定價格 | 暫時清空 PriceList 的雞心價格 | `雞心一份` | 顯示部分品項未設定價格，不發生錯誤 | N 欄為 `UNPRICED`；營收不計此品項 |  | NOT RUN |
| C04 | 部分未設定 | 高麗菜有價格、雞心無價格 | `高麗菜一份\n雞心一份` | 顯示雞心未設定價格 | 高麗菜計入營收；雞心排除 |  | NOT RUN |
| C05 | 缺貨品項 | 將雞胸肉設為已售完 | `雞胸肉一份` | 回覆雞胸肉已售完，要求重送 | Orders 不新增資料 |  | NOT RUN |
| C06 | 缺貨切換速度 | 在缺貨管理點擊高麗菜 | 畫面立即切換並顯示儲存中 | 成功後按鈕解除鎖定；PriceList 狀態正確 |  | NOT RUN |

## D. 新增與修改訂單

| ID | 測試項目 | 操作順序 | 預期結果 | Sheet 檢查 | 實際結果 | 狀態 | 備註 |
|---|---|---|---|---|---|---|---|
| D01 | 20 分鐘內自動修改 | 先送 `高麗菜一份`，再送 `雞排一份` | 第二次視為修改；沿用原顯示單號 | 新 UUID；舊 UUID 資料移除；最終只有雞排 |  | NOT RUN | 不需要輸入「修改」 |
| D02 | 20 分鐘內完整修改 | 先送 `高麗菜一份`，再送 `高麗菜兩份\n微辣` | 回覆修改成功；沿用原顯示單號 | 新 UUID；舊 UUID 資料移除；數量變 2 |  | NOT RUN |  |
| D03 | 20 分鐘內追加 | 先送 `高麗菜一份`，再送 `高麗菜一份\n雞排一份` | 第二次視為完整替換 | 最終訂單同時包含高麗菜與雞排 |  | NOT RUN | 必須重送完整內容 |
| D04 | 已完成後再次點餐 | 建立並完成訂單，再送 `雞排兩份` | 建立一張新訂單 | 已完成訂單不變；新訂單使用新顯示單號 |  | NOT RUN |  |
| D05 | 超過 20 分鐘再次點餐 | 建立訂單，超過 20 分鐘後送 `雞排兩份` | 建立一張新訂單 | 原訂單不變；新增不同 UUID 與顯示單號 |  | NOT RUN | 可在測試資料調整 created_at |
| D06 | 對話追加分組 | 先送 `大豆干兩份，小辣`，再送 `再一份一樣的，但不要辣` | 沿用原顯示單號，回覆兩個分組 | 最終同一張訂單有第一份小辣、第二份不辣；raw_message 保留兩次輸入 |  | NOT RUN | 第二次不可建立另一張訂單 |
| D07 | 指定分組修改 | 接續 D06 再送 `第二份再加一份甜不辣` | 只在第二份加入甜不辣 | 第一份內容不變；第二份包含原品項與甜不辣 |  | NOT RUN |  |
| D08 | 非訂單對話不修改 | 建立待製作訂單後送 `謝謝` | LINE 不回覆 | Orders 內容及 raw_message 完全不變 |  | NOT RUN | AI 可看見前文，但只能依本次訊息判斷 is_order |
| D09 | 備註修改保留取餐時間 | 建立系統預計 15:39 取餐的訂單後，依序送 `要加熱 不酸菜`、`第二份也不要酸菜` | 每次回覆都維持原時間 | Orders 的 pickup_time 不被 AI 改寫；O 欄為 estimated |  | NOT RUN |  |
| D10 | 明確修改取餐時間 | 接續 D09 傳送 `改成晚上六點半取餐` | 回覆取餐時間 18:30 | Orders pickup_time 更新為當日 18:30 |  | NOT RUN |  |
| D11 | 預估單追加品項延後 | 建立 pickup_time_source 為 estimated 的訂單，再追加 4 個品項單位 | 原預計時間延後 2 分鐘 | 只增加 `4 × 0.5` 分鐘並無條件進位；不重複加 1.5 分鐘 |  | NOT RUN |  |
| D12 | 指定時間追加不延後 | 建立客人指定 17:30、O 欄 requested 的訂單，再追加多個品項 | 取餐時間仍為 17:30 | 只有客人明確提供新時間才修改 |  | NOT RUN |  |
| D13 | 連續訊息依序處理 | 不等待回覆，快速傳送 `再一份...`，立刻再傳 `第二份也不要酸菜` | 先完成新增第二份，再將不酸菜加到該第二份 | 最終只有同一張訂單、兩個正確 groups；不產生競爭版本 |  | NOT RUN | 同一 LINE userId 會排隊，不同客人可同時處理 |

## E. 去重、完成與統計

| ID | 測試項目 | 操作／輸入 | 預期結果 | Sheet／Dashboard 檢查 | 實際結果 | 狀態 | 備註 |
|---|---|---|---|---|---|---|---|
| E01 | Webhook event 去重 | 使用相同 `webhookEventId` 將同一訂單送入 Apps Script 兩次 | 第二次回傳 `{"duplicate":true}` | Orders 只建立一張；WebhookEvents 只有一筆 event ID |  | NOT RUN |  |
| E02 | 去重失敗重試 | 模擬第一次寫入失敗後再次傳送相同 event ID | 第二次可重新處理 | 失敗時 event ID 已撤銷，不會永久卡住 |  | NOT RUN | 建議在測試部署執行 |
| E03 | 唯一 UUID 完成訂單 | 準備兩張顯示單號相同但 UUID 不同的訂單，完成其中一張 | 只完成指定 UUID | 另一張狀態不變；不發 LINE 完成通知 |  | NOT RUN |  |
| E04 | 無 LINE ID 的手動訂單 | 手動建立訂單後呼叫完成函式 | 訂單可完成，不發通知 | 狀態為已完成；`notified:false` |  | NOT RUN |  |
| E05 | 今日統計 | 完成數張測試訂單後開啟 Dashboard | 今日訂單、營收、待製作、完成率正確 | 訂單依 UUID 去重，不依商品列重複計數 |  | NOT RUN |  |
| E06 | 本月統計 | 確認本月多日訂單資料 | 趨勢圖及本月營收正確 | 未定價品項不計入營收 |  | NOT RUN |  |

## F. 取餐時間自動完成

| ID | 測試項目 | 前置設定／輸入 | 預期結果 | Sheet／通知檢查 | 實際結果 | 狀態 |
|---|---|---|---|---|---|---|
| F01 | 排程已安裝 | 在 Apps Script 執行一次 `installAutoCompletionTrigger()` | 回傳 `enabled: true` | `getAutoCompletionTriggerStatus()` 顯示每 5 分鐘執行 |  | NOT RUN |
| F02 | 未到取餐時間 | 建立取餐時間為 30 分鐘後的待製作訂單，執行 `autoCompleteOrdersByPickupTime()` | 訂單維持待製作 | 不發 LINE 完成通知 |  | NOT RUN |
| F03 | 到時自動完成 | 建立取餐時間已到、但不超過 12 小時的待製作訂單，執行排程函式 | 同 UUID 的所有列更新為已完成 | 不發 LINE 完成通知 |  | NOT RUN |
| F04 | 舊資料沒有取餐時間 | 準備 pickup_time 空白的舊待製作訂單 | 排程安全跳過 | 狀態維持待製作；新訂單不應再產生此狀況 |  | NOT RUN |
| F05 | 無效取餐時間 | 將 pickup_time 設為無法辨識的文字 | 排程安全跳過且不中斷其他訂單 | 狀態維持待製作 |  | NOT RUN |
| F06 | 重複執行 | 對已由 F03 完成的訂單再次執行排程 | 回傳時不重複更新 | 不重複發 LINE 通知 |  | NOT RUN |
| F07 | 明天取餐 | 輸入 `明天晚上6點取餐` | pickup_time 保存為明天的完整日期時間 | 今天不完成；明天時間到才完成 |  | NOT RUN |
| F08 | 過舊訂單保護 | 準備已超過取餐時間 12 小時的待製作舊訂單 | 排程跳過，避免首次啟用時異動過舊資料 | 狀態不變、不發通知 |  | NOT RUN |

## G. 月份封存與歷史後台

| ID | 測試項目 | 操作／前置資料 | 預期結果 | Sheet／Dashboard 檢查 | 實際結果 | 狀態 |
|---|---|---|---|---|---|---|
| G01 | 建立封存表 | 執行 `initializeOrderArchiveSheets()` | 函式成功完成 | 建立 OrdersArchive、MonthlySummary、MonthlyItemSummary 及正確表頭 |  | NOT RUN |
| G02 | 封存預覽 | 執行 `previewMonthlyArchive("上一月份")` | 只回傳預覽資料，不移動訂單 | Orders 列數不變；筆數、列數與營收可核對 |  | NOT RUN |
| G03 | 已完成整單封存 | 準備一張包含多個商品列的上月已完成訂單後執行封存 | 同 UUID 所有列一起封存 | Archive 完整保存 A:N 及 archived_at；Orders 不再有該 UUID |  | NOT RUN |
| G04 | 待製作保護 | 準備上月待製作訂單並執行封存 | 該訂單不封存 | Orders 保留完整訂單；Archive 無該 UUID |  | NOT RUN |
| G05 | 防止重複封存 | 對同月份連續執行兩次 `archiveMonthlyOrders()` | 第二次不新增重複列 | Archive 每個來源列只有一份；月統計不加倍 |  | NOT RUN |
| G06 | 未定價封存 | 準備含 `UNPRICED` 的已完成訂單 | 訂單仍可封存 | 明細保留；MonthlySummary 未定價訂單數增加且該品項營收為 0 |  | NOT RUN |
| G07 | 月商品統計 | 封存包含多品項的月份 | MonthlyItemSummary 依月份及商品彙總 | 數量與營收等於封存明細 |  | NOT RUN |
| G08 | Dashboard 月份選單 | 開啟後台並切換至已封存月份 | 畫面載入該月資料 | 營收、平均客單、訂單數、完成數、趨勢與熱銷品項正確 |  | NOT RUN |
| G09 | 雙來源去重 | 模擬同 UUID 暫時同時存在 Orders 與 OrdersArchive | Dashboard 只計算一次 | 訂單數與營收不重複 |  | NOT RUN |
| G10 | 每日補封存排程 | 執行 `installOrderArchiveTrigger()`，再將舊待製作單改成已完成 | 下一次排程封存該訂單 | Trigger 每日 03:00 左右執行，舊月份資料最終移出 Orders |  | NOT RUN |
| G11 | 六月歷史測試資料 | 執行 `generateJuneArchiveTestOrders()` 並重新整理 Dashboard | 月份選單出現 2026/06，可切換查看趨勢、營收與 25 張測試訂單 | Orders 不變；OrdersArchive 新增 `DASHBOARD-TEST-202606-` 資料 |  | NOT RUN |
| G12 | 清除六月測試資料 | 執行 `clearJuneArchiveTestOrders()` | 測試訂單消失，六月真實封存資料仍保留 | 只移除測試前綴資料並重建六月彙總 |  | NOT RUN |

## H. AI 接單開關

| ID | 測試項目 | 操作／前置資料 | 預期結果 | Sheet／LINE 檢查 | 實際結果 | 狀態 |
|---|---|---|---|---|---|---|
| H01 | 預設啟用 | 尚未切換過狀態時開啟 Dashboard | AI 接單顯示啟用 | LINE 訂單可正常解析及寫入 |  | NOT RUN |
| H02 | 暫停接單 | 在 Dashboard 關閉 AI 接單並確認，再傳送完整訂單 | 不呼叫 AI 且 LINE 不回覆 | Orders 不新增；後台與缺貨管理仍可使用 |  | NOT RUN |
| H03 | 恢復接單 | 重新啟用後傳送完整訂單 | LINE 正常回覆訂單 | Orders 正常新增；後台顯示最後切換時間 |  | NOT RUN |
| H04 | 處理中暫停 | 傳送訂單後立即關閉開關，讓切換發生在 AI 回覆前 | 該請求不得寫入訂單 | Orders 無該筆資料；LINE 不回覆 |  | NOT RUN |

## 測試結果摘要

| 指標 | 數量 |
|---|---:|
| 測試案例總數 | 69 |
| PASS |  |
| FAIL |  |
| BLOCKED |  |
| NOT RUN |  |

## 問題紀錄

| 問題 ID | 對應測試 ID | 問題描述 | Cloud Run Log 時間 | Apps Script 執行時間 | 處理狀態 |
|---|---|---|---|---|---|
| BUG-001 |  |  |  |  |  |
