# LINE 滷味訂單助手測試表

測試日期：____________　測試人員：____________　Cloud Run 修訂版本：____________　Apps Script 部署版本：____________

## 測試前準備

- 先備份 `Orders`、`Customers`、`PriceList` 工作表。
- 確認 Cloud Run 已設定 `GAS_URL`、`GAS_SHARED_SECRET`、`CLOUD_RUN_NOTIFY_SECRET`、`LINE_CHANNEL_ACCESS_TOKEN`、`LINE_CHANNEL_SECRET`。
- 確認 Apps Script 已設定 `WEBHOOK_SHARED_SECRET`、`CLOUD_RUN_NOTIFY_SECRET`、`OPENAI_API_KEY`。
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
| A05 | 完成通知錯誤 secret | 使用錯誤 `notifySecret` POST `/notifyReady` | 回傳 `401 UNAUTHORIZED`，不發 LINE 通知 |  | NOT RUN |  |

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
| B08 | 無取餐時間 | `高麗菜一份` | 訂單仍成立並顯示系統預估時間 | pickup_time 可為空；estimatedPickupTime 正常 |  | NOT RUN |  |

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

## E. 去重、完成與統計

| ID | 測試項目 | 操作／輸入 | 預期結果 | Sheet／Dashboard 檢查 | 實際結果 | 狀態 | 備註 |
|---|---|---|---|---|---|---|---|
| E01 | Webhook event 去重 | 使用相同 `webhookEventId` 將同一訂單送入 Apps Script 兩次 | 第二次回傳 `{"duplicate":true}` | Orders 只建立一張；WebhookEvents 只有一筆 event ID |  | NOT RUN |  |
| E02 | 去重失敗重試 | 模擬第一次寫入失敗後再次傳送相同 event ID | 第二次可重新處理 | 失敗時 event ID 已撤銷，不會永久卡住 |  | NOT RUN | 建議在測試部署執行 |
| E03 | 唯一 UUID 完成訂單 | 準備兩張顯示單號相同但 UUID 不同的訂單，完成其中一張 | 只完成指定 UUID | 另一張狀態不變；通知顯示人看的單號 |  | NOT RUN |  |
| E04 | 無 LINE ID 的手動訂單 | 手動建立訂單後呼叫完成函式 | 訂單可完成，不發通知 | 狀態為已完成；`notified:false` |  | NOT RUN |  |
| E05 | 今日統計 | 完成數張測試訂單後開啟 Dashboard | 今日訂單、營收、待製作、完成率正確 | 訂單依 UUID 去重，不依商品列重複計數 |  | NOT RUN |  |
| E06 | 本月統計 | 確認本月多日訂單資料 | 趨勢圖及本月營收正確 | 未定價品項不計入營收 |  | NOT RUN |  |

## 測試結果摘要

| 指標 | 數量 |
|---|---:|
| 測試案例總數 | 30 |
| PASS |  |
| FAIL |  |
| BLOCKED |  |
| NOT RUN |  |

## 問題紀錄

| 問題 ID | 對應測試 ID | 問題描述 | Cloud Run Log 時間 | Apps Script 執行時間 | 處理狀態 |
|---|---|---|---|---|---|
| BUG-001 |  |  |  |  |  |
