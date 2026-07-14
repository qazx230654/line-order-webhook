const SYSTEM_PROMPT = `
你是一位滷味店訂單解析助手。

請將客戶訊息解析為 JSON。

任務：

1. 辨識取餐時間
2. 辨識商品與數量
3. 辨識備註（辣度、酸菜、蔥、做熱等）

分組規則：

若訊息中出現：

* 第一袋、第二袋
* 分開裝
* 不同辣度
* 不同備註
* 編號 1. 2. 3.

代表同一筆訂單包含多組內容。

請建立 groups 陣列，
每個 group 代表一袋或一組獨立製作需求。

訂單判斷：

若訊息包含任一項：

* 商品名稱
* 商品數量
* 備註（辣度、酸菜、做熱等）
* 取餐時間
* 分袋需求

則：

"is_order": true

若屬於：

* 詢問價格
* 詢問營業時間
* 詢問菜單
* 一般聊天
* 感謝
* 打招呼
* 取消訂單
* 其他非點餐內容

則：

"is_order": false

回傳格式：

{
"is_order": true,
"customer_name": "",
"phone": "",
"pickup_time": "",
"groups": [
{
"name": "",
"items": [
{
"name": "",
"raw_name": "",
"suggested_name": "",
"quantity": 1
}
],
"note": ""
}
]
}

規則：

1. 僅回傳 JSON
2. 不要 markdown
3. 不要解釋
4. 缺少資料請填空字串
5. quantity 必須為數字
6. 數量轉換規則：
   - 「半份」= 0.5
   - 「半個」= 0.5
   - 「半」若明確接在商品後面，也視為 0.5
   - 「一份半」= 1.5
   - 「1份半」= 1.5
   - 「兩份半」= 2.5
7. raw_name 必須保留客戶原本輸入的商品名稱，不包含數量
8. 若能明確對應正式菜單，name 必須使用完全相同的正式名稱，suggested_name 填空字串
9. 若無法明確對應，name 與 raw_name 相同；suggested_name 可填最接近的正式名稱，無合理候選則填空字串
10. 不可自行創造正式菜單以外的商品名稱
11. 若有取餐時間，pickup_time 必須轉成 Asia/Taipei 時區的 yyyy/MM/dd HH:mm；只有時間時使用今天日期，明天或其他日期則換算成正確日期
   `;

const CLOUD_RUN_NOTIFY_URL =
  "https://line-order-webhook-171295331325.asia-east1.run.app/notifyReady";

function parseOrder(message) {
  const apiKey =
    PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  const menuItemNames = getCanonicalItemNames_();

  const currentTime = Utilities.formatDate(
    new Date(),
    "Asia/Taipei",
    "yyyy/MM/dd HH:mm",
  );

  const menuPrompt = `

目前時間（Asia/Taipei）：${currentTime}

正式菜單名稱（只能從此清單選擇）：
${JSON.stringify(menuItemNames)}

只有在能明確判斷時，才將 name 改成清單內的正式名稱。
無法確定時保留客戶原文，交由後端標記，不可猜測。
`;

  const payload = {
    model: "gpt-4.1-mini",
    temperature: 0,
    response_format: {
      type: "json_object",
    },
    messages: [
      {
        role: "system",
        content: SYSTEM_PROMPT + menuPrompt,
      },
      {
        role: "user",
        content: message,
      },
    ],
  };

  const response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/chat/completions",
    {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: `Bearer ${apiKey}`,
      },
      payload: JSON.stringify(payload),
    },
  );

  const result = JSON.parse(response.getContentText());

  const content = result.choices[0].message.content;

  return JSON.parse(content);
}

function saveOrder(
  order,
  rawMessage,
  lineUserId,
  systemOrderId = null,
  displayOrderId = null,
  createdAt = null,
) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");

  const priceMap = getPriceMap();

  const priceColumn = 14;

  systemOrderId = systemOrderId || generateSystemOrderId();

  displayOrderId = displayOrderId || generateDisplayOrderId();

  createdAt =
    createdAt ||
    Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");

  if (!order || !Array.isArray(order.groups)) {
    throw new Error("訂單格式錯誤：缺少 groups");
  }

  const rows = [];

  order.groups.forEach((group) => {
    if (!Array.isArray(group.items)) {
      return;
    }

    group.items.forEach((item) => {
      const quantity = Number(item.quantity) || 0;

      const unitPrice = quantity > 10 ? quantity : priceMap[item.name];

      rows.push([
        systemOrderId,

        displayOrderId,

        createdAt,

        order.customer_name || "",

        order.phone || "",

        group.name || "一般",

        item.name,

        quantity,

        order.pickup_time || "",

        group.note || "",

        "待製作",

        rawMessage,

        lineUserId,

        unitPrice === undefined ? "UNPRICED" : unitPrice,
      ]);
    });
  });

  if (rows.length === 0) {
    throw new Error("訂單格式錯誤：沒有可寫入的商品");
  }

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    if (!sheet.getRange(1, priceColumn).getValue()) {
      sheet.getRange(1, priceColumn).setValue("unit_price");
    }

    sheet
      .getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
      .setValues(rows);
  } finally {
    lock.releaseLock();
  }

  return {
    systemOrderId,

    displayOrderId,

    createdAt,
  };
}

function manualCreateOrder(message) {
  const order = parseOrder(message);

  if (!order.is_order) {
    throw new Error("訊息無法解析為訂單");
  }

  const normalizationResult = normalizeItems(order);

  ensureOrderPickupTime_(order, Math.max(getEstimatedWaitMinutes() + 5, 10));

  const saveResult = saveOrder(order, message);

  logUnmatchedItemsSafely_(normalizationResult.unmatchedItems, message);

  return saveResult;
}

function getOrders() {
  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");

    const values = sheet.getDataRange().getDisplayValues();

    const orders = {};

    const customerMap = getCustomerMap();

    const priceMap = getPriceMap();

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      if (row[10] !== "待製作") continue;

      const systemOrderId = row[0];

      if (!orders[systemOrderId]) {
        const lineUserId = row[12];

        const customer = customerMap[lineUserId] || {};
        orders[systemOrderId] = {
          systemOrderId: systemOrderId,
          orderId: row[1],
          createdAt: row[2].substring(11, 16),
          customer: row[3],
          pictureUrl: customer.pictureUrl || "",
          totalOrders: customer.totalOrders || 0,
          pickup: row[8],
          status: row[10],

          totalPrice: 0,
          priceMissing: false,

          missingItems: [],

          groups: {},
        };
      }

      const groupName = row[5] || "一般";

      if (!orders[systemOrderId].groups[groupName]) {
        orders[systemOrderId].groups[groupName] = [];
      }

      const itemName = row[6];

      const quantity = Number(row[7]);

      let displayItem = "";

      if (quantity > 10) {
        displayItem = `${itemName}${quantity} x1`;
      } else {
        displayItem = `${itemName} x${quantity}`;
      }

      orders[systemOrderId].groups[groupName].push({
        item: displayItem,

        note: row[9],
      });

      const priceInfo = resolveOrderPrice_(row, priceMap);

      if (priceInfo.missing) {
        orders[systemOrderId].priceMissing = true;

        if (!orders[systemOrderId].missingItems.includes(itemName)) {
          orders[systemOrderId].missingItems.push(itemName);
        }

        console.log("價格未設定:", itemName);
      }

      if (!priceInfo.missing) {
        orders[systemOrderId].totalPrice += priceInfo.lineTotal;
      }
    }

    return Object.values(orders).sort((a, b) =>
      a.orderId.localeCompare(b.orderId),
    );
  } catch (error) {
    console.error("getOrders錯誤:", error.stack);

    throw error;
  }
}

function updateOrderStatus(systemOrderId) {
  let lineUserId = "";
  let displayOrderId = "";

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");

    const values = sheet.getDataRange().getValues();

    const matchingRows = [];
    const pendingRows = [];

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] !== systemOrderId) {
        continue;
      }

      matchingRows.push(i + 1);

      if (!displayOrderId) {
        lineUserId = values[i][12];
        displayOrderId = values[i][1];
      }

      if (values[i][10] === "待製作") {
        pendingRows.push(i + 1);
      }
    }

    if (matchingRows.length === 0) {
      throw new Error("找不到指定的系統訂單 ID");
    }

    if (pendingRows.length === 0) {
      return {
        systemOrderId,
        displayOrderId,
        notified: false,
        alreadyCompleted: true,
      };
    }

    sheet
      .getRangeList(pendingRows.map((rowNumber) => `K${rowNumber}`))
      .setValue("已完成");
  } finally {
    lock.releaseLock();
  }

  if (!lineUserId) {
    return {
      systemOrderId,
      displayOrderId,
      notified: false,
    };
  }

  console.log("通知使用者:", lineUserId);

  UrlFetchApp.fetch(CLOUD_RUN_NOTIFY_URL, {
    method: "post",
    contentType: "application/json",

    payload: JSON.stringify({
      userId: lineUserId,
      orderId: displayOrderId,
      notifySecret: getRequiredScriptProperty_("CLOUD_RUN_NOTIFY_SECRET"),
    }),
  });

  return {
    systemOrderId,
    displayOrderId,
    notified: true,
  };
}

function getTaipeiDateParts_(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const formatted = Utilities.formatDate(value, "Asia/Taipei", "yyyy/MM/dd");

    const parts = formatted.split("/");

    return {
      year: Number(parts[0]),
      month: Number(parts[1]),
      day: Number(parts[2]),
    };
  }

  const match = String(value || "").match(
    /(20\d{2})[\/.\-](\d{1,2})[\/.\-](\d{1,2})/,
  );

  if (!match) {
    return null;
  }

  return {
    year: Number(match[1]),
    month: Number(match[2]),
    day: Number(match[3]),
  };
}

function shiftDateParts_(dateParts, days) {
  const date = new Date(
    Date.UTC(dateParts.year, dateParts.month - 1, dateParts.day + days),
  );

  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
  };
}

function parsePickupDateTime_(pickupValue, createdAtValue) {
  if (pickupValue instanceof Date && !Number.isNaN(pickupValue.getTime())) {
    const pickupYear = Number(
      Utilities.formatDate(pickupValue, "Asia/Taipei", "yyyy"),
    );

    if (pickupYear > 1900) {
      return pickupValue;
    }

    pickupValue = Utilities.formatDate(pickupValue, "Asia/Taipei", "HH:mm");
  }

  let text = String(pickupValue || "").trim();

  if (!text) {
    return null;
  }

  if (text.normalize) {
    text = text.normalize("NFKC");
  }

  let dateParts = getTaipeiDateParts_(text);

  const createdDateParts = getTaipeiDateParts_(createdAtValue);

  if (!dateParts) {
    const shortDateMatch = text.match(
      /(?:^|\s)(\d{1,2})[\/.\-](\d{1,2})(?:\s|$)/,
    );

    if (shortDateMatch && createdDateParts) {
      dateParts = {
        year: createdDateParts.year,
        month: Number(shortDateMatch[1]),
        day: Number(shortDateMatch[2]),
      };
    } else {
      dateParts = createdDateParts;
    }
  }

  if (!dateParts) {
    return null;
  }

  if (/明天|翌日/.test(text)) {
    dateParts = shiftDateParts_(dateParts, 1);
  }

  let timeMatch = text.match(/(\d{1,2})\s*[:：]\s*(\d{1,2})/);

  let hour;
  let minute;

  if (timeMatch) {
    hour = Number(timeMatch[1]);
    minute = Number(timeMatch[2]);
  } else {
    timeMatch = text.match(/(\d{1,2})\s*(?:點|時)\s*(半|\d{1,2})?\s*分?/);

    if (!timeMatch) {
      return null;
    }

    hour = Number(timeMatch[1]);
    minute = timeMatch[2] === "半" ? 30 : Number(timeMatch[2] || 0);
  }

  const periodMatch = text.match(/(凌晨|早上|上午|中午|下午|傍晚|晚上)/);

  const period = periodMatch ? periodMatch[1] : "";

  if (["下午", "傍晚", "晚上"].includes(period) && hour < 12) {
    hour += 12;
  }

  if (period === "中午" && hour < 11) {
    hour += 12;
  }

  if (["凌晨", "早上", "上午"].includes(period) && hour === 12) {
    hour = 0;
  }

  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) {
    return null;
  }

  const year = String(dateParts.year).padStart(4, "0");

  const month = String(dateParts.month).padStart(2, "0");

  const day = String(dateParts.day).padStart(2, "0");

  const hours = String(hour).padStart(2, "0");

  const minutes = String(minute).padStart(2, "0");

  const calendarCheck = new Date(
    Date.UTC(dateParts.year, dateParts.month - 1, dateParts.day),
  );

  if (
    calendarCheck.getUTCFullYear() !== dateParts.year ||
    calendarCheck.getUTCMonth() + 1 !== dateParts.month ||
    calendarCheck.getUTCDate() !== dateParts.day
  ) {
    return null;
  }

  const result = new Date(
    `${year}-${month}-${day}` + `T${hours}:${minutes}:00+08:00`,
  );

  return Number.isNaN(result.getTime()) ? null : result;
}

function findOrdersDueForCompletion_(values, now) {
  const maximumOverdueMs = 12 * 60 * 60 * 1000;

  const orders = Object.create(null);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    if (row[10] !== "待製作") {
      continue;
    }

    const systemOrderId = String(row[0] || "").trim();

    if (!systemOrderId || orders[systemOrderId]) {
      continue;
    }

    const pickupAt = parsePickupDateTime_(row[8], row[2]);

    if (!pickupAt) {
      continue;
    }

    const overdueMs = now.getTime() - pickupAt.getTime();

    if (overdueMs >= 0 && overdueMs <= maximumOverdueMs) {
      orders[systemOrderId] = {
        systemOrderId,
        displayOrderId: String(row[1] || ""),
        pickupAt,
      };
    }
  }

  return Object.values(orders);
}

function autoCompleteOrdersByPickupTime() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");

  if (!sheet || sheet.getLastRow() <= 1) {
    return {
      checked: 0,
      completed: 0,
      failed: [],
    };
  }

  const values = sheet.getDataRange().getValues();

  const dueOrders = findOrdersDueForCompletion_(values, new Date());

  const completed = [];
  const failed = [];

  dueOrders.forEach((order) => {
    try {
      const result = updateOrderStatus(order.systemOrderId);

      completed.push({
        systemOrderId: order.systemOrderId,
        displayOrderId: order.displayOrderId,
        notified: result.notified,
      });
    } catch (error) {
      failed.push({
        systemOrderId: order.systemOrderId,
        displayOrderId: order.displayOrderId,
        error: error.message || String(error),
      });
    }
  });

  const summary = {
    checked: dueOrders.length,
    completed: completed.length,
    completedOrders: completed,
    failed,
  };

  console.log(JSON.stringify(summary));

  return summary;
}

function installAutoCompletionTrigger() {
  const handler = "autoCompleteOrdersByPickupTime";

  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === handler)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  const trigger = ScriptApp.newTrigger(handler)
    .timeBased()
    .everyMinutes(5)
    .create();

  return {
    enabled: true,
    handler,
    everyMinutes: 5,
    triggerId: trigger.getUniqueId(),
  };
}

function getAutoCompletionTriggerStatus() {
  const handler = "autoCompleteOrdersByPickupTime";

  const triggers = ScriptApp.getProjectTriggers().filter(
    (trigger) => trigger.getHandlerFunction() === handler,
  );

  return {
    enabled: triggers.length > 0,
    handler,
    everyMinutes: 5,
    triggerCount: triggers.length,
  };
}

function removeAutoCompletionTrigger() {
  const handler = "autoCompleteOrdersByPickupTime";

  let removed = 0;

  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === handler)
    .forEach((trigger) => {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    });

  return {
    enabled: false,
    removed,
  };
}

function getCustomerMap() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");

  const values = sheet.getDataRange().getValues();

  const map = {};

  for (let i = 1; i < values.length; i++) {
    map[values[i][0]] = {
      displayName: values[i][1] || "",

      pictureUrl: values[i][2] || "",

      totalOrders: Number(values[i][5]) || 0,
    };
  }

  return map;
}

function getDashboardSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");

  const values = sheet.getDataRange().getDisplayValues();

  const orderMap = {};
  const today = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd");

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const orderDate = row[2].substring(0, 10);

    if (orderDate !== today) continue;

    orderMap[row[1]] = row[10];
  }

  let pending = 0;
  let completed = 0;

  Object.values(orderMap).forEach((status) => {
    if (status === "待製作") {
      pending++;
    }

    if (status === "已完成") {
      completed++;
    }
  });

  return {
    pending,
    completed,
  };
}

function getCompletedOrders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");

  const values = sheet.getDataRange().getDisplayValues();

  const orders = {};

  const customerMap = getCustomerMap();

  const priceMap = getPriceMap();

  const today = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd");

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    if (row[10] !== "已完成") continue;

    const orderDate = row[2].substring(0, 10);

    if (orderDate !== today) continue;

    const systemOrderId = row[0];

    if (!orders[systemOrderId]) {
      const lineUserId = row[12];

      const customer = customerMap[lineUserId] || {};
      orders[systemOrderId] = {
        systemOrderId: systemOrderId,
        orderId: row[1],
        createdAt: row[2].substring(11, 16),
        customer: row[3],
        pictureUrl: customer.pictureUrl || "",
        totalOrders: customer.totalOrders || 0,
        pickup: row[8],
        status: row[10],

        totalPrice: 0,
        priceMissing: false,

        missingItems: [],

        groups: {},
      };
    }

    const groupName = row[5] || "一般";

    if (!orders[systemOrderId].groups[groupName]) {
      orders[systemOrderId].groups[groupName] = [];
    }

    orders[systemOrderId].groups[groupName].push({
      item: `${row[6]} x${row[7]}`,

      note: row[9],
    });

    const itemName = row[6];

    const quantity = Number(row[7]);

    const priceInfo = resolveOrderPrice_(row, priceMap);

    if (priceInfo.missing) {
      orders[systemOrderId].priceMissing = true;
      console.log("價格未設定:", itemName);
    }

    if (!priceInfo.missing) {
      orders[systemOrderId].totalPrice += priceInfo.lineTotal;
    }
  }

  return Object.values(orders)
    .sort((a, b) => b.orderId.localeCompare(a.orderId))
    .slice(0, 20);
}

function generateSystemOrderId() {
  return Utilities.getUuid();
}

function generateDisplayOrderId() {
  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const props = PropertiesService.getScriptProperties();

    const today = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd");

    const savedDate = props.getProperty("ORDER_DATE");

    if (savedDate !== today) {
      props.setProperty("ORDER_DATE", today);

      props.setProperty("ORDER_COUNTER", "0");
    }

    let counter = parseInt(props.getProperty("ORDER_COUNTER") || "0");

    counter++;

    props.setProperty("ORDER_COUNTER", String(counter));

    Logger.log("Counter=" + counter);

    return "A" + String(counter).padStart(3, "0");
  } finally {
    lock.releaseLock();
  }
}

function calculateOrderPrice(order) {
  const priceMap = getPriceMap();

  let totalPrice = 0;

  const missingItems = [];

  order.groups.forEach((group) => {
    group.items.forEach((item) => {
      const price = priceMap[item.name];

      if (item.quantity > 10) {
        totalPrice += Number(item.quantity) || 0;
        return;
      }

      if (price === undefined) {
        if (!missingItems.includes(item.name)) {
          missingItems.push(item.name);
        }

        return;
      }

      totalPrice += price * item.quantity;
    });
  });

  return {
    totalPrice,
    priceMissing: missingItems.length > 0,
    missingItems,
  };
}

function getPriceMap() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PriceList");

  const values = sheet.getDataRange().getValues();

  const map = {};

  for (let i = 1; i < values.length; i++) {
    const itemName = String(values[i][0] || "").trim();

    const rawPrice = values[i][1];

    if (
      !itemName ||
      rawPrice === "" ||
      rawPrice === null ||
      rawPrice === undefined
    ) {
      continue;
    }

    const price = Number(rawPrice);

    if (!Number.isFinite(price)) {
      continue;
    }

    map[itemName] = price;
  }

  return map;
}

function getCanonicalItemNames_() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PriceList");

  if (!sheet) {
    return [];
  }

  const values = sheet.getDataRange().getValues();

  const names = [];
  const seen = Object.create(null);

  for (let i = 1; i < values.length; i++) {
    const itemName = String(values[i][0] || "").trim();

    if (itemName && !seen[itemName]) {
      seen[itemName] = true;
      names.push(itemName);
    }
  }

  return names;
}

function resolveOrderPrice_(row, priceMap) {
  const itemName = String(row[6] || "");

  const quantity = Number(row[7]) || 0;

  const storedPrice = row[13];

  if (storedPrice === "UNPRICED") {
    return {
      missing: true,
      unitPrice: null,
      lineTotal: 0,
    };
  }

  if (
    storedPrice !== "" &&
    storedPrice !== null &&
    storedPrice !== undefined &&
    Number.isFinite(Number(storedPrice))
  ) {
    const unitPrice = Number(storedPrice);

    return {
      missing: false,
      unitPrice,
      lineTotal: quantity > 10 ? unitPrice : unitPrice * quantity,
    };
  }

  // 舊訂單沒有成交單價欄位，維持原本以 PriceList 回推的行為。
  if (quantity > 10) {
    return {
      missing: false,
      unitPrice: quantity,
      lineTotal: quantity,
    };
  }

  const currentPrice = priceMap[itemName];

  if (currentPrice === undefined) {
    return {
      missing: true,
      unitPrice: null,
      lineTotal: 0,
    };
  }

  return {
    missing: false,
    unitPrice: currentPrice,
    lineTotal: currentPrice * quantity,
  };
}

function getAliasMap() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ItemAlias");

  const map = Object.create(null);

  if (sheet) {
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      const alias = String(values[i][0] || "").trim();

      const canonicalName = String(values[i][1] || "").trim();

      if (alias && canonicalName) {
        map[alias] = canonicalName;
      }
    }
  }

  const unmatchedSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UnmatchedItems");

  if (unmatchedSheet) {
    const values = unmatchedSheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      const status = String(values[i][5] || "")
        .trim()
        .toLowerCase();

      if (status !== "approved" && status !== "已核准") {
        continue;
      }

      const alias = String(values[i][0] || "").trim();

      const canonicalName = String(values[i][1] || "").trim();

      if (alias && canonicalName) {
        map[alias] = canonicalName;
      }
    }
  }

  return map;
}

function getEstimatedWaitMinutes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");

  const values = sheet.getDataRange().getDisplayValues();

  const orderIds = new Set();

  for (let i = 1; i < values.length; i++) {
    if (values[i][10] === "待製作") {
      orderIds.add(values[i][0]);
    }
  }
  //一單製作時間 預計分鐘數
  return orderIds.size * 5;
}

function getEstimatedPickupTime(waitMinutes, now = new Date()) {
  const finishTime = new Date(now.getTime() + waitMinutes * 60000);

  return Utilities.formatDate(finishTime, "Asia/Taipei", "HH:mm");
}

function getEstimatedPickupDateTime_(waitMinutes, now = new Date()) {
  const finishTime = new Date(now.getTime() + waitMinutes * 60000);

  return Utilities.formatDate(finishTime, "Asia/Taipei", "yyyy/MM/dd HH:mm");
}

function getPickupTimeDisplay_(pickupTime, createdAt = new Date()) {
  const parsed = parsePickupDateTime_(pickupTime, createdAt);

  if (parsed) {
    return Utilities.formatDate(parsed, "Asia/Taipei", "HH:mm");
  }

  return String(pickupTime || "").trim();
}

function ensureOrderPickupTime_(order, waitMinutes, now = new Date()) {
  const requestedPickupTime = String(order.pickup_time || "").trim();

  if (requestedPickupTime) {
    order.pickup_time = requestedPickupTime;

    return {
      pickupTime: requestedPickupTime,
      displayTime: getPickupTimeDisplay_(requestedPickupTime, now),
      estimated: false,
    };
  }

  const estimatedPickupTime = getEstimatedPickupDateTime_(waitMinutes, now);

  order.pickup_time = estimatedPickupTime;

  return {
    pickupTime: estimatedPickupTime,
    displayTime: getPickupTimeDisplay_(estimatedPickupTime, now),
    estimated: true,
  };
}

function normalizeItemKey_(value) {
  let text = String(value || "").trim();

  if (text.normalize) {
    text = text.normalize("NFKC");
  }

  return text
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[.,，。、:：;；'"`·・()（）\[\]【】{}]/g, "")
    .replace(/乾/g, "干")
    .replace(/臺/g, "台");
}

function addUniqueNameMapping_(map, key, canonicalName) {
  if (!key) {
    return;
  }

  if (!Object.prototype.hasOwnProperty.call(map, key)) {
    map[key] = canonicalName;
    return;
  }

  if (map[key] !== canonicalName) {
    map[key] = null;
  }
}

function getItemNormalizationData_() {
  const canonicalNames = getCanonicalItemNames_();

  const canonicalExact = Object.create(null);

  const canonicalNormalized = Object.create(null);

  canonicalNames.forEach((name) => {
    canonicalExact[name] = name;
    addUniqueNameMapping_(canonicalNormalized, normalizeItemKey_(name), name);
  });

  const aliases = getAliasMap();

  const aliasExact = Object.create(null);

  const aliasNormalized = Object.create(null);

  Object.keys(aliases).forEach((alias) => {
    const requestedCanonical = String(aliases[alias] || "").trim();

    const canonicalName =
      canonicalExact[requestedCanonical] ||
      canonicalNormalized[normalizeItemKey_(requestedCanonical)];

    if (!canonicalName) {
      return;
    }

    addUniqueNameMapping_(aliasExact, alias, canonicalName);

    addUniqueNameMapping_(
      aliasNormalized,
      normalizeItemKey_(alias),
      canonicalName,
    );
  });

  return {
    canonicalExact,
    canonicalNormalized,
    aliasExact,
    aliasNormalized,
  };
}

function resolveCanonicalItemName_(value, normalizationData) {
  const name = String(value || "").trim();

  if (!name) {
    return null;
  }

  if (normalizationData.canonicalExact[name]) {
    return normalizationData.canonicalExact[name];
  }

  const normalizedKey = normalizeItemKey_(name);

  if (normalizationData.canonicalNormalized[normalizedKey]) {
    return normalizationData.canonicalNormalized[normalizedKey];
  }

  if (normalizationData.aliasExact[name]) {
    return normalizationData.aliasExact[name];
  }

  return normalizationData.aliasNormalized[normalizedKey] || null;
}

function normalizeItems(order) {
  const normalizationData = getItemNormalizationData_();

  const unmatchedItems = [];

  if (!order || !Array.isArray(order.groups)) {
    return { unmatchedItems };
  }

  order.groups.forEach((group) => {
    if (!Array.isArray(group.items)) {
      return;
    }

    group.items.forEach((item) => {
      const parsedName = String(item.name || "").trim();

      const rawName = String(item.raw_name || parsedName).trim();

      const canonicalName =
        resolveCanonicalItemName_(parsedName, normalizationData) ||
        resolveCanonicalItemName_(rawName, normalizationData);

      item.raw_name = rawName || parsedName;

      if (canonicalName) {
        item.name = canonicalName;
        item.suggested_name = "";
        item.match_status = "matched";
        return;
      }

      const suggestedName =
        resolveCanonicalItemName_(item.suggested_name, normalizationData) || "";

      item.name = rawName || parsedName;

      item.suggested_name = suggestedName;

      item.match_status = "unmatched";

      if (item.name) {
        unmatchedItems.push({
          rawName: item.name,
          suggestedName,
        });
      }
    });
  });

  return { unmatchedItems };
}

function logUnmatchedItems_(unmatchedItems, rawMessage) {
  if (!Array.isArray(unmatchedItems) || unmatchedItems.length === 0) {
    return;
  }

  const uniqueItems = Object.create(null);

  unmatchedItems.forEach((item) => {
    const key = normalizeItemKey_(item.rawName);

    if (!key) {
      return;
    }

    if (!uniqueItems[key]) {
      uniqueItems[key] = item;
    } else if (!uniqueItems[key].suggestedName && item.suggestedName) {
      uniqueItems[key].suggestedName = item.suggestedName;
    }
  });

  const keys = Object.keys(uniqueItems);

  if (keys.length === 0) {
    return;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    let sheet = spreadsheet.getSheetByName("UnmatchedItems");

    if (!sheet) {
      sheet = spreadsheet.insertSheet("UnmatchedItems");
    }

    const headers = [
      "raw_name",
      "suggested_name",
      "raw_message",
      "occurrence_count",
      "last_seen",
      "status",
    ];

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    const values = sheet.getDataRange().getValues();

    const rowByKey = Object.create(null);

    for (let i = 1; i < values.length; i++) {
      const key = normalizeItemKey_(values[i][0]);

      if (key && !rowByKey[key]) {
        rowByKey[key] = i + 1;
      }
    }

    const now = Utilities.formatDate(
      new Date(),
      "Asia/Taipei",
      "yyyy/MM/dd HH:mm:ss",
    );

    keys.forEach((key) => {
      const item = uniqueItems[key];

      const rowIndex = rowByKey[key];

      if (rowIndex) {
        const existing = values[rowIndex - 1];

        const suggestion = item.suggestedName || existing[1] || "";

        const count = (Number(existing[3]) || 0) + 1;

        sheet
          .getRange(rowIndex, 1, 1, headers.length)
          .setValues([
            [
              existing[0] || item.rawName,
              suggestion,
              rawMessage || "",
              count,
              now,
              existing[5] || "pending",
            ],
          ]);

        return;
      }

      sheet.appendRow([
        item.rawName,
        item.suggestedName || "",
        rawMessage || "",
        1,
        now,
        "pending",
      ]);
    });
  } finally {
    lock.releaseLock();
  }
}

function logUnmatchedItemsSafely_(unmatchedItems, rawMessage) {
  try {
    logUnmatchedItems_(unmatchedItems, rawMessage);
  } catch (error) {
    console.error("Failed to log unmatched items", error);
  }
}

function incrementCustomerOrderCount(lineUserId) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === lineUserId) {
      const current = Number(values[i][5]) || 0;

      sheet.getRange(i + 1, 6).setValue(current + 1);

      return;
    }
  }
}

function saveCustomer(lineUserId, displayName, pictureUrl) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === lineUserId) {
      sheet.getRange(i + 1, 2).setValue(displayName);

      sheet.getRange(i + 1, 3).setValue(pictureUrl);

      sheet.getRange(i + 1, 5).setValue(new Date());

      return;
    }
  }

  sheet.appendRow([lineUserId, displayName, pictureUrl, "", new Date(), 0]);
}

function deleteOrder(systemOrderId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");

  const values = sheet.getDataRange().getValues();

  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][0] === systemOrderId) {
      sheet.deleteRow(i + 1);
    }
  }
}

function findEditableOrder(lineUserId) {
  if (!lineUserId) {
    return null;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");

  const values = sheet.getDataRange().getValues();

  const now = new Date();

  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][12]).trim() !== String(lineUserId).trim()) {
      continue;
    }

    if (values[i][10] !== "待製作") {
      continue;
    }

    const createdAt = new Date(values[i][2]);

    const diffMinutes = (now - createdAt) / 1000 / 60;

    // 建立後20分鐘內可修改
    if (diffMinutes >= 0 && diffMinutes <= 20) {
      return {
        systemOrderId: values[i][0],

        displayOrderId: values[i][1],

        createdAt: values[i][2],

        rawMessage: values[i][11],
      };
    }
  }

  return null;
}

function updateOrder(
  editableOrder,
  order,
  newMessage,
  lineUserId,
  pickupDisplayTime,
) {
  let orderPersisted = false;

  try {
    const saveResult = saveOrder(
      order,
      newMessage,
      lineUserId,
      null,
      editableOrder.displayOrderId,
      editableOrder.createdAt,
    );

    orderPersisted = true;

    deleteOrder(editableOrder.systemOrderId);

    const priceInfo = calculateOrderPrice(order);

    return {
      is_order: true,

      updated: true,

      orderId: saveResult.displayOrderId,

      order,

      totalPrice: priceInfo.totalPrice,

      priceMissing: priceInfo.priceMissing,

      missingItems: priceInfo.missingItems,

      estimatedPickupTime:
        pickupDisplayTime ||
        getPickupTimeDisplay_(order.pickup_time, editableOrder.createdAt),
    };
  } catch (error) {
    error.orderPersisted = orderPersisted;

    throw error;
  }
}

const ORDER_ARCHIVE_HEADERS_ = [
  "system_order_id",
  "display_order_id",
  "created_at",
  "customer_name",
  "phone",
  "group_name",
  "item_name",
  "quantity",
  "pickup_time",
  "note",
  "status",
  "raw_message",
  "line_user_id",
  "unit_price",
  "archived_at",
];

const MONTHLY_SUMMARY_HEADERS_ = [
  "month",
  "order_count",
  "completed_count",
  "revenue",
  "unpriced_order_count",
  "average_order_value",
  "archived_at",
];

const MONTHLY_ITEM_SUMMARY_HEADERS_ = [
  "month",
  "item_name",
  "quantity",
  "revenue",
];

function ensureSheetWithHeaders_(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return sheet;
}

function initializeOrderArchiveSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheetWithHeaders_(spreadsheet, "OrdersArchive", ORDER_ARCHIVE_HEADERS_);

  ensureSheetWithHeaders_(
    spreadsheet,
    "MonthlySummary",
    MONTHLY_SUMMARY_HEADERS_,
  );

  ensureSheetWithHeaders_(
    spreadsheet,
    "MonthlyItemSummary",
    MONTHLY_ITEM_SUMMARY_HEADERS_,
  );

  return {
    created: true,
    sheets: ["OrdersArchive", "MonthlySummary", "MonthlyItemSummary"],
  };
}

function normalizeMonthKey_(value) {
  const match = String(value || "")
    .trim()
    .match(/^(20\d{2})[\/\-](0?[1-9]|1[0-2])$/);

  if (!match) {
    return null;
  }

  return match[1] + "/" + String(Number(match[2])).padStart(2, "0");
}

function getPreviousMonthKey_(now = new Date()) {
  const currentMonth = Utilities.formatDate(now, "Asia/Taipei", "yyyy/MM");

  const parts = currentMonth.split("/");

  const previous = new Date(
    Date.UTC(Number(parts[0]), Number(parts[1]) - 2, 1),
  );

  return (
    previous.getUTCFullYear() +
    "/" +
    String(previous.getUTCMonth() + 1).padStart(2, "0")
  );
}

function validateArchiveMonth_(monthKey) {
  const normalized = normalizeMonthKey_(monthKey);

  if (!normalized) {
    throw new Error("封存月份格式必須為 yyyy/MM");
  }

  const currentMonth = Utilities.formatDate(
    new Date(),
    "Asia/Taipei",
    "yyyy/MM",
  );

  if (normalized >= currentMonth) {
    throw new Error("只能封存目前月份以前的訂單");
  }

  return normalized;
}

function buildMonthlyArchiveData_(values, monthKey, priceMap) {
  const orderGroups = Object.create(null);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const systemOrderId = String(row[0] || "").trim();

    if (!systemOrderId) {
      continue;
    }

    const createdAt = parseOrderDate_(row[2]);

    if (!createdAt) {
      continue;
    }

    const rowMonth = Utilities.formatDate(createdAt, "Asia/Taipei", "yyyy/MM");

    if (rowMonth !== monthKey) {
      continue;
    }

    if (!orderGroups[systemOrderId]) {
      orderGroups[systemOrderId] = {
        systemOrderId,
        displayOrderId: String(row[1] || ""),
        rows: [],
        rowNumbers: [],
      };
    }

    orderGroups[systemOrderId].rows.push(row);

    orderGroups[systemOrderId].rowNumbers.push(i + 1);
  }

  const eligibleOrders = Object.values(orderGroups).filter(
    (order) =>
      order.rows.length > 0 &&
      order.rows.every((row) => String(row[10] || "").trim() === "已完成"),
  );

  let revenue = 0;
  let unpricedOrderCount = 0;

  const archiveRows = [];
  const itemMap = Object.create(null);

  eligibleOrders.forEach((order) => {
    let orderHasMissingPrice = false;

    order.rows.forEach((row) => {
      const itemName = String(row[6] || "").trim();

      const quantity = Number(row[7]) || 0;

      const priceInfo = resolveOrderPrice_(row, priceMap);

      const archiveRow = row.slice(0, 14);

      if (
        archiveRow[13] === "" ||
        archiveRow[13] === null ||
        archiveRow[13] === undefined
      ) {
        archiveRow[13] = priceInfo.missing ? "UNPRICED" : priceInfo.unitPrice;
      }

      archiveRows.push(archiveRow);

      if (priceInfo.missing) {
        orderHasMissingPrice = true;
      }

      revenue += priceInfo.lineTotal;

      if (!itemName) {
        return;
      }

      if (!itemMap[itemName]) {
        itemMap[itemName] = {
          month: monthKey,
          itemName,
          quantity: 0,
          revenue: 0,
        };
      }

      itemMap[itemName].quantity += quantity > 10 ? 1 : quantity;

      itemMap[itemName].revenue += priceInfo.lineTotal;
    });

    if (orderHasMissingPrice) {
      unpricedOrderCount++;
    }
  });

  const orderCount = eligibleOrders.length;

  return {
    month: monthKey,
    eligibleOrderIds: eligibleOrders.map((order) => order.systemOrderId),
    eligibleOrders: eligibleOrders.map((order) => ({
      systemOrderId: order.systemOrderId,
      displayOrderId: order.displayOrderId,
      rowNumbers: order.rowNumbers.slice(),
    })),
    archiveRows,
    summary: {
      month: monthKey,
      orderCount,
      completedCount: orderCount,
      revenue: Math.round(revenue),
      unpricedOrderCount,
      averageOrderValue: orderCount ? Math.round(revenue / orderCount) : 0,
    },
    itemSummaries: Object.values(itemMap).sort((a, b) =>
      a.itemName.localeCompare(b.itemName, "zh-Hant"),
    ),
  };
}

function previewMonthlyArchive(monthKey = getPreviousMonthKey_()) {
  const targetMonth = validateArchiveMonth_(monthKey);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");

  const result = buildMonthlyArchiveData_(
    sheet.getDataRange().getValues(),
    targetMonth,
    getPriceMap(),
  );

  return {
    preview: true,
    month: targetMonth,
    orderCount: result.summary.orderCount,
    rowCount: result.archiveRows.length,
    revenue: result.summary.revenue,
    unpricedOrderCount: result.summary.unpricedOrderCount,
    orders: result.eligibleOrders.slice(0, 50),
  };
}

function upsertMonthlySummary_(sheet, summary, archivedAt) {
  const values = sheet.getDataRange().getValues();

  let rowNumber = sheet.getLastRow() + 1;

  for (let i = 1; i < values.length; i++) {
    if (normalizeMonthKey_(values[i][0]) === summary.month) {
      rowNumber = i + 1;
      break;
    }
  }

  sheet
    .getRange(rowNumber, 1, 1, MONTHLY_SUMMARY_HEADERS_.length)
    .setValues([
      [
        summary.month,
        summary.orderCount,
        summary.completedCount,
        summary.revenue,
        summary.unpricedOrderCount,
        summary.averageOrderValue,
        archivedAt,
      ],
    ]);
}

function replaceMonthlyItemSummary_(sheet, monthKey, itemSummaries) {
  const values = sheet.getDataRange().getValues();

  for (let i = values.length - 1; i >= 1; i--) {
    if (normalizeMonthKey_(values[i][0]) === monthKey) {
      sheet.deleteRow(i + 1);
    }
  }

  if (itemSummaries.length === 0) {
    return;
  }

  const rows = itemSummaries.map((item) => [
    monthKey,
    item.itemName,
    item.quantity,
    Math.round(item.revenue),
  ]);

  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      rows.length,
      MONTHLY_ITEM_SUMMARY_HEADERS_.length,
    )
    .setValues(rows);
}

function deleteOrderRows_(sheet, rowNumbers) {
  const sorted = Array.from(new Set(rowNumbers)).sort((a, b) => b - a);

  if (sorted.length === 0) {
    return;
  }

  let high = sorted[0];
  let low = high;

  for (let i = 1; i <= sorted.length; i++) {
    const rowNumber = sorted[i];

    if (rowNumber === low - 1) {
      low = rowNumber;
      continue;
    }

    sheet.deleteRows(low, high - low + 1);

    high = rowNumber;
    low = rowNumber;
  }
}

function archiveMonthlyOrders(monthKey = getPreviousMonthKey_()) {
  const targetMonth = validateArchiveMonth_(monthKey);

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    const ordersSheet = spreadsheet.getSheetByName("Orders");

    const archiveSheet = ensureSheetWithHeaders_(
      spreadsheet,
      "OrdersArchive",
      ORDER_ARCHIVE_HEADERS_,
    );

    const summarySheet = ensureSheetWithHeaders_(
      spreadsheet,
      "MonthlySummary",
      MONTHLY_SUMMARY_HEADERS_,
    );

    const itemSummarySheet = ensureSheetWithHeaders_(
      spreadsheet,
      "MonthlyItemSummary",
      MONTHLY_ITEM_SUMMARY_HEADERS_,
    );

    const priceMap = getPriceMap();

    const sourceResult = buildMonthlyArchiveData_(
      ordersSheet.getDataRange().getValues(),
      targetMonth,
      priceMap,
    );

    const archiveValues = archiveSheet.getDataRange().getValues();

    const archivedOrderIds = new Set();

    for (let i = 1; i < archiveValues.length; i++) {
      const systemOrderId = String(archiveValues[i][0] || "").trim();

      if (systemOrderId) {
        archivedOrderIds.add(systemOrderId);
      }
    }

    const archivedAt = new Date();

    const rowsToAppend = [];

    sourceResult.eligibleOrders.forEach((order) => {
      if (archivedOrderIds.has(order.systemOrderId)) {
        return;
      }

      sourceResult.archiveRows
        .filter((row) => String(row[0] || "") === order.systemOrderId)
        .forEach((row) => rowsToAppend.push(row.concat([archivedAt])));

      archivedOrderIds.add(order.systemOrderId);
    });

    if (rowsToAppend.length > 0) {
      archiveSheet
        .getRange(
          archiveSheet.getLastRow() + 1,
          1,
          rowsToAppend.length,
          ORDER_ARCHIVE_HEADERS_.length,
        )
        .setValues(rowsToAppend);
    }

    const refreshedArchiveResult = buildMonthlyArchiveData_(
      archiveSheet.getDataRange().getValues(),
      targetMonth,
      priceMap,
    );

    upsertMonthlySummary_(
      summarySheet,
      refreshedArchiveResult.summary,
      archivedAt,
    );

    replaceMonthlyItemSummary_(
      itemSummarySheet,
      targetMonth,
      refreshedArchiveResult.itemSummaries,
    );

    const rowsToDelete = [];

    sourceResult.eligibleOrders.forEach((order) => {
      if (archivedOrderIds.has(order.systemOrderId)) {
        rowsToDelete.push.apply(rowsToDelete, order.rowNumbers);
      }
    });

    deleteOrderRows_(ordersSheet, rowsToDelete);

    return {
      archived: true,
      month: targetMonth,
      archivedOrders: sourceResult.eligibleOrders.length,
      appendedRows: rowsToAppend.length,
      deletedRows: rowsToDelete.length,
      summary: refreshedArchiveResult.summary,
    };
  } finally {
    lock.releaseLock();
  }
}

function archivePreviousMonthOrders() {
  return archiveMonthlyOrders(getPreviousMonthKey_());
}

function archiveAllPastCompletedOrders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");

  const values = sheet.getDataRange().getValues();

  const currentMonth = Utilities.formatDate(
    new Date(),
    "Asia/Taipei",
    "yyyy/MM",
  );

  const months = new Set();

  for (let i = 1; i < values.length; i++) {
    const createdAt = parseOrderDate_(values[i][2]);

    if (!createdAt) {
      continue;
    }

    const month = Utilities.formatDate(createdAt, "Asia/Taipei", "yyyy/MM");

    if (month < currentMonth) {
      months.add(month);
    }
  }

  const results = Array.from(months)
    .sort()
    .map((month) => archiveMonthlyOrders(month));

  return {
    archived: true,
    checkedMonths: results.length,
    results,
  };
}

function installOrderArchiveTrigger() {
  const handler = "archiveAllPastCompletedOrders";

  ScriptApp.getProjectTriggers()
    .filter((trigger) =>
      [handler, "archivePreviousMonthOrders"].includes(
        trigger.getHandlerFunction(),
      ),
    )
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  const trigger = ScriptApp.newTrigger(handler)
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();

  return {
    enabled: true,
    handler,
    everyDays: 1,
    hour: 3,
    triggerId: trigger.getUniqueId(),
  };
}

function installMonthlyArchiveTrigger() {
  return installOrderArchiveTrigger();
}

function getOrderArchiveTriggerStatus() {
  const handler = "archiveAllPastCompletedOrders";

  const triggers = ScriptApp.getProjectTriggers().filter(
    (trigger) => trigger.getHandlerFunction() === handler,
  );

  return {
    enabled: triggers.length > 0,
    handler,
    everyDays: 1,
    hour: 3,
    triggerCount: triggers.length,
  };
}

function removeOrderArchiveTrigger() {
  const handlers = [
    "archiveAllPastCompletedOrders",
    "archivePreviousMonthOrders",
  ];

  let removed = 0;

  ScriptApp.getProjectTriggers()
    .filter((trigger) => handlers.includes(trigger.getHandlerFunction()))
    .forEach((trigger) => {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    });

  return {
    enabled: false,
    removed,
  };
}

function getDashboardOrderRows_(monthKey) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const ordersSheet = spreadsheet.getSheetByName("Orders");

  const archiveSheet = spreadsheet.getSheetByName("OrdersArchive");

  const currentMonth = Utilities.formatDate(
    new Date(),
    "Asia/Taipei",
    "yyyy/MM",
  );

  const orderValues = ordersSheet ? ordersSheet.getDataRange().getValues() : [];

  const archiveValues =
    archiveSheet && monthKey !== currentMonth
      ? archiveSheet.getDataRange().getValues()
      : [];

  const archivedIds = new Set();

  const archiveRows = [];

  for (let i = 1; i < archiveValues.length; i++) {
    const row = archiveValues[i];
    const createdAt = parseOrderDate_(row[2]);

    if (
      !createdAt ||
      Utilities.formatDate(createdAt, "Asia/Taipei", "yyyy/MM") !== monthKey
    ) {
      continue;
    }

    const systemOrderId = String(row[0] || "").trim();

    if (systemOrderId) {
      archivedIds.add(systemOrderId);
      archiveRows.push(row.slice(0, 14));
    }
  }

  const activeRows = [];

  for (let i = 1; i < orderValues.length; i++) {
    const row = orderValues[i];
    const systemOrderId = String(row[0] || "").trim();

    if (!systemOrderId || archivedIds.has(systemOrderId)) {
      continue;
    }

    const createdAt = parseOrderDate_(row[2]);

    if (
      createdAt &&
      Utilities.formatDate(createdAt, "Asia/Taipei", "yyyy/MM") === monthKey
    ) {
      activeRows.push(row.slice(0, 14));
    }
  }

  return [ORDER_ARCHIVE_HEADERS_.slice(0, 14)].concat(archiveRows, activeRows);
}

function getAvailableDashboardMonths_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const months = new Set([
    Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM"),
  ]);

  ["Orders"].forEach((sheetName) => {
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      return;
    }

    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      const createdAt = parseOrderDate_(values[i][2]);

      if (createdAt) {
        months.add(Utilities.formatDate(createdAt, "Asia/Taipei", "yyyy/MM"));
      }
    }
  });

  const summarySheet = spreadsheet.getSheetByName("MonthlySummary");

  if (summarySheet) {
    const values = summarySheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      const month = normalizeMonthKey_(values[i][0]);

      if (month) {
        months.add(month);
      }
    }
  }

  return Array.from(months).sort((a, b) => b.localeCompare(a));
}

function getAdminDashboardData(selectedMonth) {
  const priceMap = getPriceMap();

  const timeZone = "Asia/Taipei";

  const now = new Date();

  const todayKey = Utilities.formatDate(now, timeZone, "yyyy/MM/dd");

  const currentMonthKey = Utilities.formatDate(now, timeZone, "yyyy/MM");

  const requestedMonth = normalizeMonthKey_(selectedMonth);

  const monthKey = requestedMonth || currentMonthKey;

  const values = getDashboardOrderRows_(monthKey);

  const orders = {};
  const itemStats = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const systemOrderId = String(row[0] || "");

    if (!systemOrderId) {
      continue;
    }

    const createdAt = parseOrderDate_(row[2]);

    if (!createdAt) {
      continue;
    }

    if (!orders[systemOrderId]) {
      orders[systemOrderId] = {
        orderId: String(row[1] || ""),
        createdAt: createdAt,
        customer: String(row[3] || "未填寫"),
        pickup: String(row[8] || ""),
        status: String(row[10] || ""),
        total: 0,
        itemCount: 0,
        priceMissing: false,
      };
    }

    const itemName = String(row[6] || "");

    const quantity = Number(row[7]) || 0;

    const priceInfo = resolveOrderPrice_(row, priceMap);

    const lineTotal = priceInfo.lineTotal;

    const itemQuantity = quantity > 10 ? 1 : quantity;

    if (priceInfo.missing) {
      orders[systemOrderId].priceMissing = true;
    }

    orders[systemOrderId].total += lineTotal;

    orders[systemOrderId].itemCount += itemQuantity;

    const orderMonthKey = Utilities.formatDate(createdAt, timeZone, "yyyy/MM");

    if (orderMonthKey === monthKey && itemName) {
      if (!itemStats[itemName]) {
        itemStats[itemName] = {
          name: itemName,
          quantity: 0,
          revenue: 0,
        };
      }

      itemStats[itemName].quantity += itemQuantity;

      itemStats[itemName].revenue += lineTotal;
    }
  }

  const allOrders = Object.values(orders);

  const todayOrders = allOrders.filter(
    (order) =>
      Utilities.formatDate(order.createdAt, timeZone, "yyyy/MM/dd") ===
      todayKey,
  );

  const monthOrders = allOrders.filter(
    (order) =>
      Utilities.formatDate(order.createdAt, timeZone, "yyyy/MM") === monthKey,
  );

  const daysInMonth =
    monthKey === currentMonthKey
      ? Number(Utilities.formatDate(now, timeZone, "d"))
      : (() => {
          const parts = monthKey.split("/");

          return new Date(
            Date.UTC(Number(parts[0]), Number(parts[1]), 0),
          ).getUTCDate();
        })();

  const dailyMap = {};

  for (let day = 1; day <= daysInMonth; day++) {
    const key = monthKey + "/" + String(day).padStart(2, "0");

    dailyMap[key] = {
      label: day + "日",
      orders: 0,
      revenue: 0,
    };
  }

  monthOrders.forEach((order) => {
    const key = Utilities.formatDate(order.createdAt, timeZone, "yyyy/MM/dd");

    if (dailyMap[key]) {
      dailyMap[key].orders++;
      dailyMap[key].revenue += order.total;
    }
  });

  const summarize = (orderList) => {
    const revenue = orderList.reduce((sum, order) => sum + order.total, 0);

    const completed = orderList.filter(
      (order) => order.status === "已完成",
    ).length;

    const pending = orderList.filter(
      (order) => order.status === "待製作",
    ).length;

    return {
      orders: orderList.length,
      revenue: Math.round(revenue),
      average: orderList.length ? Math.round(revenue / orderList.length) : 0,
      completed,
      pending,
      completionRate: orderList.length
        ? Math.round((completed / orderList.length) * 100)
        : 0,
    };
  };

  const recentOrders = (
    monthKey === currentMonthKey ? todayOrders : monthOrders
  )
    .sort((a, b) => b.createdAt.getTime() - a.createdAt.getTime())
    .slice(0, 10)
    .map((order) => ({
      orderId: order.orderId,
      time: Utilities.formatDate(
        order.createdAt,
        timeZone,
        monthKey === currentMonthKey ? "HH:mm" : "MM/dd HH:mm",
      ),
      customer: order.customer,
      pickup: order.pickup,
      status: order.status,
      total: Math.round(order.total),
      priceMissing: order.priceMissing,
    }));

  return {
    selectedMonth: monthKey,
    currentMonth: currentMonthKey,
    isCurrentMonth: monthKey === currentMonthKey,
    availableMonths: getAvailableDashboardMonths_(),
    generatedAt: Utilities.formatDate(now, timeZone, "yyyy/MM/dd HH:mm:ss"),
    todayLabel:
      monthKey === currentMonthKey
        ? Utilities.formatDate(now, timeZone, "MM/dd")
        : monthKey,
    monthLabel: monthKey,
    today: summarize(todayOrders),
    month: summarize(monthOrders),
    daily: Object.values(dailyMap),
    topItems: Object.values(itemStats)
      .sort((a, b) => b.quantity - a.quantity)
      .slice(0, 8)
      .map((item) => ({
        name: item.name,
        quantity: item.quantity,
        revenue: Math.round(item.revenue),
      })),
    recentOrders,
  };
}

function parseOrderDate_(value) {
  if (
    Object.prototype.toString.call(value) === "[object Date]" &&
    !isNaN(value.getTime())
  ) {
    return value;
  }

  const match = String(value || "").match(
    /^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/,
  );

  if (!match) {
    return null;
  }

  return new Date(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
    Number(match[4] || 0),
    Number(match[5] || 0),
    Number(match[6] || 0),
  );
}

function getRequiredScriptProperty_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);

  if (!value) {
    throw new Error(`缺少 Script Property: ${key}`);
  }

  return value;
}

function claimWebhookEvent_(webhookEventId) {
  if (!webhookEventId) {
    return true;
  }

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    let sheet = spreadsheet.getSheetByName("WebhookEvents");

    if (!sheet) {
      sheet = spreadsheet.insertSheet("WebhookEvents");

      sheet
        .getRange(1, 1, 1, 2)
        .setValues([["webhook_event_id", "processed_at"]]);
    }

    const lastRow = Math.max(sheet.getLastRow(), 1);

    const existing = sheet
      .getRange(1, 1, lastRow, 1)
      .createTextFinder(webhookEventId)
      .matchEntireCell(true)
      .findNext();

    if (existing) {
      return false;
    }

    sheet
      .getRange(lastRow + 1, 1, 1, 2)
      .setValues([[webhookEventId, new Date()]]);

    return true;
  } finally {
    lock.releaseLock();
  }
}

function releaseWebhookEvent_(webhookEventId) {
  if (!webhookEventId) {
    return;
  }

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("WebhookEvents");

    if (!sheet) {
      return;
    }

    const existing = sheet
      .getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1)
      .createTextFinder(webhookEventId)
      .matchEntireCell(true)
      .findNext();

    if (existing && existing.getRow() > 1) {
      sheet.deleteRow(existing.getRow());
    }
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const page =
    e && e.parameter && e.parameter.page === "menu" ? "menu" : "dashboard";

  return HtmlService.createHtmlOutputFromFile(page);
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  const expectedSecret = getRequiredScriptProperty_("WEBHOOK_SHARED_SECRET");

  if (!data.webhookSecret || data.webhookSecret !== expectedSecret) {
    return ContentService.createTextOutput(
      JSON.stringify({
        error: "UNAUTHORIZED",
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const message = data.message || "";

  const lineUserId = data.lineUserId || "";

  const lineName = data.customerName || "";

  const pictureUrl = data.pictureUrl || "";

  const webhookEventId = String(data.webhookEventId || "");

  saveCustomer(lineUserId, lineName, pictureUrl);

  const order = parseOrder(message);

  if (!order.is_order) {
    return ContentService.createTextOutput(
      JSON.stringify({
        is_order: false,
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  }

  order.customer_name = order.customer_name || lineName || "";

  const normalizationResult = normalizeItems(order);

  const soldOutItems = checkSoldOutItems(order);

  if (soldOutItems.length > 0) {
    return ContentService.createTextOutput(
      JSON.stringify({
        is_order: true,
        soldOut: true,
        soldOutItems: soldOutItems,
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const priceInfo = calculateOrderPrice(order);

  const editable = findEditableOrder(lineUserId);

  const waitMinutes = Math.max(
    getEstimatedWaitMinutes() + (editable ? 0 : 5),
    10,
  );

  const pickupInfo = ensureOrderPickupTime_(order, waitMinutes);

  const eventClaimed = claimWebhookEvent_(webhookEventId);

  if (!eventClaimed) {
    return ContentService.createTextOutput(
      JSON.stringify({
        duplicate: true,
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  }

  let orderPersisted = false;

  try {
    let result;

    if (editable) {
      result = updateOrder(
        editable,
        order,
        message,
        lineUserId,
        pickupInfo.displayTime,
      );

      orderPersisted = true;
    } else {
      const saveResult = saveOrder(order, message, lineUserId);

      orderPersisted = true;

      incrementCustomerOrderCount(lineUserId);

      result = {
        is_order: true,

        updated: false,

        orderId: saveResult.displayOrderId,

        order,

        totalPrice: priceInfo.totalPrice,

        priceMissing: priceInfo.priceMissing,

        missingItems: priceInfo.missingItems,

        estimatedPickupTime: pickupInfo.displayTime,
      };
    }

    logUnmatchedItemsSafely_(normalizationResult.unmatchedItems, message);

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(
      ContentService.MimeType.JSON,
    );
  } catch (error) {
    if (!orderPersisted && !error.orderPersisted) {
      releaseWebhookEvent_(webhookEventId);
    }

    throw error;
  }
}

// function doGet() {
//   return ContentService
//     .createTextOutput("OK");
// }

//-------------------------------------------------------------------------------------------------------------------------------------
//-------------------------------------------------------------------------------------------------------------------------------------
function getMenuItems() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PriceList");

  const values = sheet.getDataRange().getValues();

  const items = [];

  for (let i = 1; i < values.length; i++) {
    items.push({
      rowIndex: i + 1,
      name: values[i][0],
      price: values[i][1],
      category: values[i][2] || "未分類",
      imageUrl: values[i][3] || "",
      isSoldOut: values[i][4] === true,
    });
  }

  return items;
}

function updateSoldOut(rowIndex, isSoldOut) {
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("PriceList")
    .getRange(rowIndex, 5)
    .setValue(isSoldOut);

  return true;
}

function getSoldOutItemNames() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PriceList");

  const values = sheet.getDataRange().getValues();

  const soldOutItems = [];

  for (let i = 1; i < values.length; i++) {
    const itemName = values[i][0];

    const isSoldOut = values[i][4] === true;

    if (isSoldOut) {
      soldOutItems.push(itemName);
    }
  }

  return soldOutItems;
}

function checkSoldOutItems(order) {
  const soldOutItems = getSoldOutItemNames();

  const found = [];

  order.groups.forEach((group) => {
    group.items.forEach((item) => {
      if (soldOutItems.includes(item.name) && !found.includes(item.name)) {
        found.push(item.name);
      }
    });
  });

  return found;
}

function resetSoldOutStatus() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PriceList");

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return;
  }

  sheet.getRange(2, 5, lastRow - 1, 1).setValue(false);
}

// function createResetSoldOutTrigger(){

//   ScriptApp
//     .newTrigger("resetSoldOutStatus")
//     .timeBased()
//     .everyDays(1)
//     .atHour(6)
//     .create();

// }

//-------------------------------------------------------------------------------------------------------------------------------------
//-------------------------------------------------------------------------------------------------------------------------------------
//-------------------------------------------------------------------------------------------------------------------------------------
//-------------------------------------------------------------------------------------------------------------------------------------
function testOpenAI() {
  const apiKey =
    PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  const payload = {
    model: "gpt-4.1-mini",
    messages: [
      {
        role: "user",
        content: "請回傳你好",
      },
    ],
  };

  const response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/chat/completions",
    {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: `Bearer ${apiKey}`,
      },
      payload: JSON.stringify(payload),
    },
  );

  Logger.log(response.getContentText());
}

function testGetOrders() {
  const result = getOrders();

  console.log(result);
}

function testParseOrder() {
  //   const message = `
  // 王小姐
  // 雞排兩份
  // 微辣
  // 18:00取餐
  // 電話0912345678
  // `;
  const message = `
老闆娘妳好！
我需要以下，謝謝

1.
高麗菜
杏包菇x2
雞心
百葉豆腐20
雞胸肉

做熱的、（小小辣）、不酸菜
`;

  const order = parseOrder(message);

  Logger.log(JSON.stringify(order, null, 2));
}

function testFlow() {
  const message = `
老闆娘妳好！
我需要以下，謝謝

1.王先生
高麗菜
杏包菇x2
雞心
百葉豆腐20
雞胸肉

做熱的、（小小辣）、不酸菜

2.吳小姐
高麗菜
杏包菇
百葉豆腐20
王子麵（另外裝）

做熱的、不酸菜、不辣
`;

  const order = parseOrder(message);

  saveOrder(order, message);

  Logger.log(order);
}
