const SYSTEM_PROMPT = `
你是滷味訂單解析器。輸入包含 current_message 與 existing_order。

只依 current_message 判斷 is_order：包含商品、數量、製作備註、分袋或取餐時間為 true；詢價、營業或菜單詢問、聊天、感謝、招呼、取消為 false。existing_order 不得使非訂單訊息變成訂單。

is_order=true 時回傳處理後的完整訂單：
- existing_order=null：update_mode="replace"。
- 有 existing_order，且本次是無「再、加、追加、改、第幾份」等承接語的完整清單：replace。
- 追加或局部修改：update_mode="merge"，保留未修改的商品、備註、客戶資料與 pickup_time。
- 「再一份／另外一份」在同一訂單新增 group，不建立第二張訂單；未指定商品則複製最近 group，再套用新備註。
- 第一袋、第二袋、分開裝、編號或不同備註代表不同 groups；指定第幾份時只修改該 group。單一「一般」group 被追加時依序改名第一份、第二份。

items 規則：
- quantity 為數字；半份／半個=0.5，一份半=1.5，兩份半=2.5。
- raw_name 保留客戶原商品文字且不含數量。
- 能明確對應 menu 時 name 使用完全相同的正式名稱；否則 name=raw_name，suggested_name 可填最接近的 menu 名稱，無候選填空字串。
- 不可創造 menu 以外的正式名稱。

pickup_time：客人有指定才轉成 Asia/Taipei 的 yyyy/MM/dd HH:mm；只有時間用今天，日期語意須正確換算。未指定時保留 existing_order.pickup_time，無 existing_order 則填空字串。

僅回傳 JSON，不要 markdown 或解釋。空缺字串填 ""。
格式：{"is_order":true,"update_mode":"replace|merge|none","customer_name":"","phone":"","pickup_time":"","groups":[{"name":"","items":[{"name":"","raw_name":"","suggested_name":"","quantity":1}],"note":""}]}
   `;

const AI_ORDER_ENABLED_PROPERTY_ = "AI_ORDER_ENABLED";
const AI_ORDER_UPDATED_AT_PROPERTY_ = "AI_ORDER_UPDATED_AT";

function isAiOrderEnabled_() {
  const value = PropertiesService.getScriptProperties().getProperty(
    AI_ORDER_ENABLED_PROPERTY_,
  );

  return String(value || "true").toLowerCase() !== "false";
}

function getAiOrderStatus() {
  const properties = PropertiesService.getScriptProperties();

  return {
    enabled: isAiOrderEnabled_(),
    updatedAt: properties.getProperty(AI_ORDER_UPDATED_AT_PROPERTY_) || "",
  };
}

function setAiOrderEnabled(enabled) {
  if (typeof enabled !== "boolean") {
    throw new Error("AI 接單狀態必須是布林值");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const properties = PropertiesService.getScriptProperties();
    const updatedAt = Utilities.formatDate(
      new Date(),
      "Asia/Taipei",
      "yyyy/MM/dd HH:mm:ss",
    );

    properties.setProperties({
      [AI_ORDER_ENABLED_PROPERTY_]: enabled ? "true" : "false",
      [AI_ORDER_UPDATED_AT_PROPERTY_]: updatedAt,
    });

    return {
      enabled,
      updatedAt,
    };
  } finally {
    lock.releaseLock();
  }
}

function assertAiOrderEnabledForWrite_() {
  if (isAiOrderEnabled_()) {
    return;
  }

  const error = new Error("AI_ORDER_DISABLED");
  error.code = "AI_ORDER_DISABLED";
  throw error;
}

function createAssistantDisabledOutput_() {
  return ContentService.createTextOutput(
    JSON.stringify({
      is_order: false,
      assistant_disabled: true,
    }),
  ).setMimeType(ContentService.MimeType.JSON);
}

function parseOrder(message, editableOrder = null) {
  const apiKey =
    PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  const menuItemNames = getCanonicalItemNames_();

  const currentTime = Utilities.formatDate(
    new Date(),
    "Asia/Taipei",
    "yyyy/MM/dd HH:mm",
  );

  const menuPrompt = `
now=${currentTime} (Asia/Taipei)
menu=${JSON.stringify(menuItemNames)}
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
        content:
          "input=" +
          JSON.stringify({
            current_message: String(message || ""),
            existing_order:
              editableOrder && editableOrder.order ? editableOrder.order : null,
          }),
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
  requireAiOrderEnabled = false,
) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");

  const priceMap = getPriceMap();

  const priceColumn = 14;
  const pickupTimeSourceColumn = 15;

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

        normalizePickupTimeSource_(order.pickup_time_source),
      ]);
    });
  });

  if (rows.length === 0) {
    throw new Error("訂單格式錯誤：沒有可寫入的商品");
  }

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    if (requireAiOrderEnabled) {
      assertAiOrderEnabledForWrite_();
    }

    if (!sheet.getRange(1, priceColumn).getValue()) {
      sheet.getRange(1, priceColumn).setValue("unit_price");
    }

    if (!sheet.getRange(1, pickupTimeSourceColumn).getValue()) {
      sheet
        .getRange(1, pickupTimeSourceColumn)
        .setValue("pickup_time_source");
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

function initializePickupTimeSourceColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    sheet.getRange(1, 15).setValue("pickup_time_source");

    if (sheet.getLastRow() <= 1) {
      return {
        initialized: true,
        updatedRows: 0,
      };
    }

    const values = sheet.getDataRange().getDisplayValues();
    let updatedRows = 0;

    const sources = values.slice(1).map((row) => {
      const existingSource = normalizePickupTimeSource_(row[14]);

      if (existingSource || !row[8]) {
        return [existingSource];
      }

      updatedRows++;

      return [
        messageUpdatesPickupTime_(row[11]) ? "requested" : "estimated",
      ];
    });

    sheet.getRange(2, 15, sources.length, 1).setValues(sources);

    return {
      initialized: true,
      updatedRows,
    };
  } finally {
    lock.releaseLock();
  }
}

function manualCreateOrder(message) {
  const order = parseOrder(message);

  if (!order.is_order) {
    throw new Error("訊息無法解析為訂單");
  }

  const normalizationResult = normalizeItems(order);

  ensureOrderPickupTime_(
    order,
    getFinalEstimatedWaitMinutes_(getEstimatedWaitMinutes(), order),
  );

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

  return {
    systemOrderId,
    displayOrderId,
    notified: false,
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
  const manualAliasKeys = Object.create(null);

  if (sheet) {
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      const alias = String(values[i][0] || "").trim();

      const canonicalName = String(values[i][1] || "").trim();

      if (alias && canonicalName) {
        map[alias] = canonicalName;
        manualAliasKeys[normalizeItemKey_(alias)] = true;
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

      if (
        alias &&
        canonicalName &&
        !manualAliasKeys[normalizeItemKey_(alias)]
      ) {
        map[alias] = canonicalName;
      }
    }
  }

  return map;
}

function getPreparationItemCount_(quantityValue) {
  const quantity = Number(quantityValue);

  if (!Number.isFinite(quantity) || quantity <= 0) {
    return 1;
  }

  // 「品項20」是既有的金額式輸入，製作量仍只算一項。
  if (quantity > 10) {
    return 1;
  }

  return Math.max(1, Math.ceil(quantity));
}

function getOrderPreparationMinutes_(order) {
  if (!order || !Array.isArray(order.groups)) {
    return 0;
  }

  let itemCount = 0;

  order.groups.forEach((group) => {
    if (!group || !Array.isArray(group.items)) {
      return;
    }

    group.items.forEach((item) => {
      itemCount += getPreparationItemCount_(item && item.quantity);
    });
  });

  return itemCount * 0.5;
}

const MINIMUM_ESTIMATED_WAIT_MINUTES_ = 10;
const ORDER_HANDLING_MINUTES_ = 1.5;

function getFinalEstimatedWaitMinutes_(queueMinutes, order) {
  const calculatedMinutes =
    Number(queueMinutes) +
    getOrderPreparationMinutes_(order) +
    ORDER_HANDLING_MINUTES_;

  return Math.max(
    Math.ceil(calculatedMinutes),
    MINIMUM_ESTIMATED_WAIT_MINUTES_,
  );
}

function getEstimatedWaitMinutes(excludedSystemOrderId = "") {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");

  const values = sheet.getDataRange().getDisplayValues();

  let itemCount = 0;

  for (let i = 1; i < values.length; i++) {
    if (values[i][10] !== "待製作") {
      continue;
    }

    if (
      excludedSystemOrderId &&
      String(values[i][0]) === String(excludedSystemOrderId)
    ) {
      continue;
    }

    itemCount += getPreparationItemCount_(values[i][7]);
  }

  return itemCount * 0.5;
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

function normalizePickupTimeSource_(value) {
  const source = String(value || "")
    .trim()
    .toLowerCase();

  return source === "requested" || source === "estimated" ? source : "";
}

function ensureOrderPickupTime_(order, waitMinutes, now = new Date()) {
  const requestedPickupTime = String(order.pickup_time || "").trim();

  if (requestedPickupTime) {
    const pickupTimeSource =
      normalizePickupTimeSource_(order.pickup_time_source) || "requested";

    order.pickup_time = requestedPickupTime;
    order.pickup_time_source = pickupTimeSource;

    return {
      pickupTime: requestedPickupTime,
      displayTime: getPickupTimeDisplay_(requestedPickupTime, now),
      estimated: pickupTimeSource === "estimated",
      pickupTimeSource,
    };
  }

  const estimatedPickupTime = getEstimatedPickupDateTime_(waitMinutes, now);

  order.pickup_time = estimatedPickupTime;
  order.pickup_time_source = "estimated";

  return {
    pickupTime: estimatedPickupTime,
    displayTime: getPickupTimeDisplay_(estimatedPickupTime, now),
    estimated: true,
    pickupTimeSource: "estimated",
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

function resolveAliasItemName_(value, normalizationData) {
  const name = String(value || "").trim();

  if (!name) {
    return null;
  }

  if (normalizationData.aliasExact[name]) {
    return normalizationData.aliasExact[name];
  }

  return normalizationData.aliasNormalized[normalizeItemKey_(name)] || null;
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
        resolveAliasItemName_(rawName, normalizationData) ||
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

function buildExistingOrderContext_(values, systemOrderId) {
  const order = {
    customer_name: "",
    phone: "",
    pickup_time: "",
    pickup_time_source: "",
    groups: [],
  };

  const groupIndexes = {};

  let rawMessage = "";

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    if (String(row[0]) !== String(systemOrderId)) {
      continue;
    }

    order.customer_name = order.customer_name || row[3] || "";
    order.phone = order.phone || row[4] || "";
    order.pickup_time = order.pickup_time || row[8] || "";
    order.pickup_time_source =
      order.pickup_time_source || normalizePickupTimeSource_(row[14]);
    rawMessage = rawMessage || row[11] || "";

    const groupName = String(row[5] || "一般");
    const groupNote = String(row[9] || "");
    const groupKey = groupName + "\u0000" + groupNote;

    if (groupIndexes[groupKey] === undefined) {
      groupIndexes[groupKey] = order.groups.length;
      order.groups.push({
        name: groupName,
        items: [],
        note: groupNote,
      });
    }

    if (row[6]) {
      order.groups[groupIndexes[groupKey]].items.push({
        name: String(row[6]),
        raw_name: String(row[6]),
        suggested_name: "",
        quantity: Number(row[7]) || 0,
      });
    }
  }

  if (!order.pickup_time_source && order.pickup_time) {
    order.pickup_time_source = messageUpdatesPickupTime_(rawMessage)
      ? "requested"
      : "estimated";
  }

  return {
    order,
    rawMessage: String(rawMessage || ""),
  };
}

function appendRawMessage_(previousMessage, currentMessage) {
  const previous = String(previousMessage || "").trim();
  const current = String(currentMessage || "").trim();

  if (!previous) {
    return current;
  }

  if (!current) {
    return previous;
  }

  return previous + "\n\n" + current;
}

function messageUpdatesPickupTime_(message) {
  const text = String(message || "").trim();

  if (!text) {
    return false;
  }

  return (
    /(取餐|取貨|拿餐|領餐|來拿|取的時間|拿的時間)/.test(text) ||
    /\d{1,2}\s*[:：]\s*\d{1,2}/.test(text) ||
    /(?:早上|上午|中午|下午|晚上|今晚|明天|今天)?\s*[零〇一二兩三四五六七八九十\d]{1,3}\s*(?:點|時)(?:半|[零〇一二兩三四五六七八九十\d]{1,2}\s*分)?/.test(
      text,
    )
  );
}

function updateEditablePickupTime_(order, editableOrder, currentMessage) {
  if (!order || !editableOrder || !editableOrder.order) {
    return;
  }

  if (messageUpdatesPickupTime_(currentMessage)) {
    order.pickup_time_source = "requested";
    return;
  }

  const existingOrder = editableOrder.order;
  const existingPickupTime = existingOrder.pickup_time || "";
  const existingSource =
    normalizePickupTimeSource_(existingOrder.pickup_time_source) ||
    (messageUpdatesPickupTime_(editableOrder.rawMessage)
      ? "requested"
      : "estimated");

  order.pickup_time = existingPickupTime;
  order.pickup_time_source = existingSource;

  if (existingSource !== "estimated" || !existingPickupTime) {
    return;
  }

  const addedPreparationMinutes =
    getOrderPreparationMinutes_(order) -
    getOrderPreparationMinutes_(existingOrder);

  if (addedPreparationMinutes <= 0) {
    return;
  }

  const existingPickupAt = parsePickupDateTime_(
    existingPickupTime,
    editableOrder.createdAt,
  );

  if (!existingPickupAt) {
    return;
  }

  const shiftedTimestamp =
    existingPickupAt.getTime() + addedPreparationMinutes * 60000;
  const roundedTimestamp = Math.ceil(shiftedTimestamp / 60000) * 60000;

  order.pickup_time = Utilities.formatDate(
    new Date(roundedTimestamp),
    "Asia/Taipei",
    "yyyy/MM/dd HH:mm",
  );
}

function findEditableOrder(lineUserId) {
  if (!lineUserId) {
    return null;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");

  const dataRange = sheet.getDataRange();

  const values = dataRange.getValues();

  const displayValues = dataRange.getDisplayValues();

  const now = new Date();

  let candidate = null;

  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][12]).trim() !== String(lineUserId).trim()) {
      continue;
    }

    if (values[i][10] !== "待製作") {
      continue;
    }

    const createdAt = parseOrderDate_(values[i][2]);

    if (!createdAt) {
      continue;
    }

    const diffMinutes = (now - createdAt) / 1000 / 60;

    // 建立後20分鐘內可修改
    if (diffMinutes >= 0 && diffMinutes <= 20) {
      candidate = {
        systemOrderId: values[i][0],

        displayOrderId: values[i][1],

        createdAt: values[i][2],
      };

      break;
    }
  }

  if (!candidate) {
    return null;
  }

  const context = buildExistingOrderContext_(
    displayValues,
    candidate.systemOrderId,
  );

  candidate.rawMessage = context.rawMessage;
  candidate.order = context.order;

  return candidate;
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
    const rawMessage = appendRawMessage_(editableOrder.rawMessage, newMessage);

    const saveResult = saveOrder(
      order,
      rawMessage,
      lineUserId,
      null,
      editableOrder.displayOrderId,
      editableOrder.createdAt,
      true,
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

function getArchivedTestOrderPrefix_(monthKey) {
  return "DASHBOARD-TEST-" + monthKey.replace("/", "") + "-";
}

function buildArchivedTestOrderRows_(monthKey, orderCount, priceMap, archivedAt) {
  const itemNames = Object.keys(priceMap).sort((a, b) =>
    a.localeCompare(b, "zh-Hant"),
  );

  if (itemNames.length === 0) {
    throw new Error("PriceList 沒有可用且已設定價格的品項");
  }

  const compactMonth = monthKey.replace("/", "");
  const prefix = getArchivedTestOrderPrefix_(monthKey);
  const rows = [];

  for (let index = 0; index < orderCount; index++) {
    const orderNumber = index + 1;
    const suffix = String(orderNumber).padStart(3, "0");
    const day = String((index % 28) + 1).padStart(2, "0");
    const hour = 11 + (index % 8);
    const minute = (index * 7) % 60;
    const pickupTotalMinutes = hour * 60 + minute + 15;
    const createdAt =
      monthKey +
      "/" +
      day +
      " " +
      String(hour).padStart(2, "0") +
      ":" +
      String(minute).padStart(2, "0") +
      ":00";
    const pickupTime =
      monthKey +
      "/" +
      day +
      " " +
      String(Math.floor(pickupTotalMinutes / 60)).padStart(2, "0") +
      ":" +
      String(pickupTotalMinutes % 60).padStart(2, "0");
    const itemCount = (index % 3) + 1;

    for (let itemIndex = 0; itemIndex < itemCount; itemIndex++) {
      const itemName = itemNames[(index + itemIndex * 3) % itemNames.length];
      const quantity = ((index + itemIndex) % 4) + 1;

      rows.push([
        prefix + suffix,
        "T" + suffix,
        createdAt,
        "六月測試顧客" + String(orderNumber).padStart(2, "0"),
        "",
        "一般",
        itemName,
        quantity,
        pickupTime,
        index % 4 === 0 ? "測試備註：不辣" : "",
        "已完成",
        "Dashboard 六月歷史訂單測試資料",
        "TEST-LINE-" + compactMonth + "-" + suffix,
        priceMap[itemName],
        archivedAt,
      ]);
    }
  }

  return rows;
}

function refreshArchiveMonthSummaries_(
  archiveSheet,
  summarySheet,
  itemSummarySheet,
  monthKey,
  priceMap,
  archivedAt,
) {
  const result = buildMonthlyArchiveData_(
    archiveSheet.getDataRange().getValues(),
    monthKey,
    priceMap,
  );

  upsertMonthlySummary_(summarySheet, result.summary, archivedAt);
  replaceMonthlyItemSummary_(
    itemSummarySheet,
    monthKey,
    result.itemSummaries,
  );

  return result;
}

function generateArchivedTestOrders_(monthKey, orderCount) {
  const targetMonth = validateArchiveMonth_(monthKey);
  const count = Math.floor(Number(orderCount));

  if (!Number.isFinite(count) || count < 1 || count > 200) {
    throw new Error("測試訂單數量必須介於 1 到 200");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
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
    const prefix = getArchivedTestOrderPrefix_(targetMonth);
    const existingValues = archiveSheet.getDataRange().getValues();
    const oldTestRows = [];

    for (let index = 1; index < existingValues.length; index++) {
      if (String(existingValues[index][0] || "").startsWith(prefix)) {
        oldTestRows.push(index + 1);
      }
    }

    deleteOrderRows_(archiveSheet, oldTestRows);

    const priceMap = getPriceMap();
    const archivedAt = new Date();
    const rows = buildArchivedTestOrderRows_(
      targetMonth,
      count,
      priceMap,
      archivedAt,
    );

    archiveSheet
      .getRange(
        archiveSheet.getLastRow() + 1,
        1,
        rows.length,
        ORDER_ARCHIVE_HEADERS_.length,
      )
      .setValues(rows);

    const result = refreshArchiveMonthSummaries_(
      archiveSheet,
      summarySheet,
      itemSummarySheet,
      targetMonth,
      priceMap,
      archivedAt,
    );

    return {
      generated: true,
      month: targetMonth,
      generatedOrders: count,
      generatedRows: rows.length,
      replacedRows: oldTestRows.length,
      monthSummary: result.summary,
    };
  } finally {
    lock.releaseLock();
  }
}

function clearArchivedTestOrders_(monthKey) {
  const targetMonth = normalizeMonthKey_(monthKey);

  if (!targetMonth) {
    throw new Error("測試月份格式必須為 yyyy/MM");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
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
    const prefix = getArchivedTestOrderPrefix_(targetMonth);
    const values = archiveSheet.getDataRange().getValues();
    const testRows = [];

    for (let index = 1; index < values.length; index++) {
      if (String(values[index][0] || "").startsWith(prefix)) {
        testRows.push(index + 1);
      }
    }

    deleteOrderRows_(archiveSheet, testRows);

    const archivedAt = new Date();
    const result = refreshArchiveMonthSummaries_(
      archiveSheet,
      summarySheet,
      itemSummarySheet,
      targetMonth,
      getPriceMap(),
      archivedAt,
    );

    return {
      cleared: true,
      month: targetMonth,
      removedRows: testRows.length,
      remainingMonthOrders: result.summary.orderCount,
    };
  } finally {
    lock.releaseLock();
  }
}

function generateJuneArchiveTestOrders() {
  return generateArchivedTestOrders_("2026/06", 25);
}

function clearJuneArchiveTestOrders() {
  return clearArchivedTestOrders_("2026/06");
}

function getJuneArchiveTestStatus() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = spreadsheet.getSheetByName("OrdersArchive");
  const prefix = getArchivedTestOrderPrefix_("2026/06");
  const orderIds = new Set();
  let testRows = 0;

  if (archiveSheet) {
    const values = archiveSheet.getDataRange().getValues();

    for (let index = 1; index < values.length; index++) {
      const systemOrderId = String(values[index][0] || "");

      if (systemOrderId.startsWith(prefix)) {
        orderIds.add(systemOrderId);
        testRows++;
      }
    }
  }

  const availableMonths = getAvailableDashboardMonths_();

  return {
    month: "2026/06",
    archiveSheetExists: Boolean(archiveSheet),
    testOrders: orderIds.size,
    testRows,
    availableMonths,
    visibleInDashboard: availableMonths.includes("2026/06"),
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

  ["Orders", "OrdersArchive"].forEach((sheetName) => {
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
      pickup: formatDashboardPickupTime_(
        order.pickup,
        order.createdAt,
        monthKey !== currentMonthKey,
      ),
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
    aiOrderStatus: getAiOrderStatus(),
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

function formatDashboardPickupTime_(pickupTime, createdAt, includeDate) {
  const parsed = parsePickupDateTime_(pickupTime, createdAt);

  if (parsed) {
    return Utilities.formatDate(
      parsed,
      "Asia/Taipei",
      includeDate ? "MM/dd HH:mm" : "HH:mm",
    );
  }

  const text = String(pickupTime || "").trim();
  const timeMatch = text.match(/(?:^|\s)(\d{1,2}):(\d{2})(?::\d{2})?/);

  if (timeMatch) {
    return String(Number(timeMatch[1])).padStart(2, "0") + ":" + timeMatch[2];
  }

  return text;
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

function acquireLineUserOrderLock_(lineUserId, waitMilliseconds = 30000) {
  const key = "line-order-lock:" + String(lineUserId || "anonymous");
  const token = Utilities.getUuid();
  const deadline = Date.now() + waitMilliseconds;

  while (Date.now() < deadline) {
    const lock = LockService.getScriptLock();

    lock.waitLock(5000);

    try {
      const properties = PropertiesService.getScriptProperties();
      const storedValue = properties.getProperty(key);

      let storedLock = null;

      try {
        storedLock = storedValue ? JSON.parse(storedValue) : null;
      } catch (error) {
        storedLock = null;
      }

      if (
        !storedLock ||
        !storedLock.token ||
        Number(storedLock.expiresAt) <= Date.now()
      ) {
        properties.setProperty(
          key,
          JSON.stringify({
            token,
            expiresAt: Date.now() + 120000,
          }),
        );

        return {
          key,
          token,
        };
      }
    } finally {
      lock.releaseLock();
    }

    Utilities.sleep(200);
  }

  throw new Error("同一位客戶的上一則訂單訊息仍在處理，請稍後重試");
}

function releaseLineUserOrderLock_(orderLock) {
  if (!orderLock || !orderLock.key || !orderLock.token) {
    return;
  }

  const lock = LockService.getScriptLock();
  let lockAcquired = false;

  try {
    lock.waitLock(5000);
    lockAcquired = true;

    const properties = PropertiesService.getScriptProperties();
    const storedValue = properties.getProperty(orderLock.key);

    let storedLock = null;

    try {
      storedLock = storedValue ? JSON.parse(storedValue) : null;
    } catch (error) {
      storedLock = null;
    }

    if (storedLock && storedLock.token === orderLock.token) {
      properties.deleteProperty(orderLock.key);
    }
  } catch (error) {
    console.error("釋放客戶訂單鎖失敗:", error);
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
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

  if (!isAiOrderEnabled_()) {
    return createAssistantDisabledOutput_();
  }

  const lineUserOrderLock = acquireLineUserOrderLock_(lineUserId);

  try {
    if (!isAiOrderEnabled_()) {
      return createAssistantDisabledOutput_();
    }

    saveCustomer(lineUserId, lineName, pictureUrl);

  const editable = findEditableOrder(lineUserId);

  const order = parseOrder(message, editable);

  if (!order.is_order) {
    return ContentService.createTextOutput(
      JSON.stringify({
        is_order: false,
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  }

  updateEditablePickupTime_(order, editable, message);

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

  const waitMinutes = getFinalEstimatedWaitMinutes_(
    getEstimatedWaitMinutes(editable ? editable.systemOrderId : ""),
    order,
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
      const saveResult = saveOrder(
        order,
        message,
        lineUserId,
        null,
        null,
        null,
        true,
      );

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
  } catch (error) {
    if (error && error.code === "AI_ORDER_DISABLED") {
      return createAssistantDisabledOutput_();
    }

    throw error;
  } finally {
    releaseLineUserOrderLock_(lineUserOrderLock);
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
