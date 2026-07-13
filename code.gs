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
   `;


const CLOUD_RUN_NOTIFY_URL =
  "https://line-order-webhook-171295331325.asia-east1.run.app/notifyReady";

function parseOrder(message) {

  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty("OPENAI_API_KEY");

  const payload = {
  model: "gpt-4.1-mini",
  temperature: 0,
  response_format: {
    type: "json_object"
  },
  messages: [
    {
      role: "system",
      content: SYSTEM_PROMPT
    },
    {
      role: "user",
      content: message
    }
  ]
};

  const response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/chat/completions",
    {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: `Bearer ${apiKey}`
      },
      payload: JSON.stringify(payload)
    }
  );

  const result =
    JSON.parse(response.getContentText());

  const content =
    result.choices[0].message.content;

  return JSON.parse(content);
}



function saveOrder(
  order,
  rawMessage,
  lineUserId,
  systemOrderId = null,
  displayOrderId = null,
  createdAt = null
){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Orders");

  const priceMap =
    getPriceMap();

  const priceColumn = 14;

  systemOrderId =
    systemOrderId ||
    generateSystemOrderId();

  displayOrderId =
    displayOrderId ||
    generateDisplayOrderId();

  createdAt =
    createdAt ||
    Utilities.formatDate(
      new Date(),
      "Asia/Taipei",
      "yyyy/MM/dd HH:mm:ss"
    );

  if(
    !order ||
    !Array.isArray(order.groups)
  ){
    throw new Error(
      "訂單格式錯誤：缺少 groups"
    );
  }

  const rows = [];

  order.groups.forEach(group => {

    if(!Array.isArray(group.items)){
      return;
    }

    group.items.forEach(item => {

      const quantity =
        Number(item.quantity) || 0;

      const unitPrice =
        quantity > 10
          ? quantity
          : priceMap[item.name];

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

        unitPrice === undefined
          ? "UNPRICED"
          : unitPrice

      ]);

    });

  });

  if(rows.length === 0){
    throw new Error(
      "訂單格式錯誤：沒有可寫入的商品"
    );
  }

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try{

    if(
      !sheet
        .getRange(1, priceColumn)
        .getValue()
    ){
      sheet
        .getRange(1, priceColumn)
        .setValue("unit_price");
    }

    sheet
      .getRange(
        sheet.getLastRow() + 1,
        1,
        rows.length,
        rows[0].length
      )
      .setValues(rows);

  }finally{
    lock.releaseLock();
  }

  return {

    systemOrderId,

    displayOrderId,

    createdAt

  };

}

function manualCreateOrder(message) {

  const order =
    parseOrder(message);

  if(!order.is_order){
    throw new Error(
      "訊息無法解析為訂單"
    );
  }

  normalizeItems(order);

  return saveOrder(
    order,
    message
  );

}

function getOrders() {
  try{
    const sheet =
      SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName("Orders");

    const values =
      sheet.getDataRange()
        .getDisplayValues();

    const orders = {};

    const customerMap =
      getCustomerMap();

    const priceMap =
      getPriceMap();

    for(let i=1;i<values.length;i++){

      const row = values[i];

      if(row[10] !== "待製作")
        continue;

      const systemOrderId =
        row[0];

      if(!orders[systemOrderId]){
        const lineUserId =
          row[12];

        const customer =
          customerMap[lineUserId]
          || {};
        orders[systemOrderId] = {

          systemOrderId: systemOrderId,
          orderId: row[1],
          createdAt: row[2].substring(11,16),
          customer: row[3],
          pictureUrl:
            customer.pictureUrl || "",
          totalOrders:
            customer.totalOrders || 0,
          pickup: row[8],
          status: row[10],

          totalPrice: 0,
          priceMissing: false,

          missingItems: [],

          groups: {}

        };

      }

      const groupName =
        row[5] || "一般";

      if(!orders[systemOrderId]
        .groups[groupName]){

        orders[systemOrderId]
          .groups[groupName] = [];

      }

      const itemName =
        row[6];

      const quantity =
        Number(row[7]);

      let displayItem = "";

      if(quantity > 10){

        displayItem =
          `${itemName}${quantity} x1`;

      }else{

        displayItem =
          `${itemName} x${quantity}`;

      }

      orders[systemOrderId]
        .groups[groupName]
        .push({

          item: displayItem,

          note: row[9]

        });

      const priceInfo =
        resolveOrderPrice_(
          row,
          priceMap
        );

      if(priceInfo.missing){

        orders[systemOrderId]
          .priceMissing = true;

        if(
          !orders[systemOrderId]
            .missingItems
            .includes(itemName)
        ){

          orders[systemOrderId]
            .missingItems
            .push(itemName);

        }
        
        console.log(
          "價格未設定:",
          itemName
        );

      }

      if(!priceInfo.missing){
        orders[systemOrderId]
          .totalPrice +=
            priceInfo.lineTotal;
      }

    }

    return Object
      .values(orders)
      .sort((a,b)=>
        a.orderId.localeCompare(b.orderId)
      );
  }catch(error){

    console.error(
      "getOrders錯誤:",
      error.stack
    );

    throw error;

  }
}

function updateOrderStatus(systemOrderId) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Orders");

  const values =
    sheet.getDataRange()
      .getValues();

  let lineUserId = "";
  let displayOrderId = "";
  const orderRows = [];

  for(let i=1;i<values.length;i++){

    if(values[i][0] === systemOrderId){

      orderRows.push(i + 1);

      if(!displayOrderId){

        lineUserId =
          values[i][12];

        displayOrderId =
          values[i][1];

      }

    }

  }

  if(orderRows.length === 0){
    throw new Error(
      "找不到指定的系統訂單 ID"
    );
  }

  sheet
    .getRangeList(
      orderRows.map(rowNumber =>
        `K${rowNumber}`
      )
    )
    .setValue("已完成");

  if(!lineUserId){
    return {
      systemOrderId,
      displayOrderId,
      notified: false
    };
  }

  console.log(
    "通知使用者:",
    lineUserId
  );

  UrlFetchApp.fetch(
    CLOUD_RUN_NOTIFY_URL,
    {
      method: "post",
      contentType:
        "application/json",

      payload:
        JSON.stringify({
          userId: lineUserId,
          orderId: displayOrderId,
          notifySecret:
            getRequiredScriptProperty_(
              "CLOUD_RUN_NOTIFY_SECRET"
            )
        })
    }
  );

  return {
    systemOrderId,
    displayOrderId,
    notified: true
  };

}

function getCustomerMap(){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        "Customers"
      );

  const values =
    sheet.getDataRange()
      .getValues();

  const map = {};

  for(
    let i = 1;
    i < values.length;
    i++
  ){

    map[
      values[i][0]
    ] = {

      displayName:
        values[i][1] || "",

      pictureUrl:
        values[i][2] || "",

      totalOrders:
        Number( values[i][5] ) || 0

    };

  }

  return map;

}

function getDashboardSummary() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Orders");

  const values =
    sheet.getDataRange()
      .getDisplayValues();

  const orderMap = {};
  const today =
    Utilities.formatDate(
      new Date(),
      "Asia/Taipei",
      "yyyy/MM/dd"
    );

  for(let i=1;i<values.length;i++){

    const row = values[i];

    const orderDate = row[2].substring(0,10);

    if(orderDate !== today)
      continue;

    orderMap[row[1]] = row[10];
  }

  let pending = 0;
  let completed = 0;

  Object.values(orderMap)
    .forEach(status => {

      if(status === "待製作"){
        pending++;
      }

      if(status === "已完成"){
        completed++;
      }

    });

  return {
    pending,
    completed
  };

}

function getCompletedOrders() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Orders");

  const values =
    sheet.getDataRange()
      .getDisplayValues();

  const orders = {};

  const customerMap =
    getCustomerMap();
  
  const priceMap =
    getPriceMap();

  const today =
    Utilities.formatDate(
      new Date(),
      "Asia/Taipei",
      "yyyy/MM/dd"
    );

  for(let i=1;i<values.length;i++){

    const row = values[i];

    if(row[10] !== "已完成")
      continue;

    const orderDate =
      row[2].substring(0,10);

    if(orderDate !== today)
      continue;

    const systemOrderId =
      row[0];

    if(!orders[systemOrderId]){
      const lineUserId =
          row[12];

        const customer =
          customerMap[lineUserId]
          || {};
        orders[systemOrderId] = {

          systemOrderId: systemOrderId,
          orderId: row[1],
          createdAt: row[2].substring(11,16),
          customer: row[3],
          pictureUrl:
            customer.pictureUrl || "",
          totalOrders:
            customer.totalOrders || 0,
          pickup: row[8],
          status: row[10],

          totalPrice: 0,
          priceMissing: false,

          missingItems: [],

          groups: {}

        };

    }

    const groupName =
      row[5] || "一般";

    if(!orders[systemOrderId]
      .groups[groupName]){

      orders[systemOrderId]
        .groups[groupName] = [];

    }

    orders[systemOrderId]
      .groups[groupName]
      .push({

        item:
          `${row[6]} x${row[7]}`,

        note:
          row[9]

      });

    const itemName =
      row[6];

    const quantity =
      Number(row[7]);

    const priceInfo =
      resolveOrderPrice_(
        row,
        priceMap
      );

    if(priceInfo.missing){
      orders[systemOrderId]
        .priceMissing = true;
      console.log(
        "價格未設定:",
        itemName
      );
    }

    if(!priceInfo.missing){
      orders[systemOrderId]
        .totalPrice +=
          priceInfo.lineTotal;
    }

  }

  return Object
    .values(orders)
    .sort((a,b)=>
      b.orderId.localeCompare(a.orderId)
    )
    .slice(0, 20);

}

function generateSystemOrderId() {

  return Utilities.getUuid();

}

function generateDisplayOrderId() {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try{

  const props =
    PropertiesService
      .getScriptProperties();

  const today =
    Utilities.formatDate(
      new Date(),
      "Asia/Taipei",
      "yyyy-MM-dd"
    );

  const savedDate =
    props.getProperty("ORDER_DATE");

  if(savedDate !== today){

    props.setProperty(
      "ORDER_DATE",
      today
    );

    props.setProperty(
      "ORDER_COUNTER",
      "0"
    );

  }

  let counter =
    parseInt(
      props.getProperty("ORDER_COUNTER") || "0"
    );

  counter++;

  props.setProperty(
    "ORDER_COUNTER",
    String(counter)
  );

  Logger.log(
    "Counter=" + counter
  );

  return (
    "A" +
    String(counter)
      .padStart(3,"0")
  );

  }finally{
    lock.releaseLock();
  }

}

function calculateOrderPrice(
  order
){

  const priceMap =
    getPriceMap();

  let totalPrice = 0;

  const missingItems = [];

  order.groups.forEach(group => {

    group.items.forEach(item => {

      const price =
        priceMap[item.name];

      if(item.quantity > 10){
        totalPrice +=
          Number(item.quantity) || 0;
        return;
      }

      if(price === undefined){

        if(
          !missingItems.includes(
            item.name
          )
        ){
          missingItems.push(
            item.name
          );
        }

        return;
      }

      totalPrice +=
        price * item.quantity;

    });

  });

  return {
    totalPrice,
    priceMissing:
      missingItems.length > 0,
    missingItems
  };

}

function getPriceMap() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        "PriceList"
      );

  const values =
    sheet
      .getDataRange()
      .getValues();

  const map = {};

  for(
    let i = 1;
    i < values.length;
    i++
  ){

    const itemName =
      String(values[i][0] || "")
        .trim();

    const rawPrice =
      values[i][1];

    if(
      !itemName ||
      rawPrice === "" ||
      rawPrice === null ||
      rawPrice === undefined
    ){
      continue;
    }

    const price =
      Number(rawPrice);

    if(!Number.isFinite(price)){
      continue;
    }

    map[itemName] = price;

  }

  return map;

}


function resolveOrderPrice_(
  row,
  priceMap
){

  const itemName =
    String(row[6] || "");

  const quantity =
    Number(row[7]) || 0;

  const storedPrice =
    row[13];

  if(storedPrice === "UNPRICED"){
    return {
      missing: true,
      unitPrice: null,
      lineTotal: 0
    };
  }

  if(
    storedPrice !== "" &&
    storedPrice !== null &&
    storedPrice !== undefined &&
    Number.isFinite(Number(storedPrice))
  ){

    const unitPrice =
      Number(storedPrice);

    return {
      missing: false,
      unitPrice,
      lineTotal:
        quantity > 10
          ? unitPrice
          : unitPrice * quantity
    };

  }

  // 舊訂單沒有成交單價欄位，維持原本以 PriceList 回推的行為。
  if(quantity > 10){
    return {
      missing: false,
      unitPrice: quantity,
      lineTotal: quantity
    };
  }

  const currentPrice =
    priceMap[itemName];

  if(currentPrice === undefined){
    return {
      missing: true,
      unitPrice: null,
      lineTotal: 0
    };
  }

  return {
    missing: false,
    unitPrice: currentPrice,
    lineTotal: currentPrice * quantity
  };

}

function getAliasMap(){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        "ItemAlias"
      );

  const values =
    sheet.getDataRange()
      .getValues();

  const map = {};

  for(let i=1;i<values.length;i++){

    map[
      values[i][0]
    ] =
      values[i][1];

  }

  return map;

}

function getEstimatedWaitMinutes() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Orders");

  const values =
    sheet.getDataRange()
      .getDisplayValues();

  const orderIds = new Set();

  for(let i = 1; i < values.length; i++){

    if(values[i][10] === "待製作"){

      orderIds.add(
        values[i][1]
      );

    }

  }
  //一單製作時間 預計分鐘數
  return orderIds.size * 5;

}

function getEstimatedPickupTime(
  waitMinutes
){

  const now =
    new Date();

  const finishTime =
    new Date(
      now.getTime()
      + waitMinutes * 60000
    );

  return Utilities.formatDate(
    finishTime,
    "Asia/Taipei",
    "HH:mm"
  );

}

function normalizeItems(order){
  const aliasMap =
    getAliasMap();

  order.groups.forEach(group => {

    group.items.forEach(item => {

      if(
        aliasMap[item.name]
      ){

        item.name =
          aliasMap[item.name];

      }
    });
  });
}


function incrementCustomerOrderCount(
  lineUserId
){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        "Customers"
      );

  const values =
    sheet.getDataRange()
      .getValues();

  for(
    let i = 1;
    i < values.length;
    i++
  ){

    if(
      values[i][0] === lineUserId
    ){

      const current =
        Number(
          values[i][5]
        ) || 0;

      sheet.getRange(
        i + 1,
        6
      ).setValue(
        current + 1
      );

      return;

    }

  }

}

function saveCustomer(
  lineUserId,
  displayName,
  pictureUrl
){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        "Customers"
      );

  const values =
    sheet.getDataRange()
      .getValues();

  for(
    let i = 1;
    i < values.length;
    i++
  ){

    if(
      values[i][0] === lineUserId
    ){

      sheet.getRange(
        i + 1,
        2
      ).setValue(
        displayName
      );

      sheet.getRange(
        i + 1,
        3
      ).setValue(
        pictureUrl
      );

      sheet.getRange(
        i + 1,
        5
      ).setValue(
        new Date()
      );

      return;

    }

  }

  sheet.appendRow([
    lineUserId,
    displayName,
    pictureUrl,
    "",
    new Date(),
    0
  ]);

}

function deleteOrder(
  systemOrderId
){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Orders");

  const values =
    sheet.getDataRange()
      .getValues();

  for(
    let i = values.length - 1;
    i >= 1;
    i--
  ){

    if(
      values[i][0] === systemOrderId
    ){

      sheet.deleteRow(
        i + 1
      );

    }

  }

}

function findEditableOrder(
  lineUserId
){

  if(!lineUserId){
    return null;
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Orders");

  const values =
    sheet.getDataRange()
      .getValues();

  const now =
    new Date();

  for(
    let i = values.length - 1;
    i >= 1;
    i--
  ){

    if(
      String(values[i][12]).trim() !==
      String(lineUserId).trim()
    ){
      continue;
    }

    if(values[i][10] !== "待製作"){
      continue;
    }

    const createdAt =
      new Date(values[i][2]);

    const diffMinutes =
      (now - createdAt)
      / 1000
      / 60;

    // 建立後20分鐘內可修改
    if(
      diffMinutes >= 0 &&
      diffMinutes <= 20
    ){

      return {

        systemOrderId:
          values[i][0],

        displayOrderId:
          values[i][1],

        createdAt:
          values[i][2],

        rawMessage:
          values[i][11]

      };

    }

  }

  return null;

}


function updateOrder(
  editableOrder,
  order,
  newMessage,
  lineUserId
){

  let orderPersisted = false;

  try{

  const saveResult =
    saveOrder(
      order,
      newMessage,
      lineUserId,
      null,
      editableOrder.displayOrderId,
      editableOrder.createdAt
    );

  orderPersisted = true;

  deleteOrder(
    editableOrder.systemOrderId
  );

  const priceInfo =
    calculateOrderPrice(
      order
    );

  const waitMinutes =
    Math.max(
      getEstimatedWaitMinutes(),
      10
    );

  const estimatedPickupTime =
    getEstimatedPickupTime(
      waitMinutes
    );

  return {

    is_order:true,

    updated:true,

    orderId:
      saveResult.displayOrderId,

    order,

    totalPrice:
      priceInfo.totalPrice,

    priceMissing:
      priceInfo.priceMissing,

    missingItems:
      priceInfo.missingItems,

    estimatedPickupTime

  };

  }catch(error){

    error.orderPersisted =
      orderPersisted;

    throw error;

  }

}


function getAdminDashboardData(){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Orders");

  const values =
    sheet.getDataRange()
      .getValues();

  const priceMap =
    getPriceMap();

  const timeZone =
    "Asia/Taipei";

  const now =
    new Date();

  const todayKey =
    Utilities.formatDate(
      now,
      timeZone,
      "yyyy/MM/dd"
    );

  const monthKey =
    Utilities.formatDate(
      now,
      timeZone,
      "yyyy/MM"
    );

  const orders = {};
  const itemStats = {};

  for(let i = 1; i < values.length; i++){

    const row = values[i];
    const systemOrderId =
      String(row[0] || "");

    if(!systemOrderId){
      continue;
    }

    const createdAt =
      parseOrderDate_(row[2]);

    if(!createdAt){
      continue;
    }

    if(!orders[systemOrderId]){

      orders[systemOrderId] = {
        orderId: String(row[1] || ""),
        createdAt: createdAt,
        customer: String(row[3] || "未填寫"),
        pickup: String(row[8] || ""),
        status: String(row[10] || ""),
        total: 0,
        itemCount: 0,
        priceMissing: false
      };

    }

    const itemName =
      String(row[6] || "");

    const quantity =
      Number(row[7]) || 0;

    const priceInfo =
      resolveOrderPrice_(
        row,
        priceMap
      );

    const lineTotal =
      priceInfo.lineTotal;

    const itemQuantity =
      quantity > 10
        ? 1
        : quantity;

    if(priceInfo.missing){
      orders[systemOrderId]
        .priceMissing = true;
    }

    orders[systemOrderId]
      .total += lineTotal;

    orders[systemOrderId]
      .itemCount += itemQuantity;

    const orderMonthKey =
      Utilities.formatDate(
        createdAt,
        timeZone,
        "yyyy/MM"
      );

    if(orderMonthKey === monthKey && itemName){

      if(!itemStats[itemName]){
        itemStats[itemName] = {
          name: itemName,
          quantity: 0,
          revenue: 0
        };
      }

      itemStats[itemName]
        .quantity += itemQuantity;

      itemStats[itemName]
        .revenue += lineTotal;

    }

  }

  const allOrders =
    Object.values(orders);

  const todayOrders =
    allOrders.filter(order =>
      Utilities.formatDate(
        order.createdAt,
        timeZone,
        "yyyy/MM/dd"
      ) === todayKey
    );

  const monthOrders =
    allOrders.filter(order =>
      Utilities.formatDate(
        order.createdAt,
        timeZone,
        "yyyy/MM"
      ) === monthKey
    );

  const daysInMonth =
    Number(
      Utilities.formatDate(
        now,
        timeZone,
        "d"
      )
    );

  const dailyMap = {};

  for(let day = 1; day <= daysInMonth; day++){
    const key =
      monthKey + "/" +
      String(day).padStart(2,"0");

    dailyMap[key] = {
      label: day + "日",
      orders: 0,
      revenue: 0
    };
  }

  monthOrders.forEach(order => {

    const key =
      Utilities.formatDate(
        order.createdAt,
        timeZone,
        "yyyy/MM/dd"
      );

    if(dailyMap[key]){
      dailyMap[key].orders++;
      dailyMap[key].revenue += order.total;
    }

  });

  const summarize =
    orderList => {

      const revenue =
        orderList.reduce(
          (sum, order) =>
            sum + order.total,
          0
        );

      const completed =
        orderList.filter(order =>
          order.status === "已完成"
        ).length;

      const pending =
        orderList.filter(order =>
          order.status === "待製作"
        ).length;

      return {
        orders: orderList.length,
        revenue: Math.round(revenue),
        average: orderList.length
          ? Math.round(revenue / orderList.length)
          : 0,
        completed,
        pending,
        completionRate: orderList.length
          ? Math.round(completed / orderList.length * 100)
          : 0
      };

    };

  const recentOrders =
    todayOrders
      .sort((a,b) =>
        b.createdAt.getTime() -
        a.createdAt.getTime()
      )
      .slice(0,10)
      .map(order => ({
        orderId: order.orderId,
        time: Utilities.formatDate(
          order.createdAt,
          timeZone,
          "HH:mm"
        ),
        customer: order.customer,
        pickup: order.pickup,
        status: order.status,
        total: Math.round(order.total),
        priceMissing: order.priceMissing
      }));

  return {
    generatedAt:
      Utilities.formatDate(
        now,
        timeZone,
        "yyyy/MM/dd HH:mm:ss"
      ),
    todayLabel:
      Utilities.formatDate(
        now,
        timeZone,
        "MM/dd"
      ),
    monthLabel:
      Utilities.formatDate(
        now,
        timeZone,
        "yyyy/MM"
      ),
    today: summarize(todayOrders),
    month: summarize(monthOrders),
    daily: Object.values(dailyMap),
    topItems: Object.values(itemStats)
      .sort((a,b) => b.quantity - a.quantity)
      .slice(0,8)
      .map(item => ({
        name: item.name,
        quantity: item.quantity,
        revenue: Math.round(item.revenue)
      })),
    recentOrders
  };

}


function parseOrderDate_(value){

  if(
    Object.prototype.toString.call(value) ===
    "[object Date]" &&
    !isNaN(value.getTime())
  ){
    return value;
  }

  const match =
    String(value || "")
      .match(
        /^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/
      );

  if(!match){
    return null;
  }

  return new Date(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
    Number(match[4] || 0),
    Number(match[5] || 0),
    Number(match[6] || 0)
  );

}


function getRequiredScriptProperty_(key){

  const value =
    PropertiesService
      .getScriptProperties()
      .getProperty(key);

  if(!value){
    throw new Error(
      `缺少 Script Property: ${key}`
    );
  }

  return value;

}


function claimWebhookEvent_(webhookEventId){

  if(!webhookEventId){
    return true;
  }

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try{

    const spreadsheet =
      SpreadsheetApp
        .getActiveSpreadsheet();

    let sheet =
      spreadsheet
        .getSheetByName(
          "WebhookEvents"
        );

    if(!sheet){
      sheet =
        spreadsheet
          .insertSheet(
            "WebhookEvents"
          );

      sheet
        .getRange(1,1,1,2)
        .setValues([[
          "webhook_event_id",
          "processed_at"
        ]]);
    }

    const lastRow =
      Math.max(
        sheet.getLastRow(),
        1
      );

    const existing =
      sheet
        .getRange(1,1,lastRow,1)
        .createTextFinder(
          webhookEventId
        )
        .matchEntireCell(true)
        .findNext();

    if(existing){
      return false;
    }

    sheet
      .getRange(
        lastRow + 1,
        1,
        1,
        2
      )
      .setValues([[
        webhookEventId,
        new Date()
      ]]);

    return true;

  }finally{
    lock.releaseLock();
  }

}


function releaseWebhookEvent_(webhookEventId){

  if(!webhookEventId){
    return;
  }

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try{

    const sheet =
      SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName(
          "WebhookEvents"
        );

    if(!sheet){
      return;
    }

    const existing =
      sheet
        .getRange(
          1,
          1,
          Math.max(
            sheet.getLastRow(),
            1
          ),
          1
        )
        .createTextFinder(
          webhookEventId
        )
        .matchEntireCell(true)
        .findNext();

    if(
      existing &&
      existing.getRow() > 1
    ){
      sheet.deleteRow(
        existing.getRow()
      );
    }

  }finally{
    lock.releaseLock();
  }

}


function doGet(e) {

  const page =
    e &&
    e.parameter &&
    e.parameter.page === "menu"
      ? "menu"
      : "dashboard";

  return HtmlService
    .createHtmlOutputFromFile(page);

}


function getWebAppUrl(){

  return ScriptApp
    .getService()
    .getUrl();

}

function doPost(e) {

  const data =
    JSON.parse(
      e.postData.contents
    );

  const expectedSecret =
    getRequiredScriptProperty_(
      "WEBHOOK_SHARED_SECRET"
    );

  if(
    !data.webhookSecret ||
    data.webhookSecret !== expectedSecret
  ){
    return ContentService
      .createTextOutput(
        JSON.stringify({
          error:"UNAUTHORIZED"
        })
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );
  }

  const message =
    data.message || "";

  const lineUserId =
    data.lineUserId || "";

  const lineName =
    data.customerName || "";

  const pictureUrl =
    data.pictureUrl || "";

  const webhookEventId =
    String(
      data.webhookEventId || ""
    );

  saveCustomer(
    lineUserId,
    lineName,
    pictureUrl
  );

  const order =
    parseOrder(
      message
    );

  if(!order.is_order){

    return ContentService
      .createTextOutput(
        JSON.stringify({
          is_order:false
        })
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );

  }

  order.customer_name =
    order.customer_name ||
    lineName ||
    "";

  normalizeItems(
    order
  );

  const soldOutItems =
    checkSoldOutItems(
      order
    );

  if(soldOutItems.length > 0){

    return ContentService
      .createTextOutput(
        JSON.stringify({
          is_order:true,
          soldOut:true,
          soldOutItems:soldOutItems
        })
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );

  }

  const priceInfo =
    calculateOrderPrice(
      order
    );

  const editable =
    findEditableOrder(
      lineUserId
    );

  const eventClaimed =
    claimWebhookEvent_(
      webhookEventId
    );

  if(!eventClaimed){

    return ContentService
      .createTextOutput(
        JSON.stringify({
          duplicate:true
        })
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );

  }

  let orderPersisted = false;

  try{

  let result;

  if(editable){

    result =
      updateOrder(
        editable,
        order,
        message,
        lineUserId
      );

    orderPersisted = true;

  }else{

    const saveResult =
      saveOrder(
        order,
        message,
        lineUserId
      );

    orderPersisted = true;

    incrementCustomerOrderCount(
      lineUserId
    );

    const waitMinutes =
      Math.max(
        getEstimatedWaitMinutes(),
        10
      );

    const estimatedPickupTime =
      getEstimatedPickupTime(
        waitMinutes
      );

    result = {

      is_order:true,

      updated:false,

      orderId:
        saveResult.displayOrderId,

      order,

      totalPrice:
        priceInfo.totalPrice,

      priceMissing:
        priceInfo.priceMissing,

      missingItems:
        priceInfo.missingItems,

      estimatedPickupTime

    };

  }

  return ContentService
    .createTextOutput(
      JSON.stringify(
        result
      )
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

  }catch(error){

    if(
      !orderPersisted &&
      !error.orderPersisted
    ){
      releaseWebhookEvent_(
        webhookEventId
      );
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
function getMenuItems(){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("PriceList");

  const values =
    sheet.getDataRange()
      .getValues();

  const items = [];

  for(let i = 1; i < values.length; i++){

    items.push({
      rowIndex: i + 1,
      name: values[i][0],
      price: values[i][1],
      category: values[i][2] || "未分類",
      imageUrl: values[i][3] || "",
      isSoldOut: values[i][4] === true
    });

  }

  return items;

}

function updateSoldOut(
  rowIndex,
  isSoldOut
){

  SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("PriceList")
    .getRange(rowIndex,5)
    .setValue(isSoldOut);

  return true;

}

function getSoldOutItemNames(){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("PriceList");

  const values =
    sheet.getDataRange()
      .getValues();

  const soldOutItems = [];

  for(let i = 1; i < values.length; i++){

    const itemName =
      values[i][0];

    const isSoldOut =
      values[i][4] === true;

    if(isSoldOut){
      soldOutItems.push(itemName);
    }

  }

  return soldOutItems;

}

function checkSoldOutItems(order){

  const soldOutItems =
    getSoldOutItemNames();

  const found = [];

  order.groups.forEach(group => {

    group.items.forEach(item => {

      if(
        soldOutItems.includes(item.name) &&
        !found.includes(item.name)
      ){
        found.push(item.name);
      }

    });

  });

  return found;

}

function resetSoldOutStatus(){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("PriceList");

  const lastRow =
    sheet.getLastRow();

  if(lastRow <= 1){
    return;
  }

  sheet
    .getRange(
      2,
      5,
      lastRow - 1,
      1
    )
    .setValue(false);

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
    PropertiesService
      .getScriptProperties()
      .getProperty("OPENAI_API_KEY");

  const payload = {
    model: "gpt-4.1-mini",
    messages: [
      {
        role: "user",
        content: "請回傳你好"
      }
    ]
  };

  const response =
    UrlFetchApp.fetch(
      "https://api.openai.com/v1/chat/completions",
      {
        method: "post",
        contentType: "application/json",
        headers: {
          Authorization: `Bearer ${apiKey}`
        },
        payload: JSON.stringify(payload)
      }
    );

  Logger.log(
    response.getContentText()
  );

}

function testGetOrders(){

  const result =
    getOrders();

  console.log(
    result
  );

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
  

  const order =
    parseOrder(message);

  Logger.log(
    JSON.stringify(order, null, 2)
  );

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

  const order =
    parseOrder(message);

  saveOrder(
    order,
    message
  );

  Logger.log(order);

}
