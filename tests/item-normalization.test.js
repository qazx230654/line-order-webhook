const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

class MemorySheet {
  constructor(values = []) {
    this.values = values.map((row) => row.slice());
  }

  getDataRange() {
    return {
      getValues: () => this.values.map((row) => row.slice()),
    };
  }

  getLastRow() {
    return this.values.length;
  }

  getRange(row, column, rowCount, columnCount) {
    return {
      setValues: (rows) => {
        for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
          const targetRow = row - 1 + rowOffset;
          if (!this.values[targetRow]) this.values[targetRow] = [];

          for (let columnOffset = 0; columnOffset < columnCount; columnOffset += 1) {
            this.values[targetRow][column - 1 + columnOffset] =
              rows[rowOffset][columnOffset];
          }
        }
      },
    };
  }

  appendRow(row) {
    this.values.push(row.slice());
  }
}

const sheets = {
  PriceList: new MemorySheet([
    ['item_name', 'price', 'category', 'image_url', 'is_sold_out'],
    ['杏鮑菇', 30, '蔬菜', '', false],
    ['大豆干', 20, '豆類', '', false],
    ['甜不辣', 25, '火鍋料', '', false],
    ['貢丸', 15, '火鍋料', '', false],
  ]),
  ItemAlias: new MemorySheet([
    ['別名', '正式名稱'],
    ['杏包菇', '杏鮑菇'],
  ]),
  UnmatchedItems: new MemorySheet([
    ['raw_name', 'suggested_name', 'raw_message', 'occurrence_count', 'last_seen', 'status'],
    ['黑輪', '甜不辣', '舊訊息', 2, '2026/07/12 12:00:00', 'approved'],
  ]),
};

const spreadsheet = {
  getSheetByName: (name) => sheets[name] || null,
  insertSheet: (name) => {
    sheets[name] = new MemorySheet();
    return sheets[name];
  },
};

let lastFetchOptions;

const context = vm.createContext({
  console,
  SpreadsheetApp: {
    getActiveSpreadsheet: () => spreadsheet,
  },
  LockService: {
    getScriptLock: () => ({
      waitLock() {},
      releaseLock() {},
    }),
  },
  Utilities: {
    formatDate: () => '2026/07/13 18:00:00',
  },
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: () => 'test-key',
    }),
  },
  UrlFetchApp: {
    fetch: (url, options) => {
      lastFetchOptions = { url, options };
      return {
        getContentText: () => JSON.stringify({
          choices: [{
            message: {
              content: JSON.stringify({ is_order: false }),
            },
          }],
        }),
      };
    },
  },
});

const code = fs.readFileSync(path.join(__dirname, '..', 'code.gs'), 'utf8');
vm.runInContext(code, context);

context.parseOrder('測試訊息');
const requestPayload = JSON.parse(lastFetchOptions.options.payload);
const prompt = requestPayload.messages[0].content;
assert.match(prompt, /正式菜單名稱/);
assert.match(prompt, /杏鮑菇/);
assert.match(prompt, /大豆干/);
assert.match(prompt, /無法確定時保留客戶原文/);

const order = {
  groups: [
    {
      items: [
        { name: '杏鮑菇', raw_name: '杏鮑菇', quantity: 1 },
        { name: '杏 鮑 菇', raw_name: '杏 鮑 菇', quantity: 1 },
        { name: '大豆乾', raw_name: '大豆乾', quantity: 1 },
        { name: '杏包菇', raw_name: '杏包菇', quantity: 1 },
        { name: '黑輪', raw_name: '黑輪', quantity: 1 },
        { name: '神秘丸', raw_name: '神秘丸', suggested_name: '貢丸', quantity: 1 },
      ],
    },
  ],
};

const result = context.normalizeItems(order);
const items = order.groups[0].items;

assert.equal(items[0].name, '杏鮑菇');
assert.equal(items[1].name, '杏鮑菇');
assert.equal(items[2].name, '大豆干');
assert.equal(items[3].name, '杏鮑菇');
assert.equal(items[4].name, '甜不辣');
assert.equal(items[5].name, '神秘丸');
assert.equal(items[5].suggested_name, '貢丸');
assert.equal(items[5].match_status, 'unmatched');
assert.equal(result.unmatchedItems.length, 1);

context.logUnmatchedItems_(result.unmatchedItems, '神秘丸一份');
context.logUnmatchedItems_([
  { rawName: '神秘 丸', suggestedName: '貢丸' },
], '再一份神秘丸');

const logged = sheets.UnmatchedItems.values.find((row) => row[0] === '神秘丸');
assert.ok(logged);
assert.equal(logged[1], '貢丸');
assert.equal(logged[2], '再一份神秘丸');
assert.equal(logged[3], 2);
assert.equal(logged[5], 'pending');

assert.deepEqual(
  Array.from(context.getCanonicalItemNames_()),
  ['杏鮑菇', '大豆干', '甜不辣', '貢丸'],
);

console.log('Item normalization tests passed');
