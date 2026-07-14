const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const orderHeader = [
  'system_order_id', 'display_order_id', 'created_at', 'customer_name',
  'phone', 'group_name', 'item_name', 'quantity', 'pickup_time', 'note',
  'status', 'raw_message', 'line_user_id', 'unit_price', 'pickup_time_source',
];

const orderRows = [
  orderHeader,
  ['uuid-old', 'A001', '2026/07/14 15:00:00', '王小姐', '', '一般', '海帶', 1, '2026/07/14 16:00', '', '待製作', '舊訂單', 'U1', 20],
  ['uuid-new', 'A002', '2026/07/14 16:10:00', '王小姐', '0912345678', '第一份', '大豆干', 2, '2026/07/14 17:00', '小辣', '待製作', '大豆干兩份，小辣', 'U1', 20],
  ['uuid-new', 'A002', '2026/07/14 16:10:00', '王小姐', '0912345678', '第二份', '甜不辣', 1, '2026/07/14 17:00', '不辣', '待製作', '大豆干兩份，小辣', 'U1', 25],
  ['uuid-other', 'A003', '2026/07/14 16:15:00', '李小姐', '', '一般', '貢丸', 1, '2026/07/14 17:10', '', '待製作', '貢丸一份', 'U2', 15],
];

class FixedDate extends Date {
  constructor(...args) {
    if (args.length) {
      super(...args);
      return;
    }

    super(2026, 6, 14, 16, 20, 0);
  }
}

const context = vm.createContext({
  console,
  Date: FixedDate,
  Utilities: {
    formatDate: (date, timeZone, format) => {
      const formatter = new Intl.DateTimeFormat('en-CA', {
        timeZone,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        hourCycle: 'h23',
      });
      const parts = Object.fromEntries(
        formatter.formatToParts(date).map((part) => [part.type, part.value]),
      );

      if (format === 'yyyy/MM/dd HH:mm') {
        return `${parts.year}/${parts.month}/${parts.day} ${parts.hour}:${parts.minute}`;
      }

      return `${parts.hour}:${parts.minute}`;
    },
  },
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ({
      getSheetByName: (name) => name === 'Orders'
        ? {
          getDataRange: () => ({
            getValues: () => orderRows.map((row, index) => {
              const copy = row.slice();

              if (index > 0) {
                copy[8] = new Date('2026-07-14T09:00:00.000Z');
              }

              return copy;
            }),
            getDisplayValues: () => orderRows.map((row) => row.slice()),
          }),
        }
        : null,
    }),
  },
});

const code = fs.readFileSync(path.join(__dirname, '..', 'code.gs'), 'utf8');
vm.runInContext(code, context);

const editable = context.findEditableOrder('U1');
const editableOrder = JSON.parse(JSON.stringify(editable.order));

assert.equal(editable.systemOrderId, 'uuid-new');
assert.equal(editable.displayOrderId, 'A002');
assert.equal(editable.rawMessage, '大豆干兩份，小辣');
assert.equal(editableOrder.customer_name, '王小姐');
assert.equal(editableOrder.phone, '0912345678');
assert.equal(editableOrder.pickup_time, '2026/07/14 17:00');
assert.equal(editableOrder.pickup_time_source, 'estimated');
assert.equal(editableOrder.groups.length, 2);
assert.deepEqual(editableOrder.groups[0], {
  name: '第一份',
  items: [{
    name: '大豆干',
    raw_name: '大豆干',
    suggested_name: '',
    quantity: 2,
  }],
  note: '小辣',
});
assert.equal(editableOrder.groups[1].name, '第二份');

assert.equal(
  context.appendRawMessage_('大豆干兩份，小辣', '再一份一樣的，但不要辣'),
  '大豆干兩份，小辣\n\n再一份一樣的，但不要辣',
);
assert.equal(context.appendRawMessage_('', '謝謝'), '謝謝');

const mergedOrder = JSON.parse(JSON.stringify(editableOrder));
mergedOrder.pickup_time = '2026/07/14 15:30';
context.updateEditablePickupTime_(
  mergedOrder,
  editable,
  '第二份也不要酸菜',
);
assert.equal(mergedOrder.pickup_time, '2026/07/14 17:00');
assert.equal(mergedOrder.pickup_time_source, 'estimated');

const expandedOrder = JSON.parse(JSON.stringify(editableOrder));
expandedOrder.groups.push({
  name: '第三份',
  items: [
    { name: '黑輪', quantity: 2 },
    { name: '芋粿', quantity: 1 },
  ],
  note: '不辣',
});
context.updateEditablePickupTime_(expandedOrder, editable, '再一份黑輪兩個、芋粿一個');
assert.equal(expandedOrder.pickup_time, '2026/07/14 17:02');
assert.equal(expandedOrder.pickup_time_source, 'estimated');

const rescheduledOrder = { pickup_time: '2026/07/14 18:30' };
context.updateEditablePickupTime_(
  rescheduledOrder,
  editable,
  '改成晚上六點半取餐',
);
assert.equal(rescheduledOrder.pickup_time, '2026/07/14 18:30');
assert.equal(rescheduledOrder.pickup_time_source, 'requested');

const requestedExpandedOrder = JSON.parse(JSON.stringify(expandedOrder));
requestedExpandedOrder.pickup_time = '2026/07/14 17:00';
context.updateEditablePickupTime_(
  requestedExpandedOrder,
  {
    createdAt: editable.createdAt,
    rawMessage: '大豆干兩份，17:00取餐',
    order: {
      ...editableOrder,
      pickup_time_source: 'requested',
    },
  },
  '再加黑輪兩個',
);
assert.equal(requestedExpandedOrder.pickup_time, '2026/07/14 17:00');
assert.equal(requestedExpandedOrder.pickup_time_source, 'requested');
assert.equal(context.messageUpdatesPickupTime_('再一份不要辣'), false);
assert.equal(context.messageUpdatesPickupTime_('17:45 取餐'), true);

console.log('Order context tests passed');
