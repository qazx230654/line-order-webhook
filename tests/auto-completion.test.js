const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

class OrdersSheet {
  constructor(values) {
    this.values = values.map((row) => row.slice());
  }

  getDataRange() {
    return {
      getValues: () => this.values.map((row) => row.slice()),
      getDisplayValues: () => this.values.map((row) => row.slice()),
    };
  }

  getRangeList(ranges) {
    return {
      setValue: (value) => {
        ranges.forEach((range) => {
          const row = Number(range.slice(1));
          this.values[row - 1][10] = value;
        });
      },
    };
  }
}

const rows = [
  ['system_order_id', 'display_order_id', 'created_at', '', '', '', '', '', 'pickup_time', '', 'status', '', 'line_user_id'],
  ['uuid-1', 'A001', '2026/07/13 17:00', '', '', '', '高麗菜', 1, '2026/07/13 18:00', '', '待製作', '', 'U123'],
  ['uuid-1', 'A001', '2026/07/13 17:00', '', '', '', '豆干', 1, '2026/07/13 18:00', '', '待製作', '', 'U123'],
  ['uuid-2', 'A001', '2026/07/13 17:30', '', '', '', '雞心', 1, '2026/07/13 19:00', '', '待製作', '', 'U456'],
];

const ordersSheet = new OrdersSheet(rows);
let notificationCount = 0;

const context = vm.createContext({
  console,
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ({
      getSheetByName: (name) => (name === 'Orders' ? ordersSheet : null),
    }),
  },
  LockService: {
    getScriptLock: () => ({
      waitLock() {},
      releaseLock() {},
    }),
  },
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

      if (format === 'yyyy') return parts.year;
      if (format === 'HH:mm') return `${parts.hour}:${parts.minute}`;
      if (format === 'yyyy/MM/dd HH:mm') {
        return `${parts.year}/${parts.month}/${parts.day} ${parts.hour}:${parts.minute}`;
      }
      return `${parts.year}/${parts.month}/${parts.day}`;
    },
  },
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: () => 'test-secret',
    }),
  },
  UrlFetchApp: {
    fetch: () => {
      notificationCount += 1;
      return {};
    },
  },
});

const code = fs.readFileSync(path.join(__dirname, '..', 'code.gs'), 'utf8');
vm.runInContext(code, context);

assert.equal(
  context.parsePickupDateTime_(
    '2026/07/13 18:30',
    '2026/07/13 12:00',
  ).toISOString(),
  '2026-07-13T10:30:00.000Z',
);

assert.equal(
  context.parsePickupDateTime_(
    '晚上6點半',
    '2026/07/13 12:00',
  ).toISOString(),
  '2026-07-13T10:30:00.000Z',
);

assert.equal(
  context.parsePickupDateTime_(
    '明天 12:00',
    '2026/07/13 12:00',
  ).toISOString(),
  '2026-07-14T04:00:00.000Z',
);

assert.equal(
  context.parsePickupDateTime_(
    '2026/02/31 12:00',
    '2026/02/28 12:00',
  ),
  null,
);

assert.equal(context.parsePickupDateTime_('', '2026/07/13 12:00'), null);

const estimatedOrder = { pickup_time: '' };
const estimatedPickup = context.ensureOrderPickupTime_(
  estimatedOrder,
  10,
  new Date('2026-07-13T10:00:00.000Z'),
);

assert.equal(estimatedOrder.pickup_time, '2026/07/13 18:10');
assert.equal(estimatedPickup.displayTime, '18:10');
assert.equal(estimatedPickup.estimated, true);

const requestedOrder = { pickup_time: '2026/07/13 19:00' };
const requestedPickup = context.ensureOrderPickupTime_(
  requestedOrder,
  10,
  new Date('2026-07-13T10:00:00.000Z'),
);

assert.equal(requestedOrder.pickup_time, '2026/07/13 19:00');
assert.equal(requestedPickup.displayTime, '19:00');
assert.equal(requestedPickup.estimated, false);

const candidates = context.findOrdersDueForCompletion_([
  rows[0],
  rows[1],
  rows[2],
  ['uuid-2', 'A002', '2026/07/13 17:00', '', '', '', '', 1, '2026/07/13 19:00', '', '待製作'],
  ['uuid-3', 'A003', '2026/07/12 08:00', '', '', '', '', 1, '2026/07/12 09:00', '', '待製作'],
  ['uuid-4', 'A004', '2026/07/13 17:00', '', '', '', '', 1, '2026/07/13 18:00', '', '已完成'],
], new Date('2026-07-13T10:05:00.000Z'));

assert.deepEqual(
  Array.from(
    candidates,
    (order) => order.systemOrderId,
  ),
  ['uuid-1'],
);

assert.equal(context.getEstimatedWaitMinutes(), 10);

const firstUpdate = context.updateOrderStatus('uuid-1');
const secondUpdate = context.updateOrderStatus('uuid-1');

assert.equal(firstUpdate.notified, true);
assert.equal(secondUpdate.alreadyCompleted, true);
assert.equal(notificationCount, 1);
assert.equal(ordersSheet.values[1][10], '已完成');
assert.equal(ordersSheet.values[2][10], '已完成');
assert.equal(ordersSheet.values[3][10], '待製作');

console.log('Auto-completion tests passed');
