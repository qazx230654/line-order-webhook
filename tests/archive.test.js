const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

class MemorySheet {
  constructor(values) {
    this.values = values.map((row) => row.slice());
  }

  getDataRange() {
    return {
      getValues: () => this.values.map((row) => row.slice()),
    };
  }
}

const orderHeader = [
  'system_order_id', 'display_order_id', 'created_at', 'customer_name',
  'phone', 'group_name', 'item_name', 'quantity', 'pickup_time', 'note',
  'status', 'raw_message', 'line_user_id', 'unit_price',
];

const archiveHeader = orderHeader.concat(['archived_at']);

const activeRows = [
  orderHeader,
  ['uuid-1', 'A001', '2026/07/02 12:00:00', '王小姐', '', '一般', '高麗菜', 2, '', '', '已完成', '', 'U1', 30],
  ['uuid-1', 'A001', '2026/07/02 12:00:00', '王小姐', '', '一般', '豆干', 1, '', '', '已完成', '', 'U1', 20],
  ['uuid-2', 'A002', '2026/07/03 12:00:00', '李小姐', '', '一般', '雞心', 1, '', '', '待製作', '', 'U2', 25],
  ['uuid-3', 'A003', '2026/07/04 12:00:00', '陳小姐', '', '一般', '未知品項', 1, '', '', '已完成', '', 'U3', 'UNPRICED'],
  ['uuid-4', 'A004', '2026/08/01 12:00:00', '林小姐', '', '一般', '高麗菜', 1, '', '', '已完成', '', 'U4', 30],
];

const archivedRows = [
  archiveHeader,
  activeRows[1].concat(['2026/08/02 03:00:00']),
  activeRows[2].concat(['2026/08/02 03:00:00']),
];

const sheets = {
  Orders: new MemorySheet(activeRows),
  OrdersArchive: new MemorySheet(archivedRows),
  PriceList: new MemorySheet([
    ['item_name', 'price', 'category', 'image_url', 'is_sold_out'],
    ['高麗菜', 30, '蔬菜', '', false],
    ['豆干', 20, '豆類', '', false],
    ['雞心', 25, '肉類', '', false],
  ]),
};

function formatDate(date, timeZone, format) {
  const formatter = new Intl.DateTimeFormat('en-CA', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hourCycle: 'h23',
  });
  const parts = Object.fromEntries(
    formatter.formatToParts(date).map((part) => [part.type, part.value]),
  );
  if (format === 'yyyy/MM') return `${parts.year}/${parts.month}`;
  if (format === 'yyyy/MM/dd') return `${parts.year}/${parts.month}/${parts.day}`;
  if (format === 'MM/dd') return `${parts.month}/${parts.day}`;
  if (format === 'd') return String(Number(parts.day));
  if (format === 'HH:mm') return `${parts.hour}:${parts.minute}`;
  if (format === 'MM/dd HH:mm') return `${parts.month}/${parts.day} ${parts.hour}:${parts.minute}`;
  if (format === 'yyyy/MM/dd HH:mm:ss') {
    return `${parts.year}/${parts.month}/${parts.day} ${parts.hour}:${parts.minute}:${parts.second}`;
  }
  return `${parts.year}/${parts.month}/${parts.day}`;
}

class FixedDate extends Date {
  constructor(...args) {
    super(...(args.length ? args : ['2026-08-14T04:00:00.000Z']));
  }

  static UTC(...args) {
    return Date.UTC(...args);
  }
}

const context = vm.createContext({
  console,
  Date: FixedDate,
  Utilities: { formatDate },
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ({
      getSheetByName: (name) => sheets[name] || null,
    }),
  },
});

const code = fs.readFileSync(path.join(__dirname, '..', 'code.gs'), 'utf8');
vm.runInContext(code, context);

const result = context.buildMonthlyArchiveData_(
  activeRows,
  '2026/07',
  {},
);

assert.deepEqual(
  Array.from(result.eligibleOrderIds),
  ['uuid-1', 'uuid-3'],
);
assert.equal(result.archiveRows.length, 3);
assert.equal(result.summary.orderCount, 2);
assert.equal(result.summary.completedCount, 2);
assert.equal(result.summary.revenue, 80);
assert.equal(result.summary.unpricedOrderCount, 1);
assert.equal(result.summary.averageOrderValue, 40);

const itemSummary = Object.fromEntries(
  Array.from(result.itemSummaries, (item) => [item.itemName, item]),
);
assert.equal(itemSummary['高麗菜'].quantity, 2);
assert.equal(itemSummary['高麗菜'].revenue, 60);
assert.equal(itemSummary['豆干'].revenue, 20);
assert.equal(itemSummary['未知品項'].revenue, 0);

const legacyResult = context.buildMonthlyArchiveData_([
  orderHeader,
  ['legacy-1', 'A099', '2026/07/05 12:00:00', '', '', '一般', '雞心', 2, '', '', '已完成', '', '', ''],
], '2026/07', { 雞心: 25 });

assert.equal(legacyResult.archiveRows[0][13], 25);
assert.equal(legacyResult.summary.revenue, 50);

const dashboardRows = context.getDashboardOrderRows_('2026/07');
assert.equal(dashboardRows.filter((row) => row[0] === 'uuid-1').length, 2);
assert.equal(dashboardRows.filter((row) => row[0] === 'uuid-2').length, 1);
assert.equal(dashboardRows.filter((row) => row[0] === 'uuid-3').length, 1);
assert.equal(dashboardRows.some((row) => row[0] === 'uuid-4'), false);

const dashboard = context.getAdminDashboardData('2026/07');
assert.equal(dashboard.selectedMonth, '2026/07');
assert.equal(dashboard.month.orders, 3);
assert.equal(dashboard.month.completed, 2);
assert.equal(dashboard.month.pending, 1);
assert.equal(dashboard.month.revenue, 105);
assert.equal(Array.isArray(dashboard.recentOrders), true);
assert.equal(dashboard.availableMonths.includes('2026/07'), true);

console.log('Archive tests passed');
