const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const properties = new Map();
let now = 1000;
let uuidCounter = 0;
let sleepCount = 0;
let onSleep = null;

class FixedDate extends Date {
  static now() {
    return now;
  }
}

const scriptProperties = {
  getProperty: (key) => properties.get(key) || null,
  setProperty: (key, value) => properties.set(key, value),
  deleteProperty: (key) => properties.delete(key),
};

const context = vm.createContext({
  console,
  Date: FixedDate,
  PropertiesService: {
    getScriptProperties: () => scriptProperties,
  },
  LockService: {
    getScriptLock: () => ({
      waitLock() {},
      releaseLock() {},
    }),
  },
  Utilities: {
    getUuid: () => `token-${++uuidCounter}`,
    sleep: (milliseconds) => {
      sleepCount += 1;
      now += milliseconds;

      if (onSleep) {
        const callback = onSleep;
        onSleep = null;
        callback();
      }
    },
  },
});

const code = fs.readFileSync(path.join(__dirname, '..', 'code.gs'), 'utf8');
vm.runInContext(code, context);

const firstUserLock = context.acquireLineUserOrderLock_('U1');
const otherUserLock = context.acquireLineUserOrderLock_('U2');

assert.notEqual(firstUserLock.key, otherUserLock.key);
assert.equal(sleepCount, 0);

onSleep = () => context.releaseLineUserOrderLock_(firstUserLock);

const queuedLock = context.acquireLineUserOrderLock_('U1');

assert.equal(queuedLock.key, firstUserLock.key);
assert.notEqual(queuedLock.token, firstUserLock.token);
assert.equal(sleepCount, 1);

context.releaseLineUserOrderLock_(firstUserLock);
assert.ok(properties.has(queuedLock.key));

context.releaseLineUserOrderLock_(queuedLock);
context.releaseLineUserOrderLock_(otherUserLock);

assert.equal(properties.size, 0);

console.log('Order lock tests passed');
