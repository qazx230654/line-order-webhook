const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const properties = new Map();

const scriptProperties = {
  getProperty: (key) => properties.get(key) || null,
  setProperty: (key, value) => properties.set(key, String(value)),
  setProperties: (values) => {
    Object.entries(values).forEach(([key, value]) => {
      properties.set(key, String(value));
    });
  },
  deleteProperty: (key) => properties.delete(key),
};

const context = vm.createContext({
  console,
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
    formatDate: () => '2026/07/14 10:30:00',
  },
  ContentService: {
    MimeType: { JSON: 'application/json' },
    createTextOutput: (content) => ({
      content,
      mimeType: '',
      setMimeType(mimeType) {
        this.mimeType = mimeType;
        return this;
      },
    }),
  },
});

const code = fs.readFileSync(path.join(__dirname, '..', 'code.gs'), 'utf8');
vm.runInContext(code, context);

assert.deepEqual(
  JSON.parse(JSON.stringify(context.getAiOrderStatus())),
  { enabled: true, updatedAt: '' },
);

const disabled = context.setAiOrderEnabled(false);
assert.deepEqual(
  JSON.parse(JSON.stringify(disabled)),
  { enabled: false, updatedAt: '2026/07/14 10:30:00' },
);
assert.equal(context.isAiOrderEnabled_(), false);
assert.throws(
  () => context.assertAiOrderEnabledForWrite_(),
  (error) => error.code === 'AI_ORDER_DISABLED',
);

const pausedOutput = context.createAssistantDisabledOutput_();
assert.equal(pausedOutput.mimeType, 'application/json');
assert.deepEqual(JSON.parse(pausedOutput.content), {
  is_order: false,
  assistant_disabled: true,
});

context.setAiOrderEnabled(true);
assert.equal(context.isAiOrderEnabled_(), true);
assert.doesNotThrow(() => context.assertAiOrderEnabledForWrite_());
assert.throws(() => context.setAiOrderEnabled('false'), /布林值/);

assert.match(code, /if \(requireAiOrderEnabled\) \{\s*assertAiOrderEnabledForWrite_\(\);/);
assert.match(code, /if \(!isAiOrderEnabled_\(\)\) \{\s*return createAssistantDisabledOutput_\(\);/);

console.log('AI order switch tests passed');
