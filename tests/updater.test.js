'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { EventEmitter } = require('node:events');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const updaterSource = fs.readFileSync(path.join(__dirname, '..', 'updater.js'), 'utf8');

function makeUpdater({ release, statusCode = 200, localVersion = '2.0.5' }) {
  const requests = [];
  const events = [];
  const errors = [];
  const module = { exports: {} };
  const ipcMain = { once() {} };
  const https = {
    get(url, options, callback) {
      requests.push({ url, options });
      queueMicrotask(() => {
        const response = new EventEmitter();
        response.statusCode = statusCode;
        response.headers = {};
        callback(response);
        if (statusCode === 200) {
          response.emit('data', JSON.stringify(release));
          response.emit('end');
        }
      });
      return new EventEmitter();
    },
  };
  const requireModule = id => {
    if (id === 'electron') return { app: { getVersion: () => localVersion }, ipcMain };
    if (id === 'https') return https;
    return require(id);
  };
  vm.runInNewContext(updaterSource, {
    module,
    require: requireModule,
    console: { log() {}, error: (...args) => errors.push(args.join(' ')) },
  }, { filename: 'updater.js' });
  const mainWindow = {
    isDestroyed: () => false,
    webContents: { send: (channel, data) => events.push({ channel, data }) },
  };
  return { check: () => module.exports.checkForUpdates(mainWindow), requests, events, errors };
}

test('checks the public release without an Authorization header and offers the Portable exe', async () => {
  const updater = makeUpdater({
    release: {
      tag_name: 'v2.0.6',
      assets: [
        { name: 'contract-generator-setup-2.0.6.exe', browser_download_url: 'https://example.com/setup' },
        { name: 'latest.yml', browser_download_url: 'https://example.com/yml' },
        { name: 'contract-generator-2.0.6.exe', browser_download_url: 'https://example.com/portable' },
      ],
    },
  });
  await updater.check();
  assert.match(updater.requests[0].url, /\/releases\/latest$/);
  assert.equal(updater.requests[0].options.headers.Authorization, undefined);
  assert.equal(updater.events.length, 1);
  assert.equal(updater.events[0].channel, 'update-available');
  assert.equal(updater.events[0].data.version, '2.0.6');
  assert.equal(updater.events[0].data.assetUrl, 'https://example.com/portable');
});

test('does not offer an installer or a blockmap as a Portable update', async () => {
  const updater = makeUpdater({
    release: {
      tag_name: 'v2.0.6',
      assets: [
        { name: 'contract-generator-setup-2.0.6.exe', browser_download_url: 'https://example.com/setup' },
        { name: 'contract-generator-setup-2.0.6.exe.blockmap', browser_download_url: 'https://example.com/blockmap' },
      ],
    },
  });
  await updater.check();
  assert.equal(updater.events.length, 0);
  assert.match(updater.errors[0], /не содержит Portable \.exe/);
});

test('does not offer the same version again', async () => {
  const updater = makeUpdater({ release: { tag_name: 'v2.0.5', assets: [] } });
  await updater.check();
  assert.equal(updater.events.length, 0);
});

test('reports an API failure instead of silently treating it as no update', async () => {
  const updater = makeUpdater({ statusCode: 401 });
  await updater.check();
  assert.equal(updater.events.length, 0);
  assert.match(updater.errors[0], /HTTP 401/);
});