'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { EventEmitter } = require('node:events');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const updaterSource = fs.readFileSync(path.join(__dirname, '..', 'updater.js'), 'utf8');

function makeUpdater({ releases, statusCode = 200, localVersion = '2.0.5', portable = true }) {
  const requests = [];
  const events = [];
  const errors = [];
  const listeners = [];
  const module = { exports: {} };
  const ipcMain = { once: channel => listeners.push(channel) };
  const https = {
    get(url, options, callback) {
      requests.push({ url, options });
      queueMicrotask(() => {
        const response = new EventEmitter();
        response.statusCode = statusCode;
        response.headers = {};
        callback(response);
        if (statusCode === 200) {
          response.emit('data', JSON.stringify(releases));
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
    process: { platform: 'win32', env: portable ? { PORTABLE_EXECUTABLE_FILE: 'C:\\old.exe' } : {} },
    console: { log() {}, error: (...args) => errors.push(args.join(' ')) },
  }, { filename: 'updater.js' });
  const mainWindow = {
    isDestroyed: () => false,
    webContents: { send: (channel, data) => events.push({ channel, data }) },
  };
  return { check: () => module.exports.checkForUpdates(mainWindow), requests, events, errors, listeners };
}

function release(tag_name, assets, extra = {}) {
  return { tag_name, assets, html_url: `https://github.com/example/releases/tag/${tag_name}`, ...extra };
}

const portableExe = version =>
  ({ name: `contract-generator-${version}.exe`, browser_download_url: `https://example.com/${version}/portable` });
const installerExe = version =>
  ({ name: `contract-generator-setup-${version}.exe`, browser_download_url: `https://example.com/${version}/setup` });

test('checks public releases without Authorization and skips a newer blockmap-only release', async () => {
  const updater = makeUpdater({
    releases: [
      release('2.0.6.1', [{ name: 'contract-generator-setup-2.0.6.exe.blockmap' }]),
      release('v2.0.6', [installerExe('2.0.6'), portableExe('2.0.6')]),
    ],
  });
  await updater.check();
  assert.match(updater.requests[0].url, /\/releases\?per_page=100$/);
  assert.equal(updater.requests[0].options.headers.Authorization, undefined);
  assert.equal(updater.events.length, 1);
  assert.equal(updater.events[0].channel, 'update-available');
  assert.equal(updater.events[0].data.version, '2.0.6');
  assert.equal(updater.events[0].data.mode, 'portable');
  assert.equal(updater.events[0].data.assetUrl, 'https://example.com/2.0.6/portable');
  assert.deepEqual(updater.listeners, ['update-start-download']);
});

test('does not offer an installer or a blockmap as a Portable update', async () => {
  const updater = makeUpdater({
    releases: [release('v2.0.6', [
      installerExe('2.0.6'),
      { name: 'contract-generator-setup-2.0.6.exe.blockmap', browser_download_url: 'https://example.com/blockmap' },
    ])],
  });
  await updater.check();
  assert.equal(updater.events.length, 0);
  assert.equal(updater.listeners.length, 0);
});

test('does not offer the same version again', async () => {
  const updater = makeUpdater({ releases: [release('v2.0.5', [portableExe('2.0.5')])] });
  await updater.check();
  assert.equal(updater.events.length, 0);
});

test('reports an API failure instead of silently treating it as no update', async () => {
  const updater = makeUpdater({ statusCode: 401 });
  await updater.check();
  assert.equal(updater.events.length, 0);
  assert.match(updater.errors[0], /HTTP 401/);
});

test('selects the highest usable version even if releases are not ordered by version', async () => {
  const updater = makeUpdater({
    localVersion: '2.0.3',
    releases: [
      release('v2.0.6', [portableExe('2.0.6')]),
      release('v2.0.7', [portableExe('2.0.7')]),
      release('2.0.6.1', [{ name: 'setup.blockmap' }]),
    ],
  });
  await updater.check();
  assert.equal(updater.events[0].data.version, '2.0.7');
});

test('compares all four numeric parts of a version', async () => {
  const updater = makeUpdater({
    localVersion: '2.0.6',
    releases: [release('2.0.6.1', [portableExe('2.0.6.1')])],
  });
  await updater.check();
  assert.equal(updater.events[0].data.version, '2.0.6.1');
});

test('ignores draft and prerelease versions', async () => {
  const updater = makeUpdater({
    releases: [
      release('v2.0.9', [portableExe('2.0.9')], { draft: true }),
      release('v2.0.8', [portableExe('2.0.8')], { prerelease: true }),
      release('v2.0.7', [portableExe('2.0.7')]),
    ],
  });
  await updater.check();
  assert.equal(updater.events[0].data.version, '2.0.7');
});

test('installed app opens the release page instead of replacing itself with a Portable exe', async () => {
  const updater = makeUpdater({
    portable: false,
    releases: [
      release('v2.0.8', [portableExe('2.0.8')]),
      release('v2.0.7', [portableExe('2.0.7'), installerExe('2.0.7')]),
    ],
  });
  await updater.check();
  assert.equal(updater.events[0].data.mode, 'manual');
  assert.equal(updater.events[0].data.version, '2.0.7');
  assert.equal(updater.events[0].data.releaseUrl, 'https://github.com/example/releases/tag/v2.0.7');
  assert.equal(updater.listeners.length, 0);
});
