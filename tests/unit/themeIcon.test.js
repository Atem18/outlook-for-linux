'use strict';

// Ported from teams-for-linux tests/unit/themeIcon.test.js.

const { describe, it, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { EventEmitter } = require('node:events');

const dbusPath = require.resolve('@homebridge/dbus-native');
const modulePath = require.resolve('../../app/appIcon/themeIcon');

describe('themeIcon', { skip: process.platform !== 'linux' }, () => {
  let tmp;
  let signals;
  let buses;
  const env = { ...process.env };

  beforeEach(() => {
    tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'theme-icon-'));
    signals = [];
    buses = [];
    process.env.XDG_DATA_HOME = tmp;
    process.env.XDG_DATA_DIRS = path.join(tmp, 'none');
    delete process.env.SNAP;
    delete process.env.FLATPAK_ID;
    delete process.env.APPIMAGE;
    require.cache[dbusPath] = {
      id: dbusPath,
      filename: dbusPath,
      loaded: true,
      exports: {
        sessionBus: () => {
          const bus = { connection: new EventEmitter(), sendSignal: (...args) => signals.push(args) };
          buses.push(bus);
          return bus;
        },
      },
    };
    delete require.cache[modulePath];
  });

  afterEach(() => {
    delete require.cache[dbusPath];
    delete require.cache[modulePath];
    for (const key of Object.keys(process.env)) delete process.env[key];
    Object.assign(process.env, env);
    fs.rmSync(tmp, { recursive: true, force: true });
  });

  const fakeImage = (width, height, tag = '') => ({
    getSize: () => ({ width, height }),
    resize: (opts) => ({ toPNG: () => Buffer.from(JSON.stringify({ ...opts, tag })) }),
  });
  const iconPath = (size, name = 'outlook-for-linux') =>
    path.join(tmp, 'icons', 'hicolor', `${size}x${size}`, 'apps', `${name}.png`);
  const readIcon = (size, name) => JSON.parse(fs.readFileSync(iconPath(size, name), 'utf8'));
  const manifest = () => JSON.parse(fs.readFileSync(path.join(tmp, 'outlook-for-linux', 'theme-icon.json'), 'utf8'));
  const writeDesktop = (file, text) => {
    const apps = path.join(tmp, 'applications');
    fs.mkdirSync(apps, { recursive: true });
    fs.writeFileSync(path.join(apps, file), text);
  };

  it('writes every hicolor size, keeping the aspect on the longer side, then removes them', () => {
    const themeIcon = require(modulePath);
    assert.strictEqual(themeIcon.installed(), false);
    assert.strictEqual(themeIcon.install(fakeImage(250, 264)), true);
    assert.strictEqual(fs.readdirSync(path.join(tmp, 'icons', 'hicolor')).length, 10);
    assert.deepStrictEqual(readIcon(48), { height: 48, tag: '' });
    assert.ok(!fs.existsSync(`${iconPath(48)}.tmp-${process.pid}`));
    assert.strictEqual(themeIcon.installed(), true);
    assert.deepStrictEqual(signals[0], ['/KIconLoader', 'org.kde.KIconLoader', 'iconChanged', 'i', [0]]);

    assert.strictEqual(themeIcon.remove(), true);
    assert.ok(!fs.existsSync(iconPath(48)));
    assert.strictEqual(themeIcon.installed(), false);
    assert.ok(!fs.existsSync(path.join(tmp, 'outlook-for-linux', 'theme-icon.json')));
  });

  it('removes only the files it wrote itself', () => {
    const themeIcon = require(modulePath);
    fs.mkdirSync(path.dirname(iconPath(1024)), { recursive: true });
    fs.writeFileSync(iconPath(1024), 'hand placed');
    assert.strictEqual(themeIcon.remove(), false);
    assert.ok(fs.existsSync(iconPath(1024)));

    themeIcon.install(fakeImage(64, 64));
    themeIcon.remove();
    assert.ok(fs.existsSync(iconPath(1024)));
    assert.ok(!fs.existsSync(iconPath(64)));
  });

  it('never overwrites a hand-placed icon at one of its sizes', () => {
    const themeIcon = require(modulePath);
    fs.mkdirSync(path.dirname(iconPath(48)), { recursive: true });
    fs.writeFileSync(iconPath(48), 'hand placed');

    assert.strictEqual(themeIcon.install(fakeImage(64, 64, 'first')), true);
    assert.strictEqual(themeIcon.install(fakeImage(64, 64, 'second')), true);
    assert.strictEqual(fs.readFileSync(iconPath(48), 'utf8'), 'hand placed');
    assert.strictEqual(readIcon(64).tag, 'second');
    assert.strictEqual(manifest()['outlook-for-linux'].length, 9);

    themeIcon.remove();
    assert.strictEqual(fs.readFileSync(iconPath(48), 'utf8'), 'hand placed');
    assert.ok(!fs.existsSync(iconPath(64)));
  });

  it('survives a failed session bus connection and reconnects on the next change', () => {
    const themeIcon = require(modulePath);
    themeIcon.install(fakeImage(64, 64));
    assert.strictEqual(buses.length, 1);
    assert.doesNotThrow(() => buses[0].connection.emit('error', new Error('connect ENOENT')));

    themeIcon.remove();
    assert.strictEqual(buses.length, 2);
    assert.strictEqual(signals.length, 2);
  });

  it('keeps the previous icon set whole when a replacement fails partway', () => {
    const themeIcon = require(modulePath);
    assert.strictEqual(themeIcon.install(fakeImage(64, 64, 'first')), true);
    fs.mkdirSync(`${iconPath(96)}.tmp-${process.pid}`);
    signals = [];

    assert.strictEqual(themeIcon.install(fakeImage(64, 64, 'second')), false);
    for (const size of [16, 48, 64, 512]) assert.strictEqual(readIcon(size).tag, 'first');
    assert.ok(!fs.existsSync(`${iconPath(16)}.tmp-${process.pid}`));
    assert.strictEqual(manifest()['outlook-for-linux'].length, 10);
    assert.strictEqual(signals.length, 0);
  });

  it('prefers the launcher for the window class over the packaged entry', () => {
    writeDesktop('outlook-for-linux.desktop', 'Exec=/usr/bin/distrobox enter outlook-for-linux\nIcon=distrobox-arch\n');
    writeDesktop('mine.desktop', 'Exec=/home/me/bin/outlook\nIcon=outlook-mine\nStartupWMClass=outlook-for-linux\n');
    writeDesktop('extra.desktop', 'Exec=/home/me/bin/outlook\nIcon=outlook-extra\nStartupWMClass=outlook-for-linux-extra\n');
    assert.strictEqual(require(modulePath).iconName(), 'outlook-mine');
  });

  it('prefers the packaged entry, then the launcher of this binary', () => {
    writeDesktop('other.desktop', `Exec=${process.execPath} %U\nIcon=other\n`);
    assert.strictEqual(require(modulePath).iconName(), 'other');

    writeDesktop('outlook-for-linux.desktop', 'Exec="/opt/Outlook for Linux/outlook-for-linux" %U\nIcon=outlook-for-linux\n');
    delete require.cache[modulePath];
    assert.strictEqual(require(modulePath).iconName(), 'outlook-for-linux');
  });

  it('falls back to the default name when the desktop file icon is not a plain name', () => {
    for (const icon of ['../../escape', '/usr/share/pixmaps/x.png', 'a b']) {
      writeDesktop('outlook-for-linux.desktop', `Icon=${icon}\n`);
      delete require.cache[modulePath];
      assert.strictEqual(require(modulePath).iconName(), 'outlook-for-linux');
    }
  });

  it('does nothing inside a snap', () => {
    process.env.SNAP = '/snap/outlook-for-linux/1';
    const themeIcon = require(modulePath);
    assert.strictEqual(themeIcon.supported(), false);
    assert.strictEqual(themeIcon.install(fakeImage(64, 64)), false);
    assert.strictEqual(themeIcon.installed(), false);
    assert.ok(!fs.existsSync(path.join(tmp, 'icons')));
  });

  it('inside a flatpak, writes the host icon theme under the flatpak id once it is granted', () => {
    process.env.FLATPAK_ID = 'io.github.Atem18.OutlookForLinux';
    process.env.XDG_DATA_HOME = path.join(tmp, 'sandbox');
    process.env.HOST_XDG_DATA_HOME = path.join(tmp, 'host');
    const themeIcon = require(modulePath);
    assert.strictEqual(themeIcon.supported(), false);

    fs.mkdirSync(path.join(tmp, 'host', 'icons'), { recursive: true });
    assert.strictEqual(themeIcon.supported(), true);
    assert.strictEqual(themeIcon.iconName(), 'io.github.Atem18.OutlookForLinux');
    assert.strictEqual(themeIcon.install(fakeImage(64, 64)), true);
    const icon = path.join(tmp, 'host', 'icons', 'hicolor', '48x48', 'apps', 'io.github.Atem18.OutlookForLinux.png');
    assert.ok(fs.existsSync(icon));
    assert.ok(fs.existsSync(path.join(tmp, 'sandbox', 'outlook-for-linux', 'theme-icon.json')));
    assert.strictEqual(themeIcon.remove(), true);
    assert.ok(!fs.existsSync(icon));
  });
});
