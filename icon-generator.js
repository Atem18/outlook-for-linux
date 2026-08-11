'use strict';

/**
 * icon-generator.js
 *
 * Pure Node.js icon renderer — zero npm dependencies.
 * Uses only the built-in `zlib` module to write valid PNG files.
 *
 * Generates:
 *  - 256×256 app icon: white "O" ring on blue rounded square
 *  - 64×64 tray icon: same base + optional red badge with unread count
 *
 * nativeImage.createFromBuffer() is used for the tray — the only
 * method that works reliably on KDE Plasma / Linux.
 */

const zlib         = require('zlib');
const fs           = require('fs');
const path         = require('path');
const { nativeImage } = require('electron');

// ─── PNG encoder ─────────────────────────────────────────────────────────────

function crc32(buf) {
  let c = 0xFFFFFFFF;
  for (const b of buf) {
    c ^= b;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
  }
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function be32(n) { const b = Buffer.alloc(4); b.writeUInt32BE(n >>> 0, 0); return b; }

function pngChunk(type, data) {
  const t = Buffer.from(type, 'ascii');
  return Buffer.concat([be32(data.length), t, data, be32(crc32(Buffer.concat([t, data])))]);
}

// ─── Pixel canvas ─────────────────────────────────────────────────────────────

class Pixels {
  constructor(w, h) {
    this.w   = w;
    this.h   = h;
    this.buf = Buffer.alloc(w * h * 4, 0); // RGBA, fully transparent
  }

  // Alpha-blend one pixel
  set(x, y, r, g, b, a = 255) {
    x = x | 0; y = y | 0;
    if (x < 0 || x >= this.w || y < 0 || y >= this.h) return;
    const i  = (y * this.w + x) * 4;
    const sa = a / 255;
    const da = this.buf[i + 3] / 255;
    const oa = sa + da * (1 - sa);
    if (oa < 0.001) return;
    this.buf[i]     = ((r * sa + this.buf[i]     * da * (1 - sa)) / oa) | 0;
    this.buf[i + 1] = ((g * sa + this.buf[i + 1] * da * (1 - sa)) / oa) | 0;
    this.buf[i + 2] = ((b * sa + this.buf[i + 2] * da * (1 - sa)) / oa) | 0;
    this.buf[i + 3] = (oa * 255) | 0;
  }

  // Anti-aliased filled circle
  circle(cx, cy, r, col) {
    const [cr, cg, cb, ca = 255] = col;
    const x0 = Math.max(0,          Math.floor(cx - r - 1));
    const x1 = Math.min(this.w - 1, Math.ceil (cx + r + 1));
    const y0 = Math.max(0,          Math.floor(cy - r - 1));
    const y1 = Math.min(this.h - 1, Math.ceil (cy + r + 1));
    for (let y = y0; y <= y1; y++) {
      for (let x = x0; x <= x1; x++) {
        const d = Math.sqrt((x - cx) ** 2 + (y - cy) ** 2);
        const cov = Math.max(0, Math.min(1, r - d + 0.5));
        if (cov > 0) this.set(x, y, cr, cg, cb, (ca * cov) | 0);
      }
    }
  }

  // Anti-aliased rounded rectangle (inclusive pixel coords)
  roundRect(x0, y0, x1, y1, r, col) {
    const [cr, cg, cb, ca = 255] = col;
    for (let y = y0; y <= y1; y++) {
      for (let x = x0; x <= x1; x++) {
        const nx  = Math.max(x0 + r, Math.min(x1 - r, x));
        const ny  = Math.max(y0 + r, Math.min(y1 - r, y));
        const d   = Math.sqrt((x - nx) ** 2 + (y - ny) ** 2);
        const cov = Math.max(0, Math.min(1, r - d + 0.5));
        if (cov > 0) this.set(x, y, cr, cg, cb, (ca * cov) | 0);
      }
    }
  }

  // Encode to PNG buffer
  toPNG() {
    const { w, h, buf } = this;
    const stride = w * 4 + 1;
    const raw    = Buffer.alloc(stride * h, 0);
    for (let y = 0; y < h; y++) {
      raw[y * stride] = 0; // filter type: None
      buf.copy(raw, y * stride + 1, y * w * 4, (y + 1) * w * 4);
    }
    return Buffer.concat([
      Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), // PNG signature
      pngChunk('IHDR', Buffer.concat([be32(w), be32(h), Buffer.from([8, 6, 0, 0, 0])])),
      pngChunk('IDAT', zlib.deflateSync(raw, { level: 6 })),
      pngChunk('IEND', Buffer.alloc(0)),
    ]);
  }
}

// ─── 3×5 bitmap font for badge digits ────────────────────────────────────────
// Each entry: array of 5 numbers, each a 3-bit row (MSB = left column)

const FONT = {
  '0': [0b111, 0b101, 0b101, 0b101, 0b111],
  '1': [0b010, 0b110, 0b010, 0b010, 0b111],
  '2': [0b111, 0b001, 0b111, 0b100, 0b111],
  '3': [0b111, 0b001, 0b011, 0b001, 0b111],
  '4': [0b101, 0b101, 0b111, 0b001, 0b001],
  '5': [0b111, 0b100, 0b111, 0b001, 0b111],
  '6': [0b111, 0b100, 0b111, 0b101, 0b111],
  '7': [0b111, 0b001, 0b010, 0b010, 0b010],
  '8': [0b111, 0b101, 0b111, 0b101, 0b111],
  '9': [0b111, 0b101, 0b111, 0b001, 0b111],
  '+': [0b000, 0b010, 0b111, 0b010, 0b000],
};

const GLYPH_W = 3;
const GLYPH_H = 5;
const GLYPH_GAP = 1;

function glyphWidth(text, scale) {
  // total px width for a string at given scale
  return text.length * (GLYPH_W + GLYPH_GAP) * scale - GLYPH_GAP * scale;
}

function drawText(px, text, startX, startY, scale, col) {
  const [cr, cg, cb] = col;
  let cx = startX;
  for (const ch of text) {
    const g = FONT[ch];
    if (g) {
      for (let row = 0; row < GLYPH_H; row++) {
        for (let col2 = 0; col2 < GLYPH_W; col2++) {
          if (g[row] & (1 << (GLYPH_W - 1 - col2))) {
            for (let dy = 0; dy < scale; dy++)
              for (let dx = 0; dx < scale; dx++)
                px.set(cx + col2 * scale + dx, startY + row * scale + dy, cr, cg, cb, 255);
          }
        }
      }
    }
    cx += (GLYPH_W + GLYPH_GAP) * scale;
  }
}

// ─── Colours ──────────────────────────────────────────────────────────────────
const BLUE  = [0,   120, 212, 255];
const WHITE = [255, 255, 255, 255];
const RED   = [211,  47,  47, 255];

// ─── Draw the base icon (blue rounded square + white O ring) ─────────────────
function drawBase(px, size) {
  const half   = size / 2;
  const corner = size * 0.20;    // rounded corner radius

  px.roundRect(0, 0, size - 1, size - 1, corner, BLUE);

  // White "O" = white disk minus inner blue disk
  const outerR = size * 0.30;
  const innerR = size * 0.16;
  px.circle(half, half, outerR, WHITE);
  px.circle(half, half, innerR, BLUE);
}

// ─── 256×256 app icon ────────────────────────────────────────────────────────
function makeAppIconPNG() {
  const px = new Pixels(256, 256);
  drawBase(px, 256);
  return px.toPNG();
}

// ─── 128×128 tray icon with optional badge ───────────────────────────────────
// Using 128px so KDE has enough pixels to scale down to 16-22px panel size
// while keeping the badge number crisp and readable.
function makeTrayIconPNG(unread) {
  const SIZE = 128;
  const px   = new Pixels(SIZE, SIZE);
  drawBase(px, SIZE);

  if (unread > 0) {
    const label = unread > 99 ? '99+' : String(unread);
    // Keep scale constant for 1–2 digits; reduce only for the 3-char "99+" label
    const scale = label.length >= 3 ? 7 : 9;
    const tw    = glyphWidth(label, scale);
    const th    = GLYPH_H * scale;
    const pad   = 12;
    // Enforce a minimum radius so single-digit badges are as bold as Teams
    const br    = Math.max(38, Math.ceil(Math.max(tw, th) / 2) + pad);
    // Badge centre — bottom-right corner, allowed to bleed past edge
    const bx    = SIZE - br;
    const by_   = SIZE - br;

    // White halo for contrast against any panel colour
    px.circle(bx, by_, br + 6, WHITE);
    // Red fill
    px.circle(bx, by_, br, RED);
    // Centred number
    const tx = Math.round(bx - tw / 2);
    const ty = Math.round(by_ - th / 2);
    drawText(px, label, tx, ty, scale, WHITE);
  }

  return px.toPNG();
}

// ─── Bootstrap ───────────────────────────────────────────────────────────────
// 1. Write assets/icon.png (BrowserWindow + Notification need a file path)
// 2. Install into ~/.local/share/icons/hicolor/ so Wayland/KDE can resolve
//    Icon=outlook-for-linux from the .desktop file. Without this the taskbar
//    and window switcher show the generic Electron icon on Wayland.
const ASSETS_DIR  = path.join(__dirname, 'assets');
const ICON_PATH   = path.join(ASSETS_DIR, 'icon.png');
const HICOLOR_DIR = path.join(
  process.env.HOME || require('os').homedir(),
  '.local', 'share', 'icons', 'hicolor'
);

(function bootstrap() {
  const iconBuf = makeAppIconPNG();

  // assets/icon.png
  if (!fs.existsSync(ICON_PATH)) {
    fs.mkdirSync(ASSETS_DIR, { recursive: true });
    fs.writeFileSync(ICON_PATH, iconBuf);
  }

  // hicolor theme at 256x256 (and a 48x48 copy — KDE prefers 48 for taskbar)
  for (const size of ['256x256', '48x48']) {
    const dest = path.join(HICOLOR_DIR, size, 'apps', 'outlook-for-linux.png');
    if (!fs.existsSync(dest)) {
      try {
        fs.mkdirSync(path.dirname(dest), { recursive: true });
        fs.writeFileSync(dest, iconBuf);   // KDE scales from 256 just fine
      } catch (_) { /* non-fatal — deploy/install.sh handles this properly */ }
    }
  }
})();

// ─── Public API ───────────────────────────────────────────────────────────────

/**
 * Returns a nativeImage for the system tray.
 * Uses createFromBuffer with a real PNG — the only method that works
 * reliably on KDE Plasma / Linux SNI.
 */
function generateTrayIcon(unread) {
  const buf = makeTrayIconPNG(Math.max(0, unread | 0));
  return nativeImage.createFromBuffer(buf, { scaleFactor: 1.0 });
}

module.exports = { generateTrayIcon, ICON_PATH };
