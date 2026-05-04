'use strict';
/**
 * scripts/gen-icon.js
 * Generates assets/icon.ico — a multi-size Windows icon in pure ICO/BMP format.
 * No external dependencies required.
 *
 * Design: document-conversion theme — blue background, white document shape
 *         with a folded corner, and three blue text lines.
 */

const fs   = require('fs');
const path = require('path');

const ASSETS_DIR = path.join(__dirname, '..', 'assets');
if (!fs.existsSync(ASSETS_DIR)) fs.mkdirSync(ASSETS_DIR, { recursive: true });

// Colour palette (R, G, B)
const C_BLUE  = [0x1d, 0x4e, 0xd8]; // #1d4ed8 — primary
const C_FOLD  = [0x3b, 0x82, 0xf6]; // #3b82f6 — lighter blue for folded flap
const C_WHITE = [0xff, 0xff, 0xff]; // document body

// ---------------------------------------------------------------------------
// Draw a single frame at the requested pixel size.
// Returns a Buffer: BITMAPINFOHEADER + BGRA pixels (bottom-up) + AND mask.
// ---------------------------------------------------------------------------
function makeFrame(sz) {
  // RGBA canvas, top-down (converted to BMP bottom-up later)
  const canvas = new Uint8Array(sz * sz * 4);

  function set(x, y, r, g, b, a = 0xff) {
    if (x < 0 || x >= sz || y < 0 || y >= sz) return;
    const i = (y * sz + x) * 4;
    canvas[i] = r; canvas[i + 1] = g; canvas[i + 2] = b; canvas[i + 3] = a;
  }

  // ── Blue background ──────────────────────────────────────────────────────
  for (let y = 0; y < sz; y++)
    for (let x = 0; x < sz; x++)
      set(x, y, ...C_BLUE);

  // ── Document boundary ────────────────────────────────────────────────────
  const ml   = Math.round(sz * 0.20);          // left edge
  const mr   = sz - Math.round(sz * 0.12) - 1; // right edge
  const mt   = Math.round(sz * 0.10);          // top edge
  const mb   = sz - Math.round(sz * 0.10) - 1; // bottom edge
  const fold = Math.round(sz * 0.24);          // folded-corner triangle size

  // White document body — skip the top-right triangular cut
  for (let y = mt; y <= mb; y++) {
    for (let x = ml; x <= mr; x++) {
      const dx = x - (mr - fold); // horizontal distance into fold zone
      const dy = y - mt;          // vertical distance from top
      if (dx > 0 && dy < fold && dx + dy < fold) continue; // inside the cut
      set(x, y, ...C_WHITE);
    }
  }

  // Lighter-blue folded flap (fills the triangle that was cut from the corner)
  for (let y = mt; y < mt + fold; y++) {
    for (let x = mr - fold + 1; x <= mr; x++) {
      const dx = x - (mr - fold);
      const dy = y - mt;
      if (dx + dy < fold) set(x, y, ...C_FOLD);
    }
  }

  // ── Text lines (blue stripes on the white body) ──────────────────────────
  const lineH = Math.max(1, Math.round(sz * 0.05));
  const lineL = ml + Math.round(sz * 0.08);
  const lineR = mr - Math.round(sz * 0.10);

  [0.50, 0.60, 0.70].forEach(pct => {
    const ly = Math.round(sz * pct);
    for (let y = ly; y < ly + lineH && y < sz; y++)
      for (let x = lineL; x <= lineR; x++)
        set(x, y, ...C_BLUE);
  });

  // ── Convert RGBA top-down → BGRA bottom-up for BMP ──────────────────────
  const pixelsBGRA = Buffer.alloc(sz * sz * 4);
  for (let y = 0; y < sz; y++) {
    const bmpY = sz - 1 - y; // BMP row order is bottom-up
    for (let x = 0; x < sz; x++) {
      const src = (y * sz + x) * 4;
      const dst = (bmpY * sz + x) * 4;
      pixelsBGRA[dst]     = canvas[src + 2]; // B
      pixelsBGRA[dst + 1] = canvas[src + 1]; // G
      pixelsBGRA[dst + 2] = canvas[src + 0]; // R
      pixelsBGRA[dst + 3] = canvas[src + 3]; // A
    }
  }

  // ── BITMAPINFOHEADER (40 bytes) ──────────────────────────────────────────
  const bih = Buffer.alloc(40);
  bih.writeInt32LE(40, 0);       // biSize
  bih.writeInt32LE(sz, 4);       // biWidth
  bih.writeInt32LE(sz * 2, 8);   // biHeight: doubled (XOR mask + AND mask in ICO)
  bih.writeUInt16LE(1, 12);      // biPlanes
  bih.writeUInt16LE(32, 14);     // biBitCount = 32-bit BGRA
  bih.writeUInt32LE(0, 16);      // biCompression = BI_RGB (0)
  // remaining 24 bytes stay zero

  // ── AND mask (all zeros = opaque; alpha channel carries actual transparency) ─
  // Each row padded to a DWORD (4-byte) boundary.
  const maskRowBytes = Math.ceil(sz / 32) * 4;
  const andMask = Buffer.alloc(maskRowBytes * sz, 0);

  return Buffer.concat([bih, pixelsBGRA, andMask]);
}

// ---------------------------------------------------------------------------
// Assemble multi-size ICO
// ---------------------------------------------------------------------------
const SIZES  = [16, 32, 48, 256];
const frames = SIZES.map(makeFrame);

// ICONDIR header (6 bytes)
const iconDir = Buffer.alloc(6);
iconDir.writeUInt16LE(0, 0);           // reserved
iconDir.writeUInt16LE(1, 2);           // type = 1 (ICO)
iconDir.writeUInt16LE(SIZES.length, 4);

// Each ICONDIRENTRY is 16 bytes; image data starts after all entries
let dataOffset = 6 + SIZES.length * 16;

const entries = SIZES.map((sz, i) => {
  const entry = Buffer.alloc(16);
  entry.writeUInt8(sz < 256 ? sz : 0, 0); // bWidth  (0 = 256 per ICO spec)
  entry.writeUInt8(sz < 256 ? sz : 0, 1); // bHeight
  entry.writeUInt8(0, 2);                 // bColorCount (0 = use biBitCount)
  entry.writeUInt8(0, 3);                 // reserved
  entry.writeUInt16LE(1, 4);              // wPlanes
  entry.writeUInt16LE(32, 6);             // wBitCount
  entry.writeUInt32LE(frames[i].length, 8);  // dwBytesInRes
  entry.writeUInt32LE(dataOffset, 12);        // dwImageOffset
  dataOffset += frames[i].length;
  return entry;
});

const ico     = Buffer.concat([iconDir, ...entries, ...frames]);
const icoPath = path.join(ASSETS_DIR, 'icon.ico');
fs.writeFileSync(icoPath, ico);

console.log(
  `Created ${path.relative(process.cwd(), icoPath)}  ` +
  `(${(ico.length / 1024).toFixed(1)} KB, sizes: ${SIZES.map(s => s + 'x' + s).join(', ')})`
);
