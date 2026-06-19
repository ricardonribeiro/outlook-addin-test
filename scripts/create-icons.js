#!/usr/bin/env node
/**
 * Generates minimal placeholder PNG icons for the Outlook add-in manifest.
 *
 * Run once from the repo root: node scripts/create-icons.js
 * Replace with real branded icons before deploying to production.
 *
 * No external dependencies — uses Node.js built-ins only.
 */

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const ICON_SIZES = [16, 32, 80];
const OUT_DIR = path.join(__dirname, '..', 'src', 'addin', 'public', 'assets');

// CRC32 table lookup
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[i] = c;
  }
  return t;
})();

function crc32(buf) {
  let crc = 0xffffffff;
  for (const b of buf) crc = CRC_TABLE[(crc ^ b) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type, data) {
  const typeBuf = Buffer.from(type, 'ascii');
  const lenBuf = Buffer.allocUnsafe(4);
  lenBuf.writeUInt32BE(data.length, 0);
  const crcInput = Buffer.concat([typeBuf, data]);
  const crcBuf = Buffer.allocUnsafe(4);
  crcBuf.writeUInt32BE(crc32(crcInput), 0);
  return Buffer.concat([lenBuf, typeBuf, data, crcBuf]);
}

function makePng(size) {
  // Build scanlines: filter byte (0x00 = None) + RGBA pixels (Indicium blue #0078D4)
  const scanline = Buffer.allocUnsafe(1 + size * 4);
  scanline[0] = 0x00; // filter: None
  for (let x = 0; x < size; x++) {
    const o = 1 + x * 4;
    scanline[o]     = 0x00; // R
    scanline[o + 1] = 0x78; // G
    scanline[o + 2] = 0xd4; // B
    scanline[o + 3] = 0xff; // A
  }
  const raw = Buffer.concat(Array.from({ length: size }, () => scanline));
  const compressed = zlib.deflateSync(raw);

  const ihdr = Buffer.allocUnsafe(13);
  ihdr.writeUInt32BE(size, 0);  // width
  ihdr.writeUInt32BE(size, 4);  // height
  ihdr[8]  = 8; // bit depth
  ihdr[9]  = 6; // colour type: RGBA
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace

  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), // PNG signature
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', compressed),
    pngChunk('IEND', Buffer.alloc(0)),
  ]);
}

fs.mkdirSync(OUT_DIR, { recursive: true });

for (const size of ICON_SIZES) {
  const png = makePng(size);
  const outPath = path.join(OUT_DIR, `icon-${size}.png`);
  fs.writeFileSync(outPath, png);
  console.log(`Created ${outPath} (${png.length} bytes)`);
}

console.log('\nDone. Replace with real icons before deploying.');
