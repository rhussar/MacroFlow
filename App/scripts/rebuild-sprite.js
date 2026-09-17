const fs = require('fs');
const path = require('path');
const sharp = require('sharp');

const RAW_DIR = path.join(__dirname, 'imagemso-raw');
const ASSETS_DIR = path.join(__dirname, '..', 'assets');
const MANIFEST_DIR = path.join(__dirname, '..', 'src', 'features', 'icons');
const ICON_SIZE = 32;
const SPRITE_COLS = 40;

const BAD = new Set([
  'AppointmentColor5','AutoSumAverage','AutoSumCount','AutoSumMax','AutoSumMin',
  'ClearContents','FontSize','Options','PageBreakInsert','PageBreakInsertExcel',
  'PageBreakRemove','PasteSpecial','PivotTableCalculatedItem','RecordAudio'
]);

function decodeBmp(buf) {
  if (buf[0] !== 0x42 || buf[1] !== 0x4D) return null;
  const dataOffset = buf.readUInt32LE(10);
  const w = buf.readInt32LE(18);
  const h = Math.abs(buf.readInt32LE(22));
  const bpp = buf.readUInt16LE(28);
  if (bpp !== 24 && bpp !== 32) return null;
  const ch = bpp / 8;
  const rowSize = Math.ceil((w * ch) / 4) * 4;

  const getPixel = (px, py) => {
    const srcRow = h - 1 - py;
    const off = dataOffset + srcRow * rowSize + px * ch;
    return [buf[off + 2], buf[off + 1], buf[off]];
  };
  const bg = getPixel(0, 0);

  let subtleCount = 0, totalNonBg = 0;
  for (let y = 0; y < h; y++) {
    const srcRow = h - 1 - y;
    for (let x = 0; x < w; x++) {
      const off = dataOffset + srcRow * rowSize + x * ch;
      const dist = Math.max(Math.abs(buf[off+2]-bg[0]), Math.abs(buf[off+1]-bg[1]), Math.abs(buf[off]-bg[2]));
      if (dist > 2) { totalNonBg++; if (dist < 60) subtleCount++; }
    }
  }
  const isSubtle = totalNonBg > 0 && (subtleCount / totalNonBg) > 0.8;
  const bgThresh = isSubtle ? 1 : 3;
  const rampEnd = isSubtle ? 15 : 40;

  const rgba = Buffer.alloc(w * h * 4);
  for (let y = 0; y < h; y++) {
    const srcRow = h - 1 - y;
    for (let x = 0; x < w; x++) {
      const srcOff = dataOffset + srcRow * rowSize + x * ch;
      const dstOff = (y * w + x) * 4;
      const r = buf[srcOff + 2], g = buf[srcOff + 1], b = buf[srcOff];
      const maxDist = Math.max(Math.abs(r-bg[0]), Math.abs(g-bg[1]), Math.abs(b-bg[2]));

      if (maxDist <= bgThresh) {
        rgba[dstOff] = 0; rgba[dstOff+1] = 0; rgba[dstOff+2] = 0; rgba[dstOff+3] = 0;
      } else if (maxDist >= rampEnd) {
        rgba[dstOff] = r; rgba[dstOff+1] = g; rgba[dstOff+2] = b; rgba[dstOff+3] = 255;
      } else {
        const alpha = Math.round(((maxDist - bgThresh) / (rampEnd - bgThresh)) * 255);
        const a = alpha / 255;
        if (a > 0.01) {
          rgba[dstOff]   = Math.min(255, Math.max(0, Math.round((r - bg[0]*(1-a)) / a)));
          rgba[dstOff+1] = Math.min(255, Math.max(0, Math.round((g - bg[1]*(1-a)) / a)));
          rgba[dstOff+2] = Math.min(255, Math.max(0, Math.round((b - bg[2]*(1-a)) / a)));
        } else {
          rgba[dstOff] = r; rgba[dstOff+1] = g; rgba[dstOff+2] = b;
        }
        rgba[dstOff+3] = alpha;
      }
    }
  }
  return { data: rgba, width: w, height: h, channels: 4 };
}

async function main() {
  const bmpFiles = fs.readdirSync(RAW_DIR).filter(f => f.endsWith('.bmp')).sort();
  const pngFiles = fs.readdirSync(RAW_DIR).filter(f => f.endsWith('.png') && !f.includes('_raw') && !f.includes('_test')).sort();
  const iconBuffers = new Map();

  for (const f of bmpFiles) {
    const name = f.replace('.bmp', '');
    if (BAD.has(name)) continue;
    const buf = fs.readFileSync(path.join(RAW_DIR, f));
    if (buf.length < 100) continue;
    const decoded = decodeBmp(buf);
    if (!decoded) continue;
    try {
      iconBuffers.set(name, await sharp(decoded.data, { raw: { width: decoded.width, height: decoded.height, channels: 4 } }).png().toBuffer());
    } catch {}
  }

  // PNGs override BMPs (custom replacements like border icons)
  for (const f of pngFiles) {
    const name = f.replace('.png', '');
    iconBuffers.set(name, await sharp(path.join(RAW_DIR, f)).resize(ICON_SIZE, ICON_SIZE).png().toBuffer());
  }

  const names = [...iconBuffers.keys()].sort();
  const manifest = {};
  const composites = [];

  for (let i = 0; i < names.length; i++) {
    const col = i % SPRITE_COLS;
    const row = Math.floor(i / SPRITE_COLS);
    manifest[names[i]] = { col, row, idx: i };
    composites.push({ input: iconBuffers.get(names[i]), left: col * ICON_SIZE, top: row * ICON_SIZE });
  }

  const rows = Math.ceil(composites.length / SPRITE_COLS);
  const w = SPRITE_COLS * ICON_SIZE;
  const h = rows * ICON_SIZE;

  await sharp({ create: { width: w, height: h, channels: 4, background: { r: 0, g: 0, b: 0, alpha: 0 } } })
    .composite(composites).png()
    .toFile(path.join(ASSETS_DIR, 'imagemso-sprite.png'));

  fs.writeFileSync(path.join(MANIFEST_DIR, 'imagemso-manifest.json'), JSON.stringify(manifest, null, 2));
  fs.writeFileSync(path.join(MANIFEST_DIR, 'imagemso-meta.json'), JSON.stringify({
    iconSize: ICON_SIZE, cols: SPRITE_COLS, rows, width: w, height: h, count: composites.length
  }, null, 2));

  console.log('Done!', composites.length, 'icons');
}

main().catch(e => { console.error(e); process.exit(1); });
