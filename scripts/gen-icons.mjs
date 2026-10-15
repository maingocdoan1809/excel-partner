#!/usr/bin/env node
import { deflateSync, inflateSync } from "node:zlib";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const OUT = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "addin", "assets");
const source = readFileSync(path.join(OUT, "excel-partner-mark.png"));
const width = source.readUInt32BE(16);
const height = source.readUInt32BE(20);
if (source[24] !== 8 || source[25] !== 6 || source[28] !== 0) {
  throw new Error("Source icon must be non-interlaced 8-bit RGBA PNG.");
}
const compressed = [];
for (let offset = 8; offset < source.length;) {
  const length = source.readUInt32BE(offset);
  if (source.toString("ascii", offset + 4, offset + 8) === "IDAT") compressed.push(source.subarray(offset + 8, offset + 8 + length));
  offset += length + 12;
}
const raw = inflateSync(Buffer.concat(compressed));
const stride = width * 4;
const pixels = Buffer.alloc(stride * height);
const paeth = (left, above, diagonal) => {
  const prediction = left + above - diagonal;
  const leftDistance = Math.abs(prediction - left);
  const aboveDistance = Math.abs(prediction - above);
  const diagonalDistance = Math.abs(prediction - diagonal);
  return leftDistance <= aboveDistance && leftDistance <= diagonalDistance ? left : aboveDistance <= diagonalDistance ? above : diagonal;
};
for (let row = 0; row < height; row++) {
  const filter = raw[row * (stride + 1)];
  if (filter > 4) throw new Error("Unsupported PNG filter.");
  for (let column = 0; column < stride; column++) {
    const position = row * stride + column;
    const left = column >= 4 ? pixels[position - 4] : 0;
    const above = row ? pixels[position - stride] : 0;
    const diagonal = row && column >= 4 ? pixels[position - stride - 4] : 0;
    const predictor = [0, left, above, Math.floor((left + above) / 2), paeth(left, above, diagonal)][filter];
    pixels[position] = (raw[row * (stride + 1) + column + 1] + predictor) & 255;
  }
}
function render(size) {
  const rows = Buffer.alloc((size * 4 + 1) * size);
  for (let row = 0; row < size; row++) {
    for (let column = 0; column < size; column++) {
      const sums = [0, 0, 0, 0];
      let total = 0;
      for (let sourceY = Math.floor(row * height / size); sourceY < Math.ceil((row + 1) * height / size); sourceY++) {
        for (let sourceX = Math.floor(column * width / size); sourceX < Math.ceil((column + 1) * width / size); sourceX++) {
          const weight = (Math.min(sourceX + 1, (column + 1) * width / size) - Math.max(sourceX, column * width / size))
            * (Math.min(sourceY + 1, (row + 1) * height / size) - Math.max(sourceY, row * height / size));
          const position = (sourceY * width + sourceX) * 4;
          const alpha = pixels[position + 3] / 255;
          for (let channel = 0; channel < 3; channel++) sums[channel] += pixels[position + channel] * alpha * weight;
          sums[3] += alpha * weight;
          total += weight;
        }
      }
      const target = row * (size * 4 + 1) + 1 + column * 4;
      for (let channel = 0; channel < 3; channel++) rows[target + channel] = sums[3] ? Math.round(sums[channel] / sums[3]) : 0;
      rows[target + 3] = Math.round(sums[3] / total * 255);
    }
  }
  return rows;
}

const crcTable = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
const crc32 = (buf) => {
  let c = 0xffffffff;
  for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};
const chunk = (type, data) => {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
};

function png(size) {
  const header = Buffer.alloc(13);
  header.writeUInt32BE(size, 0);
  header.writeUInt32BE(size, 4);
  header.set([8, 6, 0, 0, 0], 8); // 8-bit RGBA
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", header),
    chunk("IDAT", deflateSync(render(size))),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

mkdirSync(OUT, { recursive: true });
for (const size of [16, 32, 64, 80]) {
  writeFileSync(path.join(OUT, `icon-${size}.png`), png(size));
}
console.log(`icons written to ${OUT}`);
