export type ZipEntry = {
  path: string;
  data: string | Uint8Array;
};

type PreparedZipEntry = {
  path: string;
  pathBytes: Uint8Array;
  dataBytes: Uint8Array;
  crc: number;
  offset: number;
};

const encoder = new TextEncoder();

export function createZip(entries: ZipEntry[]): Uint8Array {
  let offset = 0;
  const prepared = entries.map((entry): PreparedZipEntry => {
    const dataBytes = typeof entry.data === "string" ? encoder.encode(entry.data) : entry.data;
    const pathBytes = encoder.encode(entry.path);
    const preparedEntry = {
      path: entry.path,
      pathBytes,
      dataBytes,
      crc: crc32(dataBytes),
      offset,
    };
    offset += 30 + pathBytes.length + dataBytes.length;
    return preparedEntry;
  });

  const localFiles = prepared.map(writeLocalFile);
  const centralDirectoryOffset = byteLength(localFiles);
  const centralDirectory = prepared.map(writeCentralDirectoryFile);
  const centralDirectorySize = byteLength(centralDirectory);
  const end = writeEndOfCentralDirectory(
    entries.length,
    centralDirectorySize,
    centralDirectoryOffset,
  );

  return concatBytes([...localFiles, ...centralDirectory, end]);
}

export function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;

  for (const byte of data) {
    crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ byte) & 0xff]!;
  }

  return (crc ^ 0xffffffff) >>> 0;
}

export function listZipEntries(zip: Uint8Array): string[] {
  const decoder = new TextDecoder();
  const entries: string[] = [];
  let cursor = 0;

  while (cursor + 30 < zip.length && readUint32(zip, cursor) === 0x04034b50) {
    const compressedSize = readUint32(zip, cursor + 18);
    const fileNameLength = readUint16(zip, cursor + 26);
    const extraLength = readUint16(zip, cursor + 28);
    const nameStart = cursor + 30;
    const nameEnd = nameStart + fileNameLength;
    entries.push(decoder.decode(zip.slice(nameStart, nameEnd)));
    cursor = nameEnd + extraLength + compressedSize;
  }

  return entries;
}

function writeLocalFile(entry: PreparedZipEntry): Uint8Array {
  const header = new Uint8Array(30);
  const view = new DataView(header.buffer);
  view.setUint32(0, 0x04034b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 0, true);
  view.setUint16(8, 0, true);
  view.setUint16(10, 0, true);
  view.setUint16(12, 0, true);
  view.setUint32(14, entry.crc, true);
  view.setUint32(18, entry.dataBytes.length, true);
  view.setUint32(22, entry.dataBytes.length, true);
  view.setUint16(26, entry.pathBytes.length, true);
  view.setUint16(28, 0, true);

  return concatBytes([header, entry.pathBytes, entry.dataBytes]);
}

function writeCentralDirectoryFile(entry: PreparedZipEntry): Uint8Array {
  const header = new Uint8Array(46);
  const view = new DataView(header.buffer);
  view.setUint32(0, 0x02014b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 20, true);
  view.setUint16(8, 0, true);
  view.setUint16(10, 0, true);
  view.setUint16(12, 0, true);
  view.setUint16(14, 0, true);
  view.setUint32(16, entry.crc, true);
  view.setUint32(20, entry.dataBytes.length, true);
  view.setUint32(24, entry.dataBytes.length, true);
  view.setUint16(28, entry.pathBytes.length, true);
  view.setUint16(30, 0, true);
  view.setUint16(32, 0, true);
  view.setUint16(34, 0, true);
  view.setUint16(36, 0, true);
  view.setUint32(38, 0, true);
  view.setUint32(42, entry.offset, true);

  return concatBytes([header, entry.pathBytes]);
}

function writeEndOfCentralDirectory(
  entryCount: number,
  centralDirectorySize: number,
  centralDirectoryOffset: number,
): Uint8Array {
  const end = new Uint8Array(22);
  const view = new DataView(end.buffer);
  view.setUint32(0, 0x06054b50, true);
  view.setUint16(4, 0, true);
  view.setUint16(6, 0, true);
  view.setUint16(8, entryCount, true);
  view.setUint16(10, entryCount, true);
  view.setUint32(12, centralDirectorySize, true);
  view.setUint32(16, centralDirectoryOffset, true);
  view.setUint16(20, 0, true);

  return end;
}

function concatBytes(chunks: Uint8Array[]): Uint8Array {
  const output = new Uint8Array(byteLength(chunks));
  let cursor = 0;

  for (const chunk of chunks) {
    output.set(chunk, cursor);
    cursor += chunk.length;
  }

  return output;
}

function byteLength(chunks: Uint8Array[]): number {
  return chunks.reduce((length, chunk) => length + chunk.length, 0);
}

function readUint16(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset + offset, 2).getUint16(0, true);
}

function readUint32(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset + offset, 4).getUint32(0, true);
}

const CRC_TABLE = Array.from({ length: 256 }, (_, index) => {
  let crc = index;
  for (let bit = 0; bit < 8; bit += 1) {
    crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
  }
  return crc >>> 0;
});
