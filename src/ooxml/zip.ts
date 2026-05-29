export type ZipEntry<Path extends string = string> = {
  path: Path;
  data: string | Uint8Array;
};

export type StoredZipEntry<Path extends string = string> = {
  path: Path;
  pathBytes: Uint8Array;
  dataBytes: Uint8Array;
  crc: number;
};

type PreparedZipEntry<Path extends string = string> = StoredZipEntry<Path> & {
  offset: number;
};

const encoder = new TextEncoder();

export function createZipNaive<Path extends string>(
  entries: readonly ZipEntry<Path>[],
): Uint8Array<ArrayBuffer> {
  const prepared = prepareEntries(entries);
  const localFiles = prepared.map(createLocalFile);
  const centralDirectoryOffset = byteLength(localFiles);
  const centralDirectory = prepared.map(createCentralDirectoryFile);
  const centralDirectorySize = byteLength(centralDirectory);
  const end = createEndOfCentralDirectory(
    entries.length,
    centralDirectorySize,
    centralDirectoryOffset,
  );

  return concatBytes([...localFiles, ...centralDirectory, end]);
}

export function createZipOptimized<Path extends string>(
  entries: readonly ZipEntry<Path>[],
): Uint8Array<ArrayBuffer> {
  return createZipFromStoredEntries(
    entries.map((entry) => createStoredZipEntry(entry.path, entry.data)),
  );
}

export function createZipFromStoredEntries<Path extends string>(
  entries: readonly StoredZipEntry<Path>[],
): Uint8Array<ArrayBuffer> {
  const centralDirectoryOffset = entries.reduce(
    (offset, entry) => offset + storedLocalFileLength(entry),
    0,
  );
  const centralDirectorySize = entries.reduce(
    (size, entry) => size + storedCentralDirectoryFileLength(entry),
    0,
  );
  const output = new Uint8Array(centralDirectoryOffset + centralDirectorySize + 22);
  let cursor = 0;

  for (const entry of entries) {
    cursor = writeStoredLocalFile(output, cursor, entry);
  }

  let localOffset = 0;
  for (const entry of entries) {
    cursor = writeStoredCentralDirectoryFile(output, cursor, entry, localOffset);
    localOffset += storedLocalFileLength(entry);
  }

  writeEndOfCentralDirectory(
    output,
    cursor,
    entries.length,
    centralDirectorySize,
    centralDirectoryOffset,
  );

  return output;
}

export function createStoredZipEntry<const Path extends string>(
  path: Path,
  data: string | Uint8Array,
): StoredZipEntry<Path> {
  const dataBytes = typeof data === "string" ? encoder.encode(data) : data;
  return {
    path,
    pathBytes: encoder.encode(path),
    dataBytes,
    crc: crc32(dataBytes),
  };
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

function writeStoredLocalFile(output: Uint8Array, offset: number, entry: StoredZipEntry): number {
  writeLocalFileHeader(output, offset, entry);
  output.set(entry.pathBytes, offset + 30);
  output.set(entry.dataBytes, offset + 30 + entry.pathBytes.length);

  return offset + storedLocalFileLength(entry);
}

function createLocalFile(entry: PreparedZipEntry): Uint8Array<ArrayBuffer> {
  return concatBytes([createLocalFileHeader(entry), entry.pathBytes, entry.dataBytes]);
}

function writeLocalFileHeader(output: Uint8Array, offset: number, entry: StoredZipEntry): void {
  const view = new DataView(output.buffer, output.byteOffset + offset, 30);
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
}

function createLocalFileHeader(entry: PreparedZipEntry): Uint8Array<ArrayBuffer> {
  const header = new Uint8Array(30);
  writeLocalFileHeader(header, 0, entry);
  return header;
}

function writeStoredCentralDirectoryFile(
  output: Uint8Array,
  offset: number,
  entry: StoredZipEntry,
  localOffset: number,
): number {
  writeCentralDirectoryFileHeader(output, offset, entry, localOffset);
  output.set(entry.pathBytes, offset + 46);

  return offset + storedCentralDirectoryFileLength(entry);
}

function writeCentralDirectoryFileHeader(
  output: Uint8Array,
  offset: number,
  entry: StoredZipEntry,
  localOffset: number,
): void {
  const view = new DataView(output.buffer, output.byteOffset + offset, 46);
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
  view.setUint32(42, localOffset, true);
}

function createCentralDirectoryFileHeader(entry: PreparedZipEntry): Uint8Array<ArrayBuffer> {
  const header = new Uint8Array(46);
  writeCentralDirectoryFileHeader(header, 0, entry, entry.offset);
  return header;
}

function createCentralDirectoryFile(entry: PreparedZipEntry): Uint8Array<ArrayBuffer> {
  return concatBytes([createCentralDirectoryFileHeader(entry), entry.pathBytes]);
}

function writeEndOfCentralDirectory(
  output: Uint8Array,
  offset: number,
  entryCount: number,
  centralDirectorySize: number,
  centralDirectoryOffset: number,
): void {
  const view = new DataView(output.buffer, output.byteOffset + offset, 22);
  view.setUint32(0, 0x06054b50, true);
  view.setUint16(4, 0, true);
  view.setUint16(6, 0, true);
  view.setUint16(8, entryCount, true);
  view.setUint16(10, entryCount, true);
  view.setUint32(12, centralDirectorySize, true);
  view.setUint32(16, centralDirectoryOffset, true);
  view.setUint16(20, 0, true);
}

function createEndOfCentralDirectory(
  entryCount: number,
  centralDirectorySize: number,
  centralDirectoryOffset: number,
): Uint8Array<ArrayBuffer> {
  const end = new Uint8Array(22);
  writeEndOfCentralDirectory(end, 0, entryCount, centralDirectorySize, centralDirectoryOffset);
  return end;
}

function concatBytes(chunks: Uint8Array[]): Uint8Array<ArrayBuffer> {
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

function prepareEntries<Path extends string>(
  entries: readonly ZipEntry<Path>[],
): PreparedZipEntry<Path>[] {
  return prepareStoredEntries(entries.map((entry) => createStoredZipEntry(entry.path, entry.data)));
}

function prepareStoredEntries<Path extends string>(
  entries: readonly StoredZipEntry<Path>[],
): PreparedZipEntry<Path>[] {
  let offset = 0;
  return entries.map((entry): PreparedZipEntry<Path> => {
    const preparedEntry = {
      path: entry.path,
      pathBytes: entry.pathBytes,
      dataBytes: entry.dataBytes,
      crc: entry.crc,
      offset,
    };
    offset += 30 + entry.pathBytes.length + entry.dataBytes.length;
    return preparedEntry;
  });
}

function storedLocalFileLength(entry: StoredZipEntry): number {
  return 30 + entry.pathBytes.length + entry.dataBytes.length;
}

function storedCentralDirectoryFileLength(entry: StoredZipEntry): number {
  return 46 + entry.pathBytes.length;
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
