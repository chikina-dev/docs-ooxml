export function escapeXml(value: string): string {
  let escaped = "";
  let chunkStart = 0;

  for (let index = 0; index < value.length; index += 1) {
    const replacement = xmlEntityForCharacter(value.charCodeAt(index));
    if (!replacement) {
      continue;
    }

    escaped += value.slice(chunkStart, index) + replacement;
    chunkStart = index + 1;
  }

  return escaped.length === 0 ? value : escaped + value.slice(chunkStart);
}

function xmlEntityForCharacter(charCode: number): string {
  switch (charCode) {
    case 34:
      return "&quot;";
    case 38:
      return "&amp;";
    case 39:
      return "&apos;";
    case 60:
      return "&lt;";
    case 62:
      return "&gt;";
    default:
      return "";
  }
}
