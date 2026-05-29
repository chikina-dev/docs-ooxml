export { escapeXml } from "../pipeline/xmlEscaping";

export function xmlDeclaration(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
}
