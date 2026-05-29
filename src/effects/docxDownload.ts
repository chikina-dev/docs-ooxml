import { createDocxBlob, type DocxWriteStrategy } from "../ooxml/docx.ts";
import type { OutputProjection } from "../pipeline/types.ts";

export type DocxDownloadCommand = {
  title: string;
  projection: OutputProjection;
  strategy: DocxWriteStrategy;
};

export function downloadDocx(command: DocxDownloadCommand) {
  const safeTitle = slugifyFileName(command.title || "Untitled Document");
  const blob = createDocxBlob(
    command.projection,
    {
      title: command.title || "Untitled Document",
      creator: "Docs OOXML",
      createdAt: new Date(),
    },
    command.strategy,
  );
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${safeTitle}-${command.strategy}.docx`;
  link.click();
  URL.revokeObjectURL(url);
}

function slugifyFileName(value: string): string {
  return (
    value
      .trim()
      .replace(/[^\w.-]+/g, "-")
      .replace(/^-+|-+$/g, "") || "Untitled-Document"
  );
}
