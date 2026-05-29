import type { ExportDocxCommand } from "../app/exportCommand";
import { createDocxBlob } from "../ooxml/docx";

export function downloadDocx(command: ExportDocxCommand) {
  const safeTitle = slugifyFileName(command.title || "Untitled Document");
  const blob = createDocxBlob(command.projection, {
    title: command.title || "Untitled Document",
    creator: "Docs OOXML",
    createdAt: new Date(),
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${safeTitle}.docx`;
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
