import type { OutputProjection } from "../pipeline/outputProjectionTypes";
import type { AuthoringSnapshot } from "./authoringTypes";

export type ExportDocxCommand = {
  kind: "exportDocx";
  title: string;
  projection: OutputProjection;
};

export function createExportDocxCommand(snapshot: AuthoringSnapshot): ExportDocxCommand {
  return {
    kind: "exportDocx",
    title: snapshot.title,
    projection: snapshot.pipeline.outputProjection,
  };
}
