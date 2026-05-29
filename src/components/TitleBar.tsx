import type { AuthoringCommands, AuthoringSnapshot } from "../app/authoringTypes";

export function TitleBar({
  commands,
  snapshot,
}: {
  commands: Pick<AuthoringCommands, "setTitle">;
  snapshot: Pick<AuthoringSnapshot, "title">;
}) {
  return (
    <header className="titlebar">
      <div className="brand-lockup">
        <span className="word-mark" aria-hidden="true">
          W
        </span>
        <div>
          <p className="eyebrow">Docs OOXML</p>
          <h1>Word authoring pipeline</h1>
        </div>
      </div>
      <label className="title-field">
        <span>Document title</span>
        <input value={snapshot.title} onChange={(event) => commands.setTitle(event.target.value)} />
      </label>
    </header>
  );
}
