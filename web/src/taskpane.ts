// Oberfläche des Prototyps: pro Prüfung eine Schaltfläche, Ergebnis als Ampel mit Details,
// "Alle ausführen" und "Bericht kopieren" für die Rückmeldung.

import {
  checkAnhaenge,
  checkAnmeldung,
  checkAppOrdner,
  checkEml,
  checkMail,
  checkModell,
  checkUmgebung,
  checkUpload,
  checkVerknuepfungen,
  state,
  type CheckResult,
} from "./checks";

interface CheckDef {
  id: string;
  title: string;
  run: () => Promise<CheckResult>;
  /** Wird bei "Alle ausführen" übersprungen (schreibt in eine Projektbibliothek). */
  manualOnly?: boolean;
}

const targetSelect = () => document.getElementById("target") as HTMLSelectElement;

const checks: CheckDef[] = [
  { id: "umgebung", title: "1 · Umgebung und Requirement Sets", run: checkUmgebung },
  { id: "anmeldung", title: "2 · Anmeldung (NAA)", run: checkAnmeldung },
  { id: "mail", title: "3 · Aktuelle Mail / freigegebenes Postfach", run: checkMail },
  { id: "eml", title: "4 · .eml-Export (Office.js und Graph)", run: checkEml },
  { id: "anhaenge", title: "5 · Anhänge abrufen", run: checkAnhaenge },
  { id: "verknuepfungen", title: "6 · Projektbibliotheken finden (Verknüpfungen)", run: checkVerknuepfungen },
  { id: "appordner", title: "7 · Speicherort für den Verlauf", run: checkAppOrdner },
  {
    id: "upload",
    title: "8 · Test-Upload der .eml in gewählte Bibliothek",
    run: () => checkUpload(state.targets[Number(targetSelect().value)]),
    manualOnly: true,
  },
  { id: "modell", title: "9 · Vorschlagsmodell im Aufgabenbereich", run: checkModell },
];

const results = new Map<string, CheckResult>();

function render(): void {
  const list = document.getElementById("checks")!;
  list.innerHTML = "";
  for (const check of checks) {
    const result = results.get(check.id);
    const row = document.createElement("section");
    row.className = `check ${result?.status ?? "pending"}`;
    const header = document.createElement("div");
    header.className = "check-header";
    const title = document.createElement("span");
    title.textContent = check.title;
    const button = document.createElement("button");
    button.textContent = result ? "Erneut" : "Prüfen";
    button.onclick = () => run(check);
    header.append(title, button);
    row.append(header);
    if (check.id === "upload") {
      const select = document.createElement("select");
      select.id = "target";
      if (state.targets.length === 0) select.append(new Option("– erst Prüfung 6 ausführen –", ""));
      state.targets.forEach((t, i) => select.append(new Option(t.label, String(i))));
      row.append(select);
    }
    if (result) {
      const summary = document.createElement("p");
      summary.className = "summary";
      summary.textContent = result.summary;
      const details = document.createElement("pre");
      details.textContent = result.details.join("\n");
      row.append(summary, details);
    }
    list.append(row);
  }
}

async function run(check: CheckDef): Promise<void> {
  results.set(check.id, { status: "warn", summary: "läuft …", details: [] });
  render();
  try {
    results.set(check.id, await check.run());
  } catch (error) {
    results.set(check.id, { status: "fail", summary: "Unerwarteter Fehler", details: [String(error)] });
  }
  render();
}

async function runAll(): Promise<void> {
  for (const check of checks) if (!check.manualOnly) await run(check);
}

function report(): string {
  const lines = [`MailDrop Web-Prototyp – Prüfbericht ${new Date().toLocaleString("de-DE")}`, ""];
  for (const check of checks) {
    const r = results.get(check.id);
    lines.push(`[${r ? r.status.toUpperCase() : "NICHT AUSGEFÜHRT"}] ${check.title}`);
    if (r) lines.push(`  ${r.summary}`, ...r.details.map((d) => `    ${d}`));
    lines.push("");
  }
  return lines.join("\n");
}

async function copyReport(): Promise<void> {
  const text = report();
  try {
    await navigator.clipboard.writeText(text);
    setStatus("Bericht in die Zwischenablage kopiert.");
  } catch {
    // Zwischenablage im Add-in-Kontext evtl. gesperrt: Bericht zum manuellen Kopieren anzeigen.
    const area = document.getElementById("report") as HTMLTextAreaElement;
    area.hidden = false;
    area.value = text;
    area.select();
    setStatus("Zwischenablage nicht verfügbar – Bericht unten markiert, bitte mit Strg+C kopieren.");
  }
}

function setStatus(text: string): void {
  document.getElementById("status")!.textContent = text;
}

Office.onReady(() => {
  document.getElementById("run-all")!.onclick = () => void runAll();
  document.getElementById("copy")!.onclick = () => void copyReport();
  // Angeheftete Pane: bei Mailwechsel die mailbezogenen Ergebnisse verwerfen.
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
    for (const id of ["mail", "eml", "anhaenge", "upload"]) results.delete(id);
    state.eml = undefined;
    state.emlSource = undefined;
    state.shared = undefined;
    setStatus(Office.context.mailbox.item ? "Andere Mail ausgewählt – mailbezogene Prüfungen zurückgesetzt." : "Keine einzelne Mail ausgewählt.");
    render();
  });
  render();
});
