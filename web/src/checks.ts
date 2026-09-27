// Die Machbarkeitsprüfungen des Prototyps (docs/Web-Add-in-Umbau.md, Abschnitt 10).
// Jede Prüfung liefert ein Ergebnis mit Ampel und Details, damit der Nutzer den kompletten
// Bericht per "Bericht kopieren" zurückmelden kann.

import { getAuthMode, naaSupported } from "./auth";
import { isConfigured, scopes } from "./config";
import { encodePath, GraphError, graphGet, graphPut } from "./graph";
import {
  getAttachmentBase64,
  getEmlViaGraph,
  getEmlViaOffice,
  getSharedInfo,
  readCurrentMail,
  type SharedInfo,
} from "./mail";

export type Status = "ok" | "warn" | "fail";

export interface CheckResult {
  status: Status;
  summary: string;
  details: string[];
}

export interface UploadTarget {
  label: string;
  driveId: string;
  itemId: string;
  source: "verknuepfung" | "gefolgte-site";
}

/** Zustand zwischen den Prüfungen (z. B. .eml aus Prüfung 4 für den Upload in Prüfung 8). */
export const state: { eml?: Uint8Array; emlSource?: string; targets: UploadTarget[]; shared?: SharedInfo } = {
  targets: [],
};

const ms = (start: number) => `${Math.round(performance.now() - start)} ms`;
const kb = (bytes: number) => `${(bytes / 1024).toFixed(1)} KB`;

function describeError(error: unknown): string {
  if (error instanceof GraphError) return error.message;
  if (error instanceof Error) return `${error.name}: ${error.message}`;
  return String(error);
}

// 1 ---------------------------------------------------------------------------------------------
export async function checkUmgebung(): Promise<CheckResult> {
  const req = Office.context.requirements;
  const sets = ["1.5", "1.8", "1.13", "1.14", "1.15"].map(
    (v) => `Mailbox ${v}: ${req.isSetSupported("Mailbox", v) ? "ja" : "nein"}`,
  );
  const diag = Office.context.diagnostics;
  const has114 = req.isSetSupported("Mailbox", "1.14");
  const details = [
    `Host: ${diag.host}, Plattform: ${diag.platform}, Version: ${diag.version}`,
    ...sets,
    `NestedAppAuth 1.1: ${naaSupported() ? "ja" : "nein"}`,
    `App-Registrierung konfiguriert (VITE_CLIENT_ID): ${isConfigured() ? "ja" : "NEIN"}`,
    `Browser/WebView: ${navigator.userAgent}`,
  ];
  if (!isConfigured()) {
    return { status: "fail", summary: "Client-ID fehlt – .env.local anlegen (siehe web/README.md).", details };
  }
  return {
    status: has114 && naaSupported() ? "ok" : "warn",
    summary: has114
      ? "Umgebung erkannt."
      : "Mailbox 1.14 fehlt – .eml-Export nur über die Graph-Rückfallebene möglich.",
    details,
  };
}

// 2 ---------------------------------------------------------------------------------------------
export async function checkAnmeldung(): Promise<CheckResult> {
  const start = performance.now();
  try {
    const mode = await getAuthMode();
    const me = await graphGet<{ displayName: string; userPrincipalName: string }>("/me", scopes.base);
    return {
      status: mode === "naa" ? "ok" : "warn",
      summary: `Angemeldet als ${me.displayName} (${mode === "naa" ? "NAA" : "Popup-Rückfallebene"}).`,
      details: [`UPN: ${me.userPrincipalName}`, `Verfahren: ${mode}`, `Dauer: ${ms(start)}`],
    };
  } catch (error) {
    return { status: "fail", summary: "Anmeldung fehlgeschlagen.", details: [describeError(error)] };
  }
}

// 3 ---------------------------------------------------------------------------------------------
export async function checkMail(): Promise<CheckResult> {
  try {
    const mail = readCurrentMail();
    state.shared = await getSharedInfo();
    return {
      status: "ok",
      summary: `Mail gelesen: „${mail.subject}“`,
      details: [
        `Absender: ${mail.from} <${mail.fromAddress}>`,
        `An: ${mail.to}`,
        `Datum: ${mail.received?.toISOString() ?? "-"}`,
        `Anhänge: ${mail.attachments.length} (davon inline: ${mail.attachments.filter((a) => a.isInline).length})`,
        state.shared.isShared
          ? `Freigegebenes Postfach/Stellvertretung: ja – Besitzer ${state.shared.owner}`
          : "Freigegebenes Postfach/Stellvertretung: nein (eigenes Postfach)",
      ],
    };
  } catch (error) {
    return { status: "fail", summary: "Mail konnte nicht gelesen werden.", details: [describeError(error)] };
  }
}

// 4 ---------------------------------------------------------------------------------------------
export async function checkEml(): Promise<CheckResult> {
  const details: string[] = [];
  state.shared ??= await getSharedInfo().catch(() => ({ isShared: false }));

  let start = performance.now();
  try {
    const eml = await getEmlViaOffice();
    details.push(`Office.js getAsFileAsync: ok, ${kb(eml.byteLength)}, ${ms(start)}`);
    details.push(...describeEml(eml));
    state.eml = eml;
    state.emlSource = "Office.js";
  } catch (error) {
    details.push(`Office.js getAsFileAsync: FEHLER – ${describeError(error)}`);
  }

  start = performance.now();
  try {
    const eml = await getEmlViaGraph(state.shared);
    details.push(`Graph-Rückfallebene ($value): ok, ${kb(eml.byteLength)}, ${ms(start)}`);
    if (!state.eml) {
      details.push(...describeEml(eml));
      state.eml = eml;
      state.emlSource = "Graph";
    }
  } catch (error) {
    details.push(`Graph-Rückfallebene ($value): FEHLER – ${describeError(error)}`);
  }

  if (state.emlSource === "Office.js") return { status: "ok", summary: ".eml über Office.js erhalten.", details };
  if (state.emlSource === "Graph") {
    return { status: "warn", summary: ".eml nur über die Graph-Rückfallebene erhalten.", details };
  }
  return { status: "fail", summary: "Kein Weg zum .eml-Export funktioniert.", details };
}

function describeEml(eml: Uint8Array): string[] {
  const head = new TextDecoder("latin1").decode(eml.slice(0, 64 * 1024));
  const header = (name: string) => new RegExp(`^${name}:`, "im").test(head);
  const encrypted = /application\/(x-)?pkcs7-mime|smime\.p7m/i.test(head);
  const rights = /microsoft-rights|msip_label|rpmsg/i.test(head);
  return [
    `  Kopfzeilen vorhanden – From: ${header("From")}, Subject: ${header("Subject")}, Date: ${header("Date")}, MIME-Version: ${header("MIME-Version")}`,
    `  Hinweise auf S/MIME-Verschlüsselung: ${encrypted ? "ja" : "nein"}, auf Vertraulichkeitsbezeichnung/Rechteverwaltung: ${rights ? "ja" : "nein"}`,
  ];
}

// 5 ---------------------------------------------------------------------------------------------
export async function checkAnhaenge(): Promise<CheckResult> {
  try {
    const attachments = readCurrentMail().attachments.filter((a) => !a.isInline);
    if (attachments.length === 0) {
      return { status: "warn", summary: "Mail ohne (nicht-inline) Anhänge – für diese Prüfung eine Mail mit Anhang wählen.", details: [] };
    }
    const details: string[] = [];
    let failed = 0;
    // Bewusst nacheinander: neues Outlook/Web erlaubt max. 3 gleichzeitige asynchrone Aufrufe.
    for (const a of attachments) {
      const start = performance.now();
      try {
        const content = await getAttachmentBase64(a.id);
        details.push(`${a.name} (${kb(a.size)}): ok, Format ${content.format}, ${ms(start)}`);
      } catch (error) {
        failed++;
        details.push(`${a.name} (${kb(a.size)}): FEHLER – ${describeError(error)}`);
      }
    }
    return {
      status: failed === 0 ? "ok" : "warn",
      summary: failed === 0 ? `Alle ${attachments.length} Anhänge abrufbar.` : `${failed} von ${attachments.length} Anhängen nicht abrufbar.`,
      details,
    };
  } catch (error) {
    return { status: "fail", summary: "Anhänge konnten nicht gelesen werden.", details: [describeError(error)] };
  }
}

// 6 ---------------------------------------------------------------------------------------------
interface DriveItem {
  id: string;
  name: string;
  folder?: unknown;
  remoteItem?: { id: string; name?: string; parentReference?: { driveId?: string; driveType?: string }; webUrl?: string };
}

export async function checkVerknuepfungen(): Promise<CheckResult> {
  const details: string[] = [];
  state.targets = [];
  try {
    const children = await graphGet<{ value: DriveItem[] }>("/me/drive/root/children?$top=999", scopes.files);
    const shortcuts = children.value.filter((i) => i.remoteItem);
    details.push(`Einträge im OneDrive-Stammordner: ${children.value.length}, davon Verknüpfungen (remoteItem): ${shortcuts.length}`);
    for (const s of shortcuts) {
      const driveId = s.remoteItem!.parentReference?.driveId;
      details.push(`  • ${s.name} (Laufwerkstyp ${s.remoteItem!.parentReference?.driveType ?? "?"})`);
      if (driveId) state.targets.push({ label: `${s.name} (Verknüpfung)`, driveId, itemId: s.remoteItem!.id, source: "verknuepfung" });
    }
  } catch (error) {
    details.push(`OneDrive-Stammordner: FEHLER – ${describeError(error)}`);
  }

  // Rückfallebene: vom Nutzer gefolgte Sites (Stern in SharePoint)
  try {
    const followed = await graphGet<{ value: { id: string; displayName: string }[] }>("/me/followedSites", scopes.sites);
    details.push(`Gefolgte Sites (Rückfallebene): ${followed.value.length}`);
    for (const site of followed.value.slice(0, 20)) {
      try {
        const drive = await graphGet<{ id: string; name: string }>(`/sites/${site.id}/drive`, scopes.files);
        const root = await graphGet<{ id: string }>(`/drives/${drive.id}/root`, scopes.files);
        state.targets.push({ label: `${site.displayName} / ${drive.name} (gefolgt)`, driveId: drive.id, itemId: root.id, source: "gefolgte-site" });
        details.push(`  • ${site.displayName} – Standardbibliothek „${drive.name}“`);
      } catch (error) {
        details.push(`  • ${site.displayName}: Bibliothek nicht lesbar – ${describeError(error)}`);
      }
    }
  } catch (error) {
    details.push(`Gefolgte Sites: FEHLER – ${describeError(error)}`);
  }

  const shortcutCount = state.targets.filter((t) => t.source === "verknuepfung").length;
  if (shortcutCount > 0) {
    return { status: "ok", summary: `${shortcutCount} Bibliotheks-Verknüpfung(en) über Graph gefunden.`, details };
  }
  return {
    status: state.targets.length > 0 ? "warn" : "fail",
    summary:
      "Keine Verknüpfungen über Graph sichtbar" +
      (state.targets.length > 0 ? " – Rückfallebene über gefolgte Sites funktioniert." : " und keine gefolgten Sites.") +
      " (Vorher prüfen: Ist mindestens eine Projektbibliothek per „Verknüpfung zu Meine Dateien hinzufügen“ eingebunden?)",
    details,
  };
}

// 7 ---------------------------------------------------------------------------------------------
export async function checkAppOrdner(): Promise<CheckResult> {
  const details: string[] = [];
  const payload = JSON.stringify({ test: "MailDrop-Prototyp", at: new Date().toISOString() });
  let appFolderOk = false;
  try {
    const approot = await graphGet<{ name: string; webUrl: string }>("/me/drive/special/approot", scopes.appFolder);
    await graphPut(`/me/drive/special/approot:/prototyp-test.json:/content`, scopes.appFolder, payload, "application/json");
    appFolderOk = true;
    details.push(`App-Ordner (Files.ReadWrite.AppFolder): ok – „${approot.name}“, Testdatei geschrieben`);
  } catch (error) {
    details.push(`App-Ordner (Files.ReadWrite.AppFolder): FEHLER – ${describeError(error)}`);
  }
  let fallbackOk = false;
  try {
    await graphPut(`/me/drive/root:/Apps/MailDrop/prototyp-test.json:/content`, scopes.files, payload, "application/json");
    fallbackOk = true;
    details.push("Rückfallebene Ordner Apps/MailDrop (Files.ReadWrite.All): ok, Testdatei geschrieben");
  } catch (error) {
    details.push(`Rückfallebene Ordner Apps/MailDrop: FEHLER – ${describeError(error)}`);
  }
  if (appFolderOk) return { status: "ok", summary: "App-Ordner funktioniert.", details };
  if (fallbackOk) return { status: "warn", summary: "App-Ordner nicht verfügbar – Rückfallebene funktioniert.", details };
  return { status: "fail", summary: "Kein Speicherort für den Verlauf beschreibbar.", details };
}

// 8 ---------------------------------------------------------------------------------------------
export async function checkUpload(target: UploadTarget | undefined): Promise<CheckResult> {
  if (!state.eml) return { status: "fail", summary: "Erst Prüfung 4 (.eml) ausführen.", details: [] };
  if (!target) return { status: "fail", summary: "Erst Prüfung 6 ausführen und eine Bibliothek wählen.", details: [] };
  const subject = safeName(readCurrentMail().subject || "ohne Betreff");
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const relative = `MailDrop-Prototyp-Test/${stamp}_${subject}.eml`;
  const start = performance.now();
  try {
    const created = await graphPut<{ webUrl: string; size: number }>(
      `/drives/${target.driveId}/items/${target.itemId}:/${encodePath(relative)}:/content`,
      scopes.files,
      state.eml as unknown as BodyInit,
      "message/rfc822",
    );
    return {
      status: "ok",
      summary: `.eml in „${target.label}“ abgelegt.`,
      details: [`Datei: ${relative}`, `Größe: ${kb(created.size)}, Dauer: ${ms(start)}`, `Adresse: ${created.webUrl}`, `Quelle der .eml: ${state.emlSource}`],
    };
  } catch (error) {
    return { status: "fail", summary: "Upload fehlgeschlagen.", details: [describeError(error)] };
  }
}

function safeName(text: string): string {
  // SharePoint-Regeln: * " : < > ? / \ | verboten, keine führenden/abschließenden Leerzeichen/Punkte.
  return text.replace(/[*"<>?:/\\|#%]/g, "_").replace(/\s+/g, " ").trim().replace(/\.+$/, "").slice(0, 120) || "Mail";
}

// 9 ---------------------------------------------------------------------------------------------
export async function checkModell(): Promise<CheckResult> {
  const details: string[] = [];
  const memory = () => {
    const m = (performance as Performance & { memory?: { usedJSHeapSize: number } }).memory;
    return m ? `${(m.usedJSHeapSize / 1048576).toFixed(0)} MB JS-Heap (ohne WASM-Speicher des Modells)` : "Speicher nicht messbar";
  };
  try {
    details.push(`Vorher: ${memory()}`);
    let start = performance.now();
    const [modelResponse, vocabResponse] = await Promise.all([fetch("models/model.onnx"), fetch("models/vocab.txt")]);
    if (!modelResponse.ok || !vocabResponse.ok) throw new Error(`Modelldateien nicht gefunden (HTTP ${modelResponse.status}/${vocabResponse.status}).`);
    const modelBytes = new Uint8Array(await modelResponse.arrayBuffer());
    const vocab = await vocabResponse.text();
    details.push(`Download: ${(modelBytes.byteLength / 1048576).toFixed(1)} MB in ${ms(start)}`);

    const ort = await import("onnxruntime-web/wasm");
    const { Embedder, cosine } = await import("./embedding/embedder");
    start = performance.now();
    const embedder = await Embedder.create(ort, modelBytes, vocab);
    details.push(`Modell geladen in ${Math.round(embedder.loadInfo.loadMs)} ms, danach ${memory()}`);

    start = performance.now();
    const subject = (() => {
      try {
        return readCurrentMail().subject;
      } catch {
        return "";
      }
    })();
    const a = await embedder.embed("Invoice for the office renovation project");
    const b = await embedder.embed("Bill for renovating the office");
    const c = await embedder.embed("Weekend football match results");
    if (subject) await embedder.embed(subject);
    details.push(`${subject ? 4 : 3} Embeddings in ${ms(start)}`);
    details.push(`Ähnlichkeit verwandt: ${cosine(a, b).toFixed(3)}, unverwandt: ${cosine(a, c).toFixed(3)}`);
    const plausible = cosine(a, b) > cosine(a, c) + 0.2;
    return {
      status: plausible ? "ok" : "warn",
      summary: plausible ? "Modell läuft im Aufgabenbereich." : "Modell läuft, Ähnlichkeiten unplausibel.",
      details,
    };
  } catch (error) {
    details.push(describeError(error));
    return { status: "fail", summary: "Modell konnte nicht geladen/ausgeführt werden.", details };
  }
}
