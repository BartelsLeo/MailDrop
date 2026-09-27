// Zugriff auf die aktuell angezeigte Mail über Office.js, mit Graph als Rückfallebene.

import { graphGetBytes } from "./graph";
import { scopes } from "./config";

export interface MailInfo {
  subject: string;
  from: string;
  fromAddress: string;
  to: string;
  received: Date | undefined;
  attachments: { id: string; name: string; size: number; isInline: boolean }[];
  itemId: string;
}

export interface SharedInfo {
  isShared: boolean;
  owner?: string;
  targetMailbox?: string;
}

function item(): Office.MessageRead {
  const current = Office.context.mailbox.item as Office.MessageRead | null;
  if (!current) {
    throw new Error("Keine Mail ausgewählt (bei Mehrfachauswahl oder leerer Auswahl ist item null).");
  }
  return current;
}

function promisify<T>(call: (cb: (result: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise((resolve, reject) => {
    call((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(`${result.error?.name ?? "Fehler"}: ${result.error?.message ?? "unbekannt"}`));
    });
  });
}

export function readCurrentMail(): MailInfo {
  const mail = item();
  return {
    subject: mail.subject ?? "",
    from: mail.from?.displayName ?? "",
    fromAddress: mail.from?.emailAddress ?? "",
    to: (mail.to ?? []).map((r) => r.displayName || r.emailAddress).join("; "),
    received: mail.dateTimeCreated,
    attachments: (mail.attachments ?? []).map((a) => ({
      id: a.id,
      name: a.name,
      size: a.size,
      isInline: a.isInline,
    })),
    itemId: mail.itemId,
  };
}

export async function getSharedInfo(): Promise<SharedInfo> {
  const mail = item() as Office.MessageRead & {
    getSharedPropertiesAsync?: (cb: (r: Office.AsyncResult<Office.SharedProperties>) => void) => void;
  };
  if (!Office.context.requirements.isSetSupported("Mailbox", "1.8") || !mail.getSharedPropertiesAsync) {
    return { isShared: false };
  }
  try {
    const props = await promisify<Office.SharedProperties>((cb) => mail.getSharedPropertiesAsync!(cb));
    return { isShared: true, owner: props.owner, targetMailbox: props.targetMailbox };
  } catch {
    // Im eigenen Postfach liefert die Methode je nach Client einen Fehler statt "nicht geteilt".
    return { isShared: false };
  }
}

/** .eml (MIME) über Office.js – Requirement Set Mailbox 1.14. */
export async function getEmlViaOffice(): Promise<Uint8Array> {
  const mail = item() as Office.MessageRead & {
    getAsFileAsync?: (cb: (r: Office.AsyncResult<string>) => void) => void;
  };
  if (!mail.getAsFileAsync) throw new Error("getAsFileAsync ist in diesem Outlook nicht vorhanden.");
  const base64 = await promisify<string>((cb) => mail.getAsFileAsync!(cb));
  return base64ToBytes(base64);
}

/** .eml (MIME) über Graph – Rückfallebene, braucht Mail.Read bzw. Mail.Read.Shared. */
export async function getEmlViaGraph(shared: SharedInfo): Promise<Uint8Array> {
  const restId = Office.context.mailbox.convertToRestId(item().itemId, Office.MailboxEnums.RestVersion.v2_0);
  const mailbox = shared.isShared && shared.owner ? `/users/${encodeURIComponent(shared.owner)}` : "/me";
  const needed = shared.isShared ? scopes.mailShared : scopes.mail;
  return graphGetBytes(`${mailbox}/messages/${encodeURIComponent(restId)}/$value`, needed);
}

/** Inhalt eines Anhangs über Office.js (neues Outlook/Web: bis ca. 25 MB). */
export async function getAttachmentBase64(attachmentId: string): Promise<{ format: string; content: string }> {
  const mail = item();
  const result = await promisify<Office.AttachmentContent>((cb) =>
    mail.getAttachmentContentAsync(attachmentId, cb),
  );
  return { format: result.format, content: result.content };
}

// Die .eml bleibt als Bytes erhalten (MIME-Teile können 8-Bit-Inhalte in beliebigen
// Zeichensätzen enthalten) - ein Umweg über einen JS-String würde sie beim Hochladen verfälschen.
export function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}
