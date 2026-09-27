// BERT-WordPiece-Tokenizer für das Betreff-Embedding (uncased Vokabular, 30.522 Einträge in
// Models/vocab.txt). Entspricht der Standard-Vorverarbeitung eines "uncased" BERT-Modells:
// Kleinschreibung, Akzente/Umlaute entfernen (ä -> a), an Leerraum und Satzzeichen trennen,
// dann Greedy-Longest-Match gegen das Vokabular mit "##"-Fortsetzungsstücken.
//
// Bewusst NICHT 1:1 zur VSTO-Version: Microsoft.ML.Tokenizers.BertTokenizer entfernt Akzente
// standardmäßig nicht, deutsche Umlaut-Wörter werden dort vermutlich zu [UNK] (das Vokabular
// enthält kein "ä"). Da das VSTO-Add-in keine Nutzer und keinen Altverlauf hat, gibt es keine
// Embeddings, zu denen Kompatibilität nötig wäre - hier gilt das fachlich richtige Verfahren.

export interface TokenizerOptions {
  lowerCase: boolean;
  stripAccents: boolean;
  maxInputCharsPerWord: number;
  /** Maximale Sequenzlänge inkl. [CLS]/[SEP] (Modellgrenze). */
  maxLength: number;
}

const DEFAULT_OPTIONS: TokenizerOptions = {
  lowerCase: true,
  stripAccents: true,
  maxInputCharsPerWord: 100,
  maxLength: 256,
};

export class WordPieceTokenizer {
  private readonly vocab = new Map<string, number>();
  private readonly options: TokenizerOptions;
  readonly clsId: number;
  readonly sepId: number;
  readonly unkId: number;

  constructor(vocabText: string, options: Partial<TokenizerOptions> = {}) {
    this.options = { ...DEFAULT_OPTIONS, ...options };
    const lines = vocabText.split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
      const token = lines[i];
      if (token.length > 0 && !this.vocab.has(token)) this.vocab.set(token, i);
    }
    this.clsId = this.requireToken("[CLS]");
    this.sepId = this.requireToken("[SEP]");
    this.unkId = this.requireToken("[UNK]");
  }

  get vocabSize(): number {
    return this.vocab.size;
  }

  /** Token-IDs inkl. [CLS] am Anfang und [SEP] am Ende, auf maxLength gekürzt. */
  encode(text: string): number[] {
    const ids: number[] = [this.clsId];
    const budget = this.options.maxLength - 2;
    for (const word of this.basicTokenize(text)) {
      for (const id of this.wordPiece(word)) {
        if (ids.length - 1 >= budget) break;
        ids.push(id);
      }
    }
    ids.push(this.sepId);
    return ids;
  }

  /** Nur für Tests/Diagnose: Tokens als Text. */
  tokenize(text: string): string[] {
    const reverse = new Map<number, string>();
    for (const [token, id] of this.vocab) reverse.set(id, token);
    return this.encode(text).map((id) => reverse.get(id) ?? "[UNK]");
  }

  private requireToken(token: string): number {
    const id = this.vocab.get(token);
    if (id === undefined) throw new Error(`Vokabular enthält ${token} nicht.`);
    return id;
  }

  private basicTokenize(text: string): string[] {
    let cleaned = "";
    for (const ch of text) {
      const cp = ch.codePointAt(0)!;
      if (cp === 0 || cp === 0xfffd || isControl(ch)) continue;
      cleaned += isWhitespace(ch) ? " " : ch;
    }
    if (this.options.lowerCase) cleaned = cleaned.toLowerCase();
    if (this.options.stripAccents) {
      cleaned = cleaned.normalize("NFD").replace(/\p{Mn}/gu, "");
    }
    const words: string[] = [];
    for (const chunk of cleaned.split(" ")) {
      if (!chunk) continue;
      let current = "";
      for (const ch of chunk) {
        if (isPunctuation(ch) || isCjk(ch)) {
          if (current) words.push(current);
          words.push(ch);
          current = "";
        } else {
          current += ch;
        }
      }
      if (current) words.push(current);
    }
    return words;
  }

  private wordPiece(word: string): number[] {
    const chars = Array.from(word);
    if (chars.length > this.options.maxInputCharsPerWord) return [this.unkId];
    const ids: number[] = [];
    let start = 0;
    while (start < chars.length) {
      let end = chars.length;
      let found: number | undefined;
      while (start < end) {
        const piece = (start > 0 ? "##" : "") + chars.slice(start, end).join("");
        found = this.vocab.get(piece);
        if (found !== undefined) break;
        end--;
      }
      if (found === undefined) return [this.unkId];
      ids.push(found);
      start = end;
    }
    return ids;
  }
}

function isWhitespace(ch: string): boolean {
  return /\s/u.test(ch);
}

function isControl(ch: string): boolean {
  if (ch === "\t" || ch === "\n" || ch === "\r") return false;
  return /\p{Cc}|\p{Cf}/u.test(ch);
}

function isPunctuation(ch: string): boolean {
  const cp = ch.codePointAt(0)!;
  // Wie BERT: alle ASCII-Nicht-Buchstaben/Ziffern gelten als Satzzeichen, zusätzlich Unicode-P*.
  if ((cp >= 33 && cp <= 47) || (cp >= 58 && cp <= 64) || (cp >= 91 && cp <= 96) || (cp >= 123 && cp <= 126)) {
    return true;
  }
  return /\p{P}/u.test(ch);
}

function isCjk(ch: string): boolean {
  const cp = ch.codePointAt(0)!;
  return (
    (cp >= 0x4e00 && cp <= 0x9fff) ||
    (cp >= 0x3400 && cp <= 0x4dbf) ||
    (cp >= 0x20000 && cp <= 0x2a6df) ||
    (cp >= 0xf900 && cp <= 0xfaff) ||
    (cp >= 0x2f800 && cp <= 0x2fa1f)
  );
}
