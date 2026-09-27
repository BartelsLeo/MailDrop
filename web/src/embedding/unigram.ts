// Unigram-Tokenizer (SentencePiece-Verfahren) für das deutschfähige Modell
// (gekürztes paraphrase-multilingual-MiniLM-L12-v2, siehe tools/model/).
//
// Vokabular-Format (Models/vocab.txt, erzeugt von tools/model/build_german_model.py):
// eine Zeile pro Token-ID, "<piece>\t<score>", Zeilennummer = ID. Die ersten Zeilen sind die
// Sondertokens <s>, <pad>, </s>, <unk>.
//
// Verfahren wie der Hugging-Face-Tokenizer von XLM-R: NFKC-Normalisierung, Leerraum zu einem
// Leerzeichen zusammenfassen, jedes Wort mit "▁" beginnen lassen, pro Wort per Viterbi die
// Zerlegung mit maximaler Summe der Log-Wahrscheinlichkeiten wählen; nicht abgedeckte Zeichen
// werden zu <unk> (aufeinanderfolgende zusammengefasst). Dieselbe Logik steckt in
// Services/UnigramTokenizer.vb (VSTO) und in tools/model/build_german_model.py (Referenz, prüft
// die Übereinstimmung mit dem Original-Tokenizer auf einem großen deutschen Textkorpus).
// Die Datei Models/testvectors.json aus dem Build-Skript sichert beide Umsetzungen ab.

export const WORD_PREFIX = "▁"; // "▁"

export class UnigramTokenizer {
  private readonly pieces = new Map<string, { id: number; score: number }>();
  private readonly maxPieceLength: number;
  private readonly unkScore: number;
  readonly bosId: number;
  readonly eosId: number;
  readonly unkId: number;

  constructor(
    vocabText: string,
    readonly maxLength = 128,
  ) {
    const lines = vocabText.split(/\r?\n/);
    let minScore = 0;
    let maxLen = 1;
    for (let id = 0; id < lines.length; id++) {
      const line = lines[id];
      if (!line) continue;
      const tab = line.lastIndexOf("\t");
      if (tab < 0) throw new Error(`Vokabular-Zeile ${id + 1} hat kein Tab – kein Unigram-Vokabular.`);
      const piece = line.slice(0, tab);
      const score = Number(line.slice(tab + 1));
      if (!this.pieces.has(piece)) this.pieces.set(piece, { id, score });
      if (!isSpecial(piece)) {
        minScore = Math.min(minScore, score);
        maxLen = Math.max(maxLen, Array.from(piece).length);
      }
    }
    this.maxPieceLength = maxLen;
    this.unkScore = minScore - 10; // wie Hugging Face tokenizers (Unigram)
    this.bosId = this.require("<s>");
    this.eosId = this.require("</s>");
    this.unkId = this.require("<unk>");
  }

  static looksLikeUnigramVocab(vocabText: string): boolean {
    const firstLine = vocabText.slice(0, 200).split(/\r?\n/)[0] ?? "";
    return firstLine.includes("\t");
  }

  encode(text: string): number[] {
    const ids: number[] = [this.bosId];
    const budget = this.maxLength - 2;
    outer: for (const word of this.preTokenize(text)) {
      for (const id of this.viterbi(word)) {
        if (ids.length - 1 >= budget) break outer;
        ids.push(id);
      }
    }
    ids.push(this.eosId);
    return ids;
  }

  /** Normalisierung + Wortzerlegung: jedes Wort beginnt mit "▁". */
  preTokenize(text: string): string[][] {
    const normalized = normalize(text);
    if (!normalized) return [];
    return normalized.split(" ").map((word) => Array.from(WORD_PREFIX + word));
  }

  private viterbi(chars: string[]): number[] {
    const n = chars.length;
    const bestScore = new Array<number>(n + 1).fill(-Infinity);
    const bestPrev = new Array<number>(n + 1).fill(-1);
    const bestId = new Array<number>(n + 1).fill(-1);
    bestScore[0] = 0;
    for (let start = 0; start < n; start++) {
      if (bestScore[start] === -Infinity) continue;
      let hasSingleChar = false;
      let piece = "";
      const maxEnd = Math.min(n, start + this.maxPieceLength);
      for (let end = start + 1; end <= maxEnd; end++) {
        piece += chars[end - 1];
        const entry = this.pieces.get(piece);
        if (!entry || isSpecial(piece)) continue;
        if (end === start + 1) hasSingleChar = true;
        const score = bestScore[start] + entry.score;
        if (score > bestScore[end]) {
          bestScore[end] = score;
          bestPrev[end] = start;
          bestId[end] = entry.id;
        }
      }
      if (!hasSingleChar) {
        const score = bestScore[start] + this.unkScore;
        if (score > bestScore[start + 1]) {
          bestScore[start + 1] = score;
          bestPrev[start + 1] = start;
          bestId[start + 1] = this.unkId;
        }
      }
    }
    const ids: number[] = [];
    for (let pos = n; pos > 0; pos = bestPrev[pos]) ids.push(bestId[pos]);
    ids.reverse();
    // Aufeinanderfolgende <unk> zusammenfassen (fuse_unk)
    return ids.filter((id, i) => !(id === this.unkId && i > 0 && ids[i - 1] === this.unkId));
  }

  private require(piece: string): number {
    const entry = this.pieces.get(piece);
    if (!entry) throw new Error(`Vokabular enthält ${piece} nicht.`);
    return entry.id;
  }
}

function isSpecial(piece: string): boolean {
  return piece === "<s>" || piece === "</s>" || piece === "<pad>" || piece === "<unk>" || piece === "<mask>";
}

export function normalize(text: string): string {
  let result = "";
  for (const ch of text.normalize("NFKC")) {
    const cp = ch.codePointAt(0)!;
    // Zuerst unsichtbare Zeichen verwerfen (JS zählt U+FEFF sonst zu \s), dann Leerraum vereinheitlichen.
    if (cp === 0xfeff || cp === 0x200b) continue;
    if (cp === 0x09 || cp === 0x0a || cp === 0x0d || /\s/u.test(ch)) result += " ";
    else if (cp < 0x20 || (cp >= 0x7f && cp < 0xa0)) continue;
    else result += ch;
  }
  return result.replace(/ {2,}/g, " ").trim();
}
