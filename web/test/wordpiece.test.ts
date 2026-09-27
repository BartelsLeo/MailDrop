import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { WordPieceTokenizer } from "../src/embedding/wordpiece";

const vocab = readFileSync(join(__dirname, "..", "..", "Models", "vocab.txt"), "utf8");
const tok = new WordPieceTokenizer(vocab);

describe("WordPieceTokenizer", () => {
  it("liest das komplette Vokabular", () => {
    expect(tok.vocabSize).toBe(30522);
  });

  it("umrahmt mit [CLS]/[SEP] und schreibt klein", () => {
    expect(tok.tokenize("Hello World")).toEqual(["[CLS]", "hello", "world", "[SEP]"]);
  });

  it("trennt Satzzeichen ab", () => {
    expect(tok.tokenize("RE: Offer, Version 2!")).toEqual(
      ["[CLS]", "re", ":", "offer", ",", "version", "2", "!", "[SEP]"],
    );
  });

  it("entfernt Umlaute/Akzente statt [UNK] zu erzeugen", () => {
    const tokens = tok.tokenize("Prüfung Änderung");
    expect(tokens).not.toContain("[UNK]");
    expect(tokens.join(" ")).toContain("pr");
  });

  it("zerlegt unbekannte Wörter in ##-Stücke", () => {
    const tokens = tok.tokenize("Bauantrag");
    expect(tokens.length).toBeGreaterThan(3);
    expect(tokens.slice(2, -1).every((t) => t.startsWith("##"))).toBe(true);
  });

  it("kürzt auf maxLength", () => {
    const short = new WordPieceTokenizer(vocab, { maxLength: 8 });
    const ids = short.encode("a b c d e f g h i j k l");
    expect(ids.length).toBe(8);
    expect(ids[0]).toBe(short.clsId);
    expect(ids[7]).toBe(short.sepId);
  });
});
