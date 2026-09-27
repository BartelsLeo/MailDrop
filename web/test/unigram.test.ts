import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { normalize, UnigramTokenizer } from "../src/embedding/unigram";

// Kleines künstliches Vokabular: prüft die Viterbi-Logik unabhängig vom echten Modell.
const vocab = [
  ["<s>", 0], ["<pad>", 0], ["</s>", 0], ["<unk>", 0],
  ["▁", -2], ["▁an", -3], ["▁ange", -4], ["bot", -3], ["▁angebot", -5],
  ["a", -6], ["n", -6], ["g", -6], ["e", -6], ["b", -6], ["o", -6], ["t", -6],
  ["▁pr", -4], ["ü", -5], ["fung", -4], ["▁prüfung", -6.5],
].map(([p, s]) => `${p}\t${s}`).join("\n");

describe("UnigramTokenizer (künstliches Vokabular)", () => {
  const tok = new UnigramTokenizer(vocab);
  const pieceOf = (id: number) => vocab.split("\n")[id].split("\t")[0];

  it("erkennt das Vokabularformat", () => {
    expect(UnigramTokenizer.looksLikeUnigramVocab(vocab)).toBe(true);
    expect(UnigramTokenizer.looksLikeUnigramVocab("[PAD]\n[unused0]")).toBe(false);
  });

  it("wählt die Zerlegung mit der höchsten Gesamtwahrscheinlichkeit", () => {
    // "▁angebot" (-5) schlägt "▁ange"+"bot" (-7) und "▁an"+...
    expect(tok.encode("angebot").map(pieceOf)).toEqual(["<s>", "▁angebot", "</s>"]);
  });

  it("unterscheidet Groß-/Kleinschreibung (Modell ist cased)", () => {
    expect(tok.encode("Angebot").map(pieceOf)).not.toContain("▁angebot");
  });

  it("zerlegt, wenn das ganze Wort teurer ist als die Teile", () => {
    // "▁prüfung" (-6.5) vs "▁pr"+"ü"+"fung" (-13) -> ganzes Wort
    expect(tok.encode("prüfung").map(pieceOf)).toEqual(["<s>", "▁prüfung", "</s>"]);
  });

  it("fasst nicht abgedeckte Zeichen zu einem <unk> zusammen", () => {
    expect(tok.encode("an xyz").map(pieceOf)).toEqual(["<s>", "▁an", "▁", "<unk>", "</s>"]);
  });

  it("kürzt auf maxLength inkl. <s>/</s>", () => {
    const short = new UnigramTokenizer(vocab, 4);
    expect(short.encode("angebot angebot angebot").length).toBe(4);
  });

  it("normalisiert Leerraum und Kompatibilitätszeichen", () => {
    expect(normalize("  a\t\tb c﻿ ")).toBe("a b c");
    expect(normalize("ﬁ")).toBe("fi"); // NFKC
  });
});

// Echte Prüfvektoren aus tools/model/build_german_model.py (nur vorhanden, sobald das
// deutschfähige Modell gebaut wurde): Token-IDs müssen exakt denen des Original-Tokenizers
// entsprechen.
// MAILDROP_MODELS_DIR: anderer Modellordner, z. B. die Ausgabe von tools/model/build_german_model.py.
const modelsDir = process.env.MAILDROP_MODELS_DIR ?? join(__dirname, "..", "..", "Models");
const vectorsPath = join(modelsDir, "testvectors.json");
describe.skipIf(!existsSync(vectorsPath))("UnigramTokenizer (Prüfvektoren des echten Modells)", () => {
  it("erzeugt dieselben Token-IDs wie der Hugging-Face-Tokenizer", () => {
    const vectors = JSON.parse(readFileSync(vectorsPath, "utf8")) as { text: string; ids: number[] }[];
    const tok = new UnigramTokenizer(readFileSync(join(modelsDir, "vocab.txt"), "utf8"));
    for (const v of vectors) expect(tok.encode(v.text), v.text).toEqual(v.ids);
  });
});
