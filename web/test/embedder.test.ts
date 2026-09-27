// Lädt das echte Modell mit ONNX Runtime Web (WASM) in Node - prüft, dass Verfahren und Modell
// im Browser-Laufzeitsystem funktionieren und semantisch plausible Ähnlichkeiten liefern.
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import * as ort from "onnxruntime-web";
import { Embedder, cosine } from "../src/embedding/embedder";

const modelsDir = join(__dirname, "..", "..", "Models");
const modelPath = join(modelsDir, "model.onnx");

describe.skipIf(!existsSync(modelPath))("Embedder (echtes Modell)", () => {
  it("liefert normalisierte 384-dim Vektoren mit plausiblen Ähnlichkeiten", async () => {
    const embedder = await Embedder.create(
      ort,
      new Uint8Array(readFileSync(modelPath)),
      readFileSync(join(modelsDir, "vocab.txt"), "utf8"),
    );
    const a = await embedder.embed("Invoice for the office renovation project");
    const b = await embedder.embed("Bill for renovating the office");
    const c = await embedder.embed("Weekend football match results");
    expect(a.length).toBe(384);
    const norm = Math.sqrt(a.reduce((s, v) => s + v * v, 0));
    expect(norm).toBeCloseTo(1, 4);
    expect(cosine(a, b)).toBeGreaterThan(cosine(a, c) + 0.2);
    console.log(
      `Ladezeit ${embedder.loadInfo.loadMs.toFixed(0)} ms, sim(a,b)=${cosine(a, b).toFixed(3)}, sim(a,c)=${cosine(a, c).toFixed(3)}`,
    );
  }, 120_000);
});
