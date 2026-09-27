// Betreff-Embedding im Browser: gleiches Verfahren wie Services/EmbeddingService.vb
// (Modell Models/model.onnx, 384 Dimensionen, Mittelwert über alle Tokens, L2-Normalisierung),
// aber mit ONNX Runtime Web (WASM, ein Thread) statt ONNX Runtime für .NET.

import type * as OrtNamespace from "onnxruntime-web";
import { WordPieceTokenizer } from "./wordpiece";

export interface EmbedderLoadInfo {
  modelBytes: number;
  loadMs: number;
}

export class Embedder {
  private constructor(
    private readonly ort: typeof OrtNamespace,
    private readonly session: OrtNamespace.InferenceSession,
    private readonly tokenizer: WordPieceTokenizer,
    readonly loadInfo: EmbedderLoadInfo,
  ) {}

  static async create(
    ort: typeof OrtNamespace,
    modelBytes: Uint8Array,
    vocabText: string,
  ): Promise<Embedder> {
    const started = performance.now();
    ort.env.wasm.numThreads = 1;
    const session = await ort.InferenceSession.create(modelBytes, {
      executionProviders: ["wasm"],
      graphOptimizationLevel: "all",
    });
    const tokenizer = new WordPieceTokenizer(vocabText);
    return new Embedder(ort, session, tokenizer, {
      modelBytes: modelBytes.byteLength,
      loadMs: performance.now() - started,
    });
  }

  async embed(text: string): Promise<Float32Array> {
    const ids = this.tokenizer.encode(text);
    const n = ids.length;
    const toTensor = (values: number[]) =>
      new this.ort.Tensor("int64", BigInt64Array.from(values.map((v) => BigInt(v))), [1, n]);

    const feeds: Record<string, OrtNamespace.Tensor> = {};
    const available = new Set(this.session.inputNames);
    if (available.has("input_ids")) feeds.input_ids = toTensor(ids);
    if (available.has("attention_mask")) feeds.attention_mask = toTensor(ids.map(() => 1));
    if (available.has("token_type_ids")) feeds.token_type_ids = toTensor(ids.map(() => 0));

    const results = await this.session.run(feeds);
    const output = results[this.session.outputNames[0]];
    const data = output.data as Float32Array;
    const hidden = output.dims[2];

    const embedding = new Float32Array(hidden);
    for (let t = 0; t < n; t++) {
      for (let d = 0; d < hidden; d++) embedding[d] += data[t * hidden + d];
    }
    let norm = 0;
    for (let d = 0; d < hidden; d++) {
      embedding[d] /= n;
      norm += embedding[d] * embedding[d];
    }
    norm = Math.sqrt(norm) || 1;
    for (let d = 0; d < hidden; d++) embedding[d] /= norm;
    return embedding;
  }
}

export function cosine(a: Float32Array, b: Float32Array): number {
  let dot = 0;
  for (let i = 0; i < a.length; i++) dot += a[i] * b[i];
  return dot; // beide Vektoren sind normalisiert
}
