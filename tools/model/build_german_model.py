"""Baut das deutschfähige Embedding-Modell für MailDrop (VSTO und Web-Add-in).

Ausgangsmodell: sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2 (384 Dimensionen,
mehrsprachig, gleiches Verfahren wie bisher: Mittelwert über alle Tokens + L2-Normalisierung).
Das Original hat ein Vokabular von 250.002 Tokens für ~50 Sprachen; allein die Token-Tabelle macht
~370 MB des ~470-MB-Modells aus. Dieses Skript

1. lädt Modell (ONNX-Fassung aus dem Hugging-Face-Repo) und Tokenizer in einer festen Revision,
2. zählt, welche Tokens auf einem deutschen (+ etwas englischen) Textkorpus tatsächlich vorkommen,
3. kürzt das Vokabular auf diese Tokens (plus alle einzelnen Zeichen der üblichen Schriften),
4. schneidet die Token-Tabelle des ONNX-Modells entsprechend zu und quantisiert es auf int8,
5. prüft
   - die eigene Tokenizer-Umsetzung (identisch zu web/src/embedding/unigram.ts und
     Services/UnigramTokenizer.vb) gegen den gekürzten Hugging-Face-Tokenizer,
   - die Embeddings des gekürzten, quantisierten Modells gegen das Originalmodell,
6. schreibt Models/model.onnx, Models/vocab.txt ("piece<TAB>score" je Zeile, Zeilennummer = ID),
   Models/testvectors.json (Token-IDs + Embeddings als Prüfvektoren für die TS-/VB-Umsetzung) und
   Models/MODEL_INFO.md.

Lauf: normalerweise über .github/workflows/build-german-model.yml (Hugging Face ist dort
erreichbar). Lokal: pip install -r tools/model/requirements.txt, dann
    python tools/model/build_german_model.py --out Models
Zum Testen ohne Internet: --model-dir <Ordner mit tokenizer.json und onnx/model.onnx> und
--corpus-file <Textdatei>.
"""

from __future__ import annotations

import argparse
import collections
import datetime
import hashlib
import json
import math
import sys
import tempfile
import unicodedata
from pathlib import Path

import numpy as np
import onnx
import onnxruntime as ort
from onnx import numpy_helper
from onnxruntime.quantization import QuantType, quantize_dynamic
from tokenizers import Tokenizer

REPO_ID = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"
# Feste Revision, damit ein erneuter Lauf dasselbe Modell liefert. "main" nur bewusst setzen.
DEFAULT_REVISION = "main"
MAX_LENGTH = 128  # max_seq_length des Originalmodells; gleicher Wert in unigram.ts/UnigramTokenizer.vb
SPECIAL_PIECES = {"<s>", "<pad>", "</s>", "<unk>", "<mask>"}
WORD_PREFIX = "▁"

# Zeichenbereiche, deren Einzelzeichen immer erhalten bleiben (auch wenn im Korpus nicht gesehen),
# damit seltene, aber in Betreffs plausible Zeichen nicht zu <unk> werden.
KEEP_CHAR_RANGES = [
    (0x20, 0x7E),      # ASCII
    (0xA0, 0x24F),     # Latin-1, Latin Extended-A/B
    (0x370, 0x3FF),    # Griechisch (Formelzeichen)
    (0x2000, 0x206F),  # Allgemeine Interpunktion (–, —, „, “, …)
    (0x20A0, 0x20CF),  # Währungszeichen (€)
    (0x2100, 0x214F),  # Buchstabenähnliche Symbole (№, ™)
    (0x2190, 0x21FF),  # Pfeile
    (0x2200, 0x22FF),  # Mathematische Operatoren
]

# Typische Betreffs/Wörter aus der Projektablage - gehen in den Korpus ein (Vokabular) und dienen als
# Prüftexte für Tokenizer-Gleichheit, Embedding-Qualität und die Prüfvektoren.
CURATED_TEXTS = [
    "Angebot für die Erweiterung des Bürogebäudes",
    "AW: Rechnung Nr. 2024-0815 – Zahlungserinnerung",
    "WG: Protokoll der Baubesprechung vom 12.03.2025",
    "Terminabstimmung Abnahme Heizungsanlage",
    "Nachtragsangebot Tiefbauarbeiten Los 3",
    "Bitte um Rückruf wegen Lieferverzögerung",
    "Einladung zur Projektbesprechung am Donnerstag",
    "Mängelanzeige Fenster 2. OG, Übergabe verschoben",
    "Baugenehmigung erteilt – nächste Schritte",
    "Kostenschätzung Straßenbau, Überarbeitung",
    "Schlussrechnung Elektroinstallation",
    "Änderung der Ausführungsplanung Statik",
    "Urlaubsvertretung vom 01.08. bis 15.08.",
    "Größenänderung der Fundamente (Bewehrungsplan)",
    "Stellungnahme zum Gutachten Schallschutz",
    "Zeichnungen Grundriss EG/1. OG anbei",
    "Auftragsbestätigung Bestellung 4711",
    "Frage zur Brandschutzkonzeption",
    "Re: Offer for the new office building",
    "Meeting minutes – kick-off",
    "Ölheizung, Fußbodenheizung und Wärmepumpe im Vergleich",
    "ß ẞ ä ö ü Ä Ö Ü é è à ç ñ",
    "E-Mail an Hr. Müller z. K.",
    "  Mehrfache   Leerzeichen\tund Tabs  ",
    "Preis: 1.234,56 € zzgl. MwSt. (19 %)",
    "Ticket #12345 [Dringend] {intern}",
    "ﬁnale Fassung ½ Seite ² ³",  # Kompatibilitätszeichen (NFKC)
    "Zero​Width﻿Space",
    "日本語のテキスト",  # nicht im gekürzten Vokabular -> <unk>
    "Emoji 🙂 im Betreff 👍",
    "",
    "a",
    "Donaudampfschifffahrtsgesellschaftskapitänsmütze",
]

# Paare für die Plausibilitätsprüfung (inhaltlich ähnlich vs. unähnlich), nur Bericht.
SIMILARITY_PAIRS = [
    ("Angebot für die Dachsanierung", "Kostenvoranschlag Dacharbeiten", True),
    ("Rechnung Heizungsinstallation", "Zahlungserinnerung Heizung", True),
    ("Protokoll Baubesprechung", "Niederschrift der Bausitzung", True),
    ("Angebot für die Dachsanierung", "Einladung zur Weihnachtsfeier", False),
    ("Rechnung Heizungsinstallation", "Urlaubsantrag Juli", False),
]


# --------------------------------------------------------------------------------------------
# Tokenizer - Zeile für Zeile wie web/src/embedding/unigram.ts und Services/UnigramTokenizer.vb.
# --------------------------------------------------------------------------------------------

# Leerraum wie JavaScript /\s/u bzw. .NET Char.IsWhiteSpace (ohne U+0085, s. u.).
_WHITESPACE = {0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x20, 0xA0, 0x1680, 0x2028, 0x2029, 0x202F, 0x205F, 0x3000}
_WHITESPACE.update(range(0x2000, 0x200B))


def normalize(text: str) -> str:
    result = []
    for ch in unicodedata.normalize("NFKC", text):
        cp = ord(ch)
        if cp in (0xFEFF, 0x200B):
            continue
        if cp in _WHITESPACE:
            result.append(" ")
        elif cp < 0x20 or 0x7F <= cp < 0xA0:
            continue
        else:
            result.append(ch)
    collapsed = "".join(result)
    while "  " in collapsed:
        collapsed = collapsed.replace("  ", " ")
    return collapsed.strip(" ")


class UnigramTokenizer:
    def __init__(self, vocab: list[tuple[str, float]], max_length: int = MAX_LENGTH):
        self.max_length = max_length
        self.pieces: dict[str, tuple[int, float]] = {}
        min_score = 0.0
        max_len = 1
        for i, (piece, score) in enumerate(vocab):
            if not piece:
                continue
            self.pieces.setdefault(piece, (i, score))
            if piece not in SPECIAL_PIECES:
                min_score = min(min_score, score)
                max_len = max(max_len, len(piece))
        self.max_piece_length = max_len
        self.unk_score = min_score - 10
        self.bos_id = self.pieces["<s>"][0]
        self.eos_id = self.pieces["</s>"][0]
        self.unk_id = self.pieces["<unk>"][0]

    def encode(self, text: str) -> list[int]:
        ids = [self.bos_id]
        budget = self.max_length - 2
        for word in self.pre_tokenize(text):
            for piece_id in self._viterbi(word):
                if len(ids) - 1 >= budget:
                    ids.append(self.eos_id)
                    return ids
                ids.append(piece_id)
        ids.append(self.eos_id)
        return ids

    def pre_tokenize(self, text: str) -> list[str]:
        normalized = normalize(text)
        if not normalized:
            return []
        return [WORD_PREFIX + word for word in normalized.split(" ")]

    def _viterbi(self, chars: str) -> list[int]:
        n = len(chars)
        best_score = [-math.inf] * (n + 1)
        best_prev = [-1] * (n + 1)
        best_id = [-1] * (n + 1)
        best_score[0] = 0.0
        for start in range(n):
            if best_score[start] == -math.inf:
                continue
            has_single_char = False
            for end in range(start + 1, min(n, start + self.max_piece_length) + 1):
                piece = chars[start:end]
                entry = self.pieces.get(piece)
                if entry is None or piece in SPECIAL_PIECES:
                    continue
                if end == start + 1:
                    has_single_char = True
                score = best_score[start] + entry[1]
                if score > best_score[end]:
                    best_score[end] = score
                    best_prev[end] = start
                    best_id[end] = entry[0]
            if not has_single_char:
                score = best_score[start] + self.unk_score
                if score > best_score[start + 1]:
                    best_score[start + 1] = score
                    best_prev[start + 1] = start
                    best_id[start + 1] = self.unk_id
        ids = []
        pos = n
        while pos > 0:
            ids.append(best_id[pos])
            pos = best_prev[pos]
        ids.reverse()
        return [x for i, x in enumerate(ids) if not (x == self.unk_id and i > 0 and ids[i - 1] == self.unk_id)]


# --------------------------------------------------------------------------------------------
# Schritte
# --------------------------------------------------------------------------------------------

def log(message: str) -> None:
    print(message, flush=True)


def fetch_model(revision: str) -> Path:
    from huggingface_hub import snapshot_download

    path = snapshot_download(
        REPO_ID,
        revision=revision,
        allow_patterns=["tokenizer.json", "config.json", "modules.json", "1_Pooling/config.json",
                        "sentence_bert_config.json", "onnx/model.onnx"],
    )
    return Path(path)


def load_corpus(args: argparse.Namespace) -> list[str]:
    texts = list(CURATED_TEXTS)
    if args.corpus_file:
        texts += [line for line in Path(args.corpus_file).read_text(encoding="utf-8").splitlines() if line.strip()]
        return texts
    from datasets import load_dataset

    for config, count in (("20231101.de", args.wiki_de), ("20231101.en", args.wiki_en)):
        log(f"Korpus: {count} Artikel aus wikimedia/wikipedia {config} ...")
        stream = load_dataset("wikimedia/wikipedia", config, split="train", streaming=True)
        for i, article in enumerate(stream):
            if i >= count:
                break
            texts.append(article["title"])
            # Anfang jedes Artikels: mehr verschiedene Themen statt weniger langer Artikel.
            texts += [line for line in article["text"][: args.chars_per_article].splitlines() if line.strip()]
    return texts


def count_tokens(tokenizer: Tokenizer, texts: list[str]) -> collections.Counter:
    counts: collections.Counter = collections.Counter()
    for start in range(0, len(texts), 1000):
        for encoding in tokenizer.encode_batch(texts[start:start + 1000], add_special_tokens=False):
            counts.update(encoding.ids)
    return counts


def choose_kept_ids(vocab: list[tuple[str, float]], counts: collections.Counter, corpus_chars: set[str],
                    max_vocab: int, min_count: int) -> list[int]:
    keep = {i for i, (piece, _) in enumerate(vocab) if piece in SPECIAL_PIECES}
    keep.add(next(i for i, (piece, _) in enumerate(vocab) if piece == WORD_PREFIX))
    # Einzelzeichen (mit und ohne Wortanfang): im Korpus gesehen oder in den festen Bereichen.
    for i, (piece, _) in enumerate(vocab):
        bare = piece[1:] if piece.startswith(WORD_PREFIX) else piece
        if len(bare) == 1 and (bare in corpus_chars or any(lo <= ord(bare) <= hi for lo, hi in KEEP_CHAR_RANGES)):
            keep.add(i)
    frequent = [i for i, c in counts.most_common() if c >= min_count and i not in keep]
    keep.update(frequent[: max(0, max_vocab - len(keep))])
    # Aufsteigend nach alter ID: Sondertokens <s>, <pad>, </s>, <unk> bleiben auf 0..3.
    return sorted(keep)


def pruned_hf_tokenizer(tokenizer_json: dict, kept_ids: list[int]) -> Tokenizer:
    data = json.loads(json.dumps(tokenizer_json))
    old_vocab = data["model"]["vocab"]
    data["model"]["vocab"] = [old_vocab[i] for i in kept_ids]
    remap = {old: new for new, old in enumerate(kept_ids)}
    data["model"]["unk_id"] = remap[data["model"]["unk_id"]]
    for token in data.get("added_tokens", []):
        token["id"] = remap[token["id"]]
    post = data.get("post_processor") or {}
    for special in (post.get("special_tokens") or {}).values():
        special["ids"] = [remap[i] for i in special["ids"]]
    return Tokenizer.from_str(json.dumps(data))


def prune_onnx(source: Path, target: Path, kept_ids: list[int], vocab_size: int) -> None:
    """Schneidet die Token-Tabelle auf kept_ids zu.

    Gesucht wird die Tabelle über den Gather-Knoten, der input_ids liest - nicht über die Zeilenzahl:
    das Modell hat mehr Zeilen als der Tokenizer Tokens (vocab_size 250037 vs. 250002; die übrigen
    Zeilen werden nie adressiert). Die Tabelle kann als Initializer oder als Constant-Knoten vorliegen.
    """
    model = onnx.load(str(source))
    graph = model.graph
    gathers = [n for n in graph.node if n.op_type == "Gather" and len(n.input) > 1 and n.input[1] == "input_ids"]
    initializers = {init.name: init for init in graph.initializer}
    if len(gathers) == 1:
        table_name = gathers[0].input[0]
    else:
        # Rückfall (z. B. input_ids erst über Cast/Reshape): größter 2D-Initializer mit genügend Zeilen.
        candidates = sorted((i for i in graph.initializer if len(i.dims) == 2 and i.dims[0] >= vocab_size),
                            key=lambda i: i.dims[0] * i.dims[1], reverse=True)
        if not candidates:
            raise SystemExit(f"Token-Tabelle nicht gefunden ({len(gathers)} Gather auf input_ids, kein Initializer >= {vocab_size} Zeilen).")
        table_name = candidates[0].name
    constants = {n.output[0]: n for n in graph.node if n.op_type == "Constant"}
    if table_name in initializers:
        holder = initializers[table_name]
        table = numpy_helper.to_array(holder)
    elif table_name in constants:
        holder = next(a for a in constants[table_name].attribute if a.name == "value").t
        table = numpy_helper.to_array(holder)
    else:
        raise SystemExit(f"Token-Tabelle {table_name!r} ist weder Initializer noch Constant.")
    if table.ndim != 2 or table.shape[0] < vocab_size:
        raise SystemExit(f"Token-Tabelle {table_name!r} hat Form {table.shape}, erwartet [>= {vocab_size}, H].")
    pruned = numpy_helper.from_array(np.ascontiguousarray(table[kept_ids]), holder.name)
    holder.CopyFrom(pruned)
    opset = max((o.version for o in model.opset_import if o.domain in ("", "ai.onnx")), default=0)
    if opset > 19:
        raise SystemExit(f"Opset {opset} ist zu neu für ONNX Runtime 1.16 (VSTO).")
    # ONNX Runtime 1.16 (VSTO) liest IR-Version <= 9; neuere Exporte tragen oft 10 ohne neue Features.
    model.ir_version = min(model.ir_version, 9)
    onnx.save(model, str(target))
    log(f"Token-Tabelle {table_name}: {table.shape} -> {tuple(pruned.dims)}")


class OnnxEmbedder:
    def __init__(self, path: Path):
        options = ort.SessionOptions()
        options.intra_op_num_threads = 1
        self.session = ort.InferenceSession(str(path), options, providers=["CPUExecutionProvider"])
        self.inputs = {i.name for i in self.session.get_inputs()}

    def embed(self, ids: list[int]) -> np.ndarray:
        arr = np.array([ids], dtype=np.int64)
        feeds = {"input_ids": arr}
        if "attention_mask" in self.inputs:
            feeds["attention_mask"] = np.ones_like(arr)
        if "token_type_ids" in self.inputs:
            feeds["token_type_ids"] = np.zeros_like(arr)
        hidden = self.session.run(None, feeds)[0][0]
        pooled = hidden.mean(axis=0)
        return pooled / (np.linalg.norm(pooled) or 1.0)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--out", default="Models")
    parser.add_argument("--revision", default=DEFAULT_REVISION)
    parser.add_argument("--model-dir", help="Lokaler Modellordner statt Download (tokenizer.json, onnx/model.onnx)")
    parser.add_argument("--corpus-file", help="Lokale Korpus-Textdatei statt Wikipedia")
    parser.add_argument("--wiki-de", type=int, default=40000)
    parser.add_argument("--wiki-en", type=int, default=5000)
    parser.add_argument("--chars-per-article", type=int, default=4000)
    parser.add_argument("--max-vocab", type=int, default=50000)
    parser.add_argument("--min-count", type=int, default=2)
    parser.add_argument("--agreement-sample", type=int, default=20000)
    parser.add_argument("--min-agreement", type=float, default=0.99)
    parser.add_argument("--min-mean-cosine", type=float, default=0.97)
    args = parser.parse_args()

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    model_dir = Path(args.model_dir) if args.model_dir else fetch_model(args.revision)
    source_onnx = model_dir / "onnx" / "model.onnx"
    if not source_onnx.exists():
        raise SystemExit(f"{source_onnx} fehlt - das Repo liefert in dieser Revision keine ONNX-Fassung.")

    tokenizer_json = json.loads((model_dir / "tokenizer.json").read_text(encoding="utf-8"))
    if tokenizer_json["model"]["type"] != "Unigram":
        raise SystemExit(f"Erwartet Unigram-Tokenizer, gefunden {tokenizer_json['model']['type']}.")
    full_vocab = [(p, float(s)) for p, s in tokenizer_json["model"]["vocab"]]
    original_tokenizer = Tokenizer.from_str(json.dumps(tokenizer_json))
    # tokenizer.json kann Padding/Kürzung vorgeben - fürs Zählen ganze Zeilen, ohne Auffüllen.
    original_tokenizer.no_padding()
    original_tokenizer.no_truncation()
    log(f"Originalvokabular: {len(full_vocab)} Tokens")

    texts = load_corpus(args)
    log(f"Korpus: {len(texts)} Zeilen, {sum(map(len, texts))} Zeichen")
    counts = count_tokens(original_tokenizer, texts)
    corpus_chars = {ch for text in texts for ch in normalize(text)}
    kept_ids = choose_kept_ids(full_vocab, counts, corpus_chars, args.max_vocab, args.min_count)
    bad = [full_vocab[i][0] for i in kept_ids if any(c in full_vocab[i][0] for c in "\t\r\n")]
    if bad:
        raise SystemExit(f"Tokens mit Tab/Zeilenumbruch passen nicht ins vocab.txt-Format: {bad[:5]}")
    vocab = [full_vocab[i] for i in kept_ids]
    covered = sum(c for i, c in counts.items() if i in set(kept_ids)) / max(1, sum(counts.values()))
    log(f"Gekürztes Vokabular: {len(vocab)} Tokens, deckt {covered:.2%} der Korpus-Tokens direkt ab")

    with tempfile.TemporaryDirectory() as tmp:
        pruned_path = Path(tmp) / "pruned.onnx"
        prune_onnx(source_onnx, pruned_path, kept_ids, len(full_vocab))
        model_path = out / "model.onnx"
        quantize_dynamic(str(pruned_path), str(model_path), weight_type=QuantType.QInt8,
                         op_types_to_quantize=["MatMul", "Gather"])
    log(f"Modell: {model_path.stat().st_size / 1e6:.1f} MB")

    vocab_path = out / "vocab.txt"
    vocab_path.write_text("".join(f"{p}\t{s!r}\n" for p, s in vocab), encoding="utf-8", newline="\n")

    # 1) Eigene Tokenizer-Umsetzung gegen den gekürzten Hugging-Face-Tokenizer.
    ours = UnigramTokenizer([(p, float(repr(s))) for p, s in vocab])
    reference = pruned_hf_tokenizer(tokenizer_json, kept_ids)
    reference.no_padding()
    reference.enable_truncation(MAX_LENGTH)
    sample = CURATED_TEXTS + texts[len(CURATED_TEXTS):][: args.agreement_sample]
    mismatches = []
    for text in sample:
        expected = reference.encode(text).ids
        if ours.encode(text) != expected:
            mismatches.append(text)
    agreement = 1 - len(mismatches) / len(sample)
    log(f"Tokenizer-Übereinstimmung mit Hugging Face: {agreement:.3%} von {len(sample)} Texten")
    for text in mismatches[:10]:
        log(f"  abweichend: {text[:80]!r}\n    HF:   {reference.encode(text).tokens[:20]}\n"
            f"    eig.: {[vocab[i][0] for i in ours.encode(text)][:20]}")

    # 2) Embedding-Qualität: gekürzt + quantisiert (eigener Tokenizer) gegen Original (HF-Tokenizer, fp32).
    original = OnnxEmbedder(source_onnx)
    final = OnnxEmbedder(model_path)
    original_tokenizer.enable_truncation(MAX_LENGTH)
    quality_texts = [t for t in CURATED_TEXTS if normalize(t)] + [p for pair in SIMILARITY_PAIRS for p in pair[:2]]
    cosines = [float(original.embed(original_tokenizer.encode(t).ids) @ final.embed(ours.encode(t))) for t in quality_texts]
    mean_cosine = float(np.mean(cosines))
    log(f"Embedding gekürzt+int8 vs. Original: Mittel {mean_cosine:.4f}, Minimum {min(cosines):.4f}")
    pair_lines = []
    for a, b, related in SIMILARITY_PAIRS:
        sim = float(final.embed(ours.encode(a)) @ final.embed(ours.encode(b)))
        pair_lines.append(f"| {a} | {b} | {'ähnlich' if related else 'unähnlich'} | {sim:.3f} |")
        log(f"  {sim:.3f} {'ähnlich  ' if related else 'unähnlich'} {a} / {b}")

    # 3) Prüfvektoren für web/test und tools/tokenizer-check.
    vectors = [{"text": t, "ids": ours.encode(t), "embedding": [round(float(x), 6) for x in final.embed(ours.encode(t))]}
               for t in CURATED_TEXTS + [p for pair in SIMILARITY_PAIRS for p in pair[:2]]]
    (out / "testvectors.json").write_text(json.dumps(vectors, ensure_ascii=False, indent=1), encoding="utf-8")

    info = f"""# Deutschfähiges Embedding-Modell (gebaut)

Erzeugt von `tools/model/build_german_model.py` am {datetime.datetime.now(datetime.timezone.utc):%Y-%m-%d %H:%M} UTC.
Nicht von Hand ändern - neu bauen über den Workflow `build-german-model`.

| | |
|---|---|
| Ausgangsmodell | `{REPO_ID}` |
| Revision | `{'lokal: ' + str(model_dir) if args.model_dir else args.revision + ' = ' + model_dir.name}` |
| Korpus | {'lokal: ' + args.corpus_file if args.corpus_file else f'Wikipedia de {args.wiki_de} + en {args.wiki_en} Artikel (je {args.chars_per_article} Zeichen), plus feste Prüftexte'} |
| Vokabular | {len(vocab)} von {len(full_vocab)} Tokens (Korpus-Abdeckung direkt {covered:.2%}) |
| Quantisierung | dynamisch int8 (MatMul, Gather), ONNX Runtime {ort.__version__} |
| Tokenizer-Übereinstimmung mit Hugging Face | {agreement:.3%} von {len(sample)} Texten |
| Embedding vs. Original (Kosinus) | Mittel {mean_cosine:.4f}, Minimum {min(cosines):.4f} |
| `model.onnx` | {model_path.stat().st_size / 1e6:.1f} MB, SHA-256 `{sha256(model_path)}` |
| `vocab.txt` | SHA-256 `{sha256(vocab_path)}` |

Plausibilität (Kosinus-Ähnlichkeit mit dem gebauten Modell):

| Text A | Text B | erwartet | Ähnlichkeit |
|---|---|---|---|
{chr(10).join(pair_lines)}
"""
    (out / "MODEL_INFO.md").write_text(info, encoding="utf-8")

    failed = []
    if agreement < args.min_agreement:
        failed.append(f"Tokenizer-Übereinstimmung {agreement:.3%} < {args.min_agreement:.0%}")
    if mean_cosine < args.min_mean_cosine:
        failed.append(f"mittlerer Kosinus {mean_cosine:.4f} < {args.min_mean_cosine}")
    if failed:
        log("FEHLGESCHLAGEN: " + "; ".join(failed))
        return 1
    log("Fertig.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
