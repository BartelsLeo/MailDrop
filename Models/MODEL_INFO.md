# Deutschfähiges Embedding-Modell (gebaut)

Erzeugt von `tools/model/build_german_model.py` am 2026-09-27 17:07 UTC.
Nicht von Hand ändern - neu bauen über den Workflow `build-german-model`.

| | |
|---|---|
| Ausgangsmodell | `sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2` |
| Revision | `main = e8f8c211226b894fcb81acc59f3b34ba3efd5f42` |
| Korpus | Wikipedia de 40000 + en 5000 Artikel (je 4000 Zeichen), plus feste Prüftexte |
| Vokabular | 49419 von 250002 Tokens (Korpus-Abdeckung direkt 99.98%) |
| Quantisierung | dynamisch 8 Bit (MatMul, Gather), int8-Gewichte, pro Kanal, reduce_range, ONNX Runtime 1.16.3 |
| Nur gekürzt (fp32) vs. Original (Kosinus) | Mittel 0.9918, Minimum 0.7038 |
| Tokenizer-Übereinstimmung mit Hugging Face | 99.985% von 20033 Texten |
| Embedding vs. Original (Kosinus) | Mittel 0.9844, Minimum 0.6865 |
| `model.onnx` | 41.3 MB, SHA-256 `06581603559956d2ea8fcb3038de591050c32eedc7b477f0988a7a7d54f95256` |
| `vocab.txt` | SHA-256 `2791f7fc9a7bf42acd403f8bab39254ec1c10e973ac6007cb25eaa00c60dd85b` |

Schwächste Prüftexte (gebautes Modell vs. Original):

- 0.6865: `Zero​Width﻿Space`
- 0.9546: `日本語のテキスト`
- 0.9790: `Angebot für die Dachsanierung`
- 0.9790: `Angebot für die Dachsanierung`
- 0.9833: `Nachtragsangebot Tiefbauarbeiten Los 3`

Plausibilität (Kosinus-Ähnlichkeit mit dem gebauten Modell):

| Text A | Text B | erwartet | Ähnlichkeit |
|---|---|---|---|
| Angebot für die Dachsanierung | Kostenvoranschlag Dacharbeiten | ähnlich | 0.757 |
| Rechnung Heizungsinstallation | Zahlungserinnerung Heizung | ähnlich | 0.866 |
| Protokoll Baubesprechung | Niederschrift der Bausitzung | ähnlich | 0.768 |
| Angebot für die Dachsanierung | Einladung zur Weihnachtsfeier | unähnlich | 0.314 |
| Rechnung Heizungsinstallation | Urlaubsantrag Juli | unähnlich | 0.328 |
