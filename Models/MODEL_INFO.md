# Deutschfähiges Embedding-Modell (gebaut)

Erzeugt von `tools/model/build_german_model.py` am 2026-09-27 16:54 UTC.
Nicht von Hand ändern - neu bauen über den Workflow `build-german-model`.

| | |
|---|---|
| Ausgangsmodell | `sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2` |
| Revision | `main = e8f8c211226b894fcb81acc59f3b34ba3efd5f42` |
| Korpus | Wikipedia de 40000 + en 5000 Artikel (je 4000 Zeichen), plus feste Prüftexte |
| Vokabular | 49359 von 250002 Tokens (Korpus-Abdeckung direkt 99.98%) |
| Quantisierung | dynamisch int8 (MatMul, Gather), ONNX Runtime 1.16.3 |
| Tokenizer-Übereinstimmung mit Hugging Face | 99.985% von 20033 Texten |
| Embedding vs. Original (Kosinus) | Mittel 0.9775, Minimum 0.7272 |
| `model.onnx` | 41.1 MB, SHA-256 `1c3d4150def1ce959667c7317f7c21660d1febe8f7bda44c19119819a8f213fd` |
| `vocab.txt` | SHA-256 `2f383caa1df10371281ee6e8b313b147f01dc155c3267e4b1b0add57e5a5928b` |

Plausibilität (Kosinus-Ähnlichkeit mit dem gebauten Modell):

| Text A | Text B | erwartet | Ähnlichkeit |
|---|---|---|---|
| Angebot für die Dachsanierung | Kostenvoranschlag Dacharbeiten | ähnlich | 0.623 |
| Rechnung Heizungsinstallation | Zahlungserinnerung Heizung | ähnlich | 0.846 |
| Protokoll Baubesprechung | Niederschrift der Bausitzung | ähnlich | 0.806 |
| Angebot für die Dachsanierung | Einladung zur Weihnachtsfeier | unähnlich | 0.366 |
| Rechnung Heizungsinstallation | Urlaubsantrag Juli | unähnlich | 0.345 |
