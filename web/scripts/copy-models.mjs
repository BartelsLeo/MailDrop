// Kopiert Models/model.onnx und Models/vocab.txt aus dem Repo-Stamm nach web/public/models,
// damit Vite sie ausliefert (dev) bzw. ins Build übernimmt. Die Kopien sind per .gitignore
// ausgeschlossen - einzige eingecheckte Quelle bleibt Models/. (Die WASM-Laufzeit von ONNX
// Runtime Web bündelt Vite selbst mit aus - sie kommt also vom eigenen Webspace, nicht von
// einem CDN, was Firmen-Proxy/CSP-Probleme vermeidet.)
import { copyFileSync, mkdirSync, statSync, existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const webDir = join(dirname(fileURLToPath(import.meta.url)), "..");
const source = join(webDir, "..", "Models");
const target = join(webDir, "public", "models");
mkdirSync(target, { recursive: true });

for (const file of ["model.onnx", "vocab.txt"]) {
  const from = join(source, file);
  const to = join(target, file);
  if (!existsSync(from)) {
    console.warn(`[models] ${from} fehlt - Modellprüfung wird im Prototyp fehlschlagen.`);
    continue;
  }
  if (existsSync(to) && statSync(to).size === statSync(from).size) continue;
  copyFileSync(from, to);
  console.log(`[models] ${file} kopiert.`);
}

