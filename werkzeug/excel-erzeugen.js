// Erzeugt Netzentgelte_Strom.xlsx und Netzentgelte_Gas.xlsx aus daten.json.
// Aufruf (im Repo-Hauptordner): npm install --prefix werkzeug && node werkzeug/excel-erzeugen.js
const path = require("path"), fs = require("fs");
const ExcelJS = require(path.join(__dirname, "node_modules", "exceljs"));
eval(fs.readFileSync(path.join(__dirname, "export.js"), "utf8") + ";globalThis.buildWorkbook = buildWorkbook; globalThis.buildGasWorkbook = buildGasWorkbook;");
const root = path.join(__dirname, "..");
const daten = JSON.parse(fs.readFileSync(path.join(root, "daten.json"), "utf8"));
const tpl = fs.readFileSync(path.join(__dirname, "vorlage-2026.xlsx"));
const jahre = Object.keys(daten.jahre).sort().map((y) => ({ jahr: Number(y), meta: daten.jahre[y].meta, rows: daten.jahre[y].rows }));
(async () => {
  const s = await globalThis.buildWorkbook(ExcelJS, tpl.buffer.slice(tpl.byteOffset, tpl.byteOffset + tpl.length), jahre);
  fs.writeFileSync(path.join(root, "Netzentgelte_Strom.xlsx"), Buffer.from(s));
  console.log("Netzentgelte_Strom.xlsx geschrieben:", jahre.map((j) => j.jahr).join(", "));
  if (daten.gas) {
    const g = await globalThis.buildGasWorkbook(ExcelJS, daten.gas);
    fs.writeFileSync(path.join(root, "Netzentgelte_Gas.xlsx"), Buffer.from(g));
    console.log("Netzentgelte_Gas.xlsx geschrieben");
  }
})();
