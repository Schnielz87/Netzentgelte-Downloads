/* EXPORT-START */
// Baut die Arbeitsmappe im Layout der gelieferten Excel (Spalten A–L, Bänder, Farben).
// template: ArrayBuffer der Vorlage (Reiter 2026 bleibt unverändert) oder null.
// jahre: [{ jahr, meta, rows }] – je Jahr ein Reiter.
async function buildWorkbook(ExcelJS, template, jahre) {
  const wb = new ExcelJS.Workbook();
  if (template) {
    try {
      await wb.xlsx.load(template);
      jahre.forEach((j) => { const old = wb.getWorksheet(String(j.jahr)); if (old) wb.removeWorksheet(old.id); });
    } catch (e) { /* Vorlage nicht lesbar: neue Mappe */ }
  }
  const NAVY = "FF14304A", NAVY2 = "FF193956", TEAL = "FF247E88", TXT = "FF172F41", GREY = "FF314355";
  const VAL = "FFFFF7E8", ZEBRA = "FFEFF4F6", LIGHT = "FFE7EFF2", NOTE = "FFF2F5F7", WARN = "FFFFF0C5", WARN_TXT = "FF8A4A00";
  const thin = { style: "thin", color: { argb: "FFC9D3DC" } };
  const box = { top: thin, left: thin, bottom: thin, right: thin };
  const fill = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
  const KEYS = ["gp", "ap", "et", "dt", "mme"];
  const COLKEY = { 4: "gp", 5: "ap", 6: "et", 7: "dt", 8: "mme", 11: "stand", 12: "hinweis" };

  for (const j of jahre) {
    const y = j.jahr, meta = j.meta || {}, rows = j.rows || [];
    const ws = wb.addWorksheet(String(y), {
      views: [{ state: "frozen", xSplit: 3, ySplit: 8, zoomScale: 90 }],
      pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: "8:8" }
    });
    [6, 30, 36, 15, 15, 15, 15, 17, 58, 58, 30, 70].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    const band = (r1, r2, text, font, bg, height) => {
      ws.mergeCells(r1, 1, r2, 12);
      const c = ws.getCell(r1, 1);
      c.value = text; c.font = Object.assign({ name: "Calibri" }, font); c.fill = fill(bg);
      c.alignment = { vertical: "middle", wrapText: true, indent: 1 };
      if (height) ws.getRow(r1).height = height;
    };
    const stand = meta.geprueft_text ? "Stand " + meta.geprueft_text : "noch nicht recherchiert";
    band(1, 2, "NETZENTGELTE STROM  |  " + y + "  ·  SLP-Netznutzung und Messstellenbetrieb (ET · DT · mME)", { bold: true, size: 18, color: { argb: "FFFFFFFF" } }, NAVY);
    ws.getRow(1).height = 25; ws.getRow(2).height = 25;
    band(3, 3, "SVS · Marktüberblick mit " + rows.length + " Netzbetreibern · " + stand + " · Quellen: offizielle Preisblätter der Netzbetreiber (PDF geöffnet und Tabellen ausgelesen)", { size: 10, color: { argb: "FFFFFFFF" } }, TEAL, 30);
    band(4, 4, "Geltungszeitraum: 01.01.–31.12." + y + "  ·  Alle Preise NETTO (ohne USt.)  ·  Grund-/Arbeitspreis: Niederspannung ohne Leistungsmessung (SLP), Standardtarif ohne Kommunalrabatt, ohne § 14a-Module", { bold: true, size: 11, color: { argb: NAVY } }, LIGHT, 32);
    band(5, 5, "ABGRENZUNG  |  Nicht enthalten: Konzessionsabgabe, Umlagen (KWKG, § 19 StromNEV, Offshore), Umsatzsteuer. ET/DT = konventionelle Zähler bei jährlicher Messung (Standard). mME = Jahrespreis Anschlussnutzer nach MsbG. ET, DT und mME sind alternative Gerätearten – nicht addieren.", { size: 10, color: { argb: GREY } }, NOTE, 35);
    band(6, 6, "STATUS  |  Vorläufige Netzentgelte (§ 20 Abs. 1 Satz 2 EnWG) vor Abrechnung auf endgültige Fassungen prüfen. Abweichungen und Sonderfälle in Spalte L." + (Array.isArray(meta.aenderungen) && meta.aenderungen.length ? "  Letzte Prüfung: " + meta.aenderungen.join(" ") : ""), { size: 10, color: { argb: WARN_TXT } }, "FFFFF4D6", 35);

    const headers = ["Nr.", "Bezeichnung", "Korrekte Firmierung Netzbetreiber", "Grundpreis SLP\n€/Jahr netto", "Arbeitspreis SLP\nct/kWh netto",
      "ET – Eintarifzähler\n€/Jahr netto", "DT – Doppeltarifzähler\n€/Jahr netto", "mME – moderne Messeinrichtung\n€/Jahr netto",
      "Preisblatt Netznutzung / Quelle", "Preisblatt Messstellenbetrieb / Quelle", "Stand / Gültigkeit", "Fachlicher Hinweis"];
    const hr = ws.getRow(8); hr.height = 54;
    headers.forEach((h, i) => {
      const c = hr.getCell(i + 1); c.value = h;
      c.font = { name: "Calibri", bold: true, size: 10, color: { argb: "FFFFFFFF" } };
      c.fill = fill(i + 1 <= 5 || i + 1 >= 9 ? NAVY : NAVY2);
      c.alignment = { horizontal: "center", vertical: "middle", wrapText: true }; c.border = box;
    });

    rows.forEach((r, idx) => {
      const n = idx + 1, row = ws.getRow(8 + n); row.height = 105;
      const flags = new Set(Array.isArray(r.flags) ? r.flags : []);
      if (KEYS.some((k) => typeof r[k] !== "number")) { KEYS.forEach((k) => { if (typeof r[k] !== "number") flags.add(k); }); flags.add("stand"); flags.add("hinweis"); }
      const vals = [r.nr ?? n, r.bez || "", r.firma || "", r.gp, r.ap, r.et, r.dt, r.mme, r.q_nne || "", r.q_msb || "", r.stand || "", r.hinweis || ""];
      vals.forEach((v, i) => {
        const col = i + 1, c = row.getCell(col);
        c.value = (v === undefined || v === null) ? null : v;
        c.font = { name: "Calibri", size: 10, color: { argb: TXT } }; c.border = box;
        if (col >= 4 && col <= 8) {
          c.numFmt = "#,##0.00"; c.fill = fill(VAL); c.alignment = { horizontal: "right", vertical: "middle" };
        } else {
          if (n % 2 === 0) c.fill = fill(ZEBRA);
          c.alignment = { horizontal: col === 1 ? "center" : undefined, vertical: "middle", wrapText: true };
        }
        if (r.neu && col <= 3) { c.fill = fill("FFEFEAFD"); c.font = { name: "Calibri", size: 10, bold: col === 2, color: { argb: "FF5B3FC4" } }; }
        if (r.neu && col === 2) c.value = (r.bez || "") + "  (NEU)";
        if (COLKEY[col] && flags.has(COLKEY[col])) { c.fill = fill(WARN); c.font = { name: "Calibri", size: 10, bold: true, color: { argb: WARN_TXT } }; }
      });
    });
    const foot = 8 + rows.length + 2;
    band(foot, foot, "QUELLENHINWEIS  |  Spalten I und J enthalten die direkten PDF-Links der Netzbetreiber mit Seitenangabe. Gelb markierte Werte: Besonderheit zur Gültigkeit oder noch offen, siehe Spalte L. Violett markiert: neu aufgenommene Netzbetreiber (Werte noch ungeprüft); die Markierung entfällt mit der nächsten Aktualisierung. Für abrechnungsrelevante Verwendung Zusatzleistungen (z. B. Tarifschaltgerät, Wandler), abweichende Messintervalle und endgültige Preisblätter prüfen.", { size: 10, color: { argb: NAVY } }, LIGHT, 45);
  }
  const last = jahre.length ? wb.getWorksheet(String(jahre[jahre.length - 1].jahr)) : null;
  if (last) wb.views = [{ activeTab: wb.worksheets.indexOf(last) }];
  return wb.xlsx.writeBuffer();
}
/* EXPORT-END */
