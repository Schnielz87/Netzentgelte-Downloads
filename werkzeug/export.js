/* EXPORT-START */
// Baut die Arbeitsmappe im Layout der gelieferten Excel (Spalten A–L, Bänder, Farben).
// template: ArrayBuffer der Vorlage (Reiter 2026 bleibt unverändert) oder null.
// jahre: [{ jahr, meta, rows }] – je Jahr ein Reiter.
async function buildWorkbook(ExcelJS, template, jahre, gas) {
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

// Eigene Gas-Arbeitsmappe im Aufbau der Kundenvorlage: je Netzbetreiber ein Block, Jahre nebeneinander.
async function buildGasWorkbook(ExcelJS, gas) {
  const wb = new ExcelJS.Workbook();
  const fill = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
  const thin = { style: "thin", color: { argb: "FF000000" } }, med = { style: "medium", color: { argb: "FF000000" } };
  const box = { top: thin, left: thin, bottom: thin, right: thin };
  const PAL = ["FFE2EFDA", "FFEAF1A0", "FFFCE3A9", "FFF8CBCB"]; // 2027 hellgrün · 2028 gelbgrün · 2029 orange · 2030 rosa
  const yc = (y) => PAL[Math.max(0, Number(y) - 2027) % 4];
  const num = (v) => typeof v === "number";
  const all = Object.keys(gas.jahre).sort();
  const hasData = (y) => (gas.jahre[y].rows || []).some((r) => (r.stufen || []).some((s) => num(s.ap)) || (r.zaehler || []).some((z) => num(z.msb)));
  const lastData = all.filter(hasData).pop() || all[0];
  const ys = all; // alle Jahre der App (ab 2027), wie beim Strom
  const ws = wb.addWorksheet("Gas", { views: [{ state: "frozen", xSplit: 3, ySplit: 2, zoomScale: 90 }], pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
  const W = 4, C0 = 4;
  ws.getColumn(1).width = 9; ws.getColumn(2).width = 19; ws.getColumn(3).width = 11;
  ys.forEach((y, k) => { for (let j = 0; j < W; j++) ws.getColumn(C0 + k * W + j).width = 12; });
  const t = ws.getCell(1, 1); ws.mergeCells(1, 1, 1, 3); t.value = "Netznutzungsentgelte"; t.font = { name: "Arial", bold: true, size: 14 }; t.border = { top: thin, left: thin, bottom: thin, right: thin };
  ys.forEach((y, k) => { const c = ws.getCell(1, C0 + k * W); ws.mergeCells(1, C0 + k * W, 1, C0 + k * W + W - 1); c.value = String(y); c.font = { name: "Arial", size: 22 }; c.alignment = { horizontal: "center" }; });
  ws.getRow(1).height = 32;
  const n2 = ws.getCell(2, 1); ws.mergeCells(2, 1, 2, 3 + ys.length * W);
  n2.value = "Gas · SLP · netto · AP ct/kWh · GP €/Jahr · MSB und Messung €/Jahr · Quellen und Hinweise: siehe App";
  n2.font = { name: "Arial", size: 9, color: { argb: "FF00A050" } };
  let row = 4;
  const latestRows = gas.jahre[ys[ys.length - 1]].rows || [];
  latestRows.forEach((L, idx) => {
    const per = ys.map((y) => (gas.jahre[y].rows || [])[idx] || {});
    const n = Math.max(...per.map((r) => Math.max((r.stufen || []).length, (r.zaehler || []).length)));
    const sockel = L.typ === "sockel";
    // Kopfzeile: Name + Jahr/Status
    const h = ws.getCell(row, 1); ws.mergeCells(row, 1, row, 2);
    h.value = (L.bez || "") + (L.vnb ? " (" + L.vnb + ")" : ""); h.font = { name: "Arial", bold: true, underline: true, size: 10 }; h.fill = fill(L.neu ? "FFEFEAFD" : "FFFFFF00");
    ys.forEach((y, k) => {
      const r = per[k]; const c = ws.getCell(row, C0 + k * W); c.value = Number(y); c.font = { name: "Arial", bold: true, size: 10 }; c.alignment = { horizontal: "right" };
      c.fill = fill(yc(y)); c.border = { top: med, left: med, bottom: med, right: med };
      const st = ws.getCell(row, C0 + k * W + 1); const s = String(r.stand || "");
      st.value = /endg/i.test(s) ? "endgültig" : (/vorl/i.test(s) ? "vorläufig" : "");
      st.font = { name: "Arial", bold: true, size: 10, color: { argb: "FF00A050" } };
    });
    row++;
    // Spaltenköpfe
    const hr = ws.getRow(row); hr.getCell(3).value = "ZG"; hr.getCell(3).font = { name: "Arial", bold: true, size: 10 };
    [1, 2, 3].forEach((c) => (hr.getCell(c).border = { bottom: med }));
    const heads = sockel ? ["Sockelbetrag", "durch Sockel", "AP", "MSB"] : ["AP", "GP", "MSB", "Messung"];
    ys.forEach((y, k) => heads.forEach((hd, j) => { const c = hr.getCell(C0 + k * W + j); c.value = hd; c.font = { name: "Arial", bold: true, size: 10 }; c.border = { top: med, bottom: med, left: thin, right: thin }; }));
    row++;
    for (let i = 0; i < n; i++) {
      const r = ws.getRow(row);
      const sL = per.map((p) => (p.stufen || [])[i]).filter(Boolean).pop(); const zL = per.map((p) => (p.zaehler || [])[i]).filter(Boolean).pop();
      if (sL) { r.getCell(1).value = "Stufe " + (sL.stufe || i + 1); r.getCell(2).value = sL.verbrauch || ""; }
      if (zL && zL.zg) r.getCell(3).value = zL.zg;
      [1, 2, 3].forEach((c) => { r.getCell(c).border = box; r.getCell(c).font = { name: "Arial", size: 10, bold: c === 3 }; });
      if (zL && zL.zg === "Messung") r.getCell(3).fill = fill("FFBDD7EE");
      ys.forEach((y, k) => {
        const p = per[k], s = (p.stufen || [])[i], z = (p.zaehler || [])[i];
        const isMess = z && z.zg === "Messung";
        const cols = sockel ? [[s && s.sockel, "#,##0.00", !!s], [s && s.sockel_kwh, "#,##0", false], [s && s.ap, "0.0000", !!s], [z && (isMess ? z.msb : z.msb), "#,##0.00", !!(z && z.zg && !isMess)]]
                            : [[s && s.ap, "0.0000", !!s], [s && s.gp, "#,##0.00", !!s], [z && z.msb, "#,##0.00", !!(z && z.zg)], [i === 0 && z ? z.messung : null, "#,##0.00", false]];
        cols.forEach(([v, f, must], j) => {
          const c = r.getCell(C0 + k * W + j); c.value = num(v) ? v : null; c.numFmt = f; c.alignment = { horizontal: "right" }; c.border = box; c.font = { name: "Arial", size: 10 };
          const yearHas = (p.stufen || []).some((x) => num(x.ap)) || (p.zaehler || []).some((x) => num(x.msb));
          c.fill = fill(must && !num(v) && Number(y) >= 2027 ? "FFFFF2CC" : (num(v) || (must && yearHas) ? yc(y) : "FFFFFFFF"));
        });
      });
      row++;
    }
    row++;
  });
  return wb.xlsx.writeBuffer();
}
/* EXPORT-END */
