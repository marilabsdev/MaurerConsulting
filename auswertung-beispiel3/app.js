(() => {
  const DATA = "../auswertung-beispiel2/data/";
  const FILES = {
    antraege: DATA + "01_Antraege.xlsx",
    artikel: DATA + "02_Artikelstamm.xlsx",
    gutschriften: DATA + "03_Gutschriften.xlsx",
    umsatz: DATA + "04_Umsatz.xlsx",
  };

  const state = {
    antraege: [],
    artikel: [],
    gutschriften: [],
    umsatz: [],
    years: new Set(),
    ready: false,
  };

  /** Sortierung je Tabelle: dir 1 = aufsteigend, -1 = absteigend */
  const sortState = {
    klassen: { key: "Kosten", dir: -1 },
    matrix: { key: "Antraege", dir: -1 },
    quote: { key: "Je1000", dir: -1 },
  };

  /** Datenschnitte: leeres Set = alle Werte */
  const sliceState = {
    klassen: { Warengruppe: new Set() },
    matrix: { Warengruppe: new Set(), Klasse: new Set() },
    quote: { Warengruppe: new Set() },
  };

  const SLICE_LABELS = {
    Klasse: "Fehlerklasse",
    Warengruppe: "Warengruppe",
    Lieferant: "Lieferant",
  };

  const SLICE_CONTAINERS = {
    klassen: "sliceKlassen",
    matrix: "sliceMatrix",
    quote: "sliceQuote",
  };

  const els = {
    years: document.getElementById("years"),
    kpis: document.getElementById("kpis"),
    chartKlassen: document.getElementById("chartKlassen"),
    chartKlassenSum: document.getElementById("chartKlassenSum"),
    tblKlassen: document.getElementById("tblKlassen"),
    tblMatrix: document.getElementById("tblMatrix"),
    tblQuote: document.getElementById("tblQuote"),
    freitext: document.getElementById("freitext"),
    status: document.getElementById("status"),
    error: document.getElementById("error"),
    btnRefresh: document.getElementById("btnRefresh"),
  };

  function euro(n) {
    return (Number(n) || 0).toLocaleString("de-DE", {
      style: "currency",
      currency: "EUR",
      maximumFractionDigits: 0,
    });
  }

  function num(n) {
    return (Number(n) || 0).toLocaleString("de-DE");
  }

  function yearOf(iso) {
    if (iso == null || iso === "") return null;
    if (iso instanceof Date && !isNaN(iso)) return iso.getFullYear();
    if (typeof iso === "number" && iso > 20000 && iso < 80000) {
      const d = new Date(Date.UTC(1899, 11, 30) + iso * 86400000);
      return d.getUTCFullYear();
    }
    const s = String(iso).trim();
    const m = s.match(/(20\d{2})/);
    if (m) return Number(m[1]);
    return null;
  }

  function pick(row, names) {
    for (const n of names) {
      if (row[n] != null && String(row[n]).trim() !== "") return row[n];
    }
    const keys = Object.keys(row || {});
    for (const want of names) {
      const found = keys.find((k) => k.toLowerCase().includes(want.toLowerCase().slice(0, 6)));
      if (found && row[found] != null && String(row[found]).trim() !== "") return row[found];
    }
    return "";
  }

  function normalizeDate(v) {
    if (v instanceof Date && !isNaN(v)) {
      const y = v.getFullYear();
      const m = String(v.getMonth() + 1).padStart(2, "0");
      const d = String(v.getDate()).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }
    if (typeof v === "number" && v > 20000 && v < 80000) {
      const dt = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
      const y = dt.getUTCFullYear();
      const m = String(dt.getUTCMonth() + 1).padStart(2, "0");
      const d = String(dt.getUTCDate()).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }
    const s = String(v || "").trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    const m = s.match(/(20\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})/);
    if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
    return s;
  }

  function daysBetween(a, b) {
    const da = Date.parse(a);
    const db = Date.parse(b);
    if (!da || !db) return null;
    return Math.round((db - da) / 86400000);
  }

  async function loadSheet(url) {
    const res = await fetch(url + "?t=" + Date.now());
    if (!res.ok) throw new Error("Datei nicht gefunden: " + url);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  }

  function mapAntraege(rows) {
    return rows.map((r) => ({
      Antragsdatum: normalizeDate(pick(r, ["Antragsdatum"])),
      Antragsnummer: String(pick(r, ["Antragsnummer"])),
      Artikelnummer: String(pick(r, ["Artikelnummer"])),
      Artikelname: pick(r, ["Artikelname"]),
      Kaufdatum: normalizeDate(pick(r, ["Kaufdatum"])),
      Defekt_Freitext: String(pick(r, ["Defekt_Freitext", "Defekt", "Freitext"]) || ""),
    }));
  }

  function mapArtikel(rows) {
    return rows.map((r) => ({
      Artikelnummer: String(r.Artikelnummer),
      Artikelname: r.Artikelname,
      Warengruppe: r.Warengruppe,
      Lieferant: r.Lieferant,
    }));
  }

  function mapGutschriften(rows) {
    return rows.map((r) => ({
      Antragsnummer: String(r.Antragsnummer),
      Artikelnummer: String(r.Artikelnummer),
      Art: r.Art,
      Menge: Number(r.Menge) || 0,
      Erstattete_Kosten_EUR: Number(r.Erstattete_Kosten_EUR) || 0,
    }));
  }

  function mapUmsatz(rows) {
    return rows.map((r) => ({
      Jahr: Number(r.Jahr),
      Artikelnummer: String(r.Artikelnummer),
      Artikelname: r.Artikelname,
      Warengruppe: r.Warengruppe,
      Menge_verkauft: Number(r.Menge_verkauft) || 0,
      Umsatz_EUR: Number(r.Umsatz_EUR) || 0,
    }));
  }

  function classifyDefect(text) {
    const t = String(text || "").toLowerCase();
    if (/rad|reifen|platte|schwenk|fahrwerk|bremse|lager/.test(t)) return "Räder / Fahrwerk";
    if (/gurt|schloss|knopf|reißverschluss|hosenträger|magnethalter|verdeck|druckknopf/.test(t))
      return "Gurte / Verschlüsse / Verdeck";
    if (/naht|stoff|bezug|füll|ohr|scheuer|eingelaufen|fußsack/.test(t)) return "Stoff / Nähte / Bezug";
    if (/isofix|top-?tether|stützfuß|kopfstütze|arretiert/.test(t)) return "Isofix / Autositz-Mechanik";
    if (/einrast/.test(t) && /sitz|isofix|fuß/.test(t)) return "Isofix / Autositz-Mechanik";
    if (/farbe|waschen|verfärb/.test(t)) return "Farbe / Waschen";
    if (/auge|verschluck|scharf|angesprungen|plastik/.test(t)) return "Sicherheit / Kleinteile";
    if (/griff|gestell|wackel|knack/.test(t)) return "Griff / Gestell";
    return "Sonstiges / unklar";
  }

  function artikelMap() {
    const m = {};
    state.artikel.forEach((a) => {
      m[a.Artikelnummer] = a;
    });
    return m;
  }

  function filteredAntraege() {
    return state.antraege.filter((a) => state.years.has(String(yearOf(a.Antragsdatum))));
  }

  function filteredUmsatz() {
    return state.umsatz.filter((u) => state.years.has(String(u.Jahr)));
  }

  function compute() {
    const antraege = filteredAntraege();
    const arts = artikelMap();
    const umsatz = filteredUmsatz();

    const ids = new Set(antraege.map((a) => a.Antragsnummer));
    const gs = state.gutschriften.filter((g) => ids.has(g.Antragsnummer));
    const anerkanntIds = new Set(gs.map((g) => g.Antragsnummer));
    const kosten = gs.reduce((s, g) => s + g.Erstattete_Kosten_EUR, 0);
    const quote = antraege.length ? anerkanntIds.size / antraege.length : 0;

    const kostenByAntrag = {};
    gs.forEach((g) => {
      kostenByAntrag[g.Antragsnummer] = (kostenByAntrag[g.Antragsnummer] || 0) + g.Erstattete_Kosten_EUR;
    });

    const delays = [];
    antraege.forEach((a) => {
      const d = daysBetween(a.Kaufdatum, a.Antragsdatum);
      if (d != null && d >= 0 && d < 3650) delays.push(d);
    });
    const avgDays = delays.length
      ? Math.round(delays.reduce((s, d) => s + d, 0) / delays.length)
      : null;

    // Anträge mit Warengruppe (für Fehlerklassen-Filter & Matrix)
    const antraegeWg = antraege.map((a) => {
      const wg = (arts[a.Artikelnummer] && arts[a.Artikelnummer].Warengruppe) || "?";
      return { ...a, Warengruppe: wg, Klasse: classifyDefect(a.Defekt_Freitext) };
    });
    const warengruppen = uniqueSorted(antraegeWg.map((a) => a.Warengruppe));

    // Fehlerklassen (gesamt; Übersicht wird in renderKlassen nach Warengruppe neu aggregiert)
    const byKlasse = aggregateKlassen(antraegeWg, anerkanntIds, kostenByAntrag);
    const samplesByKlasse = {};
    antraegeWg.forEach((a) => {
      if (!samplesByKlasse[a.Klasse]) samplesByKlasse[a.Klasse] = [];
      const ft = String(a.Defekt_Freitext || "").trim();
      if (ft && samplesByKlasse[a.Klasse].length < 3 && !samplesByKlasse[a.Klasse].includes(ft)) {
        samplesByKlasse[a.Klasse].push(ft);
      }
    });

    // Matrix: Warengruppe × Fehlerklasse
    const byMatrix = {};
    antraegeWg.forEach((a) => {
      const key = a.Warengruppe + "||" + a.Klasse;
      if (!byMatrix[key]) {
        byMatrix[key] = { Warengruppe: a.Warengruppe, Klasse: a.Klasse, Antraege: 0, Kosten: 0 };
      }
      byMatrix[key].Antraege += 1;
      if (anerkanntIds.has(a.Antragsnummer)) {
        byMatrix[key].Kosten += kostenByAntrag[a.Antragsnummer] || 0;
      }
    });

    // Quote je verkauftem Stück
    const soldByArt = {};
    umsatz.forEach((u) => {
      if (!soldByArt[u.Artikelnummer]) {
        soldByArt[u.Artikelnummer] = {
          Artikelnummer: u.Artikelnummer,
          Artikelname: u.Artikelname,
          Warengruppe: u.Warengruppe,
          Verkauft: 0,
        };
      }
      soldByArt[u.Artikelnummer].Verkauft += u.Menge_verkauft;
    });

    const claimsByArt = {};
    antraege.forEach((a) => {
      if (!claimsByArt[a.Artikelnummer]) {
        const meta = arts[a.Artikelnummer] || {};
        claimsByArt[a.Artikelnummer] = {
          Artikelnummer: a.Artikelnummer,
          Artikelname: a.Artikelname || meta.Artikelname || a.Artikelnummer,
          Warengruppe: meta.Warengruppe || "?",
          Antraege: 0,
          Kosten: 0,
        };
      }
      claimsByArt[a.Artikelnummer].Antraege += 1;
      if (anerkanntIds.has(a.Antragsnummer)) {
        claimsByArt[a.Artikelnummer].Kosten += kostenByAntrag[a.Antragsnummer] || 0;
      }
    });

    const quoteRows = Object.keys(claimsByArt)
      .map((k) => {
        const c = claimsByArt[k];
        const sold = (soldByArt[k] && soldByArt[k].Verkauft) || 0;
        return {
          ...c,
          Verkauft: sold,
          Je1000: sold > 0 ? (c.Antraege / sold) * 1000 : null,
        };
      })
      .filter((r) => r.Verkauft > 0);

    return {
      antraege: antraege.length,
      anerkannt: anerkanntIds.size,
      quote,
      kosten,
      avgDays,
      byKlasse: Object.values(byKlasse),
      byMatrix: Object.values(byMatrix),
      quoteRows,
      samplesByKlasse,
      antraegeWg,
      anerkanntIds,
      kostenByAntrag,
      warengruppen,
    };
  }

  function uniqueSorted(vals) {
    return [...new Set(vals.filter(Boolean))].sort((a, b) =>
      a.localeCompare(b, "de", { sensitivity: "base" })
    );
  }

  function aggregateKlassen(antraegeWg, anerkanntIds, kostenByAntrag) {
    const byKlasse = {};
    antraegeWg.forEach((a) => {
      const klasse = a.Klasse;
      if (!byKlasse[klasse]) {
        byKlasse[klasse] = { Klasse: klasse, Antraege: 0, Kosten: 0, delaySum: 0, delayN: 0 };
      }
      byKlasse[klasse].Antraege += 1;
      if (anerkanntIds.has(a.Antragsnummer)) {
        byKlasse[klasse].Kosten += kostenByAntrag[a.Antragsnummer] || 0;
      }
      const d = daysBetween(a.Kaufdatum, a.Antragsdatum);
      if (d != null && d >= 0 && d < 3650) {
        byKlasse[klasse].delaySum += d;
        byKlasse[klasse].delayN += 1;
      }
    });
    Object.values(byKlasse).forEach((row) => {
      row.AvgDays = row.delayN ? Math.round(row.delaySum / row.delayN) : null;
    });
    return byKlasse;
  }

  /** Fehlerklassen-Übersicht: nach Datenschnitt Warengruppe neu aggregieren */
  function klassenForSlicer(c) {
    const selected = sliceState.klassen.Warengruppe;
    const rows =
      selected.size === 0
        ? c.antraegeWg
        : c.antraegeWg.filter((a) => selected.has(a.Warengruppe));
    return Object.values(aggregateKlassen(rows, c.anerkanntIds, c.kostenByAntrag));
  }

  function sortRows(rows, table) {
    const { key, dir } = sortState[table];
    return [...rows].sort((a, b) => {
      let va = a[key];
      let vb = b[key];
      if (va == null && vb == null) return 0;
      if (va == null) return 1;
      if (vb == null) return -1;
      if (typeof va === "string" || typeof vb === "string") {
        return dir * String(va).localeCompare(String(vb), "de", { sensitivity: "base" });
      }
      return dir * (Number(va) - Number(vb));
    });
  }

  function sliceActive(table) {
    const dims = sliceState[table] || {};
    return Object.values(dims).some((s) => s.size > 0);
  }

  function filterRows(rows, table, exceptField) {
    const dims = sliceState[table] || {};
    return rows.filter((r) => {
      for (const [field, selected] of Object.entries(dims)) {
        if (field === exceptField) continue;
        if (!selected.size) continue;
        if (!selected.has(String(r[field] ?? ""))) return false;
      }
      return true;
    });
  }

  function prepareRows(rows, table) {
    return sortRows(filterRows(rows, table), table);
  }

  function setFilterHint(table, shown, total) {
    const el = document.querySelector('[data-hint="' + table + '"]');
    if (!el) return;
    if (!sliceActive(table)) {
      el.textContent = total ? total + " Zeilen · Klick auf Datenschnitt filtert" : "";
      return;
    }
    el.textContent = shown + " von " + total + " · erneut klicken = abwählen, „Alle“ = zurücksetzen";
  }

  function uniqueValues(rows, field) {
    return [...new Set(rows.map((r) => String(r[field] ?? "")).filter(Boolean))].sort((a, b) =>
      a.localeCompare(b, "de", { sensitivity: "base" })
    );
  }

  function renderSlicers(table, allRows) {
    const host = document.getElementById(SLICE_CONTAINERS[table]);
    const dims = sliceState[table];
    if (!host || !dims) return;
    host.innerHTML = "";
    Object.keys(dims).forEach((field) => {
      const selected = dims[field];
      // Optionen: Werte die nach den anderen Datenschnitten noch vorkommen
      const opts = uniqueValues(filterRows(allRows, table, field), field);
      const block = document.createElement("div");
      block.className = "slicer-block";
      const label = document.createElement("div");
      label.className = "slicer-label";
      label.textContent = "Datenschnitt · " + (SLICE_LABELS[field] || field);
      block.appendChild(label);
      const chips = document.createElement("div");
      chips.className = "slicer-chips";

      const allBtn = document.createElement("button");
      allBtn.type = "button";
      allBtn.className = "slicer-chip all-chip" + (selected.size === 0 ? " active" : "");
      allBtn.textContent = "Alle";
      allBtn.onclick = () => {
        selected.clear();
        render();
      };
      chips.appendChild(allBtn);

      opts.forEach((val) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "slicer-chip" + (selected.has(val) ? " active" : "");
        b.textContent = val;
        b.title = val;
        b.onclick = () => {
          if (selected.has(val)) selected.delete(val);
          else selected.add(val);
          render();
        };
        chips.appendChild(b);
      });
      block.appendChild(chips);
      host.appendChild(block);
    });
  }

  function syncSortHeaders() {
    document.querySelectorAll("th.sortable[data-table]").forEach((th) => {
      const s = sortState[th.dataset.table];
      if (!s) return;
      if (th.dataset.key === s.key) th.dataset.dir = s.dir === 1 ? "asc" : "desc";
      else delete th.dataset.dir;
    });
  }

  function wireSort() {
    document.querySelectorAll("th.sortable[data-table]").forEach((th) => {
      th.addEventListener("click", () => {
        const t = th.dataset.table;
        const k = th.dataset.key;
        if (!sortState[t]) return;
        if (sortState[t].key === k) sortState[t].dir *= -1;
        else {
          sortState[t].key = k;
          sortState[t].dir = th.dataset.type === "str" ? 1 : -1;
        }
        render();
      });
    });
  }

  function renderChips() {
    const allYears = [
      ...new Set(state.antraege.map((a) => String(yearOf(a.Antragsdatum))).filter(Boolean)),
    ].sort();
    els.years.innerHTML = "";
    allYears.forEach((y) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "chip" + (state.years.has(y) ? " active" : "");
      b.textContent = y;
      b.onclick = () => {
        if (state.years.has(y)) {
          if (state.years.size === 1) return;
          state.years.delete(y);
        } else state.years.add(y);
        render();
      };
      els.years.appendChild(b);
    });
  }

  function renderKpis(c) {
    const items = [
      ["Anträge", num(c.antraege)],
      ["Anerkannt", num(c.anerkannt)],
      ["Anerkennungsquote", (c.quote * 100).toFixed(0) + " %"],
      ["Kosten anerkannt", euro(c.kosten)],
      ["Ø Tage Kauf→Antrag", c.avgDays != null ? num(c.avgDays) : "–"],
    ];
    els.kpis.innerHTML = items
      .map(
        ([label, value]) =>
          `<div class="kpi"><div class="label">${label}</div><div class="value">${value}</div></div>`
      )
      .join("");
  }

  function renderKlassen(c) {
    // Datenschnitt nur Warengruppe — Kennzahlen werden neu aggregiert
    const sliceRows = (c.warengruppen || []).map((wg) => ({ Warengruppe: wg }));
    renderSlicers("klassen", sliceRows);

    const byKlasse = klassenForSlicer(c);
    const chartRows = [...byKlasse].sort((a, b) => b.Antraege - a.Antraege);
    const tableRows = sortRows(byKlasse, "klassen");
    const totalAntraege = (c.antraegeWg || []).length;
    const shownAntraege = chartRows.reduce((s, r) => s + r.Antraege, 0);
    setFilterHint(
      "klassen",
      shownAntraege,
      totalAntraege
    );
    // Hinweis textlich anpassen: Anträge, nicht Tabellenzeilen
    const hint = document.querySelector('[data-hint="klassen"]');
    if (hint) {
      if (!sliceActive("klassen")) {
        hint.textContent =
          byKlasse.length +
          " Fehlerklassen · Datenschnitt Warengruppe filtert die Übersicht";
      } else {
        hint.textContent =
          shownAntraege +
          " von " +
          totalAntraege +
          " Anträgen · " +
          byKlasse.length +
          " Klassen · „Alle“ setzt zurück";
      }
    }

    const max = Math.max(1, ...chartRows.map((r) => r.Antraege), 1);
    const padL = 8,
      padT = 8,
      padB = 8;
    const rowH = 28;
    const W = 720;
    const barMax = 420;
    const H = padT + padB + Math.max(chartRows.length, 1) * rowH;

    els.chartKlassenSum.textContent =
      (sliceActive("klassen") ? "Gefiltert · " : "") +
      chartRows.length +
      " Klassen · " +
      euro(chartRows.reduce((s, r) => s + r.Kosten, 0)) +
      " Kosten";

    els.chartKlassen.innerHTML =
      chartRows.length === 0
        ? `<p class="text">Keine Anträge für diese Warengruppe.</p>`
        : `
      <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Fehlerklassen">
        ${chartRows
          .map((r, i) => {
            const y = padT + i * rowH;
            const w = (r.Antraege / max) * barMax;
            const label = r.Klasse.length > 28 ? r.Klasse.slice(0, 26) + "…" : r.Klasse;
            return `
              <text x="${padL}" y="${y + 16}" font-size="12" fill="#5a3fa8">${label}</text>
              <rect x="210" y="${y + 4}" width="${w}" height="18" rx="4" fill="#a855f7" opacity="0.85"/>
              <text x="${220 + w}" y="${y + 16}" font-size="11" fill="#6b21a8">${r.Antraege} · ${euro(r.Kosten)}</text>
            `;
          })
          .join("")}
      </svg>`;

    els.tblKlassen.innerHTML =
      tableRows.length === 0
        ? `<tr><td colspan="4">Keine Treffer für diese Warengruppe.</td></tr>`
        : tableRows
            .map(
              (r) => `<tr>
        <td>${r.Klasse}</td>
        <td class="num">${num(r.Antraege)}</td>
        <td class="num">${euro(r.Kosten)}</td>
        <td class="num">${r.AvgDays != null ? num(r.AvgDays) : "–"}</td>
      </tr>`
            )
            .join("");
  }

  function renderMatrix(c) {
    renderSlicers("matrix", c.byMatrix);
    const rows = prepareRows(c.byMatrix, "matrix");
    setFilterHint("matrix", rows.length, c.byMatrix.length);
    els.tblMatrix.innerHTML =
      rows.length === 0
        ? `<tr><td colspan="4">Keine Treffer.</td></tr>`
        : rows
            .map(
              (r) => `<tr>
        <td>${r.Warengruppe}</td>
        <td>${r.Klasse}</td>
        <td class="num">${num(r.Antraege)}</td>
        <td class="num">${euro(r.Kosten)}</td>
      </tr>`
            )
            .join("");
  }

  function renderQuote(c) {
    renderSlicers("quote", c.quoteRows);
    const rows = prepareRows(c.quoteRows, "quote");
    setFilterHint("quote", rows.length, c.quoteRows.length);
    els.tblQuote.innerHTML =
      rows.length === 0
        ? `<tr><td colspan="6">Keine Treffer.</td></tr>`
        : rows
            .map(
              (r) => `<tr>
        <td>${r.Artikelname}<div class="files">${r.Artikelnummer}</div></td>
        <td>${r.Warengruppe}</td>
        <td class="num">${num(r.Verkauft)}</td>
        <td class="num">${num(r.Antraege)}</td>
        <td class="num">${r.Je1000 != null ? r.Je1000.toFixed(1) : "–"}</td>
        <td class="num">${euro(r.Kosten)}</td>
      </tr>`
            )
            .join("");
  }

  function renderFreitext(c) {
    const top = [...c.byKlasse].sort((a, b) => b.Antraege - a.Antraege).slice(0, 4);
    if (!top.length) {
      els.freitext.innerHTML = '<p class="text">Keine Freitexte.</p>';
      return;
    }
    els.freitext.innerHTML = top
      .map((k) => {
        const samples = c.samplesByKlasse[k.Klasse] || [];
        return `<div style="margin-bottom:14px">
          <div class="group-label">${k.Klasse} (${num(k.Antraege)})</div>
          ${
            samples.length
              ? samples.map((s) => `<p class="quote">„${s}"</p>`).join("")
              : '<p class="text">Keine Beispiele.</p>'
          }
        </div>`;
      })
      .join("");
  }

  function render() {
    if (!state.ready) return;
    const c = compute();
    renderChips();
    renderKpis(c);
    renderKlassen(c);
    renderMatrix(c);
    renderQuote(c);
    renderFreitext(c);
    syncSortHeaders();
  }

  async function loadAll() {
    els.error.hidden = true;
    els.btnRefresh.disabled = true;
    els.status.textContent = "Lade Excel…";
    try {
      if (typeof XLSX === "undefined") throw new Error("SheetJS nicht geladen (vendor).");
      const [antraege, artikel, gutschriften, umsatz] = await Promise.all([
        loadSheet(FILES.antraege),
        loadSheet(FILES.artikel),
        loadSheet(FILES.gutschriften),
        loadSheet(FILES.umsatz),
      ]);
      state.antraege = mapAntraege(antraege);
      state.artikel = mapArtikel(artikel);
      state.gutschriften = mapGutschriften(gutschriften);
      state.umsatz = mapUmsatz(umsatz);
      state.years = new Set(
        state.antraege.map((a) => String(yearOf(a.Antragsdatum))).filter(Boolean)
      );
      state.ready = true;
      els.status.textContent =
        "Geladen · gemeinsame Daten · " +
        state.antraege.length +
        " Anträge · " +
        [...state.years].sort().join(", ");
      render();
    } catch (e) {
      els.error.hidden = false;
      els.error.textContent = String(e.message || e);
      els.status.textContent = "Fehler";
    } finally {
      els.btnRefresh.disabled = false;
    }
  }

  els.btnRefresh.onclick = () => loadAll();
  wireSort();
  loadAll();
})();
