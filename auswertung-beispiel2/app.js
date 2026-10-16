(() => {
  const FILES = {
    antraege: "data/01_Antraege.xlsx",
    artikel: "data/02_Artikelstamm.xlsx",
    gutschriften: "data/03_Gutschriften.xlsx",
    umsatz: "data/04_Umsatz.xlsx",
  };

  const state = {
    antraege: [],
    artikel: [],
    gutschriften: [],
    umsatz: [],
    years: new Set(),
    ready: false,
  };

  const els = {
    years: document.getElementById("years"),
    kpis: document.getElementById("kpis"),
    chartYears: document.getElementById("chartYears"),
    chartSum: document.getElementById("chartSum"),
    tblArtikel: document.getElementById("tblArtikel"),
    tblLieferant: document.getElementById("tblLieferant"),
    tblKlassen: document.getElementById("tblKlassen"),
    chartKlassen: document.getElementById("chartKlassen"),
    chartKlassenSum: document.getElementById("chartKlassenSum"),
    status: document.getElementById("status"),
    error: document.getElementById("error"),
    btnRefresh: document.getElementById("btnRefresh"),
    chat: document.getElementById("chat"),
    chatForm: document.getElementById("chatForm"),
    chatInput: document.getElementById("chatInput"),
    chatSuggestions: document.getElementById("chatSuggestions"),
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
    if (!iso) return null;
    const s = String(iso).slice(0, 10);
    const y = Number(s.slice(0, 4));
    return Number.isFinite(y) ? y : null;
  }

  async function loadSheet(url) {
    const res = await fetch(url + "?t=" + Date.now());
    if (!res.ok) throw new Error("Datei nicht gefunden: " + url);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { defval: "" });
  }

  function normalizeDate(v) {
    if (v instanceof Date && !isNaN(v)) {
      const y = v.getFullYear();
      const m = String(v.getMonth() + 1).padStart(2, "0");
      const d = String(v.getDate()).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }
    const s = String(v || "").trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    return s;
  }

  function mapAntraege(rows) {
    return rows.map((r) => ({
      Kunde_Name: r.Kunde_Name,
      Kunde_Land: r.Kunde_Land,
      Antragsdatum: normalizeDate(r.Antragsdatum),
      Antragsnummer: String(r.Antragsnummer),
      Artikelnummer: String(r.Artikelnummer),
      Artikelname: r.Artikelname,
      Kaufdatum: normalizeDate(r.Kaufdatum),
      Beantragte_Kostenerstattung_EUR: Number(r.Beantragte_Kostenerstattung_EUR) || 0,
      Defekt_Freitext: r.Defekt_Freitext,
    }));
  }

  function mapArtikel(rows) {
    return rows.map((r) => ({
      Artikelnummer: String(r.Artikelnummer),
      Artikelname: r.Artikelname,
      Einkaufspreis_EUR: Number(r.Einkaufspreis_EUR) || 0,
      Warengruppe: r.Warengruppe,
      Lieferant: r.Lieferant,
    }));
  }

  function mapGutschriften(rows) {
    return rows.map((r) => ({
      Antragsnummer: String(r.Antragsnummer),
      Gutschrift_Rechnung_Nr: r.Gutschrift_Rechnung_Nr,
      Datum_Gutschrift: normalizeDate(r.Datum_Gutschrift),
      Artikelnummer: String(r.Artikelnummer),
      Artikelname: r.Artikelname,
      Art: r.Art,
      Menge: Number(r.Menge) || 0,
      Erstattete_Kosten_EUR: Number(r.Erstattete_Kosten_EUR) || 0,
      Bemerkung: r.Bemerkung,
    }));
  }

  function mapUmsatz(rows) {
    return rows.map((r) => ({
      Jahr: Number(r.Jahr),
      Monat: Number(r.Monat),
      Artikelnummer: String(r.Artikelnummer),
      Artikelname: r.Artikelname,
      Lieferant: r.Lieferant,
      Warengruppe: r.Warengruppe,
      Menge_verkauft: Number(r.Menge_verkauft) || 0,
      Umsatz_EUR: Number(r.Umsatz_EUR) || 0,
    }));
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

  function anerkanntSet(antraege) {
    const ids = new Set(antraege.map((a) => a.Antragsnummer));
    const gs = state.gutschriften.filter((g) => ids.has(g.Antragsnummer));
    return { gs, anerkanntIds: new Set(gs.map((g) => g.Antragsnummer)) };
  }

  function filteredUmsatz() {
    return state.umsatz.filter((u) => state.years.has(String(u.Jahr)));
  }

  /** Klassifiziert Kunden-Freitext (regelbasiert, ohne LLM). */
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

  function compute() {
    const antraege = filteredAntraege();
    const { gs, anerkanntIds } = anerkanntSet(antraege);
    const arts = artikelMap();
    const umsatz = filteredUmsatz();

    const kosten = gs.reduce((s, g) => s + g.Erstattete_Kosten_EUR, 0);
    const mengeErsatz = gs.filter((g) => String(g.Art).includes("Ersatz")).reduce((s, g) => s + g.Menge, 0);
    const quote = antraege.length ? anerkanntIds.size / antraege.length : 0;

    // Kosten je Antrag aus Gutschriften
    const kostenByAntrag = {};
    gs.forEach((g) => {
      kostenByAntrag[g.Antragsnummer] = (kostenByAntrag[g.Antragsnummer] || 0) + g.Erstattete_Kosten_EUR;
    });

    // Fehlerklassen aus Freitext
    const byKlasse = {};
    antraege.forEach((a) => {
      const klasse = classifyDefect(a.Defekt_Freitext);
      if (!byKlasse[klasse]) byKlasse[klasse] = { Klasse: klasse, Antraege: 0, Anerkannt: 0, Kosten: 0 };
      byKlasse[klasse].Antraege += 1;
      if (anerkanntIds.has(a.Antragsnummer)) {
        byKlasse[klasse].Anerkannt += 1;
        byKlasse[klasse].Kosten += kostenByAntrag[a.Antragsnummer] || 0;
      }
    });

    // per article
    const byArt = {};
    antraege.forEach((a) => {
      const k = a.Artikelnummer;
      if (!byArt[k]) {
        const meta = arts[k] || {};
        byArt[k] = {
          Artikelnummer: k,
          Artikelname: a.Artikelname || meta.Artikelname || k,
          Lieferant: meta.Lieferant || "?",
          Warengruppe: meta.Warengruppe || "?",
          EK: meta.Einkaufspreis_EUR || 0,
          Antraege: 0,
          Anerkannt: 0,
          Kosten: 0,
          Menge: 0,
        };
      }
      byArt[k].Antraege += 1;
      byArt[k].Menge += 1;
      if (anerkanntIds.has(a.Antragsnummer)) byArt[k].Anerkannt += 1;
    });
    gs.forEach((g) => {
      if (!byArt[g.Artikelnummer]) return;
      byArt[g.Artikelnummer].Kosten += g.Erstattete_Kosten_EUR;
      if (String(g.Art).includes("Ersatz")) byArt[g.Artikelnummer].Menge += Math.max(0, g.Menge - 1);
    });

    const umsatzByArt = {};
    const umsatzByLieferant = {};
    umsatz.forEach((u) => {
      umsatzByArt[u.Artikelnummer] = (umsatzByArt[u.Artikelnummer] || 0) + u.Umsatz_EUR;
      umsatzByLieferant[u.Lieferant] = (umsatzByLieferant[u.Lieferant] || 0) + u.Umsatz_EUR;
    });

    Object.values(byArt).forEach((row) => {
      row.Umsatz = umsatzByArt[row.Artikelnummer] || 0;
      // grobe Marge: Umsatz - EK * geschätzte Verkaufsmenge aus Umsatz/VK — hier: Umsatz * 0.45 als Proxy wenn keine VK
      row.GewinnProxy = row.Umsatz * 0.35 - row.Kosten;
    });

    const byLieferant = {};
    Object.values(byArt).forEach((row) => {
      const L = row.Lieferant;
      if (!byLieferant[L]) byLieferant[L] = { Lieferant: L, Antraege: 0, Kosten: 0, Umsatz: umsatzByLieferant[L] || 0 };
      byLieferant[L].Antraege += row.Antraege;
      byLieferant[L].Kosten += row.Kosten;
      byLieferant[L].Umsatz = umsatzByLieferant[L] || byLieferant[L].Umsatz;
    });
    // Lieferanten nur aus Umsatz, ohne Anträge
    Object.keys(umsatzByLieferant).forEach((L) => {
      if (!byLieferant[L]) byLieferant[L] = { Lieferant: L, Antraege: 0, Kosten: 0, Umsatz: umsatzByLieferant[L] };
    });

    Object.values(byLieferant).forEach((row) => {
      row.Risiko = row.Umsatz > 0 ? (row.Antraege / row.Umsatz) * 10000 : row.Antraege * 100;
    });

    // years series
    const yearKeys = [...state.years].map(Number).sort();
    const series = yearKeys.map((y) => {
      const ya = state.antraege.filter((a) => yearOf(a.Antragsdatum) === y);
      const ids = new Set(ya.map((a) => a.Antragsnummer));
      const yg = state.gutschriften.filter((g) => ids.has(g.Antragsnummer));
      const anerkannt = new Set(yg.map((g) => g.Antragsnummer)).size;
      return { year: y, antraege: ya.length, anerkannt, kosten: yg.reduce((s, g) => s + g.Erstattete_Kosten_EUR, 0) };
    });

    return {
      antraege: antraege.length,
      anerkannt: anerkanntIds.size,
      quote,
      kosten,
      mengeErsatz,
      byArt: Object.values(byArt).sort((a, b) => b.Antraege - a.Antraege),
      byLieferant: Object.values(byLieferant).sort((a, b) => b.Risiko - a.Risiko),
      byKlasse: Object.values(byKlasse).sort((a, b) => b.Antraege - a.Antraege),
      series,
    };
  }

  function renderChips() {
    const allYears = [...new Set(state.antraege.map((a) => String(yearOf(a.Antragsdatum))).filter(Boolean))].sort();
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
      ["Ersatz 1:1 (Menge)", num(c.mengeErsatz)],
    ];
    els.kpis.innerHTML = items
      .map(
        ([label, value]) => `<div class="kpi"><div class="label">${label}</div><div class="value">${value}</div></div>`
      )
      .join("");
  }

  function renderChart(series) {
    const padL = 48, padR = 16, padT = 20, padB = 36;
    const W = 720, H = 260;
    const innerW = W - padL - padR;
    const innerH = H - padT - padB;
    let maxVal = 1;
    series.forEach((s) => {
      maxVal = Math.max(maxVal, s.antraege, s.anerkannt);
    });
    els.chartSum.textContent = "Kosten gesamt: " + euro(series.reduce((a, s) => a + s.kosten, 0));

    const n = Math.max(series.length, 1);
    const groupW = innerW / n;
    const barW = Math.min(28, groupW / 2 - 8);
    let html = "";
    for (let i = 0; i <= 4; i++) {
      const y = padT + (innerH * i) / 4;
      const val = Math.round(maxVal * (1 - i / 4));
      html += `<line x1="${padL}" y1="${y}" x2="${W - padR}" y2="${y}" stroke="#e8e3f4" stroke-width="1"/>`;
      html += `<text x="${padL - 8}" y="${y + 4}" text-anchor="end" font-size="11" fill="#6b6275">${val}</text>`;
    }
    series.forEach((s, gi) => {
      const gx = padL + gi * groupW + groupW / 2;
      const draw = (value, color, offset) => {
        const h = (value / maxVal) * innerH;
        const x = gx + offset;
        const y = padT + innerH - h;
        html += `<rect x="${x}" y="${y}" width="${barW}" height="${Math.max(h, 0)}" rx="4" fill="${color}"><title>${value}</title></rect>`;
      };
      draw(s.antraege, "#2e1477", -barW - 2);
      draw(s.anerkannt, "#8b7bc4", 2);
      html += `<text x="${gx}" y="${H - 12}" text-anchor="middle" font-size="12" fill="#2b2433">${s.year}</text>`;
    });
    els.chartYears.innerHTML = html;
  }

  function renderKlassen(c) {
    const rows = c.byKlasse || [];
    els.chartKlassenSum.textContent =
      "Anträge in Klassen: " + num(rows.reduce((s, r) => s + r.Antraege, 0)) +
      " · Kosten anerkannter Fälle: " + euro(rows.reduce((s, r) => s + r.Kosten, 0));

    els.tblKlassen.innerHTML = rows
      .map(
        (r) => `<tr>
        <td>${r.Klasse}</td>
        <td class="num">${num(r.Antraege)}</td>
        <td class="num">${num(r.Anerkannt)}</td>
        <td class="num">${euro(r.Kosten)}</td>
      </tr>`
      )
      .join("");

    // Balken: Anträge je Klasse
    const padL = 48, padR = 16, padT = 20, padB = 70;
    const W = 720, H = 260;
    const innerW = W - padL - padR;
    const innerH = H - padT - padB;
    let maxVal = 1;
    rows.forEach((r) => {
      maxVal = Math.max(maxVal, r.Antraege);
    });
    const n = Math.max(rows.length, 1);
    const groupW = innerW / n;
    const barW = Math.min(36, groupW - 10);
    let html = "";
    for (let i = 0; i <= 4; i++) {
      const y = padT + (innerH * i) / 4;
      const val = Math.round(maxVal * (1 - i / 4));
      html += `<line x1="${padL}" y1="${y}" x2="${W - padR}" y2="${y}" stroke="#e8e3f4" stroke-width="1"/>`;
      html += `<text x="${padL - 8}" y="${y + 4}" text-anchor="end" font-size="11" fill="#6b6275">${val}</text>`;
    }
    rows.forEach((r, gi) => {
      const gx = padL + gi * groupW + groupW / 2;
      const h = (r.Antraege / maxVal) * innerH;
      const x = gx - barW / 2;
      const y = padT + innerH - h;
      html += `<rect x="${x}" y="${y}" width="${barW}" height="${Math.max(h, 0)}" rx="4" fill="#2e1477"><title>${r.Klasse}: ${r.Antraege}</title></rect>`;
      const label = r.Klasse.length > 14 ? r.Klasse.slice(0, 12) + "…" : r.Klasse;
      html += `<text x="${gx}" y="${H - 28}" text-anchor="middle" font-size="10" fill="#2b2433" transform="rotate(-28 ${gx} ${H - 28})">${label}</text>`;
    });
    els.chartKlassen.innerHTML = html;
  }

  function renderTables(c) {
    renderKlassen(c);

    els.tblArtikel.innerHTML = c.byArt
      .slice(0, 8)
      .map(
        (r) => `<tr>
        <td>${r.Artikelname}<div class="files">${r.Artikelnummer} · ${r.Warengruppe}</div></td>
        <td>${r.Lieferant}</td>
        <td class="num">${num(r.Antraege)}</td>
        <td class="num">${num(r.Anerkannt)}</td>
        <td class="num">${euro(r.Kosten)}</td>
      </tr>`
      )
      .join("");

    els.tblLieferant.innerHTML = c.byLieferant
      .map(
        (r) => `<tr>
        <td>${r.Lieferant}</td>
        <td class="num">${num(r.Antraege)}</td>
        <td class="num">${euro(r.Umsatz)}</td>
        <td class="num">${r.Risiko.toFixed(2)}</td>
        <td class="num">${euro(r.Kosten)}</td>
      </tr>`
      )
      .join("");
  }

  function addChat(text, who) {
    const div = document.createElement("div");
    div.className = "msg " + who;
    div.textContent = text;
    els.chat.appendChild(div);
    els.chat.scrollTop = els.chat.scrollHeight;
  }

  function answerChat(q) {
    if (!state.ready) return "Bitte zuerst die Excel-Daten laden.";
    const c = compute();
    const ql = q.toLowerCase();

    if (/trennen|lieferant|lieferanten|risiko/.test(ql)) {
      const top = c.byLieferant[0];
      if (!top) return "Keine Lieferanten-Daten im Filter.";
      const lines = c.byLieferant
        .slice(0, 3)
        .map(
          (r, i) =>
            `${i + 1}. ${r.Lieferant}: ${r.Antraege} Anträge, Umsatz ${euro(r.Umsatz)}, Risiko ${r.Risiko.toFixed(2)} Anträge/10k€`
        );
      return (
        `Nach Reklamationen relativ zum Umsatz wäre der kritischste Lieferant: ${top.Lieferant}.\n\n` +
        lines.join("\n") +
        `\n\n(Demo-Logik: Anträge je 10.000 € Umsatz — höher = problematischer.)`
      );
    }

    if (/häufig|häufigsten|kaputt|reklamiert|top.?artikel|meisten/.test(ql)) {
      const top = c.byArt.slice(0, 5);
      if (!top.length) return "Keine Anträge im gewählten Zeitraum.";
      return (
        "Am häufigsten reklamierte Artikel (gefilterte Jahre):\n\n" +
        top
          .map((r, i) => `${i + 1}. ${r.Artikelname} (${r.Lieferant}): ${r.Antraege} Anträge, Kosten ${euro(r.Kosten)}`)
          .join("\n")
      );
    }

    if (/gewinn|trotzdem|behalten|sortiment|marge/.test(ql) && /reklam/.test(ql)) {
      const list = c.byArt
        .filter((r) => r.Antraege >= 3 && r.GewinnProxy > 0)
        .sort((a, b) => b.GewinnProxy - a.GewinnProxy)
        .slice(0, 5);
      if (!list.length) return "Im Filter keine stark reklamierten Artikel mit positivem Gewinn-Proxy gefunden.";
      return (
        "Oft reklamiert, aber laut grobem Gewinn-Proxy (Umsatz×35% − Reklamationskosten) trotzdem positiv:\n\n" +
        list
          .map((r) => `• ${r.Artikelname}: ${r.Antraege} Anträge, Umsatz ${euro(r.Umsatz)}, Proxy ${euro(r.GewinnProxy)}`)
          .join("\n") +
        "\n\nHinweis: Proxy ist vereinfacht — für echte Entscheidungen EK/VK aus dem ERP nachziehen."
      );
    }

    if (/drauf|verlust|wenig verkauf|selten/.test(ql)) {
      const list = c.byArt
        .filter((r) => r.Antraege <= 5)
        .map((r) => ({ ...r, score: r.Umsatz < 5000 || r.GewinnProxy < 0 }))
        .filter((r) => r.score || r.GewinnProxy < 0)
        .sort((a, b) => a.GewinnProxy - b.GewinnProxy)
        .slice(0, 5);
      if (!list.length) {
        // fallback: lowest profit proxy
        const low = [...c.byArt].sort((a, b) => a.GewinnProxy - b.GewinnProxy).slice(0, 5);
        return (
          "Artikel mit schwachem Gewinn-Proxy (auch bei wenigen Reklamationen):\n\n" +
          low.map((r) => `• ${r.Artikelname}: Anträge ${r.Antraege}, Umsatz ${euro(r.Umsatz)}, Proxy ${euro(r.GewinnProxy)}`).join("\n")
        );
      }
      return (
        "Selten/mäßig reklamiert, aber wirtschaftlich eher schwach (wenig Umsatz oder negativer Proxy):\n\n" +
        list.map((r) => `• ${r.Artikelname}: Anträge ${r.Antraege}, Umsatz ${euro(r.Umsatz)}, Proxy ${euro(r.GewinnProxy)}`).join("\n")
      );
    }

    if (/quote|anerkannt|kosten|kennzahl|jahres/.test(ql)) {
      return (
        `Kennzahlen im Filter:\n` +
        `• Anträge: ${num(c.antraege)}\n` +
        `• Anerkannt: ${num(c.anerkannt)} (${(c.quote * 100).toFixed(0)} %)\n` +
        `• Kosten anerkannter Fälle: ${euro(c.kosten)}\n` +
        `• Ersatz 1:1 Menge: ${num(c.mengeErsatz)}\n\n` +
        c.series.map((s) => `${s.year}: ${s.antraege} Anträge, ${s.anerkannt} anerkannt, ${euro(s.kosten)}`).join("\n")
      );
    }

    return (
      "Dazu habe ich kein festes Muster. Probieren Sie z. B.:\n" +
      "• Welche Artikel werden am häufigsten reklamiert?\n" +
      "• Von welchem Lieferanten sollte ich mich trennen?\n" +
      "• Welche oft reklamierten Artikel trotzdem behalten?\n" +
      "• Welche Artikel selten reklamiert, aber wirtschaftlich schwach?"
    );
  }

  function askChat(q) {
    const text = (q || "").trim();
    if (!text) return;
    addChat(text, "user");
    addChat(answerChat(text), "bot");
    els.chatInput.value = "";
  }

  function setupChat() {
    els.chat.innerHTML = "";
    addChat(
      "Hallo! Ich beantworte Fragen zu den geladenen Excel-Daten (Demo-Logik im Browser).\nJahr-Filter oben gilt auch für meine Antworten.",
      "bot"
    );
    const prompts = [
      "Welche Artikel werden am häufigsten reklamiert?",
      "Von welchem Lieferanten sollte ich mich trennen?",
      "Welche oft reklamierten Artikel trotzdem behalten?",
      "Welche Artikel selten reklamiert, aber ich zahle drauf?",
      "Wie sind Quote und Kosten im Jahresvergleich?",
    ];
    els.chatSuggestions.innerHTML = "";
    prompts.forEach((p) => {
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = p;
      b.onclick = () => askChat(p);
      els.chatSuggestions.appendChild(b);
    });
  }

  function render() {
    if (!state.ready) return;
    renderChips();
    const c = compute();
    renderKpis(c);
    renderChart(c.series);
    renderTables(c);
  }

  async function loadAll() {
    els.error.hidden = true;
    els.btnRefresh.disabled = true;
    els.status.textContent = "Lade Excel-Dateien…";
    try {
      if (typeof XLSX === "undefined") throw new Error("SheetJS-Bibliothek nicht geladen (CDN).");
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
      const years = [...new Set(state.antraege.map((a) => String(yearOf(a.Antragsdatum))).filter(Boolean))].sort();
      state.years = new Set(years.slice(-3));
      state.ready = true;
      els.status.textContent =
        "Aktualisiert: " +
        new Date().toLocaleString("de-DE") +
        ` · ${state.antraege.length} Anträge, ${state.gutschriften.length} Gutschriften`;
      render();
      setupChat();
    } catch (err) {
      els.error.hidden = false;
      els.error.textContent = "Fehler: " + err.message;
      els.status.textContent = "";
    } finally {
      els.btnRefresh.disabled = false;
    }
  }

  els.btnRefresh.addEventListener("click", loadAll);
  els.chatForm.addEventListener("submit", (e) => {
    e.preventDefault();
    askChat(els.chatInput.value);
  });

  loadAll();
})();
