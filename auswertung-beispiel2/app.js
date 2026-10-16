(() => {
  const state = {
    rows: [],
    years: new Set(),
    region: "beide", // nord | sued | beide
  };

  const els = {
    yearBox: document.getElementById("years"),
    regionBox: document.getElementById("regions"),
    kpiUmsatz: document.getElementById("kpiUmsatz"),
    kpiTx: document.getElementById("kpiTx"),
    kpiAvg: document.getElementById("kpiAvg"),
    chart: document.getElementById("chart"),
    chartSum: document.getElementById("chartSum"),
    legend: document.getElementById("legend"),
    tableBody: document.getElementById("tableBody"),
    sourceLabel: document.getElementById("sourceLabel"),
  };

  function euro(n) {
    return n.toLocaleString("de-DE", { style: "currency", currency: "EUR", maximumFractionDigits: 0 });
  }

  function filtered() {
    return state.rows.filter((r) => {
      if (!state.years.has(String(r.Jahr))) return false;
      if (state.region === "nord" && r.Region !== "Nord") return false;
      if (state.region === "sued" && r.Region !== "Süd") return false;
      return true;
    });
  }

  function aggregateByYearRegion(rows) {
    const map = {};
    rows.forEach((r) => {
      const y = String(r.Jahr);
      if (!map[y]) map[y] = { Nord: 0, Süd: 0 };
      map[y][r.Region] += Number(r.Umsatz_EUR) || 0;
    });
    return map;
  }

  function renderChips() {
    const allYears = [...new Set(state.rows.map((r) => String(r.Jahr)))].sort();
    els.yearBox.innerHTML = "";
    allYears.forEach((y) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "chip" + (state.years.has(y) ? " active" : "");
      b.textContent = y;
      b.onclick = () => {
        if (state.years.has(y)) {
          if (state.years.size === 1) return;
          state.years.delete(y);
        } else {
          state.years.add(y);
        }
        render();
      };
      els.yearBox.appendChild(b);
    });

    els.regionBox.innerHTML = "";
    [
      ["beide", "Nord & Süd"],
      ["nord", "Nur Nord"],
      ["sued", "Nur Süd"],
    ].forEach(([id, label]) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "chip" + (state.region === id ? " active" : "");
      b.textContent = label;
      b.onclick = () => {
        state.region = id;
        render();
      };
      els.regionBox.appendChild(b);
    });
  }

  function renderKpis(rows) {
    const umsatz = rows.reduce((s, r) => s + (Number(r.Umsatz_EUR) || 0), 0);
    const tx = rows.reduce((s, r) => s + (Number(r.Transaktionen) || 0), 0);
    const avg = tx ? umsatz / tx : 0;
    els.kpiUmsatz.textContent = euro(umsatz);
    els.kpiTx.textContent = tx.toLocaleString("de-DE");
    els.kpiAvg.textContent = euro(avg);
  }

  function renderChart(rows) {
    const showNord = state.region === "beide" || state.region === "nord";
    const showSued = state.region === "beide" || state.region === "sued";
    const agg = aggregateByYearRegion(rows);
    const years = Object.keys(agg).sort();

    const padL = 52, padR = 16, padT = 20, padB = 36;
    const W = 720, H = 260;
    const innerW = W - padL - padR;
    const innerH = H - padT - padB;

    let maxVal = 1;
    let total = 0;
    years.forEach((y) => {
      if (showNord) { maxVal = Math.max(maxVal, agg[y].Nord); total += agg[y].Nord; }
      if (showSued) { maxVal = Math.max(maxVal, agg[y].Süd); total += agg[y].Süd; }
    });
    els.chartSum.textContent = "Summe: " + euro(total);

    const groupW = years.length ? innerW / years.length : innerW;
    const barsPerGroup = (showNord ? 1 : 0) + (showSued ? 1 : 0);
    const barW = Math.min(40, (groupW - 16) / Math.max(barsPerGroup, 1) - 4);

    let html = "";
    for (let i = 0; i <= 4; i++) {
      const y = padT + (innerH * i) / 4;
      const val = Math.round(maxVal * (1 - i / 4));
      html += `<line x1="${padL}" y1="${y}" x2="${W - padR}" y2="${y}" stroke="#e8e3f4" stroke-width="1"/>`;
      html += `<text x="${padL - 8}" y="${y + 4}" text-anchor="end" font-size="11" fill="#6b6275">${val.toLocaleString("de-DE")}</text>`;
    }

    years.forEach((year, gi) => {
      const gx = padL + gi * groupW + groupW / 2;
      let bi = 0;
      const drawBar = (value, color, label) => {
        const h = (value / maxVal) * innerH;
        const x = gx - (barsPerGroup * (barW + 4)) / 2 + bi * (barW + 4);
        const y = padT + innerH - h;
        html += `<rect x="${x}" y="${y}" width="${barW}" height="${Math.max(h, 0)}" rx="4" fill="${color}"><title>${label} ${year}: ${euro(value)}</title></rect>`;
        html += `<text x="${x + barW / 2}" y="${y - 6}" text-anchor="middle" font-size="10" font-weight="600" fill="#2e1477">${Math.round(value / 1000)}k</text>`;
        bi++;
      };
      if (showNord) drawBar(agg[year].Nord, "#2e1477", "Nord");
      if (showSued) drawBar(agg[year].Süd, "#8b7bc4", "Süd");
      html += `<text x="${gx}" y="${H - 12}" text-anchor="middle" font-size="12" fill="#2b2433">${year}</text>`;
    });

    els.chart.innerHTML = html;
    let leg = "";
    if (showNord) leg += `<span><span class="dot a"></span>Nord</span>`;
    if (showSued) leg += `<span><span class="dot b"></span>Süd</span>`;
    els.legend.innerHTML = leg;
  }

  function renderTable(rows) {
    const byFiliale = {};
    rows.forEach((r) => {
      const key = r.Filiale + "|" + r.Region;
      if (!byFiliale[key]) {
        byFiliale[key] = { Filiale: r.Filiale, Region: r.Region, Umsatz: 0, Tx: 0 };
      }
      byFiliale[key].Umsatz += Number(r.Umsatz_EUR) || 0;
      byFiliale[key].Tx += Number(r.Transaktionen) || 0;
    });
    const list = Object.values(byFiliale).sort((a, b) => b.Umsatz - a.Umsatz);
    els.tableBody.innerHTML = list
      .map(
        (r) => `<tr>
          <td>${r.Filiale}</td>
          <td>${r.Region}</td>
          <td>${euro(r.Umsatz)}</td>
          <td>${r.Tx.toLocaleString("de-DE")}</td>
        </tr>`
      )
      .join("");
  }

  function render() {
    renderChips();
    const rows = filtered();
    renderKpis(rows);
    renderChart(rows);
    renderTable(rows);
  }

  async function load() {
    const res = await fetch("data/beispieldaten.json");
    if (!res.ok) throw new Error("Daten konnten nicht geladen werden");
    const data = await res.json();
    state.rows = data.zeilen || [];
    const years = [...new Set(state.rows.map((r) => String(r.Jahr)))].sort();
    // Standard: letzte zwei Jahre
    state.years = new Set(years.slice(-2));
    if (els.sourceLabel && data.meta) {
      els.sourceLabel.textContent =
        (data.meta.quelle || "Beispieldaten.xlsx") + " · " + (data.meta.hinweis || "");
    }
    render();
  }

  load().catch((err) => {
    document.body.insertAdjacentHTML(
      "afterbegin",
      `<p style="color:#a00;padding:12px;">Fehler: ${err.message}. Bitte Seite über einen lokalen Server oder GitHub Pages öffnen (nicht als file://).</p>`
    );
  });
})();
