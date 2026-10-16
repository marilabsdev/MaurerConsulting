(() => {
  const chat = document.getElementById("chat");
  const form = document.getElementById("form");
  const input = document.getElementById("input");
  const suggestions = document.getElementById("suggestions");

  const prompts = [
    "Welcher Kinderwagen bei extrem unwegsamem Gelände?",
    "Welcher Kinderwagen bei kleinem Kofferraum?",
    "Welcher Autositz bei einem sehr schweren Kind?",
    "Autositz für BMW mit schräger Rückbank oder VW Bus?",
    "Mein Kind ist 105 cm, 4 Jahre, 18 kg – welcher Autositz?",
    "Was ist der Unterschied zwischen i-Size und der alten Norm?",
    "Baby mit Neurodermitis – welche Kleidung?",
    "In welchem Verhältnis Babymilch anmischen?",
    "Warum reicht Google bei Kindersitzen oft nicht?",
    "Wird bei Projekten immer dokumentiert?",
  ];

  // Freitext → Wissens-Abschnitt (ohne LLM: feste Themen-Zuordnung)
  const routes = [
    { id: "gelände", re: /gelände|unwegsam|wald|kies|feld|outdoor|wandern|jogger|luftreifen|thule|offroad|schotter/, titleHint: /gelände|unwegsam|thule/i },
    { id: "koffer", re: /koffer|klein|falt|packmaß|reise|urlaub|flug|lara|pockit|kompakt|platzspar/, titleHint: /koffer|klein|falt/i },
    { id: "schwer", re: /schwer|übergewicht|50\s*kg|36\s*kg|großes kind|dickes baby|stretch|pallas|solution/, titleHint: /schwer|stretch|50|36/i },
    { id: "spezialauto", re: /bmw|schräg|rückbank|vw\s*bus|multivan|t6|t5|isofix mitt|typenliste|besafe|level technology/, titleHint: /bmw|bus|schräg|speziell|typenliste|besafe/i },
    { id: "norm", re: /i-?size|r44|r129|norm|gewichtsklasse|körpergröße|105\s*cm|18\s*kg|jahre alt|welcher autositz/, titleHint: /norm|i-size|größe|alter|gewicht|105/i },
    { id: "haut", re: /neuro|haut|ausschlag|ekzem|atop|engel|gots|baumwolle|empfindlich|allerg/, titleHint: /neuro|haut|engel|gots|kleidung/i },
    { id: "milch", re: /milch|fläsch|anmisch|messlöffel|hipp|verhältnis|30\s*ml|pulver|nahrung/, titleHint: /milch|messlöffel|verhältnis|nahrung/i },
    { id: "google", re: /google|intern|firmenwissen|warum.*nicht|spezialist/, titleHint: /google|firmenwissen|intern/i },
    { id: "doku", re: /dokument|ausfall|abhängig|it-firma|weiterarbeit/, titleHint: /dokument|ausfall/i },
  ];

  let sections = [];

  function addMsg(text, who) {
    const div = document.createElement("div");
    div.className = "msg " + who;
    div.textContent = text;
    chat.appendChild(div);
    chat.scrollTop = chat.scrollHeight;
  }

  function parseMarkdown(md) {
    const parts = md.split(/\n(?=## )/);
    const out = [];
    parts.forEach((block) => {
      const lines = block.trim().split("\n");
      if (!lines.length) return;
      const title = lines[0].replace(/^#+\s*/, "").trim();
      const body = lines.slice(1).join("\n").trim();
      if (!title) return;
      out.push({ title, body, hay: (title + "\n" + body).toLowerCase() });
    });
    return out;
  }

  function findSectionByTitleHint(hint) {
    return sections.find((s) => hint.test(s.title) || hint.test(s.hay));
  }

  function answer(query) {
    const q = query.toLowerCase();

    // 1) Feste Themen-Routen (damit Freitext nicht „wild“ mischt)
    for (const route of routes) {
      if (route.re.test(q)) {
        const sec = findSectionByTitleHint(route.titleHint);
        if (sec) {
          return sec.title + "\n\n" + sec.body.replace(/\n{3,}/g, "\n\n").trim();
        }
      }
    }

    // 2) Fallback: bestes einzelnes Kapitel nach Wort-Score (nur 1 Treffer, Mindestscore)
    const words = q.replace(/[?!.,]/g, " ").split(/\s+/).filter((w) => w.length > 3);
    let best = null;
    let bestScore = 0;
    sections.forEach((sec) => {
      let s = 0;
      words.forEach((w) => {
        if (sec.hay.includes(w)) s += 1;
        if (sec.title.toLowerCase().includes(w)) s += 2;
      });
      if (s > bestScore) {
        bestScore = s;
        best = sec;
      }
    });

    if (best && bestScore >= 2) {
      return best.title + "\n\n" + best.body.replace(/\n{3,}/g, "\n\n").trim();
    }

    return (
      "Dazu habe ich in der Wissensdatei keinen klaren Treffer.\n\n" +
      "Hinweis: In dieser Demo steckt kein Sprachmodell (LLM), sondern Firmenwissen aus wissen.md mit Themen-Zuordnung.\n" +
      "Bitte Shortcut nutzen oder z. B. schreiben: „Gelände“, „Kofferraum“, „BMW“, „Neurodermitis“, „Babymilch“, „Dokumentation“."
    );
  }

  function ask(q) {
    const text = (q || "").trim();
    if (!text) return;
    addMsg(text, "user");
    addMsg(answer(text), "bot");
    input.value = "";
  }

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    ask(input.value);
  });

  prompts.forEach((p) => {
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = p;
    b.onclick = () => ask(p);
    suggestions.appendChild(b);
  });

  fetch("wissen.md")
    .then((r) => {
      if (!r.ok) throw new Error("wissen.md nicht gefunden");
      return r.text();
    })
    .then((md) => {
      sections = parseMarkdown(md);
      addMsg(
        "Hallo! Demo-Fachchatbot mit Babymarkt-Firmenwissen.\n" +
          "Kein LLM — Antworten kommen aus wissen.md (Themen-Matching).\n" +
          "Shortcut antippen oder frei tippen (z. B. „Waldweg Kinderwagen“, „Babymilch mischen“).",
        "bot"
      );
    })
    .catch((err) => {
      addMsg("Fehler beim Laden der Wissensdatei: " + err.message, "bot");
    });
})();
