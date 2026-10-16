(() => {
  const chat = document.getElementById("chat");
  const form = document.getElementById("form");
  const input = document.getElementById("input");
  const suggestions = document.getElementById("suggestions");

  // Maximal 10 Shortcuts — Fokus Firmenwissen Babymarkt + Doku
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
      let title = lines[0].replace(/^#+\s*/, "").trim();
      const body = lines.slice(1).join("\n").trim();
      if (!title) return;
      out.push({ title, body, hay: (title + "\n" + body).toLowerCase() });
    });
    return out;
  }

  function score(query, section) {
    const q = query.toLowerCase().replace(/[?!.,]/g, " ");
    const words = q.split(/\s+/).filter((w) => w.length > 2);
    if (!words.length) return 0;
    let s = 0;
    words.forEach((w) => {
      if (section.hay.includes(w)) s += 1;
      if (section.title.toLowerCase().includes(w)) s += 2;
    });
    // Themen-Boosts
    if (/gelände|unwegsam|wald|kies|thule|jogger/.test(q) && /gelände|thule|unwegsam/.test(section.hay)) s += 6;
    if (/koffer|klein|falt|lara|pockit|reisebuggy/.test(q) && /koffer|lara|falt/.test(section.hay)) s += 6;
    if (/schwer|50\s*kg|36\s*kg|stretch|pallas|solution/.test(q) && /schwer|stretch|50|36/.test(section.hay)) s += 6;
    if (/bmw|schräg|vw|bus|multivan|t6|isofix|besafe|level/.test(q) && /bmw|bus|schräg|typenliste|besafe/.test(section.hay)) s += 8;
    if (/105|cm|kg|jahre|alter|größe|norm|i-size|r44|r129/.test(q) && /norm|i-size|größe|r44|r129|105/.test(section.hay)) s += 6;
    if (/neuro|haut|ausschlag|engel|gots|baumwolle|ekzem/.test(q) && /neuro|engel|haut|gots/.test(section.hay)) s += 6;
    if (/milch|mischen|messlöffel|fläsch|hipp|verhältnis|30\s*ml/.test(q) && /milch|messlöffel|30|hipp/.test(section.hay)) s += 6;
    if (/google|intern|firmenwissen|spezial/.test(q) && /google|firmenwissen|typenliste/.test(section.hay)) s += 5;
    if (/dokument|ausfall|abhängig/.test(q) && /dokument|ausfall/.test(section.hay)) s += 6;
    return s;
  }

  function answer(query) {
    const ranked = sections
      .map((sec) => ({ sec, s: score(query, sec) }))
      .filter((x) => x.s > 0)
      .sort((a, b) => b.s - a.s);

    if (!ranked.length) {
      return (
        "Dazu steht in der Demo-Wissensdatei nichts Passendes.\n\n" +
        "Nutzen Sie einen der Vorschläge unten — z. B. Gelände-Kinderwagen, Kofferraum, Autositz-Normen, Neurodermitis oder Babymilch."
      );
    }

    const top = ranked.slice(0, 2);
    return top
      .map((x) => {
        const body = x.sec.body.replace(/\n{3,}/g, "\n\n").trim();
        return x.sec.title + "\n\n" + (body || "(kein Text)");
      })
      .join("\n\n—\n\n");
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
        "Hallo! Ich bin der Demo-Fachchatbot mit Babymarkt-Firmenwissen.\n" +
          "Fragen zu Kinderwagen, Autositzen (inkl. Spezialfahrzeuge), Haut/Kleidung und Babymilch — " +
          "plus warum Dokumentation bei IT-Projekten wichtig ist.\n" +
          "Tippen Sie einen Shortcut an oder schreiben Sie frei.",
        "bot"
      );
    })
    .catch((err) => {
      addMsg("Fehler beim Laden der Wissensdatei: " + err.message, "bot");
    });
})();
