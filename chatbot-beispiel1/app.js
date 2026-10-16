(() => {
  const chat = document.getElementById("chat");
  const form = document.getElementById("form");
  const input = document.getElementById("input");
  const suggestions = document.getElementById("suggestions");

  const prompts = [
    "Was bietet Maurer Consulting?",
    "Wird bei Projekten immer dokumentiert?",
    "Brauche ich n8n?",
    "Was ist mit Excel-Auswertung?",
    "Werden meine Chats gespeichert?",
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
        "Probieren Sie z. B.: Leistungen, Excel, n8n, Kontakt oder Datenschutz.\n" +
        "Das Wissen können Sie in wissen.md anpassen."
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
        "Hallo! Ich bin das Demo-Chatbot von Maurer Consulting.\n" +
          "Meine Antworten kommen aus wissen.md (fiktives Beispielwissen).\n" +
          "Stellen Sie eine Frage oder tippen Sie einen Vorschlag unten an.",
        "bot"
      );
    })
    .catch((err) => {
      addMsg("Fehler beim Laden der Wissensdatei: " + err.message, "bot");
    });
})();
