/* Code Slides — add-in de PowerPoint (Office.js, sin dependencias) */

// ---------- Temas ----------
const THEMES = {
  "Oscuro (VS Code)": { slide:"#1E1E1E", bg:"#1E1E1E", bar:"#2D2D2D", fg:"#D4D4D4", ln:"#6E7681", kw:"#569CD6", str:"#CE9178", com:"#6A9955", num:"#B5CEA8", fn:"#DCDCAA", cmd:"#4EC9B0", flag:"#9CDCFE", op:"#C586C0", var:"#9CDCFE" },
  "Dracula":          { slide:"#282A36", bg:"#282A36", bar:"#21222C", fg:"#F8F8F2", ln:"#6272A4", kw:"#FF79C6", str:"#F1FA8C", com:"#6272A4", num:"#BD93F9", fn:"#50FA7B", cmd:"#8BE9FD", flag:"#FFB86C", op:"#FF79C6", var:"#FFB86C" },
  "Monokai":          { slide:"#272822", bg:"#272822", bar:"#1E1F1A", fg:"#F8F8F2", ln:"#75715E", kw:"#F92672", str:"#E6DB74", com:"#75715E", num:"#AE81FF", fn:"#A6E22E", cmd:"#66D9EF", flag:"#FD971F", op:"#F92672", var:"#FD971F" },
  "Nord":             { slide:"#2E3440", bg:"#2E3440", bar:"#3B4252", fg:"#D8DEE9", ln:"#616E88", kw:"#81A1C1", str:"#A3BE8C", com:"#616E88", num:"#B48EAD", fn:"#88C0D0", cmd:"#8FBCBB", flag:"#EBCB8B", op:"#81A1C1", var:"#D08770" },
  "Claro":            { slide:"#FFFFFF", bg:"#FFFFFF", bar:"#EEF0F4", fg:"#24292F", ln:"#8C959F", kw:"#CF222E", str:"#0A3069", com:"#6E7781", num:"#0550AE", fn:"#8250DF", cmd:"#116329", flag:"#953800", op:"#CF222E", var:"#953800" }
};

// ---------- Lenguajes (reglas en orden de prioridad; sin grupos de captura) ----------
const STR_C = `"(?:\\\\.|[^"\\\\\\n])*"|'(?:\\\\.|[^'\\\\\\n])*'`;
const words = (s) => `\\b(?:${s})\\b`;

const CLI_COMMON = [
  ["com", `#.*`],
  ["str", STR_C],
  ["var", `\\$(?:\\{[^}\\n]*\\}|\\w+|\\()`],
  ["flag", `(?<![\\w-])--?[A-Za-z][\\w-]*`],
  ["op", `\\|\\||&&|\\||;|>>?|<|\\\\$`]
];

const LANGS = {
  "Python": { ci:false, rules: [
    ["com", `#.*`],
    ["str", `[rbfRBF]{0,2}(?:"""[\\s\\S]*?"""|'''[\\s\\S]*?'''|"(?:\\\\.|[^"\\\\\\n])*"|'(?:\\\\.|[^'\\\\\\n])*')`],
    ["op", `@\\w+(?:\\.\\w+)*`],
    ["kw", words(`and|as|assert|async|await|break|class|continue|def|del|elif|else|except|finally|for|from|global|if|import|in|is|lambda|nonlocal|not|or|pass|raise|return|try|while|with|yield|None|True|False|self`)],
    ["cmd", words(`print|len|range|str|int|float|list|dict|set|tuple|open|type|isinstance|enumerate|zip|map|filter|sorted|sum|min|max|abs|input`)],
    ["fn", `\\b[A-Za-z_]\\w*(?=\\()`],
    ["num", `\\b\\d+(?:\\.\\d+)?\\b`]
  ]},
  "JavaScript": { ci:false, rules: [
    ["com", `\\/\\/.*|\\/\\*[\\s\\S]*?\\*\\/`],
    ["str", `\`(?:\\\\[\\s\\S]|[^\`\\\\])*\`|${STR_C}`],
    ["kw", words(`async|await|break|case|catch|class|const|continue|default|delete|do|else|export|extends|finally|for|from|function|if|import|in|instanceof|let|new|of|return|static|super|switch|this|throw|try|typeof|var|void|while|yield|null|undefined|true|false`)],
    ["cmd", words(`console|document|window|Promise|Array|Object|JSON|Math|Map|Set|Date`)],
    ["op", `=>`],
    ["fn", `\\b[A-Za-z_$][\\w$]*(?=\\()`],
    ["num", `\\b\\d+(?:\\.\\d+)?\\b`]
  ]},
  "SQL": { ci:true, rules: [
    ["com", `--.*|\\/\\*[\\s\\S]*?\\*\\/`],
    ["str", `'(?:''|[^'])*'`],
    ["kw", words(`select|from|where|and|or|not|in|is|null|like|between|exists|join|inner|left|right|full|outer|cross|on|group|by|order|having|limit|top|offset|union|all|distinct|as|insert|into|values|update|set|delete|create|alter|drop|table|view|index|database|schema|primary|key|foreign|references|constraint|default|unique|check|case|when|then|else|end|with|begin|commit|rollback|transaction|asc|desc|if|procedure|function|declare|exec|truncate|grant|revoke`)],
    ["cmd", words(`int|integer|bigint|smallint|varchar|nvarchar|char|text|date|datetime|timestamp|decimal|numeric|float|boolean|bit|number|varchar2`)],
    ["fn", `\\b[A-Za-z_]\\w*(?=\\()`],
    ["num", `\\b\\d+(?:\\.\\d+)?\\b`]
  ]},
  "Git": { ci:false, rules: [
    ...CLI_COMMON,
    ["cmd", `\\bgit\\b`],
    ["fn", `(?<=\\bgit\\s+)[a-z][\\w-]*`]
  ]},
  "Bash (Linux)": { ci:false, rules: [
    ...CLI_COMMON,
    ["kw", words(`if|then|else|elif|fi|for|while|until|do|done|case|esac|function|in|return|exit|export|local|source|alias|set|unset|readonly|select`)],
    ["cmd", words(`sudo|apt|apt-get|dnf|yum|ls|cd|pwd|cat|grep|sed|awk|chmod|chown|mkdir|rm|cp|mv|ln|curl|wget|tar|ssh|scp|systemctl|journalctl|docker|find|echo|touch|ps|kill|df|du|tail|head|sort|uniq|xargs|tee|nano|vim|ping|ip|ss|ufw|crontab|useradd|passwd|man|which|whoami|git`)],
    ["num", `\\b\\d+\\b`]
  ]},
  "Azure CLI": { ci:false, rules: [
    ...CLI_COMMON,
    ["cmd", `\\baz\\b`],
    ["fn", `(?<=\\baz\\s+)[a-z][\\w-]*(?:[ \\t]+[a-z][\\w-]*){0,2}`],
    ["num", `\\b\\d+\\b`]
  ]},
  "AWS CLI": { ci:false, rules: [
    ...CLI_COMMON,
    ["cmd", `\\baws\\b`],
    ["fn", `(?<=\\baws\\s+)[a-z0-9][\\w-]*(?:[ \\t]+[a-z0-9][\\w-]*)?`],
    ["num", `\\b\\d+\\b`]
  ]}
};

const SAMPLES = {
  "Python": `def saludar(nombre: str) -> str:\n    # Devuelve un saludo\n    return f"Hola, {nombre}"\n\nfor i in range(3):\n    print(saludar("Dave"), i)`,
  "JavaScript": `// Suma con async/await\nasync function cargar(url) {\n  const res = await fetch(url);\n  return res.json();\n}\nconsole.log(42);`,
  "SQL": `-- Top 5 clientes\nSELECT TOP 5 c.nombre, COUNT(*) AS pedidos\nFROM clientes c\nJOIN pedidos p ON p.cliente_id = c.id\nWHERE p.fecha >= '2026-01-01'\nGROUP BY c.nombre\nORDER BY pedidos DESC;`,
  "Git": `git checkout -b feature/login\ngit add .\ngit commit -m "Agrega login"\ngit push --set-upstream origin feature/login`,
  "Bash (Linux)": `#!/bin/bash\n# Respaldo diario\nFECHA=$(date +%F)\nsudo tar -czf /backup/$FECHA.tar.gz /var/www | tee backup.log\nif [ $? -eq 0 ]; then echo "OK"; fi`,
  "Azure CLI": `az group create --name rg-demo --location eastus\naz vm create --resource-group rg-demo --name vm1 \\\n  --image Ubuntu2204 --admin-username dave --generate-ssh-keys`,
  "AWS CLI": `aws s3 mb s3://mi-bucket-demo --region us-east-1\naws ec2 describe-instances --query "Reservations[].Instances[].InstanceId" --output text`
};

// ---------- Tokenizador ----------
function tokenize(code, lang) {
  const L = LANGS[lang];
  const re = new RegExp(L.rules.map(r => `(${r[1]})`).join("|"), "gm" + (L.ci ? "i" : ""));
  const out = [];
  let m;
  while ((m = re.exec(code)) !== null) {
    if (m[0].length === 0) { re.lastIndex++; continue; }
    const type = L.rules[m.slice(1).findIndex(x => x !== undefined)][0];
    const last = out[out.length - 1];
    if (last && last.type === type && last.start + last.len === m.index) last.len += m[0].length;
    else out.push({ start: m.index, len: m[0].length, type });
  }
  return out;
}

// ---------- UI ----------
const $ = (id) => document.getElementById(id);
const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

function normalize(raw) {
  return raw.replace(/\r\n?/g, "\n").replace(/\t/g, "    ").replace(/\n+$/, "");
}

function renderPreview() {
  const code = normalize($("code").value), T = THEMES[$("theme").value];
  let html = "", pos = 0;
  for (const t of tokenize(code, $("lang").value)) {
    html += esc(code.slice(pos, t.start)) + `<span style="color:${T[t.type]}">${esc(code.substr(t.start, t.len))}</span>`;
    pos = t.start + t.len;
  }
  html += esc(code.slice(pos));
  const p = $("prev");
  p.style.background = T.bg; p.style.color = T.fg; p.innerHTML = html;
}

function setMsg(text, err) { const m = $("msg"); m.textContent = text; m.className = err ? "err" : ""; }

function init() {
  Object.keys(LANGS).forEach(l => $("lang").add(new Option(l, l)));
  Object.keys(THEMES).forEach(t => $("theme").add(new Option(t, t)));
  $("code").value = SAMPLES["Python"];
  let lastSample = SAMPLES["Python"];
  $("lang").addEventListener("change", () => {
    // Solo reemplaza el texto si sigue siendo el ejemplo sin tocar
    if ($("code").value === lastSample) { lastSample = SAMPLES[$("lang").value]; $("code").value = lastSample; }
    renderPreview();
  });
  ["theme", "code"].forEach(id => $(id).addEventListener("input", renderPreview));
  $("go").addEventListener("click", insertCode);
  renderPreview();
}

// ---------- Inserción en la diapositiva ----------
async function insertCode() {
  const lang = $("lang").value, T = THEMES[$("theme").value];
  const code = normalize($("code").value);
  if (!code.trim()) return setMsg("Pega o escribe algo de código primero.", true);

  const size = Math.min(40, Math.max(8, Number($("size").value) || 16));
  const start = Number($("start").value) || 1;
  const withLn = $("optLn").checked, withBar = $("optBar").checked, withBg = $("optBg").checked;
  const lines = code.split("\n");
  const longest = Math.max(...lines.map(l => l.length));

  $("go").disabled = true; setMsg("Insertando…");
  try {
    await PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.getSelectedSlides().getItemAt(0);

      // Tamaño de diapositiva (si la API no lo ofrece, se asume 16:9 = 960x540 pt)
      let SW = 960, SH = 540;
      try {
        const ps = ctx.presentation.pageSetup; ps.load("slideWidth,slideHeight"); await ctx.sync();
        SW = ps.slideWidth; SH = ps.slideHeight;
      } catch (_) { /* usa valores por defecto */ }

      const lineH = size * 1.2;                    // estimación del interlineado sencillo
      const charW = size * 0.55;                   // Consolas ≈ 0.55 em por carácter
      const pad = 16, barH = withBar ? 30 : 0;
      const lnW = withLn ? Math.max(2, String(start + lines.length - 1).length) * charW + 12 : 0;

      const cardW = Math.min(SW - 60, lnW + longest * charW + pad * 2 + 20);
      const cardH = barH + pad * 2 + lines.length * lineH;
      const cardL = (SW - cardW) / 2;
      const cardT = Math.max(20, (SH - cardH) / 2);

      const shapes = slide.shapes;
      const box = (type, l, t, w, h, color, name) => {
        const s = shapes.addGeometricShape(type, { left: l, top: t, width: w, height: h });
        s.fill.setSolidColor(color); s.lineFormat.visible = false; s.name = name; return s;
      };

      if (withBg) {
        const bg = box("Rectangle", 0, 0, SW, SH, T.slide, "CS_SlideBg");
        try { bg.setZOrder("SendToBack"); } catch (_) {}
      }
      box("Rectangle", cardL, cardT, cardW, cardH, T.bg, "CS_Card");
      box("Rectangle", cardL, cardT, 4, cardH, T.kw, "CS_Accent");

      if (withBar) {
        box("Rectangle", cardL + 4, cardT, cardW - 4, barH, T.bar, "CS_Bar");
        ["#FF5F56", "#FFBD2E", "#27C93F"].forEach((c, i) =>
          box("Ellipse", cardL + 18 + i * 18, cardT + (barH - 10) / 2, 10, 10, c, "CS_Dot" + (i + 1)));
        const lbl = shapes.addTextBox(lang, { left: cardL + cardW - 160, top: cardT + 4, width: 148, height: barH - 8 });
        lbl.name = "CS_Label";
        lbl.textFrame.wordWrap = false;
        lbl.textFrame.verticalAlignment = "Middle";
        lbl.textFrame.textRange.font.name = "Segoe UI";
        lbl.textFrame.textRange.font.size = 11;
        lbl.textFrame.textRange.font.color = T.ln;
        lbl.textFrame.textRange.paragraphFormat.horizontalAlignment = "Right";
      }

      const textTop = cardT + barH + pad;
      const mkText = (txt, l, w, name) => {
        const s = shapes.addTextBox(txt, { left: l, top: textTop, width: w, height: lines.length * lineH + 4 });
        s.name = name;
        const tf = s.textFrame;
        tf.wordWrap = false; tf.autoSizeSetting = "AutoSizeNone"; tf.verticalAlignment = "Top";
        tf.leftMargin = 0; tf.rightMargin = 0; tf.topMargin = 0; tf.bottomMargin = 0;
        tf.textRange.font.name = "Consolas"; tf.textRange.font.size = size;
        return s;
      };

      if (withLn) {
        const nums = lines.map((_, i) => String(start + i)).join("\n");
        const ln = mkText(nums, cardL + 4 + pad, lnW, "CS_LineNumbers");
        ln.textFrame.textRange.font.color = T.ln;
        ln.textFrame.textRange.paragraphFormat.horizontalAlignment = "Right";
      }

      const codeLeft = cardL + 4 + pad + (withLn ? lnW + 14 : 0);
      const cs = mkText(code, codeLeft, cardW - (codeLeft - cardL) - pad, "CS_Code");
      const tr = cs.textFrame.textRange;
      tr.font.color = T.fg;
      for (const t of tokenize(code, lang)) tr.getSubstring(t.start, t.len).font.color = T[t.type];

      await ctx.sync();
    });
    setMsg("Listo. Puedes mover o editar las formas; todas empiezan con “CS_”.");
  } catch (e) {
    setMsg("Error: " + (e && e.message ? e.message : e), true);
  } finally {
    $("go").disabled = false;
  }
}

Office.onReady(() => init());
