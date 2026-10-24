# XL2DOC — Excel to Word Converter

A beautiful, fully browser-based Excel (`.xlsx` / `.xls`) to Word (`.docx`) converter. No server, no uploads — everything runs locally in your browser.

🔗 **[Live Demo](https://Dhruva-Hegde.github.io/excel-to-word-converter)**

---

## ✨ Features

- 📂 **Drag & drop** file upload
- 🌗 **Dark / Light mode** toggle (saved to localStorage)
- 📑 **Multi-sheet support** — choose which sheets to include
- 🎨 **4 table colour themes** — Ocean Blue, Forest Green, Royal Purple, Classic Gray
- 🗒️ **Custom header & footer** text
- 📄 **Summary / title page** generation
- 📐 **Portrait / Landscape** page orientation
- ⚡ **In-browser preview** before downloading
- ⬇️ **One-click `.docx` download**
- 🔒 **100% private** — your file never leaves your device

---

## 🚀 Usage

Just open `index.html` in any modern browser (Chrome, Edge, Firefox). No installation or build step required.

```
git clone https://github.com/Dhruva-Hegde/excel-to-word-converter.git
cd excel-to-word-converter
# Open index.html in your browser
```

---

## 🛠️ Tech Stack

| Library | Purpose |
|---------|---------|
| [SheetJS (xlsx)](https://sheetjs.com/) v0.18.5 | Parse Excel files in the browser |
| [docx.js](https://docx.js.org/) v7.8.2 | Generate `.docx` Word files |
| Vanilla HTML / CSS / JS | No framework needed |

---

## 📁 Project Structure

```
excel-to-word-converter/
├── index.html   # App structure & layout
├── style.css    # Theming, animations, glassmorphism UI
├── app.js       # Core logic (parsing, docx generation, preview)
└── .gitignore
```

---

## 📄 License

MIT © [Dhruva Hegde](https://github.com/Dhruva-Hegde)
