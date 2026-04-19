# 🌟 Lumina Docs Suite

> A Google Docs sidebar add-on built with **TypeScript + Google Apps Script** that brings smart text formatting, find & replace, and URL health checking directly inside your document.

![Lumina Docs Suite](https://github.com/user-attachments/assets/0ea9db99-356c-4805-ac74-4eb39886ef16)

![Lumina Docs Suite](https://github.com/user-attachments/assets/99fba495-82f2-4d49-8bc9-0a41771b3c18)
---

## ✨ Features

| Feature | Description |
|---|---|
| 🔍 **Find & Replace** | Search across the entire document with match counter (e.g. `5 of 18`) and Prev/Next navigation |
| 🎨 **Format This** | Apply bold, italic, or highlight color to the currently selected match only |
| ⚡ **Format All** | Apply formatting to every occurrence of the searched term in one click |
| ↩️ **Undo** | Revert the last formatting action instantly |
| 🔗 **Check URL Health** | Scan all hyperlinks in the document and report their HTTP status (✅ 200 / ❌ 404) |

---

## 📸 Demo

> **Full workflow:** `clasp open → Deploy → Test Deployments → Editor Add-on → Attach Document → Extensions → Lumina Docs Suite → Open Toolbar`

<!-- Replace with your Loom/YouTube link when ready -->
🎥 **[Watch Demo Video](#)**

---

## 🛠️ Tech Stack

- **Language:** TypeScript
- **Runtime:** Google Apps Script (V8 engine)
- **CLI Tool:** Clasp v2.4.2
- **IDE:** VS Code
- **Version Control:** Git + GitHub

---

## 📁 Project Structure

```
lumina-docs-suite/
├── src/
│   ├── Main.ts              # Entry point — exposes functions to Apps Script
│   ├── LinkChecker.ts       # URL health check logic using UrlFetchApp
│   ├── StyleEngine.ts       # Formatting logic (bold, italic, highlight)
│   ├── Sidebar.html         # Sidebar UI rendered inside Google Docs
│   └── appsscript.json      # OAuth scopes and Apps Script manifest
├── .clasp.json              # Links local project to Apps Script cloud project
├── tsconfig.json            # TypeScript compiler config (target: ES2019)
├── package.json             # npm scripts and dev dependencies
└── .gitignore
```

---

## 🚀 Getting Started

### Prerequisites

- [Node.js](https://nodejs.org/) v16+
- [Clasp](https://github.com/google/clasp) v2.4.2
- A Google account

### Installation

**1. Clone the repository**
```bash
git clone https://github.com/your-username/lumina-docs-suite.git
cd lumina-docs-suite
```

**2. Install dependencies**
```bash
npm install
```

**3. Install Clasp globally (use v2.4.2 specifically)**
```bash
npm install -g @google/clasp@2.4.2
```

**4. Login to Google**
```bash
clasp login
```

**5. Link to your Apps Script project**

Update `.clasp.json` with your own `scriptId`:
```json
{
  "scriptId": "YOUR_SCRIPT_ID_HERE",
  "rootDir": "./src",
  "fileExtension": "ts"
}
```

**6. Push to Apps Script**
```bash
clasp push
```

---

## 🧪 Running as a Test Deployment

1. Run `clasp open` to open the project in the Apps Script browser editor
2. Click **Deploy → Test deployments**
3. Select type: **Editor Add-on**
4. Click **Add test** and attach a Google Doc
5. Click **Execute / Save**
6. Open the attached Google Doc
7. Go to **Extensions → Lumina Docs Suite → Open Toolbar**

---

## ⚙️ Clasp Commands Used

| Command | Purpose |
|---|---|
| `clasp login` | Authenticate with Google account |
| `clasp push` | Compile TypeScript and push to Apps Script |
| `clasp push -f` | Force push without confirmation prompt |
| `clasp open` | Open the Apps Script project in browser |
| `clasp deploy` | Create a versioned deployment |
| `clasp logs` | Stream execution logs to terminal |

---

## 🔐 OAuth Scopes

Defined in `appsscript.json`:

```json
"oauthScopes": [
  "https://www.googleapis.com/auth/documents",
  "https://www.googleapis.com/auth/script.external_request",
  "https://www.googleapis.com/auth/script.container.ui"
]
```

| Scope | Reason |
|---|---|
| `auth/documents` | Read and write document content |
| `auth/script.external_request` | Make HTTP requests for URL health checking |
| `auth/script.container.ui` | Render the sidebar UI inside Google Docs |

---

## 🗺️ Planned Improvements

- [ ] Case-sensitive search toggle
- [ ] Match counter preview before formatting
- [ ] Multi-keyword formatting in one pass
- [ ] Export URL health report to Google Sheets
- [ ] Filter broken links only (4xx / 5xx)
- [ ] Regex-based search support

---

## 🤝 Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

---

## 📄 License

[MIT](LICENSE)

---

<p align="center">Built with ❤️ using TypeScript + Google Apps Script</p>
