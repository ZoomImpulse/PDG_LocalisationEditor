# PDG Translate — Google Apps Script

A translation tool for the PDG localisation spreadsheets. Zero hosting required — runs entirely inside Google Sheets.

## Quick Start (GitHub + clasp)

1. **Clone this repo**
   ```bash
   git clone https://github.com/YOUR_ORG/PDG_LocalisationSidebar.git
   cd PDG_LocalisationSidebar
   ```

2. **Install clasp** (Google Apps Script CLI)
   ```bash
   npm install -g @google/clasp
   clasp login
   ```

3. **Configure your Script ID**
   ```bash
   cp .clasp.json.example .clasp.json
   # Edit .clasp.json and replace YOUR_LIBRARY_SCRIPT_ID_HERE
   ```

4. **Configure stub IDs** in `deploy.sh` (or `deploy.ps1` on Windows):
   - Open the script and add your 5 stub spreadsheet Script IDs to `STUB_SCRIPT_IDS`

5. **Deploy**
   - Linux/Mac: `./deploy.sh`
   - Windows: `deploy.ps1`

## Files

| File | Purpose |
|------|---------|
| `Config.gs` | Constants (status values, colors, helpers, linked spreadsheet IDs) |
| `Code.gs` | Core logic: column detection, data read/write, modal dialog |
| `StatusStorage.gs` | Status tracking in a separate hidden sheet |
| `ProgressCache.gs` | Cached progress data, refreshed every 10 minutes |
| `Overview.gs` | Auto-generated `📊 Overview` sheet with heat-map and global summary |
| `Roles.gs` | Role-based access control |
| `Modal.html` | **Main translator UI** — large modal dialog (1200×800px) |
| `Stub.gs` | **Thin wrapper** — only file needed in each spreadsheet's bound script |
| `deploy.sh` / `deploy.ps1` | Automated deployment scripts |
| `.clasp.json.example` | Template for clasp configuration |

## Deployment — Shared Library (recommended)

All real logic lives in a single **PDGLib** Apps Script Library project.
Each spreadsheet contains only a tiny `Stub.gs` that delegates to it.
When you update the library and publish a new version, all 6 spreadsheets
get the change instantly — no copy-pasting required.

### Step 1 — Deploy the main project as a library (one-time)

The library already exists — it is the **`LocalisationMasterTool`** Apps Script
project bound to the main spreadsheet.

1. Open the main spreadsheet → **Extensions → Apps Script**.
2. **Deploy → New deployment → Type: Library** → Deploy.
3. Copy the **Script ID** from **Project Settings (⚙)** — you will need it below.

### Step 2 — Add the stub to each of the 5 other spreadsheets (one-time)

1. Open the spreadsheet → **Extensions → Apps Script**.
2. Delete all existing `.gs` files and `Modal.html`.
3. Paste the contents of `Stub.gs` into the remaining (or new) script file.
4. Add the library:
   - Click **⊕** next to **Libraries** in the left panel.
   - Paste the **Script ID** from Step 1.
   - Select the latest numbered version (or **Development** for testing).
   - Set the **Identifier** to exactly `PDGLib`.
   - Click **Add**.
5. Save and reload the spreadsheet.
6. From the menu bar: **PDG Localisation → Setup: Seed Statuses**, then
   **PDG Localisation → Setup: Initialise Roles**.
7. Repeat for each of the remaining spreadsheets.

### Updating the script (ongoing)

**Automated (recommended):**

Edit files in this repo, then run the deploy script:

```bash
# Linux/Mac
./deploy.sh

# Windows PowerShell
.\deploy.ps1
```

This pushes the library and auto-updates all stub spreadsheets to use the new version.

**Manual:**

1. Edit files and push changes: `clasp push`
2. **Deploy → Manage deployments** → create a new numbered version.
3. In each stub spreadsheet's Apps Script editor, bump the version under **Libraries → PDGLib**
   — *or* use **Development mode** during active development so they always pick up the latest saved code automatically.

---

## Setup (legacy — manual copy-paste)

> Use this only if you cannot use the shared library approach above.

1. Open one of the 6 localisation spreadsheets in Google Sheets.
2. Go to **Extensions → Apps Script**.
3. Delete any existing `Code.gs` content and paste in each file:
   - Create a new script file for each `.gs` file (`+` → Script).
   - Create an HTML file (`+` → HTML) for `Modal.html`.
4. Save and reload the spreadsheet.
5. From the menu bar, click **PDG Localisation → Setup: Seed Statuses**.
   - This creates a hidden `📋 Status` sheet and seeds it with statuses from existing cell colors.
6. Repeat for each of the 6 spreadsheets.

## For Translators

1. Open the spreadsheet link (any Google account with edit access).
2. Click **PDG Translate → Open Translator** in the menu bar.
3. A large **modal dialog** opens (1200×800px), defaulting to **Untranslated only** filter.
4. Select your language and (if applicable) the column to translate via the tabs.
5. Type your translation — the status auto-sets to **Translated**.
6. Press **Ctrl+Enter** (or click **Save**) to save.
7. Use **← Prev / Next →** to navigate, or **↓ Next untranslated** to jump ahead.

### Modal Layout

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  [← Sheets]  PDG Localisation  [Translator ▾]         [👤 Admin]  [? Help] │  ← Header
├─────────────────────────────────────────────────────────────────────────────┤
│  👁 Read-only banner (hidden unless role = NONE)                            │  ← Banner
├──────────────────────────────────────────────────────┬──────────────────────┤
│  Sheet: [Sheet1 ▼]   Language: [German ▼]            │                      │
│  Search: [__________________] [Search] [Aa] [Clear]  │   Search Results     │
│  Show: [Untranslated] [All] [Translated] [Reviewed]  │   (collapsible       │
│  Fields: [Description] [Name] [Adjective]            │    right panel)      │
│                                                      │                      │
│  KEY                                                 │  ── or ──            │
│  building_army_01                                    │                      │
│                                                      │   Admin Panel        │
│  ENGLISH SOURCE                                      │   (collapsible       │
│  Army Headquarters                                   │    right panel)      │
│                                                      │                      │
│  TRANSLATION                                         │                      │
│  ┌──────────────────────────────────────────────┐    │                      │
│  │                                              │    │                      │
│  │   Large translation textarea                 │    │                      │
│  │                                              │    │                      │
│  └──────────────────────────────────────────────┘    │                      │
│                                                      │                      │
│  STATUS: [TRANSLATED ▾]                              │                      │
│  [Untranslated] [Translated Ctrl+S] [Reviewed Ctrl+E] [Disputed Ctrl+D]     │
├──────────────────────────────────────────────────────┴──────────────────────┤
│  [← Prev]  [↓ Next untranslated]  [Next →]       42 / 150    [Save]         │  ← Footer
└─────────────────────────────────────────────────────────────────────────────┘
```

- **Header**: Back-to-sheet-picker button, title, role badge, admin panel toggle, help tour button
- **Read-only banner**: Shown when the user has no translator role
- **Controls**: Sheet selector, language selector, cross-sheet search, status filter buttons, column field tabs
- **Translation panel**: Current key, English source, large translation textarea, status buttons with keyboard shortcuts
- **Search results panel**: Slides in from the right when a search is active
- **Admin panel**: Slides in from the right (admins only) — manage user roles
- **Footer**: Prev/Next navigation, entry counter, Save button

## 📊 Overview Sheet

The main spreadsheet automatically generates a `📊 Overview` tab showing translation progress across all spreadsheets.

### Sections

- **Global Overview** — one row per linked spreadsheet, reviewed % per language (main spreadsheet only)
- **Status Legend** — colour key with descriptions
- **Overall Progress per Language** — aggregated counts + sparkline progress bars
- **Progress by Sheet** — heat-map grid of reviewed % per sheet × language

### Configuration (`Config.gs`)

| Variable | Purpose |
|----------|---------|
| `MAIN_SPREADSHEET_ID` | ID of the main spreadsheet — only this one renders the Global Overview section |
| `LINKED_SPREADSHEETS` | Array of `{ name, id }` entries for all linked spreadsheets |

The overview refreshes automatically every 10 minutes (tied to the progress cache trigger) and can be manually triggered via **PDG Localisation → Refresh Overview**.

The sheet is **read-only** — it is hard-locked and cannot be edited manually.

### Stub update required

The `Stub.gs` in each linked spreadsheet must include the `refreshOverviewSheet` delegate for the menu item to appear. Re-paste the latest `Stub.gs` if updating from an older version.

---

## Status Colors

| Color | Status | Meaning |
|-------|--------|---------|
| `#d9d9d9` grey | STATIC | No translation needed |
| `#e06666` light red | UNTRANSLATED | Needs translation |
| `#ffff00` yellow | TRANSLATED | Done, awaiting review |
| `#00ff00` green | REVIEWED | Approved |
| `#ff0000` red | DISPUTED | Disputed — needs rewrite/correction |
| `#b4a7d6` purple | `%C/O%` | Empty cell — content is optional (empire-specific) |

## Sheet Rules

- **All-caps tabs** (e.g. `STATS`, `VALUES`) are completely ignored.
- **Columns where row 1 = row 2** are Stellaris export formula columns — ignored.
- **Original data is never modified** — status is tracked in a separate hidden sheet.
- **Column A** (translation keys) is write-protected with a warning.
