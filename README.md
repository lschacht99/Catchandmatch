# Cohen Expense Flow

A private, mobile-first receipt-capture and event-expense tool for **Leah & Moshe Cohen**.

Take a photo of a receipt → it lands in the right Google Drive folder with a clean name →
AI reads it → you review and approve → it's written to a Google Sheets ledger and tracked by event.

There is **no server to run, no hosting bill, and no API keys in the browser.** Everything lives in
one Google Apps Script project that runs as your own Google account.

---

## How it's built (and why)

| Piece | What it does |
|---|---|
| **Apps Script** | Serves the app *and* runs the backend. Drive + Sheets calls run as **you** — no service account, no OAuth juggling. |
| **`index.html`** | The mobile UI. Talks to the backend with `google.script.run`, so it's same-origin — **zero CORS problems**. |
| **OpenAI (server-side)** | The only outside call. The key lives in Script Properties and never reaches the phone. |
| **Image downscaling** | Photos are shrunk in the browser before upload — small payloads, faster OCR, smaller Drive files. |

> This replaces the original Next.js/Vercel plan on purpose: for a two-person private tool, Apps Script
> does the same job with far fewer moving parts and nothing to deploy or secure separately.

---

## What's in v1

- Passcode login (only people with the shared passcode get in)
- Add receipt: organization → event group → event (with suggestions + "create new event")
- Photo capture from phone camera
- Upload to Google Drive with auto-created folders and a clean filename
- AI extraction (merchant, date, totals, line items, categories, kosher flags)
- Review screen — edit everything before saving; nothing is final without your approval
- Writes to the **Receipts** and **Items** tabs, updates **PeopleBalances**, logs to **AuditLog**
- Events list + per-event dashboard (totals, category breakdown, who-owes-whom)

**Postponed to v2 (architecture is ready):** Notion sync, price comparison. Tabs already exist for both.

---

## Setup (about 15 minutes, one time)

### 1. Create the Apps Script project
1. Go to <https://script.google.com> → **New project**.
2. Delete the default `Code.gs` contents.
3. Create the files to match this repo:
   - `Code.gs` → paste in `Code.gs`
   - **+ → HTML** named `index` → paste in `index.html`
   - **Project Settings (gear) → "Show appsscript.json"** → paste in `appsscript.json`

> Prefer the command line? Use [clasp](https://github.com/google/clasp): `clasp create`, then `clasp push`.

### 2. Build the Drive folder + Google Sheet automatically
1. In the editor, choose the function **`setup`** in the toolbar dropdown → **Run**.
2. Approve the permission prompt (Drive, Sheets, external requests). It's your own account.
3. Open **Executions / Logs** — `setup()` returns the new Sheet URL and Drive folder URL.

This creates:
- Drive folder **`Cohen Expense Flow`** with `Receipts/`, `Exports/`, `Logs/` inside.
- Spreadsheet **`Cohen Expense Flow Ledger`** with all 8 tabs (plus 2 hidden helper tabs):
  `Organizations`, `EventGroups`, `Events`, `Receipts`, `Items`, `PeopleBalances`, `PriceChecks`, `AuditLog`.
- Seeded organizations: Hamsa Nomads, Moshe House, Personal, Work, Other.

You don't have to create any tabs or folders by hand — `setup()` does it and remembers the IDs.

### 3. Add your secrets (Script Properties)
**Project Settings → Script Properties → Add script property:**

| Key | Value | Required |
|---|---|---|
| `PASSCODE` | A shared passcode for Leah & Moshe (e.g. a 6-digit number) | ✅ |
| `OPENAI_API_KEY` | Your OpenAI API key (`sk-…`) | ✅ |
| `OPENAI_MODEL` | `gpt-4o-mini` (default) or `gpt-4o` for tougher receipts | optional |

`SHEET_ID` and `ROOT_FOLDER_ID` are set for you by `setup()`.

### 4. Deploy as a web app
1. **Deploy → New deployment → Web app.**
2. **Execute as:** *Me* · **Who has access:** *Anyone*.
   (“Anyone” + the passcode = simple and private. It runs as you; the passcode is the gate.)
3. Copy the **Web app URL** and open it on your phone.

---

## Add to your phone's Home Screen

- **iPhone (Safari):** open the web app URL → Share → **Add to Home Screen**.
- **Android (Chrome):** open the URL → ⋮ menu → **Add to Home screen**.

It opens full-screen like an app. (A full offline PWA can come later if you move the UI to
Cloudflare Pages — see "Going further" below.)

---

## Daily use

**Upload & approve a receipt**
1. Tap **Add receipt**.
2. Pick **Organization → Event group → Event** (or **+ Create new event**).
3. Choose **Paid by** (Leah / Moshe / Other).
4. **Tap to take a photo** of the receipt.
5. **Upload & scan** → wait a few seconds for AI.
6. On **Review**: fix merchant/date/total and line items if needed, then **Approve & save**.
   - Nothing is written as final until you approve.
   - If items don't add up to the total, you'll see a gentle warning (tax/tip/fees are normal).

**Add a new organization or event**
- New events are created right inside the Add-receipt form ("Create new event").
- New organizations: type one in the create form, or add a row to the `Organizations` tab.

---

## File + folder conventions

**Folders:** `Cohen Expense Flow / Receipts / <Org> / <Event Group> / <Event> /`

**Filenames** (after AI reads merchant + total):
```
YYYY-MM-DD__Org__EventGroup__Event__PaidBy__Merchant__Total__ReceiptID.jpg
2026-06-26__Moshe-House__Shabbat-Dinners__Moshe-House-Shabbat-Dinner-June-2026__Moshe__Costco__187-42__R-000124.jpg
```
Before AI runs, the file is saved as `PENDING__YYYY-MM-DD__PaidBy__ReceiptID.jpg` and renamed after.
Every receipt also gets a stable ID (`R-000001`) — the Sheet, not the filename, is the source of truth.

---

## Security notes

- Only people with the `PASSCODE` can use the app; every server call re-checks it.
- API keys live in Script Properties, server-side — never sent to the browser.
- The Drive folder and Sheet are private to your Google account.
- Receipt image links are Drive links (only people you share with can open them).
- The `AuditLog` tab records uploads, approvals, and deletions.

---

## Going further (v2)

- **Notion sync:** one dashboard page per event (display layer only — the Sheet stays the ledger).
- **Price comparison:** queue items over $10, compare *unit* price via SerpApi/Google Shopping,
  and **require human review for kosher-sensitive food** unless kosher equivalence is clear.
  The `PriceChecks` tab is already there.
- **True installable PWA:** move `index.html` to Cloudflare Pages / GitHub Pages and call the
  Apps Script backend. Note: cross-origin calls to Apps Script need care (CORS) — the simplest
  reliable path is keeping the UI served by Apps Script as it is now.

---

## Troubleshooting

- **"App not configured" / "Access not allowed":** set `PASSCODE` in Script Properties; make sure
  you're entering the same value.
- **AI returns empty / low confidence:** retake the photo in better light, or tap **Re-run AI**.
  You can always fill fields in by hand.
- **Permission errors on first run:** re-run `setup()` and accept the Google permission prompt.
- **Changed the code?** Re-deploy: **Deploy → Manage deployments → edit → New version.**
