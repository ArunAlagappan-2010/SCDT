# School Discipline & Grooming Tracker

A local system for logging **Late Comers**, **Defaulters**, **Uniform**, and **Nails & Hair**
cases. One PC on the school network acts as the "server" — every other PC just opens a
web page, no installation needed on the other ~60 machines.

- **Database**: a single Excel file, `data/SchoolDisciplineSystem.xlsx`, structured like a
  dashboard (KPI cards, charts, a master student list with dropdown validation).
- **Web app**: `app.py`, a small Flask server. Staff use it to log entries with
  name-autocomplete against the master list, or scan a photo of the handwritten register
  and OCR pulls out candidate names for review before anything is saved.
- **Access**: local network only. No internet/cloud tunnel needed — the other PCs just need
  to be on the same Wi-Fi/LAN as the server PC.

---

## 0. Quickest way: install it as an app (server PC)

1. Install Python 3.9+ from https://www.python.org/downloads/ (tick **Add python.exe to PATH**).
2. Double-click **`Install MCTM.bat`**. It installs the packages, creates the database if
   there isn't one (existing data is never touched), and adds:
   - a **MCTM Tracker** icon on the Desktop,
   - a Start menu folder with *MCTM Tracker*, *Stop MCTM server* and *Uninstall MCTM Tracker*,
   - an entry in **Settings > Apps > Installed apps**.
   No admin rights needed.
3. From then on, just click **MCTM Tracker**. It starts the server quietly in the background
   (no black window) and opens the tracker in its own app window. Clicking it again while
   it's running just opens another window.

Closing the window does **not** stop the server, so the other PCs keep working. To stop
it, use **Start menu > MCTM Tracker > Stop MCTM server**. Server output goes to
`data/server.log`.

**To uninstall:** Start menu > MCTM Tracker > *Uninstall MCTM Tracker*, Settings > Apps, or
double-click **`Uninstall MCTM.bat`**. It stops the server, offers to copy your data (the
Excel database, `.env` and scanned photos) to `Documents\MCTM Tracker backup <date>`,
removes the shortcuts and the Apps entry, and asks whether to remove the program folder
too. Everything it removes goes to the **Recycle Bin**; Python itself is left installed.

The manual steps below still work if you'd rather run it by hand.

## 1. One-time setup (on the server PC only)

### 1a. Install Python packages
```bash
pip install -r requirements.txt
```

### 1b. Install the Tesseract OCR engine (required for the "Scan register" feature)
`pytesseract` is just a Python wrapper — it needs the actual OCR engine installed separately:

1. Download the Windows installer from the UB Mannheim build:
   https://github.com/UB-Mannheim/tesseract/wiki
2. Run the installer (default install path is `C:\Program Files\Tesseract-OCR\`).
3. That's it — `app.py` already points at that default path. If you installed it somewhere
   else, edit the `TESSERACT_CMD` line near the top of `app.py`.

**Important, honest caveat:** Tesseract is built for *printed* text. Handwriting recognition
accuracy will be inconsistent — sometimes good, sometimes poor, depending on handwriting
and photo quality. That's exactly why the "Scan register" flow always shows a **review
screen** before saving: every OCR-guessed name has the same autocomplete search box as
manual entry, so staff correct mistakes in a few seconds rather than trusting OCR blindly.

### 1b-2. Optional: use Gemini instead of Tesseract (much better on handwriting)
Google's Gemini API reads handwriting far more reliably than Tesseract, and has a free
tier. It's **not the default** because it needs internet access on the server PC and sends
the scanned photo to Google's servers — worth deciding deliberately since register photos
contain real students' names, not something to turn on silently.

If you want it:
1. Get a free key at https://aistudio.google.com/apikey (no credit card needed).
2. Copy `.env.example` to `.env` (same folder) and paste your key into `GEMINI_API_KEY=`.
3. Restart `app.py`. It automatically uses Gemini whenever a key is present, and falls back
   to Tesseract when `.env` is blank or missing — the "Scan register" page shows which
   engine is active.

`.env` is gitignored — your key never gets committed or pushed to GitHub.

### 1c. Create the database (only once — do not re-run after you have real data)
```bash
python build_excel_template.py
```
This creates `data/SchoolDisciplineSystem.xlsx` with 5 example students already in the
Master List so you can see how everything works. **Replace/remove them** on the
"Master student list" page before rolling this out for real.

**Load the school's student list.** Put the roster workbook (one sheet per grade, e.g.
"Grade 9", "Grade IB-1A", each with *Student Name* and *House* columns) anywhere, close Excel,
stop the server, and run:
```bash
python import_roster.py "C:\path	o\SCHOOL RECORD.xlsx" --dry-run
python import_roster.py "C:\path	o\SCHOOL RECORD.xlsx"
```
The first line only shows what would happen; the second does it, after saving a backup.
Students get numbers STU001, STU002... grade by grade; "Grade IB-1A" becomes class IB-1,
section A; house spellings are cleaned (Vindyas -> Vindhyas, Aravalis -> Aravallis); two
students with the same name get their class added ("Samyuktha (7)", "Samyuktha (9)"); the
5 example students are removed. Running it again later (e.g. with new admissions) keeps
everyone's existing number and never removes a student who isn't in the new list.

**Already have a database from an older version?** Don't rebuild it. Close Excel, stop
`app.py`, and run:
```bash
python restyle_workbook.py
```
It saves a timestamped backup next to the file, then moves every student and entry onto the
current layout (new Dashboard, live per-student counts on the MasterList, a "How to use" tab).

---

## 2. Running the server

Double-click `run_server.bat`, or run:
```bash
python app.py
```
You'll see something like:
```
Running on http://192.168.1.168:5000
```
That `192.168.1.168` (yours will differ) is the address every other PC will use.

### Find the server's LAN IP address
```bash
ipconfig
```
Look for "IPv4 Address" under the active network adapter (Wi-Fi or Ethernet).

### Allow other PCs through the Windows Firewall
The first time you run the app, Windows will likely show a firewall prompt — click
**Allow access** (for both Private and Public networks, if your school network is
flagged Public). If you miss that prompt, add a rule manually:
```powershell
New-NetFirewallRule -DisplayName "Discipline Tracker" -Direction Inbound -Protocol TCP -LocalPort 5000 -Action Allow
```
(Run PowerShell as Administrator for that command.)

### On the other ~60 PCs
Just open a browser and go to:
```
http://<server-ip>:5000
```
e.g. `http://192.168.1.168:5000`. Bookmark it on each machine — no install required there.

### Keep it running
The server PC needs `app.py` running whenever staff want to use the system. For a more
"always on" setup later, you can use Windows Task Scheduler to run `run_server.bat` at
startup, or ask me to set that up.

---

## 3. Day-to-day workflow

1. **Master student list** (`/students`) — add every student once (name, class, section,
   roll no, **house**). This is the list every autocomplete/dropdown pulls from, and the
   source of truth for a student's House/Class once their name is matched anywhere else in
   the system. Keep it updated as students join/leave, or to fix a misspelled name.
2. **Manual entry** — pick a category (e.g. Late Comers) → "Log entry" → start typing a
   name → pick the matching student from the dropdown (this is what prevents duplicate or
   misspelled name records) → fill in the category-specific field + remarks → Save.
3. **Scan register** — same category → "Scan register" → upload a photo of the paper
   register page. The review screen shows one row per detected student with three boxes:
   **House**, **Name** (with the same autocomplete as manual entry), and **Class**. The
   register's own header row (the line that just says "House / Name / Class") is
   automatically filtered out and never shown as a fake entry. Picking a name from the
   autocomplete overwrites House/Class with that student's real master-list record; fix
   anything else, uncheck junk lines, then Save. Nothing is written to the database until
   you click Save on the review screen.
4. **Dashboard** (`/`) — live KPI cards (today/this week/this month/all-time) per category,
5. **Find students** (`/find`) — type a name or student ID, then narrow by register, period,
   class, section, house, or "N or more times" to find repeat late comers. Results are one row
   per student, most entries first. Click a name (here or anywhere in the app) for the
   **student profile**: totals per register, full history, usual arrival time and other
   patterns. Both pages have **Download Excel** and **Print or save PDF** buttons.
   a "most frequently flagged students" leaderboard, and a recent-activity feed.

You can also open `data/SchoolDisciplineSystem.xlsx` directly in Excel any time. Its tabs:
- **Dashboard**: cases by period, month by month (with a stacked chart), the 10 most-flagged
  students (sorted live), and a by-house table whose house names you type in the yellow cells.
- **MasterList**: one row per student; columns G-K count each student's entries live, and a
  total of 3 or more turns mint.
- **LateComers / Defaulters / Uniform / NailsHair**: one row per case; pick the Name from its
  dropdown and Class, Section and House fill in. Scanned (OCR) entries are tinted mint.
- **How to use**: what the colours mean and what not to change.
**Only open it in Excel when the server (`app.py`) isn't actively writing to it at that
exact moment** — Excel locks the file for editing, which would make the web app's next save
fail until you close it.

---

## 4. Project files
```
app.py                          the Flask web server
build_excel_template.py         one-time script that creates the Excel database
restyle_workbook.py             moves an existing database onto the current layout (keeps all data)
import_roster.py                loads the school's grade-by-grade student list into the MasterList
migrate_add_house.py            one-time script that added the House column to an existing database
requirements.txt                Python packages needed
run_server.bat                  double-click to start the server
MCTM.pyw                        one-click launcher (starts the server hidden, opens the app window)
Install MCTM.bat                installs shortcuts + Settings > Apps entry (runs installer/install.ps1)
Uninstall MCTM.bat              removes it again, backing up data first (installer/uninstall.ps1)
installer/                      install/uninstall scripts, stop_server.pyw, make_icon.py
.env.example                    template for optional Gemini API key (copy to .env)
data/SchoolDisciplineSystem.xlsx   the database (create with build_excel_template.py)
data/scans/                     uploaded register photos are kept here for reference
templates/, static/             the web app's pages and styling
tests/test_app.py               automated smoke tests (run: python tests/test_app.py)
```
