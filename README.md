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
If you later want much better handwriting accuracy, the upgrade path is a cloud OCR API
(Google Vision / Azure Document Intelligence) — that needs internet access and has a
per-page cost, which is why it isn't the default here.

### 1c. Create the database (only once — do not re-run after you have real data)
```bash
python build_excel_template.py
```
This creates `data/SchoolDisciplineSystem.xlsx` with 5 example students already in the
Master List so you can see how everything works. **Replace/remove them** on the
"Master student list" page before rolling this out for real.

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
   roll no). This is the list every autocomplete/dropdown pulls from. Keep it updated as
   students join/leave, or to fix a misspelled name.
2. **Manual entry** — pick a category (e.g. Late Comers) → "Log entry" → start typing a
   name → pick the matching student from the dropdown (this is what prevents duplicate or
   misspelled name records) → fill in the category-specific field + remarks → Save.
3. **Scan register** — same category → "Scan register" → upload a photo of the paper
   register page → review each OCR-guessed line, fix names via the same search box,
   uncheck any junk lines → Save. Nothing is written to the database until you click Save
   on the review screen.
4. **Dashboard** (`/`) — live KPI cards (today/this week/this month/all-time) per category,
   a "most frequently flagged students" leaderboard, and a recent-activity feed.

You can also open `data/SchoolDisciplineSystem.xlsx` directly in Excel any time — it's a
normal spreadsheet with its own Dashboard tab (charts + KPI formulas) and one tab per
category, and the Name columns there also have dropdown validation tied to the Master List.
**Only open it in Excel when the server (`app.py`) isn't actively writing to it at that
exact moment** — Excel locks the file for editing, which would make the web app's next save
fail until you close it.

---

## 4. Project files
```
app.py                          the Flask web server
build_excel_template.py         one-time script that creates the Excel database
requirements.txt                Python packages needed
run_server.bat                  double-click to start the server
data/SchoolDisciplineSystem.xlsx   the database (create with build_excel_template.py)
data/scans/                     uploaded register photos are kept here for reference
templates/, static/             the web app's pages and styling
```
