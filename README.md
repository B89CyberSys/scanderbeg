# scanderbeg
Intelligent file organizer for Windows. Sorts by context, not just extension. Built for German digital life.
[README.md](https://github.com/user-attachments/files/27459685/README.md)
# Scanderbeg
### Intelligent File Organizer for Windows · v8 Smart Scan

> *Built for the way people actually live with files — not the way developers imagine they should.*

---

## What it does

Most file organizers move files by extension. `.pdf` goes to Documents. `.jpg` goes to Images. Done.

Scanderbeg goes further. It reads context — filenames, folder patterns, and (optionally) file content — to figure out where a file *actually* belongs. A PDF named `Rechnung_April_2026.pdf` isn't just a PDF. It's an invoice. Scanderbeg knows the difference.

---

## Features

- **Smart Scan** — keyword-based routing for TXT, DOCX, and PDF files. 100% local. No internet required.
- **Delay-based Auto-sweep** — set it and forget it. Scanderbeg quietly sorts on your schedule.
- **File Queue** — review what's waiting before committing. Full control, always.
- **Watchdog monitoring** — watches your Desktop and Downloads in real time.
- **User-defined destinations** — every category goes exactly where you tell it.
- **Locked file handling** — gracefully skips files in use, retries later.
- **Undo last move** — made a mistake? One click back.
- **Modern UI** — dark/light adaptive, Windows Acrylic/Mica native blur, 3D button feel.
- **Multilingual** — English, Deutsch, Shqip.
- **Privacy first** — nothing leaves your machine. Ever.

---

## Built for German digital life

Scanderbeg ships with German-aware Smart Scan keywords out of the box:

| Category | Recognizes |
|---|---|
| Invoices | Rechnung, MwSt, IBAN, Beleg, Quittung... |
| Contracts | Vertrag, Kündigung, Laufzeit, AGB... |
| Umschulung | Lernfeld, IHK, Fachinformatiker, Subnetting... |
| Health | Befund, Rezept, Diagnose, Arzt... |
| Work | Auftrag, Angebot, Protokoll, Projekt... |

No other file organizer on the market thinks in German administrative life. This one does.

---

## Requirements

```
Python 3.10+
customtkinter
watchdog
```

Optional (for Smart Scan content reading):
```
PyPDF2
python-docx
```

Install dependencies:
```bash
pip install customtkinter watchdog
pip install PyPDF2 python-docx  # optional
```

---

## Run

```bash
python scanderbeg_v8_smart_scan.py
```

---

## Design philosophy

Scanderbeg is less about "cleaning files" and more about reducing digital chaos without annoying the user.

The system is designed to feel:
- **lightweight** — no background services you didn't ask for
- **non-invasive** — it suggests, you decide
- **predictable** — no surprises, no mystery moves
- **psychologically comfortable** — calm UI, no aggressive alerts

The sorting logic follows human thinking, not machine defaults.

---

## Roadmap

- [ ] Context-aware document classification (AI-assisted, optional)
- [ ] In-file text scanning expanded to more formats
- [ ] Installer packaging (Inno Setup)
- [ ] Windows tray integration
- [ ] Profile presets (Student / Office Worker / Developer)

---

## Author

**Behar Terakaj** — [@beharterakaj](https://medium.com/@beharterakaj) on Medium  
Built in Germany. Designed for real life.

---

## License

MIT — free to use, free to modify, keep the credit.
