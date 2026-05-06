"""Scanderbeg v8 - Smart Scan Edition
Fixed File Queue action bar: dropdown and Sweep Selected stay visible.
Requirements: pip install customtkinter watchdog
Optional: pip install PyPDF2 python-docx
"""
import os, sys, json, time, shutil, threading, tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from datetime import datetime
import customtkinter as ctk

try:
    import PyPDF2
    PYPDF2_AVAILABLE=True
except ImportError:
    PYPDF2_AVAILABLE=False
try:
    from docx import Document
    DOCX_AVAILABLE=True
except ImportError:
    DOCX_AVAILABLE=False
try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    WATCHDOG_AVAILABLE=True
except ImportError:
    WATCHDOG_AVAILABLE=False
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_MIXIN=TkinterDnD.DnDWrapper
    TKDND_AVAILABLE=True
except ImportError:
    DND_FILES=None
    DND_MIXIN=object
    TKDND_AVAILABLE=False
try:
    import winsound
    def play_sound(n): winsound.PlaySound(n,winsound.SND_ALIAS|winsound.SND_ASYNC)
except ImportError:
    def play_sound(n): pass

ctk.set_appearance_mode("system"); ctk.set_default_color_theme("blue")

_BaseCTkButton = ctk.CTkButton
class ModernButton(_BaseCTkButton):
    """Raised 3D button styling used by every CTkButton in the app."""
    def __init__(self,*args,**kwargs):
        kwargs.setdefault("corner_radius",13)
        kwargs.setdefault("border_width",2)
        kwargs.setdefault("border_color",("#DBEAFE","#5B7FB4"))
        kwargs.setdefault("font",ctk.CTkFont(size=13,weight="bold"))
        kwargs.setdefault("height",36)
        super().__init__(*args,**kwargs)
        self._raised_border=kwargs.get("border_color",("#DBEAFE","#5B7FB4"))
        self._raised_fg=kwargs.get("fg_color",self.cget("fg_color"))
        self._raised_hover=kwargs.get("hover_color",self.cget("hover_color"))
        self._pressed=False
        self._canvas_offset=0
        self.bind("<Enter>",self._hover_depth,add="+")
        self.bind("<ButtonPress-1>",self._press_depth,add="+")
        self.bind("<ButtonRelease-1>",self._release_depth,add="+")
        self.bind("<Leave>",self._release_depth,add="+")
    def _shift_canvas(self,dy):
        try:
            canvas=getattr(self,"_canvas",None)
            if canvas is not None:
                canvas.move("all",0,dy)
                self._canvas_offset+=dy
        except Exception:
            pass
    def _hover_depth(self,event=None):
        if self._pressed: return
        try:
            self.configure(border_color=("#FFFFFF","#7EA6D9"),border_width=2)
        except Exception:
            pass
    def _press_depth(self,event=None):
        self._pressed=True
        try:
            self.configure(border_color=("#5B7FB4","#0B2344"),border_width=2)
            if self._canvas_offset==0: self._shift_canvas(1)
        except Exception:
            pass
    def _release_depth(self,event=None):
        self._pressed=False
        try:
            if self._canvas_offset: self._shift_canvas(-self._canvas_offset)
            self.configure(border_color=self._raised_border,border_width=2)
        except Exception:
            pass

ctk.CTkButton = ModernButton

_BTN_PRI=dict(fg_color=("#2F80ED","#1d4ed8"),hover_color=("#1D4ED8","#2563eb"),text_color="#fff",border_width=2,border_color=("#BFDBFE","#60A5FA"),corner_radius=13)
_BTN_SEC=dict(fg_color=("#FFFFFF","#1A2B44"),hover_color=("#EAF1FB","#243955"),text_color=("#172033","#E2E8F0"),border_width=2,border_color=("#E2E8F0","#58708F"),corner_radius=13)
_BTN_ALT=dict(fg_color=("#2F405B","#1d4ed8"),hover_color=("#334155","#2563eb"),text_color="#fff",border_width=2,border_color=("#94A3B8","#60A5FA"),corner_radius=13)
_BTN_DNG=dict(fg_color=("#FFF7F7","#251013"),hover_color=("#FEE2E2","#351515"),text_color=("#DC2626","#FCA5A5"),border_width=2,border_color=("#FECACA","#7F1D1D"),corner_radius=13)
_SHADOW_LIGHT,_SHADOW_DARK="#BFD1E8","#020814"; _PAGE_LIGHT,_PAGE_DARK="#EAF2FA","#07111f"; _CARD_LIGHT,_CARD_DARK="#F8FBFF","#102033"; _ROW_LIGHT,_ROW_DARK="#F8FAFC","#12233a"
DEFAULT_CATEGORIES={
"Images":[".jpg",".jpeg",".png",".gif",".bmp",".webp",".svg",".ico",".tiff",".raw"],"Videos":[".mp4",".mkv",".avi",".mov",".wmv",".flv",".webm",".m4v"],"Audio":[".mp3",".wav",".flac",".aac",".ogg",".wma",".m4a"],
"Documents":[".pdf",".doc",".docx",".odt",".txt",".rtf",".md",".pages"],"Spreadsheets":[".xls",".xlsx",".csv",".ods",".numbers"],"Slides":[".ppt",".pptx",".odp",".key"],"Archives":[".zip",".rar",".7z",".tar",".gz",".bz2",".xz"],
"Installers":[".exe",".msi",".msix",".appx"],"Code":[".py",".js",".ts",".html",".css",".java",".cpp",".c",".cs",".php",".rb",".go",".rs",".sh",".bat",".ps1",".json",".xml",".yaml",".yml",".toml",".ini",".cfg"],"Fonts":[".ttf",".otf",".woff",".woff2"],"Ebooks":[".epub",".mobi",".azw",".azw3",".fb2"],"Other":[]}
ELEMENTAL_CATEGORIES={
"Images":[".jpg",".jpeg",".png",".gif",".bmp",".webp",".svg",".ico"],"Videos":[".mp4",".mkv",".avi",".mov",".wmv",".webm",".m4v"],"Audio":[".mp3",".wav",".flac",".aac",".ogg",".m4a"],
"PDF Files":[".pdf"],"Word Documents":[".doc",".docx",".odt",".pages"],"Excel Spreadsheets":[".xls",".xlsx",".csv",".ods",".numbers"],"PowerPoint Presentations":[".ppt",".pptx",".odp",".key"],"Text Documents":[".txt",".rtf",".md"],
"Archives":[".zip",".rar",".7z",".tar",".gz"],"Installers":[".exe",".msi",".msix",".appx"],"Other":[]}
OPTIONAL_CATEGORY_PRESETS={
"Code Files":DEFAULT_CATEGORIES["Code"],"Fonts":DEFAULT_CATEGORIES["Fonts"],"Ebooks":DEFAULT_CATEGORIES["Ebooks"],"Raw Photos":[".raw",".tiff",".cr2",".nef",".arw",".dng",".orf",".rw2",".raf"],
"Design Files":[".psd",".ai",".indd",".fig",".xd",".sketch",".afdesign",".afphoto"],"Config Files":[".json",".xml",".yaml",".yml",".toml",".ini",".cfg",".env"],"Database Files":[".db",".sqlite",".sqlite3",".mdb",".accdb",".sql"],
"Disk Images":[".iso",".img",".dmg",".vhd",".vhdx"],"Virtual Machines":[".ova",".ovf",".vbox",".vmx",".vmdk"],"CAD Files":[".dwg",".dxf",".stp",".step",".iges",".igs"],"3D Models":[".obj",".fbx",".stl",".glb",".gltf",".blend"],
"Subtitles":[".srt",".vtt",".ass",".ssa",".sub"],"Torrents":[".torrent"],"Backups":[".bak",".backup",".old",".orig"],"Certificates & Keys":[".crt",".cer",".pem",".pfx",".p12",".key"],"Log Files":[".log"],
"Scripts":[".ps1",".bat",".cmd",".sh",".zsh",".bash"],"Mobile Apps":[".apk",".ipa"],"Game Files":[".sav",".pak",".wad",".unitypackage"],"Custom Category":[]}
LEGACY_SPLIT_CATEGORIES={"Documents","Spreadsheets","Slides"}
DEFAULT_SMART_KEYWORDS={"Invoices":["invoice","receipt","payment","paid","total","rechnung","rechnungsnummer","quittung","beleg","betrag","summe","mwst","umsatzsteuer","iban","kundennummer"],"Contracts":["contract","agreement","signature","terms","vertrag","vereinbarung","unterschrift","kündigung","kuendigung","laufzeit","frist","agb"],"Umschulung":["umschulung","aufgabe","lösung","loesung","arbeitsblatt","lernfeld","prüfung","pruefung","ihk","fachinformatiker","netzwerk","subnetting","server","client"],"Work":["work","project","meeting","client","deadline","arbeit","projekt","besprechung","kunde","angebot","auftrag","protokoll"],"Health":["health","doctor","hospital","medicine","gesundheit","arzt","krankenhaus","befund","rezept","patient","diagnose"]}
SKIP_FILES={"desktop.ini","thumbs.db",".ds_store","ntuser.dat","ntuser.ini","iconcache.db"}; SKIP_EXTENSIONS={".lnk",".url",".tmp",".sys"}; CONFIG_FILE=Path.home()/".scanderbeg_config.json"; CONFIG_VERSION=2; MANUAL_FILE="Scanderbeg_User_Manual.html"; SUPPORT_EMAIL="beharterakaj@outlook.com"
LANGUAGES={"en":"English","de":"Deutsch","sq":"Shqip"}
TRANSLATIONS={
"en":{"dashboard":"Dashboard","file_queue":"File Queue","auto_sweep":"Auto-sweep","rules":"File Types","settings":"Settings","scan_now":"Scan Now","smart_scan":"Smart Scan","files_waiting":"Files Waiting","files_moved":"Files Moved","auto_sweeps":"Auto-sweeps","watching":"Watching","quick_actions":"Quick Actions","recent_activity":"Recent Activity","watched_folders":"Watched Folders","all_clean":"All clean. Well done! 🎉","files_waiting_status":"{n} files waiting","organized":"Your files are organized.","queue_waiting":"Your File Queue has items waiting to be sorted.","current_queue":"Current queue","this_install":"This install","total":"Total","folders":"Folders","open_downloads":"📂 Open Downloads","open_desktop":"🖥 Open Desktop","language":"UI language"},
"de":{"dashboard":"Übersicht","file_queue":"Datei-Warteschlange","auto_sweep":"Auto-Sortierung","rules":"Dateitypen","settings":"Einstellungen","scan_now":"Jetzt scannen","smart_scan":"Smart Scan","files_waiting":"Wartende Dateien","files_moved":"Verschobene Dateien","auto_sweeps":"Auto-Läufe","watching":"Überwacht","quick_actions":"Schnellaktionen","recent_activity":"Letzte Aktivität","watched_folders":"Überwachte Ordner","all_clean":"Alles sauber. Sehr gut! 🎉","files_waiting_status":"{n} Dateien warten","organized":"Deine Dateien sind organisiert.","queue_waiting":"In der Datei-Warteschlange warten Elemente.","current_queue":"Aktuelle Liste","this_install":"Diese Installation","total":"Gesamt","folders":"Ordner","open_downloads":"📂 Downloads öffnen","open_desktop":"🖥 Desktop öffnen","language":"UI-Sprache"},
"sq":{"dashboard":"Paneli","file_queue":"Radha e skedarëve","auto_sweep":"Auto-rregullim","rules":"Llojet e skedarëve","settings":"Cilësimet","scan_now":"Skano tani","smart_scan":"Skanim inteligjent","files_waiting":"Skedarë në pritje","files_moved":"Skedarë të lëvizur","auto_sweeps":"Auto-kontrolle","watching":"Në vëzhgim","quick_actions":"Veprime të shpejta","recent_activity":"Aktiviteti i fundit","watched_folders":"Dosje të vëzhguara","all_clean":"Gjithçka në rregull. Shumë mirë! 🎉","files_waiting_status":"{n} skedarë presin","organized":"Skedarët e tu janë të organizuar.","queue_waiting":"Ka skedarë që presin në radhë.","current_queue":"Radha aktuale","this_install":"Ky instalim","total":"Gjithsej","folders":"Dosje","open_downloads":"📂 Hap Downloads","open_desktop":"🖥 Hap Desktop","language":"Gjuha e UI"}}
def tx(cfg,key,**kwargs):
    lang=cfg.get("ui_language","en")
    text=TRANSLATIONS.get(lang,TRANSLATIONS["en"]).get(key,TRANSLATIONS["en"].get(key,key))
    return text.format(**kwargs) if kwargs else text
TRANSLATIONS["en"].update({"run_scan":"Run Scan","browse":"Browse","scan_folder":"Scan folder:","more_actions":"More actions","select_all":"Select all","deselect_all":"Deselect all","ignore_selected":"Ignore selected","clear_ignored_files":"Clear ignored files","open_file_location":"Open file location","move_custom":"Move to custom folder","set_dest_category":"Set destination for category","undo_last_move":"Undo last move","quick_add":"Quick Add","manual_support":"Manual & Support","open_manual":"Open User Manual","email_support":"Email Support","rescan_folder":"Rescan folder","clear_activity_log":"Clear activity log","sweep_selected":"Sweep Selected","file_name":"File Name","category":"Category","goes_to":"Goes To","size":"Size","modified":"Modified","status":"Status","queue_empty":"Queue is empty.","folder_clean":"Folder is clean. Nothing to sort.","files_to_sort":"{n} files to sort","already_sorted":"{n} already sorted","set_dest_for":"Set destinations for: {cats}","rules_intro":"Edit category names and the extensions that belong to each category.","add_category":"+ Add Category","extensions":"Extensions","save_rules":"Save Rules","reset_defaults":"Reset Defaults","remove":"Remove","auto_status_on":"● Auto-sweep is ON","auto_status_off":"● Auto-sweep is OFF","start_auto":"Start Auto-sweep","stop_auto":"Stop Auto-sweep","delay_settings":"Delay: {n} minutes - change in Settings","watched_folders_lower":"Watched folders","activity_log":"Activity log","add_folder":"+ Add Folder","auto_sort_delay":"Auto-sort delay","minutes":"Minutes:","save":"Save","smart_scan_desc":"TXT/DOCX/PDF keyword routing, 100% local.","enable_smart":"Enable Smart Scan","min_score":"Min score:","max_mb":"Max MB:","destination_categories":"Destination Categories","save_destinations":"Save Destinations","config":"Config: {path}"})
TRANSLATIONS["de"].update({"run_scan":"Scannen","browse":"Durchsuchen","scan_folder":"Scan-Ordner:","more_actions":"Weitere Aktionen","select_all":"Alle auswählen","deselect_all":"Auswahl aufheben","ignore_selected":"Ausgewählte ignorieren","clear_ignored_files":"Ignorierte Dateien löschen","open_file_location":"Dateispeicherort öffnen","move_custom":"In eigenen Ordner verschieben","set_dest_category":"Ziel für Kategorie setzen","undo_last_move":"Letzte Verschiebung rückgängig","quick_add":"Schnell hinzufügen","manual_support":"Handbuch & Support","open_manual":"Handbuch öffnen","email_support":"Support mailen","rescan_folder":"Ordner neu scannen","clear_activity_log":"Aktivitätslog löschen","sweep_selected":"Ausgewählte sortieren","file_name":"Dateiname","category":"Kategorie","goes_to":"Ziel","size":"Größe","modified":"Geändert","status":"Status","queue_empty":"Warteschlange ist leer.","folder_clean":"Ordner ist sauber. Nichts zu sortieren.","files_to_sort":"{n} Dateien zu sortieren","already_sorted":"{n} bereits sortiert","set_dest_for":"Ziele setzen für: {cats}","rules_intro":"Bearbeite Kategorienamen und die zugehörigen Erweiterungen.","add_category":"+ Kategorie hinzufügen","extensions":"Erweiterungen","save_rules":"Regeln speichern","reset_defaults":"Standardwerte","remove":"Entfernen","auto_status_on":"● Auto-Sortierung ist AN","auto_status_off":"● Auto-Sortierung ist AUS","start_auto":"Auto-Sortierung starten","stop_auto":"Auto-Sortierung stoppen","delay_settings":"Verzögerung: {n} Minuten - ändern in Einstellungen","watched_folders_lower":"Überwachte Ordner","activity_log":"Aktivitätslog","add_folder":"+ Ordner hinzufügen","auto_sort_delay":"Auto-Sortierverzögerung","minutes":"Minuten:","save":"Speichern","smart_scan_desc":"TXT/DOCX/PDF-Schlüsselwort-Routing, 100% lokal.","enable_smart":"Smart Scan aktivieren","min_score":"Mindestscore:","max_mb":"Max. MB:","destination_categories":"Ziel-Kategorien","save_destinations":"Ziele speichern","config":"Konfiguration: {path}"})
TRANSLATIONS["sq"].update({"run_scan":"Skano","browse":"Shfleto","scan_folder":"Dosja për skanim:","more_actions":"Më shumë veprime","select_all":"Zgjidh të gjitha","deselect_all":"Hiq zgjedhjen","ignore_selected":"Injoro të zgjedhurat","clear_ignored_files":"Pastro skedarët e injoruar","open_file_location":"Hap vendndodhjen","move_custom":"Zhvendos në dosje tjetër","set_dest_category":"Cakto destinacion për kategori","undo_last_move":"Kthe zhvendosjen e fundit","quick_add":"Shto shpejt","manual_support":"Manuali & Support","open_manual":"Hap manualin","email_support":"Dërgo email supportit","rescan_folder":"Riskano dosjen","clear_activity_log":"Pastro aktivitetin","sweep_selected":"Rregullo të zgjedhurat","file_name":"Emri i skedarit","category":"Kategoria","goes_to":"Shkon te","size":"Madhësia","modified":"Ndryshuar","status":"Statusi","queue_empty":"Radha është bosh.","folder_clean":"Dosja është e pastër. Nuk ka gjë për rregullim.","files_to_sort":"{n} skedarë për rregullim","already_sorted":"{n} tashmë të rregulluar","set_dest_for":"Cakto destinacione për: {cats}","rules_intro":"Ndrysho emrat e kategorive dhe prapashtesat përkatëse.","add_category":"+ Shto kategori","extensions":"Prapashtesa","save_rules":"Ruaj rregullat","reset_defaults":"Rikthe standardet","remove":"Hiq","auto_status_on":"● Auto-rregullimi është AKTIV","auto_status_off":"● Auto-rregullimi është JOAKTIV","start_auto":"Nis auto-rregullimin","stop_auto":"Ndalo auto-rregullimin","delay_settings":"Vonesa: {n} minuta - ndrysho te Cilësimet","watched_folders_lower":"Dosje të vëzhguara","activity_log":"Aktiviteti","add_folder":"+ Shto dosje","auto_sort_delay":"Vonesa e auto-rregullimit","minutes":"Minuta:","save":"Ruaj","smart_scan_desc":"Routim me fjalë kyçe për TXT/DOCX/PDF, 100% lokal.","enable_smart":"Aktivizo skanimin inteligjent","min_score":"Pikët min.:","max_mb":"MB maks.:","destination_categories":"Kategoritë e destinacionit","save_destinations":"Ruaj destinacionet","config":"Konfigurimi: {path}"})
TRANSLATIONS["en"].update({"drop_zone":"Drop Zone"})
TRANSLATIONS["de"].update({"drop_zone":"Drop Zone"})
TRANSLATIONS["sq"].update({"drop_zone":"Zona Drop"})
TRANSLATIONS["en"].update({"confidence":"Confidence","reason":"Reason","add_files":"+ Add Files","sort_listed":"Sort Listed","clear":"Clear","no_files_added":"No files added yet.","destination_folders":"Destination Folders","destination_help":"Choose where each visible file type should go. Add more file types from File Types when needed.","add_file_type_group":"Add file type group","file_types_intro":"Start with common file types. Add specialist groups only when you need them.","shown_in_queue":"shown in File Queue","no_folder_selected":"No folder selected","fallback":"fallback","no_destination_set":"No destination set","no_files_selected":"No files selected.","quick_add_empty":"Quick Add is empty.","trust_local":"100% local","trust_offline":"No internet needed","trust_undo":"Undo supported","first_run_title":"Set up Scanderbeg","first_run_body":"Start with two quick steps:\n\n1. Add the folders you want Scanderbeg to watch.\n2. Set destinations for common file types.\n\nYou can change both anytime in Settings.","open_settings":"Open Settings","later":"Later"})
TRANSLATIONS["de"].update({"confidence":"Vertrauen","reason":"Grund","add_files":"+ Dateien hinzufügen","sort_listed":"Liste sortieren","clear":"Leeren","no_files_added":"Noch keine Dateien hinzugefügt.","destination_folders":"Zielordner","destination_help":"Wähle, wohin jeder sichtbare Dateityp gehen soll. Weitere Dateitypen kannst du bei Dateitypen hinzufügen.","add_file_type_group":"Dateityp-Gruppe hinzufügen","file_types_intro":"Starte mit häufigen Dateitypen. Spezialgruppen fügst du nur bei Bedarf hinzu.","shown_in_queue":"sichtbar in der Datei-Warteschlange","no_folder_selected":"Kein Ordner ausgewählt","fallback":"Fallback","no_destination_set":"Kein Zielordner gesetzt","no_files_selected":"Keine Dateien ausgewählt.","quick_add_empty":"Schnell hinzufügen ist leer.","trust_local":"100% lokal","trust_offline":"Kein Internet nötig","trust_undo":"Undo unterstützt","first_run_title":"Scanderbeg einrichten","first_run_body":"Starte mit zwei kurzen Schritten:\n\n1. Füge die Ordner hinzu, die Scanderbeg überwachen soll.\n2. Setze Ziele für häufige Dateitypen.\n\nBeides kannst du jederzeit in den Einstellungen ändern.","open_settings":"Einstellungen öffnen","later":"Später"})
TRANSLATIONS["sq"].update({"confidence":"Besimi","reason":"Arsyeja","add_files":"+ Shto skedarë","sort_listed":"Rregullo listën","clear":"Pastro","no_files_added":"Ende nuk ka skedarë.","destination_folders":"Dosjet e destinacionit","destination_help":"Zgjidh ku duhet të shkojë çdo lloj skedari. Lloje të tjera shtohen te Llojet e skedarëve.","add_file_type_group":"Shto grup skedarësh","file_types_intro":"Fillo me llojet e zakonshme. Shto grupe speciale vetëm kur duhen.","shown_in_queue":"shfaqet në radhë","no_folder_selected":"Asnjë dosje e zgjedhur","fallback":"rezervë","no_destination_set":"Nuk ka destinacion","no_files_selected":"Nuk ka skedarë të zgjedhur.","quick_add_empty":"Shto shpejt është bosh.","trust_local":"100% lokal","trust_offline":"Pa internet","trust_undo":"Undo i mbështetur","first_run_title":"Konfiguro Scanderbeg","first_run_body":"Fillo me dy hapa të shpejtë:\n\n1. Shto dosjet që Scanderbeg duhet të vëzhgojë.\n2. Cakto destinacione për llojet e zakonshme.\n\nTë dyja mund t'i ndryshosh te Cilësimet.","open_settings":"Hap cilësimet","later":"Më vonë"})
def default_config():
    return {"config_version":CONFIG_VERSION,"setup_tip_seen":False,"scan_folder":str(Path.home()/"Desktop"),"watch_folders":[str(Path.home()/"Desktop"),str(Path.home()/"Downloads")],"category_destinations":{c:"" for c in ELEMENTAL_CATEGORIES},"rules":ELEMENTAL_CATEGORIES,"excluded_folders":["build","dist",".git","__pycache__","node_modules"],"delay_minutes":30,"first_seen":{},"log":[],"stats":{"files_moved":0,"bytes_moved":0,"sweeps":0},"last_moves":[],"move_history":[],"ignored_files":[],"routing_rules":[],"smart_scan_enabled":False,"smart_scan_max_mb":15,"smart_scan_min_score":2,"smart_keywords":DEFAULT_SMART_KEYWORDS,"ui_language":"en"}
def load_config():
    cfg=default_config(); loaded={}
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE,"r",encoding="utf-8") as f: loaded=json.load(f); cfg.update(loaded)
            for k,v in default_config().items(): cfg.setdefault(k,v)
        except Exception: pass
    migrate_config(cfg)
    if loaded and loaded.get("config_version")!=CONFIG_VERSION:
        reset_runtime_state(cfg); cfg["config_version"]=CONFIG_VERSION; save_config(cfg)
    return cfg
def reset_runtime_state(cfg):
    cfg["first_seen"]={}; cfg["log"]=[]; cfg["stats"]={"files_moved":0,"bytes_moved":0,"sweeps":0}; cfg["last_moves"]=[]; cfg["move_history"]=[]; cfg["ignored_files"]=[]
def migrate_config(cfg):
    rules=cfg.setdefault("rules",{})
    if not any(cat in rules for cat in LEGACY_SPLIT_CATEGORIES): return
    dests=cfg.setdefault("category_destinations",{})
    old_doc=dests.get("Documents",""); old_sheet=dests.get("Spreadsheets",""); old_slides=dests.get("Slides","")
    optional_names={tuple(v):k for k,v in OPTIONAL_CATEGORY_PRESETS.items() if k!="Custom Category"}
    new_rules={cat:list(exts) for cat,exts in ELEMENTAL_CATEGORIES.items()}
    for cat,exts in rules.items():
        if cat in LEGACY_SPLIT_CATEGORIES: continue
        if cat in DEFAULT_CATEGORIES and not dests.get(cat): continue
        new_cat=optional_names.get(tuple(exts),cat) if cat in DEFAULT_CATEGORIES else cat
        new_rules.setdefault(new_cat,list(exts))
        if dests.get(cat) and not dests.get(new_cat): dests[new_cat]=dests.get(cat)
    cfg["rules"]=new_rules
    for cat in new_rules: dests.setdefault(cat,"")
    for cat in ["PDF Files","Word Documents","Text Documents"]:
        if old_doc and not dests.get(cat): dests[cat]=old_doc
    if old_sheet and not dests.get("Excel Spreadsheets"): dests["Excel Spreadsheets"]=old_sheet
    if old_slides and not dests.get("PowerPoint Presentations"): dests["PowerPoint Presentations"]=old_slides
    cfg["category_destinations"]={cat:dests.get(cat,"") for cat in new_rules}
def save_config(cfg):
    try:
        with open(CONFIG_FILE,"w",encoding="utf-8") as f: json.dump(cfg,f,indent=2,ensure_ascii=False)
    except Exception: pass
def get_category(filepath,rules):
    ext=Path(filepath).suffix.lower()
    return next((cat for cat,exts in rules.items() if ext in exts),"Other")
def is_hidden(path):
    p=Path(path)
    if p.name.startswith("."): return True
    try:
        import ctypes; attrs=ctypes.windll.kernel32.GetFileAttributesW(str(p)); return attrs!=-1 and bool(attrs&2 or attrs&4)
    except Exception: return False
def should_skip(filepath):
    p=Path(filepath); return p.name.lower() in SKIP_FILES or p.suffix.lower() in SKIP_EXTENSIONS or is_hidden(p)
def is_file_locked(filepath):
    try:
        with open(filepath,"a+b"): return False
    except (IOError,PermissionError,OSError): return True
def is_already_in_destination(filepath,dest):
    if not dest: return False
    try: return Path(filepath).resolve().parent==Path(dest).resolve()
    except Exception: return False

def read_text_preview(path,cfg):
    p=Path(path); ext=p.suffix.lower()
    if ext not in {".txt",".md",".docx",".pdf"}: return ""
    try:
        if p.stat().st_size>int(cfg.get("smart_scan_max_mb",15))*1024*1024 or is_file_locked(p): return ""
    except Exception: return ""
    try:
        if ext in {".txt",".md"}:
            with open(p,"r",encoding="utf-8",errors="ignore") as f: return f.read(8000).lower()
        if ext==".docx" and DOCX_AVAILABLE:
            d=Document(p); return " ".join(x.text for x in d.paragraphs[:200] if x.text).lower()
        if ext==".pdf" and PYPDF2_AVAILABLE:
            r=PyPDF2.PdfReader(str(p)); return " ".join((r.pages[i].extract_text() or "") for i in range(min(2,len(r.pages)))).lower()
    except Exception: pass
    return ""
def smart_scan_match(filepath,cfg):
    d=smart_scan_details(filepath,cfg)
    return d["destination"],d["category"],d["score"],d["hits"]
def smart_scan_details(filepath,cfg):
    empty={"destination":"","category":"","score":0,"hits":[],"confidence":"","reason":""}
    if not cfg.get("smart_scan_enabled",False): return empty
    p=Path(filepath); content=read_text_preview(p,cfg); name=p.stem.lower()
    if not content and not name: return empty
    best=("",0,[],[])
    for cat,words in cfg.get("smart_keywords",DEFAULT_SMART_KEYWORDS).items():
        filename_hits={str(w).strip().lower() for w in words if str(w).strip().lower() and str(w).strip().lower() in name}
        content_hits={str(w).strip().lower() for w in words if content and str(w).strip().lower() and str(w).strip().lower() in content}
        score=len(filename_hits)*3+len(content_hits)*2
        hits=sorted(filename_hits|content_hits)[:8]
        if score>best[1]: best=(cat,score,hits,sorted(filename_hits)[:4])
    if not best[0]: return empty
    min_score=int(cfg.get("smart_scan_min_score",2))
    confidence="High" if best[1]>=max(8,min_score*4) else "Medium" if best[1]>=min_score else "Needs review"
    reason=f"{confidence}: {', '.join(best[2])}" if best[2] else confidence
    dest=cfg.get("category_destinations",{}).get(best[0],"") if best[1]>=min_score else ""
    return {"destination":dest,"category":best[0],"score":best[1],"hits":best[2],"confidence":confidence,"reason":reason}
def resolve_destination(filepath,cfg):
    p=Path(filepath)
    for r in cfg.get("routing_rules",[]):
        val=str(r.get("value","")).strip(); dest=str(r.get("destination","")).strip()
        if val and dest and str(r.get("type","prefix")).lower()=="prefix" and p.name.upper().startswith(val.upper()): return dest
    sd,sc,score,hits=smart_scan_match(p,cfg)
    if sd: return sd
    return cfg.get("category_destinations",{}).get(get_category(p,cfg.get("rules",DEFAULT_CATEGORIES)),"")
def new_move_batch(label):
    return {"id":datetime.now().strftime("%Y%m%d_%H%M%S_%f"),"label":label,"created":datetime.now().isoformat(timespec="seconds"),"moves":[]}
def record_move_batch(cfg,batch):
    if not batch.get("moves"): return
    cfg["last_moves"]=batch["moves"][-250:]
    history=cfg.setdefault("move_history",[])
    history.append(batch)
    cfg["move_history"]=history[-25:]
def log_smart_move(src,cfg,log):
    smart=smart_scan_details(src,cfg)
    if smart.get("destination"):
        log.append(f"[SMART]  {Path(src).name}  ->  {smart['category']}  ({smart['reason']})")
def scan_folder(folder_path,rules,excluded_folders=None):
    root=Path(os.path.normpath(folder_path)); out=[]
    if not root.exists(): return out
    excluded={x.lower() for x in (excluded_folders or [])}
    for dirpath,dirnames,filenames in os.walk(root):
        cur=Path(dirpath); dirnames[:]=[d for d in dirnames if not is_hidden(cur/d) and d.lower() not in excluded]
        for fn in filenames:
            item=cur/fn
            if should_skip(item): continue
            try: st=item.stat(); rel=item.relative_to(root)
            except Exception: continue
            out.append({"path":item,"rel":str(rel),"category":get_category(item,rules),"size":st.st_size,"modified":st.st_mtime})
    return sorted(out,key=lambda x:x["modified"],reverse=True)
def move_file(src,dest_dir,log):
    dest_dir.mkdir(parents=True,exist_ok=True); size=src.stat().st_size if src.exists() else 0; dest=dest_dir/src.name
    if dest.exists(): dest=dest_dir/f"{src.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{src.suffix}"
    try:
        shutil.move(str(src),str(dest)); entry=f"[{datetime.now().strftime('%H:%M:%S')}]  {src.name}  ->  {dest_dir}"; log.append(entry); return True,entry,str(dest),size
    except Exception as e:
        entry=f"[ERROR]  {src.name}: {e}"; log.append(entry); return False,entry,"",0
def restore_file(src,dest,log):
    src=Path(src); dest=Path(dest); dest.parent.mkdir(parents=True,exist_ok=True); size=src.stat().st_size if src.exists() else 0
    if dest.exists(): dest=dest.parent/f"{dest.stem}_restored_{datetime.now().strftime('%Y%m%d_%H%M%S')}{dest.suffix}"
    try:
        shutil.move(str(src),str(dest)); entry=f"[UNDO]  {src.name}  ->  {dest}"; log.append(entry); return True,entry,size
    except Exception as e:
        entry=f"[UNDO ERROR]  {src.name}: {e}"; log.append(entry); return False,entry,0
def app_dir():
    return Path(sys.executable).parent if getattr(sys,"frozen",False) else Path(__file__).parent
def open_manual():
    manual=app_dir()/MANUAL_FILE
    if manual.exists(): os.startfile(str(manual))
    else: messagebox.showwarning("Scanderbeg","The user manual could not be found.")
def fmt_size(b):
    b=int(b or 0)
    return f"{b} B" if b<1024 else f"{b/1024:.1f} KB" if b<1024**2 else f"{b/1024**2:.1f} MB" if b<1024**3 else f"{b/1024**3:.2f} GB"
def fmt_countdown(s):
    m,s=divmod(int(s),60); h,m=divmod(m,60); return f"{h}h {m}m" if h else f"{m}m {s}s" if m else f"{s}s"

class SweepEngine:
    def __init__(self,cfg,log_cb,status_cb,finish_cb=None):
        self.config=cfg; self.log_callback=log_cb; self.status_callback=status_cb; self.finish_callback=finish_cb; self.running=False; self._stop_event=threading.Event()
    def start(self):
        if self.running: return
        self._stop_event.clear(); threading.Thread(target=self._loop,daemon=True).start(); self.running=True
    def stop(self): self._stop_event.set(); self.running=False
    def sweep_now(self): threading.Thread(target=self._sweep,daemon=True).start()
    def _loop(self):
        while not self._stop_event.is_set():
            self._sweep(); delay=self.config.get("delay_minutes",30)*60; elapsed=0
            while elapsed<delay and not self._stop_event.is_set():
                time.sleep(5); elapsed+=5; self.status_callback(f"Next sweep in {fmt_countdown(delay-elapsed)}")
    def _sweep(self):
        self.status_callback("Sweeping..."); now=time.time(); delay=self.config.get("delay_minutes",30)*60
        first=self.config.setdefault("first_seen",{}); log=self.config.setdefault("log",[]); stats=self.config.setdefault("stats",{"files_moved":0,"bytes_moved":0,"sweeps":0})
        moved=locked=young=nodest=bytes_moved=0; batch=new_move_batch("Auto-sweep")
        for fp in self.config.get("watch_folders",[]):
            folder=Path(fp)
            if not folder.exists(): continue
            for item in folder.iterdir():
                if not item.is_file() or should_skip(item): continue
                key=str(item)
                if key not in first: first[key]=now; continue
                if now-first[key]<delay: young+=1; continue
                dest=resolve_destination(item,self.config)
                if not dest: nodest+=1; continue
                if is_file_locked(item): locked+=1; self.log_callback(f"[LOCKED]  {item.name} is in use, will retry next sweep."); continue
                ok,entry,to,size=move_file(item,Path(dest),log)
                if ok: first.pop(key,None); moved+=1; bytes_moved+=size; batch["moves"].append({"from":key,"to":to,"size":size}); log_smart_move(item,self.config,log); play_sound("SystemAsterisk"); self.log_callback(entry)
        for k in [k for k in list(first) if not Path(k).exists()]: first.pop(k,None)
        if moved: stats["files_moved"]=int(stats.get("files_moved",0))+moved; stats["bytes_moved"]=int(stats.get("bytes_moved",0))+bytes_moved; record_move_batch(self.config,batch)
        stats["sweeps"]=int(stats.get("sweeps",0))+1; save_config(self.config)
        parts=[f"Sweep done - {moved} moved"]
        if locked: parts.append(f"{locked} locked")
        if young: parts.append(f"{young} too new")
        if nodest: parts.append(f"{nodest} no destination")
        self.log_callback("  |  ".join(parts))
        if self.finish_callback: self.finish_callback()

class FileSheriffHandler(FileSystemEventHandler):
    def __init__(self,cfg,log_cb): self.config=cfg; self.log_callback=log_cb
    def on_created(self,event):
        if event.is_directory: return
        p=Path(event.src_path)
        if should_skip(p): return
        first=self.config.setdefault("first_seen",{}); key=str(p)
        if key not in first:
            first[key]=time.time(); self.log_callback(f"[SEEN]  {p.name} - will sort in {self.config.get('delay_minutes',30)} min if not in use."); save_config(self.config)

class ScanderbegApp(ctk.CTk,DND_MIXIN):
    def __init__(self):
        super().__init__()
        self.dnd_enabled=False
        if TKDND_AVAILABLE:
            try:
                self.TkdndVersion=TkinterDnD._require(self); self.dnd_enabled=True
            except Exception:
                self.dnd_enabled=False
        self.title("Scanderbeg"); self.geometry("1280x760"); self.minsize(1100,650)
        self.config_data=load_config(); self.observer=None; self.engine=None; self.scan_data=[]; self.checked={}; self.path_map={}; self.current_view=None
        self._glass_enabled=False; self._glass_mode="off"
        self.after(80,self._apply_windows_glass)
        self._build_ui(); self._show_view("dashboard"); self.after(500,self._maybe_show_first_run_tip)
    def _card(self,parent,**pk):
        sh=ctk.CTkFrame(parent,fg_color=(_SHADOW_LIGHT,_SHADOW_DARK),corner_radius=18); sh.pack(**pk)
        inner=ctk.CTkFrame(sh,fg_color=(_CARD_LIGHT,_CARD_DARK),corner_radius=16,border_width=1,border_color=("#FFFFFF","#274568")); inner.pack(fill="both",expand=pk.get("expand",False),padx=(0,3),pady=(0,4)); return inner
    def _apply_windows_glass(self):
        if os.name!="nt": return
        try:
            import ctypes
            hwnd=ctypes.windll.user32.GetParent(self.winfo_id()) or self.winfo_id()
            dark=ctypes.c_int(1 if ctk.get_appearance_mode()=="Dark" else 0)
            try:
                ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd,20,ctypes.byref(dark),ctypes.sizeof(dark))
            except Exception:
                pass
            class MARGINS(ctypes.Structure):
                _fields_=[("cxLeftWidth",ctypes.c_int),("cxRightWidth",ctypes.c_int),("cyTopHeight",ctypes.c_int),("cyBottomHeight",ctypes.c_int)]
            try:
                ctypes.windll.dwmapi.DwmExtendFrameIntoClientArea(hwnd,ctypes.byref(MARGINS(-1,-1,-1,-1)))
            except Exception:
                pass
            class ACCENTPOLICY(ctypes.Structure):
                _fields_=[("AccentState",ctypes.c_int),("AccentFlags",ctypes.c_int),("GradientColor",ctypes.c_int),("AnimationId",ctypes.c_int)]
            class WINCOMPATTRDATA(ctypes.Structure):
                _fields_=[("Attribute",ctypes.c_int),("Data",ctypes.c_void_p),("SizeOfData",ctypes.c_size_t)]
            # AccentState 4 is real Windows acrylic blur. The high byte is opacity in AABBGGRR form.
            tint=0xB8140E08 if ctk.get_appearance_mode()=="Dark" else 0xB8F8FBFF
            accent=ACCENTPOLICY(4,2,tint,0)
            data=WINCOMPATTRDATA(19,ctypes.cast(ctypes.pointer(accent),ctypes.c_void_p),ctypes.sizeof(accent))
            try:
                if ctypes.windll.user32.SetWindowCompositionAttribute(hwnd,ctypes.byref(data)):
                    self._glass_enabled=True; self._glass_mode="acrylic"; return
            except Exception:
                pass
            value=ctypes.c_int(2)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd,38,ctypes.byref(value),ctypes.sizeof(value))
            self._glass_enabled=True; self._glass_mode="mica"
        except Exception:
            self._glass_enabled=False; self._glass_mode="off"
    def _row_frame(self,parent,**pk):
        f=ctk.CTkFrame(parent,fg_color=(_ROW_LIGHT,_ROW_DARK),corner_radius=8)
        if pk: f.pack(**pk)
        return f
    def _build_ui(self):
        self.configure(fg_color=(_PAGE_LIGHT,_PAGE_DARK)); self.root=ctk.CTkFrame(self,fg_color=(_PAGE_LIGHT,_PAGE_DARK),corner_radius=0); self.root.pack(fill="both",expand=True)
        self.header=ctk.CTkFrame(self.root,height=64,corner_radius=0,fg_color=("#F8FBFF","#081526")); self.header.pack(side="top",fill="x"); self.header.pack_propagate(False)
        box=ctk.CTkFrame(self.header,fg_color="transparent"); box.pack(side="left",padx=18,pady=10)
        ctk.CTkLabel(box,text="Scanderbeg",font=ctk.CTkFont(size=26,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w")
        ctk.CTkLabel(box,text="File Organizer · v8 · Smart Scan  by Behar Terakaj",font=ctk.CTkFont(size=11),text_color=("#475569","#94A3B8")).pack(anchor="w")
        self.sweep_status_var=tk.StringVar(value="Auto-sweep: OFF"); ctk.CTkLabel(self.header,textvariable=self.sweep_status_var,font=ctk.CTkFont(size=12,weight="bold"),text_color=("#0F172A","#CBD5E1")).pack(side="right",padx=18)
        ctk.CTkLabel(self.header,text="🛡 Protected & Local",font=ctk.CTkFont(size=12,weight="bold"),text_color="#22c55e",fg_color=("#ECFDF5","#0f2a20"),corner_radius=8,padx=12,pady=6).pack(side="right",padx=8)
        self.body=ctk.CTkFrame(self.root,fg_color=(_PAGE_LIGHT,_PAGE_DARK),corner_radius=0); self.body.pack(fill="both",expand=True)
        self.sidebar=ctk.CTkFrame(self.body,width=220,corner_radius=0,fg_color=("#10233A","#071624")); self.sidebar.pack(side="left",fill="y"); self.sidebar.pack_propagate(False)
        self.content=ctk.CTkFrame(self.body,fg_color=(_PAGE_LIGHT,_PAGE_DARK),corner_radius=0); self.content.pack(side="left",fill="both",expand=True)
        self.nav_buttons={}; self.nav_meta={"dashboard":("dashboard","🏠"),"triage":("file_queue","📋"),"auto":("auto_sweep","⚙"),"rules":("rules","🧠"),"settings":("settings","🔧")}
        for view,(key,icon) in self.nav_meta.items(): self._add_sidebar_button(tx(self.config_data,key),view,icon)
        ctk.CTkFrame(self.sidebar,fg_color="transparent").pack(fill="both",expand=True)
        stor=ctk.CTkFrame(self.sidebar,fg_color=("#1B3450","#102033"),corner_radius=14,border_width=1,border_color=("#335B83","#274568")); stor.pack(fill="x",padx=16,pady=(0,16))
        ctk.CTkLabel(stor,text="FILES MOVED",font=ctk.CTkFont(size=11,weight="bold"),text_color=("#94A3B8","#CBD5E1")).pack(anchor="w",padx=14,pady=(14,4))
        self.files_moved_var=tk.StringVar(value="0"); ctk.CTkLabel(stor,textvariable=self.files_moved_var,font=ctk.CTkFont(size=26,weight="bold"),text_color="#F8FAFC").pack(anchor="w",padx=14)
        self.sweeps_var=tk.StringVar(value="0 sweeps"); ctk.CTkLabel(stor,textvariable=self.sweeps_var,font=ctk.CTkFont(size=11),text_color=("#94A3B8","#CBD5E1")).pack(anchor="w",padx=14,pady=(0,14))
        self.footer=ctk.CTkFrame(self.root,height=36,corner_radius=0,fg_color=("#F8FBFF","#081526")); self.footer.pack(side="bottom",fill="x"); self.footer.pack_propagate(False)
        self.footer_status_var=tk.StringVar(value="🛡 Your files are safe and organized."); self.footer_status_label=ctk.CTkLabel(self.footer,textvariable=self.footer_status_var,font=ctk.CTkFont(size=11),text_color=("#475569","#CBD5E1")); self.footer_status_label.pack(side="left",padx=18)
        ctk.CTkButton(self.footer,text="Export Log",height=24,width=90,**_BTN_SEC,command=self._export_log).pack(side="right",padx=8); self._update_sidebar_stats()
    def _add_sidebar_button(self,text,view,icon):
        b=ctk.CTkButton(self.sidebar,text=f"{icon}  {text}",height=42,anchor="w",fg_color="transparent",hover_color=("#1E293B","#10213a"),text_color=("#CBD5E1","#E2E8F0"),font=ctk.CTkFont(size=14),command=lambda:self._show_view(view)); b.pack(fill="x",padx=14,pady=4); self.nav_buttons[view]=b
    def _refresh_nav_labels(self):
        for view,btn in self.nav_buttons.items():
            key,icon=self.nav_meta.get(view,(view,""))
            btn.configure(text=f"{icon}  {tx(self.config_data,key)}")
    def _clear_content(self):
        for w in self.content.winfo_children(): w.destroy()
    def _show_view(self,name):
        self._clear_content(); self.current_view=name
        for v,b in self.nav_buttons.items(): b.configure(fg_color=("#1E293B","#123766") if v==name else "transparent",text_color="#fff" if v==name else "#CBD5E1",font=ctk.CTkFont(size=14,weight="bold" if v==name else "normal"))
        {"dashboard":self._build_dashboard,"triage":lambda:self._build_triage(self.content),"auto":lambda:self._build_watcher(self.content),"rules":lambda:self._build_rules(self.content),"settings":lambda:self._build_settings(self.content)}[name]()
    def _set_status_message(self,text,kind="info"):
        colors={"info":("#475569","#94a3b8"),"success":("#166534","#22c55e"),"warning":("#92400e","#f59e0b"),"error":("#991b1b","#ef4444")}; self.footer_status_var.set(text); self.footer_status_label.configure(text_color=colors.get(kind,colors["info"]))
        if hasattr(self,"_status_reset_job"):
            try: self.after_cancel(self._status_reset_job)
            except Exception: pass
        self._status_reset_job=self.after(4500,lambda:self.footer_status_var.set("🛡 Your files are safe and organized."))
    def _update_sidebar_stats(self):
        s=self.config_data.setdefault("stats",{"files_moved":0,"bytes_moved":0,"sweeps":0}); self.files_moved_var.set(str(s.get("files_moved",0))); n=s.get("sweeps",0); self.sweeps_var.set(f"{n} sweep{'s' if n!=1 else ''}")
    def _maybe_show_first_run_tip(self):
        if self.config_data.get("setup_tip_seen",False): return
        self.config_data["setup_tip_seen"]=True; save_config(self.config_data)
        if messagebox.askyesno(tx(self.config_data,"first_run_title"),tx(self.config_data,"first_run_body")+"\n\n"+tx(self.config_data,"open_settings")+"?"):
            self._show_view("settings")
    def _ask_destination_for_path(self,src):
        if not src: return ""
        cat=get_category(src,self.config_data.get("rules",DEFAULT_CATEGORIES))
        if not messagebox.askyesno("Scanderbeg",f"{tx(self.config_data,'no_destination_set')}:\n\n{cat}\n\n{tx(self.config_data,'open_settings')}?"):
            return ""
        d=filedialog.askdirectory(title=f"Choose folder for {cat}")
        if not d: return ""
        self.config_data.setdefault("category_destinations",{})[cat]=d; save_config(self.config_data)
        self._set_status_message(f"{cat} destination set.","success")
        return d
    def _get_queue_count(self):
        try: items=scan_folder(self.config_data.get("scan_folder",str(Path.home()/"Desktop")),self.config_data["rules"],self.config_data.get("excluded_folders",[]))
        except Exception: return 0
        ignored=set(self.config_data.get("ignored_files",[])); return sum(1 for e in items if str(e["path"]) not in ignored and not is_already_in_destination(e["path"],resolve_destination(e["path"],self.config_data)))
    def _build_dashboard(self):
        f=ctk.CTkScrollableFrame(self.content,fg_color=(_PAGE_LIGHT,_PAGE_DARK)); f.pack(fill="both",expand=True,padx=24,pady=22); q=self._get_queue_count(); s=self.config_data.setdefault("stats",{"files_moved":0,"bytes_moved":0,"sweeps":0})
        ctk.CTkLabel(f,text=tx(self.config_data,"all_clean") if q==0 else tx(self.config_data,"files_waiting_status",n=q),font=ctk.CTkFont(size=26,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w")
        ctk.CTkLabel(f,text=tx(self.config_data,"organized") if q==0 else tx(self.config_data,"queue_waiting"),font=ctk.CTkFont(size=13),text_color=("#475569","#94A3B8")).pack(anchor="w",pady=(2,16))
        row=ctk.CTkFrame(f,fg_color="transparent"); row.pack(fill="x",pady=(0,16))
        for title,val,sub in [(tx(self.config_data,"files_waiting"),q,tx(self.config_data,"current_queue")),(tx(self.config_data,"files_moved"),s.get("files_moved",0),tx(self.config_data,"this_install")),(tx(self.config_data,"auto_sweeps"),s.get("sweeps",0),tx(self.config_data,"total")),(tx(self.config_data,"watching"),len(self.config_data.get("watch_folders",[])),tx(self.config_data,"folders"))]: self._stat_card(row,title,str(val),sub).pack(side="left",fill="x",expand=True,padx=8)
        main=ctk.CTkFrame(f,fg_color="transparent"); main.pack(fill="both",expand=True); left=ctk.CTkFrame(main,fg_color="transparent"); left.pack(side="left",fill="both",expand=True,padx=(0,12)); right=ctk.CTkFrame(main,fg_color="transparent",width=320); right.pack(side="right",fill="y"); right.pack_propagate(False)
        quick=self._card(left,fill="x",pady=(0,12)); ctk.CTkLabel(quick,text=tx(self.config_data,"quick_actions"),font=ctk.CTkFont(size=16,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w",padx=16,pady=(14,8)); qr=ctk.CTkFrame(quick,fg_color="transparent"); qr.pack(fill="x",padx=16,pady=(0,16))
        ctk.CTkButton(qr,text=tx(self.config_data,"open_downloads"),height=38,**_BTN_SEC,command=lambda:os.startfile(str(Path.home()/"Downloads"))).pack(side="left",padx=(0,8)); ctk.CTkButton(qr,text=tx(self.config_data,"open_desktop"),height=38,**_BTN_SEC,command=lambda:os.startfile(str(Path.home()/"Desktop"))).pack(side="left",padx=8); ctk.CTkButton(qr,text=tx(self.config_data,"drop_zone") if self.dnd_enabled else tx(self.config_data,"quick_add"),height=38,**_BTN_PRI,command=self._open_drop_zone).pack(side="left",padx=8); ctk.CTkButton(qr,text="↩ "+tx(self.config_data,"undo_last_move"),height=38,**_BTN_DNG,command=self._undo_last_move).pack(side="left",padx=8)
        log_card=self._card(left,fill="both",expand=True); ctk.CTkLabel(log_card,text=tx(self.config_data,"recent_activity"),font=ctk.CTkFont(size=16,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w",padx=16,pady=(14,8)); self.dashboard_log=ctk.CTkTextbox(log_card,height=260,font=("Consolas",10),wrap="word",fg_color=(_ROW_LIGHT,"#111827"),text_color=("#1E293B","#E5E7EB")); self.dashboard_log.pack(fill="both",expand=True,padx=16,pady=(0,16)); self._fill_textbox(self.dashboard_log,"\n".join(reversed(self.config_data.get("log",[])[-30:])) or "No activity yet.")
        self._dashboard_trust_card(right); self._dashboard_auto_card(right); self._dashboard_watch_card(right)
    def _stat_card(self,p,title,value,sub):
        sh=ctk.CTkFrame(p,fg_color=(_SHADOW_LIGHT,_SHADOW_DARK),corner_radius=16); c=ctk.CTkFrame(sh,fg_color=(_CARD_LIGHT,_CARD_DARK),corner_radius=14); c.pack(fill="both",expand=True,padx=(0,3),pady=(0,4)); ctk.CTkLabel(c,text=title,font=ctk.CTkFont(size=13,weight="bold"),text_color=("#475569","#94A3B8")).pack(anchor="w",padx=16,pady=(16,0)); ctk.CTkLabel(c,text=value,font=ctk.CTkFont(size=34,weight="bold"),text_color=("#0F172A","#fff")).pack(anchor="w",padx=16); ctk.CTkLabel(c,text=sub,font=ctk.CTkFont(size=12),text_color=("#64748B","#CBD5E1")).pack(anchor="w",padx=16,pady=(0,16)); return sh
    def _dashboard_auto_card(self,p):
        card=self._card(p,fill="x",pady=(0,12)); ctk.CTkLabel(card,text=tx(self.config_data,"smart_scan"),font=ctk.CTkFont(size=16,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w",padx=14,pady=(14,6)); state="ON" if self.config_data.get("smart_scan_enabled") else "OFF"; ctk.CTkLabel(card,text=state,font=ctk.CTkFont(size=22,weight="bold"),text_color=("#16A34A","#22C55E") if state=="ON" else ("#B91C1C","#FCA5A5")).pack(pady=8); r=ctk.CTkFrame(card,fg_color="transparent"); r.pack(fill="x",padx=14,pady=(0,14)); ctk.CTkButton(r,text="🔎 "+tx(self.config_data,"scan_now"),**_BTN_PRI,command=self._scan_now_dashboard).pack(side="left",fill="x",expand=True,padx=(0,4)); ctk.CTkButton(r,text=tx(self.config_data,"smart_scan")+" ›",**_BTN_SEC,command=lambda:self._show_view("settings")).pack(side="left",fill="x",expand=True,padx=(4,0))
    def _dashboard_trust_card(self,p):
        card=self._card(p,fill="x",pady=(0,12))
        ctk.CTkLabel(card,text="Trust",font=ctk.CTkFont(size=16,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w",padx=14,pady=(14,6))
        for key in ("trust_local","trust_offline","trust_undo"):
            ctk.CTkLabel(card,text="✓ "+tx(self.config_data,key),font=ctk.CTkFont(size=12,weight="bold"),text_color=("#166534","#22c55e")).pack(anchor="w",padx=14,pady=2)
        ctk.CTkLabel(card,text=f"Support: {SUPPORT_EMAIL}",font=ctk.CTkFont(size=10),text_color=("#64748B","#94A3B8")).pack(anchor="w",padx=14,pady=(6,14))
    def _scan_now_dashboard(self):
        self._show_view("triage")
        self._do_scan()
        self._set_status_message(tx(self.config_data,"scan_now"),"success")
    def _dashboard_watch_card(self,p):
        card=self._card(p,fill="both",expand=True); ctk.CTkLabel(card,text=tx(self.config_data,"watched_folders"),font=ctk.CTkFont(size=16,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w",padx=14,pady=(14,8))
        for folder in self.config_data.get("watch_folders",[]): row=self._row_frame(card,fill="x",padx=14,pady=4); ctk.CTkLabel(row,text=f"📁 {folder}",font=ctk.CTkFont(size=11),text_color=("#1E293B","#E2E8F0"),anchor="w").pack(side="left",fill="x",expand=True,padx=10,pady=8)
        ctk.CTkButton(card,text=tx(self.config_data,"add_folder"),height=30,**_BTN_SEC,command=self._add_watch).pack(anchor="w",padx=14,pady=12)
    def _open_drop_zone(self):
        if hasattr(self,"drop_zone") and self.drop_zone.winfo_exists():
            self.drop_zone.focus(); return
        title=f"Scanderbeg {tx(self.config_data,'drop_zone')}" if self.dnd_enabled else f"Scanderbeg {tx(self.config_data,'quick_add')}"
        self.drop_zone=ctk.CTkToplevel(self); self.drop_zone.title(title); self.drop_zone.geometry("560x420"); self.drop_zone.attributes("-topmost",True); self.drop_zone.configure(fg_color=(_PAGE_LIGHT,_PAGE_DARK))
        self.drop_files=[]
        ctk.CTkLabel(self.drop_zone,text=tx(self.config_data,"drop_zone") if self.dnd_enabled else tx(self.config_data,"quick_add"),font=ctk.CTkFont(size=22,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w",padx=18,pady=(16,4))
        helper="Drop files here or add them to preview where Scanderbeg will sort them." if self.dnd_enabled else "Add files to preview where Scanderbeg will sort them."
        ctk.CTkLabel(self.drop_zone,text=helper,font=ctk.CTkFont(size=12),text_color=("#475569","#94A3B8")).pack(anchor="w",padx=18,pady=(0,12))
        bar=ctk.CTkFrame(self.drop_zone,fg_color="transparent"); bar.pack(fill="x",padx=18,pady=(0,10))
        ctk.CTkButton(bar,text=tx(self.config_data,"add_files"),**_BTN_PRI,command=self._drop_add_files).pack(side="left",padx=(0,8))
        ctk.CTkButton(bar,text=tx(self.config_data,"sort_listed"),**_BTN_SEC,command=self._drop_sort_files).pack(side="left",padx=8)
        ctk.CTkButton(bar,text=tx(self.config_data,"clear"),**_BTN_DNG,command=self._drop_clear_files).pack(side="left",padx=8)
        self.drop_box=ctk.CTkTextbox(self.drop_zone,font=("Consolas",10),wrap="word",fg_color=(_ROW_LIGHT,"#111827"),text_color=("#1E293B","#E5E7EB"))
        self.drop_box.pack(fill="both",expand=True,padx=18,pady=(0,18))
        self._register_drop_target(self.drop_zone); self._register_drop_target(self.drop_box)
        self._refresh_drop_zone()
    def _register_drop_target(self,widget):
        if not self.dnd_enabled: return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>",self._drop_files_event)
        except Exception:
            pass
    def _drop_files_event(self,event):
        for fp in self.tk.splitlist(event.data):
            p=Path(fp)
            if p.exists() and p.is_file() and p not in self.drop_files: self.drop_files.append(p)
        self._refresh_drop_zone()
        return "copy"
    def _drop_add_files(self):
        files=filedialog.askopenfilenames(title="Add files to Quick Add")
        for fp in files:
            p=Path(fp)
            if p.exists() and p.is_file() and p not in self.drop_files: self.drop_files.append(p)
        self._refresh_drop_zone()
    def _drop_clear_files(self):
        self.drop_files=[]; self._refresh_drop_zone()
    def _refresh_drop_zone(self):
        if not hasattr(self,"drop_box") or not self.drop_box.winfo_exists(): return
        lines=[]
        for p in self.drop_files:
            dest=resolve_destination(p,self.config_data); smart=smart_scan_details(p,self.config_data)
            why=f" | {smart['reason']}" if smart.get("reason") else ""
            lines.append(f"{p.name}\n  -> {dest or tx(self.config_data,'no_destination_set')}{why}")
        self._fill_textbox(self.drop_box,"\n\n".join(lines) if lines else tx(self.config_data,"no_files_added"))
    def _drop_sort_files(self):
        if not getattr(self,"drop_files",[]): self._set_status_message(tx(self.config_data,"quick_add_empty"),"warning"); return
        log=self.config_data.setdefault("log",[]); stats=self.config_data.setdefault("stats",{"files_moved":0,"bytes_moved":0,"sweeps":0}); moved=skipped=errors=locked=bytes_moved=0; batch=new_move_batch("Quick Add"); remaining=[]
        for src in self.drop_files:
            dest=resolve_destination(src,self.config_data)
            if not dest:
                dest=self._ask_destination_for_path(src)
                if not dest: skipped+=1; remaining.append(src); continue
            if not src.exists(): errors+=1; continue
            if is_file_locked(src): locked+=1; remaining.append(src); continue
            ok,_,to,size=move_file(src,Path(dest),log)
            if ok: moved+=1; bytes_moved+=size; batch["moves"].append({"from":str(src),"to":to,"size":size}); log_smart_move(src,self.config_data,log)
            else: errors+=1; remaining.append(src)
        if moved:
            stats["files_moved"]=int(stats.get("files_moved",0))+moved; stats["bytes_moved"]=int(stats.get("bytes_moved",0))+bytes_moved; record_move_batch(self.config_data,batch); play_sound("SystemAsterisk")
        self.drop_files=remaining; save_config(self.config_data); self._update_sidebar_stats(); self._refresh_drop_zone()
        msg=f"{moved} moved"
        if skipped: msg+=f" • {skipped} no destination"
        if locked: msg+=f" • {locked} locked"
        if errors: msg+=f" • {errors} error(s)"
        self._set_status_message(msg,"success" if moved and not errors and not locked else "warning")
    def _fill_textbox(self,b,t): b.configure(state="normal"); b.delete("1.0","end"); b.insert("end",t); b.configure(state="disabled")
    def _build_triage(self,parent):
        f=ctk.CTkFrame(parent,fg_color=(_PAGE_LIGHT,_PAGE_DARK)); f.pack(fill="both",expand=True,padx=24,pady=22)
        head=ctk.CTkFrame(f,fg_color="transparent"); head.pack(fill="x"); ctk.CTkLabel(head,text=tx(self.config_data,"file_queue"),font=ctk.CTkFont(size=24,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(side="left"); ctk.CTkButton(head,text=tx(self.config_data,"run_scan"),width=110,**_BTN_PRI,command=self._do_scan).pack(side="right"); ctk.CTkButton(head,text=tx(self.config_data,"browse"),width=90,**_BTN_SEC,command=self._browse_scan).pack(side="right",padx=8)
        ctrl=self._card(f,fill="x",pady=(16,10)); self.scan_var=tk.StringVar(value=self.config_data.get("scan_folder",str(Path.home()/"Desktop"))); ctk.CTkLabel(ctrl,text=tx(self.config_data,"scan_folder"),font=ctk.CTkFont(size=12),text_color=("#334155","#CBD5E1")).pack(side="left",padx=(14,8),pady=12); ctk.CTkEntry(ctrl,textvariable=self.scan_var,height=34).pack(side="left",fill="x",expand=True,padx=(0,14),pady=12)
        action=ctk.CTkFrame(f,fg_color="transparent"); action.pack(fill="x",pady=(0,8))
        # FIX: pack the controls on the right first, then let summary text use remaining space.
        right_actions=ctk.CTkFrame(action,fg_color="transparent"); right_actions.pack(side="right",anchor="e")
        ctk.CTkButton(right_actions,text="⚡ "+tx(self.config_data,"sweep_selected"),width=150,height=34,**_BTN_PRI,command=self._sort_selected).pack(side="left",padx=(0,8))
        self.action_labels={"more_actions":tx(self.config_data,"more_actions"),"select_all":tx(self.config_data,"select_all"),"deselect_all":tx(self.config_data,"deselect_all"),"ignore_selected":tx(self.config_data,"ignore_selected"),"clear_ignored_files":tx(self.config_data,"clear_ignored_files"),"open_file_location":tx(self.config_data,"open_file_location"),"move_custom":tx(self.config_data,"move_custom"),"set_dest_category":tx(self.config_data,"set_dest_category"),"undo_last_move":tx(self.config_data,"undo_last_move"),"rescan_folder":tx(self.config_data,"rescan_folder"),"clear_activity_log":tx(self.config_data,"clear_activity_log")}
        self.file_queue_action_menu=ctk.CTkOptionMenu(right_actions,width=190,height=34,values=list(self.action_labels.values()),command=self._file_queue_action,fg_color=(_ROW_LIGHT,"#1E293B"),button_color=("#CBD5E1","#334155"),button_hover_color=("#94A3B8","#475569"),dropdown_fg_color=(_CARD_LIGHT,_CARD_DARK),dropdown_hover_color=("#E2E8F0","#172033"),text_color=("#1E293B","#E2E8F0"),corner_radius=10)
        self.file_queue_action_menu.set(self.action_labels["more_actions"]); self.file_queue_action_menu.pack(side="left",padx=(0,8))
        self.selected_count_var=tk.StringVar(value="0 selected"); ctk.CTkLabel(right_actions,textvariable=self.selected_count_var,font=ctk.CTkFont(size=11,weight="bold"),text_color=("#475569","#94A3B8")).pack(side="left")
        self.summary_var=tk.StringVar(value=tx(self.config_data,"queue_empty")); ctk.CTkLabel(action,textvariable=self.summary_var,font=ctk.CTkFont(size=11),text_color=("#475569","#94A3B8"),anchor="w",justify="left").pack(side="left",fill="x",expand=True,padx=(0,12))
        wrap=self._card(f,fill="both",expand=True); cols=("check","filename","category","destination","confidence","reason","size","modified","status"); self.tree=ttk.Treeview(wrap,columns=cols,show="headings",selectmode="none")
        for col,title in {"check":"✓","filename":tx(self.config_data,"file_name"),"category":tx(self.config_data,"category"),"destination":tx(self.config_data,"goes_to"),"confidence":tx(self.config_data,"confidence"),"reason":tx(self.config_data,"reason"),"size":tx(self.config_data,"size"),"modified":tx(self.config_data,"modified"),"status":tx(self.config_data,"status")}.items(): self.tree.heading(col,text=title)
        for col,w,stretch,anchor in [("check",40,False,"center"),("filename",240,True,"w"),("category",105,False,"w"),("destination",230,True,"w"),("confidence",95,False,"center"),("reason",210,True,"w"),("size",85,False,"e"),("modified",130,False,"w"),("status",95,False,"center")]: self.tree.column(col,width=w,stretch=stretch,anchor=anchor)
        vsb=ctk.CTkScrollbar(wrap,command=self.tree.yview); self.tree.configure(yscrollcommand=vsb.set); self.tree.pack(side="left",fill="both",expand=True,padx=(10,0),pady=10); vsb.pack(side="right",fill="y",padx=(0,10),pady=10); self.tree.bind("<ButtonRelease-1>",self._toggle_row); self._restyle_tree(); self._do_scan()
    def _restyle_tree(self):
        st=ttk.Style()
        try: st.theme_use("clam")
        except Exception: pass
        dark=ctk.get_appearance_mode()=="Dark"; bg,fg,hbg,hfg,sel=("#0f1b2d","#f1f5f9","#0b1220","#94a3b8","#123766") if dark else ("#fff","#1E293B","#F1F5F9","#0F172A","#DBEAFE")
        st.configure("Treeview",background=bg,foreground=fg,fieldbackground=bg,borderwidth=0,rowheight=30,font=("Segoe UI",10)); st.configure("Treeview.Heading",background=hbg,foreground=hfg,borderwidth=0,font=("Segoe UI",10,"bold")); st.map("Treeview",background=[("selected",sel)]); st.layout("Treeview",[("Treeview.treearea",{"sticky":"nswe"})])
        if hasattr(self,"tree"):
            for tag,color in [("nodest","#ef4444"),("ready","#16a34a"),("waiting","#d97706"),("smart","#8b5cf6")]: self.tree.tag_configure(tag,foreground=color)
    def _browse_scan(self):
        d=filedialog.askdirectory()
        if d: self.scan_var.set(d)
    def _do_scan(self):
        if not hasattr(self,"tree"): return
        folder=os.path.normpath(self.scan_var.get()); self.scan_var.set(folder); self.config_data["scan_folder"]=folder; save_config(self.config_data); self.tree.delete(*self.tree.get_children()); self.checked.clear(); self.path_map.clear()
        self.scan_data=scan_folder(folder,self.config_data["rules"],self.config_data.get("excluded_folders",[])); now=time.time(); delay=self.config_data.get("delay_minutes",30)*60; first=self.config_data.get("first_seen",{}); ignored=set(self.config_data.get("ignored_files",[])); counts={}; no_dest=[]; visible=0
        for e in self.scan_data:
            p=e["path"]
            if str(p) in ignored: continue
            dest=resolve_destination(p,self.config_data)
            if is_already_in_destination(p,dest): continue
            smart=smart_scan_details(p,self.config_data)
            if smart["destination"] and dest==smart["destination"]: dest_lbl=f"[Smart] {dest}"; tag="smart"
            elif dest: dest_lbl=dest; tag=""
            else: dest_lbl=tx(self.config_data,"no_destination_set"); tag="nodest"; no_dest.append(e["category"])
            key=str(p)
            if key in first:
                rem=delay-(now-first[key]); status="Ready" if rem<=0 else fmt_countdown(rem); tag=tag or ("ready" if rem<=0 else "waiting")
            else: status="New"
            iid=self.tree.insert("","end",values=("☐",e.get("rel",p.name),e["category"],dest_lbl,smart.get("confidence",""),smart.get("reason",""),fmt_size(e["size"]),datetime.fromtimestamp(e["modified"]).strftime("%Y-%m-%d %H:%M"),status),tags=(tag,)); self.checked[iid]=False; self.path_map[iid]=p; visible+=1; counts[e["category"]]=counts.get(e["category"],0)+1
        if not self.scan_data: self.summary_var.set(tx(self.config_data,"folder_clean"))
        elif visible==0: self.summary_var.set(f"All {len(self.scan_data)} files already sorted.")
        else:
            summary=tx(self.config_data,"files_to_sort",n=visible); filtered=len(self.scan_data)-visible
            if filtered: summary+=f"  •  {tx(self.config_data,'already_sorted',n=filtered)}"
            summary+="  •  "+"  ".join(f"{c}: {n}" for c,n in sorted(counts.items()))
            if no_dest: summary+=f"  •  {tx(self.config_data,'set_dest_for',cats=', '.join(sorted(set(no_dest))))}"
            self.summary_var.set(summary)
        self._update_selected_count()
    def _toggle_row(self,event):
        row=self.tree.identify_row(event.y)
        if not row: return
        self.checked[row]=not self.checked.get(row,False); vals=list(self.tree.item(row,"values")); vals[0]="☑" if self.checked[row] else "☐"; self.tree.item(row,values=vals); self._update_selected_count()
    def _update_selected_count(self):
        if hasattr(self,"selected_count_var"): self.selected_count_var.set(f"{sum(1 for v in self.checked.values() if v)} selected  •  {len(self.checked)} shown")
    def _selected_iids(self): return [i for i,v in self.checked.items() if v and self.tree.exists(i)]
    def _selected_paths(self): return [self.path_map[i] for i in self._selected_iids() if i in self.path_map]
    def _set_all_selected(self):
        for i in list(self.checked): self.checked[i]=True; vals=list(self.tree.item(i,"values")); vals[0]="☑"; self.tree.item(i,values=vals)
        self._update_selected_count()
    def _deselect_all(self):
        for i in list(self.checked): self.checked[i]=False; vals=list(self.tree.item(i,"values")); vals[0]="☐"; self.tree.item(i,values=vals)
        self._update_selected_count()
    def _file_queue_action(self,choice):
        try:
            actions={self.action_labels["select_all"]:self._set_all_selected,self.action_labels["deselect_all"]:self._deselect_all,self.action_labels["ignore_selected"]:self._ignore_selected,self.action_labels["clear_ignored_files"]:self._clear_ignored_files,self.action_labels["open_file_location"]:self._open_selected_location,self.action_labels["move_custom"]:self._move_selected_to_custom_folder,self.action_labels["set_dest_category"]:self._set_destination_for_selected_category,self.action_labels["undo_last_move"]:self._undo_last_move,self.action_labels["rescan_folder"]:self._do_scan,self.action_labels["clear_activity_log"]:self._clear_log}
            actions.get(choice,lambda:None)()
        finally: self.file_queue_action_menu.set(self.action_labels["more_actions"])
    def _open_selected_location(self):
        paths=self._selected_paths()
        if not paths: self._set_status_message("Select one file first.","warning"); return
        import subprocess; p=paths[0]; subprocess.Popen(["explorer","/select,",str(p)] if p.exists() else ["explorer",str(p.parent)]); self._set_status_message("Opened file location.","success")
    def _move_selected_to_custom_folder(self):
        sel=self._selected_iids()
        if not sel: self._set_status_message(tx(self.config_data,"no_files_selected"),"warning"); return
        d=filedialog.askdirectory(title="Move selected files to folder")
        if d: self._move_iids(sel,Path(d),custom=True)
    def _set_destination_for_selected_category(self):
        paths=self._selected_paths()
        if not paths: self._set_status_message("Select one file first.","warning"); return
        cat=get_category(paths[0],self.config_data.get("rules",DEFAULT_CATEGORIES)); d=filedialog.askdirectory(title=f"Set destination for {cat}")
        if d: self.config_data.setdefault("category_destinations",{})[cat]=d; save_config(self.config_data); self._set_status_message(f"{cat} destination set.","success"); self._do_scan()
    def _ignore_selected(self):
        sel=self._selected_iids()
        if not sel: self._set_status_message("No files selected.","warning"); return
        ignored=self.config_data.setdefault("ignored_files",[]); added=0
        for i in sel:
            p=self.path_map.get(i)
            if p and str(p) not in ignored: ignored.append(str(p)); added+=1
            if self.tree.exists(i): self.tree.delete(i)
            self.checked.pop(i,None); self.path_map.pop(i,None)
        self.config_data.setdefault("log",[]).append(f"[{datetime.now().strftime('%H:%M:%S')}]  Ignored {added} file(s) from File Queue"); save_config(self.config_data); self._update_selected_count(); self._set_status_message(f"{added} file(s) ignored.","success")
    def _move_iids(self,sel,dest_dir,custom=False):
        log=self.config_data.setdefault("log",[]); stats=self.config_data.setdefault("stats",{"files_moved":0,"bytes_moved":0,"sweeps":0}); moved=errors=locked=bytes_moved=0; batch=new_move_batch("Custom move" if custom else "Move selected")
        for i in sel:
            src=self.path_map.get(i)
            if not src or not src.exists(): errors+=1; continue
            if is_file_locked(src): locked+=1; continue
            ok,_,to,size=move_file(src,dest_dir,log)
            if ok: moved+=1; bytes_moved+=size; batch["moves"].append({"from":str(src),"to":to,"size":size}); log_smart_move(src,self.config_data,log); self.tree.delete(i); self.checked.pop(i,None); self.path_map.pop(i,None)
            else: errors+=1
        if moved: stats["files_moved"]=int(stats.get("files_moved",0))+moved; stats["bytes_moved"]=int(stats.get("bytes_moved",0))+bytes_moved; record_move_batch(self.config_data,batch); play_sound("SystemAsterisk")
        save_config(self.config_data); self._update_sidebar_stats(); self._update_selected_count(); self._set_status_message(f"{moved} file(s) moved"+(" to custom folder." if custom else "."),"success" if moved and not errors and not locked else "warning")
    def _undo_last_move(self):
        history=self.config_data.setdefault("move_history",[])
        batch=history.pop() if history else {"moves":self.config_data.get("last_moves",[]),"label":"Last move"}
        moves=list(reversed(batch.get("moves",[])))
        if not moves:
            self._set_status_message("Nothing to undo yet.","warning"); return
        log=self.config_data.setdefault("log",[]); stats=self.config_data.setdefault("stats",{"files_moved":0,"bytes_moved":0,"sweeps":0}); restored=missing=locked=errors=bytes_restored=0
        for move in moves:
            src=Path(move.get("to","")); dest=Path(move.get("from",""))
            if not src.exists(): missing+=1; continue
            if is_file_locked(src): locked+=1; continue
            ok,_,size=restore_file(src,dest,log)
            if ok: restored+=1; bytes_restored+=size
            else: errors+=1
        if restored:
            stats["files_moved"]=max(0,int(stats.get("files_moved",0))-restored); stats["bytes_moved"]=max(0,int(stats.get("bytes_moved",0))-bytes_restored)
        self.config_data["last_moves"]=history[-1].get("moves",[]) if history else []; save_config(self.config_data); self._update_sidebar_stats(); play_sound("SystemAsterisk" if restored else "SystemHand")
        parts=[f"{restored} file(s) restored from {batch.get('label','last batch')}"]
        if missing: parts.append(f"{missing} missing")
        if locked: parts.append(f"{locked} locked")
        if errors: parts.append(f"{errors} error(s)")
        msg=" • ".join(parts); self._set_status_message(msg,"success" if restored and not errors and not locked else "warning")
        if hasattr(self,"summary_var"): self.summary_var.set(msg); self._do_scan()
    def _sort_selected(self):
        sel=self._selected_iids()
        if not sel: self._set_status_message("No files selected.","warning"); return
        log=self.config_data.setdefault("log",[]); stats=self.config_data.setdefault("stats",{"files_moved":0,"bytes_moved":0,"sweeps":0}); moved=errors=skipped=locked=bytes_moved=0; batch=new_move_batch("Sweep selected")
        for i in sel:
            src=self.path_map.get(i); dest=resolve_destination(src,self.config_data) if src else ""
            if not src: errors+=1; continue
            if not dest:
                dest=self._ask_destination_for_path(src)
                if not dest: skipped+=1; continue
            if not src.exists(): errors+=1; continue
            if is_file_locked(src): locked+=1; continue
            ok,_,to,size=move_file(src,Path(dest),log)
            if ok: moved+=1; bytes_moved+=size; batch["moves"].append({"from":str(src),"to":to,"size":size}); log_smart_move(src,self.config_data,log); self.tree.delete(i); self.checked.pop(i,None); self.path_map.pop(i,None)
            else: errors+=1
        if moved: stats["files_moved"]=int(stats.get("files_moved",0))+moved; stats["bytes_moved"]=int(stats.get("bytes_moved",0))+bytes_moved; record_move_batch(self.config_data,batch)
        save_config(self.config_data); self._update_sidebar_stats(); self._update_selected_count(); play_sound("SystemAsterisk" if moved else "SystemHand")
        parts=[f"{moved} file(s) moved"]
        if locked: parts.append(f"{locked} locked")
        if skipped: parts.append(f"{skipped} skipped")
        if errors: parts.append(f"{errors} error(s)")
        msg=" • ".join(parts); self.summary_var.set(msg); self._set_status_message(msg,"success" if not errors and not locked else "warning"); self._do_scan()
    def _build_watcher(self,parent):
        f=ctk.CTkScrollableFrame(parent,fg_color=(_PAGE_LIGHT,_PAGE_DARK)); f.pack(fill="both",expand=True,padx=24,pady=22); ctk.CTkLabel(f,text=tx(self.config_data,"auto_sweep"),font=ctk.CTkFont(size=24,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w")
        if not WATCHDOG_AVAILABLE: ctk.CTkLabel(f,text="watchdog not installed.\nRun: pip install watchdog",text_color="#ef4444").pack(pady=40); return
        card=self._card(f,fill="x",pady=(16,12)); row=ctk.CTkFrame(card,fg_color="transparent"); row.pack(fill="x",padx=16,pady=14); on=self.engine is not None and self.engine.running
        self.auto_status=ctk.CTkLabel(row,text=tx(self.config_data,"auto_status_on") if on else tx(self.config_data,"auto_status_off"),font=ctk.CTkFont(size=15,weight="bold"),text_color="#16a34a" if on else "#b91c1c"); self.auto_status.pack(side="left")
        ctk.CTkButton(row,text="⚡ "+tx(self.config_data,"scan_now"),width=120,**_BTN_SEC,command=self._sweep_now).pack(side="right",padx=4); self.auto_btn=ctk.CTkButton(row,text=tx(self.config_data,"stop_auto") if on else tx(self.config_data,"start_auto"),width=170,fg_color="#ef4444" if on else "#10b981",hover_color="#dc2626" if on else "#059669",text_color="#fff",command=self._toggle_auto); self.auto_btn.pack(side="right",padx=4)
        self.delay_lbl=ctk.CTkLabel(card,text=tx(self.config_data,"delay_settings",n=self.config_data.get("delay_minutes",30)),font=ctk.CTkFont(size=11),text_color=("#475569","#CBD5E1")); self.delay_lbl.pack(anchor="w",padx=16,pady=(0,14))
        wf=self._card(f,fill="x",pady=12); ctk.CTkLabel(wf,text=tx(self.config_data,"watched_folders_lower"),font=ctk.CTkFont(size=15,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w",padx=16,pady=(14,8)); self.watch_list_frame=ctk.CTkScrollableFrame(wf,height=110,fg_color="transparent"); self.watch_list_frame.pack(fill="x",padx=12,pady=(0,8)); self._refresh_watch_list(); ctk.CTkButton(wf,text=tx(self.config_data,"add_folder"),width=140,**_BTN_SEC,command=self._add_watch).pack(anchor="w",padx=12,pady=(0,14))
        lc=self._card(f,fill="both",expand=True,pady=(12,0)); ctk.CTkLabel(lc,text=tx(self.config_data,"activity_log"),font=ctk.CTkFont(size=15,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w",padx=16,pady=(14,6)); self.log_box=ctk.CTkTextbox(lc,font=("Consolas",10),wrap="none",height=260,fg_color=(_ROW_LIGHT,"#111827"),text_color=("#1E293B","#E5E7EB")); self.log_box.pack(fill="both",expand=True,padx=12,pady=(0,12)); self._refresh_log()
    def _refresh_watch_list(self):
        if not hasattr(self,"watch_list_frame"): return
        for w in self.watch_list_frame.winfo_children(): w.destroy()
        for folder in self.config_data.get("watch_folders",[]): row=self._row_frame(self.watch_list_frame,fill="x",pady=3); ctk.CTkLabel(row,text=folder,font=ctk.CTkFont(size=11),anchor="w",text_color=("#1E293B","#E2E8F0")).pack(side="left",padx=10,fill="x",expand=True,pady=8); ctk.CTkButton(row,text="✕",width=28,height=24,**_BTN_DNG,command=lambda fp=folder:self._remove_watch(fp)).pack(side="right",padx=4)
    def _toggle_auto(self): self._stop_auto() if self.engine and self.engine.running else self._start_auto()
    def _start_auto(self):
        if not WATCHDOG_AVAILABLE: messagebox.showerror("Scanderbeg","Run: pip install watchdog"); return
        if not self.config_data.get("watch_folders"): messagebox.showwarning("Scanderbeg","Add at least one folder to watch."); return
        self.engine=SweepEngine(self.config_data,self._async_log,self._async_status,self._async_finish); self.engine.start(); self.observer=Observer(); h=FileSheriffHandler(self.config_data,self._async_log)
        for folder in self.config_data["watch_folders"]:
            if Path(folder).exists(): self.observer.schedule(h,folder,recursive=False)
        self.observer.start(); self.sweep_status_var.set("Auto-sweep: ON"); self._show_view("auto")
    def _stop_auto(self):
        if self.engine: self.engine.stop(); self.engine=None
        if self.observer: self.observer.stop(); self.observer.join(); self.observer=None
        self.sweep_status_var.set("Auto-sweep: OFF"); self._show_view("auto")
    def _sweep_now(self):
        if self.engine: self.engine.sweep_now()
        else: threading.Thread(target=SweepEngine(self.config_data,self._async_log,self._async_status,self._async_finish)._sweep,daemon=True).start()
    def _add_watch(self):
        d=filedialog.askdirectory()
        if d and d not in self.config_data["watch_folders"]: self.config_data["watch_folders"].append(d); save_config(self.config_data); self._show_view(self.current_view)
    def _remove_watch(self,folder):
        if folder in self.config_data["watch_folders"]: self.config_data["watch_folders"].remove(folder); save_config(self.config_data); self._show_view(self.current_view)
    def _async_log(self,e): self.config_data.setdefault("log",[]).append(e); save_config(self.config_data); self.after(0,self._refresh_log)
    def _async_status(self,t): self.after(0,lambda:self.sweep_status_var.set(f"Auto-sweep: {t}")); self.after(0,lambda:getattr(self,"dashboard_countdown_var",tk.StringVar()).set(t))
    def _async_finish(self): self.after(0,self._update_sidebar_stats)
    def _refresh_log(self):
        if hasattr(self,"log_box") and self.log_box.winfo_exists():
            self._fill_textbox(self.log_box,"\n".join(reversed(self.config_data.get("log",[])[-150:])) or "No activity yet.")
    def _clear_log(self): self.config_data["log"]=[]; save_config(self.config_data); self._refresh_log()
    def _build_rules(self,parent):
        f=ctk.CTkScrollableFrame(parent,fg_color=(_PAGE_LIGHT,_PAGE_DARK)); f.pack(fill="both",expand=True,padx=24,pady=22)
        head=ctk.CTkFrame(f,fg_color="transparent"); head.pack(fill="x")
        ctk.CTkLabel(head,text=tx(self.config_data,"rules"),font=ctk.CTkFont(size=24,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(side="left")
        ctk.CTkButton(head,text=tx(self.config_data,"save_rules"),width=140,**_BTN_PRI,command=self._save_rules).pack(side="right")
        ctk.CTkButton(head,text=tx(self.config_data,"reset_defaults"),width=150,**_BTN_SEC,command=self._reset_rules).pack(side="right",padx=(0,8))
        ctk.CTkLabel(f,text=tx(self.config_data,"file_types_intro"),font=ctk.CTkFont(size=13),text_color=("#475569","#94A3B8")).pack(anchor="w",pady=(4,16))
        add_card=self._card(f,fill="x",pady=(0,12))
        add_row=ctk.CTkFrame(add_card,fg_color="transparent"); add_row.pack(fill="x",padx=14,pady=14)
        ctk.CTkLabel(add_row,text=tx(self.config_data,"add_file_type_group"),font=ctk.CTkFont(size=13,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(side="left",padx=(0,10))
        available=[c for c in OPTIONAL_CATEGORY_PRESETS if c not in self.config_data.get("rules",{})]
        self.rule_preset_var=tk.StringVar(value=available[0] if available else "Custom Category")
        self.rule_preset_menu=ctk.CTkOptionMenu(add_row,width=220,height=34,values=available or ["Custom Category"],variable=self.rule_preset_var,fg_color=(_ROW_LIGHT,"#1E293B"),button_color=("#CBD5E1","#334155"),button_hover_color=("#94A3B8","#475569"),dropdown_fg_color=(_CARD_LIGHT,_CARD_DARK),dropdown_hover_color=("#E2E8F0","#172033"),text_color=("#1E293B","#E2E8F0"),corner_radius=10)
        self.rule_preset_menu.pack(side="left",padx=(0,8))
        ctk.CTkButton(add_row,text=tx(self.config_data,"add_category"),width=140,height=34,**_BTN_SEC,command=self._add_selected_rule_preset).pack(side="left")
        card=self._card(f,fill="both",expand=True)
        self.rules_rows_frame=ctk.CTkFrame(card,fg_color="transparent"); self.rules_rows_frame.pack(fill="x",padx=14,pady=(0,8))
        self.rule_vars=[]
        rules=self._friendly_rules_for_display(self.config_data.get("rules",ELEMENTAL_CATEGORIES))
        for cat,exts in rules.items(): self._add_rule_row(cat,exts,refresh=False)
        actions=ctk.CTkFrame(card,fg_color="transparent"); actions.pack(fill="x",padx=14,pady=(4,14))
        ctk.CTkButton(actions,text=tx(self.config_data,"save_rules"),width=140,**_BTN_PRI,command=self._save_rules).pack(side="left",padx=(0,8))
        ctk.CTkButton(actions,text=tx(self.config_data,"reset_defaults"),width=150,**_BTN_SEC,command=self._reset_rules).pack(side="left")
    def _friendly_rules_for_display(self,rules):
        rules=dict(rules or {})
        display={cat:list(rules.get(cat,exts)) for cat,exts in ELEMENTAL_CATEGORIES.items()}
        optional_defaults={tuple(v):k for k,v in OPTIONAL_CATEGORY_PRESETS.items() if k!="Custom Category"}
        for cat,exts in rules.items():
            if cat in display: continue
            if cat in LEGACY_SPLIT_CATEGORIES: continue
            if cat in DEFAULT_CATEGORIES:
                preset_name=optional_defaults.get(tuple(exts),cat)
                display[preset_name]=list(exts)
            else:
                display[cat]=list(exts)
        return display
    def _add_selected_rule_preset(self):
        choice=getattr(self,"rule_preset_var",tk.StringVar(value="Custom Category")).get()
        if any(item["category"].get().strip()==choice for item in getattr(self,"rule_vars",[])):
            self._set_status_message(f"{choice} is already listed.","warning"); return
        self._add_rule_row(choice,OPTIONAL_CATEGORY_PRESETS.get(choice,[]))
        self._set_status_message(f"{choice} added. Save when you're happy with it.","success")
    def _add_rule_row(self,cat="",exts=None,refresh=True):
        if refresh and not hasattr(self,"rules_rows_frame"):
            self._show_view("rules"); return
        row=self._row_frame(self.rules_rows_frame,fill="x",pady=1)
        row.configure(height=100); row.pack_propagate(False)
        cat_var=tk.StringVar(value=cat)
        ext_text=", ".join(exts or [])
        ext_var=tk.StringVar(value=ext_text)
        dest_var=tk.StringVar(value=self.config_data.setdefault("category_destinations",{}).get(cat,""))
        left=ctk.CTkFrame(row,fg_color="transparent",width=190); left.pack(side="left",fill="y",padx=(10,8),pady=5); left.pack_propagate(False)
        ctk.CTkEntry(left,textvariable=cat_var,height=32).pack(fill="x")
        ctk.CTkLabel(left,text=tx(self.config_data,"shown_in_queue"),font=ctk.CTkFont(size=10),text_color=("#64748B","#94A3B8")).pack(anchor="w",pady=(1,0))
        middle=ctk.CTkFrame(row,fg_color="transparent"); middle.pack(side="left",fill="x",expand=True,padx=8,pady=5)
        chips=ctk.CTkFrame(middle,fg_color="transparent"); chips.pack(fill="x",pady=(0,3))
        for ext in (exts or [])[:8]:
            ctk.CTkLabel(chips,text=ext,fg_color=("#EAF1FB","#172033"),text_color=("#1E293B","#E2E8F0"),corner_radius=8,padx=8,pady=3,font=ctk.CTkFont(size=11,weight="bold")).pack(side="left",padx=(0,5),pady=1)
        if len(exts or [])>8:
            ctk.CTkLabel(chips,text=f"+{len(exts)-8}",text_color=("#64748B","#94A3B8"),font=ctk.CTkFont(size=11,weight="bold")).pack(side="left",pady=1)
        ctk.CTkEntry(middle,textvariable=ext_var,height=32,placeholder_text=".pdf, .docx, .txt").pack(fill="x")
        right=ctk.CTkFrame(row,fg_color="transparent",width=300); right.pack(side="left",fill="y",padx=(8,10),pady=5); right.pack_propagate(False)
        ctk.CTkEntry(right,textvariable=dest_var,height=30,placeholder_text=tx(self.config_data,"no_folder_selected")).pack(fill="x",pady=(0,4))
        rr=ctk.CTkFrame(right,fg_color="transparent"); rr.pack(fill="x")
        ctk.CTkButton(rr,text=tx(self.config_data,"browse"),width=90,height=28,**_BTN_ALT,command=lambda v=dest_var,c=cat_var:self._browse_rule_dest(v,c.get())).pack(side="left",padx=(0,6))
        ctk.CTkButton(rr,text=tx(self.config_data,"remove"),width=92,height=28,**_BTN_DNG,command=lambda r=row:self._remove_rule_row(r)).pack(side="left")
        self.rule_vars.append({"row":row,"category":cat_var,"extensions":ext_var,"destination":dest_var})
    def _browse_rule_dest(self,var,cat):
        d=filedialog.askdirectory(title=f"Choose folder for {cat or 'category'}")
        if d: var.set(d)
    def _remove_rule_row(self,row):
        self.rule_vars=[x for x in getattr(self,"rule_vars",[]) if x.get("row") is not row]
        row.destroy()
    def _parse_extensions(self,text):
        exts=[]
        for part in text.replace(";",",").split(","):
            ext=part.strip().lower()
            if not ext: continue
            if not ext.startswith("."): ext="."+ext
            if ext not in exts: exts.append(ext)
        return exts
    def _save_rules(self):
        rules={}; dests=self.config_data.setdefault("category_destinations",{})
        for item in getattr(self,"rule_vars",[]):
            cat=item["category"].get().strip()
            if not cat: self._set_status_message("Every rule needs a category name.","warning"); return
            if cat in rules: self._set_status_message(f"Duplicate category: {cat}","warning"); return
            rules[cat]=self._parse_extensions(item["extensions"].get())
            dests[cat]=item.get("destination",tk.StringVar(value=dests.get(cat,""))).get().strip()
        if "Other" not in rules: rules["Other"]=[]
        self.config_data["rules"]=rules
        self.config_data["category_destinations"]={cat:dests.get(cat,"") for cat in rules}
        save_config(self.config_data); self._set_status_message(f"{len(rules)} rule categories saved.","success")
    def _reset_rules(self):
        if not messagebox.askyesno("Scanderbeg","Reset category rules to defaults?"):
            return
        old=self.config_data.setdefault("category_destinations",{})
        self.config_data["rules"]={cat:list(exts) for cat,exts in ELEMENTAL_CATEGORIES.items()}
        self.config_data["category_destinations"]={cat:old.get(cat,"") for cat in ELEMENTAL_CATEGORIES}
        save_config(self.config_data); self._set_status_message("Rules reset to defaults.","success"); self._show_view("rules")
    def _build_settings(self,parent):
        f=ctk.CTkScrollableFrame(parent,fg_color=(_PAGE_LIGHT,_PAGE_DARK)); f.pack(fill="both",expand=True,padx=24,pady=22); ctk.CTkLabel(f,text=tx(self.config_data,"settings"),font=ctk.CTkFont(size=24,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w")
        lang_card=self._card(f,fill="x",pady=(16,8)); ctk.CTkLabel(lang_card,text=tx(self.config_data,"language"),font=ctk.CTkFont(size=15,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w",padx=16,pady=(14,4)); lr=ctk.CTkFrame(lang_card,fg_color="transparent"); lr.pack(fill="x",padx=16,pady=14); self.language_var=tk.StringVar(value=LANGUAGES.get(self.config_data.get("ui_language","en"),"English")); ctk.CTkOptionMenu(lr,width=180,height=34,values=list(LANGUAGES.values()),variable=self.language_var,command=self._save_language,fg_color=(_ROW_LIGHT,"#1E293B"),button_color=("#CBD5E1","#334155"),button_hover_color=("#94A3B8","#475569"),dropdown_fg_color=(_CARD_LIGHT,_CARD_DARK),dropdown_hover_color=("#E2E8F0","#172033"),text_color=("#1E293B","#E2E8F0"),corner_radius=10).pack(side="left")
        c1=self._card(f,fill="x",pady=(16,8)); ctk.CTkLabel(c1,text=tx(self.config_data,"auto_sort_delay"),font=ctk.CTkFont(size=15,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w",padx=16,pady=(14,4)); r=ctk.CTkFrame(c1,fg_color="transparent"); r.pack(fill="x",padx=16,pady=14); self.delay_var=tk.IntVar(value=self.config_data.get("delay_minutes",30)); ctk.CTkLabel(r,text=tx(self.config_data,"minutes"),text_color=("#1E293B","#E2E8F0")).pack(side="left"); ctk.CTkEntry(r,textvariable=self.delay_var,width=90,height=30).pack(side="left",padx=8); ctk.CTkButton(r,text=tx(self.config_data,"save"),width=90,height=30,**_BTN_PRI,command=self._save_delay).pack(side="left")
        smart=self._card(f,fill="x",pady=8); ctk.CTkLabel(smart,text="🧠 "+tx(self.config_data,"smart_scan"),font=ctk.CTkFont(size=15,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w",padx=16,pady=(14,8)); ctk.CTkLabel(smart,text=tx(self.config_data,"smart_scan_desc"),font=ctk.CTkFont(size=11),text_color=("#475569","#94A3B8")).pack(anchor="w",padx=16); sr=ctk.CTkFrame(smart,fg_color="transparent"); sr.pack(fill="x",padx=16,pady=14); self.smart_scan_var=tk.BooleanVar(value=bool(self.config_data.get("smart_scan_enabled",False))); ctk.CTkSwitch(sr,text=tx(self.config_data,"enable_smart"),variable=self.smart_scan_var,command=self._save_smart).pack(side="left",padx=(0,18)); self.smart_score_var=tk.IntVar(value=int(self.config_data.get("smart_scan_min_score",2))); self.smart_max_mb_var=tk.IntVar(value=int(self.config_data.get("smart_scan_max_mb",15))); ctk.CTkLabel(sr,text=tx(self.config_data,"min_score")).pack(side="left"); ctk.CTkEntry(sr,textvariable=self.smart_score_var,width=60,height=30).pack(side="left",padx=8); ctk.CTkLabel(sr,text=tx(self.config_data,"max_mb")).pack(side="left"); ctk.CTkEntry(sr,textvariable=self.smart_max_mb_var,width=70,height=30).pack(side="left",padx=8); ctk.CTkButton(sr,text=tx(self.config_data,"save"),width=90,height=30,**_BTN_PRI,command=self._save_smart).pack(side="left")
        dest=self._card(f,fill="x",pady=8); ctk.CTkLabel(dest,text=tx(self.config_data,"destination_folders"),font=ctk.CTkFont(size=15,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w",padx=16,pady=(14,4)); ctk.CTkLabel(dest,text=tx(self.config_data,"destination_help"),font=ctk.CTkFont(size=11),text_color=("#475569","#94A3B8")).pack(anchor="w",padx=16,pady=(0,10)); self.dest_vars={}
        dest_rules=self._friendly_rules_for_display(self.config_data.get("rules",ELEMENTAL_CATEGORIES))
        for cat,exts in dest_rules.items(): self._add_destination_row(dest,cat,exts)
        ctk.CTkButton(dest,text=tx(self.config_data,"save_destinations"),width=180,height=32,**_BTN_PRI,command=self._save_dests).pack(anchor="w",padx=16,pady=14)
        help_card=self._card(f,fill="x",pady=8); ctk.CTkLabel(help_card,text=tx(self.config_data,"manual_support"),font=ctk.CTkFont(size=15,weight="bold"),text_color=("#0F172A","#F8FAFC")).pack(anchor="w",padx=16,pady=(14,4)); ctk.CTkLabel(help_card,text=f"Support: {SUPPORT_EMAIL}",font=ctk.CTkFont(size=12),text_color=("#475569","#94A3B8")).pack(anchor="w",padx=16,pady=(0,10)); hr=ctk.CTkFrame(help_card,fg_color="transparent"); hr.pack(fill="x",padx=16,pady=(0,14)); ctk.CTkButton(hr,text=tx(self.config_data,"open_manual"),width=160,height=32,**_BTN_PRI,command=open_manual).pack(side="left",padx=(0,8)); ctk.CTkButton(hr,text=tx(self.config_data,"email_support"),width=130,height=32,**_BTN_SEC,command=lambda:os.startfile(f"mailto:{SUPPORT_EMAIL}?subject=Scanderbeg%20Support")).pack(side="left",padx=8)
        ctk.CTkButton(f,text=tx(self.config_data,"clear_ignored_files"),width=190,height=32,**_BTN_SEC,command=self._clear_ignored_files).pack(anchor="w",pady=10); ctk.CTkLabel(f,text=tx(self.config_data,"config",path=CONFIG_FILE),font=ctk.CTkFont(size=10),text_color=("#94A3B8","#64748B")).pack(anchor="w",pady=10)
    def _save_delay(self):
        try: m=max(1,int(self.delay_var.get()))
        except Exception: messagebox.showerror("Scanderbeg","Delay must be a number."); return
        self.config_data["delay_minutes"]=m; save_config(self.config_data); self._set_status_message(f"Delay set to {m} minute(s).","success")
    def _save_language(self,choice=None):
        selected=choice or getattr(self,"language_var",tk.StringVar(value="English")).get()
        code=next((k for k,v in LANGUAGES.items() if v==selected),"en")
        self.config_data["ui_language"]=code; save_config(self.config_data)
        self._refresh_nav_labels()
        self._show_view(self.current_view or "dashboard")
        self._set_status_message(f"{tx(self.config_data,'language')}: {LANGUAGES[code]}","success")
    def _save_smart(self):
        try: self.config_data["smart_scan_min_score"]=max(1,int(self.smart_score_var.get())); self.config_data["smart_scan_max_mb"]=max(1,int(self.smart_max_mb_var.get()))
        except Exception: self._set_status_message("Smart Scan values must be numbers.","warning"); return
        self.config_data["smart_scan_enabled"]=bool(self.smart_scan_var.get()); save_config(self.config_data); self._set_status_message("Smart Scan saved.","success")
    def _add_destination_row(self,parent,cat,exts):
        row=self._row_frame(parent,fill="x",padx=16,pady=1)
        row.configure(height=72); row.pack_propagate(False)
        left=ctk.CTkFrame(row,fg_color="transparent",width=220); left.pack(side="left",fill="y",padx=(10,8),pady=5); left.pack_propagate(False)
        ctk.CTkLabel(left,text=cat,font=ctk.CTkFont(size=13,weight="bold"),anchor="w",text_color=("#0F172A","#F8FAFC")).pack(anchor="w")
        chips=ctk.CTkFrame(left,fg_color="transparent"); chips.pack(fill="x",pady=(2,0))
        shown=list(exts or [])[:3]
        if shown:
            for ext in shown:
                ctk.CTkLabel(chips,text=ext,fg_color=("#EAF1FB","#172033"),text_color=("#1E293B","#E2E8F0"),corner_radius=8,padx=7,pady=2,font=ctk.CTkFont(size=10,weight="bold")).pack(side="left",padx=(0,4),pady=1)
            if len(exts or [])>3:
                ctk.CTkLabel(chips,text=f"+{len(exts)-3}",text_color=("#64748B","#94A3B8"),font=ctk.CTkFont(size=10,weight="bold")).pack(side="left",pady=1)
        else:
            ctk.CTkLabel(chips,text=tx(self.config_data,"fallback"),text_color=("#64748B","#94A3B8"),font=ctk.CTkFont(size=10)).pack(side="left",pady=1)
        var=tk.StringVar(value=self.config_data.setdefault("category_destinations",{}).get(cat,"")); self.dest_vars[cat]=var
        ctk.CTkEntry(row,textvariable=var,height=32,placeholder_text=tx(self.config_data,"no_folder_selected")).pack(side="left",fill="x",expand=True,padx=4,pady=5)
        ctk.CTkButton(row,text=tx(self.config_data,"browse"),width=90,height=32,**_BTN_ALT,command=lambda c=cat:self._browse_dest(c)).pack(side="left",padx=(6,10),pady=5)
    def _browse_dest(self,cat):
        d=filedialog.askdirectory(title=f"Choose folder for {cat}")
        if d: self.dest_vars[cat].set(d)
    def _save_dests(self):
        for c,v in self.dest_vars.items(): self.config_data.setdefault("category_destinations",{})[c]=v.get().strip()
        save_config(self.config_data); self._set_status_message("Destinations saved.","success")
    def _clear_ignored_files(self):
        n=len(self.config_data.get("ignored_files",[]))
        if n==0: messagebox.showinfo("Scanderbeg","No ignored files to clear."); return
        if messagebox.askyesno("Scanderbeg",f"Clear {n} ignored file(s)?\n\nThey will appear in File Queue again if they still exist."): self.config_data["ignored_files"]=[]; save_config(self.config_data); self._set_status_message("Ignored files cleared.","success")
    def _export_log(self):
        log=self.config_data.get("log",[])
        if not log: messagebox.showinfo("Scanderbeg","No log to export."); return
        fp=filedialog.asksaveasfilename(title="Export Scanderbeg log",defaultextension=".txt",filetypes=[("Text files","*.txt"),("All files","*.*")])
        if fp:
            with open(fp,"w",encoding="utf-8") as f: f.write("\n".join(log))
            messagebox.showinfo("Scanderbeg","Log exported.")
    def on_close(self):
        if self.engine: self.engine.stop()
        if self.observer: self.observer.stop(); self.observer.join()
        save_config(self.config_data); self.destroy()

def main():
    app=ScanderbegApp(); app.protocol("WM_DELETE_WINDOW",app.on_close); app.mainloop()
if __name__=="__main__": main()
