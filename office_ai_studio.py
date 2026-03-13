"""
Office AI Studio
────────────────
Local AI-powered office automation tool built on Ollama.
No cloud. No subscription. Your files stay on your machine.

New in this version:
  • Smart Drop Zones  - drag files, pick an action, run instantly
  • Script Runner     - write/run/save Python automation scripts
  • Auto Tasks        - recurring scheduled jobs (rename, convert, backup…)
  • Data Tools        - CSV clean, merge, preview, export
  • Template Engine   - fill DOCX/TXT templates from CSV or manual input
  • Meeting Notes     - paste notes → tasks + summary + email draft
  • Quick Actions     - one-click office workflows
  • AI Pipeline       - multi-step AI processing chains
  • Notepad AI        - editor with AI assistant
  • Terminal AI       - shell with AI explanations
  • History           - all outputs saved, searchable

Requirements:
    pip install requests tkinterdnd2
Ollama:
    https://ollama.com → ollama serve
Optional (more file formats):
    pip install python-docx openpyxl pandas chardet
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading, os, subprocess, json, requests, time, re, csv, io
import shutil, platform, uuid, hashlib, glob, fnmatch, textwrap
from pathlib import Path
from datetime import datetime, timedelta
from collections import defaultdict

# optional imports
try: import docx;       HAS_DOCX  = True
except ImportError:     HAS_DOCX  = False
try: import openpyxl;   HAS_XLSX  = True
except ImportError:     HAS_XLSX  = False
try: import pandas;     HAS_PD    = True
except ImportError:     HAS_PD    = False
try: import chardet;    HAS_CD    = True
except ImportError:     HAS_CD    = False
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False
    TkinterDnD = None

# ══════════════════════════════════════════════════════════════════════
#  THEME
# ══════════════════════════════════════════════════════════════════════
BG0 = "#0e0e11"; BG1 = "#141418"; BG2 = "#1c1c22"; BG3 = "#24242c"
BG4 = "#2e2e38"; BOR = "#38384a"; BOR2 = "#50505f"
FG0 = "#f5f5fa"; FG1 = "#d0d0e0"; FG2 = "#9090a8"; FG3 = "#6a6a82"
ACC = "#7c6af7"; ACC2 = "#a594f9"; ACCD = "#2d2640"
GRN = "#3dd68c"; GRND = "#1a3329"
RED = "#f25f5c"; REDD = "#2e1818"
ORG = "#f4a24a"; ORGD = "#2e2010"
CYN = "#38bdf8"; CYND = "#0f2535"
PNK = "#f472b6"; PNKD = "#2e1525"
YEL = "#fbbf24"; YELD = "#2c2008"

F_UI   = ("Segoe UI", 11)
F_UIS  = ("Segoe UI", 10)
F_UIT  = ("Segoe UI",  9)
F_UIB  = ("Segoe UI", 11, "bold")
F_H1   = ("Segoe UI", 18, "bold")
F_H2   = ("Segoe UI", 13, "bold")
F_H3   = ("Segoe UI", 11, "bold")
F_MONO = ("Consolas", 10)
F_MONOB= ("Consolas", 10, "bold")
F_MONOS= ("Consolas",  9)

OLLAMA       = "http://localhost:11434"
MODEL_DEFAULT= "llama3.2:3b"
DATA_DIR     = Path.home() / ".office_ai_studio"
DATA_DIR.mkdir(exist_ok=True)
PIPES_FILE   = DATA_DIR / "pipelines.json"
HISTORY_FILE = DATA_DIR / "history.json"
SCRIPTS_FILE = DATA_DIR / "scripts.json"
TASKS_FILE   = DATA_DIR / "tasks.json"
TEMPLATES_FILE = DATA_DIR / "templates.json"

TEXT_EXT = {".txt",".py",".js",".ts",".json",".xml",".html",".css",
            ".md",".yaml",".yml",".csv",".log",".ini",".cfg",".toml",
            ".bat",".ps1",".sh",".sql",".r",".c",".cpp",".h",".env",
            ".gitignore",".rs",".go",".rb",".php",".java",".cs",".tex"}

# ══════════════════════════════════════════════════════════════════════
#  UI HELPERS
# ══════════════════════════════════════════════════════════════════════
def ttk_style():
    s = ttk.Style()
    s.theme_use("clam")
    s.configure("Treeview", background=BG2, foreground=FG1,
                fieldbackground=BG2, font=F_UIS, rowheight=26, borderwidth=0)
    s.configure("Treeview.Heading", background=BG3, foreground=FG2,
                font=F_UIT, borderwidth=0)
    s.map("Treeview", background=[("selected", ACCD)],
          foreground=[("selected", ACC2)])
    s.configure("Vertical.TScrollbar", background=BG3, troughcolor=BG1,
                borderwidth=0, relief="flat", width=4, arrowsize=0)
    s.map("Vertical.TScrollbar", background=[("active", BG4)])
    s.configure("TCombobox", fieldbackground=BG3, background=BG3,
                foreground=FG0, arrowcolor=FG2, borderwidth=0)
    s.map("TCombobox", fieldbackground=[("readonly", BG3)],
          foreground=[("readonly", FG0)])
    s.configure("TNotebook", background=BG1, borderwidth=0)
    s.configure("TNotebook.Tab", background=BG2, foreground=FG2,
                font=F_UIS, padding=(14, 6), borderwidth=0)
    s.map("TNotebook.Tab", background=[("selected", BG1)],
          foreground=[("selected", FG0)])

def Btn(parent, text, cmd=None, style="default", **kw):
    S = {"default":(BG3,FG1,BG4,FG0), "primary":(ACC,FG0,ACC2,FG0),
         "danger":(REDD,RED,RED,BG0),  "ghost":(BG1,FG2,BG2,FG1),
         "success":(GRND,GRN,GRN,BG0), "warn":(ORGD,ORG,ORG,BG0),
         "subtle":(BG2,FG2,BG3,FG1)}
    bg,fg,abg,afg = S.get(style, S["default"])
    px = kw.pop("px", 12); py = kw.pop("py", 6)
    b = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                  activebackground=abg, activeforeground=afg,
                  font=kw.pop("font", F_UIS), bd=0, relief="flat",
                  cursor="hand2", padx=px, pady=py, **kw)
    b.bind("<Enter>", lambda e: b.config(bg=abg, fg=afg))
    b.bind("<Leave>", lambda e: b.config(bg=bg,  fg=fg))
    return b

def Lbl(parent, text="", font=F_UI, fg=FG1, bg=BG1, **kw):
    return tk.Label(parent, text=text, font=font, fg=fg, bg=bg, **kw)

def Div(parent, vertical=False, color=BOR, **kw):
    if vertical:
        tk.Frame(parent, bg=color, width=1).pack(side="left", fill="y", **kw)
    else:
        tk.Frame(parent, bg=color, height=1).pack(fill="x", **kw)

def scroll_wrap(parent, bg=BG1):
    outer = tk.Frame(parent, bg=bg)
    cv    = tk.Canvas(outer, bg=bg, highlightthickness=0)
    sb    = ttk.Scrollbar(outer, orient="vertical", command=cv.yview)
    cv.configure(yscrollcommand=sb.set)
    inner = tk.Frame(cv, bg=bg)
    win   = cv.create_window((0,0), window=inner, anchor="nw")
    def _cfg_i(e, c=cv): c.configure(scrollregion=c.bbox("all"))
    def _cfg_c(e, c=cv, w=win): c.itemconfig(w, width=e.width)
    def _scr(e, c=cv): c.yview_scroll(int(-1*(e.delta/120)), "units")
    inner.bind("<Configure>", _cfg_i)
    cv.bind("<Configure>", _cfg_c)
    cv.bind("<MouseWheel>", _scr)
    cv.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    return outer, cv, inner

def Card(parent, bg=BG2, pad=12, **kw):
    f = tk.Frame(parent, bg=bg, **kw)
    tk.Frame(f, bg=bg, height=pad).pack()
    return f

def SectionHdr(parent, title, subtitle="", bg=BG1, btn_text=None, btn_cmd=None):
    f = tk.Frame(parent, bg=bg)
    f.pack(fill="x", padx=18, pady=(16,6))
    Lbl(f, title, F_H2, FG0, bg).pack(side="left")
    if subtitle:
        Lbl(f, subtitle, F_UIT, FG3, bg).pack(side="left", padx=10, pady=2)
    if btn_text:
        Btn(f, btn_text, btn_cmd, "primary", font=F_UIT, px=12, py=5
            ).pack(side="right")
    return f

def fsize(n):
    for u in ["B","KB","MB","GB"]:
        if n < 1024: return f"{n:.0f} {u}"
        n /= 1024
    return f"{n:.1f} GB"

def ficon(ext):
    return {"py":"PY","js":"JS","json":"{}","html":"<>","css":"CS",
            "md":"MD","txt":"TX","pdf":"PDF","docx":"DOC","xlsx":"XLS",
            "csv":"CSV","sql":"SQL","sh":"SH","bat":"BAT","log":"LOG",
            "png":"IMG","jpg":"IMG","jpeg":"IMG","zip":"ZIP","py":"PY",
            }.get(ext.lower().lstrip("."), "  ")

def detect_enc(path):
    if HAS_CD:
        raw = open(path,"rb").read(32768)
        r   = chardet.detect(raw)
        return r.get("encoding") or "utf-8"
    return "utf-8"

def read_text(path, max_bytes=60000):
    enc = detect_enc(path)
    try:
        return open(path, encoding=enc, errors="replace").read(max_bytes)
    except:
        return open(path, encoding="utf-8", errors="replace").read(max_bytes)

# ══════════════════════════════════════════════════════════════════════
#  DATA MODELS
# ══════════════════════════════════════════════════════════════════════
class PipeStep:
    def __init__(self, name, color, dim_bg, instruction):
        self.uid = uuid.uuid4().hex[:6]
        self.name = name; self.color = color; self.dim_bg = dim_bg
        self.instruction = instruction
        self.source = "prev"; self.files = []; self.text = ""
        self.output = ""; self.status = "idle"

class Script:
    def __init__(self, name="", code="", desc=""):
        self.uid = uuid.uuid4().hex[:8]
        self.name = name; self.code = code; self.desc = desc
        self.created = datetime.now().isoformat()
        self.last_run = ""; self.last_result = ""

class AutoTask:
    def __init__(self, name="", trigger="drop", pattern="*",
                 action="", target_dir="", enabled=True):
        self.uid = uuid.uuid4().hex[:8]
        self.name = name; self.trigger = trigger; self.pattern = pattern
        self.action = action; self.target_dir = target_dir
        self.enabled = enabled
        self.runs = 0; self.last_run = ""

# ══════════════════════════════════════════════════════════════════════
#  PERSISTENCE
# ══════════════════════════════════════════════════════════════════════
def jload(p, default=None):
    if default is None: default = []
    try: return json.loads(p.read_text("utf-8")) if p.exists() else default
    except: return default

def jsave(p, data):
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), "utf-8")

def pipes_load():
    raw = jload(PIPES_FILE)
    out = []
    for p in raw:
        steps = []
        for s in p.get("steps",[]):
            st = PipeStep(s["name"],s["color"],s.get("dim_bg",BG3),s["instruction"])
            st.source=s.get("source","prev"); st.files=s.get("files",[])
            st.text=s.get("text",""); steps.append(st)
        out.append({"name":p["name"],"steps":steps})
    return out

def pipes_save(pipes):
    jsave(PIPES_FILE, [{"name":p["name"],"steps":[{
        "name":s.name,"color":s.color,"dim_bg":s.dim_bg,
        "instruction":s.instruction,"source":s.source,
        "files":s.files,"text":s.text}
        for s in p["steps"]]} for p in pipes])

def scripts_load():
    raw = jload(SCRIPTS_FILE)
    out = []
    for r in raw:
        s = Script(r.get("name",""), r.get("code",""), r.get("desc",""))
        s.uid=r.get("uid",s.uid); s.last_run=r.get("last_run","")
        s.last_result=r.get("last_result",""); out.append(s)
    return out

def scripts_save(scripts):
    jsave(SCRIPTS_FILE,[{"uid":s.uid,"name":s.name,"code":s.code,
        "desc":s.desc,"last_run":s.last_run,"last_result":s.last_result}
        for s in scripts])

def tasks_load():
    raw = jload(TASKS_FILE)
    out = []
    for r in raw:
        t = AutoTask(r.get("name",""),r.get("trigger","drop"),
                     r.get("pattern","*"),r.get("action",""),
                     r.get("target_dir",""),r.get("enabled",True))
        t.uid=r.get("uid",t.uid); t.runs=r.get("runs",0)
        t.last_run=r.get("last_run",""); out.append(t)
    return out

def tasks_save(tasks):
    jsave(TASKS_FILE,[{"uid":t.uid,"name":t.name,"trigger":t.trigger,
        "pattern":t.pattern,"action":t.action,"target_dir":t.target_dir,
        "enabled":t.enabled,"runs":t.runs,"last_run":t.last_run}
        for t in tasks])

def history_load(): return jload(HISTORY_FILE)
def history_save(items): jsave(HISTORY_FILE, items[-500:])

# ══════════════════════════════════════════════════════════════════════
#  OLLAMA
# ══════════════════════════════════════════════════════════════════════
def ollama_stream(model, messages, on_token, is_stopped, timeout=180):
    full = ""
    def _chat():
        nonlocal full
        r = requests.post(f"{OLLAMA}/api/chat",
            json={"model":model,"messages":messages,"stream":True},
            stream=True, timeout=timeout)
        r.raise_for_status()
        for raw in r.iter_lines():
            if is_stopped(): break
            if not raw: continue
            try: data = json.loads(raw)
            except: continue
            if "error" in data: raise RuntimeError(data["error"])
            t = (data.get("message") or {}).get("content") or ""
            if t: full += t; on_token(t)
        return full
    def _generate():
        nonlocal full; full = ""
        prompt = "\n\n".join(
            f"{'User' if m['role']=='user' else 'Assistant'}: {m['content']}"
            for m in messages) + "\n\nAssistant:"
        r = requests.post(f"{OLLAMA}/api/generate",
            json={"model":model,"prompt":prompt,"stream":True},
            stream=True, timeout=timeout)
        r.raise_for_status()
        for raw in r.iter_lines():
            if is_stopped(): break
            if not raw: continue
            try: data = json.loads(raw)
            except: continue
            if "error" in data: raise RuntimeError(data["error"])
            t = data.get("response") or ""
            if t: full += t; on_token(t)
        return full
    try: return _chat()
    except RuntimeError as e:
        if any(k in str(e).lower() for k in ("template","chat","does not support")):
            return _generate()
        raise

# ══════════════════════════════════════════════════════════════════════
#  FILE AUTOMATION ACTIONS
# ══════════════════════════════════════════════════════════════════════
BUILTIN_ACTIONS = {
    "rename_date": {
        "label": "Prefix with today's date",
        "desc":  "Rename: YYYY-MM-DD_filename.ext",
        "fn": lambda p: p.rename(p.parent / f"{datetime.now().strftime('%Y-%m-%d')}_{p.name}"),
    },
    "rename_lower": {
        "label": "Lowercase filename",
        "fn": lambda p: p.rename(p.parent / p.name.lower()),
    },
    "rename_spaces": {
        "label": "Replace spaces with underscores",
        "fn": lambda p: p.rename(p.parent / p.name.replace(" ","_")),
    },
    "copy_desktop": {
        "label": "Copy to Desktop",
        "fn": lambda p: shutil.copy2(str(p), str(Path.home()/"Desktop"/p.name)),
    },
    "move_docs": {
        "label": "Move to Documents",
        "fn": lambda p: shutil.move(str(p), str(Path.home()/"Documents"/p.name)),
    },
    "count_lines": {
        "label": "Count lines (text files)",
        "fn": lambda p: f"{sum(1 for _ in open(p, encoding='utf-8', errors='replace'))} lines in {p.name}",
    },
    "word_count": {
        "label": "Word count",
        "fn": lambda p: f"{len(open(p,encoding='utf-8',errors='replace').read().split())} words in {p.name}",
    },
    "csv_preview": {
        "label": "CSV: preview first 5 rows",
        "fn": lambda p: "\n".join([",".join(r) for r in list(csv.reader(open(p,encoding=detect_enc(p),errors="replace")))[:6]]),
    },
    "hash_md5": {
        "label": "Calculate MD5 hash",
        "fn": lambda p: f"MD5({p.name}) = {hashlib.md5(p.read_bytes()).hexdigest()}",
    },
    "duplicate": {
        "label": "Duplicate file",
        "fn": lambda p: shutil.copy2(str(p), str(p.parent/f"{p.stem}_copy{p.suffix}")),
    },
    "open_folder": {
        "label": "Open containing folder",
        "fn": lambda p: (os.startfile(str(p.parent)) if platform.system()=="Windows"
                         else subprocess.Popen(["xdg-open", str(p.parent)])),
    },
}

def run_action(action_key, files, log_fn=None):
    action = BUILTIN_ACTIONS.get(action_key)
    if not action: return f"Unknown action: {action_key}"
    results = []
    for p in files:
        p = Path(p)
        try:
            r = action["fn"](p)
            msg = str(r) if r else f"✓  {action['label']}  →  {p.name}"
            results.append(msg)
            if log_fn: log_fn(msg + "\n")
        except Exception as ex:
            msg = f"✗  {p.name}  →  {ex}"
            results.append(msg)
            if log_fn: log_fn(msg + "\n")
    return "\n".join(results)

# ══════════════════════════════════════════════════════════════════════
#  PIPELINE TEMPLATES
# ══════════════════════════════════════════════════════════════════════
STEP_TEMPLATES = [
    ("Summarize",    ACC,  ACCD, "Summarize the following concisely:\n\n{input}"),
    ("Translate EN", CYN,  CYND, "Translate to English:\n\n{input}"),
    ("Translate PL", CYN,  CYND, "Translate to Polish:\n\n{input}"),
    ("Improve Style",PNK,  PNKD, "Improve style and grammar. Keep the language:\n\n{input}"),
    ("Bullet Points",GRN,  GRND, "Turn into a clear bullet-point list:\n\n{input}"),
    ("Review Code",  ORG,  ORGD, "Review this code - bugs, improvements, best practices:\n\n{input}"),
    ("Explain Code", ORG,  ORGD, "Explain what this code does, step by step:\n\n{input}"),
    ("Extract Data", GRN,  GRND, "Extract all numbers, dates, names and key facts:\n\n{input}"),
    ("Action Items", GRN,  GRND, "Extract a concrete action-item list:\n\n{input}"),
    ("Write Email",  ACC,  ACCD, "Write a professional email based on:\n\n{input}"),
    ("Write Report", PNK,  PNKD, "Write a professional report based on:\n\n{input}"),
    ("Meeting Notes",YEL,  YELD, "From these meeting notes extract:\n1. Key decisions\n2. Action items with owners\n3. Next steps\n\nNotes:\n{input}"),
    ("Q&A List",     ORG,  ORGD, "Create a Q&A list based on:\n\n{input}"),
    ("Classify",     CYN,  CYND, "Classify and group these items:\n\n{input}"),
    ("Custom…",      FG2,  BG3,  "{input}"),
]

# ══════════════════════════════════════════════════════════════════════
#  FILE PICKER
# ══════════════════════════════════════════════════════════════════════
class FilePicker(tk.Toplevel):
    def __init__(self, parent, initial_files=None):
        super().__init__(parent)
        self.title("Select Files")
        self.geometry("940x600")
        self.configure(bg=BG1)
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.cur_path = Path.home()
        self.selected = list(initial_files or [])
        self.result = None
        ttk_style()
        self._build()
        self._refresh()

    def _build(self):
        hdr = tk.Frame(self, bg=BG0, pady=10)
        hdr.pack(fill="x")
        Lbl(hdr, "Select Files", F_H2, FG0, BG0).pack(side="left", padx=16)
        Div(self)
        tb = tk.Frame(self, bg=BG1, pady=6)
        tb.pack(fill="x", padx=10)
        Btn(tb, "↑ Up", self._up, "ghost").pack(side="left", padx=(0,4))
        Btn(tb, "⌂ Home", lambda: self._nav(Path.home()), "ghost").pack(side="left")
        self.pv = tk.StringVar(value=str(self.cur_path))
        e = tk.Entry(tb, textvariable=self.pv, bg=BG3, fg=FG0, font=F_MONO,
                     bd=0, relief="flat", insertbackground=ACC)
        e.pack(side="left", fill="x", expand=True, padx=10, ipady=5, ipadx=8)
        e.bind("<Return>", lambda ev: self._nav(Path(self.pv.get())))
        Div(self)
        body = tk.Frame(self, bg=BG1)
        body.pack(fill="both", expand=True)
        left = tk.Frame(body, bg=BG1)
        left.pack(side="left", fill="both", expand=True)
        ab = tk.Frame(left, bg=BG2, pady=8)
        ab.pack(fill="x")
        self.add_btn = Btn(ab, "  +  Add Selected  ", self._do_add, "primary",
                           font=F_UIB, px=16, py=7)
        self.add_btn.pack(side="left", padx=10)
        self.browser_lbl = Lbl(ab, "", F_UIT, FG2, BG2)
        self.browser_lbl.pack(side="left")
        cols = ("name","size","modified")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", selectmode="extended")
        for c,w,l in [("name",320,"Name"),("size",80,"Size"),("modified",150,"Modified")]:
            self.tree.heading(c, text=l, anchor="w")
            self.tree.column(c, width=w, anchor="w" if c!="size" else "e")
        tsb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        tsb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._dbl)
        self.tree.bind("<<TreeviewSelect>>", self._on_sel)
        Div(body, vertical=True)
        right = tk.Frame(body, bg=BG2, width=260)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)
        rh = tk.Frame(right, bg=BG2, pady=10)
        rh.pack(fill="x", padx=14)
        Lbl(rh, "Selected", F_H3, FG0, BG2).pack(side="left")
        self._cnt = Lbl(rh, "0", F_UIT, FG2, BG2)
        self._cnt.pack(side="right")
        self.sel_lb = tk.Listbox(right, bg=BG3, fg=FG1, font=F_MONOS,
                                  bd=0, relief="flat", selectmode="extended",
                                  activestyle="none", selectbackground=ACCD,
                                  selectforeground=ACC2)
        slsb = ttk.Scrollbar(right, orient="vertical", command=self.sel_lb.yview)
        self.sel_lb.configure(yscrollcommand=slsb.set)
        self.sel_lb.pack(side="left", fill="both", expand=True, padx=(8,0), pady=4)
        slsb.pack(side="right", fill="y", pady=4, padx=(0,4))
        bf = tk.Frame(right, bg=BG2, pady=8)
        bf.pack(fill="x", padx=8)
        Btn(bf, "Remove", self._remove, "danger", font=F_UIT).pack(fill="x", pady=2)
        Btn(bf, "Clear",  self._clear,  "ghost",  font=F_UIT).pack(fill="x", pady=2)
        Div(self)
        ft = tk.Frame(self, bg=BG0, pady=10)
        ft.pack(fill="x", padx=14)
        self.st = Lbl(ft, "", F_UIT, FG2, BG0)
        self.st.pack(side="left")
        Btn(ft, "Cancel", self._cancel, "ghost", px=14, py=7).pack(side="right", padx=(8,0))
        Btn(ft, "✓  Confirm", self._ok, "primary", font=F_UIB, px=18, py=7).pack(side="right")

    def _refresh(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        try:
            items = sorted(self.cur_path.iterdir(),
                           key=lambda x:(not x.is_dir(), x.name.lower()))
        except PermissionError: self.browser_lbl.config(text="Access denied"); return
        n = 0
        for e in items:
            if e.name.startswith("."): continue
            try:
                s = e.stat()
                sz = fsize(s.st_size) if e.is_file() else "-"
                m  = datetime.fromtimestamp(s.st_mtime).strftime("%d %b %Y  %H:%M")
                ic = "[dir]" if e.is_dir() else f"[{ficon(e.suffix)}]"
                self.tree.insert("","end",iid=str(e),values=(f"{ic}  {e.name}",sz,m))
                n += 1
            except: pass
        self.pv.set(str(self.cur_path))
        self.browser_lbl.config(text=f"{n} items")
        self._refresh_sel()

    def _refresh_sel(self):
        self.sel_lb.delete(0,"end")
        for p in self.selected: self.sel_lb.insert("end", f"  {Path(p).name}")
        self._cnt.config(text=str(len(self.selected)))

    def _nav(self, p):
        p = Path(p)
        if p.is_dir(): self.cur_path = p; self._refresh()

    def _up(self):
        p = self.cur_path.parent
        if p != self.cur_path: self._nav(p)

    def _dbl(self, ev):
        row = self.tree.identify_row(ev.y)
        if not row: return
        p = Path(row)
        if p.is_dir(): self._nav(p)
        elif str(p) not in self.selected:
            self.selected.append(str(p)); self._refresh_sel()

    def _on_sel(self, _):
        n = len([s for s in self.tree.selection() if Path(s).is_file()])
        self.add_btn.config(text=f"  +  Add {n} file{'s' if n!=1 else ''}  " if n
                            else "  +  Add Selected  ")

    def _do_add(self):
        added = 0
        for iid in self.tree.selection():
            p = Path(iid)
            if p.is_file() and str(p) not in self.selected:
                self.selected.append(str(p)); added += 1
        if added: self._refresh_sel(); self.st.config(text=f"✓ Added {added}")
        else: self.st.config(text="Select files then click + Add")

    def _remove(self):
        for i in reversed(list(self.sel_lb.curselection())): del self.selected[i]
        self._refresh_sel()

    def _clear(self): self.selected.clear(); self._refresh_sel()
    def _ok(self): self.result = list(self.selected); self.destroy()
    def _cancel(self): self.result = None; self.destroy()

# ══════════════════════════════════════════════════════════════════════
#  SMART DROP ZONE DIALOG
# ══════════════════════════════════════════════════════════════════════
class SmartDropDialog(tk.Toplevel):
    """Shown when files are dropped onto the main window drop zone."""
    def __init__(self, parent, files, on_action, model_var):
        super().__init__(parent)
        self.title("What should I do with these files?")
        self.geometry("680x560")
        self.configure(bg=BG1)
        self.resizable(False, False)
        self.files = files
        self.on_action = on_action
        self.model_var = model_var
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=BG0, pady=12)
        hdr.pack(fill="x")
        Lbl(hdr, "Smart Drop", F_H2, FG0, BG0).pack(side="left", padx=16)
        Lbl(hdr, f"{len(self.files)} file{'s' if len(self.files)!=1 else ''}",
            F_UIS, ACC2, BG0).pack(side="left", padx=8, pady=2)
        Div(self)

        # file list
        fl = tk.Frame(self, bg=BG2)
        fl.pack(fill="x", padx=16, pady=(10,4))
        for p in self.files[:5]:
            Lbl(fl, f"  📄  {Path(p).name}", F_UIT, FG2, BG2, anchor="w"
                ).pack(fill="x", padx=8, pady=1)
        if len(self.files) > 5:
            Lbl(fl, f"  … and {len(self.files)-5} more", F_UIT, FG3, BG2
                ).pack(anchor="w", padx=8, pady=2)
        Div(self, pady=4)

        Lbl(self, "Choose an action:", F_H3, FG0, BG1).pack(anchor="w", padx=16, pady=(8,4))

        # ── AI actions ────────────────────────────────────────────
        Lbl(self, "  AI  ·  powered by Ollama", F_UIT, ACC2, BG1
            ).pack(anchor="w", padx=16, pady=(4,2))
        ai_grid = tk.Frame(self, bg=BG1)
        ai_grid.pack(fill="x", padx=16, pady=4)
        ai_actions = [
            ("⬡  Summarize",        "ai_summarize"),
            ("⬡  Extract key data", "ai_extract"),
            ("⬡  List action items","ai_actions"),
            ("⬡  Translate to EN",  "ai_translate_en"),
            ("⬡  Translate to PL",  "ai_translate_pl"),
            ("⬡  Review / critique","ai_review"),
            ("⬡  Add to pipeline",  "add_pipeline"),
            ("⬡  Open in Chat",     "open_chat"),
        ]
        for i, (label, key) in enumerate(ai_actions):
            Btn(ai_grid, label, lambda k=key: self._pick(k),
                "subtle", font=F_UIT, px=12, py=7
                ).grid(row=i//4, column=i%4, padx=4, pady=3, sticky="ew")
        for c in range(4): ai_grid.columnconfigure(c, weight=1)

        Div(self, pady=6)

        # ── File actions ──────────────────────────────────────────
        Lbl(self, "  File automation  ·  no AI needed", F_UIT, GRN, BG1
            ).pack(anchor="w", padx=16, pady=(4,2))
        fa_grid = tk.Frame(self, bg=BG1)
        fa_grid.pack(fill="x", padx=16, pady=4)
        file_actions = [
            ("📅 Prefix date",      "rename_date"),
            ("🔡 Lowercase name",   "rename_lower"),
            ("_  Fix spaces",       "rename_spaces"),
            ("📋 Copy to Desktop",  "copy_desktop"),
            ("📂 Move to Docs",     "move_docs"),
            ("# Count lines",       "count_lines"),
            ("Σ  Word count",       "word_count"),
            ("🔒 MD5 hash",         "hash_md5"),
            ("⧉  Duplicate",        "duplicate"),
            ("📁 Open folder",      "open_folder"),
        ]
        for i, (label, key) in enumerate(file_actions):
            Btn(fa_grid, label, lambda k=key: self._pick(k),
                "default", font=F_UIT, px=10, py=6
                ).grid(row=i//5, column=i%5, padx=3, pady=3, sticky="ew")
        for c in range(5): fa_grid.columnconfigure(c, weight=1)

        Div(self)
        ft = tk.Frame(self, bg=BG0, pady=8)
        ft.pack(fill="x", padx=14)
        Btn(ft, "Cancel", self.destroy, "ghost", py=7).pack(side="right")

    def _pick(self, key):
        self.destroy()
        self.on_action(key, self.files)

# ══════════════════════════════════════════════════════════════════════
#  NOTEPAD AI
# ══════════════════════════════════════════════════════════════════════
class NotepadAI(tk.Toplevel):
    def __init__(self, parent, model_var):
        super().__init__(parent)
        self.title("Notepad AI")
        self.geometry("1100x700")
        self.configure(bg=BG1)
        self.model_var = model_var
        self.generating = False; self.stop_flag = False
        self._build()

    def _build(self):
        tb = tk.Frame(self, bg=BG0, pady=8)
        tb.pack(fill="x")
        Lbl(tb, "Notepad AI", F_H3, FG0, BG0).pack(side="left", padx=16)
        for label, cmd in [("Open", self._open), ("Save", self._save), ("Clear", self._clear)]:
            Btn(tb, label, cmd, "ghost", font=F_UIT).pack(side="left", padx=3)
        self.wc_lbl = Lbl(tb, "0 words", F_UIT, FG3, BG0)
        self.wc_lbl.pack(side="right", padx=16)
        Div(self)
        split = tk.Frame(self, bg=BG1)
        split.pack(fill="both", expand=True)
        ew = tk.Frame(split, bg=BG1)
        ew.pack(side="left", fill="both", expand=True)
        self.editor = tk.Text(ew, bg=BG1, fg=FG0, font=("Segoe UI",12),
                               bd=0, relief="flat", wrap="word",
                               insertbackground=ACC, padx=28, pady=20,
                               selectbackground=ACCD)
        esb = ttk.Scrollbar(ew, orient="vertical", command=self.editor.yview)
        self.editor.configure(yscrollcommand=esb.set)
        self.editor.pack(side="left", fill="both", expand=True)
        esb.pack(side="right", fill="y")
        self.editor.bind("<KeyRelease>", self._wc)
        Div(split, vertical=True)
        ai = tk.Frame(split, bg=BG2, width=340)
        ai.pack(side="right", fill="y")
        ai.pack_propagate(False)
        Lbl(ai, "AI Assistant", F_H3, FG0, BG2).pack(anchor="w", padx=16, pady=(16,4))
        qf = tk.Frame(ai, bg=BG2)
        qf.pack(fill="x", padx=12, pady=(0,8))
        for label, prompt in [
            ("Summarize",   "Summarize concisely:\n\n{text}"),
            ("Fix Grammar", "Fix grammar and style:\n\n{text}"),
            ("Continue",    "Continue naturally:\n\n{text}"),
            ("Translate EN","Translate to English:\n\n{text}"),
        ]:
            Btn(qf, label, lambda p=prompt: self._quick(p), "default",
                font=F_UIT, px=8, py=4).pack(side="left", padx=2, pady=2)
        Div(ai, pady=4)
        self.ai_inp = tk.Text(ai, bg=BG3, fg=FG0, font=F_UI, bd=0,
                               relief="flat", height=3, insertbackground=ACC,
                               wrap="word", padx=10, pady=8)
        self.ai_inp.pack(fill="x", padx=12)
        self.ai_inp.bind("<Control-Return>", lambda e: self._send())
        bf = tk.Frame(ai, bg=BG2)
        bf.pack(fill="x", padx=12, pady=6)
        self.send_b = Btn(bf, "▶ Send (Ctrl+↵)", self._send, "primary", font=F_UIB)
        self.send_b.pack(side="left")
        Btn(bf, "■ Stop", lambda: setattr(self,"stop_flag",True), "danger"
            ).pack(side="left", padx=6)
        Div(ai, pady=4)
        ow = tk.Frame(ai, bg=BG3)
        ow.pack(fill="both", expand=True, padx=12, pady=(0,8))
        self.ai_out = tk.Text(ow, bg=BG3, fg=FG0, font=F_MONO, bd=0,
                               relief="flat", wrap="word", state="disabled",
                               padx=10, pady=10)
        osb = ttk.Scrollbar(ow, orient="vertical", command=self.ai_out.yview)
        self.ai_out.configure(yscrollcommand=osb.set)
        self.ai_out.pack(side="left", fill="both", expand=True)
        osb.pack(side="right", fill="y")
        af = tk.Frame(ai, bg=BG2)
        af.pack(fill="x", padx=12, pady=(0,12))
        Btn(af, "↓ Insert", self._insert, "default", font=F_UIT, px=10, py=4).pack(side="left")
        Btn(af, "↺ Replace", self._replace, "default", font=F_UIT, px=10, py=4).pack(side="left",padx=4)

    def _wc(self, _=None):
        t = self.editor.get("1.0","end").strip()
        self.wc_lbl.config(text=f"{len(t.split()) if t else 0} words")

    def _quick(self, prompt):
        t = self.editor.get("1.0","end").strip()
        if t: self._run_ai(prompt.replace("{text}", t[:6000]))

    def _send(self):
        cmd = self.ai_inp.get("1.0","end").strip()
        t   = self.editor.get("1.0","end").strip()
        if not cmd: return
        self._run_ai(f"{cmd}\n\nText:\n{t[:6000]}" if t else cmd)

    def _run_ai(self, prompt):
        if self.generating: return
        self.generating = True; self.stop_flag = False
        self.send_b.config(state="disabled")
        self.ai_out.config(state="normal"); self.ai_out.delete("1.0","end")
        self.ai_out.config(state="disabled")
        threading.Thread(target=self._stream, args=(prompt,), daemon=True).start()

    def _stream(self, prompt):
        try:
            ollama_stream(self.model_var.get(),
                [{"role":"user","content":prompt}],
                lambda t: self.after(0, self._tok, t), lambda: self.stop_flag)
        except Exception as ex: self.after(0, self._tok, f"\n⚠ {ex}")
        finally: self.after(0, lambda: self.send_b.config(state="normal")); self.generating=False

    def _tok(self, t):
        self.ai_out.config(state="normal")
        self.ai_out.insert("end", t); self.ai_out.see("end")
        self.ai_out.config(state="disabled")

    def _insert(self):
        t = self.ai_out.get("1.0","end").strip()
        if t: self.editor.insert("end", "\n\n" + t)

    def _replace(self):
        t = self.ai_out.get("1.0","end").strip()
        if not t: return
        sel = self.editor.tag_ranges("sel")
        if sel: self.editor.delete(sel[0],sel[1]); self.editor.insert(sel[0],t)
        else: self.editor.delete("1.0","end"); self.editor.insert("1.0",t)

    def _open(self):
        p = filedialog.askopenfilename(
            filetypes=[("Text","*.txt *.md *.py *.json"),("All","*.*")])
        if p:
            try:
                self.editor.delete("1.0","end")
                self.editor.insert("1.0", read_text(p))
                self._wc()
            except Exception as e: messagebox.showerror("Error", str(e))

    def _save(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text","*.txt"),("Markdown","*.md"),("All","*.*")])
        if p:
            try: open(p,"w",encoding="utf-8").write(self.editor.get("1.0","end"))
            except Exception as e: messagebox.showerror("Error", str(e))

    def _clear(self):
        if messagebox.askyesno("Clear","Clear the editor?"): self.editor.delete("1.0","end")

# ══════════════════════════════════════════════════════════════════════
#  TERMINAL AI
# ══════════════════════════════════════════════════════════════════════
class TerminalAI(tk.Toplevel):
    def __init__(self, parent, model_var):
        super().__init__(parent)
        self.title("Terminal AI")
        self.geometry("900x600")
        self.configure(bg="#0d0d0f")
        self.model_var = model_var
        self.cwd = Path.home(); self.hist = []; self.hist_idx = -1
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg="#0a0a0c", pady=8)
        hdr.pack(fill="x")
        Lbl(hdr,"Terminal AI",F_H3,GRN,"#0a0a0c").pack(side="left",padx=16)
        Lbl(hdr,"type commands  ·  ?? <question> = ask AI  ·  ↑↓ history",
            F_UIT,FG3,"#0a0a0c").pack(side="left",padx=8)
        self.ai_mode = tk.BooleanVar(value=False)
        tk.Checkbutton(hdr, text="AI explains output", variable=self.ai_mode,
                       bg="#0a0a0c", fg=FG2, font=F_UIT, selectcolor="#0a0a0c",
                       activebackground="#0a0a0c", activeforeground=GRN
                       ).pack(side="right", padx=16)
        ow = tk.Frame(self, bg=BG0)
        ow.pack(fill="both", expand=True)
        self.out = tk.Text(ow, bg=BG0, fg="#c8c8d8", font=("Consolas",11),
                            bd=0, relief="flat", state="disabled", wrap="word",
                            padx=14, pady=12, selectbackground="#1e2840",
                            insertbackground=GRN)
        osb = ttk.Scrollbar(ow, orient="vertical", command=self.out.yview)
        self.out.configure(yscrollcommand=osb.set)
        self.out.pack(side="left", fill="both", expand=True)
        osb.pack(side="right", fill="y")
        for tag, fg, fnt in [("prompt",GRN,None),("error",RED,None),
                              ("ai",ACC2,None),("ai_hdr",ACC,("Consolas",11,"bold")),
                              ("info",FG3,None)]:
            self.out.tag_config(tag, foreground=fg, **({"font":fnt} if fnt else {}))
        self._print(f"Terminal AI  ─  {platform.system()} {platform.release()}\n","info")
        self._print(f"  cwd: {self.cwd}\n","info")
        self._print("  ?? <question> to ask AI about anything\n\n","info")
        ir = tk.Frame(self, bg="#0c0c0e", pady=8)
        ir.pack(fill="x")
        self.prompt_lbl = tk.Label(ir, text=f"  {self.cwd.name} ❯ ",
                                    bg="#0c0c0e", fg=GRN, font=("Consolas",11))
        self.prompt_lbl.pack(side="left")
        self.inp = tk.Entry(ir, bg="#0c0c0e", fg="#e0e0f0", font=("Consolas",11),
                             bd=0, relief="flat", insertbackground=GRN)
        self.inp.pack(side="left", fill="x", expand=True, padx=(0,14))
        self.inp.bind("<Return>", self._exec)
        self.inp.bind("<Up>",     self._hist_up)
        self.inp.bind("<Down>",   self._hist_dn)
        self.inp.focus()

    def _print(self, text, tag="output"):
        self.out.config(state="normal")
        self.out.insert("end", text, tag)
        self.out.see("end")
        self.out.config(state="disabled")

    def _exec(self, _=None):
        cmd = self.inp.get().strip()
        if not cmd: return
        self.inp.delete(0,"end")
        self.hist.insert(0, cmd); self.hist_idx = -1
        self._print(f"❯ {cmd}\n","prompt")
        if cmd.startswith("??"): self._ask_ai(cmd[2:].strip()); return
        if cmd.startswith("cd "):
            t = cmd[3:].strip().strip('"')
            try:
                n = (self.cwd / t).resolve()
                if n.is_dir():
                    self.cwd = n
                    self.prompt_lbl.config(text=f"  {self.cwd.name} ❯ ")
                    self._print(f"  → {self.cwd}\n","info")
                else: self._print(f"  Not found: {t}\n","error")
            except Exception as ex: self._print(f"  {ex}\n","error")
            return
        threading.Thread(target=self._run, args=(cmd,), daemon=True).start()

    def _run(self, cmd):
        try:
            enc = "cp1250" if platform.system()=="Windows" else "utf-8"
            p = subprocess.run(cmd, shell=True, capture_output=True, text=True,
                               cwd=str(self.cwd), timeout=30,
                               encoding=enc, errors="replace")
            if p.stdout: self.after(0, self._print, p.stdout)
            if p.stderr: self.after(0, self._print, p.stderr, "error")
            if not p.stdout and not p.stderr: self.after(0, self._print,"  (no output)\n","info")
            if self.ai_mode.get():
                out = (p.stdout or "")+(p.stderr or "")
                if out: self.after(0, self._explain, cmd, out[:2000])
        except subprocess.TimeoutExpired: self.after(0,self._print,"  ⚠ Timeout\n","error")
        except Exception as ex: self.after(0,self._print,f"  ⚠ {ex}\n","error")

    def _ask_ai(self, q):
        self._print("⬡ AI\n","ai_hdr")
        p = (f"You are a terminal expert. Answer concisely. "
             f"Include command examples if helpful. System: {platform.system()}.\n\nQuestion: {q}")
        threading.Thread(target=self._stream_ai, args=(p,), daemon=True).start()

    def _explain(self, cmd, out):
        self._print("\n⬡ AI explains:\n","ai_hdr")
        p = (f"Briefly explain this command output. If errors, explain and suggest fix.\n\n"
             f"Command: {cmd}\nOutput:\n{out}")
        threading.Thread(target=self._stream_ai, args=(p,), daemon=True).start()

    def _stream_ai(self, prompt):
        try:
            ollama_stream(self.model_var.get(),[{"role":"user","content":prompt}],
                lambda t: self.after(0,self._print,t,"ai"), lambda: False, timeout=60)
            self.after(0,self._print,"\n\n")
        except Exception as ex: self.after(0,self._print,f"\n⚠ {ex}\n","error")

    def _hist_up(self,_):
        if self.hist and self.hist_idx < len(self.hist)-1:
            self.hist_idx+=1; self.inp.delete(0,"end"); self.inp.insert(0,self.hist[self.hist_idx])
    def _hist_dn(self,_):
        if self.hist_idx > 0:
            self.hist_idx-=1; self.inp.delete(0,"end"); self.inp.insert(0,self.hist[self.hist_idx])
        elif self.hist_idx==0: self.hist_idx=-1; self.inp.delete(0,"end")

# ══════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ══════════════════════════════════════════════════════════════════════
BaseClass = TkinterDnD.Tk if HAS_DND else tk.Tk

class App(BaseClass):
    def __init__(self):
        super().__init__()
        self.title("Office AI Studio")
        self.geometry("1560x920")
        self.minsize(1100, 700)
        self.configure(bg=BG0)

        self.model_var   = tk.StringVar(value="Loading models…")
        self.pipes       = pipes_load()
        self.pipe_steps  = []
        self.generating  = False
        self.stop_flag   = False
        self.ai_history  = history_load()
        self.scripts     = scripts_load()
        self.auto_tasks  = tasks_load()
        self.cur_path    = Path.home()
        self.sel_path    = ""
        self.sel_content = ""

        ttk_style()
        self._build_ui()
        self._show("drop")
        threading.Thread(target=self._poll_ollama, daemon=True).start()
        self._tick()
        self._setup_main_drop()
        # fetch models immediately on startup, don't wait 12s
        threading.Thread(target=self._fetch_models_once, daemon=True).start()

    # ─────────────────────────────────────────────────────────────────
    #  LAYOUT
    # ─────────────────────────────────────────────────────────────────
    def _build_ui(self):
        self._titlebar()
        tk.Frame(self, bg=BOR, height=1).pack(fill="x")
        body = tk.Frame(self, bg=BG0)
        body.pack(fill="both", expand=True)
        self._sidebar(body)
        tk.Frame(body, bg=BOR, width=1).pack(side="left", fill="y")
        self.content = tk.Frame(body, bg=BG1)
        self.content.pack(side="left", fill="both", expand=True)
        self.pages = {
            "drop":     self._pg_drop(),
            "pipeline": self._pg_pipeline(),
            "scripts":  self._pg_scripts(),
            "tasks":    self._pg_tasks(),
            "data":     self._pg_data(),
            "meeting":  self._pg_meeting(),
            "files":    self._pg_files(),
            "chat":     self._pg_chat(),
            "history":  self._pg_history(),
        }

    def _titlebar(self):
        tb = tk.Frame(self, bg=BG0, height=52)
        tb.pack(fill="x"); tb.pack_propagate(False)
        lf = tk.Frame(tb, bg=BG0)
        lf.pack(side="left", padx=16, pady=15)
        cv = tk.Canvas(lf, width=22, height=22, bg=BG0, highlightthickness=0)
        cv.pack(side="left")
        cv.create_rectangle(0,0,10,10,   fill=ACC,  outline="")
        cv.create_rectangle(12,0,22,10,  fill=ACC2, outline="")
        cv.create_rectangle(0,12,10,22,  fill=GRN,  outline="")
        cv.create_rectangle(12,12,22,22, fill=CYN,  outline="")
        Lbl(lf, "Office AI Studio", F_UIB, FG0, BG0).pack(side="left", padx=8)
        rf = tk.Frame(tb, bg=BG0)
        rf.pack(side="right", padx=14)
        for label, cmd in [("Notepad AI", lambda: NotepadAI(self, self.model_var)),
                            ("Terminal AI", lambda: TerminalAI(self, self.model_var))]:
            Btn(rf, label, cmd, "ghost", font=F_UIT).pack(side="left", padx=3)
        tk.Frame(rf, bg=BOR2, width=1, height=22).pack(side="left", padx=10)
        self._dot = tk.Canvas(rf, width=8, height=8, bg=BG0, highlightthickness=0)
        self._dot.pack(side="left", pady=22)
        self._dot_item = self._dot.create_oval(1,1,7,7, fill=FG3, outline="")
        self.status_lbl = Lbl(rf, "connecting…", F_UIT, FG3, BG0)
        self.status_lbl.pack(side="left", padx=(4,12))
        Lbl(rf, "Model", F_UIT, FG3, BG0).pack(side="left")
        self.model_cb = ttk.Combobox(rf, textvariable=self.model_var,
                                      values=[], width=28,
                                      state="readonly", font=F_UIS)
        self.model_cb.pack(side="left", padx=(6,0), ipady=2)

    def _sidebar(self, parent):
        sb = tk.Frame(parent, bg=BG0, width=196)
        sb.pack(side="left", fill="y"); sb.pack_propagate(False)
        tk.Frame(sb, bg=BG0, height=8).pack()
        self._nav_items = {}
        nav = [
            ("drop",     "⬇", "Smart Drop"),
            ("pipeline", "⚡", "Pipelines"),
            ("scripts",  "»",  "Scripts"),
            ("tasks",    "⏲", "Auto Tasks"),
            ("data",     "⊞",  "Data Tools"),
            ("meeting",  "◎",  "Meeting Notes"),
            ("files",    "≡",  "Files"),
            ("chat",     "◉",  "Chat"),
            ("history",  "⏱",  "History"),
        ]
        for key, icon, label in nav:
            f   = tk.Frame(sb, bg=BG0, cursor="hand2")
            bar = tk.Frame(f,  bg=BG0, width=3)
            il  = Lbl(f, icon,  F_UIS, FG3, BG0, width=2)
            tl  = Lbl(f, label, F_UIS, FG2, BG0, anchor="w")
            f.pack(fill="x", padx=6, pady=1)
            bar.pack(side="left", fill="y")
            il.pack(side="left", padx=(8,4), pady=8)
            tl.pack(side="left", fill="x", expand=True)
            self._nav_items[key] = (f, bar, il, tl)
            def _click(e=None, k=key): self._show(k)
            def _enter(e, ws=(f,il,tl), k=key):
                if self._cur != k:
                    for w in ws: w.config(bg=BG2)
                    f.config(bg=BG2)
            def _leave(e, ws=(f,il,tl), k=key):
                if self._cur != k:
                    for w in ws: w.config(bg=BG0)
                    f.config(bg=BG0)
            for w in [f, bar, il, tl]:
                w.bind("<Button-1>", _click)
                w.bind("<Enter>",    _enter)
                w.bind("<Leave>",    _leave)
        tk.Frame(sb, bg=BOR, height=1).pack(fill="x", padx=16, pady=10)
        Lbl(sb, "QUICK ACCESS", F_UIT, FG3, BG0, anchor="w"
            ).pack(fill="x", padx=22, pady=(0,4))
        for label, path in [("Desktop",   Path.home()/"Desktop"),
                             ("Downloads", Path.home()/"Downloads"),
                             ("Documents", Path.home()/"Documents")]:
            b = tk.Button(sb, text=f"  {label}", bg=BG0, fg=FG3, font=F_UIT,
                          bd=0, relief="flat", anchor="w", padx=16, pady=4,
                          cursor="hand2", activebackground=BG2, activeforeground=FG1,
                          command=lambda p=path: self._nav_files(p))
            b.pack(fill="x", padx=6)
        self._clock_lbl = Lbl(sb, "", F_UIT, FG3, BG0, justify="left")
        self._clock_lbl.pack(side="bottom", anchor="w", padx=22, pady=14)

    _cur = "drop"
    def _show(self, key):
        for p in self.pages.values(): p.pack_forget()
        self.pages[key].pack(fill="both", expand=True)
        self._cur = key
        for k, (f, bar, il, tl) in self._nav_items.items():
            sel = (k == key)
            bg   = BG2  if sel else BG0
            bbg  = ACC  if sel else BG0
            for w in [f, il, tl]: w.config(bg=bg)
            bar.config(bg=bbg)
            il.config(fg=ACC if sel else FG3)
            tl.config(fg=FG0 if sel else FG2)

    # ═════════════════════════════════════════════════════════════════
    #  PAGE: SMART DROP
    # ═════════════════════════════════════════════════════════════════
    def _pg_drop(self):
        page = tk.Frame(self.content, bg=BG1)
        SectionHdr(page, "Smart Drop Zone", "Drag files here - choose what to do with them", BG1)
        Div(page)

        # big drop area
        dz = tk.Frame(page, bg=BG2, height=180)
        dz.pack(fill="x", padx=20, pady=16)
        dz.pack_propagate(False)
        cv = tk.Canvas(dz, bg=BG2, highlightthickness=2,
                       highlightbackground=BOR)
        cv.pack(fill="both", expand=True, padx=2, pady=2)
        self._drop_cv = cv
        Lbl(cv, "⬇  Drop files here", ("Segoe UI",22,"bold"), FG3, BG2
            ).place(relx=0.5, rely=0.4, anchor="center")
        Lbl(cv, "or click to browse",  F_UIS, FG3, BG2
            ).place(relx=0.5, rely=0.62, anchor="center")
        cv.bind("<Button-1>", lambda e: self._drop_browse())
        self._drop_log_var = tk.StringVar(value="")

        # quick actions grid
        SectionHdr(page, "Quick File Actions", "No AI needed - instant automation", BG1)
        grid_f = tk.Frame(page, bg=BG1)
        grid_f.pack(fill="x", padx=20, pady=4)
        quick = [
            ("📅 Prefix date",     "rename_date",   BG2),
            ("🔡 Lowercase names", "rename_lower",  BG2),
            ("_  Fix spaces→_",   "rename_spaces",  BG2),
            ("⧉  Duplicate",      "duplicate",       BG2),
            ("📋 Copy to Desktop","copy_desktop",    BG2),
            ("📂 Move to Docs",   "move_docs",       BG2),
            ("#  Count lines",    "count_lines",     BG2),
            ("Σ  Word count",     "word_count",      BG2),
            ("🔒 MD5 hash",       "hash_md5",        BG2),
            ("📁 Open folder",    "open_folder",     BG2),
        ]
        for i, (label, key, bg) in enumerate(quick):
            b = Btn(grid_f, label,
                    lambda k=key: self._qa_run(k),
                    "default", font=F_UIS, px=14, py=10)
            b.grid(row=i//5, column=i%5, padx=6, pady=5, sticky="ew")
        for c in range(5): grid_f.columnconfigure(c, weight=1)

        # log
        Div(page, pady=8)
        Lbl(page, "Action log", F_UIT, FG2, BG1, anchor="w"
            ).pack(anchor="w", padx=20)
        lw = tk.Frame(page, bg=BG3)
        lw.pack(fill="both", expand=True, padx=20, pady=(4,16))
        self.drop_log = tk.Text(lw, bg=BG3, fg=FG1, font=F_MONO, bd=0,
                                 relief="flat", state="disabled",
                                 padx=12, pady=10, height=6)
        lsb = ttk.Scrollbar(lw, orient="vertical", command=self.drop_log.yview)
        self.drop_log.configure(yscrollcommand=lsb.set)
        self.drop_log.pack(side="left", fill="both", expand=True)
        lsb.pack(side="right", fill="y")
        self.drop_log.tag_config("ok",  foreground=GRN)
        self.drop_log.tag_config("err", foreground=RED)
        self.drop_log.tag_config("hdr", foreground=ACC2)

        self._drop_files = []

        # register main drop zone
        self.after(300, lambda: self._setup_drop_zone(cv))
        return page

    def _setup_drop_zone(self, widget):
        if not HAS_DND: return
        widget.drop_target_register(DND_FILES)
        def _on_drop(e):
            files = self._parse_drop(e.data)
            if files:
                self._drop_files = files
                self._drop_cv.config(highlightbackground=ACC)
                SmartDropDialog(self, files, self._handle_drop_action, self.model_var)
        widget.dnd_bind('<<Drop>>', _on_drop)

    def _setup_main_drop(self):
        """Register entire window as drop target."""
        if not HAS_DND: return
        self.drop_target_register(DND_FILES)
        def _on_drop(e):
            files = self._parse_drop(e.data)
            if files:
                self._drop_files = files
                SmartDropDialog(self, files, self._handle_drop_action, self.model_var)
        self.dnd_bind('<<Drop>>', _on_drop)

    def _drop_browse(self):
        files = filedialog.askopenfilenames()
        if files:
            self._drop_files = list(files)
            SmartDropDialog(self, list(files), self._handle_drop_action, self.model_var)

    def _qa_run(self, key):
        if not self._drop_files:
            files = filedialog.askopenfilenames()
            if not files: return
            self._drop_files = list(files)
        self._handle_drop_action(key, self._drop_files)

    def _handle_drop_action(self, key, files):
        self._show("drop")
        self._log_drop(f"▶  {key}  on {len(files)} file(s)\n", "hdr")
        if key.startswith("ai_") or key in ("add_pipeline","open_chat"):
            self._drop_ai_action(key, files)
        else:
            threading.Thread(target=self._drop_file_action,
                             args=(key, files), daemon=True).start()

    def _drop_file_action(self, key, files):
        for f in files:
            p = Path(f)
            action = BUILTIN_ACTIONS.get(key)
            if not action: continue
            try:
                r = action["fn"](p)
                msg = str(r) if r else f"✓  {p.name}"
                self.after(0, self._log_drop, msg+"\n", "ok")
            except Exception as ex:
                self.after(0, self._log_drop, f"✗  {p.name}  →  {ex}\n", "err")

    def _drop_ai_action(self, key, files):
        ai_prompts = {
            "ai_summarize":    "Summarize the following content concisely:\n\n{input}",
            "ai_extract":      "Extract all key data, numbers, dates and facts:\n\n{input}",
            "ai_actions":      "Extract a concrete action-item list:\n\n{input}",
            "ai_translate_en": "Translate to English:\n\n{input}",
            "ai_translate_pl": "Translate to Polish:\n\n{input}",
            "ai_review":       "Review and critique the following. Be specific:\n\n{input}",
        }
        if key == "add_pipeline":
            self._show("pipeline")
            if self.pipe_steps:
                self.pipe_steps[-1].files = files
                self.pipe_steps[-1].source = "files"
            else:
                st = PipeStep("Dropped Files", ACC, ACCD, "Process:\n\n{input}")
                st.source = "files"; st.files = files
                self.pipe_steps.append(st)
            self._draw_steps()
            self._toast(f"Added {len(files)} files to pipeline"); return
        if key == "open_chat":
            self._show("chat")
            content = "\n\n".join(
                f"--- {Path(f).name} ---\n{read_text(f, 3000)}" for f in files[:3])
            self.chat_inp.delete("1.0","end")
            self.chat_inp.insert("1.0", f"Please help me with these files:\n\n{content}")
            return
        prompt_tpl = ai_prompts.get(key, "Process:\n\n{input}")
        content = "\n\n".join(
            f"--- {Path(f).name} ---\n{read_text(f, 4000)}" for f in files)
        prompt = prompt_tpl.replace("{input}", content)
        self._log_drop("⬡ Running AI…\n", "hdr")
        self.generating = True; self.stop_flag = False
        threading.Thread(target=self._drop_ai_stream,
                         args=(prompt,), daemon=True).start()

    def _drop_ai_stream(self, prompt):
        full = ""
        try:
            full = ollama_stream(
                self.model_var.get(),
                [{"role":"user","content":prompt}],
                lambda t: self.after(0, self._log_drop, t),
                lambda: self.stop_flag)
            self.after(0, self._log_drop, "\n")
        except Exception as ex:
            self.after(0, self._log_drop, f"\n⚠ {ex}\n", "err")
        finally:
            self.generating = False
            if full:
                self.ai_history.append({
                    "ts":datetime.now().isoformat(), "step":"Smart Drop",
                    "model":self.model_var.get(), "input":"(dropped files)",
                    "output":full[:600]})
                history_save(self.ai_history)

    def _log_drop(self, text, tag=None):
        self.drop_log.config(state="normal")
        self.drop_log.insert("end", text, tag or "")
        self.drop_log.see("end")
        self.drop_log.config(state="disabled")

    # ═════════════════════════════════════════════════════════════════
    #  PAGE: SCRIPTS
    # ═════════════════════════════════════════════════════════════════
    def _pg_scripts(self):
        page = tk.Frame(self.content, bg=BG1)
        SectionHdr(page, "Scripts", "Write and run Python automation scripts", BG1,
                   "New Script", self._script_new)
        Div(page)
        split = tk.Frame(page, bg=BG1)
        split.pack(fill="both", expand=True)

        # list
        left = tk.Frame(split, bg=BG2, width=240)
        left.pack(side="left", fill="y"); left.pack_propagate(False)
        Lbl(left, "My Scripts", F_H3, FG0, BG2).pack(anchor="w", padx=14, pady=(14,6))
        self.script_lb = tk.Listbox(left, bg=BG2, fg=FG1, font=F_UIS,
                                     bd=0, relief="flat", selectmode="browse",
                                     activestyle="none", selectbackground=ACCD,
                                     selectforeground=ACC2)
        slsb = ttk.Scrollbar(left, orient="vertical", command=self.script_lb.yview)
        self.script_lb.configure(yscrollcommand=slsb.set)
        self.script_lb.pack(side="left", fill="both", expand=True)
        slsb.pack(side="right", fill="y")
        self.script_lb.bind("<<ListboxSelect>>", self._script_sel)

        lbf = tk.Frame(left, bg=BG2, pady=8)
        lbf.pack(fill="x", padx=8)
        Btn(lbf, "+ New",   self._script_new, "primary", font=F_UIT
            ).pack(side="left", padx=(0,4))
        Btn(lbf, "Delete",  self._script_del, "danger",  font=F_UIT).pack(side="left")

        Div(split, vertical=True)

        # editor
        right = tk.Frame(split, bg=BG1)
        right.pack(side="right", fill="both", expand=True)

        meta = tk.Frame(right, bg=BG2, pady=10)
        meta.pack(fill="x", padx=14, pady=(8,4))
        Lbl(meta, "Name:", F_UIT, FG2, BG2).pack(side="left")
        self.script_name_v = tk.StringVar()
        tk.Entry(meta, textvariable=self.script_name_v, bg=BG3, fg=FG0,
                 font=F_UIB, bd=0, relief="flat", insertbackground=ACC, width=28
                 ).pack(side="left", ipady=5, ipadx=8, padx=(6,16))
        Lbl(meta, "Desc:", F_UIT, FG2, BG2).pack(side="left")
        self.script_desc_v = tk.StringVar()
        tk.Entry(meta, textvariable=self.script_desc_v, bg=BG3, fg=FG0,
                 font=F_UIS, bd=0, relief="flat", insertbackground=ACC, width=40
                 ).pack(side="left", ipady=5, ipadx=8, padx=(6,0))

        # toolbar
        stb = tk.Frame(right, bg=BG1, pady=4)
        stb.pack(fill="x", padx=14)
        Btn(stb, "▶  Run (F5)", self._script_run, "primary", font=F_UIB, px=14, py=7
            ).pack(side="left")
        Btn(stb, "💾 Save",     self._script_save, "default", font=F_UIS
            ).pack(side="left", padx=8)
        Btn(stb, "⬡ Generate with AI", self._script_gen_ai, "subtle", font=F_UIS
            ).pack(side="left")
        Btn(stb, "⬡ Explain", self._script_explain, "subtle", font=F_UIS
            ).pack(side="left", padx=6)
        self.script_status = Lbl(stb, "", F_UIT, FG3, BG1)
        self.script_status.pack(side="right")
        Div(right, pady=4)

        # code editor
        cw = tk.Frame(right, bg=BG2)
        cw.pack(fill="both", expand=True, padx=14, pady=4)
        self.code_ed = tk.Text(cw, bg=BG2, fg=FG0, font=F_MONO,
                                bd=0, relief="flat", insertbackground=ACC,
                                wrap="none", padx=14, pady=12,
                                selectbackground=ACCD, tabs=("4c",))
        self.code_ed.bind("<F5>", lambda e: self._script_run())
        csb_v = ttk.Scrollbar(cw, orient="vertical", command=self.code_ed.yview)
        csb_h = ttk.Scrollbar(cw, orient="horizontal", command=self.code_ed.xview)
        self.code_ed.configure(yscrollcommand=csb_v.set, xscrollcommand=csb_h.set)
        csb_v.pack(side="right", fill="y")
        csb_h.pack(side="bottom", fill="x")
        self.code_ed.pack(side="left", fill="both", expand=True)

        Div(right, pady=2)
        Lbl(right, "Output", F_UIT, FG2, BG1).pack(anchor="w", padx=14, pady=(4,2))
        ow = tk.Frame(right, bg=BG3)
        ow.pack(fill="x", padx=14, pady=(0,10))
        self.script_out = tk.Text(ow, bg=BG3, fg=FG1, font=F_MONOS,
                                   bd=0, relief="flat", height=7,
                                   state="disabled", padx=10, pady=8)
        osb = ttk.Scrollbar(ow, orient="vertical", command=self.script_out.yview)
        self.script_out.configure(yscrollcommand=osb.set)
        self.script_out.pack(side="left", fill="both", expand=True)
        osb.pack(side="right", fill="y")
        self.script_out.tag_config("err", foreground=RED)
        self.script_out.tag_config("ok",  foreground=GRN)
        self.script_out.tag_config("ai",  foreground=ACC2)

        self._script_load_starters()
        self._script_refresh_lb()
        return page

    def _script_load_starters(self):
        if self.scripts: return
        starters = [
            Script("Rename with date", textwrap.dedent("""\
                # Rename all files in a folder: prefix with today's date
                import os
                from pathlib import Path
                from datetime import date

                folder = Path(r"C:/Users/YourName/Desktop")  # ← change this
                today  = date.today().strftime("%Y-%m-%d")

                for f in folder.iterdir():
                    if f.is_file() and not f.name.startswith(today):
                        f.rename(f.parent / f"{today}_{f.name}")
                        print(f"Renamed: {f.name}")

                print("Done!")
                """), "Prefix all files in a folder with today's date"),

            Script("CSV to summary", textwrap.dedent("""\
                # Read a CSV and print a quick summary
                import csv
                from pathlib import Path
                from collections import Counter

                path = Path(r"C:/path/to/your/file.csv")  # ← change this

                with open(path, encoding="utf-8", errors="replace") as f:
                    rows = list(csv.DictReader(f))

                print(f"Rows: {len(rows)}")
                print(f"Columns: {list(rows[0].keys()) if rows else []}")
                print(f"\\nFirst 3 rows:")
                for r in rows[:3]:
                    print(dict(r))
                """), "Read a CSV file and print summary stats"),

            Script("Find duplicate files", textwrap.dedent("""\
                # Find duplicate files in a folder by MD5 hash
                import hashlib
                from pathlib import Path
                from collections import defaultdict

                folder = Path(r"C:/Users/YourName/Downloads")  # ← change this

                hashes = defaultdict(list)
                for f in folder.rglob("*"):
                    if f.is_file():
                        h = hashlib.md5(f.read_bytes()).hexdigest()
                        hashes[h].append(f)

                dups = {h: fs for h, fs in hashes.items() if len(fs) > 1}
                if dups:
                    print(f"Found {len(dups)} duplicate group(s):")
                    for h, files in dups.items():
                        print(f"\\n  Hash {h[:8]}…")
                        for ff in files:
                            print(f"    {ff}")
                else:
                    print("No duplicates found.")
                """), "Find duplicate files by MD5 hash"),

            Script("Batch replace in text files", textwrap.dedent("""\
                # Batch find-and-replace across all .txt files in a folder
                from pathlib import Path

                folder      = Path(r"C:/your/folder")   # ← change this
                find_text   = "old company name"         # ← what to find
                replace_with= "new company name"         # ← replace with

                count = 0
                for f in folder.glob("*.txt"):
                    text = f.read_text(encoding="utf-8", errors="replace")
                    if find_text in text:
                        f.write_text(text.replace(find_text, replace_with), "utf-8")
                        print(f"Updated: {f.name}")
                        count += 1

                print(f"\\nDone - {count} file(s) updated.")
                """), "Find and replace text across multiple files"),

            Script("Organise files by extension", textwrap.dedent("""\
                # Sort files in a folder into subfolders by extension
                import shutil
                from pathlib import Path

                source = Path(r"C:/Users/YourName/Downloads")  # ← change this
                dry_run = True   # ← set False to actually move files

                moved = 0
                for f in source.iterdir():
                    if not f.is_file(): continue
                    ext = f.suffix.lower().lstrip(".") or "other"
                    dest_dir = source / ext.upper()
                    if not dry_run:
                        dest_dir.mkdir(exist_ok=True)
                        shutil.move(str(f), str(dest_dir / f.name))
                    print(f"{'[DRY]' if dry_run else 'MOVED'} {f.name} → {ext.upper()}/")
                    moved += 1

                print(f"\\n{'Would move' if dry_run else 'Moved'} {moved} file(s).")
                print("Set dry_run=False to actually move files.")
                """), "Sort files into subfolders by extension"),
        ]
        self.scripts = starters
        scripts_save(self.scripts)

    def _script_refresh_lb(self):
        self.script_lb.delete(0,"end")
        for s in self.scripts:
            self.script_lb.insert("end", f"  {s.name}")

    def _script_sel(self, _):
        sel = self.script_lb.curselection()
        if not sel: return
        s = self.scripts[sel[0]]
        self.script_name_v.set(s.name)
        self.script_desc_v.set(s.desc)
        self.code_ed.delete("1.0","end")
        self.code_ed.insert("1.0", s.code)
        self.script_status.config(text=f"Last run: {s.last_run[:16] or 'never'}")

    def _script_new(self):
        s = Script("New Script", "# Your script here\nprint('Hello!')", "")
        self.scripts.append(s)
        scripts_save(self.scripts)
        self._script_refresh_lb()
        self.script_lb.selection_clear(0,"end")
        self.script_lb.selection_set(len(self.scripts)-1)
        self._script_sel(None)

    def _script_save(self):
        sel = self.script_lb.curselection()
        if not sel: return
        s = self.scripts[sel[0]]
        s.name = self.script_name_v.get()
        s.desc = self.script_desc_v.get()
        s.code = self.code_ed.get("1.0","end")
        scripts_save(self.scripts)
        self._script_refresh_lb()
        self._toast("Script saved")

    def _script_del(self):
        sel = self.script_lb.curselection()
        if not sel: return
        if messagebox.askyesno("Delete", f"Delete '{self.scripts[sel[0]].name}'?"):
            del self.scripts[sel[0]]
            scripts_save(self.scripts)
            self._script_refresh_lb()
            self.code_ed.delete("1.0","end")

    def _script_run(self):
        code = self.code_ed.get("1.0","end")
        sel  = self.script_lb.curselection()
        self.script_out.config(state="normal")
        self.script_out.delete("1.0","end")
        self.script_out.config(state="disabled")
        self.script_status.config(text="running…")
        threading.Thread(target=self._run_script,
                         args=(code, sel[0] if sel else None), daemon=True).start()

    def _run_script(self, code, idx):
        import io as _io
        import sys as _sys
        buf = _io.StringIO()
        old_stdout = _sys.stdout
        old_stderr = _sys.stderr
        _sys.stdout = buf; _sys.stderr = buf
        err = None
        try:
            exec(compile(code, "<script>", "exec"), {"__name__":"__main__"})
        except Exception as ex:
            import traceback
            buf.write(traceback.format_exc())
            err = str(ex)
        finally:
            _sys.stdout = old_stdout; _sys.stderr = old_stderr
        out = buf.getvalue()
        self.after(0, self._script_done, out, err, idx)

    def _script_done(self, out, err, idx):
        self.script_out.config(state="normal")
        self.script_out.insert("1.0", out, "err" if err else "ok")
        self.script_out.config(state="disabled")
        ts = datetime.now().strftime("%H:%M:%S")
        self.script_status.config(text=f"Ran at {ts}" + (f"  ✗ {err[:40]}" if err else "  ✓ OK"),
                                   fg=RED if err else GRN)
        if idx is not None and 0 <= idx < len(self.scripts):
            self.scripts[idx].last_run = datetime.now().isoformat()
            self.scripts[idx].last_result = out[:200]
            scripts_save(self.scripts)

    def _script_gen_ai(self):
        task = self.script_name_v.get() or "Python automation script"
        prompt = (f"Write a clean, well-commented Python script that does the following:\n{task}\n\n"
                  f"Include: error handling, print statements showing progress, "
                  f"a clear comment at the top explaining usage. "
                  f"Use only Python standard library. Output ONLY the script, no explanation.")
        self.code_ed.delete("1.0","end")
        self.code_ed.insert("1.0","# Generating…\n")
        self.generating = True; self.stop_flag = False
        def _stream():
            self.after(0, lambda: (self.code_ed.delete("1.0","end")))
            try:
                ollama_stream(self.model_var.get(),
                    [{"role":"user","content":prompt}],
                    lambda t: self.after(0, self._code_tok, t),
                    lambda: self.stop_flag)
            except Exception as ex:
                self.after(0, self._code_tok, f"\n# Error: {ex}")
            finally: self.generating = False
        threading.Thread(target=_stream, daemon=True).start()

    def _code_tok(self, t):
        self.code_ed.insert("end", t)
        self.code_ed.see("end")

    def _script_explain(self):
        code = self.code_ed.get("1.0","end").strip()
        if not code: return
        prompt = f"Explain what this Python script does in plain English, step by step:\n\n```python\n{code[:4000]}\n```"
        self.script_out.config(state="normal")
        self.script_out.delete("1.0","end")
        self.script_out.config(state="disabled")
        self.generating = True; self.stop_flag = False
        def _stream():
            try:
                ollama_stream(self.model_var.get(),
                    [{"role":"user","content":prompt}],
                    lambda t: self.after(0, self._script_out_tok, t),
                    lambda: self.stop_flag)
            except Exception as ex:
                self.after(0, self._script_out_tok, f"\n⚠ {ex}")
            finally: self.generating = False
        threading.Thread(target=_stream, daemon=True).start()

    def _script_out_tok(self, t):
        self.script_out.config(state="normal")
        self.script_out.insert("end", t, "ai")
        self.script_out.see("end")
        self.script_out.config(state="disabled")

    # ═════════════════════════════════════════════════════════════════
    #  PAGE: AUTO TASKS
    # ═════════════════════════════════════════════════════════════════
    def _pg_tasks(self):
        page = tk.Frame(self.content, bg=BG1)
        SectionHdr(page, "Auto Tasks", "File automation rules - runs on drop or on schedule", BG1,
                   "+ New Task", self._task_new)
        Div(page)

        # task list
        sw, _, sf = scroll_wrap(page, BG1)
        sw.pack(fill="both", expand=True)
        self._task_scroll = sf
        self._task_scroll_wrap = sw

        Lbl(page,
            "Tip: Drop files on any task card to run it immediately, "
            "or enable scheduled tasks for automatic processing.",
            F_UIT, FG3, BG1, anchor="w", padx=18, pady=6).pack(fill="x", side="bottom")

        self._draw_tasks()
        return page

    def _draw_tasks(self):
        for w in self._task_scroll.winfo_children(): w.destroy()

        if not self.auto_tasks:
            Lbl(self._task_scroll,
                "No tasks yet.\nClick '+ New Task' to create your first automation rule.",
                F_UI, FG3, BG1, justify="center").pack(pady=60)
            return

        for task in self.auto_tasks:
            self._task_card(task)

    def _task_card(self, task):
        card = tk.Frame(self._task_scroll, bg=BG2)
        card.pack(fill="x", padx=16, pady=6)
        tk.Frame(card, bg=GRN if task.enabled else FG3, width=3).pack(side="left",fill="y")
        body = tk.Frame(card, bg=BG2)
        body.pack(side="left", fill="x", expand=True, padx=14, pady=12)

        hrow = tk.Frame(body, bg=BG2)
        hrow.pack(fill="x")
        nv = tk.StringVar(value=task.name)
        tk.Entry(hrow, textvariable=nv, bg=BG2, fg=FG0, font=F_UIB,
                 bd=0, relief="flat", insertbackground=ACC, width=28
                 ).pack(side="left")
        nv.trace("w", lambda *a, t=task, v=nv: setattr(t,"name",v.get()))

        ev = tk.BooleanVar(value=task.enabled)
        tk.Checkbutton(hrow, text="Enabled", variable=ev,
                       bg=BG2, fg=FG2, font=F_UIT, selectcolor=ACCD,
                       activebackground=BG2, activeforeground=GRN,
                       command=lambda t=task, v=ev: (setattr(t,"enabled",v.get()),
                                                      tasks_save(self.auto_tasks),
                                                      self._draw_tasks())
                       ).pack(side="right")

        meta = tk.Frame(body, bg=BG2)
        meta.pack(fill="x", pady=(6,0))
        Lbl(meta, f"Pattern: {task.pattern}  ·  Action: {task.action}  ·  Runs: {task.runs}",
            F_UIT, FG2, BG2).pack(side="left")

        bf = tk.Frame(body, bg=BG2)
        bf.pack(fill="x", pady=(8,0))
        Btn(bf, "▶ Run Now", lambda t=task: self._task_run_now(t), "success", font=F_UIT, px=10).pack(side="left")
        Btn(bf, "Edit",      lambda t=task: self._task_edit(t),    "default", font=F_UIT).pack(side="left",padx=6)
        Btn(bf, "Delete",    lambda t=task: self._task_delete(t),  "danger",  font=F_UIT).pack(side="left")
        if task.last_run:
            Lbl(bf, f"Last: {task.last_run[:16]}", F_UIT, FG3, BG2).pack(side="right")

        if HAS_DND:
            card.drop_target_register(DND_FILES)
            def _on_drop(e, t=task):
                files = self._parse_drop(e.data)
                if files: self._task_run_on_files(t, files)
            card.dnd_bind('<<Drop>>', _on_drop)

    def _task_new(self):
        t = AutoTask("New Task", "drop", "*.txt",
                     "rename_date", str(Path.home()), True)
        self.auto_tasks.append(t)
        tasks_save(self.auto_tasks)
        self._task_edit(t)

    def _task_edit(self, task):
        dlg = tk.Toplevel(self)
        dlg.title(f"Edit Task: {task.name}")
        dlg.geometry("500x420")
        dlg.configure(bg=BG1)
        dlg.resizable(False, False)

        tk.Frame(dlg, bg=BG0, pady=10).pack(fill="x")
        Lbl(dlg, "Edit Auto Task", F_H2, FG0, BG0).pack(anchor="w", padx=16, pady=(0,10))
        Div(dlg)

        fields = [
            ("Task Name:",   "name",       task.name),
            ("File Pattern:","pattern",    task.pattern),
            ("Target Dir:",  "target_dir", task.target_dir),
        ]
        vars_ = {}
        for label, key, val in fields:
            row = tk.Frame(dlg, bg=BG1)
            row.pack(fill="x", padx=16, pady=4)
            Lbl(row, label, F_UIS, FG2, BG1, width=14, anchor="w").pack(side="left")
            v = tk.StringVar(value=val)
            vars_[key] = v
            tk.Entry(row, textvariable=v, bg=BG3, fg=FG0, font=F_UIS,
                     bd=0, relief="flat", insertbackground=ACC
                     ).pack(side="left", fill="x", expand=True, ipady=5, ipadx=8)

        row = tk.Frame(dlg, bg=BG1)
        row.pack(fill="x", padx=16, pady=4)
        Lbl(row, "Action:", F_UIS, FG2, BG1, width=14, anchor="w").pack(side="left")
        av = tk.StringVar(value=task.action)
        action_names = list(BUILTIN_ACTIONS.keys())
        ttk.Combobox(row, textvariable=av, values=action_names,
                     state="readonly", width=28, font=F_UIS
                     ).pack(side="left", ipady=4)
        vars_["action"] = av

        Div(dlg, pady=8)
        ft = tk.Frame(dlg, bg=BG0, pady=10)
        ft.pack(fill="x", padx=14)
        def _save():
            task.name       = vars_["name"].get()
            task.pattern    = vars_["pattern"].get()
            task.target_dir = vars_["target_dir"].get()
            task.action     = vars_["action"].get()
            tasks_save(self.auto_tasks)
            self._draw_tasks()
            dlg.destroy()
        Btn(ft, "Cancel",    dlg.destroy, "ghost",   py=7).pack(side="right", padx=(8,0))
        Btn(ft, "✓ Save",    _save,       "primary", font=F_UIB, py=7).pack(side="right")

    def _task_delete(self, task):
        if messagebox.askyesno("Delete", f"Delete task '{task.name}'?"):
            self.auto_tasks = [t for t in self.auto_tasks if t.uid != task.uid]
            tasks_save(self.auto_tasks)
            self._draw_tasks()

    def _task_run_now(self, task):
        files = filedialog.askopenfilenames(title=f"Select files for: {task.name}")
        if files: self._task_run_on_files(task, list(files))

    def _task_run_on_files(self, task, files):
        matched = [f for f in files
                   if fnmatch.fnmatch(Path(f).name, task.pattern)]
        if not matched:
            self._toast(f"No files match pattern '{task.pattern}'", "info")
            return
        task.runs += len(matched)
        task.last_run = datetime.now().isoformat()
        tasks_save(self.auto_tasks)
        threading.Thread(
            target=run_action,
            args=(task.action, matched,
                  lambda msg: self.after(0, self._log_drop, msg)),
            daemon=True).start()
        self._show("drop")
        self._draw_tasks()
        self._toast(f"Running '{task.name}' on {len(matched)} file(s)")

    # ═════════════════════════════════════════════════════════════════
    #  PAGE: DATA TOOLS
    # ═════════════════════════════════════════════════════════════════
    def _pg_data(self):
        page = tk.Frame(self.content, bg=BG1)
        SectionHdr(page, "Data Tools", "CSV preview, clean, merge and export", BG1)
        Div(page)

        tb = tk.Frame(page, bg=BG1, pady=8)
        tb.pack(fill="x", padx=16)
        Btn(tb, "Open CSV",     self._data_open,   "primary",  font=F_UIS).pack(side="left")
        Btn(tb, "Merge CSVs",   self._data_merge,  "default",  font=F_UIS).pack(side="left",padx=6)
        Btn(tb, "Clean & Fix",  self._data_clean,  "default",  font=F_UIS).pack(side="left")
        Btn(tb, "Export TXT",   self._data_export, "default",  font=F_UIS).pack(side="left",padx=6)
        Btn(tb, "⬡ Analyse with AI", self._data_ai,"subtle",  font=F_UIS).pack(side="left")
        self.data_info = Lbl(tb, "", F_UIT, FG3, BG1)
        self.data_info.pack(side="right")

        Div(page, pady=2)

        cols_f = tk.Frame(page, bg=BG1)
        cols_f.pack(fill="x", padx=16, pady=4)
        Lbl(cols_f, "Columns:", F_UIT, FG2, BG1).pack(side="left")
        self.data_cols_lbl = Lbl(cols_f, "-", F_UIT, FG1, BG1)
        self.data_cols_lbl.pack(side="left", padx=6)

        # treeview for data
        tw = tk.Frame(page, bg=BG1)
        tw.pack(fill="both", expand=True, padx=16, pady=4)
        self.data_tree = ttk.Treeview(tw, show="headings", selectmode="browse")
        dtsb_v = ttk.Scrollbar(tw, orient="vertical",   command=self.data_tree.yview)
        dtsb_h = ttk.Scrollbar(tw, orient="horizontal", command=self.data_tree.xview)
        self.data_tree.configure(yscrollcommand=dtsb_v.set, xscrollcommand=dtsb_h.set)
        dtsb_v.pack(side="right", fill="y")
        dtsb_h.pack(side="bottom", fill="x")
        self.data_tree.pack(fill="both", expand=True)

        self._data_rows = []
        self._data_headers = []
        self._data_path = ""

        # AI result
        Div(page, pady=2)
        aw = tk.Frame(page, bg=BG3)
        aw.pack(fill="x", padx=16, pady=(0,10))
        self.data_ai_out = tk.Text(aw, bg=BG3, fg=FG1, font=F_MONO,
                                    bd=0, relief="flat", height=5,
                                    state="disabled", padx=10, pady=8)
        asb = ttk.Scrollbar(aw, orient="vertical", command=self.data_ai_out.yview)
        self.data_ai_out.configure(yscrollcommand=asb.set)
        self.data_ai_out.pack(side="left", fill="both", expand=True)
        asb.pack(side="right", fill="y")

        return page

    def _data_open(self, path=None):
        if not path:
            path = filedialog.askopenfilename(
                filetypes=[("CSV files","*.csv"),("All","*.*")])
        if not path: return
        self._data_path = path
        enc = detect_enc(path)
        try:
            rows = list(csv.reader(open(path, encoding=enc, errors="replace")))
        except Exception as ex:
            self._toast(str(ex), "err"); return
        if not rows: self._toast("Empty file","err"); return
        headers = rows[0]; data = rows[1:]
        self._data_headers = headers; self._data_rows = data
        self.data_tree.delete(*self.data_tree.get_children())
        self.data_tree["columns"] = headers
        for h in headers:
            self.data_tree.heading(h, text=h, anchor="w")
            self.data_tree.column(h, width=max(80, min(200, 10*len(h)+40)), minwidth=50)
        for row in data[:500]:
            self.data_tree.insert("","end", values=row)
        self.data_info.config(
            text=f"{len(data)} rows  ·  {len(headers)} cols  ·  {Path(path).name}")
        self.data_cols_lbl.config(text="  |  ".join(headers[:10]) +
                                   (f"  … +{len(headers)-10}" if len(headers)>10 else ""))
        self._toast(f"Loaded {len(data)} rows")

    def _data_merge(self):
        files = filedialog.askopenfilenames(
            title="Select CSV files to merge",
            filetypes=[("CSV","*.csv")])
        if not files: return
        all_rows = []; headers = None
        for f in files:
            enc = detect_enc(f)
            rows = list(csv.reader(open(f, encoding=enc, errors="replace")))
            if not rows: continue
            if headers is None: headers = rows[0]; all_rows.extend(rows[1:])
            else:
                if rows[0] == headers: all_rows.extend(rows[1:])
                else: all_rows.extend(rows[1:])  # merge anyway
        if not headers: self._toast("No data","err"); return
        out_path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV","*.csv")])
        if out_path:
            with open(out_path,"w",newline="",encoding="utf-8") as f:
                w = csv.writer(f); w.writerow(headers); w.writerows(all_rows)
            self._toast(f"Merged {len(all_rows)} rows → {Path(out_path).name}")
            self._data_open(out_path)

    def _data_clean(self):
        if not self._data_rows: self._toast("Open a CSV first","err"); return
        cleaned = []
        removed = 0
        seen = set()
        for row in self._data_rows:
            key = tuple(c.strip() for c in row)
            if key in seen or all(c.strip()=="" for c in row):
                removed += 1; continue
            seen.add(key)
            cleaned.append([c.strip() for c in row])
        self._data_rows = cleaned
        self.data_tree.delete(*self.data_tree.get_children())
        for row in cleaned[:500]:
            self.data_tree.insert("","end", values=row)
        self.data_info.config(text=f"{len(cleaned)} rows after clean (removed {removed})")
        self._toast(f"Cleaned: removed {removed} duplicates/empty rows")

    def _data_export(self):
        if not self._data_rows: self._toast("Open a CSV first","err"); return
        out = filedialog.asksaveasfilename(
            defaultextension=".txt", filetypes=[("Text","*.txt"),("CSV","*.csv")])
        if out:
            with open(out,"w",encoding="utf-8") as f:
                if out.endswith(".csv"):
                    w = csv.writer(f)
                    w.writerow(self._data_headers); w.writerows(self._data_rows)
                else:
                    f.write("\t".join(self._data_headers)+"\n")
                    for r in self._data_rows: f.write("\t".join(r)+"\n")
            self._toast(f"Exported → {Path(out).name}")

    def _data_ai(self):
        if not self._data_rows: self._toast("Open a CSV first","err"); return
        preview = "\t".join(self._data_headers) + "\n"
        for r in self._data_rows[:20]: preview += "\t".join(r) + "\n"
        prompt = (f"Analyse this CSV data ({len(self._data_rows)} rows, "
                  f"{len(self._data_headers)} columns).\n\n"
                  f"Describe:\n1. What this data appears to be about\n"
                  f"2. Key statistics or patterns\n3. Any data quality issues\n"
                  f"4. Useful insights or recommendations\n\nData preview:\n{preview}")
        self.data_ai_out.config(state="normal")
        self.data_ai_out.delete("1.0","end")
        self.data_ai_out.config(state="disabled")
        self.generating = True; self.stop_flag = False
        def _stream():
            try:
                ollama_stream(self.model_var.get(),
                    [{"role":"user","content":prompt}],
                    lambda t: self.after(0, self._data_ai_tok, t),
                    lambda: self.stop_flag)
            except Exception as ex:
                self.after(0, self._data_ai_tok, f"\n⚠ {ex}")
            finally: self.generating = False
        threading.Thread(target=_stream, daemon=True).start()

    def _data_ai_tok(self, t):
        self.data_ai_out.config(state="normal")
        self.data_ai_out.insert("end", t)
        self.data_ai_out.see("end")
        self.data_ai_out.config(state="disabled")

    # ═════════════════════════════════════════════════════════════════
    #  PAGE: MEETING NOTES
    # ═════════════════════════════════════════════════════════════════
    def _pg_meeting(self):
        page = tk.Frame(self.content, bg=BG1)
        SectionHdr(page, "Meeting Notes", "Paste notes → get tasks, summary & email draft", BG1)
        Div(page)

        split = tk.Frame(page, bg=BG1)
        split.pack(fill="both", expand=True)

        # input
        left = tk.Frame(split, bg=BG1)
        left.pack(side="left", fill="both", expand=True)

        Lbl(left, "Paste meeting notes here:", F_UIT, FG2, BG1
            ).pack(anchor="w", padx=16, pady=(10,4))
        nw = tk.Frame(left, bg=BG2)
        nw.pack(fill="both", expand=True, padx=16)
        self.meeting_inp = tk.Text(nw, bg=BG2, fg=FG0, font=F_UI,
                                    bd=0, relief="flat", wrap="word",
                                    insertbackground=ACC, padx=14, pady=12)
        msb = ttk.Scrollbar(nw, orient="vertical", command=self.meeting_inp.yview)
        self.meeting_inp.configure(yscrollcommand=msb.set)
        self.meeting_inp.pack(side="left", fill="both", expand=True)
        msb.pack(side="right", fill="y")

        bf = tk.Frame(left, bg=BG1, pady=8)
        bf.pack(fill="x", padx=16)
        Btn(bf, "⬡ Extract Tasks & Summary", self._meeting_run, "primary",
            font=F_UIB, px=18, py=8).pack(side="left")
        Btn(bf, "⬡ Draft Follow-up Email", self._meeting_email, "default",
            font=F_UIS).pack(side="left", padx=8)
        Btn(bf, "Clear", lambda: self.meeting_inp.delete("1.0","end"),
            "ghost", font=F_UIT).pack(side="left")

        Div(split, vertical=True)

        # output tabs
        right = tk.Frame(split, bg=BG1, width=480)
        right.pack(side="right", fill="both", expand=True)
        right.pack_propagate(False)

        self.meeting_nb = ttk.Notebook(right)
        self.meeting_nb.pack(fill="both", expand=True, padx=10, pady=10)

        for tab_name in ["Tasks & Decisions", "Summary", "Email Draft"]:
            f = tk.Frame(self.meeting_nb, bg=BG2)
            self.meeting_nb.add(f, text=f"  {tab_name}  ")
            tw = tk.Text(f, bg=BG2, fg=FG0, font=F_UI, bd=0,
                          relief="flat", wrap="word", state="disabled",
                          padx=14, pady=12, insertbackground=ACC)
            sb = ttk.Scrollbar(f, orient="vertical", command=tw.yview)
            tw.configure(yscrollcommand=sb.set)
            tw.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")

        # store references
        self._mtabs = [self.meeting_nb.tabs()[i] for i in range(3)]
        self._mtexts = []
        for tab in self.meeting_nb.tabs():
            f = self.meeting_nb.nametowidget(tab)
            self._mtexts.append(f.winfo_children()[0])

        return page

    def _meeting_run(self):
        notes = self.meeting_inp.get("1.0","end").strip()
        if not notes: self._toast("Paste some meeting notes first","err"); return
        prompt = (
            "You are a professional meeting assistant.\n"
            "From the meeting notes below, extract:\n\n"
            "## ACTION ITEMS\n"
            "List every task with: owner (if mentioned), deadline (if mentioned), priority\n\n"
            "## KEY DECISIONS\n"
            "List all decisions made\n\n"
            "## OPEN QUESTIONS\n"
            "List unresolved questions or items needing follow-up\n\n"
            f"Meeting notes:\n{notes[:5000]}"
        )
        summary_prompt = (
            "Write a concise executive summary (3-5 sentences) of these meeting notes. "
            "Focus on: purpose of meeting, key outcomes, next steps.\n\n"
            f"Notes:\n{notes[:5000]}"
        )
        for t in self._mtexts:
            t.config(state="normal"); t.delete("1.0","end"); t.config(state="disabled")
        self.generating = True; self.stop_flag = False
        def _go():
            try:
                # tasks tab
                ollama_stream(self.model_var.get(),
                    [{"role":"user","content":prompt}],
                    lambda t: self.after(0, self._mtok, 0, t),
                    lambda: self.stop_flag)
                # summary tab
                if not self.stop_flag:
                    ollama_stream(self.model_var.get(),
                        [{"role":"user","content":summary_prompt}],
                        lambda t: self.after(0, self._mtok, 1, t),
                        lambda: self.stop_flag)
            except Exception as ex:
                self.after(0, self._mtok, 0, f"\n⚠ {ex}")
            finally: self.generating = False
        threading.Thread(target=_go, daemon=True).start()

    def _meeting_email(self):
        notes = self.meeting_inp.get("1.0","end").strip()
        if not notes: self._toast("Paste meeting notes first","err"); return
        tasks_out = self._mtexts[0].get("1.0","end").strip()
        context = tasks_out if tasks_out else notes
        prompt = (
            "Write a professional follow-up email to meeting participants.\n"
            "Include: brief summary of what was discussed, action items with owners, "
            "next meeting / deadline if applicable.\n"
            "Tone: professional but warm. Length: concise.\n\n"
            f"Meeting content:\n{context[:3000]}"
        )
        t = self._mtexts[2]
        t.config(state="normal"); t.delete("1.0","end"); t.config(state="disabled")
        self.meeting_nb.select(2)
        self.generating = True; self.stop_flag = False
        def _go():
            try:
                ollama_stream(self.model_var.get(),
                    [{"role":"user","content":prompt}],
                    lambda tok: self.after(0, self._mtok, 2, tok),
                    lambda: self.stop_flag)
            except Exception as ex: self.after(0, self._mtok, 2, f"\n⚠ {ex}")
            finally: self.generating = False
        threading.Thread(target=_go, daemon=True).start()

    def _mtok(self, idx, t):
        tw = self._mtexts[idx]
        tw.config(state="normal"); tw.insert("end", t); tw.see("end")
        tw.config(state="disabled")

    # ═════════════════════════════════════════════════════════════════
    #  PAGE: PIPELINE  (preserved from v4 + improved)
    # ═════════════════════════════════════════════════════════════════
    def _pg_pipeline(self):
        page = tk.Frame(self.content, bg=BG1)
        top  = tk.Frame(page, bg=BG0, pady=10)
        top.pack(fill="x")
        Lbl(top, "Pipeline Builder", F_H1, FG0, BG0).pack(side="left", padx=20)
        self.run_btn  = Btn(top, "  ▶  Run  ", self._pipe_run, "primary", font=F_UIB, px=18, py=8)
        self.run_btn.pack(side="right", padx=16)
        self.stop_btn = Btn(top, "■ Stop",
                            lambda: setattr(self,"stop_flag",True), "danger", py=8)
        self.stop_btn.pack(side="right", padx=(0,4))
        self.stop_btn.config(state="disabled")
        tk.Frame(top, bg=BOR2, width=1, height=28).pack(side="right", padx=12)
        self.saved_cb = ttk.Combobox(top, state="readonly", width=20, font=F_UIS)
        self.saved_cb.pack(side="right")
        self.saved_cb.bind("<<ComboboxSelected>>", lambda e: self._pipe_load())
        Lbl(top, "Load:", F_UIT, FG2, BG0).pack(side="right", padx=(12,4))
        Btn(top, "Save",   self._pipe_save, "default", font=F_UIT).pack(side="right",padx=4)
        Btn(top, "New",    self._pipe_new,  "ghost",   font=F_UIT).pack(side="right")
        Btn(top, "Delete", self._pipe_del,  "danger",  font=F_UIT).pack(side="right",padx=4)
        self.pipe_name_v = tk.StringVar(value="My Pipeline")
        tk.Entry(top, textvariable=self.pipe_name_v, bg=BG3, fg=FG0, font=F_UIS,
                 bd=0, relief="flat", insertbackground=ACC, width=18
                 ).pack(side="right", ipady=5, ipadx=8, padx=(0,4))
        Lbl(top, "Name:", F_UIT, FG2, BG0).pack(side="right", padx=(12,4))
        tk.Frame(page, bg=BOR, height=1).pack(fill="x")

        cols = tk.Frame(page, bg=BG1)
        cols.pack(fill="both", expand=True)
        colA = tk.Frame(cols, bg=BG2, width=200)
        colA.pack(side="left", fill="y"); colA.pack_propagate(False)
        tk.Frame(cols, bg=BOR, width=1).pack(side="left", fill="y")
        colB = tk.Frame(cols, bg=BG1)
        colB.pack(side="left", fill="both", expand=True)
        tk.Frame(cols, bg=BOR, width=1).pack(side="left", fill="y")
        colC = tk.Frame(cols, bg=BG2, width=360)
        colC.pack(side="right", fill="y"); colC.pack_propagate(False)

        self._build_tpl_panel(colA)
        self._build_builder_panel(colB)
        self._build_output_panel(colC)
        self._saved_refresh()
        return page

    def _build_tpl_panel(self, parent):
        Lbl(parent,"Step Templates",F_H3,FG0,BG2).pack(anchor="w",padx=14,pady=(16,2))
        Lbl(parent,"Click to add →",F_UIT,FG3,BG2).pack(anchor="w",padx=14,pady=(0,8))
        tk.Frame(parent,bg=BOR,height=1).pack(fill="x",padx=14,pady=4)
        sw,_,si = scroll_wrap(parent,BG2)
        sw.pack(fill="both",expand=True)
        for name,color,dim_bg,instr in STEP_TEMPLATES:
            f   = tk.Frame(si,bg=BG2,cursor="hand2")
            bar = tk.Frame(f,bg=color,width=2)
            r   = tk.Frame(f,bg=BG2)
            nl  = Lbl(r,name,F_UIS,FG1,BG2,anchor="w")
            f.pack(fill="x",padx=8,pady=1)
            bar.pack(side="left",fill="y")
            r.pack(side="left",fill="x",expand=True,padx=10,pady=7)
            nl.pack(anchor="w")
            def _click(e=None,n=name,c=color,d=dim_bg,i=instr):
                st=PipeStep(n,c,d,i); self.pipe_steps.append(st)
                self._draw_steps(); self._toast(f"Added: {n}")
            for w in [f,bar,r,nl]:
                w.bind("<Button-1>",_click)
                w.bind("<Enter>", lambda e,ws=(f,r,nl): [x.config(bg=BG3) for x in ws])
                w.bind("<Leave>", lambda e,ws=(f,r,nl): [x.config(bg=BG2) for x in ws])

    def _build_builder_panel(self, parent):
        bh = tk.Frame(parent,bg=BG1,pady=8)
        bh.pack(fill="x",padx=16)
        Lbl(bh,"Builder",F_H3,FG0,BG1).pack(side="left")
        Btn(bh,"Clear All",self._pipe_new,"ghost",font=F_UIT).pack(side="right")
        dnd_hint = "  ⬇  Drop files onto a step" if HAS_DND else ""
        self._empty_lbl = Lbl(parent,
            f"← Add a step template\n\nEach step can use:\n"
            f"  • Output of previous step\n  • One or more files\n  • Custom text\n\n{dnd_hint}",
            F_UI,FG3,BG1,justify="center")
        self._empty_lbl.pack(pady=60)
        self._builder_sw,_,self._builder_frame = scroll_wrap(parent,BG1)
        self._builder_sw.pack_forget()
        self.after(300, lambda p=parent: self._register_drop(p))

    def _draw_steps(self):
        for w in self._builder_frame.winfo_children(): w.destroy()
        if not self.pipe_steps:
            self._empty_lbl.pack(pady=60); self._builder_sw.pack_forget()
            self.out_step_cb["values"] = []; return
        self._empty_lbl.pack_forget()
        self._builder_sw.pack(fill="both",expand=True)
        for i,step in enumerate(self.pipe_steps):
            if i>0: Lbl(self._builder_frame,"↓",F_H2,FG3,BG1).pack(pady=2)
            self._step_card(i,step)
        names=[f"Step {i+1}: {s.name}" for i,s in enumerate(self.pipe_steps)]
        self.out_step_cb["values"]=names
        if names: self.out_step_v.set(names[-1])

    def _step_card(self, idx, step):
        wrap = tk.Frame(self._builder_frame,bg=BG1)
        wrap.pack(fill="x",padx=14,pady=2)
        self.after(50, lambda w=wrap,s=step: self._register_drop_recursive(w,s))
        sc={"idle":FG3,"running":ORG,"done":GRN,"error":RED}.get(step.status,FG3)
        ss={"idle":"○","running":"◌","done":"●","error":"✕"}.get(step.status,"○")
        Lbl(wrap,ss,F_H3,sc,BG1).pack(side="left",anchor="n",pady=10,padx=(0,6))
        card=tk.Frame(wrap,bg=BG2); card.pack(side="left",fill="x",expand=True)
        tk.Frame(card,bg=step.color,width=3).pack(side="left",fill="y")
        body=tk.Frame(card,bg=BG2)
        body.pack(side="left",fill="x",expand=True,padx=12,pady=10)
        hrow=tk.Frame(body,bg=BG2); hrow.pack(fill="x")
        nb=tk.Frame(hrow,bg=step.dim_bg,padx=6,pady=1); nb.pack(side="left")
        Lbl(nb,f"{idx+1}",F_UIT,step.color,step.dim_bg).pack()
        nv=tk.StringVar(value=step.name)
        ne=tk.Entry(hrow,textvariable=nv,bg=BG2,fg=FG0,font=F_UIB,bd=0,
                    relief="flat",insertbackground=ACC,width=22); ne.pack(side="left",padx=8)
        nv.trace("w",lambda *a,s=step,v=nv: setattr(s,"name",v.get()))
        if step.status=="running": Lbl(hrow,"generating…",F_UIT,ORG,BG2).pack(side="left",padx=8)
        elif step.status=="done":  Lbl(hrow,"done",F_UIT,GRN,BG2).pack(side="left",padx=8)
        elif step.status=="error": Lbl(hrow,"error",F_UIT,RED,BG2).pack(side="left",padx=8)
        cf=tk.Frame(hrow,bg=BG2); cf.pack(side="right")
        for sym,fn in [("↑",lambda i=idx: self._mv(i,-1)),
                       ("↓",lambda i=idx: self._mv(i,+1)),
                       ("✕",lambda i=idx: self._rm(i))]:
            Btn(cf,sym,fn,"ghost",font=F_UIT,px=7,py=3).pack(side="left",padx=1)
        sr=tk.Frame(body,bg=BG2); sr.pack(fill="x",pady=(8,0))
        Lbl(sr,"Input:",F_UIT,FG3,BG2).pack(side="left")
        sv=tk.StringVar(value=step.source)
        for val,lbl in [("prev","↑ previous"),("files","📄 files"),("text","✎ text")]:
            tk.Radiobutton(sr,text=lbl,variable=sv,value=val,bg=BG2,fg=FG2,font=F_UIT,
                activebackground=BG2,activeforeground=FG0,selectcolor=ACCD,
                command=lambda s=step,v=sv:(setattr(s,"source",v.get()),self._draw_steps())
                ).pack(side="left",padx=8)
        if step.source=="files":
            fr=tk.Frame(body,bg=BG3); fr.pack(fill="x",pady=(6,0))
            self._register_drop(fr,step)
            hint="⬇ drop files  |  " if HAS_DND else ""
            fl=Lbl(fr,hint+self._files_lbl(step.files),F_UIT,CYN if step.files else FG3,BG3,
                   wraplength=420,justify="left")
            fl.pack(side="left",fill="x",expand=True,padx=8,pady=6)
            Btn(fr,"Browse…",lambda s=step: self._pick_files(s),"default",font=F_UIT,px=10,py=4
                ).pack(side="right",padx=6,pady=4)
        if step.source=="text":
            Lbl(body,"Text input:",F_UIT,FG3,BG2).pack(anchor="w",pady=(6,2))
            ct=tk.Text(body,bg=BG3,fg=FG0,font=F_UI,bd=0,relief="flat",
                       height=3,insertbackground=ACC,wrap="word",padx=8,pady=6)
            ct.insert("1.0",step.text); ct.pack(fill="x")
            ct.bind("<FocusOut>",lambda e,s=step,t=ct: setattr(s,"text",t.get("1.0","end").strip()))
        tk.Frame(body,bg=BOR,height=1).pack(fill="x",pady=(8,4))
        Lbl(body,"Instruction:",F_UIT,FG3,BG2).pack(anchor="w",pady=(0,2))
        it=tk.Text(body,bg=BG3,fg=FG1,font=F_MONOS,bd=0,relief="flat",
                   height=3,insertbackground=ACC,wrap="word",padx=8,pady=6)
        it.insert("1.0",step.instruction); it.pack(fill="x")
        it.bind("<FocusOut>",lambda e,s=step,t=it: setattr(s,"instruction",t.get("1.0","end").strip()))
        if step.status in ("done","error") and step.output:
            pbg=GRND if step.status=="done" else REDD
            pfg=GRN  if step.status=="done" else RED
            prev=step.output[:100].replace("\n"," ")
            Lbl(body,f"{ss}  {prev}{'…' if len(step.output)>100 else ''}",
                F_UIT,pfg,pbg,wraplength=440,justify="left",padx=8,pady=4
                ).pack(fill="x",pady=(6,0))

    def _files_lbl(self, files):
        if not files: return "No files selected"
        n=len(files); names=", ".join(Path(f).name for f in files[:3])
        return f"{names}{f'  +{n-3} more' if n>3 else ''}"

    def _pick_files(self, step):
        dlg = FilePicker(self, step.files)
        self.wait_window(dlg)
        if dlg.result is not None:
            step.files = dlg.result; self._draw_steps()

    def _mv(self,i,d):
        j=i+d
        if 0<=j<len(self.pipe_steps):
            self.pipe_steps[i],self.pipe_steps[j]=self.pipe_steps[j],self.pipe_steps[i]
            self._draw_steps()

    def _rm(self,i): del self.pipe_steps[i]; self._draw_steps()

    def _build_output_panel(self, parent):
        Lbl(parent,"Output",F_H3,FG0,BG2).pack(anchor="w",padx=14,pady=(16,4))
        self.out_step_v = tk.StringVar()
        self.out_step_cb = ttk.Combobox(parent,textvariable=self.out_step_v,
                                         state="readonly",width=28,font=F_UIS)
        self.out_step_cb.pack(anchor="w",padx=14,pady=(0,8))
        self.out_step_cb.bind("<<ComboboxSelected>>",self._show_step_out)
        tk.Frame(parent,bg=BOR,height=1).pack(fill="x",padx=14,pady=4)
        ow=tk.Frame(parent,bg=BG3)
        ow.pack(fill="both",expand=True,padx=10,pady=4)
        self.out_text=tk.Text(ow,bg=BG3,fg=FG1,font=F_MONO,bd=0,relief="flat",
                               wrap="word",state="disabled",padx=12,pady=12)
        osb=ttk.Scrollbar(ow,orient="vertical",command=self.out_text.yview)
        self.out_text.configure(yscrollcommand=osb.set)
        self.out_text.pack(side="left",fill="both",expand=True); osb.pack(side="right",fill="y")
        tk.Frame(parent,bg=BOR,height=1).pack(fill="x",padx=14,pady=(4,0))
        af=tk.Frame(parent,bg=BG2,pady=8); af.pack(fill="x",padx=10)
        Btn(af,"Copy",    self._copy_out, "default",font=F_UIT,px=10,py=5).pack(side="left",padx=(0,4))
        Btn(af,"→ Chat",  self._out2chat, "default",font=F_UIT,px=10,py=5).pack(side="left",padx=(0,4))
        Btn(af,"Save…",   self._save_out, "default",font=F_UIT,px=10,py=5).pack(side="left")
        pf=tk.Frame(parent,bg=BG2); pf.pack(fill="x",padx=14,pady=(0,14))
        self._prog_lbl=Lbl(pf,"",F_UIT,FG3,BG2); self._prog_lbl.pack(anchor="w")
        pb_bg=tk.Frame(pf,bg=BG4,height=3); pb_bg.pack(fill="x",pady=(3,0))
        self._pbar=tk.Frame(pb_bg,bg=ACC,height=3,width=0)
        self._pbar.place(x=0,y=0,relheight=1)

    def _upd_pbar(self,frac):
        self.update_idletasks()
        w=int(self._pbar.master.winfo_width()*frac)
        self._pbar.place(x=0,y=0,relheight=1,width=max(0,w))

    def _pipe_new(self):
        self.pipe_steps.clear(); self.pipe_name_v.set("My Pipeline")
        self._draw_steps(); self._out_clear()

    def _pipe_save(self):
        if not self.pipe_steps: self._toast("Add steps first","err"); return
        name=self.pipe_name_v.get().strip() or "Pipeline"
        self.pipes=[p for p in self.pipes if p["name"]!=name]
        self.pipes.insert(0,{"name":name,"steps":list(self.pipe_steps)})
        pipes_save(self.pipes); self._saved_refresh(); self._toast(f"Saved: {name}")

    def _pipe_load(self,_=None):
        name=self.saved_cb.get()
        p=next((x for x in self.pipes if x["name"]==name),None)
        if not p: return
        self.pipe_steps=[]
        for s in p["steps"]:
            st=PipeStep(s.name,s.color,s.dim_bg,s.instruction)
            st.source=s.source; st.files=list(s.files); st.text=s.text
            self.pipe_steps.append(st)
        self.pipe_name_v.set(name); self._draw_steps(); self._toast(f"Loaded: {name}")

    def _pipe_del(self):
        name=self.saved_cb.get()
        if not name: return
        self.pipes=[p for p in self.pipes if p["name"]!=name]
        pipes_save(self.pipes); self._saved_refresh(); self._toast(f"Deleted: {name}")

    def _saved_refresh(self):
        names=[p["name"] for p in self.pipes]
        self.saved_cb["values"]=names
        if names: self.saved_cb.set(names[0])

    def _pipe_run(self):
        if not self.pipe_steps: self._toast("No steps","err"); return
        if self.generating:     self._toast("Already running","err"); return
        for s in self.pipe_steps: s.status="idle"; s.output=""
        self._draw_steps(); self._out_clear()
        self.generating=True; self.stop_flag=False
        self.run_btn.config(state="disabled"); self.stop_btn.config(state="normal")
        threading.Thread(target=self._run_thread, daemon=True).start()

    def _run_thread(self):
        prev=""; total=len(self.pipe_steps)
        for i,step in enumerate(self.pipe_steps):
            if self.stop_flag:
                self.after(0,self._toast,"Stopped","err"); break
            step.status="running"
            self.after(0,self._draw_steps)
            self.after(0,self._prog_lbl.config,{"text":f"Step {i+1}/{total}: {step.name}…"})
            self.after(0,self._upd_pbar,i/total)
            if step.source=="prev": inp=prev
            elif step.source=="files":
                parts=[]
                for fp in step.files:
                    try: parts.append(f"--- {Path(fp).name} ---\n{read_text(fp,8000)}")
                    except Exception as ex: parts.append(f"--- {Path(fp).name} --- ERROR: {ex}")
                inp="\n\n".join(parts) if parts else "(no files)"
            else: inp=step.text
            if not inp.strip(): inp="(no input)"
            prompt=step.instruction.replace("{input}",inp)
            self.after(0,self._out_append,f"\n{'─'*52}\n  Step {i+1}: {step.name}\n{'─'*52}\n")
            result=""
            try:
                result=ollama_stream(
                    self.model_var.get(),
                    [{"role":"user","content":prompt}],
                    lambda t: self.after(0,self._out_append,t),
                    lambda: self.stop_flag, timeout=180)
                self.after(0,self._out_append,"\n")
                step.output=result; step.status="done"; prev=result
                self.ai_history.append({"ts":datetime.now().isoformat(),"step":step.name,
                    "model":self.model_var.get(),"input":inp[:300],"output":result[:600]})
                history_save(self.ai_history)
            except requests.ConnectionError:
                step.status="error"; step.output="No connection to Ollama"
                self.after(0,self._out_append,"\n⚠  No connection to Ollama\n"); break
            except Exception as ex:
                step.status="error"; step.output=str(ex)
                self.after(0,self._out_append,f"\n⚠  {ex}\n"); break
            self.after(0,self._draw_steps)
        self.generating=False
        self.after(0,self._pipe_finish,total)

    def _pipe_finish(self,total):
        self.run_btn.config(state="normal"); self.stop_btn.config(state="disabled")
        self._prog_lbl.config(text=f"✓  Done - {total} step{'s' if total!=1 else ''}")
        self._upd_pbar(1.0); self._toast(f"Pipeline complete ({total} steps)")
        try: self._build_hist_list()
        except: pass

    def _out_append(self,t):
        self.out_text.config(state="normal"); self.out_text.insert("end",t)
        self.out_text.see("end"); self.out_text.config(state="disabled")

    def _out_clear(self):
        self.out_text.config(state="normal"); self.out_text.delete("1.0","end")
        self.out_text.config(state="disabled"); self._prog_lbl.config(text=""); self._upd_pbar(0)

    def _show_step_out(self,_=None):
        val=self.out_step_v.get()
        try:
            i=int(val.split(":")[0].replace("Step","").strip())-1
            s=self.pipe_steps[i]
            self.out_text.config(state="normal"); self.out_text.delete("1.0","end")
            self.out_text.insert("1.0",s.output or "(no output)")
            self.out_text.config(state="disabled")
        except: pass

    def _copy_out(self):
        t=self.out_text.get("1.0","end").strip()
        if t: self.clipboard_clear(); self.clipboard_append(t); self._toast("Copied")

    def _save_out(self):
        t=self.out_text.get("1.0","end").strip()
        if not t: self._toast("Nothing to save","err"); return
        p=filedialog.asksaveasfilename(defaultextension=".txt",
            filetypes=[("Text","*.txt"),("Markdown","*.md"),("All","*.*")])
        if p:
            try: open(p,"w",encoding="utf-8").write(t); self._toast("Saved")
            except Exception as ex: self._toast(str(ex),"err")

    def _out2chat(self):
        t=self.out_text.get("1.0","end").strip()
        if t:
            self._show("chat"); self.chat_inp.delete("1.0","end")
            self.chat_inp.insert("1.0",f"Pipeline result:\n\n{t[:3000]}")

    # ═════════════════════════════════════════════════════════════════
    #  PAGE: FILES
    # ═════════════════════════════════════════════════════════════════
    def _pg_files(self):
        page=tk.Frame(self.content,bg=BG1)
        tb=tk.Frame(page,bg=BG0,pady=8); tb.pack(fill="x")
        Lbl(tb,"Files",F_H1,FG0,BG0).pack(side="left",padx=20)
        for label,cmd in [("↑ Up",self._go_up),("⌂ Home",lambda: self._nav_files(Path.home())),
                          ("↻",self._files_refresh)]:
            Btn(tb,label,cmd,"ghost",font=F_UIS).pack(side="left",padx=3)
        self.path_v=tk.StringVar(value=str(self.cur_path))
        pe=tk.Entry(tb,textvariable=self.path_v,bg=BG3,fg=FG0,font=F_MONO,
                    bd=0,relief="flat",insertbackground=ACC)
        pe.pack(side="left",fill="x",expand=True,padx=12,ipady=5,ipadx=8)
        pe.bind("<Return>",lambda e: self._nav_files(Path(self.path_v.get())))
        tk.Frame(page,bg=BOR,height=1).pack(fill="x")
        split=tk.Frame(page,bg=BG1); split.pack(fill="both",expand=True)
        left=tk.Frame(split,bg=BG1); left.pack(side="left",fill="both",expand=True)
        cols=("name","size","modified")
        self.ftree=ttk.Treeview(left,columns=cols,show="headings",selectmode="browse")
        for c,w,l in [("name",320,"Name"),("size",80,"Size"),("modified",150,"Modified")]:
            self.ftree.heading(c,text=l,anchor="w"); self.ftree.column(c,width=w,anchor="w" if c!="size" else "e")
        fsb=ttk.Scrollbar(left,orient="vertical",command=self.ftree.yview)
        self.ftree.configure(yscrollcommand=fsb.set)
        self.ftree.pack(side="left",fill="both",expand=True); fsb.pack(side="right",fill="y")
        self.ftree.bind("<Double-1>",self._fdbl)
        self.ftree.bind("<<TreeviewSelect>>",self._fsel)
        self.ftree.bind("<Button-3>",self._fctx)
        tk.Frame(split,bg=BOR,width=1).pack(side="left",fill="y")
        pv=tk.Frame(split,bg=BG2,width=300); pv.pack(side="right",fill="y"); pv.pack_propagate(False)
        ph=tk.Frame(pv,bg=BG2,pady=12); ph.pack(fill="x",padx=14)
        self.pv_name=Lbl(ph,"Select a file",F_H3,FG0,BG2,wraplength=260); self.pv_name.pack(anchor="w")
        self.pv_info=Lbl(ph,"",F_UIT,FG3,BG2,justify="left",wraplength=260); self.pv_info.pack(anchor="w",pady=(4,0))
        tk.Frame(pv,bg=BOR,height=1).pack(fill="x",padx=14)
        pw=tk.Frame(pv,bg=BG3); pw.pack(fill="both",expand=True,padx=10,pady=6)
        self.pv_text=tk.Text(pw,bg=BG3,fg=FG1,font=F_MONOS,bd=0,relief="flat",
                              wrap="word",state="disabled",padx=8,pady=8)
        ptsb=ttk.Scrollbar(pw,orient="vertical",command=self.pv_text.yview)
        self.pv_text.configure(yscrollcommand=ptsb.set)
        self.pv_text.pack(side="left",fill="both",expand=True); ptsb.pack(side="right",fill="y")
        tk.Frame(pv,bg=BOR,height=1).pack(fill="x",padx=14)
        af=tk.Frame(pv,bg=BG2,pady=8); af.pack(fill="x",padx=10)
        for label,cmd,style in [("→ Chat",self._file2chat,"default"),
            ("→ Pipeline",self._file2pipe,"default"),("→ Data Tools",self._file2data,"default"),
            ("Open",self._fopen,"default"),("Delete",self._fdel,"danger")]:
            Btn(af,label,cmd,style,font=F_UIT,px=10,py=5).pack(fill="x",pady=2)
        self.fstatus=Lbl(page,"",F_UIT,FG3,BG0,anchor="w",padx=16,pady=4)
        self.fstatus.pack(fill="x",side="bottom")
        tk.Frame(page,bg=BOR,height=1).pack(fill="x",side="bottom")
        self._files_refresh()
        return page

    def _nav_files(self,p):
        p=Path(p)
        if p.is_dir(): self.cur_path=p; self._files_refresh(); self._show("files")

    def _files_refresh(self):
        for i in self.ftree.get_children(): self.ftree.delete(i)
        try: items=sorted(self.cur_path.iterdir(),key=lambda x:(not x.is_dir(),x.name.lower()))
        except PermissionError: self._toast("Access denied","err"); return
        nd=nf=0
        for e in items:
            if e.name.startswith("."): continue
            try:
                s=e.stat(); sz=fsize(s.st_size) if e.is_file() else "-"
                m=datetime.fromtimestamp(s.st_mtime).strftime("%d %b %Y  %H:%M")
                ic="[dir]" if e.is_dir() else f"[{ficon(e.suffix)}]"
                self.ftree.insert("","end",iid=str(e),values=(f"{ic}  {e.name}",sz,m))
                nd+=e.is_dir(); nf+=e.is_file()
            except: pass
        self.path_v.set(str(self.cur_path))
        self.fstatus.config(text=f"  {nd} folders  ·  {nf} files    {self.cur_path}")

    def _go_up(self):
        p=self.cur_path.parent
        if p!=self.cur_path: self._nav_files(p)

    def _fdbl(self,e):
        sel=self.ftree.selection()
        if not sel: return
        p=Path(sel[0])
        if p.is_dir(): self._nav_files(p)
        else: self._fopen_path(p)

    def _fsel(self,_):
        sel=self.ftree.selection()
        if not sel: return
        p=Path(sel[0]); self.sel_path=str(p)
        self.pv_name.config(text=p.name)
        try:
            s=p.stat()
            self.pv_info.config(text=f"{'Dir' if p.is_dir() else (p.suffix[1:].upper() or 'File')}"
                +(f"  ·  {fsize(s.st_size)}" if p.is_file() else "")
                +f"\n{datetime.fromtimestamp(s.st_mtime).strftime('%d %b %Y  %H:%M')}")
        except: pass
        self.pv_text.config(state="normal"); self.pv_text.delete("1.0","end")
        self.sel_content=""
        if p.is_file() and p.suffix.lower() in TEXT_EXT:
            try:
                content=read_text(str(p),10000); self.sel_content=content
                self.pv_text.insert("1.0",content)
            except Exception as ex: self.pv_text.insert("1.0",str(ex))
        elif p.is_dir():
            try:
                ls=sorted(p.iterdir(),key=lambda x:(not x.is_dir(),x.name))[:40]
                self.pv_text.insert("1.0","\n".join(
                    ("[dir] " if i.is_dir() else "      ")+i.name for i in ls))
            except: pass
        else: self.pv_text.insert("1.0","[Preview not available]")
        self.pv_text.config(state="disabled")

    def _fopen(self):
        sel=self.ftree.selection()
        if sel: self._fopen_path(Path(sel[0]))

    def _fopen_path(self,p):
        try:
            if platform.system()=="Windows": os.startfile(str(p))
            else: subprocess.Popen(["xdg-open",str(p)])
        except Exception as ex: self._toast(str(ex),"err")

    def _fdel(self):
        sel=self.ftree.selection()
        if not sel: return
        p=Path(sel[0])
        if messagebox.askyesno("Delete",f"Delete  {p.name}?"):
            try:
                shutil.rmtree(p) if p.is_dir() else p.unlink()
                self._files_refresh(); self._toast("Deleted")
            except Exception as ex: self._toast(str(ex),"err")

    def _fctx(self,e):
        row=self.ftree.identify_row(e.y)
        if not row: return
        self.ftree.selection_set(row)
        m=tk.Menu(self,tearoff=0,bg=BG3,fg=FG1,activebackground=ACCD,
                   activeforeground=ACC2,font=F_UIS,bd=0,relief="solid")
        m.add_command(label="  Open",                command=self._fopen)
        m.add_command(label="  → Send to Chat",      command=self._file2chat)
        m.add_command(label="  → Add to Pipeline",   command=self._file2pipe)
        m.add_command(label="  → Data Tools (CSV)",  command=self._file2data)
        m.add_separator()
        m.add_command(label="  Copy Path",
                      command=lambda: (self.clipboard_clear(),self.clipboard_append(row)))
        m.add_separator()
        m.add_command(label="  Delete",              command=self._fdel)
        m.tk_popup(e.x_root,e.y_root)

    def _file2chat(self):
        if not self.sel_content: self._toast("No text content","err"); return
        self._show("chat"); self.chat_inp.delete("1.0","end")
        self.chat_inp.insert("1.0",
            f"Analyze the file `{Path(self.sel_path).name}`:\n\n"
            f"```\n{self.sel_content[:4000]}\n```")

    def _file2pipe(self):
        if not self.sel_path: self._toast("Select a file first","err"); return
        self._show("pipeline")
        if self.pipe_steps:
            last=self.pipe_steps[-1]
            if self.sel_path not in last.files: last.files.append(self.sel_path); last.source="files"
            self._draw_steps(); self._toast(f"Added to step: {last.name}")
        else: self._toast("Add a pipeline step first","err")

    def _file2data(self):
        if not self.sel_path: self._toast("Select a file first","err"); return
        if not self.sel_path.lower().endswith(".csv"):
            self._toast("Data Tools works with CSV files","info"); return
        self._show("data"); self._data_open(self.sel_path)

    # ═════════════════════════════════════════════════════════════════
    #  PAGE: CHAT
    # ═════════════════════════════════════════════════════════════════
    def _pg_chat(self):
        page=tk.Frame(self.content,bg=BG1)
        tb=tk.Frame(page,bg=BG0,pady=10); tb.pack(fill="x")
        Lbl(tb,"Chat",F_H1,FG0,BG0).pack(side="left",padx=20)
        tk.Frame(page,bg=BOR,height=1).pack(fill="x")
        Btn(tb,"New conversation",self._chat_clear,"ghost",font=F_UIS).pack(side="right",padx=16)
        msg_outer,self.chat_cv,self.chat_frame=scroll_wrap(page,BG1)
        msg_outer.pack(fill="both",expand=True)
        self._chat_sys("Send a message to start.")
        inp=tk.Frame(page,bg=BG0,pady=10); inp.pack(fill="x",padx=16,pady=(0,12))
        self.chat_inp=tk.Text(inp,bg=BG2,fg=FG0,font=F_UI,bd=0,relief="flat",
                               height=4,insertbackground=ACC,wrap="word",padx=12,pady=10)
        self.chat_inp.pack(side="left",fill="both",expand=True)
        self.chat_inp.bind("<Control-Return>",lambda e: self._chat_send())
        rb=tk.Frame(inp,bg=BG0); rb.pack(side="right",padx=(10,0))
        self.chat_send_b=Btn(rb,"▶ Send\n(Ctrl+↵)",self._chat_send,"primary",font=F_UIB,px=14,py=10)
        self.chat_send_b.pack()
        self.chat_stop_b=Btn(rb,"■ Stop",lambda: setattr(self,"stop_flag",True),"danger",py=6)
        self.chat_stop_b.pack(pady=(6,0)); self.chat_stop_b.config(state="disabled")
        self._chat_hist=[]
        return page

    def _chat_sys(self,text):
        row=tk.Frame(self.chat_frame,bg=BG1,pady=4); row.pack(fill="x",padx=40)
        Lbl(row,text,F_UIT,FG3,BG2,padx=14,pady=6,wraplength=700).pack(anchor="center")

    def _chat_user(self,text):
        row=tk.Frame(self.chat_frame,bg=BG1,pady=6); row.pack(fill="x",padx=40)
        bub=tk.Frame(row,bg=ACCD); bub.pack(side="right")
        Lbl(bub,text[:500]+("…" if len(text)>500 else ""),F_UI,ACC2,ACCD,
            wraplength=560,justify="left",padx=14,pady=10).pack()
        Lbl(row,"you",F_UIT,FG3,BG1).pack(side="right",padx=8,anchor="s")

    def _chat_ai_box(self):
        row=tk.Frame(self.chat_frame,bg=BG1,pady=6); row.pack(fill="x",padx=40)
        Lbl(row,f"⬡  {self.model_var.get()}",F_UIT,FG3,BG1).pack(anchor="w")
        bub=tk.Frame(row,bg=BG2); bub.pack(anchor="w")
        tw=tk.Text(bub,bg=BG2,fg=FG0,font=F_UI,bd=0,relief="flat",wrap="word",
                   width=66,height=2,padx=14,pady=10,selectbackground=ACCD,insertbackground=ACC)
        tw.pack(); return tw

    def _chat_send(self):
        if self.generating: return
        text=self.chat_inp.get("1.0","end").strip()
        if not text: return
        self.chat_inp.delete("1.0","end"); self._chat_user(text)
        self._chat_hist.append({"role":"user","content":text})
        self.generating=True; self.stop_flag=False
        self.chat_send_b.config(state="disabled"); self.chat_stop_b.config(state="normal")
        tw=self._chat_ai_box(); self._chat_scroll()
        threading.Thread(target=self._stream_chat,args=(tw,),daemon=True).start()

    def _stream_chat(self,tw):
        full=""
        try:
            full=ollama_stream(self.model_var.get(),self._chat_hist,
                lambda t: self.after(0,self._chat_tok,tw,t),lambda: self.stop_flag)
        except requests.ConnectionError:
            self.after(0,self._chat_tok,tw,"\n⚠  No connection to Ollama.")
        except Exception as ex: self.after(0,self._chat_tok,tw,f"\n⚠  {ex}")
        finally:
            if full: self._chat_hist.append({"role":"assistant","content":full})
            self.after(0,self._chat_done)

    def _chat_tok(self,tw,t):
        try:
            tw.config(state="normal") if tw.cget("state")=="disabled" else None
            tw.insert("end",t)
            lines=int(tw.index("end-1c").split(".")[0])
            tw.config(height=min(max(lines,2),40))
            self._chat_scroll()
        except tk.TclError: pass

    def _chat_done(self):
        self.generating=False; self.chat_send_b.config(state="normal")
        self.chat_stop_b.config(state="disabled")

    def _chat_clear(self):
        for w in self.chat_frame.winfo_children(): w.destroy()
        self._chat_hist.clear(); self._chat_sys("New conversation started.")

    def _chat_scroll(self):
        self.chat_cv.update_idletasks(); self.chat_cv.yview_moveto(1.0)

    # ═════════════════════════════════════════════════════════════════
    #  PAGE: HISTORY
    # ═════════════════════════════════════════════════════════════════
    def _pg_history(self):
        page=tk.Frame(self.content,bg=BG1)
        tb=tk.Frame(page,bg=BG0,pady=10); tb.pack(fill="x")
        Lbl(tb,"History",F_H1,FG0,BG0).pack(side="left",padx=20)
        tk.Frame(page,bg=BOR,height=1).pack(fill="x")
        Btn(tb,"Clear All",self._clear_hist,"danger",font=F_UIS).pack(side="right",padx=16)
        split=tk.Frame(page,bg=BG1); split.pack(fill="both",expand=True)
        left=tk.Frame(split,bg=BG2,width=300); left.pack(side="left",fill="y"); left.pack_propagate(False)
        Lbl(left,"Recent runs",F_H3,FG0,BG2).pack(anchor="w",padx=14,pady=(14,6))
        self.hist_lb=tk.Listbox(left,bg=BG2,fg=FG1,font=F_UIS,bd=0,relief="flat",
                                 selectmode="browse",activestyle="none",
                                 selectbackground=ACCD,selectforeground=ACC2)
        hlsb=ttk.Scrollbar(left,orient="vertical",command=self.hist_lb.yview)
        self.hist_lb.configure(yscrollcommand=hlsb.set)
        self.hist_lb.pack(side="left",fill="both",expand=True); hlsb.pack(side="right",fill="y")
        self.hist_lb.bind("<<ListboxSelect>>",self._hist_sel)
        tk.Frame(split,bg=BOR,width=1).pack(side="left",fill="y")
        right=tk.Frame(split,bg=BG1); right.pack(side="right",fill="both",expand=True)
        self.hist_meta=Lbl(right,"",F_UIS,FG2,BG2,justify="left",wraplength=600,anchor="w",padx=16,pady=10)
        self.hist_meta.pack(fill="x",padx=14,pady=(14,4))
        dw=tk.Frame(right,bg=BG3); dw.pack(fill="both",expand=True,padx=14,pady=4)
        self.hist_out=tk.Text(dw,bg=BG3,fg=FG1,font=F_MONO,bd=0,relief="flat",
                               wrap="word",state="disabled",padx=12,pady=12)
        dsb=ttk.Scrollbar(dw,orient="vertical",command=self.hist_out.yview)
        self.hist_out.configure(yscrollcommand=dsb.set)
        self.hist_out.pack(side="left",fill="both",expand=True); dsb.pack(side="right",fill="y")
        af=tk.Frame(right,bg=BG1,pady=8); af.pack(fill="x",padx=14)
        Btn(af,"Copy Output",self._hist_copy,"default",font=F_UIT).pack(side="left")
        Btn(af,"→ Chat",     self._hist2chat,"default",font=F_UIT).pack(side="left",padx=6)
        self._build_hist_list()
        return page

    def _build_hist_list(self):
        try:
            self.hist_lb.delete(0,"end")
            for item in reversed(self.ai_history[-100:]):
                ts=item.get("ts","")[:16].replace("T"," ")
                self.hist_lb.insert("end",f"  {ts}  {item.get('step','?')}")
        except: pass

    def _hist_sel(self,_):
        sel=self.hist_lb.curselection()
        if not sel: return
        idx=len(self.ai_history)-1-sel[0]
        if 0<=idx<len(self.ai_history):
            item=self.ai_history[idx]
            self.hist_meta.config(text=f"Step: {item.get('step','')}   ·   "
                f"Model: {item.get('model','')}   ·   {item.get('ts','')[:16].replace('T',' ')}\n"
                f"Input: {item.get('input','')[:200]}")
            self.hist_out.config(state="normal"); self.hist_out.delete("1.0","end")
            self.hist_out.insert("1.0",item.get("output","")); self.hist_out.config(state="disabled")

    def _hist_copy(self):
        t=self.hist_out.get("1.0","end").strip()
        if t: self.clipboard_clear(); self.clipboard_append(t); self._toast("Copied")

    def _hist2chat(self):
        t=self.hist_out.get("1.0","end").strip()
        if t: self._show("chat"); self.chat_inp.delete("1.0","end"); self.chat_inp.insert("1.0",t[:3000])

    def _clear_hist(self):
        if messagebox.askyesno("Clear History","Delete all history?"):
            self.ai_history.clear(); history_save(self.ai_history); self._build_hist_list()

    # ═════════════════════════════════════════════════════════════════
    #  DRAG & DROP HELPERS
    # ═════════════════════════════════════════════════════════════════
    def _parse_drop(self, data):
        paths=[]
        for m in re.finditer(r'\{([^}]+)\}', data): paths.append(m.group(1))
        remainder=re.sub(r'\{[^}]+\}','',data).strip()
        if remainder: paths.extend(p for p in remainder.split() if p)
        return [p for p in paths if Path(p).is_file()]

    def _register_drop(self, widget, step=None):
        if not HAS_DND: return
        widget.drop_target_register(DND_FILES)
        def _on_drop(event, s=step):
            files=self._parse_drop(event.data)
            if not files: return
            if s is not None: target=s
            elif self.pipe_steps: target=self.pipe_steps[-1]
            else:
                target=PipeStep("Dropped Files",ACC,ACCD,"Process:\n\n{input}")
                target.source="files"; target.files=[]; self.pipe_steps.append(target)
            added=0
            for f in files:
                if f not in target.files: target.files.append(f); added+=1
            target.source="files"; self._draw_steps(); self._show("pipeline")
            if added: self._toast(f"Added {added} file{'s' if added!=1 else ''} → {target.name}")
            else: self._toast("Already in step","info")
        widget.dnd_bind('<<Drop>>',_on_drop)

    def _register_drop_recursive(self, widget, step):
        self._register_drop(widget, step)
        for child in widget.winfo_children():
            self._register_drop_recursive(child, step)

    # ═════════════════════════════════════════════════════════════════
    #  UTILITIES
    # ═════════════════════════════════════════════════════════════════
    def _fetch_models_once(self):
        """Fetch available Ollama models immediately on startup."""
        for attempt in range(5):
            try:
                r = requests.get(f"{OLLAMA}/api/tags", timeout=4)
                models = [m["name"] for m in r.json().get("models", [])]
                if models:
                    self.after(0, self._online, models)
                else:
                    self.after(0, self._no_models)
                return
            except: pass
            time.sleep(1.5)
        self.after(0, self._offline)

    def _poll_ollama(self):
        while True:
            try:
                r = requests.get(f"{OLLAMA}/api/tags", timeout=4)
                models = [m["name"] for m in r.json().get("models", [])]
                if models: self.after(0, self._online, models)
                else: self.after(0, self._no_models)
            except: self.after(0, self._offline)
            time.sleep(12)

    def _online(self, models):
        self._dot.itemconfig(self._dot_item, fill=GRN)
        self.status_lbl.config(
            text=f"ollama  ·  {len(models)} model{'s' if len(models)!=1 else ''}",
            fg=GRN)
        self.model_cb["values"] = models
        if models and self.model_var.get() not in models:
            self.model_var.set(models[0])

    def _no_models(self):
        self._dot.itemconfig(self._dot_item, fill=ORG)
        self.status_lbl.config(text="ollama running  ·  no models - run: ollama pull llama3.2:3b", fg=ORG)
        self.model_cb["values"] = []
        self.model_var.set("")

    def _offline(self):
        self._dot.itemconfig(self._dot_item, fill=RED)
        self.status_lbl.config(text="ollama offline  ·  run: ollama serve", fg=RED)
        self.model_cb["values"] = []

    def _tick(self):
        try:
            now=datetime.now()
            self._clock_lbl.config(text=f"{now.strftime('%H:%M:%S')}\n{now.strftime('%d %b %Y')}")
        except: pass
        self.after(1000,self._tick)

    def _toast(self,msg,kind="ok"):
        cols={"ok":(GRND,GRN),"err":(REDD,RED),"info":(ACCD,ACC2)}.get(kind,(BG3,FG0))
        t=tk.Toplevel(self); t.overrideredirect(True)
        t.attributes("-topmost",True); t.attributes("-alpha",0.95)
        tk.Frame(t,bg=BOR2,padx=1,pady=1).pack()
        inner=tk.Frame(t.children["!frame"],bg=cols[0],padx=18,pady=10); inner.pack()
        Lbl(inner,msg,F_UIS,cols[1],cols[0]).pack()
        self.update_idletasks()
        x=self.winfo_x()+self.winfo_width()-340
        y=self.winfo_y()+self.winfo_height()-80
        t.geometry(f"+{x}+{y}"); t.after(2800,t.destroy)

# ══════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    App().mainloop()
