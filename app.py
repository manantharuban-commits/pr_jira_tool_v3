"""
PR  Jira Ticket Tool  v3
─────────────────────────────────────────────────────────────────────────────
NEW in v3
  • Word Document Preview  – live HTML-rendered preview panel (screenshots +
    comments shown inline) with Print / Open-in-Word buttons
  • Dynamic Jira Fields    – add / remove / reorder fields; each field has a
    label, type (text / dropdown / date / number), default value, and an
    "include in Jira" toggle
  • Custom fields          – "Add Field" button lets user create any field on
    the fly with label + optional dropdown choices
  • Default-value manager  – Settings → Field Defaults; saved to config.json
    so every fresh launch pre-fills values automatically
─────────────────────────────────────────────────────────────────────────────
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import json, os, re, threading, webbrowser, base64, copy, tempfile, html
from datetime import datetime

import requests
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─────────────────────────────────────────────────────────────────────────────
#  Paths & palette
# ─────────────────────────────────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

APP_NAME     = "PRism"
APP_SUBTITLE = "PR  →  Jira · Doc · SQL"


def _draw_prism(d, sz):
    """Draw PRism logo onto an ImageDraw canvas of size sz×sz (RGB, dark bg)."""
    cx, cy = sz // 2 - sz // 25, sz // 2
    # prism body
    tl   = (cx - sz * 24 // 100, cy - sz * 30 // 100)
    bl   = (cx - sz * 24 // 100, cy + sz * 30 // 100)
    apex = (cx + sz * 30 // 100, cy)
    d.polygon([tl, bl, apex], fill=(30, 40, 60))
    bw = max(2, sz // 50)
    d.polygon([tl, bl, apex], outline=(88, 166, 255), width=bw)
    # incoming beam
    d.line([(sz * 4 // 100, cy), (tl[0], cy)],
           fill=(210, 225, 255), width=max(2, sz // 60))
    # refracted rays
    ex = sz - sz * 6 // 100
    for ey, col in [
        (cy - sz * 25 // 100, (63, 185, 80)),
        (cy,                  (88, 166, 255)),
        (cy + sz * 25 // 100, (210, 153, 34)),
    ]:
        lw = max(2, sz // 48)
        d.line([(apex[0], apex[1]), (ex, ey)], fill=col, width=lw)
    return tl, bl, apex


def _generate_app_assets():
    """Generate prism_icon.ico + prism_logo.png in BASE_DIR.
    Uses RGB images (no alpha) for maximum ICO compatibility.
    Returns (logo_path, icon_path) or (None, None) on failure."""
    try:
        from PIL import Image, ImageDraw, ImageFont

        icon_path = os.path.join(BASE_DIR, "prism_icon.ico")
        logo_path = os.path.join(BASE_DIR, "prism_logo.png")
        BG = (13, 17, 23)   # #0d1117

        # Build one image per ICO size to keep quality at small sizes
        ico_imgs = []
        for sz in (256, 128, 64, 48, 32, 16):
            img = Image.new("RGB", (sz, sz), BG)
            d   = ImageDraw.Draw(img)
            # rounded-square bg tint
            pad = max(1, sz // 20)
            try:
                d.rounded_rectangle([pad, pad, sz - pad - 1, sz - pad - 1],
                                    radius=sz // 7, fill=(22, 27, 34))
            except AttributeError:
                d.rectangle([pad, pad, sz - pad - 1, sz - pad - 1],
                            fill=(22, 27, 34))
            _draw_prism(d, sz)
            # "PR" text only for larger sizes
            if sz >= 48:
                fnt = None
                for fp in ["C:/Windows/Fonts/arialbd.ttf",
                           "C:/Windows/Fonts/calibrib.ttf",
                           "C:/Windows/Fonts/segoeui.ttf"]:
                    try:
                        fnt = ImageFont.truetype(fp, sz * 22 // 100)
                        break
                    except Exception:
                        pass
                if fnt:
                    tx = sz * 18 // 100
                    ty = sz * 34 // 100
                    d.text((tx, ty), "PR", font=fnt, fill=(255, 255, 255))
            ico_imgs.append(img)

        # Save ICO — first image is primary, rest are alternate sizes
        ico_imgs[0].save(
            icon_path, format="ICO",
            append_images=ico_imgs[1:],
            sizes=[(i.width, i.height) for i in ico_imgs],
        )

        # Save PNG logo at 256
        ico_imgs[0].save(logo_path, format="PNG")
        return logo_path, icon_path

    except Exception as exc:
        return None, None

THEMES = {
    "dark": dict(
        bg         = "#0d1117",
        surface    = "#161b22",
        card       = "#1c2128",
        border     = "#30363d",
        accent     = "#58a6ff",
        green      = "#3fb950",
        yellow     = "#d29922",
        red        = "#f85149",
        purple     = "#bc8cff",
        orange     = "#e3b341",
        text       = "#e6edf3",
        muted      = "#8b949e",
        inp        = "#161b22",
        hover      = "#1f6feb",
        cb_sel     = "#2ea043",   # checkbox checked indicator (bright green, visible on dark bg)
        tag        = "dark",
        st_badge   = dict(added="#1f6931", modified="#1c4a6e", removed="#8a1f1f",
                          renamed="#5a2d82", copied="#1c4a6e", changed="#1c4a6e"),
        st_row     = dict(added="#0e2318", modified="#0c1c2c", removed="#200e0e",
                          renamed="#180e2a", copied="#0c1c2c", changed="#0c1c2c"),
    ),
    "light": dict(
        bg         = "#f6f8fa",
        surface    = "#ffffff",
        card       = "#ffffff",
        border     = "#d0d7de",
        accent     = "#0969da",
        green      = "#1a7f37",
        yellow     = "#9a6700",
        red        = "#cf222e",
        purple     = "#8250df",
        orange     = "#9a6700",
        text       = "#1f2328",
        muted      = "#57606a",
        inp        = "#f6f8fa",
        hover      = "#0550ae",
        cb_sel     = "#cce5ff",   # checkbox checked indicator
        tag        = "light",
        st_badge   = dict(added="#1a7f37", modified="#0969da", removed="#cf222e",
                          renamed="#8250df", copied="#0969da", changed="#0969da"),
        st_row     = dict(added="#eafbee", modified="#eef6ff", removed="#fff5f5",
                          renamed="#f7f0ff", copied="#eef6ff", changed="#eef6ff"),
    ),
}

C = dict(THEMES["dark"])   # mutable — updated in-place on theme switch

STATUS_EMOJI  = dict(added="A", modified="M", removed="D", renamed="R", copied="C", changed="~")
STATUS_BADGE  = dict(C["st_badge"])
STATUS_ROW    = dict(C["st_row"])

STATUS_VERB   = dict(added="Add", modified="Update", removed="Remove",
                     renamed="Rename", copied="Copy", changed="Change")

FIELD_HINTS = {
    "owner":       "Person responsible for resolving or triaging this issue.",
    "environment": "Environment where the issue was discovered (e.g. Production, Staging).",
    "issue_type":  "Jira issue category — Bug, Task, Story, Epic, etc.",
    "dependency":  "Enter 'Yes' if this ticket depends on another PR or ticket.",
    "git_link":    "GitHub Pull Request URL — auto-filled when a PR is fetched.",
    "doc_name":    "Output file name for the generated Word document (no .docx extension). Auto-synced from Jira Summary.",
    "doc_title":   "Heading shown at the top of the Word document. Auto-synced from Jira Summary; edit to override.",
    "summary":     "Jira ticket title — becomes the Issue Summary field in Jira.",
    "priority":    "Priority level shown in Jira (Highest → Lowest).",
    "component":   "Application component or module affected by this issue.",
    "fix_version": "Release version in which this issue is fixed.",
}

# ─────────────────────────────────────────────────────────────────────────────
#  Tooltip helper
# ─────────────────────────────────────────────────────────────────────────────

class Tooltip:
    """Lightweight hover tooltip attached to any widget."""
    def __init__(self, widget, text):
        self._w    = widget
        self._text = text
        self._tip  = None
        widget.bind("<Enter>",       self._show, "+")
        widget.bind("<Leave>",       self._hide, "+")
        widget.bind("<ButtonPress>", self._hide, "+")

    def _show(self, _=None):
        if self._tip or not self._text:
            return
        wx, wy = self._w.winfo_rootx(), self._w.winfo_rooty()
        wh     = self._w.winfo_height()
        self._tip = tk.Toplevel(self._w)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{wx + 16}+{wy + wh + 6}")
        tk.Label(self._tip, text=self._text,
                 bg="#24292f", fg="#e6edf3",
                 font=("Segoe UI", 9), relief="flat",
                 padx=10, pady=7, wraplength=300,
                 justify="left").pack()

    def _hide(self, _=None):
        if self._tip:
            self._tip.destroy()
            self._tip = None

# ─────────────────────────────────────────────────────────────────────────────
#  Dependency-order scoring for PR files
# ─────────────────────────────────────────────────────────────────────────────

def _dep_score(filename):
    """Lower score = fewer dependencies = should appear first in the step list."""
    fname = filename.lower()
    base  = os.path.basename(fname)
    name  = os.path.splitext(base)[0]

    # Test files first — a file named test_auth_service.py is a test, not a service
    if any(x in fname for x in ['test', 'spec', '__tests__', '.test.', '.spec.']):
        return 5
    if any(x in name for x in ['config', 'const', 'constant', 'setting', 'settings',
                                 'setup', 'init', 'env', 'properties', 'defaults']):
        return 0
    if any(x in name for x in ['type', 'types', 'model', 'models', 'schema', 'schemas',
                                 'interface', 'interfaces', 'enum', 'enums', 'dto', 'entity']):
        return 1
    if any(x in name for x in ['util', 'utils', 'helper', 'helpers', 'common',
                                 'shared', 'base', 'mixin', 'lib', 'libs', 'core']):
        return 2
    if any(x in name for x in ['service', 'services', 'api', 'client', 'repo',
                                 'repository', 'store', 'stores', 'hook', 'hooks', 'dao']):
        return 3
    if any(x in name for x in ['component', 'controller', 'view', 'page',
                                 'screen', 'widget', 'handler', 'middleware', 'route']):
        return 4
    return 3   # default: mid-tier

# ─────────────────────────────────────────────────────────────────────────────
#  Built-in field definitions  (label, key, type, choices, default, jira_key)
# ─────────────────────────────────────────────────────────────────────────────
BUILTIN_FIELDS = [
    # ── Jira ticket fields (sent directly to Jira API, shown in "Jira Fields" card) ──
    {"label": "Project",      "key": "project",     "type": "text",
     "choices": [], "default": "", "jira_key": "project",    "jira_field": True,
     "required": True,  "enabled": True, "show_label_in_jira": False},
    {"label": "Jira Summary", "key": "summary",     "type": "text",
     "choices": [], "default": "", "jira_key": "summary",    "jira_field": True,
     "required": True,  "enabled": True, "show_label_in_jira": False},
    {"label": "Issue Type",   "key": "issue_type",  "type": "dropdown",
     "choices": ["Bug","Enhancement","Task","Story","Epic","Sub-task","Incident"],
     "default": "Bug", "jira_key": "issuetype",              "jira_field": True,
     "required": True,  "enabled": True, "show_label_in_jira": False},
    {"label": "Reporter",     "key": "reporter",    "type": "text",
     "choices": [], "default": "", "jira_key": "reporter",   "jira_field": True,
     "required": False, "enabled": True, "show_label_in_jira": False},
    # ── Description-only fields (appear in Jira ticket body, not as top-level fields) ──
    {"label": "Dependency",   "key": "dependency",  "type": "text",
     "choices": [], "default": "", "jira_key": "dependency",
     "required": False, "enabled": True, "show_label_in_jira": True},
    {"label": "Git PR Link",  "key": "git_link",    "type": "text",
     "choices": [], "default": "", "jira_key": "git_link",
     "required": False, "enabled": True, "show_label_in_jira": True},
    {"label": "Doc File Name", "key": "doc_name",    "type": "text",
     "choices": [], "default": "", "jira_key": None,
     "required": False, "enabled": True, "show_label_in_jira": False},
    {"label": "Doc Title",    "key": "doc_title",   "type": "text",
     "choices": [], "default": "", "jira_key": None,
     "required": False, "enabled": True, "show_label_in_jira": False},
    {"label": "Priority",     "key": "priority",    "type": "dropdown",
     "choices": ["Highest","High","Medium","Low","Lowest"],
     "default": "Medium", "jira_key": "priority",
     "required": False, "enabled": True, "show_label_in_jira": True},
    {"label": "Component",    "key": "component",   "type": "text",
     "choices": [], "default": "", "jira_key": "components",
     "required": False, "enabled": False, "show_label_in_jira": True},
    {"label": "Fix Version",  "key": "fix_version", "type": "text",
     "choices": [], "default": "", "jira_key": "fixVersions",
     "required": False, "enabled": False, "show_label_in_jira": True},
]

DEFAULT_CONFIG = {
    "github_token_file":   os.path.join(BASE_DIR, "github_token.txt"),
    "jira_token_file":     os.path.join(BASE_DIR, "jira_token.txt"),
    "jira_base_url":       "https://yourcompany.atlassian.net",
    "jira_project_key":    "PROJ",
    "jira_email":          "you@company.com",
    "github_owner":        "",
    "github_repo":         "",
    "word_doc_output_dir": BASE_DIR,
    "fields":              BUILTIN_FIELDS,
    "field_defaults":      {},   # key → default_value overrides
}

# ─────────────────────────────────────────────────────────────────────────────
#  Config helpers
# ─────────────────────────────────────────────────────────────────────────────

def load_config():
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE) as f:
            saved = json.load(f)
        # apply non-field settings (tokens, urls, etc.)
        for k, v in saved.items():
            if k != "fields":
                cfg[k] = v
        # merge saved per-field overrides (enabled, default) into BUILTIN_FIELDS
        saved_by_key = {fd["key"]: fd for fd in saved.get("fields", [])}
        for fd in cfg["fields"]:
            if fd["key"] in saved_by_key:
                sf = saved_by_key[fd["key"]]
                if "enabled" in sf:
                    fd["enabled"] = sf["enabled"]
                if "default" in sf:
                    fd["default"] = sf["default"]
        # also apply field_defaults overrides
        for fd in cfg["fields"]:
            if fd["key"] in cfg.get("field_defaults", {}):
                fd["default"] = cfg["field_defaults"][fd["key"]]
    return cfg

def save_config(cfg):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)

def read_token(path):
    path = os.path.expanduser(path)
    if not os.path.exists(path):
        raise FileNotFoundError(f"Token file not found:\n{path}")
    with open(path) as f:
        return f.read().strip()

def gh_headers(token):
    return {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}

def jira_auth_headers(email, token):
    creds = base64.b64encode(f"{email}:{token}".encode()).decode()
    return {"Authorization": f"Basic {creds}",
            "Content-Type": "application/json", "Accept": "application/json"}

# ─────────────────────────────────────────────────────────────────────────────
#  SQL helpers
# ─────────────────────────────────────────────────────────────────────────────

def _detect_sql_object(filename, patch):
    """Return a short label like 'TABLE users' or 'VIEW active_orders' from
    the patch/filename.  Used as the default step comment in the SQL popup."""
    # Gather added/changed lines from patch (first 60) for detection
    search = ""
    if patch:
        added = [l[1:] for l in patch.splitlines()
                 if l.startswith('+') and not l.startswith('+++')]
        search = "\n".join(added[:60])
    if not search:
        search = os.path.basename(filename)

    DDL = (r'(CREATE|ALTER|DROP)\s+(?:OR\s+REPLACE\s+)?(?:FORCE\s+)?'
           r'(?:EDITIONABLE\s+)?'
           r'(TABLE|VIEW|PROCEDURE|PROC|FUNCTION|TRIGGER|PACKAGE|SEQUENCE|'
           r'TYPE|INDEX|SCHEMA)\s+(?:\w+\.)?(\w+)')
    m = re.search(DDL, search, re.IGNORECASE | re.MULTILINE)
    if m:
        return f"{m.group(2).upper()} {m.group(3)}"

    DML = r'(?:INSERT\s+INTO|UPDATE|DELETE\s+FROM)\s+(?:\w+\.)?(\w+)'
    m = re.search(DML, search, re.IGNORECASE)
    if m:
        return m.group(1)

    base = os.path.splitext(os.path.basename(filename))[0]
    base = re.sub(r'^\d+[_\-]+', '', base)   # strip leading "01_"
    return base




# ─────────────────────────────────────────────────────────────────────────────
#  GitHub-style diff renderer for Word docs
# ─────────────────────────────────────────────────────────────────────────────

def _word_diff_block(doc, patch, status, adds, dels):
    """Print-friendly diff: light backgrounds, dark text readable on white paper."""
    # colour palette (light backgrounds, dark foreground text)
    BG_ADD  = "E6FFED"; FG_ADD  = RGBColor(0x1a, 0x7f, 0x37)  # light green / dark green
    BG_DEL  = "FFEBE9"; FG_DEL  = RGBColor(0xcf, 0x22, 0x2e)  # light red   / dark red
    BG_MOD  = "FFFBDD"; FG_MOD  = RGBColor(0x7c, 0x4a, 0x00)  # light amber / dark amber
    BG_HUNK = "DDF4FF"; FG_HUNK = RGBColor(0x09, 0x69, 0xda)  # light blue  / dark blue
    BG_CTX  = "F6F8FA"; FG_CTX  = RGBColor(0x57, 0x60, 0x6a)  # light grey  / dark grey
    BG_NONE = "F6F8FA"; FG_NONE = RGBColor(0x57, 0x60, 0x6a)

    def _sp(text, bg_hex, fg_rgb, bold=False):
        p = doc.add_paragraph(style="Normal")
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)
        p.paragraph_format.left_indent  = Cm(0.5)
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"),   "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"),  bg_hex)
        pPr.append(shd)
        r = p.add_run(text)
        r.font.name      = "Consolas"
        r.font.size      = Pt(8)
        r.font.color.rgb = fg_rgb
        r.bold           = bold
        return p

    if status in ("removed", "deleted"):
        _sp(f"  −  FILE DELETED  —  {dels} line{'s' if dels != 1 else ''} removed",
            BG_DEL, FG_DEL, bold=True)
        if patch:
            old_ln = 1
            for raw in patch.splitlines():
                if raw.startswith('@@'):
                    m2 = re.match(r'@@ -(\d+)(?:,\d+)? \+(\d+)(?:,\d+)? @@', raw)
                    if m2:
                        old_ln = int(m2.group(1))
                    _sp(f"  {raw}", BG_HUNK, FG_HUNK)
                elif raw.startswith('-') and not raw.startswith('---'):
                    _sp(f"  {old_ln:>4}  −  {raw[1:]}", BG_DEL, FG_DEL)
                    old_ln += 1
        return

    if not patch:
        _sp("     (Diff not available — file too large or binary)", BG_NONE, FG_NONE)
        return

    if status == "added":
        _sp(f"  +  NEW FILE  —  {adds} line{'s' if adds != 1 else ''} added",
            BG_ADD, FG_ADD, bold=True)
        new_ln = 1
        for raw in patch.splitlines():
            if raw.startswith('+') and not raw.startswith('+++'):
                _sp(f"  {new_ln:>4}  +  {raw[1:]}", BG_ADD, FG_ADD)
                new_ln += 1
        return

    # modified / renamed
    old_ln = new_ln = 0
    has_changes = False
    lines = patch.splitlines()
    i = 0
    while i < len(lines):
        raw = lines[i]
        if raw.startswith('@@'):
            m = re.match(r'@@ -(\d+)(?:,\d+)? \+(\d+)(?:,\d+)? @@(.*)', raw)
            if m:
                old_ln = int(m.group(1))
                new_ln = int(m.group(2))
                _sp(f"  {raw}", BG_HUNK, FG_HUNK)
            i += 1
        elif (raw.startswith('-') and not raw.startswith('---')) or \
             (raw.startswith('+') and not raw.startswith('+++')):
            block_dels, block_adds = [], []
            while i < len(lines) and lines[i].startswith('-') and not lines[i].startswith('---'):
                block_dels.append(lines[i][1:])
                old_ln += 1
                i += 1
            new_ln_start = new_ln
            while i < len(lines) and lines[i].startswith('+') and not lines[i].startswith('+++'):
                block_adds.append(lines[i][1:])
                new_ln += 1
                i += 1
            for j, add_text in enumerate(block_adds):
                ln = new_ln_start + j
                if j < len(block_dels):
                    _sp(f"  {ln:>4}  ~  {add_text}", BG_MOD, FG_MOD)
                else:
                    _sp(f"  {ln:>4}  +  {add_text}", BG_ADD, FG_ADD)
            has_changes = has_changes or bool(block_dels or block_adds)
        elif raw.startswith(' '):
            old_ln += 1; new_ln += 1
            i += 1
        else:
            i += 1

    if not has_changes:
        _sp("     (No changed lines detected)", BG_NONE, FG_NONE)


# ─────────────────────────────────────────────────────────────────────────────
#  Word document generator
# ─────────────────────────────────────────────────────────────────────────────

def generate_word_doc(field_values, field_defs, pr_data, file_comments, output_dir):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # ── title ──
    doc_title_val = (field_values.get("doc_title") or "").strip() or \
                    (pr_data.get("title") if pr_data else None) or "Issue Report"
    title = doc.add_heading(doc_title_val, 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].font.color.rgb = RGBColor(0x8b, 0x94, 0x9e)
    doc.add_paragraph("")

    # ── dynamic fields table ──
    tbl = doc.add_table(rows=0, cols=2)
    tbl.style = "Light List Accent 1"

    def add_row(label, value, bold_label=True):
        row = tbl.add_row()
        row.cells[0].text = label
        row.cells[1].text = str(value) if value else "N/A"
        if bold_label and row.cells[0].paragraphs[0].runs:
            row.cells[0].paragraphs[0].runs[0].bold = True

    enabled = [fd for fd in field_defs if fd.get("enabled", True)]
    for fd in enabled:
        val = field_values.get(fd["key"], fd.get("default", ""))
        add_row(fd["label"], val)

    # ── description ──
    doc.add_paragraph("")
    doc.add_heading("Description", level=1)
    doc.add_paragraph(field_values.get("description") or "No description provided.")

    # ── PR details ──
    if pr_data:
        doc.add_paragraph("")
        doc.add_heading("GitHub PR Details", level=1)
        pt = doc.add_table(rows=0, cols=2)
        pt.style = "Light List Accent 1"

        def pr_add(label, value):
            row = pt.add_row()
            row.cells[0].text = label
            row.cells[1].text = str(value) if value else "N/A"
            if row.cells[0].paragraphs[0].runs:
                row.cells[0].paragraphs[0].runs[0].bold = True

        pr_add("PR Title", pr_data.get("title"))
        pr_add("PR State", pr_data.get("state", "").upper())
        pr_add("Author",   pr_data.get("user", {}).get("login"))
        pr_add("Created",  (pr_data.get("created_at") or "")[:10])
        pr_add("Updated",  (pr_data.get("updated_at") or "")[:10])
        pr_add("Branch",   pr_data.get("head", {}).get("ref"))
        pr_add("Base",     pr_data.get("base", {}).get("ref"))
        if pr_data.get("body"):
            doc.add_paragraph("")
            doc.add_heading("PR Body", level=2)
            doc.add_paragraph(pr_data["body"])

    # ── file changes ──
    if file_comments:
        doc.add_paragraph("")
        doc.add_heading("Code Changes", level=1)
        for idx, fc in enumerate(file_comments, 1):
            fname   = fc.get("filename", "unknown")
            status  = fc.get("status", "modified")
            comment = (fc.get("word_comment") or fc.get("comment", "")).strip()
            shots   = fc.get("screenshots", [])
            adds    = fc.get("additions", 0)
            dels    = fc.get("deletions", 0)
            patch   = fc.get("patch", "")

            doc.add_paragraph("")

            # ── file header ──────────────────────────────────────────────────
            fh = doc.add_paragraph()
            fh.paragraph_format.space_before = Pt(4)
            fh.paragraph_format.space_after  = Pt(2)
            r1 = fh.add_run(f"Step {idx}  ")
            r1.bold = True; r1.font.size = Pt(11)
            r1.font.color.rgb = RGBColor(0x00, 0x5c, 0xb8)   # dark blue
            badge_lbl = STATUS_EMOJI.get(status, "~")
            r2 = fh.add_run(f"[{badge_lbl} {status.upper()}]  ")
            r2.bold = True; r2.font.size = Pt(10)
            if status == "added":
                r2.font.color.rgb = RGBColor(0x1a, 0x7f, 0x37)   # dark green
            elif status in ("removed", "deleted"):
                r2.font.color.rgb = RGBColor(0xcf, 0x22, 0x2e)   # dark red
            else:
                r2.font.color.rgb = RGBColor(0x7c, 0x4a, 0x00)   # dark amber
            r3 = fh.add_run(fname)
            r3.bold = True; r3.font.size = Pt(10)
            r3.font.name = "Consolas"
            r3.font.color.rgb = RGBColor(0x1f, 0x2d, 0x3d)       # dark navy

            if adds or dels:
                sp = doc.add_paragraph()
                sp.paragraph_format.left_indent = Cm(0.5)
                sp.paragraph_format.space_before = Pt(0)
                sp.paragraph_format.space_after  = Pt(4)
                ra = sp.add_run(f"+{adds} additions  ")
                ra.font.size = Pt(9); ra.font.color.rgb = RGBColor(0x1a, 0x7f, 0x37)
                rd = sp.add_run(f"−{dels} deletions")
                rd.font.size = Pt(9); rd.font.color.rgb = RGBColor(0xcf, 0x22, 0x2e)

            # ── change description ────────────────────────────────────────────
            if comment:
                ch = doc.add_paragraph()
                ch.paragraph_format.left_indent   = Cm(0.5)
                ch.paragraph_format.right_indent  = Cm(0.5)
                ch.paragraph_format.space_before  = Pt(6)
                ch.paragraph_format.space_after   = Pt(6)
                # light yellow shading so it stands out on the page
                pPr = ch._p.get_or_add_pPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"),   "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"),  "FFF8DC")
                pPr.append(shd)
                rl = ch.add_run("Change Description:  ")
                rl.bold = True; rl.font.size = Pt(10)
                rl.font.color.rgb = RGBColor(0x7c, 0x4a, 0x00)   # dark amber label
                rc = ch.add_run(comment)
                rc.font.size = Pt(11)
                rc.bold = False
                rc.font.color.rgb = RGBColor(0x1a, 0x1a, 0x1a)   # near-black — readable

            # ── diff block label ──────────────────────────────────────────────
            doc.add_paragraph("")
            dh = doc.add_paragraph()
            dh.paragraph_format.left_indent = Cm(0.5)
            dh.paragraph_format.space_before = Pt(0)
            dh.paragraph_format.space_after  = Pt(2)
            rh = dh.add_run("Diff:")
            rh.bold = True; rh.font.size = Pt(9)
            rh.font.color.rgb = RGBColor(0x57, 0x60, 0x6a)       # dark grey

            _word_diff_block(doc, patch, status, adds, dels)

            for si, shot in enumerate(shots, 1):
                if isinstance(shot, dict):
                    sp_path     = shot.get("path", "")
                    shot_cmt    = shot.get("comment", "")
                else:
                    sp_path  = str(shot)
                    shot_cmt = ""
                if not sp_path or not os.path.exists(sp_path):
                    continue
                doc.add_paragraph("")
                cap_lbl = shot_cmt or f"Screenshot {si}  —  {os.path.basename(sp_path)}"
                cap = doc.add_paragraph(f"  {cap_lbl}")
                if cap.runs:
                    cap.runs[0].font.size = Pt(9)
                    cap.runs[0].font.color.rgb = RGBColor(0x8b, 0x94, 0x9e)
                try:
                    doc.add_picture(sp_path, width=Inches(5.5))
                    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                except Exception as e:
                    doc.add_paragraph(f"  [Cannot embed image: {e}]")

            # divider rule
            div = doc.add_paragraph("")
            pPr = div._p.get_or_add_pPr()
            pBdr = OxmlElement("w:pBdr")
            bot  = OxmlElement("w:bottom")
            bot.set(qn("w:val"), "single")
            bot.set(qn("w:sz"), "4")
            bot.set(qn("w:space"), "1")
            bot.set(qn("w:color"), "BFBFBF")
            pBdr.append(bot)
            pPr.append(pBdr)

    doc_name = field_values.get("doc_name") or \
               f"Issue_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    if not doc_name.endswith(".docx"):
        doc_name += ".docx"
    pr_num = str(pr_data.get("number", "")) if pr_data else ""
    save_dir = os.path.join(output_dir, f"PR_{pr_num}") if pr_num else output_dir
    os.makedirs(save_dir, exist_ok=True)
    out = os.path.join(save_dir, doc_name)
    doc.save(out)
    return out

# ─────────────────────────────────────────────────────────────────────────────
#  HTML preview builder  (rendered in a tk.Text widget with embedded images)
# ─────────────────────────────────────────────────────────────────────────────

def build_preview_html(field_values, field_defs, pr_data, file_comments):
    """Return an HTML string that visually matches the Word doc layout."""
    e = html.escape

    def trow(label, value):
        return (f'<tr><td style="font-weight:700;padding:4px 12px 4px 4px;'
                f'color:#e6edf3;white-space:nowrap">{e(label)}</td>'
                f'<td style="padding:4px;color:#adbac7">{e(str(value) if value else "N/A")}'
                f'</td></tr>')

    css = """
    body{background:#0d1117;color:#e6edf3;font-family:Calibri,Segoe UI,sans-serif;
         font-size:14px;margin:0;padding:24px}
    h1{color:#58a6ff;font-size:20px;border-bottom:1px solid #30363d;padding-bottom:6px}
    h2{color:#8b949e;font-size:15px;margin-top:18px}
    h0{color:#e6edf3;font-size:24px;text-align:center;display:block;margin-bottom:4px}
    .sub{text-align:center;color:#8b949e;font-size:12px;margin-bottom:20px}
    table{border-collapse:collapse;width:100%;max-width:700px;margin-bottom:16px}
    td{vertical-align:top}
    .section{background:#1c2128;border:1px solid #30363d;border-radius:6px;
             padding:16px;margin-bottom:18px}
    .file-card{border-left:3px solid #58a6ff;margin:12px 0;padding:10px 14px;
               background:#161b22;border-radius:0 6px 6px 0}
    .file-added{border-left-color:#3fb950}
    .file-modified{border-left-color:#58a6ff}
    .file-removed{border-left-color:#f85149}
    .file-renamed{border-left-color:#bc8cff}
    .badge{display:inline-block;padding:2px 8px;border-radius:4px;
           font-size:11px;font-weight:700;margin-right:8px}
    .badge-A{background:#1f6931;color:#fff}
    .badge-M{background:#1c4a6e;color:#fff}
    .badge-D{background:#8a1f1f;color:#fff}
    .badge-R{background:#5a2d82;color:#fff}
    .badge-C{background:#1c4a6e;color:#fff}
    .fname{font-family:Consolas,monospace;font-size:13px;font-weight:700}
    .stats{font-size:11px;color:#8b949e;margin-top:4px}
    .adds{color:#3fb950} .dels{color:#f85149}
    .comment-box{background:#0d1117;border-left:3px solid #d29922;
                 border-radius:0 4px 4px 0;padding:8px 12px;
                 margin-top:10px;font-size:13px;color:#e6edf3}
    .comment-label{font-size:11px;color:#d29922;font-weight:700;
                   margin-bottom:4px;text-transform:uppercase;letter-spacing:.05em}
    .screenshot-wrap{margin-top:12px;text-align:center}
    .screenshot-cap{font-size:11px;color:#8b949e;margin-bottom:6px}
    img{max-width:100%;border-radius:4px;border:1px solid #30363d}
    .divider{border:none;border-top:1px solid #30363d;margin:16px 0}
    .pr-body{white-space:pre-wrap;font-size:12px;color:#adbac7;
             background:#161b22;padding:10px;border-radius:4px}
    .diff-block{margin-top:10px;background:#0a0e14;border-radius:4px;
                border:1px solid #30363d}
    .diff-block summary{padding:6px 10px;cursor:pointer;color:#8b949e;
                        font-size:11px;user-select:none}
    .diff-content{font-family:Consolas,monospace;font-size:11px;
                  padding:8px;overflow-x:auto;max-height:320px;overflow-y:auto}
    .diff-add{background:#0e2318;color:#3fb950;white-space:pre}
    .diff-del{background:#200e0e;color:#f85149;white-space:pre}
    .diff-hunk{background:#0c1c2c;color:#58a6ff;white-space:pre}
    .diff-ctx{color:#6e7681;white-space:pre}
    """

    body = [f"<style>{css}</style>"]
    body.append(f'<p class="sub">Generated: {datetime.now().strftime("%Y-%m-%d  %H:%M:%S")}</p>')

    # dynamic fields
    body.append('<div class="section"><h1>Issue Details</h1>')
    body.append('<table>')
    enabled = [fd for fd in field_defs if fd.get("enabled", True)]
    for fd in enabled:
        val = field_values.get(fd["key"], fd.get("default", ""))
        body.append(trow(fd["label"], val))
    body.append('</table></div>')

    # description
    desc = field_values.get("description") or "No description provided."
    body.append(f'<div class="section"><h1>Description</h1>'
                f'<p style="white-space:pre-wrap">{e(desc)}</p></div>')

    # PR details
    if pr_data:
        body.append('<div class="section"><h1>GitHub PR Details</h1><table>')
        body.append(trow("PR Title", pr_data.get("title")))
        body.append(trow("PR State", (pr_data.get("state") or "").upper()))
        body.append(trow("Author",   pr_data.get("user", {}).get("login")))
        body.append(trow("Created",  (pr_data.get("created_at") or "")[:10]))
        body.append(trow("Updated",  (pr_data.get("updated_at") or "")[:10]))
        body.append(trow("Branch",   pr_data.get("head", {}).get("ref")))
        body.append(trow("Base",     pr_data.get("base", {}).get("ref")))
        body.append('</table>')
        if pr_data.get("body"):
            body.append(f'<h2>PR Body</h2>'
                        f'<div class="pr-body">{e(pr_data["body"])}</div>')
        body.append('</div>')

    # file changes
    if file_comments:
        body.append('<div class="section"><h1>Code Changes</h1>')
        for idx, fc in enumerate(file_comments, 1):
            fname        = fc.get("filename", "unknown")
            status       = fc.get("status", "modified")
            word_comment = (fc.get("word_comment") or fc.get("comment", "")).strip()
            jira_comment = fc.get("jira_comment", "").strip()
            shots        = fc.get("screenshots", [])
            adds         = fc.get("additions", 0)
            dels         = fc.get("deletions", 0)
            patch        = fc.get("patch", "")
            badge        = STATUS_EMOJI.get(status, "~")

            body.append(f'<div class="file-card file-{status}">')
            body.append(f'<span class="badge badge-{badge}">{badge} {status.upper()}</span>'
                        f'<span class="fname">{e(fname)}</span>')
            body.append(f'<div class="stats">'
                        f'<span class="adds">+{adds}</span>&nbsp;&nbsp;'
                        f'<span class="dels">-{dels}</span></div>')

            # NEW FILE banner
            if status == "added":
                body.append(f'<div style="background:#0e2318;border-radius:4px;'
                            f'padding:6px 12px;margin-top:8px;font-family:Consolas,monospace;'
                            f'font-size:12px;color:#3fb950;font-weight:700">'
                            f'+ NEW FILE &nbsp;—&nbsp; {adds} line{"s" if adds != 1 else ""} added'
                            f'</div>')

            if word_comment:
                body.append(f'<div class="comment-box">'
                            f'<div class="comment-label">Change Description</div>'
                            f'{e(word_comment)}</div>')
            if jira_comment and jira_comment != word_comment:
                body.append(f'<div class="comment-box" style="border-left-color:#58a6ff">'
                            f'<div class="comment-label" style="color:#58a6ff">Jira Comment</div>'
                            f'{e(jira_comment)}</div>')

            # GitHub-style diff block
            if patch and status != "added":
                patch_lines = patch.splitlines()
                shown = patch_lines[:120]
                diff_rows = []
                for dl in shown:
                    if dl.startswith('+') and not dl.startswith('+++'):
                        cls, sym = "diff-add", "+"
                        diff_rows.append(
                            f'<div class="{cls}"><span style="opacity:.6;margin-right:6px">{sym}</span>{e(dl[1:])}</div>')
                    elif dl.startswith('-') and not dl.startswith('---'):
                        cls, sym = "diff-del", "−"
                        diff_rows.append(
                            f'<div class="{cls}"><span style="opacity:.6;margin-right:6px">{sym}</span>{e(dl[1:])}</div>')
                    elif dl.startswith('@'):
                        diff_rows.append(
                            f'<div class="diff-hunk">{e(dl)}</div>')
                    else:
                        diff_rows.append(
                            f'<div class="diff-ctx"><span style="opacity:.4;margin-right:6px">&nbsp;</span>{e(dl[1:] if dl and dl[0]==" " else dl)}</div>')
                if len(patch_lines) > 120:
                    diff_rows.append(
                        f'<div class="diff-ctx" style="color:#d29922">'
                        f'⋯  {len(patch_lines)-120} more lines not shown</div>')
                body.append(
                    f'<details class="diff-block" open>'
                    f'<summary>Code Changes &nbsp;'
                    f'<span style="color:#3fb950">+{adds}</span>'
                    f'&nbsp;<span style="color:#f85149">−{dels}</span>'
                    f'</summary>'
                    f'<div class="diff-content">{"".join(diff_rows)}</div>'
                    f'</details>')

            for si, shot in enumerate(shots, 1):
                if isinstance(shot, dict):
                    sp_path  = shot.get("path", "")
                    shot_cmt = shot.get("comment", "")
                else:
                    sp_path  = str(shot)
                    shot_cmt = ""
                if not sp_path or not os.path.exists(sp_path):
                    continue
                import base64 as b64
                try:
                    ext = os.path.splitext(sp_path)[1].lower().lstrip(".")
                    mime = {"jpg":"jpeg","jpeg":"jpeg","png":"png",
                            "gif":"gif","webp":"webp","bmp":"bmp"}.get(ext, "png")
                    with open(sp_path, "rb") as fh:
                        data = b64.b64encode(fh.read()).decode()
                    cap_html = (f'<div class="screenshot-cap">{e(shot_cmt)}</div>'
                                if shot_cmt else
                                f'<div class="screenshot-cap">'
                                f'Screenshot {si} — {e(os.path.basename(sp_path))}'
                                f'</div>')
                    body.append(f'<div class="screenshot-wrap">'
                                f'{cap_html}'
                                f'<img src="data:image/{mime};base64,{data}" />'
                                f'</div>')
                except Exception:
                    body.append(f'<div class="stats">[Cannot load screenshot {si}]</div>')

            body.append('</div><hr class="divider">')
        body.append('</div>')

    return "<html><body>" + "".join(body) + "</body></html>"


# ─────────────────────────────────────────────────────────────────────────────
#  DocPreviewWindow
# ─────────────────────────────────────────────────────────────────────────────

class DocPreviewWindow(tk.Toplevel):
    """Preview window with left comment panel (per-file) and right HTML preview."""

    def __init__(self, parent, html_content, doc_path=None,
                 pr_files=None, file_comments=None, on_comments_saved=None,
                 build_html_func=None):
        super().__init__(parent)
        self.title(f"{APP_NAME}  —  Code Changes")
        self.geometry("1300x820")
        self.minsize(900, 560)
        self.configure(bg=C["bg"])
        self._doc_path         = doc_path
        self._html_content     = html_content
        self._pr_files         = pr_files or []
        self._on_save          = on_comments_saved
        self._build_html_func  = build_html_func
        self._html_frame       = None
        self._comment_rows     = []   # {filename, word_var, jira_var, status, ...}

        # pre-index existing comments by filename
        self._existing = {fc["filename"]: fc for fc in (file_comments or [])}

        self._build()

    # ── build ─────────────────────────────────────────────────────────────────

    def _build(self):
        # ── toolbar ───────────────────────────────────────────────────────────
        self._bar = tk.Frame(self, bg=C["surface"])
        self._bar.pack(fill="x")
        tk.Label(self._bar, text="Code Changes — Add Comments",
                 bg=C["surface"], fg=C["accent"],
                 font=("Segoe UI", 12, "bold")).pack(side="left", padx=16, pady=10)

        # Open in Word button — enabled only after doc is generated
        doc_exists = bool(self._doc_path and os.path.exists(self._doc_path))
        self._open_word_btn = tk.Button(
            self._bar, text="  Open in Word  ",
            bg=C["green"] if doc_exists else C["surface"],
            fg="#0d1117" if doc_exists else C["muted"],
            relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=3,
            activebackground="#2ea043", cursor="hand2" if doc_exists else "arrow",
            state="normal" if doc_exists else "disabled",
            command=self._open_word)
        self._open_word_btn.pack(side="right", padx=8, pady=8)

        if self._pr_files and self._on_save:
            self._status_lbl = tk.Label(self._bar, text="",
                                        bg=C["surface"], fg=C["green"],
                                        font=("Segoe UI", 9))
            self._status_lbl.pack(side="right", padx=8)
            tk.Button(self._bar, text="  Generate Doc  ",
                      bg=C["green"], fg="#0d1117", relief="flat",
                      font=("Segoe UI", 9, "bold"), padx=10, pady=3,
                      activebackground="#2ea043", cursor="hand2",
                      command=self._generate_doc).pack(side="right", padx=4, pady=8)

        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")

        # ── body: full-screen comment panel ───────────────────────────────────
        body = tk.Frame(self, bg=C["bg"])
        body.pack(fill="both", expand=True)

        if self._pr_files:
            self._build_comment_panel(body)

    def _build_comment_panel(self, parent):
        panel = tk.Frame(parent, bg=C["bg"])
        panel.pack(fill="both", expand=True)

        # header
        ph = tk.Frame(panel, bg=C["surface"])
        ph.pack(fill="x")
        tk.Frame(ph, bg=C["yellow"], width=4).pack(side="left", fill="y")
        tk.Label(ph, text=f"  Code Changes  ({len(self._pr_files)} files)",
                 bg=C["surface"], fg=C["yellow"],
                 font=("Segoe UI", 10, "bold")).pack(side="left", padx=6, pady=8)
        tk.Label(ph, text="  Add change descriptions for Word doc",
                 bg=C["surface"], fg=C["muted"],
                 font=("Segoe UI", 8)).pack(side="left")

        tk.Frame(panel, bg=C["border"], height=1).pack(fill="x")

        # scrollable file cards
        wrap = tk.Frame(panel, bg=C["bg"]); wrap.pack(fill="both", expand=True)
        canvas = tk.Canvas(wrap, bg=C["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        lf = tk.Frame(canvas, bg=C["bg"])
        wid = canvas.create_window((0, 0), window=lf, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(wid, width=e.width))
        lf.bind("<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        def _mw(e): canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        canvas.bind("<MouseWheel>", _mw)

        sorted_files = sorted(self._pr_files,
                              key=lambda f: (_dep_score(f.get("filename", "")),
                                             f.get("filename", "")))

        for idx, f in enumerate(sorted_files, 1):
            fname  = f.get("filename", "")
            status = f.get("status", "modified")
            adds   = f.get("additions", 0)
            dels   = f.get("deletions", 0)
            rbg    = STATUS_ROW.get(status, C["card"])
            bbg    = STATUS_BADGE.get(status, C["surface"])
            em     = STATUS_EMOJI.get(status, "~")
            ex     = self._existing.get(fname, {})

            card = tk.Frame(lf, bg=rbg,
                            highlightbackground=C["border"], highlightthickness=1)
            card.pack(fill="x", padx=6, pady=3)
            card.bind("<MouseWheel>", _mw)

            # file header
            fhdr = tk.Frame(card, bg=rbg); fhdr.pack(fill="x", padx=6, pady=(6, 2))
            fhdr.bind("<MouseWheel>", _mw)
            tk.Label(fhdr, text=f"Step {idx}",
                     bg=C["accent"], fg="#fff",
                     font=("Segoe UI", 8, "bold"), padx=4, pady=1
                     ).pack(side="left", padx=(0, 4))
            tk.Label(fhdr, text=f" {em} {status.upper()} ",
                     bg=bbg, fg="#fff",
                     font=("Segoe UI", 8, "bold"), padx=3, pady=1
                     ).pack(side="left", padx=(0, 6))
            sfg = C["green"] if adds > dels else C["red"] if dels > adds else C["muted"]
            tk.Label(fhdr, text=f"+{adds} −{dels}",
                     bg=rbg, fg=sfg,
                     font=("Segoe UI", 8, "bold")).pack(side="right")

            # filename + inline Show Code Change button
            parts = fname.split("/")
            pre   = "/".join(parts[:-2]) + "/" if len(parts) > 2 else ""
            tail  = "/".join(parts[-2:])
            nf    = tk.Frame(card, bg=rbg); nf.pack(fill="x", padx=6, pady=(0, 4))
            nf.bind("<MouseWheel>", _mw)
            if pre:
                tk.Label(nf, text=pre, bg=rbg, fg=C["muted"],
                         font=("Consolas", 8)).pack(side="left")
            tk.Label(nf, text=tail, bg=rbg, fg=C["text"],
                     font=("Consolas", 9, "bold")).pack(side="left")
            patch_str = f.get("patch", "")
            if patch_str:
                n_lines = len(patch_str.splitlines())
                tog_btn = tk.Button(
                    nf,
                    text=f"⊞ Code Change ({n_lines})",
                    bg=C["accent"], fg="#ffffff", relief="flat",
                    font=("Segoe UI", 8, "bold"),
                    padx=6, pady=1, cursor="hand2",
                    activebackground=C["hover"], activeforeground="#ffffff",
                    command=lambda fn=fname, p=patch_str, st=status, a=adds, d=dels:
                        DiffViewPopup(self, fn, p, st, a, d))
                tog_btn.pack(side="left", padx=(8, 0))
                tog_btn.bind("<MouseWheel>", _mw)

            # Change Description
            tk.Label(card, text="Change Description:",
                     bg=rbg, fg=C["yellow"],
                     font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=6)
            word_var = tk.StringVar(value=ex.get("word_comment", ex.get("comment", "")))
            word_ent = scrolledtext.ScrolledText(
                card, height=3, bg=C["inp"], fg=C["text"],
                font=("Segoe UI", 10), relief="flat",
                insertbackground=C["text"], wrap="word")
            word_ent.insert("1.0", word_var.get())
            word_ent.pack(fill="x", padx=6, pady=(2, 8))
            word_ent.bind("<MouseWheel>", _mw)

            self._comment_rows.append(dict(
                filename=fname, status=status,
                additions=adds, deletions=dels,
                patch=f.get("patch", ""),
                screenshots=ex.get("screenshots", []),
                word_ent=word_ent,
            ))

    def _build_preview(self, parent):
        pf = tk.Frame(parent, bg=C["bg"])
        pf.pack(side="left", fill="both", expand=True)
        try:
            from tkinterweb import HtmlFrame
            frame = HtmlFrame(pf, horizontal_scrollbar="auto")
            frame.pack(fill="both", expand=True)
            frame.load_html(self._html_content)
            self._html_frame = frame
        except ImportError:
            self._fallback_viewer(pf, self._html_content)

    def _fallback_viewer(self, parent, html_content):
        text = re.sub(r'<style[^>]*>.*?</style>', '', html_content, flags=re.DOTALL)
        text = re.sub(r'<[^>]+>', '', text)
        text = text.replace("&amp;", "&").replace("&lt;", "<").replace(
            "&gt;", ">").replace("&#39;", "'").replace("&quot;", '"')
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        text  = "\n".join(lines)
        tk.Label(parent,
                 text="Install  pip install tkinterweb  for full HTML preview  |  "
                      "Showing plain-text fallback",
                 bg=C["yellow"], fg="#0d1117",
                 font=("Segoe UI", 9), pady=4).pack(fill="x")
        st = scrolledtext.ScrolledText(parent, bg=C["card"], fg=C["text"],
                                       font=("Segoe UI", 10), relief="flat",
                                       wrap="word")
        st.pack(fill="both", expand=True, padx=4, pady=4)
        st.insert("1.0", text)
        st.config(state="disabled")

    # ── actions ───────────────────────────────────────────────────────────────

    def _generate_doc(self):
        result = []
        for r in self._comment_rows:
            wc = r["word_ent"].get("1.0", "end-1c").strip()
            result.append(dict(
                filename=r["filename"], status=r["status"],
                additions=r["additions"], deletions=r["deletions"],
                patch=r["patch"], screenshots=r["screenshots"],
                comment=wc, word_comment=wc, jira_comment="",
                include=True,
            ))
        n_c = sum(1 for r in result if r["comment"])
        if hasattr(self, "_status_lbl"):
            self._status_lbl.config(text=f"Generating…  {len(result)} files, {n_c} comments")
        self.update_idletasks()
        if self._on_save:
            self._on_save(result)

    def _open_word(self):
        if self._doc_path and os.path.exists(self._doc_path):
            import subprocess, sys
            if sys.platform == "win32":
                os.startfile(self._doc_path)
            elif sys.platform == "darwin":
                subprocess.run(["open", self._doc_path])
            else:
                subprocess.run(["xdg-open", self._doc_path])

    def _enable_open_word(self, path):
        self._doc_path = path
        self._open_word_btn.config(
            bg=C["green"], fg="#0d1117",
            activebackground="#2ea043", cursor="hand2",
            state="normal")


# ─────────────────────────────────────────────────────────────────────────────
#  FieldManagerDialog  – add / remove / reorder / set defaults for Jira fields
# ─────────────────────────────────────────────────────────────────────────────

class FieldManagerDialog(tk.Toplevel):
    """
    Shows all field definitions.  User can:
      • toggle enabled
      • change default value
      • move up / down
      • delete custom fields
      • add new custom fields
    Returns via .result  (list of field dicts) after apply.
    """

    # Column spec: (header_label, min_pixel_width, tooltip)
    _COLS = [
        ("",            28,  ""),
        ("On",          34,  "Enable / disable field"),
        ("Field Label", 155, "Name shown in UI and Word doc"),
        ("Type",        68,  "text / dropdown / number / date"),
        ("Default",     175, "Pre-filled value when creating a ticket"),
        ("Choices",     64,  "Edit the dropdown options list"),
        ("Jira API",    56,  "Map this field directly to a Jira API field"),
        ("Show Label",  72,  "Prefix 'FieldName: value' in Jira description"),
        ("",            92,  ""),
    ]

    def __init__(self, parent, fields):
        super().__init__(parent)
        self.title(f"{APP_NAME}  —  Manage Fields")
        self.geometry("980x660")
        self.minsize(820, 480)
        self.configure(bg=C["bg"])
        self.grab_set()
        self.transient(parent)

        self._fields     = copy.deepcopy(fields)
        self.result      = None
        self._row_frames = []

        self._build()

    # ── build ─────────────────────────────────────────────────────────────────

    def _build(self):
        # ── title bar ────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=C["surface"])
        hdr.pack(fill="x")
        tk.Frame(hdr, bg=C["purple"], width=4).pack(side="left", fill="y")
        ti = tk.Frame(hdr, bg=C["surface"]); ti.pack(side="left", padx=14, pady=10)
        tk.Label(ti, text="Manage Fields",
                 bg=C["surface"], fg=C["text"],
                 font=("Segoe UI", 12, "bold")).pack(anchor="w")
        tk.Label(ti, text="Enable / disable fields  •  set defaults  •  control Jira mapping",
                 bg=C["surface"], fg=C["muted"],
                 font=("Segoe UI", 8)).pack(anchor="w")
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")

        # ── column header (grid-based, matches _render_rows columns exactly) ──
        col_hdr = tk.Frame(self, bg=C["card"])
        col_hdr.pack(fill="x")
        # left indent strip to visually align with row content
        tk.Frame(col_hdr, bg=C["card"], width=2).pack(side="left", fill="y")
        hdr_inner = tk.Frame(col_hdr, bg=C["card"])
        hdr_inner.pack(side="left", fill="x", expand=True)
        for c, (txt, minw, tip) in enumerate(self._COLS):
            hdr_inner.columnconfigure(c, minsize=minw, weight=1 if c == 2 else 0)
        for c, (txt, minw, tip) in enumerate(self._COLS):
            cell = tk.Label(hdr_inner, text=txt, bg=C["card"], fg=C["accent"],
                            font=("Segoe UI", 8, "bold"), anchor="w", padx=6)
            cell.grid(row=0, column=c, sticky="ew", ipady=6)
            if tip:
                Tooltip(cell, tip)
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")

        # ── scrollable list ───────────────────────────────────────────────────
        outer = tk.Frame(self, bg=C["bg"])
        outer.pack(fill="both", expand=True)
        self._canvas = tk.Canvas(outer, bg=C["bg"], highlightthickness=0)
        vbar = ttk.Scrollbar(outer, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vbar.set)
        vbar.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self._list_frame = tk.Frame(self._canvas, bg=C["bg"])
        self._win = self._canvas.create_window((0, 0), window=self._list_frame, anchor="nw")
        self._canvas.bind("<Configure>",
                          lambda e: self._canvas.itemconfig(self._win, width=e.width))
        self._list_frame.bind("<Configure>",
                              lambda e: self._canvas.configure(
                                  scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<MouseWheel>",
                          lambda e: self._canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        self._render_rows()

        # ── bottom action bar ─────────────────────────────────────────────────
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")
        bot = tk.Frame(self, bg=C["surface"])
        bot.pack(fill="x", side="bottom", pady=0)
        inner_bot = tk.Frame(bot, bg=C["surface"])
        inner_bot.pack(fill="x", padx=14, pady=10)

        tk.Button(inner_bot, text="+ Add Custom Field",
                  bg=C["accent"], fg="#ffffff", relief="flat",
                  font=("Segoe UI", 9, "bold"), padx=12, pady=5,
                  activebackground=C["hover"], cursor="hand2",
                  command=self._add_field).pack(side="left")

        tk.Button(inner_bot, text="Apply & Close",
                  bg=C["green"], fg="#0d1117", relief="flat",
                  font=("Segoe UI", 10, "bold"), padx=16, pady=5,
                  activebackground="#2ea043", cursor="hand2",
                  command=self._apply).pack(side="right")
        tk.Button(inner_bot, text="Cancel",
                  bg=C["red"], fg="#ffffff", relief="flat",
                  font=("Segoe UI", 10), padx=12, pady=5,
                  activebackground="#c0392b", cursor="hand2",
                  command=self.destroy).pack(side="right", padx=(0, 8))

    def _render_rows(self):
        for w in self._list_frame.winfo_children():
            w.destroy()
        self._row_frames = []

        # configure grid columns to match header exactly
        for c, (_, minw, _) in enumerate(self._COLS):
            self._list_frame.columnconfigure(c, minsize=minw, weight=1 if c == 2 else 0)

        type_clrs = {"text": C["muted"], "dropdown": C["accent"],
                     "number": C["yellow"], "date": C["green"]}

        for i, fd in enumerate(self._fields):
            r      = i * 2          # data on even rows, separator on odd
            bg     = C["card"] if i % 2 == 0 else C["surface"]
            enabled = fd.get("enabled", True)
            fg_txt = C["text"] if enabled else C["muted"]
            ftype  = fd.get("type", "text")

            def _cell(text="", col=0, fg=None, font=("Segoe UI", 9), anchor="w"):
                tk.Label(self._list_frame, text=text, bg=bg, fg=fg or fg_txt,
                         font=font, anchor=anchor, padx=6
                         ).grid(row=r, column=col, sticky="nsew", ipady=7)

            # col 0 — index
            _cell(str(i + 1), col=0, fg=C["muted"], font=("Segoe UI", 8))

            # col 1 — enabled checkbox
            en_var = tk.BooleanVar(value=enabled)
            tk.Checkbutton(self._list_frame, variable=en_var, bg=bg, fg=C["text"],
                           selectcolor=C["cb_sel"], activebackground=bg,
                           activeforeground=C["text"],
                           relief="flat", bd=0, cursor="hand2"
                           ).grid(row=r, column=1, sticky="ns", pady=4)

            # col 2 — field label
            lbl_var = tk.StringVar(value=fd["label"])
            if fd.get("custom"):
                e = ttk.Entry(self._list_frame, textvariable=lbl_var)
                e.grid(row=r, column=2, sticky="ew", padx=(6, 4), pady=5)
            else:
                hint = FIELD_HINTS.get(fd.get("key", ""), "")
                lw = tk.Label(self._list_frame, textvariable=lbl_var, bg=bg, fg=fg_txt,
                              anchor="w", padx=6,
                              font=("Segoe UI", 9, "bold" if fd.get("required") else "normal"))
                lw.grid(row=r, column=2, sticky="nsew", ipady=7)
                if hint:
                    Tooltip(lw, hint)

            # col 3 — type
            _cell(ftype, col=3, fg=type_clrs.get(ftype, C["muted"]),
                  font=("Segoe UI", 8, "italic"))

            # col 4 — default value
            def_var    = tk.StringVar(value=fd.get("default", ""))
            choices_list = fd.get("choices", [])
            choices_cb   = None
            if choices_list:
                choices_cb = ttk.Combobox(self._list_frame, textvariable=def_var,
                                          values=choices_list, state="normal")
                choices_cb.grid(row=r, column=4, sticky="ew", padx=(6, 4), pady=5)
            else:
                ttk.Entry(self._list_frame, textvariable=def_var).grid(
                    row=r, column=4, sticky="ew", padx=(6, 4), pady=5)

            # col 5 — edit choices (dropdown only)
            if ftype == "dropdown":
                def _open_edit(fd=fd, cb=choices_cb, dv=def_var):
                    self._edit_choices_dialog(fd, cb, dv)
                tk.Button(self._list_frame, text="✎", bg=bg, fg=C["accent"],
                          relief="flat", font=("Segoe UI", 10), padx=6, pady=0,
                          activebackground=C["border"], cursor="hand2",
                          command=_open_edit).grid(row=r, column=5, sticky="ns", pady=5)
            else:
                tk.Label(self._list_frame, bg=bg).grid(row=r, column=5, sticky="nsew")

            # col 6 — Jira API checkbox
            jira_var = tk.BooleanVar(value=bool(fd.get("jira_key")))
            tk.Checkbutton(self._list_frame, variable=jira_var, bg=bg, fg=C["text"],
                           selectcolor=C["cb_sel"], activebackground=bg,
                           activeforeground=C["text"],
                           relief="flat", bd=0, cursor="hand2"
                           ).grid(row=r, column=6, sticky="ns", pady=4)

            # col 7 — show label checkbox
            show_lbl_var = tk.BooleanVar(value=fd.get("show_label_in_jira", True))
            tk.Checkbutton(self._list_frame, variable=show_lbl_var, bg=bg, fg=C["text"],
                           selectcolor=C["cb_sel"], activebackground=bg,
                           activeforeground=C["text"],
                           relief="flat", bd=0, cursor="hand2"
                           ).grid(row=r, column=7, sticky="ns", pady=4)

            # col 8 — ▲▼✕ actions
            af = tk.Frame(self._list_frame, bg=bg)
            af.grid(row=r, column=8, sticky="ns", padx=4, pady=4)

            def _up(idx=i):
                if idx > 0:
                    self._fields[idx], self._fields[idx-1] = \
                        self._fields[idx-1], self._fields[idx]
                    self._render_rows()
            def _down(idx=i):
                if idx < len(self._fields) - 1:
                    self._fields[idx], self._fields[idx+1] = \
                        self._fields[idx+1], self._fields[idx]
                    self._render_rows()
            def _del(idx=i, fd=fd):
                if not fd.get("custom") and fd.get("required"):
                    messagebox.showwarning("Cannot delete",
                        f'"{fd["label"]}" is required — disable via checkbox instead.',
                        parent=self)
                    return
                if messagebox.askyesno("Remove field",
                        f'Remove "{fd["label"]}"?', parent=self):
                    del self._fields[idx]
                    self._render_rows()

            for txt, cmd, fg in [("▲", _up, C["muted"]),
                                  ("▼", _down, C["muted"]),
                                  ("✕", _del, C["red"])]:
                tk.Button(af, text=txt, bg=bg, fg=fg,
                          relief="flat", font=("Segoe UI", 9), padx=4,
                          activebackground=C["border"],
                          command=cmd, cursor="hand2").pack(side="left")

            # thin separator row
            tk.Frame(self._list_frame, bg=C["border"], height=1).grid(
                row=r + 1, column=0, columnspan=len(self._COLS), sticky="ew")

            self._row_frames.append(dict(
                fd=fd, en_var=en_var, lbl_var=lbl_var,
                def_var=def_var, jira_var=jira_var,
                show_lbl_var=show_lbl_var))

    def _add_field(self):
        """Dialog to create a new custom field."""
        dlg = tk.Toplevel(self)
        dlg.title("Add Custom Field")
        dlg.geometry("520x420")
        dlg.configure(bg=C["bg"])
        dlg.grab_set()

        tk.Label(dlg, text="New Custom Field",
                 bg=C["bg"], fg=C["accent"],
                 font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(16, 12))

        def row(parent, label):
            f = tk.Frame(parent, bg=C["bg"])
            f.pack(fill="x", padx=20, pady=4)
            tk.Label(f, text=label, bg=C["bg"], fg=C["muted"],
                     width=16, anchor="w").pack(side="left")
            return f

        # label
        lbl_var = tk.StringVar()
        f1 = row(dlg, "Field Label *")
        ttk.Entry(f1, textvariable=lbl_var, width=28).pack(side="left")

        # type
        type_var = tk.StringVar(value="text")
        f2 = row(dlg, "Field Type")
        ttk.Combobox(f2, textvariable=type_var,
                     values=["text", "dropdown", "number", "date"],
                     state="readonly", width=14).pack(side="left")

        # default
        def_var = tk.StringVar()
        f3 = row(dlg, "Default Value")
        ttk.Entry(f3, textvariable=def_var, width=28).pack(side="left")

        # dropdown choices (shown when type=dropdown)
        choices_var = tk.StringVar()
        f4 = row(dlg, "Choices (csv)")
        choices_entry = ttk.Entry(f4, textvariable=choices_var, width=28)

        def _on_type(*_):
            if type_var.get() == "dropdown":
                choices_entry.pack(side="left")
            else:
                choices_entry.pack_forget()
        type_var.trace_add("write", _on_type)
        _on_type()

        # jira key
        jira_key_var = tk.StringVar()
        f5 = row(dlg, "Jira Field Key")
        ttk.Entry(f5, textvariable=jira_key_var, width=28).pack(side="left")
        tk.Label(dlg, text="Leave blank to exclude from Jira ticket",
                 bg=C["bg"], fg=C["muted"], font=("Segoe UI", 8)
                 ).pack(anchor="w", padx=36)

        # show label in Jira
        show_lbl_new = tk.BooleanVar(value=True)
        f6 = row(dlg, "Show Label in Jira")
        tk.Checkbutton(f6, variable=show_lbl_new, bg=C["bg"], fg=C["text"],
                       selectcolor=C["cb_sel"], activebackground=C["bg"],
                       activeforeground=C["text"], relief="flat").pack(side="left")
        tk.Label(f6, text="(show 'FieldName: value' in Jira description)",
                 bg=C["bg"], fg=C["muted"], font=("Segoe UI", 8)).pack(side="left", padx=6)

        def _create():
            label = lbl_var.get().strip()
            if not label:
                messagebox.showwarning("Required", "Field Label is required", parent=dlg)
                return
            key     = label.lower().replace(" ", "_")
            choices = [c.strip() for c in choices_var.get().split(",") if c.strip()] \
                      if type_var.get() == "dropdown" else []
            new_fd  = dict(
                label             = label,
                key               = key,
                type              = type_var.get(),
                choices           = choices,
                default           = def_var.get().strip(),
                jira_key          = jira_key_var.get().strip() or None,
                show_label_in_jira= show_lbl_new.get(),
                required          = False,
                enabled           = True,
                custom            = True,
            )
            self._fields.append(new_fd)
            self._render_rows()
            dlg.destroy()

        brow = tk.Frame(dlg, bg=C["bg"])
        brow.pack(fill="x", padx=20, pady=20)
        tk.Button(brow, text="Add Field",
                  bg=C["green"], fg="#0d1117", relief="flat",
                  font=("Segoe UI", 10, "bold"), padx=14, pady=5,
                  cursor="hand2", command=_create).pack(side="right")
        tk.Button(brow, text="Cancel",
                  bg=C["red"], fg="#ffffff", relief="flat",
                  font=("Segoe UI", 10), padx=12, pady=5,
                  activebackground="#c0392b", cursor="hand2",
                  command=dlg.destroy).pack(side="right", padx=(0, 8))

    def _edit_choices_dialog(self, fd, combobox, def_var):
        dlg = tk.Toplevel(self)
        dlg.title(f"Edit choices — {fd['label']}")
        dlg.geometry("400x380")
        dlg.configure(bg=C["bg"])
        dlg.grab_set()
        dlg.transient(self)

        tk.Label(dlg, text=f"Choices for \"{fd['label']}\"",
                 bg=C["bg"], fg=C["accent"],
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(dlg, text="One value per line  •  Order here = order in dropdown",
                 bg=C["bg"], fg=C["muted"],
                 font=("Segoe UI", 8)).pack(anchor="w", padx=16, pady=(0, 8))

        from tkinter import scrolledtext as st_mod
        txt = st_mod.ScrolledText(dlg, height=12,
                                  bg=C["inp"], fg=C["text"],
                                  insertbackground=C["text"],
                                  relief="flat", font=("Segoe UI", 10),
                                  highlightthickness=1,
                                  highlightbackground=C["border"])
        txt.pack(fill="both", expand=True, padx=16, pady=(0, 10))
        txt.insert("1.0", "\n".join(fd.get("choices", [])))

        add_var = tk.StringVar()
        add_row = tk.Frame(dlg, bg=C["bg"]); add_row.pack(fill="x", padx=16, pady=(0, 6))
        ttk.Entry(add_row, textvariable=add_var).pack(side="left", fill="x", expand=True)
        def _add_value():
            v = add_var.get().strip()
            if v:
                txt.insert("end", ("\n" if txt.get("1.0", "end-1c") else "") + v)
                add_var.set("")
        tk.Button(add_row, text="+ Add", bg=C["accent"], fg="#fff",
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=3,
                  cursor="hand2", command=_add_value).pack(side="left", padx=(6, 0))

        def _save():
            lines = [l.strip() for l in txt.get("1.0", "end-1c").splitlines() if l.strip()]
            fd["choices"] = lines
            if combobox:
                combobox.config(values=lines)
                if def_var.get() not in lines and lines:
                    def_var.set(lines[0])
            dlg.destroy()

        br = tk.Frame(dlg, bg=C["bg"]); br.pack(fill="x", padx=16, pady=(0, 14))
        tk.Button(br, text="Save", bg=C["green"], fg="#0d1117", relief="flat",
                  font=("Segoe UI", 10, "bold"), padx=14, pady=5,
                  cursor="hand2", command=_save).pack(side="right")
        tk.Button(br, text="Cancel", bg=C["red"], fg="#ffffff", relief="flat",
                  font=("Segoe UI", 10), padx=12, pady=5,
                  activebackground="#c0392b", cursor="hand2",
                  command=dlg.destroy).pack(side="right", padx=(0, 8))

    def _apply(self):
        # read current widget values back into field dicts
        for meta in self._row_frames:
            fd = meta["fd"]
            fd["enabled"] = meta["en_var"].get()
            fd["default"] = meta["def_var"].get()
            fd["show_label_in_jira"] = meta["show_lbl_var"].get()
            if fd.get("custom"):
                fd["label"] = meta["lbl_var"].get().strip()
            if not meta["jira_var"].get():
                fd["jira_key"] = None
        self.result = self._fields
        self.destroy()


# ─────────────────────────────────────────────────────────────────────────────
#  DiffViewPopup  — GitHub-style split diff viewer
# ─────────────────────────────────────────────────────────────────────────────

class DiffViewPopup(tk.Toplevel):

    def __init__(self, parent, filename, patch, status, adds=0, dels=0):
        super().__init__(parent)
        self.title(f"{APP_NAME}  —  {os.path.basename(filename)}")
        self.geometry("1300x720")
        self.minsize(900, 480)
        self.configure(bg=C["bg"])
        self.transient(parent)
        self._build(filename, patch, status, adds, dels)

    # ── patch parser ──────────────────────────────────────────────────────────

    @staticmethod
    def _split_rows(patch):
        """Convert unified diff patch to split-diff row list.

        Each row: (l_type, l_ln, l_text, r_type, r_ln, r_text)
        l/r_type: 'del' | 'add' | 'ctx' | 'hunk' | 'empty'
        """
        unified = []
        old_ln = new_ln = 0
        for line in patch.splitlines():
            if line.startswith("@@"):
                m = re.match(r"@@ -(\d+)(?:,\d+)? \+(\d+)(?:,\d+)? @@(.*)", line)
                if m:
                    old_ln = int(m.group(1))
                    new_ln = int(m.group(2))
                unified.append(("hunk", None, None, line))
            elif line.startswith("+") and not line.startswith("+++"):
                unified.append(("add", None, new_ln, line[1:]))
                new_ln += 1
            elif line.startswith("-") and not line.startswith("---"):
                unified.append(("del", old_ln, None, line[1:]))
                old_ln += 1
            elif line and line[0] not in ("\\",):
                text = line[1:] if line.startswith(" ") else line
                unified.append(("ctx", old_ln, new_ln, text))
                old_ln += 1; new_ln += 1

        rows = []
        i = 0
        while i < len(unified):
            rtype, a, b, text = unified[i]
            if rtype == "hunk":
                rows.append(("hunk", None, text, "hunk", None, text))
                i += 1
            elif rtype == "ctx":
                rows.append(("ctx", a, text, "ctx", b, text))
                i += 1
            else:
                dels, adds = [], []
                while i < len(unified) and unified[i][0] in ("del", "add"):
                    t, oa, nb, tx = unified[i]
                    (dels if t == "del" else adds).append(
                        (oa if t == "del" else nb, tx))
                    i += 1
                for j in range(max(len(dels), len(adds))):
                    lv = dels[j] if j < len(dels) else None
                    rv = adds[j] if j < len(adds) else None
                    rows.append((
                        "del"   if lv else "empty", lv[0] if lv else None, lv[1] if lv else "",
                        "add"   if rv else "empty", rv[0] if rv else None, rv[1] if rv else "",
                    ))
        return rows

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build(self, filename, patch, status, adds, dels):
        # ── header bar ───────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=C["surface"])
        hdr.pack(fill="x")
        tk.Frame(hdr, bg=C["accent"], width=4).pack(side="left", fill="y")
        tk.Label(hdr, text=f"  {filename}",
                 bg=C["surface"], fg=C["text"],
                 font=("Consolas", 11, "bold")).pack(side="left", padx=8, pady=8)
        badge_bg = STATUS_BADGE.get(status, C["surface"])
        tk.Label(hdr, text=f"  {STATUS_EMOJI.get(status,'~')} {status.upper()}  ",
                 bg=badge_bg, fg="#ffffff",
                 font=("Segoe UI", 9, "bold"), padx=4).pack(side="left", padx=(0, 10))
        sfg = C["green"] if adds > dels else C["red"] if dels > adds else C["muted"]
        tk.Label(hdr, text=f"+{adds}  −{dels}",
                 bg=C["surface"], fg=sfg,
                 font=("Segoe UI", 9, "bold")).pack(side="left")

        # ── column labels ─────────────────────────────────────────────────────
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")
        col_hdr = tk.Frame(self, bg="#161b22")
        col_hdr.pack(fill="x")
        tk.Label(col_hdr, text="  Before (old)",
                 bg="#161b22", fg=C["muted"],
                 font=("Segoe UI", 9, "bold"), anchor="w",
                 ).pack(side="left", fill="x", expand=True, padx=8, pady=4)
        tk.Frame(col_hdr, bg=C["border"], width=1).pack(side="left", fill="y")
        tk.Label(col_hdr, text="  After (new)",
                 bg="#161b22", fg=C["muted"],
                 font=("Segoe UI", 9, "bold"), anchor="w",
                 ).pack(side="left", fill="x", expand=True, padx=8, pady=4)
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")

        # ── pane area ─────────────────────────────────────────────────────────
        pane = tk.Frame(self, bg="#0d1117")
        pane.pack(fill="both", expand=True)

        vsb = ttk.Scrollbar(pane, orient="vertical")
        vsb.pack(side="right", fill="y")

        xsb = ttk.Scrollbar(self, orient="horizontal")
        xsb.pack(fill="x", side="bottom")

        lf = tk.Frame(pane, bg="#0d1117"); lf.pack(side="left", fill="both", expand=True)
        tk.Frame(pane, bg=C["border"], width=1).pack(side="left", fill="y")
        rf = tk.Frame(pane, bg="#0d1117"); rf.pack(side="left", fill="both", expand=True)

        txt_kw = dict(
            bg="#0d1117", fg=C["text"],
            font=("Consolas", 9), relief="flat",
            state="disabled", wrap="none",
            highlightthickness=0, insertbackground=C["text"],
        )

        self._lt = tk.Text(lf, **txt_kw)
        self._rt = tk.Text(rf, **txt_kw)

        _TAG = {
            "del":   dict(background="#200e0e", foreground="#f85149"),
            "add":   dict(background="#0e2318", foreground="#3fb950"),
            "ctx":   dict(foreground="#8b949e"),
            "hunk":  dict(background="#0c1c2c", foreground="#58a6ff"),
            "empty": dict(background="#161b22"),
            "ln":    dict(foreground="#484f58"),
        }
        for t in (self._lt, self._rt):
            for tag, cfg in _TAG.items():
                t.tag_config(tag, **cfg)

        # ── populate ──────────────────────────────────────────────────────────
        rows = self._split_rows(patch)
        self._lt.config(state="normal")
        self._rt.config(state="normal")

        for l_type, l_ln, l_text, r_type, r_ln, r_text in rows:
            if l_type == "hunk":
                self._lt.insert("end", f"  {l_text}\n", "hunk")
                self._rt.insert("end", f"  {r_text}\n", "hunk")
            else:
                l_ln_s = f"{l_ln:>5} " if l_ln is not None else "       "
                r_ln_s = f"{r_ln:>5} " if r_ln is not None else "       "
                sym_l  = "−" if l_type == "del" else (" " if l_type == "ctx" else " ")
                sym_r  = "+" if r_type == "add" else (" " if r_type == "ctx" else " ")

                if l_type == "empty":
                    self._lt.insert("end", "\n", "empty")
                else:
                    self._lt.insert("end", l_ln_s, "ln")
                    self._lt.insert("end", f"{sym_l} {l_text}\n", l_type)

                if r_type == "empty":
                    self._rt.insert("end", "\n", "empty")
                else:
                    self._rt.insert("end", r_ln_s, "ln")
                    self._rt.insert("end", f"{sym_r} {r_text}\n", r_type)

        self._lt.config(state="disabled")
        self._rt.config(state="disabled")

        # ── scrollbar wiring ─────────────────────────────────────────────────
        def _yscroll(*args):
            self._lt.yview(*args)
            self._rt.yview(*args)

        def _on_scroll_l(first, last):
            vsb.set(first, last)
            self._rt.yview_moveto(first)

        def _on_scroll_r(first, last):
            vsb.set(first, last)
            self._lt.yview_moveto(first)

        vsb.config(command=_yscroll)
        self._lt.config(yscrollcommand=_on_scroll_l, xscrollcommand=xsb.set)
        self._rt.config(yscrollcommand=_on_scroll_r, xscrollcommand=xsb.set)
        xsb.config(command=lambda *a: [self._lt.xview(*a), self._rt.xview(*a)])

        self._lt.pack(fill="both", expand=True)
        self._rt.pack(fill="both", expand=True)

        def _mw(e):
            self._lt.yview_scroll(int(-1*(e.delta/120)), "units")
            self._rt.yview_scroll(int(-1*(e.delta/120)), "units")
            return "break"
        self._lt.bind("<MouseWheel>", _mw)
        self._rt.bind("<MouseWheel>", _mw)


# ─────────────────────────────────────────────────────────────────────────────
#  FileChangesPopup  (unchanged from v2, just cleaner)
# ─────────────────────────────────────────────────────────────────────────────

class FileChangesPopup(tk.Toplevel):

    def __init__(self, parent, files, pr_title):
        super().__init__(parent)
        self.title(f"{APP_NAME}  —  {pr_title[:70]}")
        self.geometry("1120x760")
        self.minsize(900, 560)
        self.configure(bg=C["bg"])
        self.grab_set()
        self.transient(parent)
        # sort by dependency score: no-dep files first
        self._files  = sorted(files,
                              key=lambda f: (_dep_score(f.get("filename", "")),
                                             f.get("filename", "")))
        self._rows   = []
        self._result = None
        self._build()

    def _build(self):
        bar = tk.Frame(self, bg=C["surface"])
        bar.pack(fill="x")
        tk.Frame(bar, bg=C["accent"], width=4).pack(side="left", fill="y")
        hdr_f = tk.Frame(bar, bg=C["surface"])
        hdr_f.pack(side="left", padx=14, pady=10, fill="x", expand=True)
        tk.Label(hdr_f,
                 text=f"{len(self._files)} files changed",
                 bg=C["surface"], fg=C["text"],
                 font=("Segoe UI", 13, "bold")).pack(side="left")
        tk.Label(hdr_f,
                 text="   Sorted by dependency  •  Change Description → Word doc  •  Jira Comment → Jira ticket",
                 bg=C["surface"], fg=C["muted"],
                 font=("Segoe UI", 9)).pack(side="left", pady=(3, 0))

        for lbl, st, abg, afg in [
            ("Select All",   True,  C["accent"],  "#ffffff"),
            ("Deselect All", False, C["orange"],  "#0d1117"),
        ]:
            tk.Button(bar, text=lbl,
                      bg=abg, fg=afg, relief="flat",
                      font=("Segoe UI", 9, "bold"), padx=10, pady=4,
                      activebackground=C["border"], cursor="hand2",
                      command=lambda s=st: self._toggle_all(s)
                      ).pack(side="right", padx=4, pady=10)

        tk.Button(bar, text="Reset Step Comments",
                  bg=C["yellow"], fg="#0d1117", relief="flat",
                  font=("Segoe UI", 9, "bold"), padx=10, pady=4,
                  activebackground=C["orange"], cursor="hand2",
                  command=self._reset_step_comments
                  ).pack(side="right", padx=4, pady=10)

        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")

        wrap = tk.Frame(self, bg=C["bg"])
        wrap.pack(fill="both", expand=True)
        self._canvas = tk.Canvas(wrap, bg=C["bg"], highlightthickness=0)
        vbar = ttk.Scrollbar(wrap, orient="vertical", command=self._canvas.yview)
        hbar = ttk.Scrollbar(wrap, orient="horizontal", command=self._canvas.xview)
        self._canvas.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        hbar.pack(side="bottom", fill="x")
        vbar.pack(side="right",  fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self._lf = tk.Frame(self._canvas, bg=C["bg"])
        wid = self._canvas.create_window((0, 0), window=self._lf, anchor="nw")
        self._canvas.bind("<Configure>",
                          lambda e: self._canvas.itemconfig(wid, width=e.width))
        self._lf.bind("<Configure>",
                      lambda e: self._canvas.configure(
                          scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<MouseWheel>",
                          lambda e: self._canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        self._build_rows()

        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")
        bot = tk.Frame(self, bg=C["surface"], height=52)
        bot.pack(fill="x", side="bottom"); bot.pack_propagate(False)
        self._sv = tk.StringVar(value="")
        tk.Label(bot, textvariable=self._sv, bg=C["surface"], fg=C["green"],
                 font=("Segoe UI", 9)).pack(side="left", padx=16)
        tk.Button(bot, text="Cancel",
                  bg=C["red"], fg="#ffffff", relief="flat",
                  font=("Segoe UI", 10), padx=12, pady=5,
                  activebackground="#c0392b", cursor="hand2",
                  command=self.destroy).pack(side="right", padx=8, pady=10)
        tk.Button(bot, text="  Save & Close  ",
                  bg=C["green"], fg="#0d1117", relief="flat",
                  font=("Segoe UI", 10, "bold"), padx=14, pady=5,
                  activebackground="#2ea043", cursor="hand2",
                  command=self._save).pack(side="right", padx=4, pady=10)

    @staticmethod
    def _default_step_comment(step, filename, status):
        verb = STATUS_VERB.get(status, "Modify")
        base = os.path.basename(filename)
        return f"Step {step}: {verb} {base}"

    def _build_rows(self):
        total = len(self._files)
        for idx, f in enumerate(self._files, 1):
            fname   = f.get("filename", "")
            status  = f.get("status", "modified")
            adds    = f.get("additions", 0)
            dels    = f.get("deletions", 0)
            rbg     = STATUS_ROW.get(status, C["card"])
            bbg     = STATUS_BADGE.get(status, C["surface"])
            em      = STATUS_EMOJI.get(status, "~")

            card = tk.Frame(self._lf, bg=rbg,
                            highlightbackground=C["border"], highlightthickness=1)
            card.pack(fill="x", padx=10, pady=3)

            def _sw(e, c=self._canvas):
                c.yview_scroll(int(-1*(e.delta/120)), "units")
            card.bind("<MouseWheel>", _sw)

            hdr = tk.Frame(card, bg=rbg)
            hdr.pack(fill="x", padx=8, pady=(8, 4))

            inc = tk.BooleanVar(value=True)
            tk.Checkbutton(hdr, variable=inc, bg=rbg, fg=C["text"],
                           selectcolor=C["cb_sel"],
                           activebackground=rbg, activeforeground=C["text"],
                           text="Include", font=("Segoe UI", 8),
                           bd=0, relief="flat", cursor="hand2").pack(side="left")

            # step badge
            tk.Label(hdr, text=f"  Step {idx}/{total}  ",
                     bg=C["accent"], fg="#ffffff",
                     font=("Segoe UI", 8, "bold"), padx=4, pady=2
                     ).pack(side="left", padx=(8, 6))

            tk.Label(hdr, text=f"  {em}  {status.upper()}  ",
                     bg=bbg, fg="#ffffff",
                     font=("Segoe UI", 8, "bold"), padx=4, pady=2
                     ).pack(side="left", padx=(0, 10))

            parts = fname.split("/")
            pre   = "/".join(parts[:-2]) + "/" if len(parts) > 2 else ""
            tail  = "/".join(parts[-2:])
            pf    = tk.Frame(hdr, bg=rbg); pf.pack(side="left", fill="x", expand=True)
            if pre:
                tk.Label(pf, text=pre, bg=rbg, fg=C["muted"],
                         font=("Consolas", 10)).pack(side="left")
            tk.Label(pf, text=tail, bg=rbg, fg=C["text"],
                     font=("Consolas", 10, "bold")).pack(side="left")

            sfg = C["green"] if adds > dels else C["red"] if dels > adds else C["muted"]
            tk.Label(hdr, text=f"+{adds}  -{dels}", bg=rbg, fg=sfg,
                     font=("Segoe UI", 9, "bold")).pack(side="right", padx=8)

            body = tk.Frame(card, bg=rbg); body.pack(fill="x", padx=8, pady=(0, 8))

            # ── GitHub link + diff toggle row ────────────────────────────────
            gh_row = tk.Frame(body, bg=rbg)
            gh_row.pack(fill="x", pady=(0, 4))

            gh_url = f.get("blob_url") or f.get("html_url", "")
            if gh_url:
                tk.Button(gh_row, text="View on GitHub ↗",
                          bg=rbg, fg=C["accent"], relief="flat",
                          font=("Segoe UI", 8), padx=6, pady=2, cursor="hand2",
                          activebackground=C["border"],
                          command=lambda u=gh_url: webbrowser.open(u)
                          ).pack(side="left")

            patch = f.get("patch", "")
            if patch:
                n_lines = len(patch.splitlines())
                tk.Button(
                    gh_row,
                    text=f"⊞ View Changes  ({n_lines} lines)",
                    bg=rbg, fg=C["accent"], relief="flat",
                    font=("Segoe UI", 8, "bold"), padx=6, pady=2, cursor="hand2",
                    activebackground=C["border"],
                    command=lambda fn=fname, p=patch, st=status, a=adds, d=dels:
                        DiffViewPopup(self, fn, p, st, a, d)
                ).pack(side="left", padx=(6, 0))

            # ── two comment columns ──────────────────────────────────────────
            comment_row = tk.Frame(body, bg=rbg)
            comment_row.pack(fill="x", pady=(0, 4))

            # Word comment (left)
            word_col = tk.Frame(comment_row, bg=rbg)
            word_col.pack(side="left", fill="both", expand=True, padx=(0, 6))
            tk.Label(word_col, text="Change Description (Word doc):", bg=rbg, fg=C["yellow"],
                     font=("Segoe UI", 9, "bold")).pack(anchor="w")
            word_cbox = scrolledtext.ScrolledText(
                word_col, height=3, bg=C["inp"], fg=C["text"],
                insertbackground=C["text"], selectbackground=C["hover"],
                font=("Segoe UI", 10), relief="flat", wrap="word")
            word_cbox.pack(fill="both", expand=True, pady=(2, 0))
            word_cbox.bind("<MouseWheel>", _sw)
            word_cbox.insert("1.0", self._default_step_comment(idx, fname, status))

            # Jira comment (right)
            jira_col = tk.Frame(comment_row, bg=rbg)
            jira_col.pack(side="left", fill="both", expand=True)
            tk.Label(jira_col, text="Jira Comment:", bg=rbg, fg=C["accent"],
                     font=("Segoe UI", 9, "bold")).pack(anchor="w")
            jira_cbox = scrolledtext.ScrolledText(
                jira_col, height=3, bg=C["inp"], fg=C["text"],
                insertbackground=C["text"], selectbackground=C["hover"],
                font=("Segoe UI", 10), relief="flat", wrap="word")
            jira_cbox.pack(fill="both", expand=True, pady=(2, 0))
            jira_cbox.bind("<MouseWheel>", _sw)
            jira_cbox.insert("1.0", self._default_step_comment(idx, fname, status))

            # ── screenshots with per-screenshot comments ─────────────────────
            shots      = []   # list of {"path": str, "var": StringVar, "frame": Frame}
            shots_cv   = tk.StringVar(value="No screenshots")
            shots_area = tk.Frame(body, bg=rbg)
            shots_area.pack(fill="x", pady=(6, 0))

            def _add_shot_row(path, sa=shots_area, s=shots, sv=shots_cv, _rbg=rbg):
                var = tk.StringVar()
                fr  = tk.Frame(sa, bg=C["surface"],
                               highlightbackground=C["border"], highlightthickness=1)
                fr.pack(fill="x", pady=2)
                item = {"path": path, "var": var, "frame": fr}
                s.append(item)
                # thumbnail
                try:
                    from PIL import Image, ImageTk
                    img = Image.open(path); img.thumbnail((56, 42))
                    ph  = ImageTk.PhotoImage(img)
                    lb  = tk.Label(fr, image=ph, bg=C["surface"], cursor="hand2")
                    lb.pack(side="left", padx=(4, 6), pady=4)
                    lb.image = ph
                    lb.bind("<Button-1>", lambda e, p=path: webbrowser.open(p))
                except Exception:
                    tk.Label(fr, text="IMG", bg=C["surface"], fg=C["accent"],
                             font=("Consolas", 9), padx=6).pack(
                                 side="left", padx=(4, 6), pady=4)
                # filename label
                tk.Label(fr, text=os.path.basename(path)[:22],
                         bg=C["surface"], fg=C["muted"],
                         font=("Consolas", 8)).pack(side="left")
                # comment entry
                ent = ttk.Entry(fr, textvariable=var)
                ent.pack(side="left", fill="x", expand=True, padx=(6, 4), pady=4)
                ent.insert(0, "Screenshot caption...")
                ent.bind("<FocusIn>",
                         lambda e, w=ent: w.delete(0, "end")
                         if w.get() == "Screenshot caption..." else None)
                ent.bind("<MouseWheel>", _sw)
                # remove button
                def _rm(i=item, f=fr, s=s, sv=sv):
                    s.remove(i)
                    f.destroy()
                    n = len(s)
                    sv.set(f"{n} screenshot{'s' if n!=1 else ''}" if n else "No screenshots")
                tk.Button(fr, text="✕", bg=C["surface"], fg=C["red"], relief="flat",
                          font=("Segoe UI", 9, "bold"), padx=6, pady=2, cursor="hand2",
                          command=_rm).pack(side="right", padx=4, pady=4)
                n = len(s)
                sv.set(f"{n} screenshot{'s' if n!=1 else ''}" if n else "No screenshots")

            srow = tk.Frame(body, bg=rbg); srow.pack(fill="x")
            tk.Label(srow, textvariable=shots_cv, bg=rbg, fg=C["muted"],
                     font=("Segoe UI", 9)).pack(side="left", padx=(0, 8))

            def _add_shots(sa=shots_area, s=shots, sv=shots_cv):
                paths = filedialog.askopenfilenames(
                    title="Select screenshots",
                    filetypes=[("Images","*.png *.jpg *.jpeg *.gif *.bmp *.webp *.tiff"),
                               ("All","*.*")])
                for p in paths:
                    if p and not any(x["path"] == p for x in s):
                        _add_shot_row(p, sa, s, sv)

            def _clr_shots(s=shots, sa=shots_area, sv=shots_cv):
                for x in s:
                    x["frame"].destroy()
                s.clear()
                sv.set("No screenshots")

            tk.Button(srow, text="+ Add Screenshot",
                      bg=C["green"], fg="#0d1117", relief="flat",
                      font=("Segoe UI", 9, "bold"), padx=10, pady=3,
                      activebackground="#2ea043", cursor="hand2",
                      command=_add_shots).pack(side="left", padx=(0, 6))
            tk.Button(srow, text="Clear",
                      bg=C["orange"], fg="#0d1117", relief="flat",
                      font=("Segoe UI", 9, "bold"), padx=8, pady=3,
                      activebackground=C["yellow"], cursor="hand2",
                      command=_clr_shots).pack(side="left")

            self._rows.append(dict(filename=fname, status=status,
                                   additions=adds, deletions=dels,
                                   word_cbox=word_cbox, jira_cbox=jira_cbox,
                                   shots=shots, inc=inc))

    def _toggle_all(self, state):
        for r in self._rows: r["inc"].set(state)

    def _reset_step_comments(self):
        """Re-fill both comment boxes with default step comments."""
        for idx, r in enumerate(self._rows, 1):
            default = self._default_step_comment(idx, r["filename"], r["status"])
            for box in (r["word_cbox"], r["jira_cbox"]):
                box.delete("1.0", "end")
                box.insert("1.0", default)

    def _save(self):
        res = []
        for r in self._rows:
            if not r["inc"].get(): continue
            word_comment = r["word_cbox"].get("1.0", "end-1c").strip()
            jira_comment = r["jira_cbox"].get("1.0", "end-1c").strip()
            shots_out = [
                {"path": x["path"],
                 "comment": x["var"].get().strip()
                            if x["var"].get().strip() != "Screenshot caption..." else ""}
                for x in r["shots"]
            ]
            res.append(dict(filename=r["filename"], status=r["status"],
                            additions=r["additions"], deletions=r["deletions"],
                            word_comment=word_comment,
                            jira_comment=jira_comment,
                            comment=word_comment,   # backward compat
                            screenshots=shots_out))
        self._result = res
        n_wc = sum(1 for x in res if x["word_comment"])
        n_jc = sum(1 for x in res if x["jira_comment"])
        n_s  = sum(1 for x in res if x["screenshots"])
        self._sv.set(
            f"Saved {len(res)} files  —  {n_wc} word comments  |  "
            f"{n_jc} jira comments  |  {n_s} screenshots")
        self.after(400, self.destroy)

    def get_result(self):
        return self._result if self._result is not None else []


# ─────────────────────────────────────────────────────────────────────────────
#  SQL PR File popup  –  combine changed files into a single runnable script
# ─────────────────────────────────────────────────────────────────────────────

class SqlFilePopup(tk.Toplevel):

    def __init__(self, parent, files, pr_data, config_data):
        super().__init__(parent)
        self.title(f"{APP_NAME}  —  SQL File Generator")
        self.geometry("1020x700")
        self.minsize(820, 520)
        self.configure(bg=C["bg"])
        self.grab_set()
        self.transient(parent)

        self._files   = sorted(files,
                               key=lambda f: (_dep_score(f.get("filename", "")),
                                              f.get("filename", "")))
        self._pr      = pr_data
        self._cfg     = config_data
        self._rows    = []
        self._result  = None

        pr_num = pr_data.get("number", "combined")
        out_dir = config_data.get("word_doc_output_dir", BASE_DIR)
        pr_dir  = os.path.join(out_dir, f"PR_{pr_num}")
        self._out_var = tk.StringVar(
            value=os.path.join(pr_dir, f"PR_{pr_num}.sql"))

        self._build()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build(self):
        bar = tk.Frame(self, bg=C["surface"])
        bar.pack(fill="x")
        tk.Frame(bar, bg=C["yellow"], width=4).pack(side="left", fill="y")
        hf = tk.Frame(bar, bg=C["surface"])
        hf.pack(side="left", padx=14, pady=10, fill="x", expand=True)
        tk.Label(hf, text="SQL PR File",
                 bg=C["surface"], fg=C["yellow"],
                 font=("Segoe UI", 13, "bold")).pack(side="left")
        tk.Label(hf,
                 text="   Files sorted by dependency  •  Edit SQL comment per file  "
                      "•  Combined into one runnable .sql script",
                 bg=C["surface"], fg=C["muted"],
                 font=("Segoe UI", 9)).pack(side="left")
        for lbl, st, abg, afg in [
            ("Select All",   True,  C["accent"], "#ffffff"),
            ("Deselect All", False, C["orange"], "#0d1117"),
        ]:
            tk.Button(bar, text=lbl, bg=abg, fg=afg,
                      relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=4,
                      activebackground=C["border"], cursor="hand2",
                      command=lambda s=st: self._toggle_all(s)
                      ).pack(side="right", padx=4, pady=10)

        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")

        wrap = tk.Frame(self, bg=C["bg"]); wrap.pack(fill="both", expand=True)
        self._canvas = tk.Canvas(wrap, bg=C["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self._lf = tk.Frame(self._canvas, bg=C["bg"])
        wid = self._canvas.create_window((0, 0), window=self._lf, anchor="nw")
        self._canvas.bind("<Configure>",
                          lambda e: self._canvas.itemconfig(wid, width=e.width))
        self._lf.bind("<Configure>",
                      lambda e: self._canvas.configure(
                          scrollregion=(0, 0, e.width, e.height)))
        self._canvas.bind("<MouseWheel>",
                          lambda e: self._canvas.yview_scroll(
                              int(-1*(e.delta/120)), "units"))

        self._build_rows()

        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")

        bot = tk.Frame(self, bg=C["surface"]); bot.pack(fill="x", side="bottom")
        out_row = tk.Frame(bot, bg=C["surface"]); out_row.pack(fill="x", padx=16, pady=(10, 4))
        tk.Label(out_row, text="Output file:",
                 bg=C["surface"], fg=C["muted"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))
        ttk.Entry(out_row, textvariable=self._out_var).pack(side="left", fill="x", expand=True)
        tk.Button(out_row, text="Browse",
                  bg=C["purple"], fg="#ffffff", relief="flat",
                  font=("Segoe UI", 9, "bold"), padx=8, pady=3, cursor="hand2",
                  activebackground="#9966cc",
                  command=self._browse).pack(side="left", padx=(6, 0))

        btn_row = tk.Frame(bot, bg=C["surface"]); btn_row.pack(fill="x", padx=16, pady=(0, 12))
        self._sv = tk.StringVar(value="")
        tk.Label(btn_row, textvariable=self._sv,
                 bg=C["surface"], fg=C["green"],
                 font=("Segoe UI", 9)).pack(side="left")
        tk.Button(btn_row, text="Cancel",
                  bg=C["red"], fg="#ffffff", relief="flat",
                  font=("Segoe UI", 10), padx=12, pady=5, cursor="hand2",
                  activebackground="#c0392b",
                  command=self.destroy).pack(side="right", padx=8)
        tk.Button(btn_row, text="  Generate SQL File  ",
                  bg=C["green"], fg="#0d1117", relief="flat",
                  font=("Segoe UI", 10, "bold"), padx=14, pady=5, cursor="hand2",
                  activebackground="#2ea043",
                  command=self._generate).pack(side="right", padx=4)

    def _build_rows(self):
        for f in self._files:
            fname    = f.get("filename", "")
            status   = f.get("status", "modified")
            adds     = f.get("additions", 0)
            dels     = f.get("deletions", 0)
            patch    = f.get("patch", "")
            obj_name = _detect_sql_object(fname, patch)
            self._rows.append(dict(
                filename=fname, status=status,
                additions=adds, deletions=dels,
                patch=patch, obj_name=obj_name,
                inc=tk.BooleanVar(value=(status != "removed")),
                cmt_var=tk.StringVar(value=obj_name),
                raw_url=f.get("raw_url", ""),
                full_content=f.get("full_content"),
                gh_url=f.get("blob_url") or f.get("html_url", ""),
            ))
        self._render_rows()

    def _render_rows(self):
        for w in self._lf.winfo_children():
            w.destroy()
        total = len(self._rows)

        def _sw(e, c=self._canvas):
            c.yview_scroll(int(-1*(e.delta/120)), "units")

        for i, r in enumerate(self._rows):
            fname    = r["filename"]
            status   = r["status"]
            adds     = r["additions"]
            dels     = r["deletions"]
            inc      = r["inc"]
            cmt_var  = r["cmt_var"]
            obj_name = r["obj_name"]
            gh_url   = r["gh_url"]
            rbg      = STATUS_ROW.get(status, C["card"])
            bbg      = STATUS_BADGE.get(status, C["surface"])
            em       = STATUS_EMOJI.get(status, "~")
            idx      = i + 1

            card = tk.Frame(self._lf, bg=rbg,
                            highlightbackground=C["border"], highlightthickness=1)
            card.pack(fill="x", padx=10, pady=3)
            card.bind("<MouseWheel>", _sw)

            hdr = tk.Frame(card, bg=rbg); hdr.pack(fill="x", padx=8, pady=(8, 4))

            # ▲▼ reorder buttons
            nav = tk.Frame(hdr, bg=rbg); nav.pack(side="left", padx=(0, 4))
            tk.Button(nav, text="▲", bg=C["card"], fg=C["text"], relief="flat",
                      font=("Segoe UI", 7), padx=3, pady=0, cursor="hand2",
                      activebackground=C["border"],
                      state="normal" if i > 0 else "disabled",
                      command=lambda ii=i: self._move_row(ii, -1)
                      ).pack(side="top")
            tk.Button(nav, text="▼", bg=C["card"], fg=C["text"], relief="flat",
                      font=("Segoe UI", 7), padx=3, pady=0, cursor="hand2",
                      activebackground=C["border"],
                      state="normal" if i < total - 1 else "disabled",
                      command=lambda ii=i: self._move_row(ii, 1)
                      ).pack(side="top")

            tk.Checkbutton(hdr, variable=inc, bg=rbg, fg=C["text"],
                           selectcolor=C["cb_sel"],
                           activebackground=rbg, activeforeground=C["text"],
                           text="Include", font=("Segoe UI", 8),
                           bd=0, relief="flat", cursor="hand2").pack(side="left")

            tk.Label(hdr, text=f"  Step {idx}/{total}  ",
                     bg=C["accent"], fg="#ffffff",
                     font=("Segoe UI", 8, "bold"), padx=4, pady=2
                     ).pack(side="left", padx=(8, 6))
            tk.Label(hdr, text=f"  {em}  {status.upper()}  ",
                     bg=bbg, fg="#ffffff",
                     font=("Segoe UI", 8, "bold"), padx=4, pady=2
                     ).pack(side="left", padx=(0, 10))

            parts = fname.split("/")
            pre   = "/".join(parts[:-2]) + "/" if len(parts) > 2 else ""
            tail  = "/".join(parts[-2:])
            pf    = tk.Frame(hdr, bg=rbg); pf.pack(side="left", fill="x", expand=True)
            if pre:
                tk.Label(pf, text=pre, bg=rbg, fg=C["muted"],
                         font=("Consolas", 10)).pack(side="left")
            tk.Label(pf, text=tail, bg=rbg, fg=C["text"],
                     font=("Consolas", 10, "bold")).pack(side="left")
            # Show Code Change — compact, next to filename
            patch = r["patch"]
            if patch:
                n_lines = len(patch.splitlines())
                tk.Button(pf,
                          text=f"⊞ Code Change ({n_lines})",
                          bg=C["accent"], fg="#ffffff", relief="flat",
                          font=("Segoe UI", 8, "bold"),
                          padx=6, pady=1, cursor="hand2",
                          activebackground=C["hover"], activeforeground="#ffffff",
                          command=lambda fn=fname, p=patch, st=status, a=adds, d=dels:
                              DiffViewPopup(self, fn, p, st, a, d)
                          ).pack(side="left", padx=(8, 0))

            sfg = C["green"] if adds > dels else C["red"] if dels > adds else C["muted"]
            tk.Label(hdr, text=f"+{adds}  −{dels}", bg=rbg, fg=sfg,
                     font=("Segoe UI", 9, "bold")).pack(side="right", padx=8)

            body = tk.Frame(card, bg=rbg); body.pack(fill="x", padx=8, pady=(0, 8))

            # SQL comment row — "Step N:" prefix is read-only label, entry holds description only
            cmt_row = tk.Frame(body, bg=rbg); cmt_row.pack(fill="x", pady=(0, 4))
            tk.Label(cmt_row, text="SQL Comment:",
                     bg=rbg, fg=C["yellow"], font=("Segoe UI", 9, "bold")
                     ).pack(side="left", padx=(0, 6))
            tk.Label(cmt_row, text=f"Step {idx}:",
                     bg=rbg, fg=C["accent"], font=("Segoe UI", 9, "bold")
                     ).pack(side="left", padx=(0, 4))
            cmt_ent = ttk.Entry(cmt_row, textvariable=cmt_var)
            cmt_ent.pack(side="left", fill="x", expand=True)
            cmt_ent.bind("<MouseWheel>", _sw)

            info_row = tk.Frame(body, bg=rbg); info_row.pack(fill="x")
            tk.Label(info_row, text=f"Detected: {obj_name}",
                     bg=rbg, fg=C["muted"],
                     font=("Segoe UI", 8, "italic")).pack(side="left")
            if gh_url:
                tk.Button(info_row, text="View on GitHub ↗",
                          bg=rbg, fg=C["accent"], relief="flat",
                          font=("Segoe UI", 8), padx=6, pady=1, cursor="hand2",
                          activebackground=C["border"],
                          command=lambda u=gh_url: webbrowser.open(u)
                          ).pack(side="right")

    def _move_row(self, i, direction):
        j = i + direction
        if 0 <= j < len(self._rows):
            self._rows[i], self._rows[j] = self._rows[j], self._rows[i]
        self._render_rows()

    # ── actions ───────────────────────────────────────────────────────────────

    def _toggle_all(self, state):
        for r in self._rows: r["inc"].set(state)

    def _browse(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".sql",
            filetypes=[("SQL files", "*.sql"), ("All files", "*.*")],
            initialfile=os.path.basename(self._out_var.get()),
            initialdir=os.path.dirname(self._out_var.get()) or BASE_DIR,
            parent=self)
        if p:
            self._out_var.set(p)

    def _fetch_content(self, raw_url):
        if not raw_url:
            return None
        try:
            tok_file = self._cfg.get("github_token_file", "")
            token = ""
            if tok_file and os.path.exists(tok_file):
                with open(tok_file) as fh:
                    token = fh.read().strip()
            headers = {"Authorization": f"token {token}"} if token else {}
            r = requests.get(raw_url, headers=headers, timeout=15)
            return r.text if r.ok else None
        except Exception:
            return None

    def _generate(self):
        included = [r for r in self._rows if r["inc"].get()]
        if not included:
            messagebox.showwarning("No files", "Select at least one file.", parent=self)
            return
        out = self._out_var.get().strip()
        if not out:
            messagebox.showwarning("No path", "Set an output file path.", parent=self)
            return
        if not out.lower().endswith(".sql"):
            out += ".sql"

        self._sv.set(f"Fetching {len(included)} file(s)…")
        self.update_idletasks()

        pr_num   = self._pr.get("number", "")
        pr_title = self._pr.get("title", "")

        blk = []
        blk.append("-- " + "=" * 68)
        blk.append(f"-- PR #{pr_num}: {pr_title}")
        blk.append(f"-- Generated : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        blk.append(f"-- Files     : {len(included)}")
        blk.append("-- " + "=" * 68)
        blk.append("")

        for idx, r in enumerate(included, 1):
            fname   = r["filename"]
            status  = r["status"]
            adds    = r["additions"]
            dels    = r["deletions"]
            desc    = r["cmt_var"].get().strip() or os.path.basename(fname)
            comment = f"Step {idx}: {desc}"

            blk.append("-- " + "-" * 68)
            blk.append(f"-- {comment}")
            blk.append(f"-- File  : {fname}")
            blk.append(f"-- Status: {status.upper()}  (+{adds} / -{dels})")
            blk.append("-- " + "-" * 68)
            blk.append("")

            content = None

            # 1. raw_url → full file from GitHub (real production use)
            if r.get("raw_url"):
                content = self._fetch_content(r["raw_url"])

            # 2. pre-provided full content (mock data / offline)
            if content is None and r.get("full_content"):
                content = r["full_content"]

            # 3. added files only — reconstruct from patch + lines
            if content is None:
                patch = r.get("patch", "")
                if patch and status == "added":
                    content = "\n".join(
                        l[1:] for l in patch.splitlines()
                        if l.startswith('+') and not l.startswith('+++'))

            if content:
                lines = content.rstrip().splitlines()
                while lines and lines[-1].strip() == "/":
                    lines.pop()
                blk.append("\n".join(lines).rstrip())
            else:
                blk.append(f"-- *** CONTENT UNAVAILABLE for {fname}")
                blk.append(f"-- *** Configure GitHub token or provide raw_url to fetch full file.")

            blk.append("")
            blk.append("")

        try:
            os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
            with open(out, "w", encoding="utf-8") as fh:
                fh.write("\n".join(blk))
            self._sv.set(f"Saved: {os.path.basename(out)}")
            self._result = out
            if messagebox.askyesno("Done",
                                   f"SQL file saved:\n{out}\n\nOpen now?",
                                   parent=self):
                webbrowser.open(out)
        except Exception as ex:
            self._sv.set(f"Error: {ex}")
            messagebox.showerror("Error", str(ex), parent=self)

    def get_result(self):
        return self._result


# ─────────────────────────────────────────────────────────────────────────────
#  Main Application
# ─────────────────────────────────────────────────────────────────────────────

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.config_data    = load_config()
        self.pr_cache       = None
        self.pr_files       = []
        self.file_comments  = []
        self._jira_url      = ""
        self._last_doc_path = None
        self._last_sql_path = None
        self._field_widgets = {}
        self._theme_name    = "dark"

        self.title(APP_NAME)
        self.geometry("1060x940")
        self.minsize(880, 720)
        self.configure(bg=C["bg"])
        self._apply_theme()
        self._load_app_icon()
        self._build_ui()

    # ── icon ──────────────────────────────────────────────────────────────────

    def _load_app_icon(self):
        icon_path = os.path.join(BASE_DIR, "prism_icon.ico")

        # Regenerate if missing or zero-byte
        if not os.path.exists(icon_path) or os.path.getsize(icon_path) == 0:
            _, icon_path2 = _generate_app_assets()
            if icon_path2:
                icon_path = icon_path2

        if os.path.exists(icon_path) and os.path.getsize(icon_path) > 0:
            try:
                # default= propagates to taskbar AND all child windows on Windows
                self.iconbitmap(default=icon_path)
                return
            except Exception:
                pass

        # Fallback: iconphoto via PIL (non-Windows or .ico unavailable)
        try:
            from PIL import Image, ImageTk
            logo = os.path.join(BASE_DIR, "prism_logo.png")
            src  = logo if os.path.exists(logo) else icon_path
            if os.path.exists(src):
                im = Image.open(src).resize((64, 64), Image.LANCZOS)
                self._app_photo = ImageTk.PhotoImage(im)
                self.iconphoto(True, self._app_photo)
        except Exception:
            pass

    # ── theme ─────────────────────────────────────────────────────────────────

    def _apply_theme(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".",
                    background=C["bg"], foreground=C["text"],
                    fieldbackground=C["inp"], bordercolor=C["border"],
                    darkcolor=C["surface"], lightcolor=C["surface"],
                    troughcolor=C["surface"],
                    selectbackground=C["hover"], selectforeground=C["text"],
                    font=("Segoe UI", 10))
        s.configure("TFrame",  background=C["bg"])
        s.configure("TLabel",  background=C["bg"], foreground=C["text"])
        s.configure("TEntry",
                    fieldbackground=C["inp"], foreground=C["text"],
                    insertcolor=C["text"], bordercolor=C["border"],
                    padding=(6, 4))
        s.configure("TCombobox",
                    fieldbackground=C["inp"], foreground=C["text"],
                    selectbackground=C["hover"], arrowcolor=C["accent"],
                    padding=(4, 4))
        s.map("TCombobox",
              fieldbackground=[("readonly", C["inp"])],
              foreground=[("readonly", C["text"])])
        s.configure("TScrollbar",
                    background=C["surface"], troughcolor=C["surface"],
                    bordercolor=C["border"], arrowcolor=C["muted"],
                    relief="flat")

        btn_text = "#ffffff" if C["tag"] == "dark" else "#ffffff"
        for name, bg, fg, active in [
            ("Accent",  C["accent"],  "#ffffff", C["hover"]),
            ("Success", C["green"],   "#ffffff", C["green"]),
            ("Purple",  C["purple"],  "#ffffff", C["purple"]),
            ("Orange",  C["orange"],  "#ffffff", C["orange"]),
            ("Ghost",   C["surface"], C["muted"], C["border"]),
            ("Danger",  C["red"],     "#ffffff", C["red"]),
        ]:
            s.configure(f"{name}.TButton",
                        background=bg, foreground=fg,
                        padding=(14, 7), relief="flat",
                        font=("Segoe UI", 10, "bold"),
                        borderwidth=0)
            s.map(f"{name}.TButton",
                  background=[("active", active), ("pressed", active)],
                  foreground=[("active", fg)])

        s.configure("TNotebook", background=C["bg"], borderwidth=0)
        s.configure("TNotebook.Tab",
                    background=C["surface"], foreground=C["muted"],
                    padding=[18, 8], font=("Segoe UI", 10, "bold"),
                    borderwidth=0)
        s.map("TNotebook.Tab",
              background=[("selected", C["bg"])],
              foreground=[("selected", C["accent"])],
              expand=[("selected", [0, 0, 0, 2])])
        s.configure("TProgressbar",
                    troughcolor=C["border"],
                    background=C["accent"],
                    bordercolor=C["border"],
                    thickness=6)

    # ── theme toggle ──────────────────────────────────────────────────────────

    def _toggle_theme(self):
        state = self._save_app_state()
        self._theme_name = "light" if self._theme_name == "dark" else "dark"
        C.update(THEMES[self._theme_name])
        STATUS_BADGE.clear(); STATUS_BADGE.update(C["st_badge"])
        STATUS_ROW.clear();   STATUS_ROW.update(C["st_row"])
        for w in self.winfo_children():
            w.destroy()
        self.configure(bg=C["bg"])
        self._apply_theme()
        self._build_ui()
        self._restore_app_state(state)

    def _save_app_state(self):
        state = {
            "pr_url":        getattr(self, "pr_url_var", tk.StringVar()).get(),
            "pr_cache":      self.pr_cache,
            "pr_files":      self.pr_files,
            "file_comments": self.file_comments,
            "_jira_url":     self._jira_url,
            "_last_doc_path":self._last_doc_path,
            "field_values":  {k: v.get() for k, v in self._field_widgets.items()},
            "description":   "",
        }
        if hasattr(self, "desc_text"):
            d = self.desc_text.get("1.0", "end-1c")
            if d != "Describe the issue in detail...":
                state["description"] = d
        return state

    def _restore_app_state(self, state):
        if state.get("pr_url") and hasattr(self, "pr_url_var"):
            self.pr_url_var.set(state["pr_url"])
        self.pr_cache       = state.get("pr_cache")
        self.pr_files       = state.get("pr_files", [])
        self.file_comments  = state.get("file_comments", [])
        self._jira_url      = state.get("_jira_url", "")
        self._last_doc_path = state.get("_last_doc_path")
        if self.pr_cache and hasattr(self, "pr_info_var"):
            pr = self.pr_cache
            self.pr_info_var.set(
                f"  #{pr['number']}  {pr['title']}   "
                f"{pr['state'].upper()}   by {pr['user']['login']}")
        if self.pr_files and hasattr(self, "sql_btn"):
            self.sql_btn.config(state="normal")
        for k, v in state.get("field_values", {}).items():
            if k in self._field_widgets:
                self._field_widgets[k].set(v)
        if state.get("description") and hasattr(self, "desc_text"):
            self.desc_text.delete("1.0", "end")
            self.desc_text.insert("1.0", state["description"])
            self.desc_text.config(fg=C["text"])

    # ── UI root ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── header bar ────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=C["surface"])
        hdr.pack(fill="x")
        tk.Frame(hdr, bg=C["accent"], width=4).pack(side="left", fill="y")

        # logo / title
        title_f = tk.Frame(hdr, bg=C["surface"])
        title_f.pack(side="left", padx=(12, 0), pady=8)

        # canvas-drawn prism icon
        ico = tk.Canvas(title_f, width=38, height=38, bg=C["surface"],
                        highlightthickness=0)
        ico.pack(side="left", padx=(0, 8))
        # dark rounded bg
        ico.create_oval(2, 2, 36, 36, fill=C["card"], outline=C["border"], width=1)
        # prism triangle
        pts = [10, 6,  10, 32,  33, 19]
        ico.create_polygon(pts, fill=C["accent"], outline="#aad4ff", width=1)
        # refracted rays
        for dy, col in [(-8, C["green"]), (0, C["accent"]), (8, C["yellow"])]:
            ico.create_line(33, 19, 37, 19 + dy, fill=col, width=2)

        # "PRism" — "PR" in accent, "ism" in text
        name_f = tk.Frame(title_f, bg=C["surface"])
        name_f.pack(side="left")
        tk.Label(name_f, text="PR",
                 bg=C["surface"], fg=C["accent"],
                 font=("Segoe UI", 16, "bold")).pack(side="left")
        tk.Label(name_f, text="ism",
                 bg=C["surface"], fg=C["text"],
                 font=("Segoe UI", 16, "bold")).pack(side="left")
        tk.Label(name_f, text=f"  —  {APP_SUBTITLE}",
                 bg=C["surface"], fg=C["muted"],
                 font=("Segoe UI", 9)).pack(side="left", pady=(4, 0))

        # right-side controls
        ctrl_f = tk.Frame(hdr, bg=C["surface"])
        ctrl_f.pack(side="right", padx=12, pady=10)

        theme_lbl = "   Light Mode   " if self._theme_name == "dark" else "   Dark Mode   "
        theme_ico = "☀" if self._theme_name == "dark" else "\U0001f319"
        tk.Button(ctrl_f, text=f"{theme_ico}{theme_lbl}",
                  bg=C["purple"], fg="#ffffff",
                  relief="flat", font=("Segoe UI", 9, "bold"),
                  padx=8, pady=5, cursor="hand2",
                  activebackground="#9966cc", activeforeground="#ffffff",
                  command=self._toggle_theme).pack(side="right", padx=(6, 0))

        for txt, cmd, bg, fg in [
            (" ⚡ Manage Fields ", self._open_field_mgr, C["purple"],  "#ffffff"),
            ("  ⚙ Settings  ",    self._open_settings,  C["orange"],  "#0d1117"),
        ]:
            tk.Button(ctrl_f, text=txt, bg=bg, fg=fg, relief="flat",
                      font=("Segoe UI", 9, "bold"), padx=8, pady=5,
                      activebackground=C["border"], activeforeground=C["text"],
                      cursor="hand2", command=cmd).pack(side="right", padx=3)

        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")

        # ── main content ──────────────────────────────────────────────────────
        self.tab_create = ttk.Frame(self)
        self.tab_create.pack(fill="both", expand=True)

        self._build_create_tab()

        # ── status bar ────────────────────────────────────────────────────────
        self.status_var = tk.StringVar(value="Ready  —  fetch a PR to begin")
        sbar = tk.Frame(self, bg=C["surface"])
        sbar.pack(fill="x", side="bottom")
        tk.Frame(sbar, bg=C["border"], height=1).pack(fill="x")
        sbar_inner = tk.Frame(sbar, bg=C["surface"])
        sbar_inner.pack(fill="x")
        tk.Label(sbar_inner, text="●",
                 bg=C["surface"], fg=C["green"],
                 font=("Segoe UI", 8)).pack(side="left", padx=(10, 4), pady=5)
        tk.Label(sbar_inner, textvariable=self.status_var,
                 bg=C["surface"], fg=C["muted"],
                 font=("Segoe UI", 9)).pack(side="left")
        tk.Label(sbar_inner,
                 text=f"  {self._theme_name.capitalize()} mode",
                 bg=C["surface"], fg=C["muted"],
                 font=("Segoe UI", 8, "italic")).pack(side="right", padx=12, pady=5)

    # ── scrollable frame helper ───────────────────────────────────────────────

    def _scrollable(self, parent):
        c = tk.Canvas(parent, bg=C["bg"], highlightthickness=0)
        v = ttk.Scrollbar(parent, orient="vertical", command=c.yview)
        c.configure(yscrollcommand=v.set)
        v.pack(side="right", fill="y")
        c.pack(side="left", fill="both", expand=True)
        f = tk.Frame(c, bg=C["bg"])
        w = c.create_window((0, 0), window=f, anchor="nw")
        c.bind("<Configure>", lambda e: c.itemconfig(w, width=e.width))
        f.bind("<Configure>", lambda e: c.configure(scrollregion=(0, 0, e.width, e.height)))
        c.bind_all("<MouseWheel>",
                   lambda e: c.yview_scroll(int(-1*(e.delta/120)), "units"))
        return f

    def _card(self, parent, title="", icon="", accent_color=None):
        accent_color = accent_color or C["accent"]
        border_f = tk.Frame(parent, bg=C["border"])
        border_f.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        inner_wrap = tk.Frame(border_f, bg=C["card"])
        inner_wrap.pack(fill="both", expand=True, padx=1, pady=1)
        accent_strip = tk.Frame(inner_wrap, bg=accent_color, width=3)
        accent_strip.pack(side="left", fill="y")
        content = tk.Frame(inner_wrap, bg=C["card"])
        content.pack(side="left", fill="both", expand=True, padx=12, pady=10)
        if title:
            th = tk.Frame(content, bg=C["card"])
            th.pack(fill="x", pady=(0, 8))
            if icon:
                tk.Label(th, text=icon, bg=C["card"], fg=accent_color,
                         font=("Segoe UI", 11)).pack(side="left", padx=(0, 6))
            tk.Label(th, text=title, bg=C["card"], fg=C["text"],
                     font=("Segoe UI", 10, "bold")).pack(side="left")
            tk.Frame(content, bg=C["border"], height=1).pack(fill="x", pady=(0, 10))
        return content

    # ── Create Ticket tab ────────────────────────────────────────────────────

    def _build_create_tab(self):
        outer = tk.Frame(self.tab_create, bg=C["bg"])
        outer.pack(fill="both", expand=True)

        # ── Right sidebar: actions ────────────────────────────────────────────
        sidebar = tk.Frame(outer, bg=C["surface"], width=190)
        sidebar.pack(side="right", fill="y")
        sidebar.pack_propagate(False)
        tk.Frame(sidebar, bg=C["border"], width=1).pack(side="left", fill="y")
        rp = tk.Frame(sidebar, bg=C["surface"])
        rp.pack(side="left", fill="both", expand=True, padx=12, pady=14)

        tk.Label(rp, text="ACTIONS", bg=C["surface"], fg=C["muted"],
                 font=("Segoe UI", 7, "bold")).pack(anchor="w", pady=(0, 10))

        prev_btn = ttk.Button(rp, text="Preview Doc", style="Orange.TButton",
                              command=self._show_preview)
        prev_btn.pack(fill="x", pady=(0, 6))
        Tooltip(prev_btn, "Add change comments and generate the Word document.")

        self.sql_btn = ttk.Button(rp, text="SQL PR File", style="Accent.TButton",
                                  state="disabled", command=self._open_sql_popup)
        self.sql_btn.pack(fill="x", pady=(0, 6))
        Tooltip(self.sql_btn, "Combine all PR SQL files into a single runnable script.")

        jira_btn = ttk.Button(rp, text="Create Jira Ticket", style="Success.TButton",
                              command=self._create_jira_thread)
        jira_btn.pack(fill="x", pady=(0, 14))
        Tooltip(jira_btn, "Create Jira ticket and attach Word doc + SQL file.")

        tk.Frame(rp, bg=C["border"], height=1).pack(fill="x", pady=(0, 10))

        self.progress = ttk.Progressbar(rp, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 8))

        self.result_var = tk.StringVar()
        rl = tk.Label(rp, textvariable=self.result_var,
                      bg=C["surface"], fg=C["green"],
                      font=("Segoe UI", 8, "bold"), wraplength=160,
                      justify="left", cursor="hand2")
        rl.pack(anchor="w")
        rl.bind("<Button-1>", lambda e: self._jira_url and webbrowser.open(self._jira_url))

        # ── Left content pane ─────────────────────────────────────────────────
        lp = tk.Frame(outer, bg=C["bg"])
        lp.pack(side="left", fill="both", expand=True)

        def _sep():
            tk.Frame(lp, bg=C["border"], height=1).pack(fill="x", padx=12, pady=5)

        def _sh(parent, text):
            tk.Label(parent, text=text, bg=C["bg"], fg=C["accent"],
                     font=("Segoe UI", 7, "bold")).pack(anchor="w", padx=14, pady=(6, 2))

        def _frow(parent, fd):
            key = fd["key"]; label = fd["label"]
            ftype = fd.get("type","text"); choices = fd.get("choices",[])
            req = fd.get("required", False)
            if key == "project":
                default = self.config_data.get("jira_project_key", fd.get("default",""))
            else:
                default = fd.get("default","")
            fr = tk.Frame(parent, bg=C["bg"]); fr.pack(fill="x", padx=12, pady=2)
            lf = tk.Frame(fr, bg=C["bg"], width=115); lf.pack(side="left", fill="y")
            lf.pack_propagate(False)
            lbl_w = tk.Label(lf, text=f"{label}{'*' if req else ''}",
                             bg=C["bg"], fg=C["text"],
                             font=("Segoe UI", 9, "bold" if req else "normal"), anchor="w")
            lbl_w.pack(anchor="w", pady=3)
            hint = FIELD_HINTS.get(key, "")
            if hint:
                Tooltip(lbl_w, hint)
            var = tk.StringVar(value=default)
            if ftype == "dropdown" and choices:
                w = ttk.Combobox(fr, textvariable=var, values=choices, state="readonly")
            else:
                w = ttk.Entry(fr, textvariable=var)
            w.pack(side="left", fill="x", expand=True)
            if key == "doc_name":
                tk.Button(fr, text="Browse", bg=C["purple"], fg="#ffffff",
                          relief="flat", font=("Segoe UI", 8, "bold"), padx=6, pady=2,
                          activebackground="#9966cc", cursor="hand2",
                          command=self._choose_doc_dir).pack(side="left", padx=(4, 0))
            if ftype == "dropdown":
                def _edit(fd=fd, cb=w, v=var):
                    self._edit_dropdown_choices(fd, cb, v)
                tk.Button(fr, text="✎", bg=C["bg"], fg=C["accent"], relief="flat",
                          font=("Segoe UI", 9), padx=4, activebackground=C["border"],
                          cursor="hand2", command=_edit).pack(side="left", padx=(2, 0))
            return var

        # ── PR URL ────────────────────────────────────────────────────────────
        _sh(lp, "GITHUB PULL REQUEST")
        pr_row = tk.Frame(lp, bg=C["bg"]); pr_row.pack(fill="x", padx=12, pady=(0, 3))
        lf = tk.Frame(pr_row, bg=C["bg"], width=115); lf.pack(side="left", fill="y")
        lf.pack_propagate(False)
        lbl = tk.Label(lf, text="PR URL / Number", bg=C["bg"], fg=C["text"],
                       font=("Segoe UI", 9), anchor="w")
        lbl.pack(anchor="w", pady=3)
        Tooltip(lbl, "Paste full GitHub PR URL or just the number (e.g. 42)")
        self.pr_url_var = tk.StringVar()
        ttk.Entry(pr_row, textvariable=self.pr_url_var).pack(
            side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(pr_row, text="Fetch PR", style="Accent.TButton",
                   command=self._fetch_pr_thread).pack(side="left", padx=(0, 4))
        ttk.Button(pr_row, text="⚗ Test Data", style="Orange.TButton",
                   command=self._load_mock_data).pack(side="left")

        info_f = tk.Frame(lp, bg=C["bg"]); info_f.pack(fill="x", padx=14, pady=(0, 2))
        self.pr_info_var = tk.StringVar(
            value="No PR loaded — enter a PR URL or number above")
        tk.Label(info_f, textvariable=self.pr_info_var, bg=C["bg"], fg=C["muted"],
                 font=("Segoe UI", 8), anchor="w").pack(anchor="w")
        self.files_sum_var = tk.StringVar(value="")
        tk.Label(info_f, textvariable=self.files_sum_var, bg=C["bg"], fg=C["green"],
                 font=("Segoe UI", 8, "bold"), anchor="w").pack(anchor="w")

        _sep()

        # ── Description (compact) ─────────────────────────────────────────────
        _sh(lp, "DESCRIPTION")
        self.desc_text = scrolledtext.ScrolledText(
            lp, height=3, bg=C["inp"], fg=C["muted"],
            insertbackground=C["text"], relief="flat",
            font=("Segoe UI", 9), wrap="word", selectbackground=C["hover"],
            highlightthickness=1, highlightbackground=C["border"],
            highlightcolor=C["accent"])
        self.desc_text.pack(fill="x", padx=12, pady=(0, 4))
        self.desc_text.insert("1.0", "Describe the issue in detail...")
        def _clr(e):
            if self.desc_text.get("1.0", "end-1c") == "Describe the issue in detail...":
                self.desc_text.delete("1.0", "end")
                self.desc_text.config(fg=C["text"])
        self.desc_text.bind("<FocusIn>", _clr)

        _sep()

        # ── Two columns: Jira Fields | Issue Details ──────────────────────────
        mid = tk.Frame(lp, bg=C["bg"]); mid.pack(fill="x", expand=False)
        jf_col = tk.Frame(mid, bg=C["bg"]); jf_col.pack(side="left", fill="both", expand=True)
        tk.Frame(mid, bg=C["border"], width=1).pack(side="left", fill="y", pady=6)
        id_col = tk.Frame(mid, bg=C["bg"]); id_col.pack(side="left", fill="both", expand=True)

        # Jira Fields
        _sh(jf_col, "JIRA FIELDS")
        self._jira_field_widgets = {}
        for fd in [fd for fd in self.config_data["fields"]
                   if fd.get("jira_field") and fd.get("enabled", True)]:
            self._jira_field_widgets[fd["key"]] = _frow(jf_col, fd)

        if "summary" in self._jira_field_widgets:
            def _sanitize(s):
                s = re.sub(r'[\[\](){}<>:"/\\|?*\n\r\t]', '_', s)
                s = re.sub(r'\s+', '_', s.strip())
                return re.sub(r'_+', '_', s).strip('_')[:80] if s else ""
            self._last_auto_doc_name  = _sanitize(self._jira_field_widgets["summary"].get())
            self._last_auto_doc_title = self._jira_field_widgets["summary"].get()
            self._syncing = False
            def _on_sum(*_):
                if getattr(self, "_syncing", False): return
                self._syncing = True
                try:
                    val = self._jira_field_widgets["summary"].get()
                    fw  = getattr(self, "_field_widgets", {})
                    if "doc_name" in fw:
                        cur = fw["doc_name"].get()
                        if cur in (getattr(self,"_last_auto_doc_name",None), ""):
                            new = _sanitize(val)
                            fw["doc_name"].set(new); self._last_auto_doc_name = new
                    if "doc_title" in fw:
                        cur = fw["doc_title"].get()
                        if cur in (getattr(self,"_last_auto_doc_title",None), ""):
                            fw["doc_title"].set(val); self._last_auto_doc_title = val
                finally:
                    self._syncing = False
            self._jira_field_widgets["summary"].trace_add("write", _on_sum)

        # Issue Details
        _sh(id_col, "ISSUE DETAILS")
        self._fields_card_parent = id_col
        self._fields_card = None
        self._rebuild_fields_card()

    # PR fetch card
    def _build_pr_card(self, parent):
        card = self._card(parent, "GitHub Pull Request", "")

        row = tk.Frame(card, bg=C["card"]); row.pack(fill="x", pady=(0, 6))
        lf = tk.Frame(row, bg=C["card"], width=148); lf.pack(side="left")
        lf.pack_propagate(False)
        lbl = tk.Label(lf, text="PR URL or Number",
                       bg=C["card"], fg=C["text"],
                       font=("Segoe UI", 10))
        lbl.pack(anchor="w", pady=4)
        Tooltip(lbl, "Paste the full GitHub PR URL  or  just the PR number (e.g. 42).\n"
                     "GitHub Owner and Repo must be set in Settings if using a number.")
        self.pr_url_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.pr_url_var).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="  Fetch PR  ", style="Accent.TButton",
                   command=self._fetch_pr_thread).pack(side="left", padx=(8, 0))
        ttk.Button(row, text="⚗ Test Data", style="Orange.TButton",
                   command=self._load_mock_data).pack(side="left", padx=(4, 0))

        info_f = tk.Frame(card, bg=C["card"]); info_f.pack(fill="x", pady=(6, 0))
        self.pr_info_var = tk.StringVar(value="No PR loaded — enter a PR URL or number above")
        tk.Label(info_f, textvariable=self.pr_info_var,
                 bg=C["card"], fg=C["muted"],
                 font=("Segoe UI", 9), wraplength=820, justify="left"
                 ).pack(anchor="w")
        self.files_sum_var = tk.StringVar(value="")
        tk.Label(info_f, textvariable=self.files_sum_var,
                 bg=C["card"], fg=C["green"],
                 font=("Segoe UI", 9, "bold"), wraplength=820, justify="left"
                 ).pack(anchor="w")

    # Jira Fields card — fixed set of 4 fields sent directly to Jira API
    def _build_jira_fields_card(self, parent):
        card = self._card(parent, "Jira Fields", "")
        self._jira_field_widgets = {}
        jira_fds = [fd for fd in self.config_data["fields"]
                    if fd.get("jira_field") and fd.get("enabled", True)]

        for fd in jira_fds:
            key     = fd["key"]
            label   = fd["label"]
            ftype   = fd.get("type", "text")
            choices = fd.get("choices", [])
            req     = fd.get("required", False)

            if key == "project":
                default = self.config_data.get("jira_project_key", fd.get("default", ""))
            else:
                default = fd.get("default", "")

            row = tk.Frame(card, bg=C["card"]); row.pack(fill="x", pady=(0, 5))
            lf  = tk.Frame(row, bg=C["card"], width=130); lf.pack(side="left", fill="y")
            lf.pack_propagate(False)
            req_badge = " *" if req else ""
            tk.Label(lf, text=f"{label}{req_badge}",
                     bg=C["card"], fg=C["text"],
                     font=("Segoe UI", 9, "bold" if req else "normal"),
                     anchor="w").pack(side="left", padx=(0, 4), pady=4)

            var = tk.StringVar(value=default)
            if ftype == "dropdown" and choices:
                w = ttk.Combobox(row, textvariable=var, values=choices, state="readonly")
            else:
                w = ttk.Entry(row, textvariable=var)
            w.pack(side="left", fill="x", expand=True)

            if ftype == "dropdown":
                def _edit_jf(fd=fd, cb=w, v=var):
                    self._edit_dropdown_choices(fd, cb, v)
                tk.Button(row, text="✎", bg=C["surface"], fg=C["accent"],
                          relief="flat", font=("Segoe UI", 10), padx=6, pady=2,
                          activebackground=C["border"], cursor="hand2",
                          command=_edit_jf
                          ).pack(side="left", padx=(4, 0))

            self._jira_field_widgets[key] = var

        # auto-sync: Jira Summary → Doc File Name + Doc Title
        if "summary" in self._jira_field_widgets:
            def _sanitize_fname(s):
                s = re.sub(r'[\[\](){}<>:"/\\|?*\n\r\t]', '_', s)
                s = re.sub(r'\s+', '_', s.strip())
                s = re.sub(r'_+', '_', s).strip('_')
                return s[:80] if s else ""

            _cur_sum = self._jira_field_widgets["summary"].get()
            self._last_auto_doc_name  = _sanitize_fname(_cur_sum)
            self._last_auto_doc_title = _cur_sum
            self._syncing = False

            def _on_summary_change(*_):
                if getattr(self, "_syncing", False):
                    return
                self._syncing = True
                try:
                    sv = self._jira_field_widgets.get("summary")
                    if sv is None:
                        return
                    val = sv.get()
                    fw = getattr(self, "_field_widgets", {})
                    if "doc_name" in fw:
                        cur = fw["doc_name"].get()
                        if cur == getattr(self, "_last_auto_doc_name", None) or cur == "":
                            new = _sanitize_fname(val)
                            fw["doc_name"].set(new)
                            self._last_auto_doc_name = new
                    if "doc_title" in fw:
                        cur = fw["doc_title"].get()
                        if cur == getattr(self, "_last_auto_doc_title", None) or cur == "":
                            fw["doc_title"].set(val)
                            self._last_auto_doc_title = val
                finally:
                    self._syncing = False

            self._jira_field_widgets["summary"].trace_add("write", _on_summary_change)

    # Dynamic fields card — rebuilt whenever field definitions change
    def _build_dynamic_fields_card(self, parent):
        self._fields_card_parent = parent
        self._rebuild_fields_card()

    def _rebuild_fields_card(self):
        if hasattr(self, "_fields_card") and self._fields_card:
            self._fields_card.destroy()

        flat = tk.Frame(self._fields_card_parent, bg=C["bg"])
        flat.pack(fill="x")
        self._fields_card = flat

        self._field_widgets = {}
        enabled = [fd for fd in self.config_data["fields"]
                   if fd.get("enabled", True) and not fd.get("jira_field")]

        for fd in enabled:
            key = fd["key"]; label = fd["label"]
            ftype = fd.get("type","text"); choices = fd.get("choices",[])
            default = fd.get("default",""); req = fd.get("required", False)

            row = tk.Frame(flat, bg=C["bg"]); row.pack(fill="x", padx=12, pady=2)
            lf  = tk.Frame(row, bg=C["bg"], width=115); lf.pack(side="left", fill="y")
            lf.pack_propagate(False)

            lbl_w = tk.Label(lf, text=f"{label}{'*' if req else ''}",
                             bg=C["bg"], fg=C["text"],
                             font=("Segoe UI", 9, "bold" if req else "normal"),
                             anchor="w")
            lbl_w.pack(anchor="w", pady=3)
            hint = FIELD_HINTS.get(key, "")
            if hint:
                Tooltip(lbl_w, hint)

            var = tk.StringVar(value=default)
            if ftype == "dropdown" and choices:
                w = ttk.Combobox(row, textvariable=var, values=choices, state="readonly")
            else:
                w = ttk.Entry(row, textvariable=var)
            w.pack(side="left", fill="x", expand=True)

            if key == "doc_name":
                tk.Button(row, text="Browse", bg=C["purple"], fg="#ffffff",
                          relief="flat", font=("Segoe UI", 8, "bold"), padx=6, pady=2,
                          activebackground="#9966cc", cursor="hand2",
                          command=self._choose_doc_dir).pack(side="left", padx=(4, 0))

            if ftype == "dropdown":
                def _edit_choices(fd=fd, cb=w, v=var):
                    self._edit_dropdown_choices(fd, cb, v)
                tk.Button(row, text="✎", bg=C["bg"], fg=C["accent"],
                          relief="flat", font=("Segoe UI", 9), padx=4,
                          activebackground=C["border"], cursor="hand2",
                          command=_edit_choices).pack(side="left", padx=(2, 0))

            self._field_widgets[key] = var


    def _edit_dropdown_choices(self, fd, combobox, var):
        dlg = tk.Toplevel(self)
        dlg.title(f"Edit choices — {fd['label']}")
        dlg.geometry("380x340")
        dlg.configure(bg=C["bg"])
        dlg.grab_set()
        dlg.transient(self)

        tk.Label(dlg, text=f"Choices for  \"{fd['label']}\"",
                 bg=C["bg"], fg=C["accent"],
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=16, pady=(14, 4))
        tk.Label(dlg, text="One value per line",
                 bg=C["bg"], fg=C["muted"],
                 font=("Segoe UI", 8)).pack(anchor="w", padx=16, pady=(0, 6))

        txt = scrolledtext.ScrolledText(dlg, height=10,
                                        bg=C["inp"], fg=C["text"],
                                        insertbackground=C["text"],
                                        relief="flat", font=("Segoe UI", 10),
                                        highlightthickness=1,
                                        highlightbackground=C["border"])
        txt.pack(fill="both", expand=True, padx=16, pady=(0, 10))
        txt.insert("1.0", "\n".join(fd.get("choices", [])))

        def _save():
            lines = [l.strip() for l in txt.get("1.0", "end-1c").splitlines() if l.strip()]
            fd["choices"] = lines
            if combobox:
                combobox.config(values=lines)
            # keep current value if still valid, else reset to first
            if var.get() not in lines and lines:
                var.set(lines[0])
            save_config(self.config_data)
            dlg.destroy()

        br = tk.Frame(dlg, bg=C["bg"]); br.pack(fill="x", padx=16, pady=(0, 14))
        tk.Button(br, text="Save", bg=C["green"], fg="#0d1117", relief="flat",
                  font=("Segoe UI", 10, "bold"), padx=14, pady=5,
                  cursor="hand2", command=_save).pack(side="right")
        tk.Button(br, text="Cancel", bg=C["red"], fg="#ffffff", relief="flat",
                  font=("Segoe UI", 10), padx=12, pady=5,
                  activebackground="#c0392b", cursor="hand2",
                  command=dlg.destroy).pack(side="right", padx=(0, 8))

    def _build_desc_card(self, parent):
        card = self._card(parent, "Description", "", accent_color=C["muted"])
        self.desc_text = scrolledtext.ScrolledText(
            card, height=7, bg=C["inp"], fg=C["muted"],
            insertbackground=C["text"], relief="flat",
            font=("Segoe UI", 10), wrap="word", selectbackground=C["hover"],
            highlightthickness=1, highlightbackground=C["border"],
            highlightcolor=C["accent"])
        self.desc_text.pack(fill="x")
        self.desc_text.insert("1.0", "Describe the issue in detail...")

        def _clr(e):
            if self.desc_text.get("1.0", "end-1c") == "Describe the issue in detail...":
                self.desc_text.delete("1.0", "end")
                self.desc_text.config(fg=C["text"])
        self.desc_text.bind("<FocusIn>", _clr)

    def _build_actions_card(self, parent):
        card = self._card(parent, "Actions", "", accent_color=C["green"])

        # status label + progress bar
        self.result_var = tk.StringVar()
        rl = tk.Label(card, textvariable=self.result_var,
                      bg=C["card"], fg=C["green"],
                      font=("Segoe UI", 9, "bold"), cursor="hand2",
                      wraplength=400, justify="left")
        rl.pack(anchor="w", pady=(0, 6))
        rl.bind("<Button-1>", lambda e: self._jira_url and webbrowser.open(self._jira_url))

        self.progress = ttk.Progressbar(card, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 16))

        # buttons — stacked full-width, right-aligned in the half-width column
        prev_btn = ttk.Button(card, text="  Preview Doc  ", style="Orange.TButton",
                              command=self._show_preview)
        prev_btn.pack(fill="x", pady=(0, 6))
        Tooltip(prev_btn, "Add change comments then generate the Word document.")

        self.sql_btn = ttk.Button(card, text="  SQL PR File  ", style="Accent.TButton",
                                  state="disabled", command=self._open_sql_popup)
        self.sql_btn.pack(fill="x", pady=(0, 6))
        Tooltip(self.sql_btn, "Combine all PR SQL files into a single runnable script.")

        jira_btn = ttk.Button(card, text="  Create Jira Ticket  ", style="Success.TButton",
                              command=self._create_jira_thread)
        jira_btn.pack(fill="x")
        Tooltip(jira_btn, "Create a Jira ticket and attach Word doc + SQL file.")

    # ── Settings ─────────────────────────────────────────────────────────────


    # ── settings window ───────────────────────────────────────────────────────

    def _open_settings(self):
        win = tk.Toplevel(self)
        win.title(f"{APP_NAME}  —  Settings")
        win.geometry("640x540")
        win.configure(bg=C["bg"])
        win.grab_set()

        tk.Label(win, text="Connection Settings",
                 bg=C["bg"], fg=C["accent"],
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=24, pady=(18, 14))

        c = tk.Canvas(win, bg=C["bg"], highlightthickness=0)
        v = ttk.Scrollbar(win, orient="vertical", command=c.yview)
        c.configure(yscrollcommand=v.set)
        v.pack(side="right", fill="y")
        c.pack(side="left", fill="both", expand=True)
        sf  = tk.Frame(c, bg=C["bg"])
        wid = c.create_window((0, 0), window=sf, anchor="nw")
        c.bind("<Configure>", lambda e: c.itemconfig(wid, width=e.width))
        sf.bind("<Configure>", lambda e: c.configure(scrollregion=c.bbox("all")))

        defs = [
            ("GitHub Token File",   "github_token_file"),
            ("Jira Token File",     "jira_token_file"),
            ("Jira Base URL",       "jira_base_url"),
            ("Jira Project Key",    "jira_project_key"),
            ("Jira Email",          "jira_email"),
            ("GitHub Owner",        "github_owner"),
            ("GitHub Repo",         "github_repo"),
            ("Output Directory",    "word_doc_output_dir"),
        ]
        vars_ = {}
        for label, key in defs:
            row = tk.Frame(sf, bg=C["bg"]); row.pack(fill="x", padx=20, pady=5)
            tk.Label(row, text=label, bg=C["bg"], fg=C["muted"],
                     width=22, anchor="w", font=("Segoe UI", 10)).pack(side="left")
            v2 = tk.StringVar(value=str(self.config_data.get(key, "")))
            vars_[key] = v2
            ttk.Entry(row, textvariable=v2).pack(side="left", fill="x", expand=True)
            if "file" in key or "dir" in key.lower():
                is_dir = "dir" in key.lower()
                def _browse(var=v2, d=is_dir):
                    p = filedialog.askdirectory() if d else \
                        filedialog.askopenfilename(
                            filetypes=[("Text","*.txt"),("All","*.*")])
                    if p: var.set(p)
                ttk.Button(row, text="...", style="Purple.TButton",
                           command=_browse).pack(side="left", padx=(6, 0))

        def _save():
            for k, v2 in vars_.items():
                self.config_data[k] = v2.get()
            save_config(self.config_data)
            self._set_status("Settings saved")
            win.destroy()

        br = tk.Frame(sf, bg=C["bg"]); br.pack(fill="x", padx=20, pady=20)
        ttk.Button(br, text="Save Settings", style="Success.TButton",
                   command=_save).pack(side="right")
        ttk.Button(br, text="Cancel", style="Danger.TButton",
                   command=win.destroy).pack(side="right", padx=(0, 8))

    # ── Field manager ─────────────────────────────────────────────────────────

    def _open_field_mgr(self):
        dlg = FieldManagerDialog(self, self.config_data["fields"])
        self.wait_window(dlg)
        if dlg.result is not None:
            self.config_data["fields"] = dlg.result
            # persist defaults
            self.config_data["field_defaults"] = {
                fd["key"]: fd["default"] for fd in dlg.result}
            save_config(self.config_data)
            self._rebuild_fields_card()
            self._set_status("Fields updated — defaults saved")

    # ── Fetch PR ──────────────────────────────────────────────────────────────

    def _fetch_pr_thread(self):
        threading.Thread(target=self._fetch_pr, daemon=True).start()

    def _fetch_pr(self):
        self.after(0, self.progress.start)
        self.after(0, lambda: self._set_status("Fetching PR from GitHub..."))
        try:
            token = read_token(self.config_data["github_token_file"])
        except FileNotFoundError as e:
            self.after(0, lambda m=str(e): messagebox.showerror("Token Error", m))
            self.after(0, self.progress.stop); return

        raw   = self.pr_url_var.get().strip()
        owner = self.config_data.get("github_owner", "")
        repo  = self.config_data.get("github_repo", "")
        num   = None

        if raw.isdigit():
            num = raw
        elif "github.com" in raw:
            parts = raw.rstrip("/").split("/")
            try:
                idx = parts.index("pull")
                num = parts[idx+1]; owner = parts[idx-2]; repo = parts[idx-1]
            except (ValueError, IndexError):
                pass

        if not num:
            self.after(0, lambda: messagebox.showerror(
                "Error", "Enter a PR number or full GitHub PR URL"))
            self.after(0, self.progress.stop); return
        if not owner or not repo:
            self.after(0, lambda: messagebox.showerror(
                "Error", "Set GitHub Owner and Repo in Settings"))
            self.after(0, self.progress.stop); return

        hdrs = gh_headers(token)
        try:
            pr_r = requests.get(
                f"https://api.github.com/repos/{owner}/{repo}/pulls/{num}",
                headers=hdrs, timeout=15)
            pr_r.raise_for_status()
            pr   = pr_r.json()

            files_r = requests.get(
                f"https://api.github.com/repos/{owner}/{repo}/pulls/{num}/files?per_page=100",
                headers=hdrs, timeout=15)
            files_r.raise_for_status()
            files = files_r.json()
        except requests.HTTPError as e:
            msg = f"GitHub API error {e.response.status_code}:\n{e.response.text[:300]}"
            self.after(0, lambda m=msg: messagebox.showerror("GitHub Error", m))
            self.after(0, self.progress.stop); return
        except Exception as e:
            self.after(0, lambda m=str(e): messagebox.showerror("Network Error", m))
            self.after(0, self.progress.stop); return

        self.pr_cache = pr; self.pr_files = files; self.file_comments = []

        def _ui():
            self.pr_info_var.set(
                f"  #{pr['number']}  {pr['title']}   "
                f"{pr['state'].upper()}   by {pr['user']['login']}   "
                f"{(pr.get('created_at') or '')[:10]}")
            # auto-fill fields from PR
            if "git_link" in self._field_widgets:
                self._field_widgets["git_link"].set(pr.get("html_url", ""))
            if "summary" in self._field_widgets and not self._field_widgets["summary"].get():
                itype = self._field_widgets.get("issue_type",
                                                tk.StringVar(value="")).get()
                self._field_widgets["summary"].set(f"[{itype}] {pr['title']}")
            self.pr_url_var.set(pr.get("html_url", ""))

            total = len(files)
            adds  = sum(1 for f in files if f.get("status")=="added")
            mods  = sum(1 for f in files if f.get("status")=="modified")
            rems  = sum(1 for f in files if f.get("status")=="removed")
            other = total - adds - mods - rems
            parts = [f"{total} file{'s' if total!=1 else ''} changed"]
            if adds:  parts.append(f"A {adds} added")
            if mods:  parts.append(f"M {mods} modified")
            if rems:  parts.append(f"D {rems} removed")
            if other: parts.append(f"~ {other} other")
            self.files_sum_var.set("   ".join(parts))
            if hasattr(self, "sql_btn"):
                self.sql_btn.config(state="normal")
            self.progress.stop()
            self._set_status(f"PR #{pr['number']} loaded — {total} files")

        self.after(0, _ui)

    # ── File changes ──────────────────────────────────────────────────────────

    def _open_sql_popup(self):
        if not self.pr_files:
            messagebox.showinfo("No PR loaded", "Fetch a PR first.", parent=self)
            return
        dlg = SqlFilePopup(self, self.pr_files, self.pr_cache or {}, self.config_data)
        self.wait_window(dlg)
        result = dlg.get_result()
        if result and os.path.exists(result):
            self._last_sql_path = result

    # ── Mock / test data ──────────────────────────────────────────────────────

    def _load_mock_data(self):
        self.pr_cache = {
            "number":     42,
            "title":      "feat: inventory & order management SQL overhaul",
            "state":      "open",
            "html_url":   "https://github.com/example/db-repo/pull/42",
            "created_at": "2026-05-10T09:00:00Z",
            "updated_at": "2026-05-19T14:00:00Z",
            "user":       {"login": "anantharuban"},
            "head":       {"ref": "feature/sql-overhaul"},
            "base":       {"ref": "main"},
            "body": (
                "## Summary\n"
                "- Refactored `v_customer_summary` view to include loyalty tier\n"
                "- Updated `sp_update_order_status` procedure for new workflow states\n"
                "- Added new `pkg_inventory_mgmt` package for stock tracking\n"
                "- Removed deprecated `v_order_summary_old` view (replaced by pkg_inventory_mgmt)\n\n"
                "## Test plan\n"
                "- [ ] Run regression on order status transitions\n"
                "- [ ] Verify inventory deduction on new orders\n"
                "- [ ] Check view returns correct loyalty tier labels\n"
                "- [ ] Confirm no dependencies on removed view"
            ),
        }

        # ── File 1: SMALL change — view, 3 additions / 2 deletions ─────────────
        patch_small = (
            "@@ -1,14 +1,15 @@\n"
            " CREATE OR REPLACE VIEW v_customer_summary AS\n"
            " SELECT\n"
            "     c.customer_id,\n"
            "     c.full_name,\n"
            "     c.email,\n"
            "-    c.created_date,\n"
            "-    COUNT(o.order_id)  AS total_orders\n"
            "+    c.created_date,\n"
            "+    COUNT(o.order_id)     AS total_orders,\n"
            "+    NVL(lp.tier, 'BRONZE') AS loyalty_tier\n"
            " FROM customers c\n"
            " LEFT JOIN orders o  ON o.customer_id = c.customer_id\n"
            "+LEFT JOIN loyalty_points lp ON lp.customer_id = c.customer_id\n"
            " WHERE c.status = 'ACTIVE'\n"
            " GROUP BY\n"
            "     c.customer_id, c.full_name, c.email,\n"
            "-    c.created_date;\n"
            "+    c.created_date, lp.tier;\n"
        )

        # ── File 2: MEDIUM change — stored procedure, 18 add / 8 del ──────────
        patch_medium = (
            "@@ -1,6 +1,8 @@\n"
            " CREATE OR REPLACE PROCEDURE sp_update_order_status (\n"
            "     p_order_id   IN  orders.order_id%TYPE,\n"
            "     p_new_status IN  VARCHAR2,\n"
            "+    p_updated_by IN  VARCHAR2 DEFAULT USER,\n"
            "+    p_notes      IN  VARCHAR2 DEFAULT NULL,\n"
            "     p_result     OUT VARCHAR2\n"
            " ) AS\n"
            "@@ -12,18 +14,26 @@\n"
            "     v_current_status  VARCHAR2(30);\n"
            "     v_allowed_next    VARCHAR2(200);\n"
            "+    v_audit_id        audit_log.audit_id%TYPE;\n"
            " BEGIN\n"
            "     SELECT status\n"
            "     INTO   v_current_status\n"
            "     FROM   orders\n"
            "     WHERE  order_id = p_order_id;\n"
            " \n"
            "-    IF p_new_status NOT IN ('PENDING','PROCESSING','SHIPPED','DELIVERED','CANCELLED') THEN\n"
            "+    IF p_new_status NOT IN ('PENDING','PROCESSING','PICKED','SHIPPED',\n"
            "+                            'OUT_FOR_DELIVERY','DELIVERED','CANCELLED','RETURNED') THEN\n"
            "         p_result := 'ERROR: invalid status';\n"
            "         RETURN;\n"
            "     END IF;\n"
            " \n"
            "-    UPDATE orders SET status = p_new_status WHERE order_id = p_order_id;\n"
            "+    UPDATE orders\n"
            "+    SET    status      = p_new_status,\n"
            "+           updated_by  = p_updated_by,\n"
            "+           updated_date = SYSDATE\n"
            "+    WHERE  order_id = p_order_id;\n"
            " \n"
            "+    INSERT INTO order_status_history (order_id, old_status, new_status,\n"
            "+                                      changed_by, changed_date, notes)\n"
            "+    VALUES (p_order_id, v_current_status, p_new_status,\n"
            "+            p_updated_by, SYSDATE, p_notes);\n"
            "+\n"
            "-    COMMIT;\n"
            "+    COMMIT;\n"
            "     p_result := 'OK';\n"
            " EXCEPTION\n"
            "     WHEN OTHERS THEN\n"
            "-        ROLLBACK;\n"
            "+        ROLLBACK;\n"
            "         p_result := 'ERROR: ' || SQLERRM;\n"
            " END sp_update_order_status;\n"
            " /\n"
        )

        # ── File 3: BIG change — new package, 72 additions ─────────────────────
        patch_big = (
            "@@ -0,0 +1,72 @@\n"
            "+CREATE OR REPLACE PACKAGE pkg_inventory_mgmt AS\n"
            "+\n"
            "+    -- Reserve stock for an order line\n"
            "+    PROCEDURE reserve_stock (\n"
            "+        p_product_id  IN  NUMBER,\n"
            "+        p_qty         IN  NUMBER,\n"
            "+        p_order_id    IN  NUMBER,\n"
            "+        p_result      OUT VARCHAR2\n"
            "+    );\n"
            "+\n"
            "+    -- Release reserved stock (cancellation / rejection)\n"
            "+    PROCEDURE release_stock (\n"
            "+        p_product_id  IN  NUMBER,\n"
            "+        p_qty         IN  NUMBER,\n"
            "+        p_order_id    IN  NUMBER,\n"
            "+        p_result      OUT VARCHAR2\n"
            "+    );\n"
            "+\n"
            "+    -- Confirm stock deduction on dispatch\n"
            "+    PROCEDURE confirm_dispatch (\n"
            "+        p_order_id    IN  NUMBER,\n"
            "+        p_result      OUT VARCHAR2\n"
            "+    );\n"
            "+\n"
            "+    -- Returns current available qty (on_hand - reserved)\n"
            "+    FUNCTION get_available_qty (\n"
            "+        p_product_id  IN  NUMBER\n"
            "+    ) RETURN NUMBER;\n"
            "+\n"
            "+END pkg_inventory_mgmt;\n"
            "+/\n"
            "+\n"
            "+CREATE OR REPLACE PACKAGE BODY pkg_inventory_mgmt AS\n"
            "+\n"
            "+    PROCEDURE reserve_stock (p_product_id IN NUMBER, p_qty IN NUMBER,\n"
            "+                             p_order_id IN NUMBER, p_result OUT VARCHAR2) AS\n"
            "+        v_avail NUMBER;\n"
            "+    BEGIN\n"
            "+        SELECT on_hand_qty - reserved_qty\n"
            "+        INTO   v_avail\n"
            "+        FROM   inventory\n"
            "+        WHERE  product_id = p_product_id\n"
            "+        FOR UPDATE;\n"
            "+\n"
            "+        IF v_avail < p_qty THEN\n"
            "+            p_result := 'ERROR: insufficient stock (' || v_avail || ' available)';\n"
            "+            RETURN;\n"
            "+        END IF;\n"
            "+\n"
            "+        UPDATE inventory\n"
            "+        SET    reserved_qty = reserved_qty + p_qty\n"
            "+        WHERE  product_id = p_product_id;\n"
            "+\n"
            "+        INSERT INTO inventory_reservations (product_id, order_id, qty, reserved_date)\n"
            "+        VALUES (p_product_id, p_order_id, p_qty, SYSDATE);\n"
            "+\n"
            "+        COMMIT;\n"
            "+        p_result := 'OK';\n"
            "+    EXCEPTION WHEN OTHERS THEN ROLLBACK; p_result := 'ERROR: ' || SQLERRM;\n"
            "+    END reserve_stock;\n"
            "+\n"
            "+    PROCEDURE release_stock (p_product_id IN NUMBER, p_qty IN NUMBER,\n"
            "+                             p_order_id IN NUMBER, p_result OUT VARCHAR2) AS\n"
            "+    BEGIN\n"
            "+        UPDATE inventory\n"
            "+        SET    reserved_qty = GREATEST(0, reserved_qty - p_qty)\n"
            "+        WHERE  product_id = p_product_id;\n"
            "+\n"
            "+        DELETE FROM inventory_reservations\n"
            "+        WHERE  product_id = p_product_id AND order_id = p_order_id;\n"
            "+\n"
            "+        COMMIT; p_result := 'OK';\n"
            "+    EXCEPTION WHEN OTHERS THEN ROLLBACK; p_result := 'ERROR: ' || SQLERRM;\n"
            "+    END release_stock;\n"
            "+\n"
            "+    PROCEDURE confirm_dispatch (p_order_id IN NUMBER, p_result OUT VARCHAR2) AS\n"
            "+    BEGIN\n"
            "+        UPDATE inventory i\n"
            "+        SET    on_hand_qty  = on_hand_qty  - ir.qty,\n"
            "+               reserved_qty = reserved_qty - ir.qty\n"
            "+        FROM   inventory_reservations ir\n"
            "+        WHERE  ir.product_id = i.product_id AND ir.order_id = p_order_id;\n"
            "+\n"
            "+        DELETE FROM inventory_reservations WHERE order_id = p_order_id;\n"
            "+        COMMIT; p_result := 'OK';\n"
            "+    EXCEPTION WHEN OTHERS THEN ROLLBACK; p_result := 'ERROR: ' || SQLERRM;\n"
            "+    END confirm_dispatch;\n"
            "+\n"
            "+    FUNCTION get_available_qty (p_product_id IN NUMBER) RETURN NUMBER AS\n"
            "+        v_qty NUMBER := 0;\n"
            "+    BEGIN\n"
            "+        SELECT on_hand_qty - NVL(reserved_qty, 0)\n"
            "+        INTO   v_qty FROM inventory WHERE product_id = p_product_id;\n"
            "+        RETURN v_qty;\n"
            "+    EXCEPTION WHEN NO_DATA_FOUND THEN RETURN 0;\n"
            "+    END get_available_qty;\n"
            "+\n"
            "+END pkg_inventory_mgmt;\n"
            "+/\n"
        )

        # ── File 4: DELETED file — deprecated view, 14 deletions ──────────────
        patch_deleted = (
            "@@ -1,14 +0,0 @@\n"
            "-CREATE OR REPLACE VIEW v_order_summary_old AS\n"
            "-SELECT\n"
            "-    o.order_id,\n"
            "-    o.customer_id,\n"
            "-    o.status,\n"
            "-    o.order_total,\n"
            "-    o.created_date,\n"
            "-    o.updated_date,\n"
            "-    c.full_name        AS customer_name,\n"
            "-    c.email            AS customer_email\n"
            "-FROM orders o\n"
            "-JOIN customers c ON c.customer_id = o.customer_id\n"
            "-WHERE o.created_date < ADD_MONTHS(SYSDATE, -24)\n"
            "-  AND o.status IN ('CANCELLED', 'RETURNED');\n"
        )

        # Full post-change file content for modified files (simulates raw_url fetch)
        full_small = (
            "CREATE OR REPLACE VIEW v_customer_summary AS\n"
            "SELECT\n"
            "    c.customer_id,\n"
            "    c.full_name,\n"
            "    c.email,\n"
            "    c.created_date,\n"
            "    COUNT(o.order_id)       AS total_orders,\n"
            "    NVL(lp.tier, 'BRONZE')  AS loyalty_tier\n"
            "FROM customers c\n"
            "LEFT JOIN orders o         ON o.customer_id  = c.customer_id\n"
            "LEFT JOIN loyalty_points lp ON lp.customer_id = c.customer_id\n"
            "WHERE c.status = 'ACTIVE'\n"
            "GROUP BY\n"
            "    c.customer_id, c.full_name, c.email,\n"
            "    c.created_date, lp.tier;\n"
        )

        full_medium = (
            "CREATE OR REPLACE PROCEDURE sp_update_order_status (\n"
            "    p_order_id   IN  orders.order_id%TYPE,\n"
            "    p_new_status IN  VARCHAR2,\n"
            "    p_updated_by IN  VARCHAR2 DEFAULT USER,\n"
            "    p_notes      IN  VARCHAR2 DEFAULT NULL,\n"
            "    p_result     OUT VARCHAR2\n"
            ") AS\n"
            "    v_current_status  VARCHAR2(30);\n"
            "    v_allowed_next    VARCHAR2(200);\n"
            "    v_audit_id        audit_log.audit_id%TYPE;\n"
            "BEGIN\n"
            "    SELECT status\n"
            "    INTO   v_current_status\n"
            "    FROM   orders\n"
            "    WHERE  order_id = p_order_id;\n"
            "\n"
            "    IF p_new_status NOT IN ('PENDING','PROCESSING','PICKED','SHIPPED',\n"
            "                            'OUT_FOR_DELIVERY','DELIVERED','CANCELLED','RETURNED') THEN\n"
            "        p_result := 'ERROR: invalid status';\n"
            "        RETURN;\n"
            "    END IF;\n"
            "\n"
            "    UPDATE orders\n"
            "    SET    status       = p_new_status,\n"
            "           updated_by   = p_updated_by,\n"
            "           updated_date = SYSDATE\n"
            "    WHERE  order_id = p_order_id;\n"
            "\n"
            "    INSERT INTO order_status_history\n"
            "           (order_id, old_status, new_status, changed_by, changed_date, notes)\n"
            "    VALUES (p_order_id, v_current_status, p_new_status,\n"
            "            p_updated_by, SYSDATE, p_notes);\n"
            "\n"
            "    COMMIT;\n"
            "    p_result := 'OK';\n"
            "EXCEPTION\n"
            "    WHEN OTHERS THEN\n"
            "        ROLLBACK;\n"
            "        p_result := 'ERROR: ' || SQLERRM;\n"
            "END sp_update_order_status;\n"
            "/\n"
        )

        self.pr_files = [
            {
                "filename":     "db/views/v_customer_summary.sql",
                "status":       "modified",
                "additions":    3,
                "deletions":    2,
                "changes":      5,
                "patch":        patch_small,
                "full_content": full_small,
                "blob_url":     "",
                "raw_url":      "",
                "contents_url": "",
            },
            {
                "filename":     "db/procedures/sp_update_order_status.sql",
                "status":       "modified",
                "additions":    18,
                "deletions":    8,
                "changes":      26,
                "patch":        patch_medium,
                "full_content": full_medium,
                "blob_url":     "",
                "raw_url":      "",
                "contents_url": "",
            },
            {
                "filename":     "db/packages/pkg_inventory_mgmt.sql",
                "status":       "added",
                "additions":    72,
                "deletions":    0,
                "changes":      72,
                "patch":        patch_big,
                "full_content": None,
                "blob_url":     "",
                "raw_url":      "",
                "contents_url": "",
            },
            {
                "filename":     "db/views/v_order_summary_old.sql",
                "status":       "removed",
                "additions":    0,
                "deletions":    14,
                "changes":      14,
                "patch":        patch_deleted,
                "full_content": None,
                "blob_url":     "",
                "raw_url":      "",
                "contents_url": "",
            },
        ]
        self.file_comments = []

        pr = self.pr_cache
        self.pr_info_var.set(
            f"  #{pr['number']}  {pr['title']}   "
            f"{pr['state'].upper()}   by {pr['user']['login']}   "
            f"2026-05-10  [MOCK DATA]"
        )
        self.files_sum_var.set(
            "4 files changed   A 1 added   M 2 modified   D 1 removed"
        )
        if hasattr(self, "sql_btn"):
            self.sql_btn.config(state="normal")
        if hasattr(self, "git_link") and "git_link" in self._field_widgets:
            self._field_widgets["git_link"].set(pr["html_url"])
        self._set_status("Mock data loaded — 4 SQL files (small / medium / big / deleted)")

    # ── Preview ───────────────────────────────────────────────────────────────

    def _show_preview(self):
        fv = self._collect_fields()
        html_content = build_preview_html(
            fv, self.config_data["fields"], self.pr_cache, self.file_comments)

        popup_ref = [None]

        def _on_generate(updated_comments):
            self.file_comments = updated_comments

            def _after_save(path):
                p = popup_ref[0]
                if p and p.winfo_exists():
                    p._enable_open_word(path)

            self._gen_word_thread(after_save=_after_save)

        def _build_html(fc):
            fv2 = self._collect_fields()
            return build_preview_html(fv2, self.config_data["fields"], self.pr_cache, fc)

        popup = DocPreviewWindow(self, html_content, self._last_doc_path,
                                 pr_files=self.pr_files,
                                 file_comments=self.file_comments,
                                 on_comments_saved=_on_generate,
                                 build_html_func=_build_html)
        popup_ref[0] = popup

    # ── Word doc ─────────────────────────────────────────────────────────────

    def _gen_word_thread(self, after_save=None):
        threading.Thread(target=self._gen_word, args=(after_save,), daemon=True).start()

    def _gen_word(self, after_save=None):
        self.after(0, self.progress.start)
        self.after(0, lambda: self._set_status("Generating Word document..."))
        fv      = self._collect_fields()
        out_dir = self.config_data.get("word_doc_output_dir", BASE_DIR)
        try:
            path = generate_word_doc(fv, self.config_data["fields"],
                                     self.pr_cache, self.file_comments, out_dir)
        except Exception as e:
            self.after(0, lambda m=str(e): messagebox.showerror("Word Error", m))
            self.after(0, self.progress.stop); return
        self._last_doc_path = path
        self.after(0, self.progress.stop)
        self.after(0, lambda: self._set_status("Word document saved"))
        if after_save:
            self.after(0, lambda p=path: after_save(p))

    # ── Jira ticket ───────────────────────────────────────────────────────────

    def _create_jira_thread(self):
        threading.Thread(target=self._create_jira, daemon=True).start()

    def _create_jira(self):
        self.after(0, self.progress.start)
        self.after(0, lambda: self._set_status("Creating Jira ticket..."))

        fv = self._collect_fields()

        # validate required fields
        for fd in self.config_data["fields"]:
            if fd.get("required") and fd.get("enabled", True):
                if not fv.get(fd["key"]):
                    msg = f"{fd['label']} is required"
                    self.after(0, lambda m=msg: messagebox.showerror("Validation", m))
                    self.after(0, self.progress.stop); return

        try:
            jira_tok = read_token(self.config_data["jira_token_file"])
        except FileNotFoundError as e:
            self.after(0, lambda m=str(e): messagebox.showerror("Token Error", m))
            self.after(0, self.progress.stop); return

        email    = self.config_data.get("jira_email", "")
        base_url = self.config_data.get("jira_base_url", "").rstrip("/")
        proj_key = self.config_data.get("jira_project_key", "")

        # build ADF description
        detail_bullets = []
        for fd in self.config_data["fields"]:
            if not fd.get("enabled", True): continue
            if fd.get("jira_field"): continue
            val = fv.get(fd["key"], "")
            if val:
                # show_label_in_jira: True → "Label: value", False → just "value"
                show_lbl = fd.get("show_label_in_jira", True)
                txt = f"{fd['label']}: {val}" if show_lbl else val
                detail_bullets.append({
                    "type": "listItem",
                    "content": [{"type": "paragraph", "content": [
                        {"type": "text", "text": txt}]}]
                })

        fc_bullets = []
        for fc in self.file_comments:
            jira_cmt = (fc.get("jira_comment") or fc.get("comment", "")).strip()
            txt = f"{STATUS_EMOJI.get(fc['status'],'~')} {fc['filename']}  [{fc['status'].upper()}]"
            if jira_cmt:
                txt += f"\n    {jira_cmt}"
            fc_bullets.append({
                "type": "listItem",
                "content": [{"type": "paragraph",
                             "content": [{"type": "text", "text": txt}]}]
            })

        adf = {"type": "doc", "version": 1, "content": [
            {"type": "heading", "attrs": {"level": 3},
             "content": [{"type": "text", "text": "Issue Details"}]},
            {"type": "bulletList", "content": detail_bullets} if detail_bullets else
            {"type": "paragraph", "content": [{"type": "text", "text": ""}]},
            {"type": "heading", "attrs": {"level": 3},
             "content": [{"type": "text", "text": "Description"}]},
            {"type": "paragraph", "content": [
                {"type": "text", "text": fv.get("description") or "No description."}]},
        ] + ([{"type": "heading", "attrs": {"level": 3},
               "content": [{"type": "text", "text": "Changed Files"}]},
              {"type": "bulletList", "content": fc_bullets}]
             if fc_bullets else [])}

        type_map = {"Bug":"Bug","Enhancement":"Story","Task":"Task",
                    "Story":"Story","Epic":"Epic","Sub-task":"Sub-task","Incident":"Bug"}
        issue_type = type_map.get(fv.get("issue_type","Bug"), "Bug")

        payload = {"fields": {
            "project":     {"key": fv.get("project") or proj_key},
            "summary":     fv.get("summary") or "(no summary)",
            "issuetype":   {"name": issue_type},
            "description": adf,
        }}

        if fv.get("reporter"):
            payload["fields"]["reporter"] = {"name": fv["reporter"]}

        # description-only fields that have jira_key → send to API directly
        _MULTI = {"components", "fixVersions", "labels"}
        _NAME  = {"priority", "issuetype"}
        for fd in self.config_data["fields"]:
            if fd.get("jira_field"): continue
            jk = fd.get("jira_key")
            if not jk: continue
            val = fv.get(fd["key"], "").strip()
            if not val: continue
            if jk in _MULTI:
                payload["fields"][jk] = [{"name": val}]
            elif jk in _NAME:
                payload["fields"][jk] = {"name": val}
            else:
                payload["fields"][jk] = val

        hdrs = jira_auth_headers(email, jira_tok)
        try:
            resp = requests.post(f"{base_url}/rest/api/3/issue",
                                 json=payload, headers=hdrs, timeout=20)
            resp.raise_for_status()
            data = resp.json()
        except requests.HTTPError as e:
            msg = f"Jira API error {e.response.status_code}:\n{e.response.text[:400]}"
            self.after(0, lambda m=msg: messagebox.showerror("Jira Error", m))
            self.after(0, self.progress.stop); return
        except Exception as e:
            self.after(0, lambda m=str(e): messagebox.showerror("Network Error", m))
            self.after(0, self.progress.stop); return

        issue_key = data.get("key", "")
        issue_url = f"{base_url}/browse/{issue_key}"
        self._jira_url = issue_url

        # generate word doc
        doc_path = None
        try:
            doc_path = generate_word_doc(fv, self.config_data["fields"],
                                         self.pr_cache, self.file_comments,
                                         self.config_data.get("word_doc_output_dir", BASE_DIR))
            self._last_doc_path = doc_path
        except Exception:
            pass

        def _attach_file(path, mime):
            try:
                ah = jira_auth_headers(email, jira_tok)
                ah["X-Atlassian-Token"] = "no-check"
                ah.pop("Content-Type", None)
                with open(path, "rb") as fh:
                    requests.post(
                        f"{base_url}/rest/api/3/issue/{issue_key}/attachments",
                        headers=ah,
                        files={"file": (os.path.basename(path), fh, mime)},
                        timeout=30)
            except Exception:
                pass

        # attach word doc
        if doc_path:
            _attach_file(doc_path,
                         "application/vnd.openxmlformats-officedocument"
                         ".wordprocessingml.document")

        # attach SQL file if generated
        sql_path = self._last_sql_path
        if sql_path and os.path.exists(sql_path):
            _attach_file(sql_path, "text/plain")

        attached = []
        if doc_path:   attached.append(os.path.basename(doc_path))
        if sql_path and os.path.exists(sql_path):
            attached.append(os.path.basename(sql_path))

        def _done():
            self.progress.stop()
            self.result_var.set(f"  Ticket created: {issue_key}   Click to open in Jira")
            self._set_status(f"Ticket {issue_key} created successfully")
            attach_note = ("\n\nAttached: " + ", ".join(attached)) if attached else ""
            messagebox.showinfo("Ticket Created",
                f"Jira ticket created!\n\n{issue_key}\n{issue_url}{attach_note}")
        self.after(0, _done)

    # ── helpers ───────────────────────────────────────────────────────────────

    def _collect_fields(self):
        fv = {}
        for key, var in self._field_widgets.items():
            fv[key] = var.get().strip()
        for key, var in getattr(self, "_jira_field_widgets", {}).items():
            fv[key] = var.get().strip()
        desc = self.desc_text.get("1.0", "end-1c").strip()
        if desc == "Describe the issue in detail...":
            desc = ""
        fv["description"] = desc
        fv["pr_url"]      = self.pr_url_var.get().strip()
        return fv

    def _choose_doc_dir(self):
        p = filedialog.askdirectory()
        if p:
            self.config_data["word_doc_output_dir"] = p

    def _open_word(self):
        import subprocess, sys
        if not self._last_doc_path or not os.path.exists(self._last_doc_path):
            messagebox.showinfo("No document", "Generate a Word doc first."); return
        if sys.platform == "win32":
            os.startfile(self._last_doc_path)
        elif sys.platform == "darwin":
            subprocess.run(["open", self._last_doc_path])
        else:
            subprocess.run(["xdg-open", self._last_doc_path])

    def _set_status(self, msg):
        self.status_var.set(msg)


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
