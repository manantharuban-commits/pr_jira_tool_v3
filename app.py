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
        cb_sel     = "#1a3a5c",   # checkbox checked indicator
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
    {"label": "Issue Owner",  "key": "owner",       "type": "text",
     "choices": [], "default": "", "jira_key": "assignee",   "required": True,  "enabled": True,
     "show_label_in_jira": True},
    {"label": "Environment",  "key": "environment", "type": "dropdown",
     "choices": ["Production","Staging","Development","QA","UAT","Pre-Production"],
     "default": "Production", "jira_key": "environment",     "required": True,  "enabled": True,
     "show_label_in_jira": True},
    {"label": "Issue Type",   "key": "issue_type",  "type": "dropdown",
     "choices": ["Bug","Enhancement","Task","Story","Epic","Sub-task","Incident"],
     "default": "Bug", "jira_key": "issuetype",              "required": True,  "enabled": True,
     "show_label_in_jira": False},
    {"label": "Dependency",   "key": "dependency",  "type": "text",
     "choices": [], "default": "", "jira_key": "dependency",  "required": False, "enabled": True,
     "show_label_in_jira": True},
    {"label": "Git PR Link",  "key": "git_link",    "type": "text",
     "choices": [], "default": "", "jira_key": "git_link",    "required": False, "enabled": True,
     "show_label_in_jira": True},
    {"label": "Doc File Name", "key": "doc_name",    "type": "text",
     "choices": [], "default": "", "jira_key": None,          "required": False, "enabled": True,
     "show_label_in_jira": False},
    {"label": "Doc Title",    "key": "doc_title",   "type": "text",
     "choices": [], "default": "", "jira_key": None,          "required": False, "enabled": True,
     "show_label_in_jira": False},
    {"label": "Jira Summary", "key": "summary",     "type": "text",
     "choices": [], "default": "", "jira_key": "summary",     "required": True,  "enabled": True,
     "show_label_in_jira": False},
    {"label": "Priority",     "key": "priority",    "type": "dropdown",
     "choices": ["Highest","High","Medium","Low","Lowest"],
     "default": "Medium", "jira_key": "priority",            "required": False, "enabled": True,
     "show_label_in_jira": True},
    {"label": "Component",    "key": "component",   "type": "text",
     "choices": [], "default": "", "jira_key": "components",  "required": False, "enabled": False,
     "show_label_in_jira": True},
    {"label": "Fix Version",  "key": "fix_version", "type": "text",
     "choices": [], "default": "", "jira_key": "fixVersions", "required": False, "enabled": False,
     "show_label_in_jira": True},
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
        cfg.update(saved)
        # merge saved field_defaults into field definitions
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
        doc.add_heading("File Changes & Review Comments", level=1)
        for idx, fc in enumerate(file_comments, 1):
            fname   = fc.get("filename", "unknown")
            status  = fc.get("status", "modified")
            comment = (fc.get("word_comment") or fc.get("comment", "")).strip()
            shots   = fc.get("screenshots", [])
            adds    = fc.get("additions", 0)
            dels    = fc.get("deletions", 0)

            doc.add_paragraph("")
            fh  = doc.add_paragraph()
            r1  = fh.add_run(f"[{STATUS_EMOJI.get(status,'~')}]  ")
            r1.bold = True; r1.font.size = Pt(11)
            r1.font.color.rgb = RGBColor(0x58, 0xa6, 0xff)
            r2  = fh.add_run(f"{idx}. {fname}")
            r2.bold = True; r2.font.size = Pt(11)
            r3  = fh.add_run(f"   ({status.upper()})")
            r3.font.size = Pt(9)
            r3.font.color.rgb = RGBColor(0x8b, 0x94, 0x9e)

            if adds or dels:
                sp = doc.add_paragraph(f"    +{adds} additions   -{dels} deletions")
                sp.paragraph_format.left_indent = Cm(0.5)
                if sp.runs:
                    sp.runs[0].font.size = Pt(9)
                    sp.runs[0].font.color.rgb = RGBColor(0x8b, 0x94, 0x9e)

            if comment:
                doc.add_paragraph("")
                ch = doc.add_paragraph("  Review Comment:")
                ch.runs[0].bold = True
                ch.runs[0].font.size = Pt(10)
                ch.runs[0].font.color.rgb = RGBColor(0x58, 0xa6, 0xff)
                cp = doc.add_paragraph(f"  {comment}")
                cp.paragraph_format.left_indent = Cm(0.8)

            # ── code diff/patch ──
            patch = fc.get("patch", "")
            if patch:
                doc.add_paragraph("")
                ph = doc.add_paragraph("  Code Diff:")
                if ph.runs:
                    ph.runs[0].bold = True
                    ph.runs[0].font.size = Pt(9)
                    ph.runs[0].font.color.rgb = RGBColor(0x8b, 0x94, 0x9e)
                patch_lines = patch.splitlines()
                shown = patch_lines[:80]
                for dl in shown:
                    pp = doc.add_paragraph(style="Normal")
                    pp.paragraph_format.left_indent = Cm(0.8)
                    r = pp.add_run(dl)
                    r.font.name = "Consolas"
                    r.font.size = Pt(8)
                    if dl.startswith('+'):
                        r.font.color.rgb = RGBColor(0x3f, 0xb9, 0x50)
                    elif dl.startswith('-'):
                        r.font.color.rgb = RGBColor(0xf8, 0x51, 0x49)
                    elif dl.startswith('@'):
                        r.font.color.rgb = RGBColor(0x58, 0xa6, 0xff)
                    else:
                        r.font.color.rgb = RGBColor(0x8b, 0x94, 0x9e)
                if len(patch_lines) > 80:
                    tp = doc.add_paragraph(
                        f"  ... {len(patch_lines) - 80} more lines (truncated)")
                    tp.paragraph_format.left_indent = Cm(0.8)
                    if tp.runs:
                        tp.runs[0].font.size = Pt(8)
                        tp.runs[0].font.color.rgb = RGBColor(0xd2, 0x99, 0x22)

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
            bot.set(qn("w:color"), "30363d")
            pBdr.append(bot)
            pPr.append(pBdr)

    doc_name = field_values.get("doc_name") or \
               f"Issue_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    if not doc_name.endswith(".docx"):
        doc_name += ".docx"
    out = os.path.join(output_dir, doc_name)
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
    body.append('<span class="h0">Issue Report</span>')
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
        body.append('<div class="section"><h1>File Changes &amp; Review Comments</h1>')
        for idx, fc in enumerate(file_comments, 1):
            fname        = fc.get("filename", "unknown")
            status       = fc.get("status", "modified")
            word_comment = (fc.get("word_comment") or fc.get("comment", "")).strip()
            jira_comment = fc.get("jira_comment", "").strip()
            shots        = fc.get("screenshots", [])
            adds         = fc.get("additions", 0)
            dels         = fc.get("deletions", 0)
            badge        = STATUS_EMOJI.get(status, "~")

            body.append(f'<div class="file-card file-{status}">')
            body.append(f'<span class="badge badge-{badge}">{badge} {status.upper()}</span>'
                        f'<span class="fname">{e(fname)}</span>')
            body.append(f'<div class="stats">'
                        f'<span class="adds">+{adds}</span>&nbsp;&nbsp;'
                        f'<span class="dels">-{dels}</span></div>')

            if word_comment:
                body.append(f'<div class="comment-box">'
                            f'<div class="comment-label">Word Comment</div>'
                            f'{e(word_comment)}</div>')
            if jira_comment and jira_comment != word_comment:
                body.append(f'<div class="comment-box" style="border-left-color:#58a6ff">'
                            f'<div class="comment-label" style="color:#58a6ff">Jira Comment</div>'
                            f'{e(jira_comment)}</div>')

            # code diff/patch
            patch = fc.get("patch", "")
            if patch:
                patch_lines = patch.splitlines()
                shown = patch_lines[:100]
                diff_rows = []
                for dl in shown:
                    if dl.startswith('+'):   cls = "diff-add"
                    elif dl.startswith('-'): cls = "diff-del"
                    elif dl.startswith('@'): cls = "diff-hunk"
                    else:                    cls = "diff-ctx"
                    diff_rows.append(f'<div class="{cls}">{e(dl)}</div>')
                if len(patch_lines) > 100:
                    diff_rows.append(
                        f'<div class="diff-ctx" style="color:#d29922">'
                        f'... {len(patch_lines)-100} more lines (truncated)</div>')
                body.append(
                    f'<details class="diff-block">'
                    f'<summary>View Diff &nbsp;'
                    f'<span style="color:#3fb950">+{adds}</span>'
                    f'&nbsp;<span style="color:#f85149">-{dels}</span>'
                    f'&nbsp;lines</summary>'
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
    """Renders the HTML preview in a scrollable tk window using tk.Text + image tags."""

    def __init__(self, parent, html_content, doc_path=None):
        super().__init__(parent)
        self.title("Document Preview")
        self.geometry("860x780")
        self.minsize(700, 500)
        self.configure(bg=C["bg"])
        self._doc_path = doc_path
        self._build(html_content)

    def _build(self, html_content):
        # toolbar
        bar = tk.Frame(self, bg=C["surface"], height=46)
        bar.pack(fill="x"); bar.pack_propagate(False)
        tk.Label(bar, text="Document Preview",
                 bg=C["surface"], fg=C["accent"],
                 font=("Segoe UI", 12, "bold")).pack(side="left", padx=16)

        if self._doc_path and os.path.exists(self._doc_path):
            tk.Button(bar, text="Open in Word",
                      bg=C["green"], fg="#0d1117", relief="flat",
                      font=("Segoe UI", 9, "bold"), padx=10, pady=3,
                      activebackground="#2ea043", cursor="hand2",
                      command=self._open_word).pack(side="right", padx=8, pady=8)

        tk.Button(bar, text="Save HTML",
                  bg=C["card"], fg=C["muted"], relief="flat",
                  font=("Segoe UI", 9), padx=10, pady=3,
                  activebackground=C["border"], cursor="hand2",
                  command=lambda: self._save_html(html_content)
                  ).pack(side="right", padx=4, pady=8)

        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")

        # try tkinterweb for proper HTML rendering
        try:
            from tkinterweb import HtmlFrame
            frame = HtmlFrame(self, horizontal_scrollbar="auto")
            frame.pack(fill="both", expand=True)
            frame.load_html(html_content)
        except ImportError:
            self._fallback_viewer(html_content)

    def _fallback_viewer(self, html_content):
        """Plain-text fallback when tkinterweb is not installed."""
        import re
        # strip tags
        text = re.sub(r'<style[^>]*>.*?</style>', '', html_content, flags=re.DOTALL)
        text = re.sub(r'<[^>]+>', '', text)
        text = text.replace("&amp;", "&").replace("&lt;", "<").replace(
            "&gt;", ">").replace("&#39;", "'").replace("&quot;", '"')
        # collapse whitespace
        import textwrap
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        text  = "\n".join(lines)

        # banner
        info = tk.Label(self,
            text="Install  pip install tkinterweb  for full HTML preview  |  "
                 "Showing plain-text fallback",
            bg=C["yellow"], fg="#0d1117",
            font=("Segoe UI", 9), pady=4)
        info.pack(fill="x")

        st = scrolledtext.ScrolledText(self, bg=C["card"], fg=C["text"],
                                       font=("Segoe UI", 10), relief="flat",
                                       wrap="word")
        st.pack(fill="both", expand=True, padx=4, pady=4)
        st.insert("1.0", text)
        st.config(state="disabled")

    def _open_word(self):
        if self._doc_path and os.path.exists(self._doc_path):
            import subprocess, sys
            if sys.platform == "win32":
                os.startfile(self._doc_path)
            elif sys.platform == "darwin":
                subprocess.run(["open", self._doc_path])
            else:
                subprocess.run(["xdg-open", self._doc_path])

    def _save_html(self, html_content):
        p = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("HTML", "*.html"), ("All", "*.*")])
        if p:
            with open(p, "w", encoding="utf-8") as f:
                f.write(html_content)
            messagebox.showinfo("Saved", f"HTML preview saved:\n{p}")


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

    def __init__(self, parent, fields):
        super().__init__(parent)
        self.title("Manage Jira Fields & Defaults")
        self.geometry("920x680")
        self.minsize(780, 520)
        self.configure(bg=C["bg"])
        self.grab_set()
        self.transient(parent)

        self._fields  = copy.deepcopy(fields)   # working copy
        self.result   = None
        self._row_frames = []

        self._build()

    # ── build ─────────────────────────────────────────────────────────────────

    def _build(self):
        # header
        hdr = tk.Frame(self, bg=C["surface"])
        hdr.pack(fill="x")
        tk.Frame(hdr, bg=C["purple"], width=4).pack(side="left", fill="y")
        tk.Label(hdr, text="  Manage Jira Fields & Default Values",
                 bg=C["surface"], fg=C["text"],
                 font=("Segoe UI", 13, "bold")).pack(side="left", padx=12, pady=14)
        tk.Label(hdr, text="Enable/disable fields, set defaults, control Jira output",
                 bg=C["surface"], fg=C["muted"],
                 font=("Segoe UI", 9)).pack(side="left", pady=4)
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")

        # column headers with tooltips
        col_hdr = tk.Frame(self, bg=C["surface"])
        col_hdr.pack(fill="x", padx=8, pady=(8, 2))

        col_defs = [
            ("", 32,       ""),
            ("Field Label",168, "The display name of the field in the UI and Word doc."),
            ("Type",       88,  "Data type: text, dropdown, number, or date."),
            ("Default Value",200,"Pre-filled value shown when a new ticket is created."),
            ("In Jira",    58,  "Include this field in the Jira ticket payload."),
            ("Show Label", 72,  "Prefix 'FieldName: ' before the value in Jira description.\n"
                                "e.g. 'Dependency: Yes'  vs  just 'Yes'."),
            ("",           100, ""),
        ]
        for txt, w, tip in col_defs:
            lbl = tk.Label(col_hdr, text=txt, bg=C["surface"], fg=C["accent"],
                           font=("Segoe UI", 9, "bold"), width=w//8, anchor="w")
            lbl.pack(side="left", padx=2)
            if tip:
                Tooltip(lbl, tip)

        tk.Frame(self, bg=C["border"], height=1).pack(fill="x", padx=8)

        # scrollable list
        outer = tk.Frame(self, bg=C["bg"])
        outer.pack(fill="both", expand=True, padx=8, pady=4)
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

        # bottom bar
        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")
        bot = tk.Frame(self, bg=C["surface"], height=52)
        bot.pack(fill="x", side="bottom"); bot.pack_propagate(False)

        tk.Button(bot, text="+ Add Custom Field",
                  bg=C["accent"], fg="#ffffff", relief="flat",
                  font=("Segoe UI", 9, "bold"), padx=12, pady=4,
                  activebackground=C["hover"], cursor="hand2",
                  command=self._add_field).pack(side="left", padx=10, pady=10)

        tk.Button(bot, text="Cancel",
                  bg=C["card"], fg=C["muted"], relief="flat",
                  font=("Segoe UI", 10), padx=12, pady=5,
                  activebackground=C["border"], cursor="hand2",
                  command=self.destroy).pack(side="right", padx=8, pady=10)
        tk.Button(bot, text="Apply & Close",
                  bg=C["green"], fg="#0d1117", relief="flat",
                  font=("Segoe UI", 10, "bold"), padx=14, pady=5,
                  activebackground="#2ea043", cursor="hand2",
                  command=self._apply).pack(side="right", padx=4, pady=10)

    def _render_rows(self):
        for w in self._list_frame.winfo_children():
            w.destroy()
        self._row_frames = []

        for i, fd in enumerate(self._fields):
            bg      = C["card"] if i % 2 == 0 else C["surface"]
            enabled = fd.get("enabled", True)
            txt_fg  = C["text"] if enabled else C["muted"]

            row = tk.Frame(self._list_frame, bg=bg,
                           highlightbackground=C["border"], highlightthickness=1)
            row.pack(fill="x", pady=1, padx=4)

            # index
            tk.Label(row, text=str(i+1), bg=bg, fg=C["muted"],
                     width=3, font=("Segoe UI", 8)).pack(side="left", padx=(6, 0))

            # ── enabled toggle with visible label ──
            en_var = tk.BooleanVar(value=enabled)
            en_f   = tk.Frame(row, bg=bg); en_f.pack(side="left", padx=4)
            tk.Checkbutton(en_f, variable=en_var, bg=bg, fg=C["text"],
                           selectcolor=C["cb_sel"],
                           activebackground=bg, activeforeground=C["text"],
                           relief="flat", bd=0, cursor="hand2",
                           font=("Segoe UI", 8),
                           command=lambda v=en_var, r=row, bg=bg: None
                           ).pack()

            # label
            lbl_var = tk.StringVar(value=fd["label"])
            if fd.get("custom"):
                ttk.Entry(row, textvariable=lbl_var, width=19).pack(side="left", padx=4)
            else:
                hint = FIELD_HINTS.get(fd.get("key", ""), "")
                lbl_w = tk.Label(row, textvariable=lbl_var, bg=bg, fg=txt_fg,
                                 width=19, anchor="w",
                                 font=("Segoe UI", 10, "bold" if fd.get("required") else "normal"))
                lbl_w.pack(side="left", padx=4)
                if hint:
                    Tooltip(lbl_w, hint)

            # type badge
            ftype     = fd.get("type", "text")
            type_clrs = {"text": C["muted"], "dropdown": C["accent"],
                         "number": C["yellow"], "date": C["green"]}
            tk.Label(row, text=ftype, bg=bg, fg=type_clrs.get(ftype, C["muted"]),
                     width=10, font=("Segoe UI", 9, "italic")).pack(side="left", padx=2)

            # default value
            def_var = tk.StringVar(value=fd.get("default", ""))
            choices = fd.get("choices", [])
            if choices:
                cb = ttk.Combobox(row, textvariable=def_var,
                                  values=choices, width=22, state="normal")
                cb.pack(side="left", padx=4)
            else:
                ttk.Entry(row, textvariable=def_var, width=25).pack(side="left", padx=4)

            # ── In Jira toggle ────────────────────────────────────────────────
            jira_var = tk.BooleanVar(value=bool(fd.get("jira_key")))
            jira_cb  = tk.Checkbutton(row, variable=jira_var, bg=bg, fg=C["text"],
                                      selectcolor=C["cb_sel"],
                                      activebackground=bg, activeforeground=C["text"],
                                      relief="flat", bd=0, cursor="hand2",
                                      font=("Segoe UI", 8))
            jira_cb.pack(side="left", padx=(10, 4))

            # ── Show Label toggle ─────────────────────────────────────────────
            show_lbl_var = tk.BooleanVar(value=fd.get("show_label_in_jira", True))
            show_lbl_cb  = tk.Checkbutton(row, variable=show_lbl_var, bg=bg, fg=C["text"],
                                           selectcolor=C["cb_sel"],
                                           activebackground=bg, activeforeground=C["text"],
                                           relief="flat", bd=0, cursor="hand2",
                                           font=("Segoe UI", 8))
            show_lbl_cb.pack(side="left", padx=(10, 4))

            # move up / down / delete
            btn_frame = tk.Frame(row, bg=bg)
            btn_frame.pack(side="right", padx=6)

            def _up(idx=i):
                if idx > 0:
                    self._fields[idx], self._fields[idx-1] = \
                        self._fields[idx-1], self._fields[idx]
                    self._render_rows()
            def _down(idx=i):
                if idx < len(self._fields)-1:
                    self._fields[idx], self._fields[idx+1] = \
                        self._fields[idx+1], self._fields[idx]
                    self._render_rows()
            def _del(idx=i, fd=fd):
                if not fd.get("custom") and fd.get("required"):
                    messagebox.showwarning("Cannot delete",
                        f'"{fd["label"]}" is a required field and cannot be deleted.\n'
                        'Disable it using the checkbox instead.')
                    return
                if messagebox.askyesno("Delete field",
                    f'Remove "{fd["label"]}"?', parent=self):
                    del self._fields[idx]
                    self._render_rows()

            for txt, cmd, fg in [("▲", _up, C["muted"]),
                                  ("▼", _down, C["muted"]),
                                  ("✕", _del, C["red"])]:
                tk.Button(btn_frame, text=txt, bg=bg, fg=fg,
                          relief="flat", font=("Segoe UI", 9), padx=4,
                          activebackground=C["border"],
                          command=cmd, cursor="hand2").pack(side="left")

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
        tk.Checkbutton(f6, variable=show_lbl_new, bg=C["bg"],
                       selectcolor=C["cb_sel"], activebackground=C["bg"],
                       relief="flat").pack(side="left")
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
                  bg=C["card"], fg=C["muted"], relief="flat",
                  font=("Segoe UI", 10), padx=12, pady=5,
                  cursor="hand2", command=dlg.destroy).pack(side="right", padx=(0, 8))

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
#  FileChangesPopup  (unchanged from v2, just cleaner)
# ─────────────────────────────────────────────────────────────────────────────

class FileChangesPopup(tk.Toplevel):

    def __init__(self, parent, files, pr_title):
        super().__init__(parent)
        self.title(f"File Changes  —  {pr_title[:70]}")
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
                 text="   Sorted by dependency  •  Step comments pre-filled  •  Word & Jira comments are independent",
                 bg=C["surface"], fg=C["muted"],
                 font=("Segoe UI", 9)).pack(side="left", pady=(3, 0))

        for lbl, st in [("Select All", True), ("Deselect All", False)]:
            tk.Button(bar, text=lbl,
                      bg=C["card"], fg=C["muted"], relief="flat",
                      font=("Segoe UI", 9), padx=10, pady=4,
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
                  bg=C["card"], fg=C["muted"], relief="flat",
                  font=("Segoe UI", 10), padx=12, pady=5,
                  activebackground=C["border"], cursor="hand2",
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
                diff_visible = tk.BooleanVar(value=False)
                diff_content = tk.Frame(body, bg="#0a0e14",
                                        highlightbackground=C["border"],
                                        highlightthickness=1)
                diff_text = tk.Text(diff_content, bg="#0a0e14", fg=C["text"],
                                    font=("Consolas", 8), height=10,
                                    relief="flat", insertbackground=C["text"],
                                    state="disabled", wrap="none",
                                    highlightthickness=0)
                diff_text.tag_config("add",  background="#0e2318", foreground="#3fb950")
                diff_text.tag_config("del",  background="#200e0e", foreground="#f85149")
                diff_text.tag_config("hunk", background="#0c1c2c", foreground="#58a6ff")
                diff_text.tag_config("ctx",  foreground="#6e7681")
                diff_text.config(state="normal")
                patch_lines = patch.splitlines()
                for dl in patch_lines[:80]:
                    if dl.startswith('+'):   tg = "add"
                    elif dl.startswith('-'): tg = "del"
                    elif dl.startswith('@'): tg = "hunk"
                    else:                    tg = "ctx"
                    diff_text.insert("end", dl + "\n", tg)
                if len(patch_lines) > 80:
                    diff_text.insert("end",
                        f"... {len(patch_lines)-80} more lines (truncated)\n", "hunk")
                diff_text.config(state="disabled")
                diff_xsb = ttk.Scrollbar(diff_content, orient="horizontal",
                                         command=diff_text.xview)
                diff_text.configure(xscrollcommand=diff_xsb.set)
                diff_text.pack(fill="both", expand=True, padx=1, pady=(1, 0))
                diff_xsb.pack(fill="x", padx=1, pady=(0, 1))

                def _sw_dt(e, dt=diff_text):
                    dt.yview_scroll(int(-1*(e.delta/120)), "units")
                diff_text.bind("<MouseWheel>", _sw_dt)

                def _toggle_diff(dv=diff_visible, dc=diff_content):
                    if dv.get():
                        dc.pack_forget()
                        dv.set(False)
                        diff_btn.config(text="▶ View Diff")
                    else:
                        dc.pack(fill="x", padx=0, pady=(2, 4))
                        dv.set(True)
                        diff_btn.config(text="▼ Hide Diff")

                n_lines = len(patch_lines)
                diff_btn = tk.Button(
                    gh_row,
                    text=f"▶ View Diff  ({n_lines} lines)",
                    bg=rbg, fg=C["muted"], relief="flat",
                    font=("Segoe UI", 8), padx=6, pady=2, cursor="hand2",
                    activebackground=C["border"],
                    command=_toggle_diff)
                diff_btn.pack(side="left", padx=(6, 0))

            # ── two comment columns ──────────────────────────────────────────
            comment_row = tk.Frame(body, bg=rbg)
            comment_row.pack(fill="x", pady=(0, 4))

            # Word comment (left)
            word_col = tk.Frame(comment_row, bg=rbg)
            word_col.pack(side="left", fill="both", expand=True, padx=(0, 6))
            tk.Label(word_col, text="Word Comment:", bg=rbg, fg=C["yellow"],
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
                      bg=C["surface"], fg=C["accent"], relief="flat",
                      font=("Segoe UI", 9), padx=10, pady=3,
                      activebackground=C["border"], cursor="hand2",
                      command=_add_shots).pack(side="left", padx=(0, 6))
            tk.Button(srow, text="Clear",
                      bg=C["surface"], fg=C["muted"], relief="flat",
                      font=("Segoe UI", 9), padx=8, pady=3,
                      activebackground=C["border"], cursor="hand2",
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
        self._field_widgets = {}
        self._theme_name    = "dark"

        self.title("PR  to  Jira  v3")
        self.geometry("1060x940")
        self.minsize(880, 720)
        self.configure(bg=C["bg"])
        self._apply_theme()
        self._build_ui()

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
        if self.pr_files and hasattr(self, "fc_btn"):
            self.fc_btn.config(state="normal")
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
        title_f.pack(side="left", padx=(16, 0), pady=12)
        tk.Label(title_f, text="PR  ▶  Jira",
                 bg=C["surface"], fg=C["accent"],
                 font=("Segoe UI", 15, "bold")).pack(side="left")
        tk.Label(title_f, text="  v3",
                 bg=C["surface"], fg=C["muted"],
                 font=("Segoe UI", 10)).pack(side="left", pady=(4, 0))

        # right-side controls
        ctrl_f = tk.Frame(hdr, bg=C["surface"])
        ctrl_f.pack(side="right", padx=12, pady=10)

        theme_lbl = "   Light Mode   " if self._theme_name == "dark" else "   Dark Mode   "
        theme_ico = "☀" if self._theme_name == "dark" else "\U0001f319"
        tk.Button(ctrl_f, text=f"{theme_ico}{theme_lbl}",
                  bg=C["inp"], fg=C["muted"],
                  relief="flat", font=("Segoe UI", 9, "bold"),
                  padx=8, pady=5, cursor="hand2",
                  activebackground=C["border"], activeforeground=C["text"],
                  command=self._toggle_theme).pack(side="right", padx=(6, 0))

        for txt, cmd, bg, fg in [
            (" ⚡ Manage Fields ", self._open_field_mgr, C["purple"], "#ffffff"),
            ("  ⚙ Settings  ",    self._open_settings,  C["surface"], C["muted"]),
        ]:
            tk.Button(ctrl_f, text=txt, bg=bg, fg=fg, relief="flat",
                      font=("Segoe UI", 9, "bold"), padx=8, pady=5,
                      activebackground=C["border"], activeforeground=C["text"],
                      cursor="hand2", command=cmd).pack(side="right", padx=3)

        tk.Frame(self, bg=C["border"], height=1).pack(fill="x")

        # ── notebook ──────────────────────────────────────────────────────────
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        self.tab_create  = ttk.Frame(nb)
        self.tab_preview = ttk.Frame(nb)
        self.tab_pr_view = ttk.Frame(nb)

        nb.add(self.tab_create,  text="   Create Ticket   ")
        nb.add(self.tab_preview, text="   PR Preview   ")
        nb.add(self.tab_pr_view, text="   Doc Preview   ")
        self._nb = nb

        self._build_create_tab()
        self._build_pr_view_tab()
        self._build_doc_preview_tab()

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
        border_f.pack(fill="x", padx=20, pady=(0, 14))
        # 1-px border simulation
        inner_wrap = tk.Frame(border_f, bg=C["card"])
        inner_wrap.pack(fill="x", padx=1, pady=1)
        # left accent strip
        accent_strip = tk.Frame(inner_wrap, bg=accent_color, width=3)
        accent_strip.pack(side="left", fill="y")
        content = tk.Frame(inner_wrap, bg=C["card"])
        content.pack(side="left", fill="both", expand=True, padx=16, pady=14)
        if title:
            th = tk.Frame(content, bg=C["card"])
            th.pack(fill="x", pady=(0, 10))
            if icon:
                tk.Label(th, text=icon, bg=C["card"], fg=accent_color,
                         font=("Segoe UI", 13)).pack(side="left", padx=(0, 8))
            tk.Label(th, text=title, bg=C["card"], fg=C["text"],
                     font=("Segoe UI", 12, "bold")).pack(side="left")
            tk.Frame(content, bg=C["border"], height=1).pack(fill="x", pady=(0, 12))
        return content

    # ── Create Ticket tab ────────────────────────────────────────────────────

    def _build_create_tab(self):
        sf = self._scrollable(self.tab_create)
        self._build_pr_card(sf)
        self._build_dynamic_fields_card(sf)
        self._build_desc_card(sf)
        self._build_actions_card(sf)

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
        self.fc_btn = ttk.Button(row, text="  View File Changes  ", style="Purple.TButton",
                                 state="disabled", command=self._open_file_changes)
        self.fc_btn.pack(side="left", padx=(6, 0))

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

    # Dynamic fields card — rebuilt whenever field definitions change
    def _build_dynamic_fields_card(self, parent):
        self._fields_card_parent = parent
        self._rebuild_fields_card()

    def _rebuild_fields_card(self):
        if hasattr(self, "_fields_card") and self._fields_card:
            self._fields_card.destroy()

        card = self._card(self._fields_card_parent, "Issue Details", "")
        # card hierarchy: content → inner_wrap → border_f → parent
        # store border_f so destroy/rebuild replaces only this card, not the parent
        self._fields_card = card.master.master

        self._field_widgets = {}
        enabled = [fd for fd in self.config_data["fields"] if fd.get("enabled", True)]

        for fd in enabled:
            key     = fd["key"]
            label   = fd["label"]
            ftype   = fd.get("type", "text")
            choices = fd.get("choices", [])
            default = fd.get("default", "")
            req     = fd.get("required", False)

            row = tk.Frame(card, bg=C["card"]); row.pack(fill="x", pady=(0, 8))
            # fill="y" ensures the label column takes the full row height even
            # with pack_propagate(False) (otherwise height stays 0 and label is invisible)
            lf  = tk.Frame(row, bg=C["card"], width=180); lf.pack(side="left", fill="y")
            lf.pack_propagate(False)

            req_badge = " *" if req else ""
            lbl_w = tk.Label(lf, text=f"{label}{req_badge}",
                             bg=C["card"], fg=C["text"],
                             font=("Segoe UI", 10, "bold" if req else "normal"),
                             anchor="w")
            lbl_w.pack(side="left", padx=(0, 4), pady=4)

            hint = FIELD_HINTS.get(key, "")
            if hint:
                help_btn = tk.Label(lf, text="?",
                                    bg=C["card"], fg=C["accent"],
                                    font=("Segoe UI", 9, "bold"), cursor="hand2")
                help_btn.pack(side="left", pady=4)
                Tooltip(help_btn, hint)

            var = tk.StringVar(value=default)

            if ftype == "dropdown" and choices:
                w = ttk.Combobox(row, textvariable=var, values=choices,
                                 state="readonly")
            else:
                w = ttk.Entry(row, textvariable=var)

            w.pack(side="left", fill="x", expand=True)

            if key == "doc_name":
                tk.Button(row, text="Browse", bg=C["surface"], fg=C["muted"],
                          relief="flat", font=("Segoe UI", 9), padx=8, pady=3,
                          activebackground=C["border"], cursor="hand2",
                          command=self._choose_doc_dir
                          ).pack(side="left", padx=(6, 0))

            self._field_widgets[key] = var

        # ── auto-sync: Jira Summary → Doc File Name + Doc Title ──────────────
        if "summary" in self._field_widgets:
            def _sanitize_fname(s):
                s = re.sub(r'[\[\](){}<>:"/\\|?*\n\r\t]', '_', s)
                s = re.sub(r'\s+', '_', s.strip())
                s = re.sub(r'_+', '_', s).strip('_')
                return s[:80] if s else ""

            _cur_sum = self._field_widgets["summary"].get()
            self._last_auto_doc_name  = _sanitize_fname(_cur_sum)
            self._last_auto_doc_title = _cur_sum
            self._syncing = False

            def _on_summary_change(*_):
                if getattr(self, "_syncing", False):
                    return
                self._syncing = True
                try:
                    sv = self._field_widgets.get("summary")
                    if sv is None:
                        return
                    val = sv.get()
                    if "doc_name" in self._field_widgets:
                        cur = self._field_widgets["doc_name"].get()
                        if cur == getattr(self, "_last_auto_doc_name", None) or cur == "":
                            new = _sanitize_fname(val)
                            self._field_widgets["doc_name"].set(new)
                            self._last_auto_doc_name = new
                    if "doc_title" in self._field_widgets:
                        cur = self._field_widgets["doc_title"].get()
                        if cur == getattr(self, "_last_auto_doc_title", None) or cur == "":
                            self._field_widgets["doc_title"].set(val)
                            self._last_auto_doc_title = val
                finally:
                    self._syncing = False

            self._field_widgets["summary"].trace_add("write", _on_summary_change)

    def _build_desc_card(self, parent):
        card = self._card(parent, "Description", "", accent_color=C["muted"])
        tk.Label(card, text="Detailed description of the issue — included in the Word doc and Jira ticket.",
                 bg=C["card"], fg=C["muted"],
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 6))
        self.desc_text = scrolledtext.ScrolledText(
            card, height=5, bg=C["inp"], fg=C["muted"],
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
        btn  = tk.Frame(card, bg=C["card"]); btn.pack(fill="x")

        for txt, style, cmd, tip in [
            ("  Preview Doc  ",       "Orange.TButton",  self._show_preview,
             "Open a live HTML preview of the Word document in a pop-out window."),
            ("  Generate Word Doc  ", "Ghost.TButton",   self._gen_word_thread,
             "Save the Word .docx file to the output directory."),
            ("  Create Jira Ticket  ","Success.TButton", self._create_jira_thread,
             "Create a Jira ticket with all details and attach the Word doc."),
        ]:
            b = ttk.Button(btn, text=txt, style=style, command=cmd)
            b.pack(side="left", padx=(0, 8))
            Tooltip(b, tip)

        self.progress = ttk.Progressbar(card, mode="indeterminate")
        self.progress.pack(fill="x", pady=(12, 0))

        self.result_var = tk.StringVar()
        rl = tk.Label(card, textvariable=self.result_var,
                      bg=C["card"], fg=C["green"],
                      font=("Segoe UI", 10, "bold"), cursor="hand2",
                      wraplength=820, justify="left")
        rl.pack(anchor="w", pady=(6, 0))
        rl.bind("<Button-1>", lambda e: self._jira_url and webbrowser.open(self._jira_url))

    # ── PR Preview tab ────────────────────────────────────────────────────────

    def _build_pr_view_tab(self):
        tk.Label(self.tab_preview, text="PR Preview",
                 bg=C["bg"], fg=C["accent"],
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=20, pady=16)
        self.pr_preview_text = scrolledtext.ScrolledText(
            self.tab_preview, bg=C["inp"], fg=C["text"],
            font=("Consolas", 9), relief="flat", insertbackground=C["text"])
        self.pr_preview_text.pack(fill="both", expand=True, padx=20, pady=(0, 20))

    def _update_pr_preview(self, pr):
        self.pr_preview_text.config(state="normal")
        self.pr_preview_text.delete("1.0", "end")
        lines = [
            f"PR #{pr.get('number')}  —  {pr.get('title', '')}",
            f"State:    {pr.get('state', '').upper()}",
            f"Author:   {pr.get('user', {}).get('login', '')}",
            f"Branch:   {pr.get('head', {}).get('ref', '')} -> {pr.get('base', {}).get('ref', '')}",
            f"Created:  {(pr.get('created_at') or '')[:10]}",
            f"Updated:  {(pr.get('updated_at') or '')[:10]}",
            f"URL:      {pr.get('html_url', '')}",
            "", "-"*60, "",
            pr.get("body") or "(No PR description)",
        ]
        self.pr_preview_text.insert("1.0", "\n".join(lines))
        self.pr_preview_text.config(state="disabled")

    # ── Doc Preview tab ───────────────────────────────────────────────────────

    def _build_doc_preview_tab(self):
        """Embedded preview panel inside the main window."""
        top = tk.Frame(self.tab_pr_view, bg=C["surface"], height=46)
        top.pack(fill="x"); top.pack_propagate(False)
        tk.Label(top, text="Document Preview",
                 bg=C["surface"], fg=C["accent"],
                 font=("Segoe UI", 12, "bold")).pack(side="left", padx=16)

        tk.Button(top, text="Refresh Preview",
                  bg=C["card"], fg=C["muted"], relief="flat",
                  font=("Segoe UI", 9), padx=10, pady=3,
                  activebackground=C["border"], cursor="hand2",
                  command=self._refresh_embedded_preview
                  ).pack(side="right", padx=8, pady=8)

        tk.Button(top, text="Open Word Doc",
                  bg=C["green"], fg="#0d1117", relief="flat",
                  font=("Segoe UI", 9, "bold"), padx=10, pady=3,
                  activebackground="#2ea043", cursor="hand2",
                  command=self._open_word).pack(side="right", padx=4, pady=8)

        tk.Button(top, text="Pop-out Preview",
                  bg=C["purple"], fg="#0d1117", relief="flat",
                  font=("Segoe UI", 9, "bold"), padx=10, pady=3,
                  activebackground="#a070e0", cursor="hand2",
                  command=self._show_preview).pack(side="right", padx=4, pady=8)

        tk.Frame(self.tab_pr_view, bg=C["border"], height=1).pack(fill="x")

        # banner for when nothing is generated yet
        self._preview_placeholder = tk.Label(
            self.tab_pr_view,
            text="Click  Preview Doc  or  Generate Word Doc  to see the document preview here.",
            bg=C["bg"], fg=C["muted"],
            font=("Segoe UI", 11), wraplength=600)
        self._preview_placeholder.pack(expand=True)

        # container for the actual preview widget
        self._preview_container = tk.Frame(self.tab_pr_view, bg=C["bg"])

    def _refresh_embedded_preview(self):
        """Rebuild embedded preview from current state (no file needed)."""
        html_content = build_preview_html(
            self._collect_fields(),
            self.config_data["fields"],
            self.pr_cache,
            self.file_comments,
        )
        self._embed_preview(html_content)
        self._nb.select(self.tab_pr_view)

    def _embed_preview(self, html_content):
        self._preview_placeholder.pack_forget()
        for w in self._preview_container.winfo_children():
            w.destroy()
        self._preview_container.pack(fill="both", expand=True)

        try:
            from tkinterweb import HtmlFrame
            hf = HtmlFrame(self._preview_container, horizontal_scrollbar="auto")
            hf.pack(fill="both", expand=True)
            hf.load_html(html_content)
        except ImportError:
            self._embed_fallback(self._preview_container, html_content)

    def _embed_fallback(self, parent, html_content):
        import re
        text = re.sub(r'<style[^>]*>.*?</style>', '', html_content, flags=re.DOTALL)
        text = re.sub(r'<[^>]+>', '', text)
        for ent, ch in [("&amp;","&"),("&lt;","<"),("&gt;",">"),
                        ("&#39;","'"),("&quot;",'"'),("&nbsp;"," ")]:
            text = text.replace(ent, ch)
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        text  = "\n".join(lines)

        info = tk.Label(parent,
            text="Install  pip install tkinterweb  for rich HTML preview",
            bg=C["yellow"], fg="#0d1117", font=("Segoe UI", 9), pady=4)
        info.pack(fill="x")

        st = scrolledtext.ScrolledText(parent, bg=C["card"], fg=C["text"],
                                       font=("Segoe UI", 10), relief="flat", wrap="word")
        st.pack(fill="both", expand=True, padx=4, pady=4)
        st.insert("1.0", text)
        st.config(state="disabled")

    # ── settings window ───────────────────────────────────────────────────────

    def _open_settings(self):
        win = tk.Toplevel(self)
        win.title("Settings")
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
            ("Word Doc Output Dir", "word_doc_output_dir"),
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
                ttk.Button(row, text="...", style="Ghost.TButton",
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
        ttk.Button(br, text="Cancel", style="Ghost.TButton",
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
            self.files_sum_var.set("   ".join(parts) +
                                   "   — click View File Changes to review")
            self.fc_btn.config(state="normal")
            self._update_pr_preview(pr)
            self.progress.stop()
            self._set_status(f"PR #{pr['number']} loaded — {total} files")

        self.after(0, _ui)

    # ── File changes ──────────────────────────────────────────────────────────

    def _open_file_changes(self):
        if not self.pr_files:
            messagebox.showinfo("No files", "Fetch a PR first."); return
        title = (self.pr_cache or {}).get("title", "")
        pop   = FileChangesPopup(self, self.pr_files, title)
        self.wait_window(pop)
        res = pop.get_result()
        if res:
            self.file_comments = res
            n_c = sum(1 for r in res if r["comment"])
            n_s = sum(1 for r in res if r["screenshots"])
            base = self.files_sum_var.get().split("—")[0].strip()
            self.files_sum_var.set(
                f"{base}   —   {len(res)} reviewed  |  {n_c} comments  |  {n_s} screenshots")
            self._set_status(
                f"Review saved: {len(res)} files, {n_c} comments, {n_s} screenshots")

    # ── Preview ───────────────────────────────────────────────────────────────

    def _show_preview(self):
        """Open preview in a pop-out window."""
        html_content = build_preview_html(
            self._collect_fields(),
            self.config_data["fields"],
            self.pr_cache,
            self.file_comments,
        )
        DocPreviewWindow(self, html_content, self._last_doc_path)

    # ── Word doc ─────────────────────────────────────────────────────────────

    def _gen_word_thread(self):
        threading.Thread(target=self._gen_word, daemon=True).start()

    def _gen_word(self):
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
        # refresh embedded preview
        html_content = build_preview_html(
            fv, self.config_data["fields"], self.pr_cache, self.file_comments)
        self.after(0, lambda: self._embed_preview(html_content))
        self.after(0, self.progress.stop)
        self.after(0, lambda p=path: messagebox.showinfo("Done", f"Word doc saved:\n{p}"))
        self.after(0, lambda: self._set_status("Word document saved"))

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
            if fd.get("jira_key") in ("summary", "issuetype", "assignee"): continue
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
            "project":     {"key": proj_key},
            "summary":     fv.get("summary") or "(no summary)",
            "issuetype":   {"name": issue_type},
            "description": adf,
        }}

        # priority
        if fv.get("priority"):
            payload["fields"]["priority"] = {"name": fv["priority"]}

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

        # attach word doc
        if doc_path:
            try:
                ah = jira_auth_headers(email, jira_tok)
                ah["X-Atlassian-Token"] = "no-check"
                ah.pop("Content-Type", None)
                with open(doc_path, "rb") as fh:
                    requests.post(
                        f"{base_url}/rest/api/3/issue/{issue_key}/attachments",
                        headers=ah,
                        files={"file": (os.path.basename(doc_path), fh,
                               "application/vnd.openxmlformats-officedocument"
                               ".wordprocessingml.document")},
                        timeout=30)
            except Exception:
                pass

        html_content = build_preview_html(
            fv, self.config_data["fields"], self.pr_cache, self.file_comments)

        def _done():
            self.progress.stop()
            self._embed_preview(html_content)
            self.result_var.set(f"  Ticket created: {issue_key}   Click to open in Jira")
            self._set_status(f"Ticket {issue_key} created successfully")
            messagebox.showinfo("Ticket Created",
                f"Jira ticket created!\n\n{issue_key}\n{issue_url}"
                + (f"\n\nWord doc attached: {os.path.basename(doc_path)}" if doc_path else ""))
        self.after(0, _done)

    # ── helpers ───────────────────────────────────────────────────────────────

    def _collect_fields(self):
        fv = {}
        for key, var in self._field_widgets.items():
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
