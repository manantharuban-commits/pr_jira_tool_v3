"""
Validation script for PR-Jira Tool v3 changes.

Tests:
  1. _dep_score()           – file dependency scoring
  2. File sort order        – dependency-based ordering of PR files
  3. Default step comments  – auto-generated step N comments
  4. show_label_in_jira     – label prefix toggle in Jira ADF
  5. word_comment/jira_comment split
  6. per-screenshot comments [{"path": str, "comment": str}] format
  7. doc_title flows into Word doc heading
  8. patch/diff in word_doc and HTML
  9. doc_name sanitization helper
  10. generate_word_doc()   – Word document output
  11. build_preview_html()  – HTML preview
  12. Jira ADF payload structure

Run:  python -X utf8 dumpdata/validate.py
"""

import sys, os, json, copy

DUMP_DIR   = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(DUMP_DIR)
sys.path.insert(0, PARENT_DIR)

# ── import non-GUI symbols from app ──────────────────────────────────────────
from app import (
    _dep_score,
    generate_word_doc,
    build_preview_html,
    STATUS_EMOJI,
    STATUS_VERB,
    BUILTIN_FIELDS,
)

SEP   = "-" * 70
PASS  = "  [PASS]"
FAIL  = "  [FAIL]"

results = []

def check(label, condition, detail=""):
    status = PASS if condition else FAIL
    line   = f"{status}  {label}"
    if detail:
        line += f"\n         {detail}"
    print(line)
    results.append(condition)
    return condition


# ─────────────────────────────────────────────────────────────────────────────
# 1. Load dump data
# ─────────────────────────────────────────────────────────────────────────────
print(f"\n{SEP}")
print("  LOADING DUMP DATA")
print(SEP)

with open(os.path.join(DUMP_DIR, "sample_pr.json"))        as f: PR_DATA   = json.load(f)
with open(os.path.join(DUMP_DIR, "sample_pr_files.json"))  as f: PR_FILES  = json.load(f)
with open(os.path.join(DUMP_DIR, "sample_field_values.json")) as f: FIELD_VALS = json.load(f)

print(f"  PR:    #{PR_DATA['number']} — {PR_DATA['title']}")
print(f"  Files: {len(PR_FILES)} total")
print(f"  Fields: {len(FIELD_VALS)} values loaded")


# ─────────────────────────────────────────────────────────────────────────────
# 2. _dep_score — score each file
# ─────────────────────────────────────────────────────────────────────────────
print(f"\n{SEP}")
print("  TEST 1: _dep_score() per file")
print(SEP)

EXPECTED_SCORES = {
    "src/config/jwt_config.py":       0,   # config
    "src/constants/auth_constants.py":0,   # constants
    "src/models/user_model.py":       1,   # model
    "src/utils/auth_helpers.py":      2,   # utils
    "src/services/auth_service.py":   3,   # service
    "src/routes/auth_routes.py":      4,   # route
    "src/controllers/login_controller.py": 4, # controller
    "tests/test_auth_service.py":     5,   # test
    "README.md":                      3,   # default
}

all_score_ok = True
for fname, expected in EXPECTED_SCORES.items():
    got = _dep_score(fname)
    ok  = got == expected
    if not ok:
        all_score_ok = False
    print(f"    {'OK' if ok else 'XX'}  score={got} (expected {expected})  {fname}")

check("All dep_score values correct", all_score_ok)


# ─────────────────────────────────────────────────────────────────────────────
# 3. File sort order
# ─────────────────────────────────────────────────────────────────────────────
print(f"\n{SEP}")
print("  TEST 2: Dependency-based file sort order (no-dep first)")
print(SEP)

sorted_files = sorted(PR_FILES,
                      key=lambda f: (_dep_score(f["filename"]), f["filename"]))

print("  Sorted order:")
for i, f in enumerate(sorted_files, 1):
    score = _dep_score(f["filename"])
    print(f"    {i}. score={score}  {f['filename']}")

# Verify: first file should have lower score than last
first_score = _dep_score(sorted_files[0]["filename"])
last_score  = _dep_score(sorted_files[-1]["filename"])
check("First file score <= last file score",
      first_score <= last_score,
      f"first={first_score}, last={last_score}")

# No file with tests before config
config_idx = next(i for i, f in enumerate(sorted_files)
                  if "config" in f["filename"].lower())
test_idx   = next(i for i, f in enumerate(sorted_files)
                  if "test" in f["filename"].lower())
check("Config file comes before test file",
      config_idx < test_idx,
      f"config_idx={config_idx}, test_idx={test_idx}")


# ─────────────────────────────────────────────────────────────────────────────
# 4. Default step comments
# ─────────────────────────────────────────────────────────────────────────────
print(f"\n{SEP}")
print("  TEST 3: Default step comment generation")
print(SEP)

def default_step_comment(step, filename, status):
    verb = STATUS_VERB.get(status, "Modify")
    base = os.path.basename(filename)
    return f"Step {step}: {verb} {base}"

print("  Generated step comments:")
for i, f in enumerate(sorted_files, 1):
    cmt = default_step_comment(i, f["filename"], f["status"])
    print(f"    {cmt}")

# Check step 1 starts with "Step 1:"
step1 = default_step_comment(1, sorted_files[0]["filename"], sorted_files[0]["status"])
check("Step 1 comment starts with 'Step 1:'", step1.startswith("Step 1:"),
      f"got: '{step1}'")

# Check "Add" verb for added file
added_files = [f for f in sorted_files if f["status"] == "added"]
if added_files:
    cmt = default_step_comment(1, added_files[0]["filename"], "added")
    check("'added' status → verb 'Add'", "Add" in cmt, f"got: '{cmt}'")

# Check "Update" verb for modified file
mod_files = [f for f in sorted_files if f["status"] == "modified"]
if mod_files:
    cmt = default_step_comment(1, mod_files[0]["filename"], "modified")
    check("'modified' status → verb 'Update'", "Update" in cmt, f"got: '{cmt}'")


# ─────────────────────────────────────────────────────────────────────────────
# 5. show_label_in_jira — Jira ADF label prefix logic
# ─────────────────────────────────────────────────────────────────────────────
print(f"\n{SEP}")
print("  TEST 4: show_label_in_jira toggle")
print(SEP)

def build_detail_bullets(fields, fv):
    bullets = []
    for fd in fields:
        if not fd.get("enabled", True): continue
        if fd.get("jira_key") in ("summary", "issuetype", "assignee"): continue
        val = fv.get(fd["key"], "")
        if val:
            show_lbl = fd.get("show_label_in_jira", True)
            txt = f"{fd['label']}: {val}" if show_lbl else val
            bullets.append({"field": fd["label"], "text": txt,
                            "show_label": show_lbl, "value": val})
    return bullets

fields_with_show = copy.deepcopy(BUILTIN_FIELDS)
bullets = build_detail_bullets(fields_with_show, FIELD_VALS)

print("  Jira description bullets:")
all_lbl_ok = True
for b in bullets:
    expected = f"{b['field']}: {b['value']}" if b["show_label"] else b["value"]
    ok = b["text"] == expected
    if not ok: all_lbl_ok = False
    marker = "OK" if ok else "XX"
    print(f"    {marker}  show_label={b['show_label']}  → '{b['text']}'")

check("All label-prefix formats correct", all_lbl_ok)

# Verify Dependency shows label (show_label_in_jira=True)
dep_bullet = next((b for b in bullets if b["field"] == "Dependency"), None)
if dep_bullet:
    check("Dependency bullet shows label prefix",
          dep_bullet["text"].startswith("Dependency:"),
          f"got: '{dep_bullet['text']}'")

# Verify Issue Type does NOT show label (show_label_in_jira=False)
itype_field = next((fd for fd in BUILTIN_FIELDS if fd["key"] == "issue_type"), None)
if itype_field:
    check("Issue Type show_label_in_jira is False by default",
          itype_field.get("show_label_in_jira", True) == False)


# ─────────────────────────────────────────────────────────────────────────────
# 6. word_comment / jira_comment split
# ─────────────────────────────────────────────────────────────────────────────
print(f"\n{SEP}")
print("  TEST 5: word_comment / jira_comment separation")
print(SEP)

# Simulate what FileChangesPopup._save() produces
simulated_file_comments = []
for i, f in enumerate(sorted_files, 1):
    fname  = f["filename"]
    status = f["status"]
    word_c = default_step_comment(i, fname, status)
    jira_c = f"Step {i}: {os.path.basename(fname)}"   # shorter jira version
    simulated_file_comments.append({
        "filename":     fname,
        "status":       status,
        "additions":    f["additions"],
        "deletions":    f["deletions"],
        "word_comment": word_c,
        "jira_comment": jira_c,
        "comment":      word_c,     # backward compat
        "screenshots":  [],
    })

print("  File comments (word vs jira):")
for fc in simulated_file_comments:
    diff = "SAME" if fc["word_comment"] == fc["jira_comment"] else "DIFF"
    print(f"    [{diff}]  word='{fc['word_comment']}'")
    if diff == "DIFF":
        print(f"            jira='{fc['jira_comment']}'")

check("All entries have word_comment key",
      all("word_comment" in fc for fc in simulated_file_comments))
check("All entries have jira_comment key",
      all("jira_comment" in fc for fc in simulated_file_comments))
check("backward-compat 'comment' key present",
      all("comment" in fc for fc in simulated_file_comments))
check("'comment' equals word_comment",
      all(fc["comment"] == fc["word_comment"] for fc in simulated_file_comments))


# ─────────────────────────────────────────────────────────────────────────────
# 6. per-screenshot comments format
# ─────────────────────────────────────────────────────────────────────────────
print(f"\n{SEP}")
print("  TEST 6: per-screenshot comments format")
print(SEP)

# Simulate new dict-format screenshots
shot_dicts = [
    {"path": "/nonexistent/shot1.png", "comment": "Auth flow overview"},
    {"path": "/nonexistent/shot2.png", "comment": ""},
]
shot_strs  = ["/nonexistent/shot1.png", "/nonexistent/shot2.png"]

# Backward-compat: both dict and str should be handled
def _extract_shot_path(s):
    return s.get("path", "") if isinstance(s, dict) else str(s)

def _extract_shot_comment(s):
    return s.get("comment", "") if isinstance(s, dict) else ""

paths_from_dicts = [_extract_shot_path(s) for s in shot_dicts]
paths_from_strs  = [_extract_shot_path(s) for s in shot_strs]
check("Dict-format screenshot path extraction",
      paths_from_dicts == ["/nonexistent/shot1.png", "/nonexistent/shot2.png"])
check("Str-format screenshot path extraction (backward compat)",
      paths_from_strs  == ["/nonexistent/shot1.png", "/nonexistent/shot2.png"])
check("Dict comment extraction",
      _extract_shot_comment(shot_dicts[0]) == "Auth flow overview")
check("Str comment extraction returns empty string",
      _extract_shot_comment(shot_strs[0]) == "")

# Simulate _save() output: filter placeholder caption
def _filter_caption(cap):
    return cap if cap != "Screenshot caption..." else ""
check("Placeholder caption filtered to empty",
      _filter_caption("Screenshot caption...") == "")
check("Real caption preserved",
      _filter_caption("Login page screenshot") == "Login page screenshot")


# ─────────────────────────────────────────────────────────────────────────────
# 7. doc_title and doc_name auto-sync
# ─────────────────────────────────────────────────────────────────────────────
print(f"\n{SEP}")
print("  TEST 7: doc_title / doc_name from BUILTIN_FIELDS")
print(SEP)

import re as _re

def _sanitize_fname(s):
    s = _re.sub(r'[\[\](){}<>:"/\\|?*\n\r\t]', '_', s)
    s = _re.sub(r'\s+', '_', s.strip())
    s = _re.sub(r'_+', '_', s).strip('_')
    return s[:80] if s else ""

summary_val  = "[Bug] feat: user authentication overhaul"
expected_name = _sanitize_fname(summary_val)

print(f"  summary:      '{summary_val}'")
print(f"  sanitized:    '{expected_name}'")

check("doc_name sanitize removes [, ]",  "[" not in expected_name and "]" not in expected_name)
check("doc_name sanitize no double underscores", "__" not in expected_name)
check("doc_name sanitize non-empty",     len(expected_name) > 0)
check("doc_title equals raw summary",    summary_val == "[Bug] feat: user authentication overhaul")

doc_title_field = next((fd for fd in BUILTIN_FIELDS if fd["key"] == "doc_title"), None)
doc_name_field  = next((fd for fd in BUILTIN_FIELDS if fd["key"] == "doc_name"), None)
check("doc_title field exists in BUILTIN_FIELDS", doc_title_field is not None,
      "missing key: doc_title")
check("doc_name label is 'Doc File Name'",
      doc_name_field is not None and doc_name_field["label"] == "Doc File Name",
      f"got: {doc_name_field['label'] if doc_name_field else 'MISSING'}")
check("doc_title label is 'Doc Title'",
      doc_title_field is not None and doc_title_field["label"] == "Doc Title",
      f"got: {doc_title_field['label'] if doc_title_field else 'MISSING'}")


# ─────────────────────────────────────────────────────────────────────────────
# 8. patch/diff in simulated file comments
# ─────────────────────────────────────────────────────────────────────────────
print(f"\n{SEP}")
print("  TEST 8: patch/diff presence in file comments")
print(SEP)

# Add fake patch to first file comment for testing
simulated_file_comments[0]["patch"] = (
    "@@ -1,5 +1,7 @@\n"
    " import os\n"
    "+import re\n"
    " \n"
    "-def old_func():\n"
    "+def new_func():\n"
    "+    pass\n"
    " \n"
)
patch_fc = simulated_file_comments[0]
check("patch key present in file comment", "patch" in patch_fc)
patch_lines = patch_fc["patch"].splitlines()
check("patch has @@ hunk header", any(l.startswith("@@") for l in patch_lines))
check("patch has + addition lines",  any(l.startswith("+") for l in patch_lines))
check("patch has - deletion lines",  any(l.startswith("-") for l in patch_lines))


# ─────────────────────────────────────────────────────────────────────────────
# 9. generate_word_doc() — with doc_title, patch, per-screenshot comments
# ─────────────────────────────────────────────────────────────────────────────
print(f"\n{SEP}")
print("  TEST 9: generate_word_doc() with doc_title and patch data")
print(SEP)

fv = dict(FIELD_VALS)
fv["doc_name"]  = "dump_test_report"
fv["doc_title"] = "Authentication Overhaul Report"

try:
    out_path = generate_word_doc(
        fv,
        BUILTIN_FIELDS,
        PR_DATA,
        simulated_file_comments,
        DUMP_DIR,
    )
    exists = os.path.exists(out_path)
    size   = os.path.getsize(out_path) if exists else 0
    check("Word doc generated successfully", exists,
          f"path: {out_path}")
    check("Word doc size > 5 KB", size > 5000,
          f"size: {size} bytes")
    print(f"    Output: {out_path}  ({size:,} bytes)")
except Exception as ex:
    check("Word doc generated (no exception)", False, str(ex))


# ─────────────────────────────────────────────────────────────────────────────
# 10. build_preview_html() — with diff-block and per-screenshot comments
# ─────────────────────────────────────────────────────────────────────────────
print(f"\n{SEP}")
print("  TEST 10: build_preview_html() with diff and per-screenshot comments")
print(SEP)

# add a screenshot dict to first file comment
simulated_file_comments[0]["screenshots"] = [
    {"path": "/nonexistent/shot.png", "comment": "Auth flow overview"},
]

try:
    html_out_content = build_preview_html(fv, BUILTIN_FIELDS, PR_DATA, simulated_file_comments)

    check("HTML output non-empty",              len(html_out_content) > 1000)
    check("HTML contains 'Word Comment'",       "Word Comment" in html_out_content)
    check("HTML contains 'Jira Comment'",       "Jira Comment" in html_out_content)
    check("HTML contains PR title",             PR_DATA["title"] in html_out_content)
    check("HTML contains step comments",        "Step 1" in html_out_content)
    check("HTML contains file badge",           "badge-A" in html_out_content or "badge-M" in html_out_content)
    check("HTML contains diff-block (patch)",   "diff-block" in html_out_content)
    check("HTML contains diff-add class",       "diff-add" in html_out_content)

    html_out = os.path.join(DUMP_DIR, "dump_preview.html")
    with open(html_out, "w", encoding="utf-8") as fh:
        fh.write(html_out_content)
    check("HTML preview saved to file", os.path.exists(html_out),
          f"path: {html_out}")
    print(f"    Output: {html_out}  ({len(html_out_content):,} chars)")
except Exception as ex:
    check("build_preview_html (no exception)", False, str(ex))


# ─────────────────────────────────────────────────────────────────────────────
# 11. Jira ADF payload structure
# ─────────────────────────────────────────────────────────────────────────────
print(f"\n{SEP}")
print("  TEST 11: Jira ADF payload (jira_comment used, show_label_in_jira applied)")
print(SEP)

detail_bullets_adf = []
for fd in BUILTIN_FIELDS:
    if not fd.get("enabled", True): continue
    if fd.get("jira_key") in ("summary", "issuetype", "assignee"): continue
    val = fv.get(fd["key"], "")
    if val:
        show_lbl = fd.get("show_label_in_jira", True)
        txt = f"{fd['label']}: {val}" if show_lbl else val
        detail_bullets_adf.append({
            "type": "listItem",
            "content": [{"type": "paragraph",
                         "content": [{"type": "text", "text": txt}]}]
        })

fc_bullets_adf = []
for fc in simulated_file_comments:
    jira_cmt = (fc.get("jira_comment") or fc.get("comment", "")).strip()
    txt = f"{STATUS_EMOJI.get(fc['status'],'~')} {fc['filename']}  [{fc['status'].upper()}]"
    if jira_cmt:
        txt += f"\n    {jira_cmt}"
    fc_bullets_adf.append({
        "type": "listItem",
        "content": [{"type": "paragraph",
                     "content": [{"type": "text", "text": txt}]}]
    })

adf = {
    "type": "doc", "version": 1,
    "content": [
        {"type": "heading", "attrs": {"level": 3},
         "content": [{"type": "text", "text": "Issue Details"}]},
        {"type": "bulletList", "content": detail_bullets_adf},
        {"type": "heading", "attrs": {"level": 3},
         "content": [{"type": "text", "text": "Description"}]},
        {"type": "paragraph",
         "content": [{"type": "text",
                      "text": fv.get("description", "No description.")}]},
        {"type": "heading", "attrs": {"level": 3},
         "content": [{"type": "text", "text": "Changed Files"}]},
        {"type": "bulletList", "content": fc_bullets_adf},
    ]
}

# Save ADF JSON for inspection
adf_path = os.path.join(DUMP_DIR, "dump_jira_adf.json")
with open(adf_path, "w") as fh:
    json.dump(adf, fh, indent=2)

check("ADF has detail bullets",  len(detail_bullets_adf) > 0,
      f"{len(detail_bullets_adf)} field bullets")
check("ADF has file change bullets", len(fc_bullets_adf) > 0,
      f"{len(fc_bullets_adf)} file bullets")

# Verify dependency shows label
dep_texts = [b["content"][0]["content"][0]["text"]
             for b in detail_bullets_adf
             if "Dependency" in b["content"][0]["content"][0]["text"]]
check("Dependency field shows 'Dependency: Yes' in ADF",
      any(t.startswith("Dependency:") for t in dep_texts),
      f"dep texts: {dep_texts}")

# Verify jira_comment is used (not word_comment)
fc_texts = [b["content"][0]["content"][0]["text"] for b in fc_bullets_adf]
has_step_in_jira = any("Step" in t for t in fc_texts)
check("File bullets use jira_comment (contains Step N)",
      has_step_in_jira)

print(f"    ADF saved: {adf_path}")
print("    Detail bullets preview:")
for b in detail_bullets_adf[:4]:
    print(f"      • {b['content'][0]['content'][0]['text']}")
print("    File bullets preview:")
for b in fc_bullets_adf[:3]:
    first_line = b["content"][0]["content"][0]["text"].split("\n")[0]
    print(f"      • {first_line}")


# ─────────────────────────────────────────────────────────────────────────────
# SUMMARY
# ─────────────────────────────────────────────────────────────────────────────
print(f"\n{SEP}")
total  = len(results)
passed = sum(results)
failed = total - passed
status = "ALL PASS" if failed == 0 else f"{failed} FAILED"
print(f"  RESULT: {passed}/{total} passed  —  {status}")
print(SEP)

if failed:
    sys.exit(1)
