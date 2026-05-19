# PR → Jira Ticket Tool

A dark-themed Python Tkinter GUI to fetch GitHub PRs and create Jira tickets with auto-generated Word documents.

---

## 📦 Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Create token files (plain text, one token per file)

```
github_token.txt   ← your GitHub Personal Access Token
jira_token.txt     ← your Jira API Token
```

Place them next to `app.py`, or configure custom paths in Settings.

**GitHub token** — needs `repo` scope:
https://github.com/settings/tokens

**Jira API token**:
https://id.atlassian.com/manage-profile/security/api-tokens

### 3. Run the app

```bash
python app.py
```

---

## ⚙️ First-time Configuration (Settings window)

| Field              | Example                            |
|--------------------|------------------------------------|
| GitHub Token File  | `github_token.txt`                 |
| Jira Token File    | `jira_token.txt`                   |
| Jira Base URL      | `https://yourcompany.atlassian.net`|
| Jira Project Key   | `PROJ`                             |
| Jira Email         | `you@company.com`                  |
| GitHub Owner       | `your-org`                         |
| GitHub Repo        | `your-repo`                        |
| Default Owner      | your username                      |
| Default Environment| `Production`                       |
| Word Doc Output Dir| any folder path                    |

Settings are saved to `config.json` automatically.

---

## 🚀 Workflow

1. Paste a GitHub PR URL or PR number into the **PR URL** field
2. Click **Fetch PR** — fields auto-populate
3. Fill in **Issue Owner**, **Environment**, **Issue Type**, **Dependency**
4. Optionally set a **Word Doc Name** (defaults to timestamp)
5. Click **Generate Word Doc** — saves `.docx` to output folder
6. Click **Create Jira Ticket** — creates ticket, attaches Word doc, shows link

---

## 📁 Files

```
app.py               — main application
requirements.txt     — pip dependencies
config.json          — auto-generated settings (gitignore this)
github_token.txt     — your GitHub PAT  (gitignore this!)
jira_token.txt       — your Jira token  (gitignore this!)
```

> ⚠️ Add `*.txt` and `config.json` to your `.gitignore` to avoid committing tokens.
