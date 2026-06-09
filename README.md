# 🧪 Agent QA Eval Suite

A self-contained web app to run Excel-driven evaluations against the deployed
agent, score each case with an LLM judge, and download a results workbook —
including, for every failure, the agent's own **self-diagnosis** of why it failed.

---

## ✅ Easiest way to run it (no make, no Docker)

For anyone — including non-developers. You only need two things installed once:

1. **Azure CLI** — https://learn.microsoft.com/cli/azure/install-azure-cli
2. **Python 3** — https://www.python.org/downloads/ (on Windows, tick *“Add
   python.exe to PATH”* during install)

Then:

| Your computer | What to do |
|---------------|------------|
| **macOS**     | Double-click **`run.command`** |
| **Windows**   | Double-click **`run.bat`** |

The first launch sets everything up (about a minute) and signs you in to Azure
in your browser. After that it starts in seconds and opens
**http://localhost:8080** automatically. Keep the little window open while you
use the app; close it to stop.

> macOS may say *“cannot be opened because it is from an unidentified
> developer.”* Right-click `run.command` → **Open** → **Open** (once).

In the browser:

1. Check the top-right badge shows **✓ signed in** (if not, the launcher will run
   `az login` for you).
2. Pick a **test workbook** (or **Upload** a new `.xlsx`).
3. Click **▶ Run eval** and watch cases stream in live.
4. Click **⬇ Download results .xlsx** when it finishes.

---

## 🛠 Developer way (Docker + make)

```bash
az login            # once
make up             # build + start the container
make open           # http://localhost:8080
make down           # stop
```

---

## How a run works

For each row in the workbook:

1. **Ask the agent** — sends the `User Query` to the agent's `/chat` API
   (authenticated as you, via `az login`).
2. **Judge it** — an LLM judge compares the answer to the row's expected
   intent / route / incidents / output and returns **PASS** or **FAIL** with a
   reason and score.
3. **Diagnose failures** — on FAIL, the agent is re-asked *why* it failed; its
   root-cause explanation is saved to the **Agent Failure Reasoning** column.

Runs are **sequential** by design (the agent currently errors under parallel
load). Expect roughly 30–45 s per case. The results workbook has four sheets:
**Summary**, **All Results**, **PASS**, **FAIL**.

---

## Folders

| Folder         | Purpose                                                          |
|----------------|------------------------------------------------------------------|
| `test_excels/` | Input workbooks. Drop `.xlsx` files here (or upload via the UI). |
| `results/`     | Output. Each run writes `results_<workbook>_<timestamp>.xlsx`.   |
| `app/`         | The application (backend + UI). You don't need to touch this.    |

---

## Workbook format

The judge reads these columns (first sheet, or the one named in `DEFAULT_SHEET`):
`Test ID`, `User Query`, `Expected Intent`, `Route To`,
`Data Source (Resolved)`, `STTM Lookup Required?`,
`Expected Incident Numbers Returned`, `Expected Output Summary`.

A row needs at least `Test ID` and `User Query` to run. A sample workbook is in
`test_excels/`.

---

## Configuration

Settings live in `.env` (copied from `.env.example` on first run). Defaults point
at the dev agent; change `CHAT_API_URL` (and `CHAT_API_SCOPE` if the audience
differs) to test another environment.

| Variable                       | What it does                                              |
|--------------------------------|-----------------------------------------------------------|
| `CHAT_API_URL`                 | Base URL of the agent to test.                            |
| `CHAT_API_SCOPE`               | Entra scope for the agent token (minted from `az login`). |
| `AZURE_OPENAI_ENDPOINT`        | Azure OpenAI resource used by the judge.                  |
| `AZURE_OPENAI_CHAT_DEPLOYMENT` | Judge model deployment name.                              |
| `AZURE_OPENAI_API_KEY`         | Optional. Leave blank to use `az login` (AAD).            |
| `APP_PORT`                     | Port for the UI (default 8080).                           |

---

## Authentication

Everything uses your **`az login`** — one sign-in covers both the agent and the
judge. The launcher runs it for you; if the badge ever shows *not signed in*,
just run `az login` again.

> The judge uses `az login` (AAD) by default, which needs the **Cognitive
> Services OpenAI User** role on the Azure OpenAI resource. If you don't have it,
> set `AZURE_OPENAI_API_KEY` in `.env` instead.

---

## Troubleshooting

- **“Python 3 / Azure CLI is required”** — install the missing one (links above)
  and re-launch.
- **Badge says not signed in** — run `az login` (or just re-launch).
- **`401`/`403` on a case** — your Azure session expired; `az login` again.
- **Want a clean reinstall** — delete the `.venv` folder and re-launch.
