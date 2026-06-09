# 🧪 Agent QA Eval Suite

A self-contained web app for the QA team to run Excel-driven evaluations against
the deployed agent, score each case with an LLM judge, and download a results
workbook — including, for every failure, the agent's own **self-diagnosis** of
why it failed.



---

## Quick start

```bash
# 1. Clone and enter the repo
git clone <repo-url> && cd agent-eval-suite

# 2. Log in to Azure (once; the app reuses this session)
make login          # or: az login

# 3. Start the app
make up              # builds the container and starts it

# 4. Open the UI
#    http://localhost:8080   (or: make open)
```

That's it. In the browser:

1. Confirm the top-right badge shows **✓ logged in** (if not, run `make login`).
2. Pick a **test workbook** (or click **Upload** to add a new `.xlsx`).
3. Click **▶ Run eval**.
4. Watch cases stream in live (PASS/FAIL, score, judge reason, and an expandable
   **agent self-diagnosis** on failures).
5. Click **⬇ Download results .xlsx** when it finishes.

Stop the app with `make down`.

---

## Folders

| Folder         | Purpose                                                            |
|----------------|--------------------------------------------------------------------|
| `test_excels/` | Input workbooks. Drop `.xlsx` files here (or upload via the UI).   |
| `results/`     | Output. Each run writes `results_<workbook>_<timestamp>.xlsx`.     |
| `app/`         | The application (backend + UI). You don't need to touch this.      |

Both folders are mounted into the container, so files you add/receive show up on
your machine too.

---

## How a run works

For each row in the workbook:

1. **Ask the agent** — sends the `User Query` to the deployed `/chat` API.
2. **Judge it** — an LLM judge compares the answer to the row's expected
   intent / route / incidents / output and returns **PASS** or **FAIL** with a
   reason and score.
3. **Diagnose failures** — on FAIL, the agent is re-asked (in the same session)
   *why* it failed; its root-cause explanation is saved to the
   **Agent Failure Reasoning** column.

Runs are **sequential** by design (the agent currently errors under parallel
load). Expect roughly 30–45 s per case.

The results workbook has four sheets: **Summary**, **All Results**, **PASS**,
**FAIL**.

---

## Workbook format

The judge reads these columns (first sheet, or the one named in `DEFAULT_SHEET`):

`Test ID`, `User Query`, `Expected Intent`, `Route To`,
`Data Source (Resolved)`, `STTM Lookup Required?`,
`Expected Incident Numbers Returned`, `Expected Output Summary`

A row needs at least `Test ID` and `User Query` to run. A sample workbook is
included in `test_excels/`.

---

## Configuration

Settings live in `.env` (created from `.env.example` on first `make up`). The
defaults already point at the **dev** agent. To test a different environment,
change `CHAT_API_URL` (and `CHAT_API_SCOPE` if the API's audience differs).

| Variable                       | What it does                                            |
|--------------------------------|---------------------------------------------------------|
| `CHAT_API_URL`                 | Base URL of the agent to test.                          |
| `CHAT_API_SCOPE`               | Entra scope for the agent token (minted from az login). |
| `AZURE_OPENAI_ENDPOINT`        | Azure OpenAI resource used by the judge.                |
| `AZURE_OPENAI_CHAT_DEPLOYMENT` | Judge model deployment name.                            |
| `AZURE_OPENAI_API_KEY`         | Optional. Leave blank to use az login (AAD) for the judge. |
| `APP_PORT`                     | Port for the UI (default 8080).                         |

---

## Authentication

The app reuses your local `az login` — no tokens to paste. Your `~/.azure`
session is mounted into the container, and the backend mints short-lived tokens
from it on demand (auto-refreshing). If the UI shows **Not logged in**, run
`make login` and refresh.

> The judge uses the same login by default (AAD). This requires your account to
> have the **Cognitive Services OpenAI User** role on the Azure OpenAI resource.
> If you don't have that, set `AZURE_OPENAI_API_KEY` in `.env` instead.

---

## Make targets

```
make up        Build and start the app
make down      Stop the app
make login     az login
make logs      Follow logs
make open      Open the UI
make restart   Restart the app
```
