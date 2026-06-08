"""FastAPI app: serves the QA eval UI and API.

Endpoints:
  GET  /                      -> UI
  GET  /api/config            -> target URL / sheet / port
  GET  /api/auth              -> az login status
  GET  /api/excels            -> list uploaded test workbooks
  POST /api/excels            -> upload a workbook
  POST /api/run               -> start an eval run (background)
  GET  /api/run/{run_id}      -> live status + partial results
  GET  /api/results           -> list result workbooks
  GET  /api/download/{name}   -> download a result workbook
"""
from __future__ import annotations

import json
import os
import subprocess
import threading
import uuid
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from . import eval_runner

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent.parent
TEST_EXCELS = Path(os.getenv("TEST_EXCELS_DIR", BASE_DIR / "test_excels"))
RESULTS_DIR = Path(os.getenv("RESULTS_DIR", BASE_DIR / "results"))
STATIC_DIR = Path(__file__).resolve().parent / "static"
TEST_EXCELS.mkdir(parents=True, exist_ok=True)
RESULTS_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Agent QA Eval Suite")

# In-memory run registry. Runs are sequential and local, so this is sufficient.
RUNS: dict[str, dict[str, Any]] = {}
_LOCK = threading.Lock()


# --------------------------------------------------------------------------- #
# UI
# --------------------------------------------------------------------------- #
@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return (STATIC_DIR / "index.html").read_text(encoding="utf-8")


@app.get("/api/config")
def config() -> dict[str, Any]:
    return {
        "url": (os.getenv("CHAT_API_URL") or "").rstrip("/"),
        "sheet": os.getenv("DEFAULT_SHEET", "Intent Test Cases"),
        "judge_auth": "api-key" if os.getenv("AZURE_OPENAI_API_KEY") else "az-login (AAD)",
    }


# --------------------------------------------------------------------------- #
# Auth status (reuses host az login)
# --------------------------------------------------------------------------- #
@app.get("/api/auth")
def auth_status() -> dict[str, Any]:
    try:
        out = subprocess.run(
            ["az", "account", "show", "-o", "json"],
            capture_output=True, text=True, timeout=20,
        )
        if out.returncode != 0:
            return {"logged_in": False, "detail": (out.stderr or "not logged in").strip()[:200]}
        acct = json.loads(out.stdout)
        return {
            "logged_in": True,
            "user": (acct.get("user") or {}).get("name"),
            "subscription": acct.get("name"),
            "tenant": acct.get("tenantId"),
        }
    except FileNotFoundError:
        return {"logged_in": False, "detail": "Azure CLI not available in this environment."}
    except Exception as exc:  # noqa: BLE001
        return {"logged_in": False, "detail": str(exc)[:200]}


# --------------------------------------------------------------------------- #
# Excel management
# --------------------------------------------------------------------------- #
@app.get("/api/excels")
def list_excels() -> list[dict[str, Any]]:
    items = []
    for p in sorted(TEST_EXCELS.glob("*.xlsx")):
        if p.name.startswith("~$"):
            continue
        try:
            sheets = eval_runner.list_sheets(p)
        except Exception:  # noqa: BLE001
            sheets = []
        items.append({"name": p.name, "sheets": sheets, "size_kb": round(p.stat().st_size / 1024, 1)})
    return items


@app.post("/api/excels")
async def upload_excel(file: UploadFile) -> dict[str, Any]:
    if not file.filename or not file.filename.endswith(".xlsx"):
        raise HTTPException(400, "Only .xlsx files are accepted.")
    safe = Path(file.filename).name
    dest = TEST_EXCELS / safe
    dest.write_bytes(await file.read())
    return {"name": safe}


# --------------------------------------------------------------------------- #
# Runs
# --------------------------------------------------------------------------- #
def _run_worker(run_id: str, excel: str, sheet: str | None, limit: int | None) -> None:
    state = RUNS[run_id]

    def cb(ev: dict[str, Any]) -> None:
        with _LOCK:
            if ev["type"] == "start":
                state.update(status="running", total=ev["total"], url=ev["url"], sheet=ev["sheet"])
            elif ev["type"] == "case":
                row = ev["row"]
                state["done"] = ev["index"]
                state["rows"].append({
                    "test_id": row.get("Test ID"),
                    "verdict": row.get("Judge Verdict"),
                    "score": row.get("Judge Score"),
                    "failure_stage": row.get("Failure Stage"),
                    "query": row.get("User Query"),
                    "judge_reason": row.get("Judge Reason"),
                    "agent_failure_reasoning": row.get("Agent Failure Reasoning"),
                    "answer": row.get("Agent Answer"),
                })
                if row.get("Judge Verdict") == "PASS":
                    state["passed"] += 1
                else:
                    state["failed"] += 1
            elif ev["type"] == "done":
                state.update(status="done", result_file=ev["result_file"], pass_rate=ev["pass_rate"])

    try:
        eval_runner.run_eval(
            TEST_EXCELS / excel, RESULTS_DIR,
            sheet=sheet, limit=limit, progress_cb=cb,
        )
    except Exception as exc:  # noqa: BLE001
        with _LOCK:
            state.update(status="error", error=f"{type(exc).__name__}: {exc}")


@app.post("/api/run")
def start_run(body: dict[str, Any]) -> dict[str, str]:
    excel = body.get("excel")
    if not excel or not (TEST_EXCELS / Path(excel).name).exists():
        raise HTTPException(400, "Unknown excel file. Upload it first.")
    sheet = body.get("sheet") or None
    limit = body.get("limit")
    limit = int(limit) if limit else None

    run_id = uuid.uuid4().hex[:12]
    RUNS[run_id] = {
        "status": "starting", "excel": Path(excel).name, "total": 0, "done": 0,
        "passed": 0, "failed": 0, "rows": [], "result_file": None, "error": None,
    }
    threading.Thread(target=_run_worker, args=(run_id, Path(excel).name, sheet, limit), daemon=True).start()
    return {"run_id": run_id}


@app.get("/api/run/{run_id}")
def run_status(run_id: str) -> JSONResponse:
    state = RUNS.get(run_id)
    if not state:
        raise HTTPException(404, "Unknown run id.")
    return JSONResponse({**state, "run_id": run_id})


@app.get("/api/active")
def active_run() -> JSONResponse:
    """Return the in-progress run (or, if none, the most recent run) so a page
    refresh can reconnect to it instead of losing state."""
    if not RUNS:
        return JSONResponse({})
    in_progress = [rid for rid, s in RUNS.items() if s.get("status") in ("starting", "running")]
    run_id = in_progress[-1] if in_progress else list(RUNS.keys())[-1]
    return JSONResponse({**RUNS[run_id], "run_id": run_id})


@app.get("/api/results")
def list_results() -> list[dict[str, Any]]:
    items = []
    for p in sorted(RESULTS_DIR.glob("*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True):
        items.append({"name": p.name, "size_kb": round(p.stat().st_size / 1024, 1)})
    return items


@app.get("/api/download/{name}")
def download(name: str) -> FileResponse:
    safe = Path(name).name
    path = RESULTS_DIR / safe
    if not path.exists():
        raise HTTPException(404, "Result file not found.")
    return FileResponse(path, filename=safe,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
