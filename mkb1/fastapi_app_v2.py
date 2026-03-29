from __future__ import annotations

import os
from typing import Optional

from fastapi import FastAPI, Header
from fastapi.responses import JSONResponse
from pydantic import BaseModel, ConfigDict, Field

import main
from task_runner import process_tasks


app = FastAPI(title="Document Extraction Service", version="2.1.0")
INPUT_FIELDS = ("zipName", "folderName", "fileName", "fileFetchPath")


class TaskPayload(BaseModel):
    model_config = ConfigDict(extra="forbid")
    zipName: list[str] = Field(...)
    folderName: list[str] = Field(...)
    fileName: list[str] = Field(...)
    fileFetchPath: list[str] = Field(...)


class RunRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    tasks: list[TaskPayload] | None = None
    zipName: list[str] | None = None
    folderName: list[str] | None = None
    fileName: list[str] | None = None
    fileFetchPath: list[str] | None = None

    def to_payload(self) -> dict:
        if self.tasks is not None:
            return {"tasks": [task.model_dump() for task in self.tasks]}
        return {field: getattr(self, field) for field in INPUT_FIELDS}


def _check_auth(x_api_token: Optional[str]) -> Optional[JSONResponse]:
    expected = os.getenv("API_TOKEN", "").strip()
    if expected and x_api_token != expected:
        return JSONResponse(status_code=401, content={"detail": "Unauthorized"})
    return None


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/run")
def run_job(req: RunRequest, x_api_token: Optional[str] = Header(default=None, alias="X-API-Token")):
    auth_error = _check_auth(x_api_token)
    if auth_error is not None:
        return auth_error
    try:
        return process_tasks(req.to_payload(), config_path="./config.json")
    except ValueError as exc:
        return JSONResponse(status_code=400, content={"status": "fail", "detail": str(exc)})
    except Exception:
        size = len(req.zipName or []) if req.tasks is None else len(req.tasks)
        fail_payload = main.build_fail_output(size)
        return JSONResponse(status_code=500, content=fail_payload)
