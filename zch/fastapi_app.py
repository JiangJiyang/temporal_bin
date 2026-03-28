from __future__ import annotations

import os
from typing import Optional

from fastapi import FastAPI, Header
from fastapi.responses import JSONResponse
from pydantic import BaseModel, ConfigDict, Field

import main


app = FastAPI(title="Document Extraction Service", version="2.0.0")


class RunRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    zipName: list[str] = Field(...)
    folderName: list[str] = Field(...)
    fileName: list[str] = Field(...)
    fileFetchPath: list[str] = Field(...)


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
    payload = req.model_dump()
    try:
        return main.process_request(payload, config_path="./config.json", persist_result=False)
    except Exception:
        fail_payload = main.build_fail_output(len(req.zipName))
        return JSONResponse(status_code=500, content=fail_payload)
