"""Aplicação FastAPI — AI DOCX Editor."""

from __future__ import annotations

import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles

from app import __version__
from app.api import api_router
from app.config import ensure_directories, settings
from app.exceptions import AppError
from app.models.responses import ApiResponse

logger = logging.getLogger("aidoc")


def configure_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
    )


@asynccontextmanager
async def lifespan(_app: FastAPI):
    configure_logging()
    ensure_directories()
    logger.info("server_start version=%s", __version__)
    yield
    logger.info("server_stop")


app = FastAPI(
    title=settings.app_name,
    version=__version__,
    description="Motor profissional de edição de documentos DOCX controlável por IA.",
    lifespan=lifespan,
)


@app.exception_handler(AppError)
async def app_error_handler(_request: Request, exc: AppError) -> JSONResponse:
    logger.error("app_error code=%s message=%s", exc.code, exc.message)
    body = ApiResponse.fail(exc.code, exc.message)
    return JSONResponse(status_code=exc.status, content=body.model_dump())


@app.exception_handler(RequestValidationError)
async def validation_handler(_request: Request, exc: RequestValidationError) -> JSONResponse:
    body = ApiResponse.fail("VALIDATION_ERROR", str(exc.errors()))
    return JSONResponse(status_code=422, content=body.model_dump())


@app.exception_handler(Exception)
async def unhandled_handler(_request: Request, exc: Exception) -> JSONResponse:
    logger.exception("unhandled_error")
    body = ApiResponse.fail("INTERNAL_ERROR", "An unexpected error occurred.")
    return JSONResponse(status_code=500, content=body.model_dump())


app.include_router(api_router)

ensure_directories()
app.mount("/previews", StaticFiles(directory=str(settings.previews_dir)), name="previews")


@app.get("/health")
def health() -> ApiResponse:
    return ApiResponse.ok(data={"status": "ok", "version": __version__})


def run() -> None:
    import uvicorn

    uvicorn.run("app.main:app", host=settings.host, port=settings.port, reload=False)


if __name__ == "__main__":
    run()
