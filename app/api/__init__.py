from fastapi import APIRouter

from app.api.blocks import router as blocks_router
from app.api.documents import router as documents_router
from app.api.rendering import router as rendering_router
from app.api.styles import router as styles_router
from app.api.tables import router as tables_router

api_router = APIRouter()
api_router.include_router(documents_router)
api_router.include_router(blocks_router)
api_router.include_router(tables_router)
api_router.include_router(styles_router)
api_router.include_router(rendering_router)
