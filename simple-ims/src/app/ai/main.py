"""Standalone AI order-entry test service.

Run with:
    uvicorn app.ai.main:app --host 0.0.0.0 --port 8091

Or use the provided start_ai_test.ps1 script.
"""

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.ai.routes import router as ai_router
from app.db import Base, engine, ensure_schema

Base.metadata.create_all(bind=engine)
ensure_schema(engine)

app = FastAPI(title="Simple IMS AI Test Service")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
app.include_router(ai_router)


@app.get("/health")
def health_check():
    return {"status": "ok", "service": "ai-test"}


@app.get("/")
def root():
    return {
        "service": "Simple IMS AI Test Service",
        "version": "0.1.0",
        "endpoints": {
            "health": "GET /health",
            "create_order": "POST /ai/orders/create (multipart/form-data, field: image)",
            "demo": "GET /ai/demo",
        },
        "note": "Access /health to check status, or use /ai/demo to upload an image.",
    }
