from fastapi import FastAPI
from app.db import engine, Base, SessionLocal, ensure_schema
from app.web.routes import home, owners, products, orders, stock, inbounds, outbounds
from app.ai.routes import router as ai_router
from app.services.owners import OwnerService

Base.metadata.create_all(bind=engine)
ensure_schema(engine)

db = SessionLocal()
try:
    OwnerService.seed_defaults(db)
finally:
    db.close()

app = FastAPI()
app.include_router(home.router)
app.include_router(owners.router)
app.include_router(products.router)
app.include_router(orders.router)
app.include_router(stock.router)
app.include_router(inbounds.router)
app.include_router(outbounds.router)
app.include_router(ai_router)


@app.get("/health")
def health_check():
    return {"status": "ok", "service": "simple-ims"}
