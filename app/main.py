# app/main.py
from fastapi import FastAPI
from starlette.middleware.sessions import SessionMiddleware
from fastapi.middleware.cors import CORSMiddleware
from app.api.remote_trigger import router as remote_trigger_router
from app.api.tower_manager import router as tower_manager_router

# ===== FastAPI App 设置 =====
app = FastAPI(title="TOWER")

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        'http://10.8.3.32:5001',
        'http://clound.gxtower.cn:3980'
    ],
    allow_credentials=True,
    allow_methods=["GET", "POST", "OPTIONS", "DELETE", "PUT"],
    allow_headers=["X-Role", "Content-Type", "Accept", "Origin"],
)

# Session
app.add_middleware(
    SessionMiddleware,
    secret_key="a1b2c3d4e5f678901234567890abcdef1234567890abcdef1234567890abcdef",
    session_cookie="session",
    same_site="lax",
)

# Routers
app.include_router(remote_trigger_router)
app.include_router(tower_manager_router)
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app.main:app", host="0.0.0.0", port=38114, reload=True)