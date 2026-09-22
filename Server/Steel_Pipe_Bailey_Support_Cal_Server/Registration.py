import os
import uuid
import json
import requests
from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
import sqlite3
from datetime import datetime, timedelta
from pydantic import BaseModel

app = FastAPI()
templates = Jinja2Templates(directory="templates")
DB_FILE = "license.db"

def init_db():
    try:
        conn = sqlite3.connect("license.db")
        cursor = conn.cursor()
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS licenses (
            username TEXT PRIMARY KEY,
            license_key TEXT NOT NULL,
            register_time TEXT NOT NULL,
            expire_time TEXT NOT NULL,
            device_id TEXT,
            created_at TEXT
        )
        """)
        conn.commit()
    except Exception as e:
        print("初始化数据库失败:", e)
    finally:
        conn.close()

init_db()

# 检查用户是否存在
def user_exists(email):
    try:
        conn = sqlite3.connect(DB_FILE)
        c.execute("SELECT 1 FROM licenses WHERE username=?", (email,))
        c.execute("SELECT 1 FROM users WHERE email=?", (email,))
        result = c.fetchone()
        return result is not None
    except Exception as e:
        print("查询用户失败:", e)
        return False
    finally:
        conn.close()

# 创建用户记录
def create_user(email, license_key, device_id, expire_time = 365 ):
    now = datetime.now()
    expire = now + timedelta(days=expire_time)
    try:
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("""
            INSERT INTO users (email, license_key, register_time, expire_time, device_id)
            VALUES (?, ?, ?, ?, ?)
        """, (email, license_key, now.strftime("%Y-%m-%d"), expire.strftime("%Y-%m-%d"), None))
        conn.commit()
    except Exception as e:
        print("创建用户失败:", e)
    finally:
        conn.close()

# 获取用户记录
def get_user_record(email):
    try:
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT email, license_key, device_id FROM users WHERE email=?", (email,))
        result = c.fetchone()
        return result
    except Exception as e:
        print("获取用户记录失败:", e)
    finally:
        conn.close()

# 绑定社保
def bind_device(email, device_id):
    try:
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("UPDATE users SET device_id=? WHERE email=?", (device_id, email))
        conn.commit()
    except Exception as e:
        print("绑定设备失败:", e)
    finally:
        conn.close()

# 注册请求
class LicenseRequest(BaseModel):
    email: str
    license_key: str
    device_id: str

# 检查服务器接口状态
@app.get("/ping")
def ping():
    return {"status": "ok"}

# 注册页GET路由
@app.get("/register", response_class=HTMLResponse)
async def register_get(request: Request):
    return templates.TemplateResponse("register.html", {"request": request})

# 注册页POST路由
@app.post("/register", response_class=HTMLResponse)
async def register_post(request: Request, email: str = Form(...)):
    # 简单验证邮箱格式...
    license_key = str(uuid.uuid4()).upper()
    now = datetime.now()
    expire = now + timedelta(days=365)
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT OR IGNORE INTO users(email, license_key, register_time, expire_time, device_id)
        VALUES (?, ?, ?, ?, NULL)
    """, (email, license_key, now.strftime("%Y-%m-%d"), expire.strftime("%Y-%m-%d")))
    conn.commit()
    conn.close()

    return templates.TemplateResponse("register.html", {
        "request": request,
        "success": True,
        "email": email,
        "license_key": license_key,
        "expire": expire.strftime("%Y-%m-%d")
    })

# 验证注册码
@app.post("/verify_license")
async def verify_license(data: LicenseRequest):
    email = data.email.strip()
    license_key = data.license_key.strip()
    device_id = data.device_id.strip()

    if not email or not license_key or not device_id:
        return JSONResponse(status_code=400, content={"msg": "信息不完整"})

    record = get_user_record(email)
    if not record:
        return JSONResponse(status_code=401, content={"msg": "用户不存在"})

    try:
        stored_email, stored_key, stored_device_id, expire_time = record

        if stored_key != license_key:
            return JSONResponse(status_code=403, content={"msg": "密钥错误"})

        if datetime.strptime(expire_time, "%Y-%m-%d") < datetime.now():
            return JSONResponse(status_code=403, content={"msg": "密钥已过期"})

        if stored_device_id is None:
            bind_device(email, device_id)
            return JSONResponse(status_code=200, content={"msg": "验证成功，设备已绑定", "status": "success"})

        if stored_device_id != device_id:
            return JSONResponse(status_code=403, content={"msg": "当前设备未绑定，密钥无效"})

        return JSONResponse(status_code=200, content={"msg": "验证成功，设备已匹配", "status": "success"})

    except Exception as e:
        return JSONResponse(status_code=500, content={"msg": f"服务器错误: {str(e)}"})

# 定期检查服务状态
@app.get("/heartbeat")
async def heartbeat():
    return JSONResponse(content={
        "status": "alive",
        "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    })