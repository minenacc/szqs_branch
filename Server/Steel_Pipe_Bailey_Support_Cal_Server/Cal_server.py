import os
import uuid
import json
import pickle
import requests
from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
import threading
import rsa
import base64
import sqlite3
from datetime import datetime, timedelta
from pydantic import BaseModel
from email_send import send_license_email
import re
from pathlib import Path

TASK_FOLDER = os.path.join(os.path.dirname(__file__), "tasks")
INDEX_FILE = os.path.join(TASK_FOLDER, "task_index.pkl")
os.makedirs(TASK_FOLDER, exist_ok=True)
index_lock = threading.Lock()
POLL_INTERVAL = 5  # 每5秒检查一次是否有新任务

app = FastAPI()
# 基础目录设为当前脚本所在目录
BASE_DIR = Path(__file__).resolve().parent
# templates 目录设置
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
# 数据库文件的绝对路径
DB_FILE = BASE_DIR / "license.db"

# 配置FastAPI应用程序，允许跨域请求
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# 服务端地址
server_url = "http://10.142.35.199:8000" 

# 初始化数据库
def init_db():
    try:
        conn = sqlite3.connect(str(DB_FILE))
        cursor = conn.cursor()
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            email TEXT PRIMARY KEY,
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

# 保存单个任务
def save_task(task):
    task_file = os.path.join(TASK_FOLDER, f"task_{task['task_id']}.pkl")
    with open(task_file, "wb") as f:
        pickle.dump(task, f)

# 加载单个任务
def load_task(task_id):
    task_file = os.path.join(TASK_FOLDER, f"task_{task_id}.pkl")
    if os.path.exists(task_file):
        with open(task_file, "rb") as f:
            return pickle.load(f)
    return None

# 加载任务索引列表
def load_task_index():
    if os.path.exists(INDEX_FILE):
        with open(INDEX_FILE, "rb") as f:
            return pickle.load(f)
    return []

# 保存任务索引
def save_task_index(index_list):
    with open(INDEX_FILE, "wb") as f:
        pickle.dump(index_list, f)

# 添加任务到索引（上传任务时调用）
def add_task_id(task_id):
    with index_lock:
        index_list = load_task_index()
        if task_id not in index_list:
            index_list.append(task_id)
            save_task_index(index_list)

# 检查服务器接口状态
@app.get("/ping")
def ping():
    return {"status": "ok"}

# 服务端创建任务对象
@app.post("/submit_task")
async def submit_task(
    mcb_file: UploadFile = File(...),
    txt_file: UploadFile = File(...),
    foundation_name: str = Form(...),
    foundation_standard: str = Form(...),
    foundation_values: str =Form(...)
    ):
    print(f"收到mcb文件: {mcb_file.filename}")
    print(f"收到txt文件: {txt_file.filename}")
    print(f"基础类型: {foundation_name}")
    print(f"标准类型: {foundation_standard}")
    print(f"参数字典: {foundation_values}")
    # 生成任务ID
    task_id = str(uuid.uuid4())
    # 创建临时目录保存本次任务的数据
    temp_dir = os.path.join("task", task_id)
    os.makedirs(temp_dir, exist_ok=True)
    print(f"[任务目录] 创建成功：{temp_dir}")

    # 保存客户端上传的mcb文件至任务目录
    mcb_path = os.path.join(temp_dir, mcb_file.filename)
    with open(mcb_path, "wb") as f:
        f.write(await mcb_file.read())
    print(f"[MCB文件]已保存至：{mcb_path}")
    txt_path = os.path.join(temp_dir, txt_file.filename)
    with open(txt_path, "wb") as f:
        f.write(await txt_file.read())
    print(f"[TXT文件]已保存至：{txt_path}")
    
    # 将前端提交的参数 JSON 字符串转成字典
    try:
        foundation_values_dict = json.loads(foundation_values)
    except:
        foundation_values_dict = {}

    # 定义任务结构，初始化状态为pending
    task = {
        "task_id": task_id,
        "mcb_path": mcb_path,
        "txt_path": mcb_path,
        "foundation_name": foundation_name,
        "foundation_standard": foundation_standard,
        "foundation_values": foundation_values_dict,
        "temp_dir": temp_dir,
        "status": "pending",
        "msg": "任务已提交",
        "docx_path": None,
        "error": None
    }

    # 保存当前任务
    save_task(task)
    add_task_id(task_id)

    return {"task_id": task_id}

# 客户端点击“计算”时调用，改状态为 queued
@app.post("/run_analysis")
async def run_analysis(request: Request):
    # 获取任务ID
    form = await request.form()
    task_id = form.get("task_id")
    # 如果任务是 pending，则转为 queued 状态，等待后台线程拾取并执行
    task = load_task(task_id)
    if not task:
        return JSONResponse(status_code=404, content={"msg": "未找到任务"})
        
    if task["task_id"] == task_id:
        if task["status"] == "pending":
            task["status"] = "queued"
            task["msg"] = "等待执行"
            save_task(task)
            return {"msg": "任务已排队执行"}
        elif task["status"] == "error":
            return JSONResponse(status_code=400, content={"msg": "计算失败，请重新上传模型后再试"})
        elif task["status"] in ("queued", "running"):
            return {"msg": f"计算进行中，请耐心等待"}
        elif task["status"] == "done":
            return {"msg": "任务已完成，无需重复计算"}
        else:
            return JSONResponse(status_code=500, content={"msg": "未知错误"})

# 客户端轮询该接口，查看任务当前进度
@app.post("/check_status")
async def check_status(request: Request):
    form = await request.form()
    task_id = form.get("task_id")
    task = load_task(task_id)
    if not task:
        return JSONResponse(status_code=404, content={"msg": "未找到任务"})

    index_list = load_task_index()
    if task_id not in index_list:
        return JSONResponse(status_code=404, content={"msg": "任务不在队列中"})
    current_index = index_list.index(task_id)
    ahead = sum(1 for i in index_list[:current_index] if load_task(i)["status"] == "queued") + 1

    msg = task["msg"]
    if task["status"] == "queued":
        msg += f"，前方还有 {ahead} 个任务排队"
    return {"status": task["status"], "msg": msg}

# 客户端点击下载时触发
@app.post("/download_docx")
async def download_docx(request: Request):
    form = await request.form()
    task_id = form.get("task_id")
    task = load_task(task_id)
    if not task:
        return JSONResponse(status_code=404, content={"msg": "任务不存在"})
    if task["status"] == "done" and task["docx_path"] and os.path.exists(task["docx_path"]):
        return FileResponse(task["docx_path"], filename=os.path.basename(task["docx_path"]))
    else:
        return JSONResponse(status_code=400, content={"msg": "计算书尚未生成或任务未完成"})


# 定期检查服务状态
@app.get("/heartbeat")
async def heartbeat():
    return JSONResponse(content={
        "status": "alive",
        "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    })

# 检查用户是否存在
def user_exists(email):
    try:
        conn = sqlite3.connect(DB_FILE)
        conn.execute("SELECT 1 FROM users WHERE email=?", (email,))
        result = conn.fetchone()
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
        c.execute("SELECT email, license_key, device_id, expire_time FROM users WHERE email=?", (email,))
        result = c.fetchone()
        return result
    except Exception as e:
        print("获取用户记录失败:", e)
    finally:
        conn.close()

# 绑定设备
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

# 注册页GET路由
@app.get("/register", response_class=HTMLResponse)
async def register_get(request: Request):
    return templates.TemplateResponse("register.html", {"request": request})

# 注册页POST路由
@app.post("/register", response_class=HTMLResponse)
async def register_post(request: Request, email: str = Form(...)):
    # 校验邮箱格式
    if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
        return templates.TemplateResponse("register.html", {
            "request": request,
            "error": "邮箱格式错误"
        })
    print(f"收到注册请求: {email}")
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    # 查询邮箱是否已注册
    cursor.execute("SELECT * FROM users WHERE email = ?", (email,))
    record = cursor.fetchone()

    if record:
        conn.close()
        return templates.TemplateResponse("register.html", {
            "request": request,
            "error": "该邮箱已注册，请勿重复注册。",
            "email": email
        })
    # 注册流程
    license_key = str(uuid.uuid4()).upper()
    now = datetime.now()
    expire_time = now + timedelta(days=365)

    cursor.execute("""
        INSERT OR IGNORE INTO users(email, license_key, register_time, expire_time, device_id)
        VALUES (?, ?, ?, ?, NULL)
    """, (email, license_key, now.strftime("%Y-%m-%d"), expire_time.strftime("%Y-%m-%d")))
    conn.commit()
    conn.close()

    # 发送授权邮件
    success = send_license_email(email, license_key, expire_time.strftime("%Y-%m-%d"))

    if success:
        return templates.TemplateResponse("register.html", {
            "request": request,
            "success": True,
            "email": email,
            "license_key": license_key,
            "expire_time": expire_time.strftime("%Y-%m-%d")
        })
    else:
        return templates.TemplateResponse("register.html", {
            "request": request,
            "error": "邮件发送失败，请稍后联系管理员获取授权密钥。",
            "email": email
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
            print("验证成功，设备已绑定")
        if stored_device_id != device_id:
            return JSONResponse(status_code=403, content={"msg": "已绑定设备，密钥无效"})

        if stored_device_id == device_id and datetime.strptime(expire_time, "%Y-%m-%d") > datetime.now() and stored_key == license_key:
            return JSONResponse(status_code=200, content={"msg": "验证成功", "status": "success"})
            print("验证成功")
            
    except Exception as e:
        return JSONResponse(status_code=500, content={"msg": f"服务器错误: {str(e)}"})