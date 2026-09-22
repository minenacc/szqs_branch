import uvicorn
import threading
from Cal_server import app
import Cal_task_executor_new
import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

def start_background_task_executor():
    thread = threading.Thread(target=Cal_task_executor_new.run_task_loop, daemon=True)
    thread.start()

def start_heartbeat_logger():
    def log():
        import time
        while True:
            print("[心跳] 服务正在运行...", flush=True)
            time.sleep(10)
    thread = threading.Thread(target=log, daemon=True)
    thread.start()

if __name__ == "__main__":
    start_background_task_executor()
    start_heartbeat_logger()
    uvicorn.run(app, host="0.0.0.0", port=8000)
