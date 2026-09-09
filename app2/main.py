# 新区域打卡统计 Web 后端（衍生自 app/main.py，功能与接口保持一致）
# 差异：PREFIX=/daka2、独立 Cookie/token、调用 processor2/config2、overview 列名映射

import io
import os
import shutil
import threading
import uuid
from contextlib import redirect_stdout, redirect_stderr
from secrets import compare_digest

from fastapi import FastAPI, UploadFile, File, HTTPException, Query, Request, APIRouter
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse

import processor2 as processor
import config2 as config

# ============ 配置 ============
SITE_PASSWORD = os.environ.get("SITE_PASSWORD", "")
COOKIE_NAME = "daka2_auth"
AUTH_TOKEN = "daka2_auth_ok"
PREFIX = "/daka2"
# ==============================

if not SITE_PASSWORD:
    print("警告：未设置 SITE_PASSWORD 环境变量，登录功能不可用（在 .env 或 compose environment 中配置）")

app = FastAPI(title="打卡数据整理助手 API（二区）")
router = APIRouter(prefix=PREFIX)

from fastapi.middleware.cors import CORSMiddleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")

# 线程锁，防止 os.chdir 并发冲突
_chdir_lock = threading.Lock()

# ============ 自动清理过期文件（7天） ============
CLEANUP_MAX_AGE_DAYS = 7


def _cleanup_old_dirs(base_dir: str):
    """删除超过 N 天的目录"""
    import time
    now = time.time()
    if not os.path.exists(base_dir):
        return
    for name in os.listdir(base_dir):
        dir_path = os.path.join(base_dir, name)
        if not os.path.isdir(dir_path):
            continue
        try:
            mtime = os.path.getmtime(dir_path)
            if (now - mtime) > CLEANUP_MAX_AGE_DAYS * 86400:
                shutil.rmtree(dir_path, ignore_errors=True)
        except OSError:
            pass


@app.on_event("startup")
async def startup_cleanup():
    _cleanup_old_dirs(UPLOAD_DIR)
    _cleanup_old_dirs(OUTPUT_DIR)


# ============ 根路径返回空 ============
@app.get("/")
async def root():
    return JSONResponse({"status": "ok"}, status_code=200)


# ============ 密码认证中间件 ============
def check_auth(request: Request) -> bool:
    token = request.cookies.get(COOKIE_NAME, "")
    if token and compare_digest(token, AUTH_TOKEN):
        return True
    auth = request.headers.get("Authorization", "")
    if auth and compare_digest(auth, f"Bearer {AUTH_TOKEN}"):
        return True
    q_token = request.query_params.get("token", "")
    if q_token and compare_digest(q_token, AUTH_TOKEN):
        return True
    return False


@app.middleware("http")
async def auth_middleware(request: Request, call_next):
    path = request.url.path

    # 不属于 /daka2 的路径直接放行
    if not path.startswith(PREFIX):
        return await call_next(request)

    # 登录接口不拦截
    if path == f"{PREFIX}/api/login":
        return await call_next(request)

    # 页面和静态资源不拦截（前端自己判断登录状态）
    # upload 用 cookie 认证（FormData 不能带自定义 header）
    if path == f"{PREFIX}/" or path == f"{PREFIX}/api/upload" or path.endswith(('.css', '.js', '.ico', '.png', '.jpg', '.html')):
        return await call_next(request)

    # 已登录放行
    if check_auth(request):
        return await call_next(request)

    # 未登录：API 返回 401
    return JSONResponse(status_code=401, content={"detail": "未登录"})


# ============ 登录/登出 ============
@router.post("/api/login")
async def login(request: Request):
    body = await request.json()
    password = body.get("password", "")
    if not SITE_PASSWORD:
        return JSONResponse({"ok": False, "detail": "站点未设置密码（SITE_PASSWORD）"}, status_code=500)
    if compare_digest(password, SITE_PASSWORD):
        resp = JSONResponse({"ok": True})
        resp.set_cookie(COOKIE_NAME, AUTH_TOKEN, httponly=True, max_age=86400 * 30, path="/")
        return resp
    return JSONResponse({"ok": False, "detail": "密码错误"}, status_code=401)


@router.post("/api/logout")
async def logout():
    resp = JSONResponse({"ok": True})
    resp.delete_cookie(COOKIE_NAME, path="/")
    return resp


# ============ 页面路由 ============
@router.get("/")
async def index():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))


@router.get("/login.html")
async def login_page():
    return FileResponse(os.path.join(STATIC_DIR, "login.html"))


# ============ API 路由 ============
@router.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="仅支持 .xlsx / .xls 文件")

    file_id = uuid.uuid4().hex[:12]
    file_dir = os.path.join(UPLOAD_DIR, file_id)
    os.makedirs(file_dir, exist_ok=True)

    file_path = os.path.join(file_dir, file.filename)
    with open(file_path, "wb") as f:
        content = await file.read()
        f.write(content)

    return {"file_id": file_id, "filename": file.filename}


@router.post("/api/process")
async def process_file(file_id: str = Query(...), show_missing: bool = Query(True)):
    file_dir = os.path.join(UPLOAD_DIR, file_id)
    if not os.path.exists(file_dir):
        raise HTTPException(status_code=404, detail="文件不存在，请先上传")

    xlsx_files = [f for f in os.listdir(file_dir) if f.endswith(('.xlsx', '.xls'))]
    if not xlsx_files:
        raise HTTPException(status_code=404, detail="目录中没有 Excel 文件")

    filename = xlsx_files[0]
    file_path = os.path.join(file_dir, filename)

    output_dir = os.path.join(OUTPUT_DIR, file_id)
    os.makedirs(output_dir, exist_ok=True)

    stdout_buffer = io.StringIO()
    stderr_buffer = io.StringIO()
    error_msg = None

    def run_processor():
        with _chdir_lock:
            original_dir = os.getcwd()
            try:
                os.chdir(output_dir)
                with redirect_stdout(stdout_buffer), redirect_stderr(stderr_buffer):
                    processor.main(file_path, show_missing=show_missing)
            finally:
                os.chdir(original_dir)

    try:
        loop = __import__('asyncio').get_event_loop()
        await loop.run_in_executor(None, run_processor)
    except Exception as e:
        error_msg = str(e)

    logs = stdout_buffer.getvalue()
    if stderr_buffer.getvalue():
        logs += "\n" + stderr_buffer.getvalue()

    image_filename = None
    for f in os.listdir(output_dir):
        if f.endswith('.png'):
            image_filename = f
            break

    html_filename = None
    for f in os.listdir(output_dir):
        if f.endswith('.html'):
            html_filename = f
            break

    return {
        "file_id": file_id,
        "logs": logs,
        "image_url": f"{PREFIX}/api/image/{file_id}/{image_filename}" if image_filename else None,
        "html_url": f"{PREFIX}/api/html/{file_id}/{html_filename}" if html_filename else None,
        "error": error_msg,
    }


@router.get("/api/image/{file_id}/{image_name}")
async def get_image(file_id: str, image_name: str):
    image_path = os.path.join(OUTPUT_DIR, file_id, image_name)
    if not os.path.exists(image_path):
        raise HTTPException(status_code=404, detail="图片不存在")
    return FileResponse(image_path, media_type="image/png", filename=image_name)


@router.get("/api/html/{file_id}/{html_name}")
async def get_html(file_id: str, html_name: str):
    html_path = os.path.join(OUTPUT_DIR, file_id, html_name)
    if not os.path.exists(html_path):
        raise HTTPException(status_code=404, detail="HTML 不存在")
    return FileResponse(html_path, media_type="text/html", filename=html_name)


@router.get("/api/overview/{file_id}")
async def get_overview(file_id: str):
    file_dir = os.path.join(UPLOAD_DIR, file_id)
    if not os.path.exists(file_dir):
        raise HTTPException(status_code=404, detail="文件不存在")

    xlsx_files = [f for f in os.listdir(file_dir) if f.endswith(('.xlsx', '.xls'))]
    if not xlsx_files:
        raise HTTPException(status_code=404, detail="没有 Excel 文件")

    file_path = os.path.join(file_dir, xlsx_files[0])

    try:
        df, cols = processor.read_excel_overview(file_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"读取失败: {str(e)}")

    display_names = {
        config.COL_NAME: '姓名',
        config.COL_ID: '工号',
        config.COL_START: '上班时间',
        config.COL_END: '下班时间',
        config.COL_START_V2: '上班时间',
        config.COL_END_V2: '下班时间',
        config.COL_LOCATION: '工作地点',
        config.COL_HANDOVER: '4p接班',
        config.COL_SHIFT_TODAY: '今日班次',
        config.COL_TOMORROW: '明日工作地点',
    }

    import pandas as pd
    rows = []
    for idx, row in df.iterrows():
        row_data = {"_idx": idx}
        for col in cols:
            row_data[col] = str(row[col]) if pd.notna(row[col]) else ""
        rows.append(row_data)

    headers = [{"key": c, "label": display_names.get(c, c)} for c in cols]
    return {"headers": headers, "rows": rows, "columns": cols}


@router.post("/api/save/{file_id}")
async def save_changes(file_id: str, body: dict = None):
    if body is None:
        raise HTTPException(status_code=400, detail="请求体为空")

    file_dir = os.path.join(UPLOAD_DIR, file_id)
    if not os.path.exists(file_dir):
        raise HTTPException(status_code=404, detail="文件不存在")

    xlsx_files = [f for f in os.listdir(file_dir) if f.endswith(('.xlsx', '.xls'))]
    if not xlsx_files:
        raise HTTPException(status_code=404, detail="没有 Excel 文件")

    file_path = os.path.join(file_dir, xlsx_files[0])

    changes = body.get("changes", {})
    new_rows = body.get("new_rows", [])

    if not changes and not new_rows:
        return {"message": "没有修改", "saved": 0}

    try:
        df, cols = processor.read_excel_overview(file_path)

        logs = []
        if changes:
            def log_fn(text):
                logs.append(text)
            processor.save_excel_overview(file_path, changes, df, log=log_fn)

        for row_data in new_rows:
            def log_fn2(text):
                logs.append(text)
            processor.patch_excel_data(file_path, row_data, log=log_fn2)

        total = len(changes) + len(new_rows)
        return {"message": f"已保存 {total} 条修改", "saved": total, "logs": logs}

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"保存失败: {str(e)}")


# 注册路由
app.include_router(router)
