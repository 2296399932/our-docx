from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import uvicorn
import os
import logging
import sys

# 导入自定义模块前确保路径正确
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
# 添加调试信息
print(f"当前工作目录: {os.getcwd()}")
print(f"Python路径: {sys.path}")

# 尝试导入util包中的StyleAnalyzer
try:
    from util.style_analyzer import StyleAnalyzer
    print("成功导入StyleAnalyzer")
except ImportError as e:
    print(f"导入StyleAnalyzer失败: {e}")
    # 如果导入失败，尝试查看util目录内容
    util_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "util")
    if os.path.exists(util_dir):
        print(f"util目录存在，内容: {os.listdir(util_dir)}")
    else:
        print(f"util目录不存在: {util_dir}")

from models.error_models import ErrorResponse
from routers import document_routes

# 导入服务模块
from services.document_service import DocumentService

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()]
)

logger = logging.getLogger(__name__)

# API服务器配置
API_HOST = os.getenv("API_HOST", "localhost")
API_PORT = int(os.getenv("API_PORT", "8000"))
API_PROTOCOL = os.getenv("API_PROTOCOL", "http")

# 图片服务器配置
# 默认与API服务器相同，但可以通过环境变量单独配置
IMAGE_SERVER_URL = os.getenv("IMAGE_SERVER_URL")  # 如果未设置，则使用API_BASE_URL

# 前端服务器配置（用于CORS）
FRONTEND_URL = os.getenv("FRONTEND_URL", "http://localhost:3000")
# 检查是否设置了允许所有来源
ALLOW_ALL_ORIGINS = os.getenv("ALLOW_ALL_ORIGINS", "").lower() in ("true", "1", "yes")
# 默认允许localhost和常见的局域网IP访问
DEFAULT_ALLOWED_ORIGINS = [
    "http://localhost:3000",
    "http://127.0.0.1:3000",
    "http://192.168.0.101:3000"  # 添加访问IP
]
# 允许的前端来源列表，包括默认允许的URLs和环境变量中配置的URLs
EXTRA_ORIGINS = os.getenv("EXTRA_ALLOWED_ORIGINS", "").split(",") if os.getenv("EXTRA_ALLOWED_ORIGINS") else []
ALLOWED_ORIGINS = ["*"] if ALLOW_ALL_ORIGINS else (DEFAULT_ALLOWED_ORIGINS + [origin for origin in EXTRA_ORIGINS if origin])

# 打印CORS配置信息
if ALLOW_ALL_ORIGINS:
    logger.info("CORS配置为允许所有来源")
else:
    logger.info(f"CORS允许的来源: {ALLOWED_ORIGINS}")

# 创建FastAPI应用实例
app = FastAPI(
    title="Word文档处理API",
    description="用于处理和分析Word文档的REST API",
    version="0.1.0"
)

# 设置API基础URL
API_BASE_URL = f"{API_PROTOCOL}://{API_HOST}:{API_PORT}"
DocumentService.set_api_base_url(API_BASE_URL)
logger.info(f"已设置API基础URL为: {API_BASE_URL}")

# 设置图片服务器URL（如果配置了）
if IMAGE_SERVER_URL:
    DocumentService.set_image_server_url(IMAGE_SERVER_URL)
    logger.info(f"已设置图片服务器URL为: {IMAGE_SERVER_URL}")
else:
    logger.info(f"图片服务器URL未单独配置，将使用API基础URL: {API_BASE_URL}")

# 配置CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,  # 使用配置的前端URL列表
    allow_credentials=True,
    allow_methods=["*"],  # 允许所有方法
    allow_headers=["*"],  # 允许所有头部
    expose_headers=["Content-Disposition"],  # 暴露Content-Disposition头，用于文件下载
)

# 注册路由
app.include_router(document_routes.router)
print(f"已注册路由: {[route.path for route in app.routes]}")

# 创建上传目录
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# 创建输出目录
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 创建静态图片目录
STATIC_DIR = "static"
os.makedirs(os.path.join(STATIC_DIR, "images"), exist_ok=True)

# 挂载静态文件目录
app.mount("/images", StaticFiles(directory=os.path.join(STATIC_DIR, "images")), name="images")

# 全局异常处理
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    logger.error(f"全局异常: {str(exc)}", exc_info=True)
    return JSONResponse(
        status_code=500,
        content={"message": f"服务器内部错误: {str(exc)}"}
    )

@app.on_event("startup")
async def startup_event():
    """应用启动时的初始化操作"""
    # 确保各种目录存在
    for directory in [UPLOAD_DIR, OUTPUT_DIR, STATIC_DIR, os.path.join(STATIC_DIR, "images")]:
        os.makedirs(directory, exist_ok=True)
        logger.info(f"确保目录存在: {directory}")
    
    # 重新配置API基础URL（以防启动后端口变更）
    API_BASE_URL = f"{API_PROTOCOL}://{API_HOST}:{API_PORT}"
    DocumentService.set_api_base_url(API_BASE_URL)
    logger.info(f"应用启动，API基础URL: {API_BASE_URL}")
    
    # 重新配置图片服务器URL
    if IMAGE_SERVER_URL:
        DocumentService.set_image_server_url(IMAGE_SERVER_URL)
        logger.info(f"应用启动，图片服务器URL: {IMAGE_SERVER_URL}")
    else:
        logger.info(f"应用启动，图片服务器URL未单独配置，将使用API基础URL")

@app.get("/")
async def root():
    return {"status": "服务正常运行"}

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True) 