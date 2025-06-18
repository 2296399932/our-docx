from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
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

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()]
)

logger = logging.getLogger(__name__)

# 创建FastAPI应用实例
app = FastAPI(
    title="Word文档处理API",
    description="用于处理和分析Word文档的REST API",
    version="0.1.0"
)

# 配置CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 允许所有来源，生产环境中应该限制为特定域名
    allow_credentials=True,
    allow_methods=["*"],  # 允许所有方法
    allow_headers=["*"],  # 允许所有头部
)

# 注册路由
app.include_router(document_routes.router)

# 创建上传目录
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# 创建输出目录
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 全局异常处理
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    logger.error(f"全局异常: {str(exc)}", exc_info=True)
    return JSONResponse(
        status_code=500,
        content={"message": f"服务器内部错误: {str(exc)}"}
    )

@app.get("/")
async def root():
    return {"status": "服务正常运行"}

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True) 