import uvicorn

if __name__ == "__main__":
    print("启动文档处理API服务...")
    uvicorn.run(
        "main:app", 
        host="0.0.0.0", 
        port=8000, 
        reload=True,
        log_level="info"
    ) 