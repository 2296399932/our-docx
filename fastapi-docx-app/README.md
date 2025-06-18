# FastAPI Word文档处理服务

一个基于FastAPI和异步处理技术的Word文档处理RESTful API服务。

## 功能特点

- 使用FastAPI构建高性能异步API
- 支持Word文档(.docx/.doc)上传和管理
- 异步处理文档内容和结构分析
- 使用Pydantic模型进行数据验证和序列化

## 安装

1. 克隆仓库：

```bash
git clone <仓库地址>
cd fastapi-docx-app
```

2. 安装依赖：

```bash
pip install -r requirements.txt
```

## 运行服务

```bash
uvicorn main:app --reload
```

服务将在 http://127.0.0.1:8000 上运行，API文档可在 http://127.0.0.1:8000/docs 访问。

## API端点

- `GET /` - API根端点，返回欢迎信息
- `POST /documents/upload/` - 上传Word文档
- `GET /documents/list/` - 获取已上传文档列表
- `GET /documents/analyze/{filename}` - 分析指定文档
- `DELETE /documents/{filename}` - 删除指定文档

## 项目结构

```
fastapi-docx-app/
├── main.py           # 应用主文件
├── requirements.txt  # 项目依赖
├── uploads/          # 上传文件存储目录
├── routers/          # API路由
│   ├── __init__.py
│   └── document_routes.py
├── models/           # 数据模型
│   ├── __init__.py
│   └── document_models.py
└── services/         # 业务逻辑服务
    ├── __init__.py
    └── document_service.py
```

## 技术栈

- FastAPI - Web框架
- Uvicorn - ASGI服务器
- python-docx - Word文档处理
- Pydantic - 数据验证
- aiofiles - 异步文件I/O 