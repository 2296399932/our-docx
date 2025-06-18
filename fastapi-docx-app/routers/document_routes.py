from fastapi import APIRouter, UploadFile, File, HTTPException, Depends, BackgroundTasks
from fastapi.responses import JSONResponse, FileResponse
from typing import List, Dict, Any
import os
import shutil
import asyncio
import aiofiles
from aiofiles import os as aio_os
from datetime import datetime
import json

from services import DocumentService


router = APIRouter(
    prefix="/documents",
    tags=["文档处理"],
    responses={404: {"description": "未找到"}},
)

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "output"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

@router.post("/upload/", response_model=Dict[str, Any])
async def upload_document(file: UploadFile = File(...)):
    """上传并分析Word文档"""
    if not file.filename.endswith(('.docx')):
        raise HTTPException(status_code=400, detail="只接受Word文档(.docx)格式")

    try:
        # 1. 保存文件
        content = await file.read()
        document_response = await DocumentService.save_uploaded_file(
            file_content=content,
            filename=file.filename,
            upload_dir=UPLOAD_DIR
        )
        
        # 2. 立即分析文件
        analysis_result = await DocumentService.process_document(document_response.file_path)
        
        # 3. 合并文件信息和分析结果
        result = {
            "file_info": {
                "filename": document_response.filename,
                "status": document_response.status,
                "file_path": document_response.file_path,
                "upload_time": document_response.upload_time
            },
            "analysis": analysis_result
        }

        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"文件处理失败: {str(e)}")

@router.post("/export/", response_class=FileResponse)
async def export_document(request_data: Dict[str, Any]):
    """导出文档为DOCX格式"""
    try:
        content = request_data.get("content")
        options = request_data.get("options", {})
        format_type = options.get("format", "docx")
        file_name = options.get("fileName", f"document-export-{datetime.now().strftime('%Y%m%d%H%M%S')}.{format_type}")
        
        output_path = os.path.join(OUTPUT_DIR, file_name)
        
        # 将内容保存为临时JSON文件
        temp_json_path = os.path.join(OUTPUT_DIR, f"temp_{datetime.now().strftime('%Y%m%d%H%M%S')}.json")
        with open(temp_json_path, "w", encoding="utf-8") as f:
            json.dump(content, f, ensure_ascii=False, indent=2)
        
        # 如果是docx格式，尝试调用服务转换
        if format_type.lower() == "docx":
            # 这里应该调用DocumentService的方法进行转换
            # 由于目前没有实现，我们暂时返回原始JSON文件
            return FileResponse(
                path=temp_json_path,
                filename=file_name.replace(".docx", ".json"),
                media_type="application/json"
            )
        else:
            # 直接返回JSON文件
            return FileResponse(
                path=temp_json_path,
                filename=file_name,
                media_type="application/json"
            )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"导出文档失败: {str(e)}")


