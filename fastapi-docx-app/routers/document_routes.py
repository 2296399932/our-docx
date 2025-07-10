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
import traceback
import logging
import uuid

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
    logger = logging.getLogger(__name__)
    logger.info(f"接收到文件: {file.filename}")
    if not file.filename.endswith(('.docx')):
        raise HTTPException(status_code=400, detail="只接受Word文档(.docx)格式")

    try:
        # 1. 保存文件
        logger.info(f"开始保存文件: {file.filename}")
        content = await file.read()
        document_response = await DocumentService.save_uploaded_file(
            file_content=content,
            filename=file.filename,
            upload_dir=UPLOAD_DIR
        )
        logger.info(f"文件保存成功: {document_response.file_path}")
        
        # 2. 立即分析文件
        logger.info(f"开始分析文件: {document_response.file_path}")
        analysis_result = await DocumentService.process_document(document_response.file_path)
        logger.info("文件分析完成")
        
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
        logger.error(f"文件处理失败: {str(e)}", exc_info=True)
        error_traceback = traceback.format_exc()
        logger.error(f"错误详情:\n{error_traceback}")
        # 打印到控制台以便在服务器日志中查看
        print(f"ERROR: 文件处理失败: {str(e)}")
        print(f"ERROR: 错误详情:\n{error_traceback}")
        # 返回简化的错误信息，避免过长
        error_message = str(e) if str(e) else "未知错误"
        raise HTTPException(status_code=500, detail=f"文件处理失败: {error_message}")

@router.post("/export/", response_class=FileResponse)
async def export_document(request_data: Dict[str, Any]):
    """导出文档为DOCX格式"""
    logger = logging.getLogger(__name__)
    logger.info("接收到导出文档请求")

    try:
        content = request_data.get("content")
        options = request_data.get("options", {})
        format_type = options.get("format", "docx")
        file_name = options.get("fileName", f"document-export-{datetime.now().strftime('%Y%m%d%H%M%S')}.{format_type}")

        if not content:
            raise HTTPException(status_code=400, detail="文档内容不能为空")

        logger.info(f"开始导出文档，格式: {format_type}")

        # 调用DocumentService导出文档
        output_path = await DocumentService.export_document(
            content=content,
            format_type=format_type,
            original_file_path=request_data.get("original_file_path"),  # 添加这个参数
            output_dir=OUTPUT_DIR
        )

        

        # 获取文件名并设置正确的扩展名
        actual_ext = os.path.splitext(output_path)[1].lstrip('.')
        download_filename = file_name.replace(f".{format_type}", f".{actual_ext}")
        
        logger.info(f"文档导出成功，提供下载: {output_path}，下载文件名: {download_filename}")
        
        # 返回文件响应
        return FileResponse(
            path=output_path,
            filename=download_filename,
            media_type="application/json" if actual_ext == "json" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        logger.error(f"导出文档失败: {str(e)}", exc_info=True)
        error_traceback = traceback.format_exc()
        logger.error(f"错误详情:\n{error_traceback}")
        print(f"ERROR: 导出文档失败: {str(e)}")
        print(f"ERROR: 错误详情:\n{error_traceback}")
        raise HTTPException(status_code=500, detail=f"导出文档失败: {str(e)}")




