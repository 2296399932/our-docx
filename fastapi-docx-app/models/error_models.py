from pydantic import BaseModel, Field
from typing import Optional, List, Dict, Any

class ErrorResponse(BaseModel):
    """API错误响应模型"""
    message: str = Field(..., description="错误消息")
    error_code: Optional[int] = Field(None, description="错误代码")
    details: Optional[Dict[str, Any]] = Field(None, description="错误详情")
    
    class Config:
        schema_extra = {
            "example": {
                "message": "文件处理失败",
                "error_code": 500,
                "details": {"reason": "无法解析文档格式"}
            }
        } 