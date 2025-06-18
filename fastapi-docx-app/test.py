from services import *
import asyncio

file_path="F:\danzi\our-docx\\1.docx"
# 文档路径

async def main():
    # 立即分析文件
    analysis_result = await DocumentService.process_document(file_path)
    print(analysis_result)

# 运行异步主函数
asyncio.run(main())

