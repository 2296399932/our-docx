import google.generativeai as genai
import os

# 1. 配置 API Key (从环境变量获取)
try:
    genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
except Exception as e:
    print(f"API Key 配置失败: {e}")
    print("请确保已设置 GOOGLE_API_KEY 环境变量。")
    exit()

# 2. 指定模型名称
MODEL_NAME = "gemini-2.5-flash-preview-05-20"

# 3. 实例化模型
try:
    model = genai.GenerativeModel(MODEL_NAME)
    print(f"模型 '{MODEL_NAME}' 实例化成功。")
except Exception as e:
    print(f"模型 '{MODEL_NAME}' 实例化失败: {e}")
    print("请检查模型名称是否正确，或API Key是否有访问该模型的权限。")
    exit()

# 4. 发送请求
user_prompt = "你好，请简单介绍一下你自己。"

print(f"\nUser: {user_prompt}")
try:
    response = model.generate_content(user_prompt)

    # 5. 处理响应
    if response and response.text:
        print(f"AI: {response.text}")
    else:
        print("AI没有返回文本内容。")
        # 打印原始响应对象以供调试
        print(f"原始响应对象: {response}")
except Exception as e:
    print(f"AI调用发生错误: {e}")
    # 对于更详细的错误信息，可以检查 e 的类型，例如 google.api_core.exceptions.ResourceExhausted