import openai
import os
import base64
from core.config import settings
def main(image_path):
    try:
        def encode_image(image_path):
            """将本地图片编码为Base64格式"""
            with open(image_path, "rb") as image_file:
                return base64.b64encode(image_file.read()).decode('utf-8')
        # 检查图片文件是否存在
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"图片文件不存在: {image_path}")

        # 将图片编码为Base64
        base64_image = encode_image(image_path)

        client = openai.OpenAI(
            api_key=settings.ai_api_key,
            base_url=settings.ai_api_url
        )

        response = client.chat.completions.create(
            model="Qwen/Qwen2.5-VL-32B-Instruct",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": "识别图片中的4位验证码，结果只用输出验证码不要包含别的部分"},
                        {"type": "image_url", "image_url": {
                            "url": f"data:image/jpeg;base64,{base64_image}"  # 使用Base64格式的图片
                        }}
                    ]
                }
            ],
            max_tokens=2000
        )
        return response.choices[0].message.content
    except Exception as e:print(e)

def text_ai(text):
    client=openai.OpenAI(
        api_key="sk-J4ndZW4RLXwrd32D40289e4dAdAc4306B4634cBbBcD5BdE7",
        base_url="https://llm-oneapi.bytebroad.com.cn/v1")
    res=client.chat.completions.create(messages=[{"role":"user","content":text}],model="Qwen/Qwen2.5-VL-72B-Instruct")
    return res.choices[0].message.content