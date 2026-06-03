"""
C宝视觉识别 - 调用豆包模型分析图片
用法: python vision.py <图片路径> [可选：你的问题]
"""
import sys
import os
import base64
import json
import urllib.request

# 解决 Windows 控制台 GBK 编码问题
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

# ========== 配置 ==========
API_KEY = "ark-74990dfa-9292-47a5-9ade-1ac81567a98c-052c0"
API_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
MODEL = "doubao-seed-2-0-lite-260428"

def encode_image(image_path):
    """把图片转成 base64"""
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

def get_mime_type(image_path):
    ext = os.path.splitext(image_path)[1].lower()
    mime_map = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
                ".gif": "image/gif", ".webp": "image/webp", ".bmp": "image/bmp"}
    return mime_map.get(ext, "image/png")

def analyze_image(image_path, question="请详细描述这张图片的内容，包括文字、图表、人物、场景等所有细节。"):
    """调用豆包视觉模型分析图片"""
    print(f"[C宝] 正在看图片: {image_path}")

    b64 = encode_image(image_path)
    mime = get_mime_type(image_path)

    payload = {
        "model": MODEL,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                {"type": "text", "text": question}
            ]
        }],
        "max_tokens": 2000
    }

    req = urllib.request.Request(API_URL, data=json.dumps(payload).encode(), method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("Authorization", f"Bearer {API_KEY}")

    try:
        resp = urllib.request.urlopen(req, timeout=60)
        result = json.loads(resp.read())
        answer = result["choices"][0]["message"]["content"]
        return answer
    except Exception as e:
        return f"[错误] {e}"

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python vision.py <图片路径> [你的问题]")
        print("示例: python vision.py C:/Users/33164/Desktop/图片/1.png")
        print("示例: python vision.py C:/Users/33164/Desktop/图片/1.png \"图片里有什么文字？\"")
        sys.exit(1)

    image_path = sys.argv[1]
    question = sys.argv[2] if len(sys.argv) > 2 else "请详细描述这张图片的内容，包括文字、图表、人物、场景等所有细节。"

    if not os.path.exists(image_path):
        print(f"[C宝] 找不到图片: {image_path}")
        sys.exit(1)

    result = analyze_image(image_path, question)
    print(f"\n[C宝] 分析结果:\n{result}")
