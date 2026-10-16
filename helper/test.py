from dotenv import load_dotenv
import os
import requests
import json

# Tải env
load_dotenv()

# Xuất env ra biến
OPENROUTER_API_KEY = os.getenv('OPENROUTER_API_KEY')
OPENROUTER_MODEL_CONFIG = os.getenv('OPENROUTER_MODEL_CONFIG')


# Hàm test API
def TestAPI():
    response = requests.post(
        url="https://openrouter.ai/api/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {OPENROUTER_API_KEY}",
            "Content-Type": "application/json",
        },
        data=json.dumps({
            "model": OPENROUTER_MODEL_CONFIG,
            "messages": [
                {
                    "role": "system",
                    "content": "Only answer in 10 token, this is like a ping request"
                },
                {
                    "role": "user",
                    "content": "Can you say hi?"
                }
            ],
            "reasoning": {"enabled": False}
        })
    )

    # Trích xuất câu trả lời
    raw_response = response.json()
    final_respone = raw_response['choices'][0]['message']['content']
    total_tokens = raw_response['usage']['total_tokens']

    return f"response: {final_respone}, total_tokens: {total_tokens}"
