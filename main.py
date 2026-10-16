from dotenv import load_dotenv
import os
import requests
import json
# from helper import test
import asyncio
from contextlib import AsyncExitStack
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from openai import OpenAI

# Tải env
load_dotenv()

# Xuất env ra biến
OPENROUTER_API_KEY = os.getenv('OPENROUTER_API_KEY')
OPENROUTER_MODEL_CONFIG = os.getenv('OPENROUTER_MODEL_CONFIG')

# Cấu hình khởi chạy FastMCP
SERVER_CONFIG = {
    "command": "python",
    "args": ["my_server.py"],
    "env": None,
}


def convert_tool_format(tool):
    # Hàm chuyển đổi định dạng tool từ MCP sang OpenAI-compatible
    return {
        "type": "function",
        "function": {
            "name": tool.name,
            "description": tool.description,
            "parameters": {
                "type": "object",
                "properties": tool.inputSchema.get("properties", {}),
                "required": tool.inputSchema.get("required", []),
            },
        },
    }


class MCPClient:
    def __init__(self):
        self.session = None
        self.exit_stack = AsyncExitStack()
        self.openai = OpenAI(
            base_url="https://openrouter.ai/api/v1",
            api_key=OPENROUTER_API_KEY,
        )
        self.messages = []

    async def connect_to_server(self, server_config):
        """Kết nối tới MCP server qua stdio."""
        server_params = StdioServerParameters(**server_config)
        stdio_transport = await self.exit_stack.enter_async_context(
            stdio_client(server_params)
        )
        self.stdio, self.write = stdio_transport
        self.session = await self.exit_stack.enter_async_context(
            ClientSession(self.stdio, self.write)
        )
        await self.session.initialize()

        # Lấy danh sách các tool từ server
        response = await self.session.list_tools()
        print(
            f"\nĐã kết nối tới server với các tool: {[t.name for t in response.tools]}")

    async def process_query(self, query: str) -> str:
        """Xử lý câu hỏi của người dùng qua OpenRouter với tool calling."""
        self.messages.append({"role": "user", "content": query})

        # Lấy danh sách tool có sẵn và chuyển đổi định dạng
        response = await self.session.list_tools()
        available_tools = [convert_tool_format(t) for t in response.tools]

        # Gọi OpenRouter API
        response = self.openai.chat.completions.create(
            model=MODEL,
            messages=self.messages,
            tools=available_tools,
            temperature=0,  # Đặt temperature = 0 để đầu ra ổn định
        )

        message = response.choices[0].message

        # Xử lý nếu model muốn gọi tool
        if message.tool_calls:
            tool_call = message.tool_calls[0]
            tool_name = tool_call.function.name
            tool_args = json.loads(tool_call.function.arguments)

            # Thực thi tool thông qua MCP session
            result = await self.session.call_tool(tool_name, tool_args)
            print(f"Đã gọi tool: {tool_name} với args: {tool_args}")
            print(f"Kết quả từ tool: {result.content}")

            # Gửi kết quả tool trở lại model để có câu trả lời cuối cùng
            # Thêm phản hồi của assistant (chứa tool_calls)
            self.messages.append(message)
            self.messages.append({
                "role": "tool",
                "tool_call_id": tool_call.id,
                "content": str(result.content),
            })
            final_response = self.openai.chat.completions.create(
                model=MODEL,
                messages=self.messages,
                temperature=0,
            )
            return final_response.choices[0].message.content
        else:
            # Model trả lời trực tiếp mà không cần gọi tool
            return message.content

    async def cleanup(self):
        """Dọn dẹp tài nguyên."""
        await self.exit_stack.aclose()


if __name__ == "__main__":
    pass
