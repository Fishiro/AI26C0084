# main.py
import os
import sys
import re
import csv
import json
import argparse
import asyncio
from typing import List, Dict, Any
from contextlib import AsyncExitStack
from dotenv import load_dotenv
from openai import AsyncOpenAI
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

# Tải biến môi trường từ file .env (khi test local)
load_dotenv()

SYSTEM_PROMPT = """Bạn là hệ thống AI chuyên phân tích và truy vấn bảng tính Excel (.xlsx) với độ chính xác tuyệt đối.

QUY TRÌNH XỬ LÝ:
1. Bạn đã được cung cấp bản xem trước cấu trúc tệp Excel (đã tự động tách ô gộp - merged cells, kèm tọa độ Dòng R1, R2... và Cột C1(A), C2(B)...).
2. Luôn sử dụng công cụ `run_python_analysis` để viết code Python/Pandas trích xuất số liệu và tính toán (tổng, hiệu, trung bình, tỷ lệ, tìm min/max, xếp hạng, ngày tháng).
   - Trong `run_python_analysis`, hàm `load_unmerged_df(sheet_name=None)` đã có sẵn để lấy DataFrame toàn bộ sheet đã tách ô gộp (lưu ý: dòng R1 ứng với index 0, cột C1(A) ứng với cột 0 trong DataFrame).
3. Nếu bảng dài hơn bản xem trước, hãy dùng `read_sheet_grid`, `search_in_workbook` hoặc viết code trong `run_python_analysis` để kiểm tra toàn bộ dòng.
4. Tuân thủ NGHIÊM NGẶT định dạng đầu ra mà câu hỏi yêu cầu (số nguyên không dấu phẩy, số chữ số thập phân, thứ tự liệt kê, Yes/No, ngôn ngữ...).
5. Đặt đáp án cuối cùng vào trong cặp thẻ <FINAL_ANSWER>...</FINAL_ANSWER>. Tuyệt đối không viết thêm bất kỳ lời giải thích hay đơn vị không được yêu cầu nào bên trong thẻ này."""


def convert_tool_format(tool) -> dict:
    """Chuyển đổi định dạng tool từ MCP sang chuẩn OpenAI Function Calling."""
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


def normalize_final_answer(query: str, raw_text: str) -> str:
    """
    Hàm kiểm soát và chuẩn hóa mẫu câu trả lời cuối cùng theo yêu cầu câu hỏi (Exact Match).
    """
    if not raw_text:
        return ""

    # 1. Trích xuất nội dung trong <FINAL_ANSWER>
    matches = re.findall(r"<FINAL_ANSWER>(.*?)</FINAL_ANSWER>",
                         raw_text, re.DOTALL | re.IGNORECASE)
    ans = matches[-1].strip() if matches else raw_text.strip().splitlines()[-1].strip()

    # Loại bỏ tiền tố dư thừa
    ans = re.sub(
        r"^(Answer|Final Answer|Đáp án|Kết quả|Output)\s*[:=-]\s*", "", ans, flags=re.IGNORECASE).strip()
    if (ans.startswith('"') and ans.endswith('"')) or (ans.startswith("'") and ans.endswith("'")):
        ans = ans[1:-1].strip()
    if ans.endswith(".") and not query.strip().endswith("."):
        ans = ans[:-1].strip()

    q_lower = query.lower()

    # 2. Chuẩn hóa câu hỏi Yes / No
    if "yes or no" in q_lower or "'yes' or 'no'" in q_lower:
        if re.search(r"\byes\b", ans, re.IGNORECASE):
            return "Yes"
        if re.search(r"\bno\b", ans, re.IGNORECASE):
            return "No"

    # 3. Xác định yêu cầu làm tròn chữ số thập phân
    decimals_map = [
        (["one decimal place", "1 decimal place",
         "một chữ số thập phân", "une décimale"], 1),
        (["two decimal places", "2 decimal places",
         "hai chữ số thập phân", "deux décimales"], 2),
        (["three decimal places", "3 decimal places",
         "ba chữ số thập phân", "trois décimales"], 3),
    ]
    target_decimals = None
    for phrases, dec in decimals_map:
        if any(p in q_lower for p in phrases):
            target_decimals = dec
            break

    # 4. Xác định yêu cầu trả về số nguyên
    is_integer_req = any(
        kw in q_lower
        for kw in [
            "only the integer", "as an integer", "as integers",
            "two integers", "integer value", "số nguyên", "only the year",
        ]
    )

    # 5. Nếu câu hỏi yêu cầu trả về nhiều giá trị phân tách bằng dấu phẩy
    # SỬA LỖI BUG: Chuyển kiểm tra từ ", " sang "," để bắt lỗi LLM quên khoảng trắng
    if "," in ans and ("separated by a comma" in q_lower or "answer as `" in q_lower):
        parts = [p.strip() for p in ans.split(",")]
        formatted_parts = []
        for idx, p in enumerate(parts):
            clean_p = p.replace(",", "").replace(
                "\$", "").replace("%", "").strip()
            try:
                val = float(clean_p)
                if target_decimals is not None and ("." in clean_p or idx == len(parts) - 1):
                    formatted_parts.append(f"{val:.{target_decimals}f}")
                elif is_integer_req or clean_p.lstrip("-").isdigit():
                    formatted_parts.append(str(int(round(val))))
                else:
                    formatted_parts.append(p)
            except ValueError:
                formatted_parts.append(p)
        return ", ".join(formatted_parts)

    # 6. Nếu câu trả lời là một số đơn lẻ
    clean_single = ans.replace(",", "").replace(
        "\$", "").replace("%", "").strip()
    try:
        val = float(clean_single)
        if target_decimals is not None:
            return f"{val:.{target_decimals}f}"
        if is_integer_req:
            return str(int(round(val)))
        if ("without commas" in q_lower or "không có dấu phẩy" in q_lower) and clean_single.lstrip("-").isdigit():
            return str(int(round(val)))
    except ValueError:
        pass

    return ans


class MCPClient:
    def __init__(self, api_key: str, model: str, base_url: str):
        self.session: ClientSession | None = None
        self.exit_stack = AsyncExitStack()
        self.openai = AsyncOpenAI(
            base_url=base_url,
            api_key=api_key,
            timeout=90.0,
            max_retries=2,
        )
        self.model = model
        self.available_tools = []
        self.overview_cache = {}

    async def connect_to_server(self):
        """Khởi chạy và kết nối tới helper/excel_mcp_server.py qua giao thức stdio."""
        server_script = os.path.join(os.path.dirname(
            os.path.abspath(__file__)), "helper", "excel_mcp_server.py")
        server_params = StdioServerParameters(
            command=sys.executable,
            args=[server_script],
            env=os.environ.copy(),
        )
        stdio_transport = await self.exit_stack.enter_async_context(stdio_client(server_params))
        self.stdio, self.write = stdio_transport
        self.session = await self.exit_stack.enter_async_context(ClientSession(self.stdio, self.write))
        await self.session.initialize()

        response = await self.session.list_tools()
        self.available_tools = [
            convert_tool_format(t) for t in response.tools if t.name != "get_workbook_overview"
        ]
        print(
            f"Đã kết nối MCP Server. Tools sẵn sàng: {[t.name for t in response.tools]}", file=sys.stderr)

    async def get_workbook_context(self, file_path: str) -> str:
        """Lấy và cache tổng quan file Excel để tái sử dụng."""
        if file_path not in self.overview_cache:
            assert self.session is not None  # Fix Pylance NoneType error
            res = await self.session.call_tool("get_workbook_overview", {"file_path": file_path})
            self.overview_cache[file_path] = "\n".join(
                # Fix ImageContent Pylance error
                item.text for item in res.content if getattr(item, "type", "") == "text"
            )
        return self.overview_cache[file_path]

    async def process_query(self, q_id, file_path: str, query: str) -> str:
        """Xử lý từng câu hỏi tự nhiên qua vòng lặp MCP Tool Calling."""
        overview = await self.get_workbook_context(file_path)

        user_message = (
            f"File Path: `{file_path}`\n"
            f"Question: {query}\n\n"
            f"{overview}\n\n"
            f"Hãy dùng `run_python_analysis` để tính ra kết quả chính xác và trả lời trong thẻ <FINAL_ANSWER>...</FINAL_ANSWER>."
        )

        # Cấp định dạng type chính xác cho messages để tránh lỗi Pylance
        messages: List[Dict[str, Any]] = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_message},
        ]

        for step in range(6):
            response = await self.openai.chat.completions.create(
                model=self.model,
                messages=messages,
                tools=self.available_tools,
                tool_choice="auto",
                temperature=0.0,
            )

            message = response.choices[0].message

            # SỬA LỖI BUG: Chuyển đổi ChatCompletionMessage thành Dict chuẩn để giải quyết toàn bộ xung đột Type của Pylance và thư viện v1
            assistant_msg: Dict[str, Any] = {
                "role": "assistant",
                "content": message.content or ""
            }

            if message.tool_calls:
                assistant_msg["tool_calls"] = [
                    {
                        "id": tc.id,
                        "type": "function",
                        "function": {
                            "name": tc.function.name,
                            "arguments": tc.function.arguments
                        }
                    } for tc in message.tool_calls
                ]

            messages.append(assistant_msg)

            if not message.tool_calls:
                raw_ans = message.content or ""
                return normalize_final_answer(query, raw_ans)

            for tool_call in message.tool_calls:
                tool_name = tool_call.function.name
                try:
                    tool_args = json.loads(tool_call.function.arguments)
                except Exception:
                    tool_args = {}

                print(
                    f"  [ID {q_id} | Step {step+1}] Gọi tool: {tool_name}", file=sys.stderr)
                try:
                    assert self.session is not None  # Fix Pylance NoneType error
                    result = await self.session.call_tool(tool_name, tool_args)

                    # Fix Pylance ImageContent type error
                    tool_output = "\n".join(
                        item.text for item in result.content if getattr(item, "type", "") == "text"
                    )
                except Exception as e:
                    tool_output = f"Lỗi khi gọi tool {tool_name}: {str(e)}"

                messages.append({
                    "role": "tool",
                    "tool_call_id": tool_call.id,
                    "name": tool_name,
                    "content": str(tool_output),
                })

        messages.append({
            "role": "user",
            "content": "Dựa trên toàn bộ kết quả đã chạy ở trên, hãy đưa ra đáp án cuối cùng ngay lập tức trong thẻ <FINAL_ANSWER>...</FINAL_ANSWER>.",
        })
        final_resp = await self.openai.chat.completions.create(
            model=self.model,
            messages=messages,
            temperature=0.0,
        )
        return normalize_final_answer(query, final_resp.choices[0].message.content or "")

    async def cleanup(self):
        await self.exit_stack.aclose()


async def run_pipeline(args):
    out_dir = os.path.dirname(os.path.abspath(args.output))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    with open(args.queries, "r", encoding="utf-8") as f:
        queries_list = json.load(f)

    client = MCPClient(
        api_key=args.openrouter_key,
        model=args.openrouter_model,
        base_url=args.openrouter_base_url,
    )

    results = []
    try:
        await client.connect_to_server()

        for item in queries_list:
            q_id = item.get("id")
            file_name = item.get("file_name") or item.get("file name")
            query_text = item.get("query", "")
            excel_path = os.path.join(args.xlsx_dir, file_name)

            print(
                f"\n>>> Đang xử lý ID: {q_id} | File: {file_name}", file=sys.stderr)
            try:
                answer = await client.process_query(q_id, excel_path, query_text)
            except Exception as e:
                print(f"[LỖI ID {q_id}]: {str(e)}", file=sys.stderr)
                answer = ""

            print(f"<<< Đáp án ID {q_id}: {answer}", file=sys.stderr)
            results.append({"id": q_id, "answer": answer})

            # Ghi kết quả ra file CSV theo chuẩn UTF-8
            with open(args.output, "w", newline="", encoding="utf-8") as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=["id", "answer"])
                writer.writeheader()
                writer.writerows(results)

    finally:
        await client.cleanup()

    print(
        f"\nHoàn tất! Đã ghi {len(results)} câu trả lời vào '{args.output}'.", file=sys.stderr)


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="AI26C0084 - Giải pháp Hỏi đáp trên Bảng tính Excel với MCP")
    parser.add_argument("--xlsx-dir", type=str,
                        default="excel_input", help="Thư mục chứa các tệp .xlsx")
    parser.add_argument("--queries", type=str,
                        default="queries.json", help="Đường dẫn tệp queries.json")
    parser.add_argument("--output", type=str, default=os.path.join(
        "submission_output", "submission.csv"), help="Đường dẫn tệp kết quả")
    parser.add_argument("--openrouter-key", type=str,
                        default=os.getenv("OPENROUTER_API_KEY", ""), help="Khóa API")
    parser.add_argument("--openrouter-model", type=str, default=os.getenv(
        "OPENROUTER_MODEL_CONFIG", "qwen/qwen3.7-flash"), help="Model ID")
    parser.add_argument("--openrouter-base-url", type=str,
                        default="https://openrouter.ai/api/v1", help="Base URL")
    return parser.parse_args()


if __name__ == "__main__":
    cli_args = parse_arguments()
    if not cli_args.openrouter_key:
        print("Lỗi: Thiếu khóa API (--openrouter-key hoặc biến môi trường OPENROUTER_API_KEY).", file=sys.stderr)
        sys.exit(1)

    asyncio.run(run_pipeline(cli_args))
