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

# Tải biến môi trường
load_dotenv()

# Cấu hình giá tiền tham khảo cho qwen/qwen3.7-flash (USD/1 triệu token)
# Lưu ý: Cập nhật lại giá này theo đúng bảng giá của OpenRouter tại thời điểm thi.
PRICE_1M_PROMPT = 0.10
PRICE_1M_COMPLETION = 0.20

SYSTEM_PROMPT = """Bạn là hệ thống AI chuyên phân tích và truy vấn bảng tính Excel (.xlsx) với độ chính xác tuyệt đối.

QUY TRÌNH XỬ LÝ BẮT BUỘC (CHAIN OF THOUGHT):
Mỗi lần phản hồi, bạn PHẢI tuân thủ 2 bước sau:

1. BƯỚC SUY NGHĨ: Đặt toàn bộ phân tích, kế hoạch và lý luận vào thẻ <THOUGHT>...</THOUGHT>.
   - Tiền chuẩn hóa: Xem xét kỹ cấu trúc file, data types (chuỗi, số, ngày tháng) trước khi gọi tool.
   - Nếu gọi tool bị lỗi, hãy phân tích lỗi trong thẻ này và đề xuất code Python mới. KHÔNG lặp lại code cũ.

2. BƯỚC HÀNH ĐỘNG / TRẢ LỜI:
   - Nếu cần gọi tool: Dùng `run_python_analysis` (hàm `load_unmerged_df()` đã có sẵn). BẮT BUỘC dùng `print()` để xuất kết quả.
   - Nếu ĐÃ CÓ kết quả cuối cùng: Đặt ĐÁP ÁN DUY NHẤT vào thẻ <FINAL_ANSWER>...</FINAL_ANSWER>. Tuyệt đối không giải thích thêm trong thẻ này.

LƯU Ý QUAN TRỌNG:
- Trả về đúng định dạng yêu cầu (số nguyên, số thập phân, Yes/No).
- Mọi suy nghĩ phải nằm trong <THOUGHT>, mọi đáp án cuối nằm trong <FINAL_ANSWER>.
"""


def convert_tool_format(tool) -> dict:
    return {
        "type": "function",
        "function": {
            "name": tool.name,
            "description": tool.description,
            "parameters": {
                "type": "object",
                "properties": tool.input_schema.get("properties", {}),
                "required": tool.input_schema.get("required", []),
            },
        },
    }


def normalize_final_answer(query: str, raw_text: str) -> str:
    """Tiền chuẩn hóa và hậu xử lý kết quả."""
    if not raw_text:
        return ""

    # Trích xuất nội dung trong <FINAL_ANSWER>
    matches = re.findall(r"<FINAL_ANSWER>(.*?)</FINAL_ANSWER>",
                         raw_text, re.DOTALL | re.IGNORECASE)
    if matches:
        ans = matches[-1].strip()
    else:
        # Fallback nếu LLM quên thẻ (loại bỏ phần thought nếu có)
        clean_text = re.sub(r"<THOUGHT>.*?</THOUGHT>", "",
                            raw_text, flags=re.DOTALL).strip()
        ans = clean_text.splitlines()[-1].strip() if clean_text else ""

    ans = re.sub(
        r"^(Answer|Final Answer|Đáp án|Kết quả|Output)\s*[:=-]\s*", "", ans, flags=re.IGNORECASE).strip()
    if (ans.startswith('"') and ans.endswith('"')) or (ans.startswith("'") and ans.endswith("'")):
        ans = ans[1:-1].strip()
    if ans.endswith(".") and not query.strip().endswith("."):
        ans = ans[:-1].strip()

    q_lower = query.lower()

    if "yes or no" in q_lower or "'yes' or 'no'" in q_lower:
        if re.search(r"\byes\b", ans, re.IGNORECASE):
            return "Yes"
        if re.search(r"\bno\b", ans, re.IGNORECASE):
            return "No"

    decimals_map = [
        (["one decimal place", "1 decimal place"], 1),
        (["two decimal places", "2 decimal places"], 2),
        (["three decimal places", "3 decimal places"], 3),
    ]
    target_decimals = None
    for phrases, dec in decimals_map:
        if any(p in q_lower for p in phrases):
            target_decimals = dec
            break

    is_integer_req = any(kw in q_lower for kw in [
                         "only the integer", "as an integer", "số nguyên"])

    if "," in ans and ("separated by a comma" in q_lower or "answer as `" in q_lower):
        parts = [p.strip() for p in ans.split(",")]
        formatted_parts = []
        for idx, p in enumerate(parts):
            clean_p = p.replace(",", "").replace(
                "$", "").replace("%", "").strip()
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

    clean_single = ans.replace(",", "").replace(
        "$", "").replace("%", "").strip()
    try:
        val = float(clean_single)
        if target_decimals is not None:
            return f"{val:.{target_decimals}f}"
        if is_integer_req:
            return str(int(round(val)))
    except ValueError:
        pass

    return ans


class MCPClient:
    def __init__(self, api_key: str, model: str, base_url: str):
        self.session: ClientSession | None = None
        self.exit_stack = AsyncExitStack()
        self.openai = AsyncOpenAI(
            base_url=base_url, api_key=api_key, timeout=90.0, max_retries=2)
        self.model = model
        self.available_tools = []
        self.overview_cache = {}

    async def connect_to_server(self):
        server_script = os.path.join(os.path.dirname(
            os.path.abspath(__file__)), "helper", "excel_mcp_server.py")
        server_params = StdioServerParameters(command=sys.executable, args=[
                                              server_script], env=os.environ.copy())
        stdio_transport = await self.exit_stack.enter_async_context(stdio_client(server_params))
        self.stdio, self.write = stdio_transport
        self.session = await self.exit_stack.enter_async_context(ClientSession(self.stdio, self.write))
        await self.session.initialize()

        response = await self.session.list_tools()
        self.available_tools = [convert_tool_format(
            t) for t in response.tools if t.name != "get_workbook_overview"]
        print(
            f"[*] Đã kết nối MCP. Tools: {[t.name for t in response.tools]}", file=sys.stderr)

    async def get_workbook_context(self, file_path: str) -> str:
        if file_path not in self.overview_cache:
            assert self.session is not None
            res = await self.session.call_tool("get_workbook_overview", {"file_path": file_path})
            self.overview_cache[file_path] = "\n".join(
                item.text for item in res.content if getattr(item, "type", "") == "text")
        return self.overview_cache[file_path]

    async def process_query(self, q_id, file_path: str, query: str) -> Dict[str, Any]:
        overview = await self.get_workbook_context(file_path)
        user_message = f"File Path: `{file_path}`\nQuestion: {query}\n\n{overview}"

        messages: List[Dict[str, Any]] = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_message},
        ]

        total_p_tokens = 0
        total_c_tokens = 0
        total_cost = 0.0
        final_answer = ""

        for step in range(8):  # Tăng lên 8 step đề phòng dữ liệu phức tạp
            response = await self.openai.chat.completions.create(
                model=self.model,
                messages=messages,  # type: ignore
                tools=self.available_tools,  # type: ignore
                tool_choice="auto",
                temperature=0.0,  # Ép Temperature = 0 để tránh bịa đặt
            )

            message = response.choices[0].message

            # Xử lý Token & Chi phí
            if response.usage:
                p_tokens = response.usage.prompt_tokens
                c_tokens = response.usage.completion_tokens
                total_p_tokens += p_tokens
                total_c_tokens += c_tokens
                step_cost = (p_tokens / 1_000_000 * PRICE_1M_PROMPT) + \
                    (c_tokens / 1_000_000 * PRICE_1M_COMPLETION)
                total_cost += step_cost
                print(
                    f"  [Step {step+1}] Tokens: Prompt={p_tokens}, Comp={c_tokens} | Chi phí: ${step_cost:.6f}", file=sys.stderr)

            # In luồng suy nghĩ (Trực quan hóa)
            if message.content:
                thoughts = re.findall(
                    r"<THOUGHT>(.*?)</THOUGHT>", message.content, re.DOTALL | re.IGNORECASE)
                for t in thoughts:
                    print(f"  [💡 THOUGHT]: {t.strip()}", file=sys.stderr)

            assistant_msg: Dict[str, Any] = {
                "role": "assistant", "content": message.content or ""}

            if message.tool_calls:
                assistant_msg["tool_calls"] = []
                for tc in message.tool_calls:
                    if tc.type == "function":
                        assistant_msg["tool_calls"].append({
                            "id": tc.id, "type": "function", "function": {
                                "name": tc.function.name,
                                "arguments": tc.function.arguments}
                        })
                    else:
                        assistant_msg["tool_calls"].append({
                            "id": tc.id, "type": "custom", "custom": {
                                "name": tc.custom.name,
                                "input": tc.custom.input}
                        })
                messages.append(assistant_msg)

                for tool_call in message.tool_calls:
                    if tool_call.type == "function":
                        tool_name = tool_call.function.name
                        raw_tool_args = tool_call.function.arguments
                    else:
                        tool_name = tool_call.custom.name
                        raw_tool_args = tool_call.custom.input
                    try:
                        tool_args = json.loads(raw_tool_args)
                    except Exception:
                        tool_args = {}

                    print(
                        f"  [🛠️ CALL]: {tool_name} {tool_args}", file=sys.stderr)
                    try:
                        assert self.session is not None
                        result = await self.session.call_tool(tool_name, tool_args)
                        tool_output = "\n".join(
                            item.text for item in result.content if getattr(item, "type", "") == "text")
                    except Exception as e:
                        tool_output = f"Lỗi khi gọi tool {tool_name}: {str(e)}"

                    # Chống loop: Cảnh báo LLM nếu tool output rỗng hoặc lỗi
                    if not tool_output.strip() or "Lỗi" in tool_output:
                        tool_output += "\nCẢNH BÁO: Phân tích nguyên nhân lỗi trong <THOUGHT> và viết phương pháp mới. KHÔNG lặp lại code cũ."

                    print(
                        f"  [📩 OUTPUT]: {tool_output[:200]}...", file=sys.stderr)
                    messages.append({"role": "tool", "tool_call_id": tool_call.id,
                                    "name": tool_name, "content": str(tool_output)})
            else:
                messages.append(assistant_msg)
                final_answer = normalize_final_answer(
                    query, message.content or "")
                if final_answer:
                    break
                else:
                    # Nếu LLM dừng nhưng không đưa ra <FINAL_ANSWER>
                    messages.append(
                        {"role": "user", "content": "Bạn chưa cung cấp <FINAL_ANSWER>. Hãy phân tích lại kết quả và xuất đáp án cuối."})

        # Xử lý nếu hết vòng lặp mà chưa có thẻ FINAL_ANSWER
        if not final_answer:
            messages.append(
                {"role": "user", "content": "Dựa vào dữ liệu trên, hãy đưa ra đáp án NGAY LẬP TỨC trong <FINAL_ANSWER>...</FINAL_ANSWER>."})

            final_resp = await self.openai.chat.completions.create(model=self.model, messages=messages, temperature=0.0)

            if final_resp.usage:
                total_p_tokens += final_resp.usage.prompt_tokens
                total_c_tokens += final_resp.usage.completion_tokens
                total_cost += (final_resp.usage.prompt_tokens / 1_000_000 * PRICE_1M_PROMPT) + (
                    final_resp.usage.completion_tokens / 1_000_000 * PRICE_1M_COMPLETION)

            final_answer = normalize_final_answer(
                query, final_resp.choices[0].message.content or "")

        return {"answer": final_answer, "cost": total_cost, "tokens": total_p_tokens + total_c_tokens}

    async def cleanup(self):
        await self.exit_stack.aclose()


async def run_pipeline(args):
    out_dir = os.path.dirname(os.path.abspath(args.output))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    with open(args.queries, "r", encoding="utf-8") as f:
        queries_list = json.load(f)

    client = MCPClient(api_key=args.openrouter_key,
                       model=args.openrouter_model, base_url=args.openrouter_base_url)
    results = []

    total_campaign_cost = 0.0
    total_campaign_tokens = 0

    try:
        await client.connect_to_server()

        for item in queries_list:
            q_id = item.get("id")
            file_name = item.get("file_name") or item.get("file name")
            query_text = item.get("query", "")
            excel_path = os.path.join(args.xlsx_dir, file_name)

            print(
                f"\n{'='*50}\n>>> Đang xử lý ID: {q_id} | File: {file_name}", file=sys.stderr)
            try:
                res_data = await client.process_query(q_id, excel_path, query_text)
                answer = res_data["answer"]
                cost = res_data["cost"]
                tokens = res_data["tokens"]

                total_campaign_cost += cost
                total_campaign_tokens += tokens

                print(
                    f"\n[THỐNG KÊ ID {q_id}] Chi phí: ${cost:.6f} | Tổng Token: {tokens}", file=sys.stderr)
            except Exception as e:
                print(f"[LỖI CRITICAL ID {q_id}]: {str(e)}", file=sys.stderr)
                answer = ""

            print(
                f"<<< ĐÁP ÁN CHÍNH THỨC ID {q_id}: {answer}", file=sys.stderr)
            results.append({"id": q_id, "answer": answer})

            # Ghi file liên tục (tránh crash giữa chừng mất data)
            with open(args.output, "w", newline="", encoding="utf-8") as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=["id", "answer"])
                writer.writeheader()
                writer.writerows(results)

    finally:
        await client.cleanup()

    print(f"\n{'='*50}", file=sys.stderr)
    print(f"HOÀN TẤT CÔNG VIỆC!", file=sys.stderr)
    print(f"- Đã xử lý: {len(results)} câu", file=sys.stderr)
    print(f"- Tổng Token sử dụng: {total_campaign_tokens}", file=sys.stderr)
    print(f"- TỔNG CHI PHÍ: ${total_campaign_cost:.6f}", file=sys.stderr)
    print(f"*(Lưu ý: Giới hạn BTC là 0.10 USD/câu hỏi)*", file=sys.stderr)
    print(f"{'='*50}\n", file=sys.stderr)


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
