from fastmcp import FastMCP

# Test MCP
mcp = FastMCP("Test MCP")


@mcp.tool
def greeting(name: str) -> str:
    # Chào người dùng với tên được cung cấp
    return f"Hello, {name}!"


if __name__ == "__main__":
    mcp.run()
