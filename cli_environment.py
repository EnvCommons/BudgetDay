from __future__ import annotations

from typing import Any

from openreward.environments import Environment, JSONObject, TextBlock, ToolOutput, tool
from pydantic import BaseModel, Field


class BashInput(BaseModel):
    command: str = Field(..., description="The bash command to execute")


class ReadInput(BaseModel):
    file_path: str = Field(..., description="Path to the file to read")


class WriteInput(BaseModel):
    file_path: str = Field(..., description="Path to the file to write")
    content: str = Field(..., description="Content to write to the file")


class GrepInput(BaseModel):
    pattern: str = Field(..., description="Pattern to search for")
    path: str = Field(default=".", description="Path to search in")


class GlobInput(BaseModel):
    pattern: str = Field(..., description="Glob pattern to match files")


def sanitize_utf8(text: str) -> str:
    """
    Sanitize a string to ensure it only contains valid UTF-8 characters.
    Replaces invalid bytes and surrogate pairs with replacement characters.
    """
    if not text:
        return text
    # Encode to UTF-8 with error handling, then decode back
    # This replaces invalid sequences (including surrogates) with U+FFFD (�)
    return text.encode('utf-8', errors='replace').decode('utf-8', errors='replace')


class CLIEnvironment(Environment):
    """Base class for environments that provide CLI-like tools (bash, read, write, etc.)"""

    def __init__(self, task_spec: JSONObject, secrets: dict[str, str] = {}) -> None:
        super().__init__(task_spec)

    @tool
    async def bash(self, params: BashInput) -> ToolOutput:
        """Execute a bash command in the sandbox"""
        try:
            output, exit_code = await self.sandbox.run(params.command)

            # Sanitize output to prevent JSON serialization errors
            output = sanitize_utf8(output)

            display_text = f"Command: {params.command}\n"
            display_text += f"Exit code: {exit_code}\n\n"
            display_text += output if output else "(no output)"

            return ToolOutput(
                blocks=[TextBlock(text=display_text)],
                metadata={"command": params.command, "exit_code": exit_code, "output": output},
                reward=0.0,
                finished=False,
            )
        except Exception as e:
            return ToolOutput(
                blocks=[TextBlock(text=f"Error executing command: {str(e)}")],
                metadata={"error": str(e)},
                reward=0.0,
                finished=False,
            )

    @tool
    async def read(self, params: ReadInput) -> ToolOutput:
        """Read a file from the sandbox"""
        try:
            content_bytes = await self.sandbox.download(params.file_path)
            content = content_bytes.decode("utf-8", errors='replace')

            # Sanitize content to prevent JSON serialization errors
            content = sanitize_utf8(content)

            # Truncate if too long for display
            max_display = 10000
            if len(content) > max_display:
                display_content = content[:max_display] + f"\n\n... (truncated, {len(content)} total chars)"
            else:
                display_content = content

            return ToolOutput(
                blocks=[TextBlock(text=f"File: {params.file_path}\n\n{display_content}")],
                metadata={"file_path": params.file_path, "size": len(content)},
                reward=0.0,
                finished=False,
            )
        except Exception as e:
            return ToolOutput(
                blocks=[TextBlock(text=f"Error reading file {params.file_path}: {str(e)}")],
                metadata={"error": str(e)},
                reward=0.0,
                finished=False,
            )

    @tool
    async def write(self, params: WriteInput) -> ToolOutput:
        """Write content to a file in the sandbox"""
        try:
            content_bytes = params.content.encode("utf-8")
            await self.sandbox.upload_bytes(content_bytes, params.file_path)

            return ToolOutput(
                blocks=[TextBlock(text=f"Successfully wrote {len(content_bytes)} bytes to {params.file_path}")],
                metadata={"file_path": params.file_path, "size": len(content_bytes)},
                reward=0.0,
                finished=False,
            )
        except Exception as e:
            return ToolOutput(
                blocks=[TextBlock(text=f"Error writing to {params.file_path}: {str(e)}")],
                metadata={"error": str(e)},
                reward=0.0,
                finished=False,
            )

    @tool
    async def grep(self, params: GrepInput) -> ToolOutput:
        """Search for a pattern in files"""
        try:
            # Use grep via bash
            grep_cmd = f"grep -r '{params.pattern}' {params.path}"
            output, exit_code = await self.sandbox.run(grep_cmd)

            # Sanitize output to prevent JSON serialization errors
            output = sanitize_utf8(output)

            if exit_code == 0:
                display_text = f"Search results for '{params.pattern}':\n\n{output}"
            elif exit_code == 1:
                display_text = f"No matches found for '{params.pattern}'"
            else:
                display_text = f"Error searching (exit {exit_code}): {output}"

            return ToolOutput(
                blocks=[TextBlock(text=display_text)],
                metadata={"pattern": params.pattern, "path": params.path, "exit_code": exit_code},
                reward=0.0,
                finished=False,
            )
        except Exception as e:
            return ToolOutput(
                blocks=[TextBlock(text=f"Error searching: {str(e)}")],
                metadata={"error": str(e)},
                reward=0.0,
                finished=False,
            )

    @tool
    async def glob(self, params: GlobInput) -> ToolOutput:
        """Find files matching a glob pattern"""
        try:
            # Use find command
            find_cmd = f"find /orwd_data -name '{params.pattern}' -type f"
            output, exit_code = await self.sandbox.run(find_cmd)

            # Sanitize output to prevent JSON serialization errors
            output = sanitize_utf8(output)

            if exit_code == 0 and output:
                files = output.strip().split("\n")
                display_text = f"Found {len(files)} file(s) matching '{params.pattern}':\n\n"
                display_text += "\n".join(files)
            else:
                display_text = f"No files found matching '{params.pattern}'"
                files = []

            return ToolOutput(
                blocks=[TextBlock(text=display_text)],
                metadata={"pattern": params.pattern, "files": files},
                reward=0.0,
                finished=False,
            )
        except Exception as e:
            return ToolOutput(
                blocks=[TextBlock(text=f"Error finding files: {str(e)}")],
                metadata={"error": str(e)},
                reward=0.0,
                finished=False,
            )
