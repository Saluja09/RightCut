"""
RightCut — Agent orchestrator.
Runs the Gemini function-calling loop with manual tool execution
so every step can be streamed to the frontend tool-call timeline.
"""

from __future__ import annotations

import asyncio
import logging
import time
from typing import TYPE_CHECKING

from google import genai
from google.genai import types

from agent.prompts import SYSTEM_PROMPT
from agent.tool_schemas import RIGHTCUT_TOOL
from agent.tools import ToolExecutor
from config import (
    AGENT_TEMPERATURE,
    GEMINI_MODEL,
    MAX_BACKOFF_SECONDS,
    MAX_TOOL_ITERATIONS,
    RATE_LIMIT_DELAY_BASE,
)
from excel.engine import WorkbookEngine
from excel.serializer import serialize_workbook
from models import ToolStep

logger = logging.getLogger(__name__)


class AgentOrchestrator:
    """
    Drives the Gemini agentic loop with manual AFC disabled.
    Each tool call is streamed to the WebSocket as it happens.
    """

    def __init__(self, api_key: str) -> None:
        self.client = genai.Client(api_key=api_key)

    def _make_config(self) -> types.GenerateContentConfig:
        return types.GenerateContentConfig(
            system_instruction=SYSTEM_PROMPT,
            tools=[RIGHTCUT_TOOL],
            # CRITICAL: disable automatic function calling so we can stream each step
            automatic_function_calling=types.AutomaticFunctionCallingConfig(
                disable=True
            ),
            tool_config=types.ToolConfig(
                function_calling_config=types.FunctionCallingConfig(
                    mode=types.FunctionCallingConfigMode.AUTO
                )
            ),
            temperature=AGENT_TEMPERATURE,
            max_output_tokens=8192,
        )

    async def run(
        self,
        user_message: str,
        conversation_history: list[types.Content],
        executor: ToolExecutor,
        workbook: WorkbookEngine,
        websocket,
    ) -> tuple[str, list[ToolStep]]:
        """
        Execute the agentic loop for a single user message.
        Returns (final_text, timeline).
        Mutates conversation_history in place.
        """
        # Build working contents: history + new user turn
        contents: list[types.Content] = list(conversation_history) + [
            types.Content(
                role="user",
                parts=[types.Part.from_text(text=user_message)],
            )
        ]

        timeline: list[ToolStep] = []
        iteration = 0

        while iteration < MAX_TOOL_ITERATIONS:
            iteration += 1

            # Signal to frontend that we're calling Gemini
            await _safe_send(websocket, {"type": "thinking", "iteration": iteration})

            # Call Gemini (with retry on rate limits)
            response = await self._call_with_retry(contents)
            candidate = response.candidates[0]

            # Detect terminal states
            finish = candidate.finish_reason
            has_function_calls = any(
                hasattr(p, "function_call") and p.function_call
                for p in candidate.content.parts
            )
            has_text = any(
                hasattr(p, "text") and p.text
                for p in candidate.content.parts
            )

            if not has_function_calls:
                # Agent is done — return text response
                final_text = "".join(
                    p.text for p in candidate.content.parts if hasattr(p, "text") and p.text
                ) or "Done."

                # Final workbook state push
                wb_state = serialize_workbook(workbook.wb, workbook._charts)
                await _safe_send(websocket, {"type": "workbook_update", "state": wb_state})

                # Push full sheet tab list
                for sheet_name in workbook.get_all_sheet_names():
                    await _safe_send(websocket, {
                        "type": "new_tab",
                        "tab": {"id": sheet_name, "name": sheet_name, "type": "sheet"},
                    })

                await _safe_send(websocket, {
                    "type": "agent_response",
                    "text": final_text,
                    "timeline": [s.model_dump() for s in timeline],
                })

                # Update conversation history
                conversation_history.append(
                    types.Content(role="user", parts=[types.Part.from_text(text=user_message)])
                )
                conversation_history.append(candidate.content)

                return final_text, timeline

            # ── Process function calls ────────────────────────────────────────
            function_call_parts = [
                p for p in candidate.content.parts
                if hasattr(p, "function_call") and p.function_call
            ]

            # Add model's turn to working contents
            contents.append(candidate.content)

            function_response_parts: list[types.Part] = []

            for fc_part in function_call_parts:
                fc = fc_part.function_call
                tool_name = fc.name
                tool_args = dict(fc.args) if fc.args else {}

                start_ts = time.perf_counter()
                result = await executor.execute(tool_name, tool_args)
                duration_ms = round((time.perf_counter() - start_ts) * 1000)

                step = ToolStep(
                    tool=tool_name,
                    args=tool_args,
                    result_summary=result.summary,
                    duration_ms=duration_ms,
                    success=result.success,
                    error=result.error,
                )
                timeline.append(step)

                # Stream tool step to frontend immediately
                await _safe_send(websocket, {"type": "tool_call", "step": step.model_dump()})

                # Push incremental workbook updates for mutating tools
                if tool_name in {
                    "create_sheet", "insert_data", "add_formula",
                    "edit_cell", "apply_formatting", "sort_range",
                }:
                    wb_state = serialize_workbook(workbook.wb, workbook._charts)
                    await _safe_send(websocket, {"type": "workbook_update", "state": wb_state})
                    # Announce new tabs
                    for sheet_name in workbook.get_all_sheet_names():
                        await _safe_send(websocket, {
                            "type": "new_tab",
                            "tab": {"id": sheet_name, "name": sheet_name, "type": "sheet"},
                        })

                # Build function response part
                function_response_parts.append(
                    types.Part.from_function_response(
                        name=tool_name,
                        response=result.data if result.data else {"result": result.summary},
                    )
                )

            # Feed all responses back in one Content
            contents.append(
                types.Content(role="user", parts=function_response_parts)
            )

        # ── Hit iteration cap ─────────────────────────────────────────────────
        logger.warning(f"Agent hit MAX_TOOL_ITERATIONS={MAX_TOOL_ITERATIONS}")
        await _safe_send(websocket, {
            "type": "error",
            "message": (
                f"Agent reached the maximum of {MAX_TOOL_ITERATIONS} tool calls. "
                "Partial results are available in the workbook. "
                "Please send a follow-up message to continue."
            ),
        })
        return f"Reached maximum iterations ({MAX_TOOL_ITERATIONS}).", timeline

    async def _call_with_retry(
        self,
        contents: list[types.Content],
        max_retries: int = 8,
    ) -> types.GenerateContentResponse:
        """Call Gemini with exponential backoff on rate limit errors."""
        delay = RATE_LIMIT_DELAY_BASE

        for attempt in range(max_retries):
            try:
                return await self.client.aio.models.generate_content(
                    model=GEMINI_MODEL,
                    contents=contents,
                    config=self._make_config(),
                )
            except Exception as exc:
                err = str(exc).lower()
                is_rate_limit = any(k in err for k in ("429", "503", "quota", "rate limit", "resource_exhausted", "unavailable", "overloaded"))

                if is_rate_limit and attempt < max_retries - 1:
                    # Exponential backoff with jitter
                    jitter = (time.time() % 1)  # 0–1 second of jitter
                    sleep_for = min(delay * (2 ** attempt) + jitter, MAX_BACKOFF_SECONDS)
                    logger.warning(
                        f"Rate limit hit (attempt {attempt + 1}/{max_retries}), "
                        f"sleeping {sleep_for:.1f}s"
                    )
                    await asyncio.sleep(sleep_for)
                else:
                    raise

        raise RuntimeError("Exceeded retry attempts")  # unreachable


async def _safe_send(websocket, data: dict) -> None:
    """Send a JSON message to the websocket, swallowing disconnect errors."""
    try:
        await websocket.send_json(data)
    except Exception as e:
        logger.debug(f"WebSocket send failed (client likely disconnected): {e}")
