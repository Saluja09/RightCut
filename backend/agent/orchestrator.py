"""
RightCut — Agent orchestrator.
Runs the Gemini function-calling loop with manual tool execution
so every step can be streamed to the frontend tool-call timeline.
"""

from __future__ import annotations

import asyncio
import logging
import time

from google import genai
from google.genai import types

from agent.compaction import estimate_tokens, run_compaction_pipeline
from agent.prompts import SYSTEM_PROMPT
from agent.tool_schemas import RIGHTCUT_TOOL
from agent.tools import ToolExecutor
from config import (
    AGENT_TEMPERATURE,
    COMPACT_SUMMARY_KEEP_LAST,
    COMPACT_TOKEN_BUDGET,
    COMPACT_TOOL_KEEP_LAST,
    COMPACT_WINDOW_KEEP_LAST,
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
    Applies a three-stage compaction pipeline after each turn to control costs.
    """

    def __init__(self, api_key: str) -> None:
        self.client = genai.Client(api_key=api_key)

    def _make_config(self, system_prompt: str | None = None) -> types.GenerateContentConfig:
        return types.GenerateContentConfig(
            system_instruction=system_prompt or SYSTEM_PROMPT,
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
        system_prompt: str | None = None,
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

        def _flush_history_from_contents() -> None:
            """
            Copy the completed working contents back into conversation_history.
            Called on both clean exit and error exit so the model always has context.
            Skips the initial user message (index 0 in contents = len(history) at entry)
            to avoid double-appending.
            """
            existing_len = len(conversation_history)
            # contents[0..existing_len-1] are the pre-existing history items we copied in
            # contents[existing_len] is the new user message (not yet in history)
            # everything from existing_len onward is new this turn
            new_items = contents[existing_len:]
            if new_items:
                conversation_history.extend(new_items)

        while iteration < MAX_TOOL_ITERATIONS:
            iteration += 1

            # Signal to frontend that we're calling Gemini
            await _safe_send(websocket, {"type": "thinking", "iteration": iteration})

            # Call Gemini (with retry on rate limits)
            response = await self._call_with_retry(contents, system_prompt=system_prompt)

            # Guard: no candidates (safety block, quota, etc.)
            if not response.candidates:
                logger.warning("Gemini returned no candidates")
                await _safe_send(websocket, {
                    "type": "error",
                    "message": "The model returned no response. Please try rephrasing your request.",
                })
                return "No response from model.", timeline

            candidate = response.candidates[0]

            # Guard: content or parts can be None on certain finish_reason values
            parts = (candidate.content.parts if candidate.content else None) or []

            # Detect terminal states
            has_function_calls = any(
                hasattr(p, "function_call") and p.function_call
                for p in parts
            )

            if not has_function_calls:
                # Agent is done — return text response
                final_text = "".join(
                    p.text for p in parts if hasattr(p, "text") and p.text
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

                # ── Flush completed turn into persistent history ──────────────
                _flush_history_from_contents()

                # ── Run compaction pipeline on persistent history ─────────────
                tokens_before = estimate_tokens(conversation_history)
                compacted, strategy = await run_compaction_pipeline(
                    history=conversation_history,
                    client=self.client,
                    model=GEMINI_MODEL,
                    token_budget=COMPACT_TOKEN_BUDGET,
                    tool_keep_last=COMPACT_TOOL_KEEP_LAST,
                    summary_keep_last=COMPACT_SUMMARY_KEEP_LAST,
                    window_keep_last=COMPACT_WINDOW_KEEP_LAST,
                )
                if compacted:
                    tokens_after = estimate_tokens(conversation_history)
                    await _safe_send(websocket, {
                        "type": "history_compacted",
                        "strategy": strategy,
                        "tokens_before": tokens_before,
                        "tokens_after": tokens_after,
                    })

                return final_text, timeline

            # ── Process function calls ────────────────────────────────────────
            function_call_parts = [
                p for p in parts
                if hasattr(p, "function_call") and p.function_call
            ]

            # Add model's turn to working contents
            contents.append(candidate.content)

            function_response_parts: list[types.Part] = []

            try:
                for fc_part in function_call_parts:
                    fc = fc_part.function_call
                    tool_name = fc.name
                    tool_args = dict(fc.args) if fc.args else {}

                    # Log tool call with key args (truncate large values)
                    args_summary = {
                        k: (str(v)[:60] + "..." if len(str(v)) > 60 else v)
                        for k, v in tool_args.items()
                        if k != "rows"  # skip bulky row data
                    }
                    if "rows" in tool_args:
                        args_summary["rows"] = f"[{len(tool_args['rows'])} rows]"
                    logger.info(f"  tool[{iteration}] → {tool_name}({args_summary})")

                    start_ts = time.perf_counter()
                    result = await executor.execute(tool_name, tool_args)
                    duration_ms = round((time.perf_counter() - start_ts) * 1000)

                    status = "✓" if result.success else "✗"
                    logger.info(
                        f"  tool[{iteration}] ← {tool_name} {status} ({duration_ms}ms) "
                        f"{result.summary[:100]}"
                    )

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
                        "create_model_scaffold", "clean_data",
                    }:
                        wb_state = serialize_workbook(workbook.wb, workbook._charts)
                        await _safe_send(websocket, {"type": "workbook_update", "state": wb_state})
                        for sheet_name in workbook.get_all_sheet_names():
                            await _safe_send(websocket, {
                                "type": "new_tab",
                                "tab": {"id": sheet_name, "name": sheet_name, "type": "sheet"},
                            })

                    # Trim tool response before adding to contents — reduces tokens sent to Gemini
                    # Read tools (get_sheet_state, parse_document) are passed through untouched
                    trimmed_response = _trim_tool_response(tool_name, result.data, result.summary)
                    function_response_parts.append(
                        types.Part.from_function_response(
                            name=tool_name,
                            response=trimmed_response,
                        )
                    )

                # Feed all responses back in one Content
                contents.append(
                    types.Content(role="user", parts=function_response_parts)
                )

            except Exception as tool_exc:
                # Preserve all history built so far before re-raising so the model
                # retains context of completed actions on the next user message.
                _flush_history_from_contents()
                raise tool_exc

        # ── Hit iteration cap ─────────────────────────────────────────────────
        logger.warning(f"Agent hit MAX_TOOL_ITERATIONS={MAX_TOOL_ITERATIONS}")
        _flush_history_from_contents()
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
        system_prompt: str | None = None,
    ) -> types.GenerateContentResponse:
        """Call Gemini with exponential backoff on rate limit errors."""
        delay = RATE_LIMIT_DELAY_BASE
        n_turns = len(contents)

        for attempt in range(max_retries):
            try:
                t0 = time.perf_counter()
                result = await self.client.aio.models.generate_content(
                    model=GEMINI_MODEL,
                    contents=contents,
                    config=self._make_config(system_prompt=system_prompt),
                )
                elapsed = round((time.perf_counter() - t0) * 1000)
                # Log response metadata
                cand = result.candidates[0] if result.candidates else None
                n_parts = len(cand.content.parts) if cand and cand.content else 0
                has_fc = any(
                    hasattr(p, "function_call") and p.function_call
                    for p in (cand.content.parts if cand and cand.content else [])
                )
                logger.info(
                    f"  gemini ← {elapsed}ms  "
                    f"turns={n_turns}  parts={n_parts}  "
                    f"has_tool_calls={has_fc}  "
                    f"finish={cand.finish_reason if cand else 'none'}"
                )
                return result
            except Exception as exc:
                err = str(exc).lower()
                is_rate_limit = any(k in err for k in (
                    "429", "503", "quota", "rate limit",
                    "resource_exhausted", "unavailable", "overloaded"
                ))

                if is_rate_limit and attempt < max_retries - 1:
                    jitter = (time.time() % 1)
                    sleep_for = min(delay * (2 ** attempt) + jitter, MAX_BACKOFF_SECONDS)
                    logger.warning(
                        f"  gemini rate-limit (attempt {attempt + 1}/{max_retries}), "
                        f"retry in {sleep_for:.1f}s — {str(exc)[:120]}"
                    )
                    await asyncio.sleep(sleep_for)
                else:
                    logger.error(f"  gemini FATAL error: {str(exc)[:200]}")
                    raise

        raise RuntimeError("Exceeded retry attempts")  # unreachable


# ── Helpers ───────────────────────────────────────────────────────────────────

def _trim_tool_response(tool_name: str, data: dict | None, summary: str) -> dict:
    """
    Strip bulky fields from tool responses for write-only tools before sending
    back to Gemini. Only safe to trim tools where the agent does NOT need to
    read the response to continue reasoning.

    Read tools (get_sheet_state, parse_document, validate_workbook) are passed
    through untouched — stripping their data would break formula writing and
    document extraction.
    """
    # These tools are called specifically so Gemini can READ the result.
    # Never trim them — the agent needs the full data to act correctly.
    _READ_TOOLS = {
        "get_sheet_state",
        "parse_document",
        "validate_workbook",
        "get_all_sheet_names",
    }
    if tool_name in _READ_TOOLS or not data:
        return data if data else {"result": summary}

    # For write/mutation tools the agent only needs confirmation, not full data.
    _BULKY_FIELDS = {
        "rows", "cells", "row_data", "full_state",
        "applied_cells", "issues",
    }
    trimmed = {k: v for k, v in data.items() if k not in _BULKY_FIELDS}
    trimmed["_summary"] = summary
    return trimmed


async def _safe_send(websocket, data: dict) -> None:
    """Send a JSON message to the websocket, swallowing disconnect errors."""
    try:
        await websocket.send_json(data)
    except Exception as e:
        logger.debug(f"WebSocket send failed (client likely disconnected): {e}")
