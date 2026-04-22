"""
RightCut — Conversation compaction strategies.

Modelled after the Microsoft Agent Framework compaction spec:
https://learn.microsoft.com/en-us/agent-framework/agents/conversations/compaction

Three layered strategies applied in order (gentlest → most aggressive):
  1. ToolResultCompactionStrategy   — collapse old tool call/response groups into 1-line summaries (free)
  2. SummarizationStrategy          — LLM-summarize old user↔agent turns when token budget exceeded (cheap)
  3. SlidingWindowStrategy          — hard backstop: keep last N groups regardless

The pipeline runs AFTER each completed turn, mutating conversation_history in place.
Returns (compacted: bool, strategy_used: str | None).
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from enum import Enum
from typing import TYPE_CHECKING

from google.genai import types

logger = logging.getLogger(__name__)


# ── Group classification ──────────────────────────────────────────────────────

class GroupKind(str, Enum):
    SYSTEM        = "system"
    USER_TEXT     = "user_text"       # plain user message (no file/tool content)
    ASSISTANT_TEXT = "assistant_text" # plain model text response
    TOOL_CALL     = "tool_call"       # model function_call Content
    TOOL_RESPONSE = "tool_response"   # user function_response Content (paired with TOOL_CALL)


@dataclass
class MessageGroup:
    kind: GroupKind
    contents: list[types.Content]   # 1 or 2 Content objects (tool_call + tool_response are paired)

    @property
    def token_estimate(self) -> int:
        """Character-count / 4 heuristic — same as CharacterEstimatorTokenizer."""
        total = 0
        for c in self.contents:
            for p in (c.parts or []):
                if hasattr(p, "text") and p.text:
                    total += len(p.text)
                else:
                    # function_call / function_response — serialize roughly
                    total += len(str(p))
        return max(1, total // 4)


def classify_history(history: list[types.Content]) -> list[MessageGroup]:
    """
    Group a flat Content list into MessageGroup atoms.
    Tool call (model) + tool response (user) are paired into one TOOL_CALL group.
    """
    groups: list[MessageGroup] = []
    i = 0
    while i < len(history):
        c = history[i]
        parts = c.parts or []

        # Detect function_call parts in a model turn
        has_fc = any(hasattr(p, "function_call") and p.function_call for p in parts)
        # Detect function_response parts in a user turn
        has_fr = any(hasattr(p, "function_response") and p.function_response for p in parts)

        if has_fc and c.role == "model":
            # Pair with the next user turn if it's a function_response
            if i + 1 < len(history):
                nxt = history[i + 1]
                nxt_parts = nxt.parts or []
                nxt_has_fr = any(
                    hasattr(p, "function_response") and p.function_response
                    for p in nxt_parts
                )
                if nxt_has_fr:
                    groups.append(MessageGroup(kind=GroupKind.TOOL_CALL, contents=[c, nxt]))
                    i += 2
                    continue
            # Unpaired function_call (shouldn't happen in well-formed history)
            groups.append(MessageGroup(kind=GroupKind.TOOL_CALL, contents=[c]))
        elif has_fr and c.role == "user":
            # Orphaned function_response — treat as tool_response
            groups.append(MessageGroup(kind=GroupKind.TOOL_RESPONSE, contents=[c]))
        elif c.role == "model":
            groups.append(MessageGroup(kind=GroupKind.ASSISTANT_TEXT, contents=[c]))
        else:
            # user text, system, context injections
            groups.append(MessageGroup(kind=GroupKind.USER_TEXT, contents=[c]))

        i += 1

    return groups


def flatten_groups(groups: list[MessageGroup]) -> list[types.Content]:
    result: list[types.Content] = []
    for g in groups:
        result.extend(g.contents)
    return result


def estimate_tokens(history: list[types.Content]) -> int:
    total = 0
    for c in history:
        for p in (c.parts or []):
            if hasattr(p, "text") and p.text:
                total += len(p.text)
            else:
                total += len(str(p))
    return max(1, total // 4)


# ── Strategy 1: ToolResultCompactionStrategy ─────────────────────────────────

def compact_tool_results(
    groups: list[MessageGroup],
    keep_last: int = 2,
) -> tuple[list[MessageGroup], bool]:
    """
    Collapse TOOL_CALL groups (except the most recent `keep_last`) into
    a single-line summary Content. No LLM needed.

    Equivalent to Microsoft's ToolResultCompactionStrategy(keep_last_tool_call_groups=keep_last).
    """
    tool_indices = [i for i, g in enumerate(groups) if g.kind == GroupKind.TOOL_CALL]

    if len(tool_indices) <= keep_last:
        return groups, False

    to_collapse_indices = set(tool_indices[:-keep_last])
    changed = False
    new_groups: list[MessageGroup] = []

    for i, g in enumerate(groups):
        if i in to_collapse_indices:
            # Build a short summary of what tools were called
            summaries: list[str] = []
            for c in g.contents:
                for p in (c.parts or []):
                    if hasattr(p, "function_call") and p.function_call:
                        fc = p.function_call
                        summaries.append(f"{fc.name}")
                    elif hasattr(p, "function_response") and p.function_response:
                        fr = p.function_response
                        # Grab first meaningful field from response
                        resp = fr.response or {}
                        summary_val = (
                            resp.get("_summary")
                            or resp.get("result")
                            or resp.get("sheet_name")
                            or resp.get("rows_written")
                            or "ok"
                        )
                        summaries.append(f"→ {str(summary_val)[:60]}")

            compact_text = "[Tool: " + " ".join(summaries) + "]"
            synthetic = types.Content(
                role="user",
                parts=[types.Part.from_text(text=compact_text)],
            )
            new_groups.append(MessageGroup(kind=GroupKind.USER_TEXT, contents=[synthetic]))
            changed = True
        else:
            new_groups.append(g)

    return new_groups, changed


# ── Strategy 2: SummarizationStrategy ────────────────────────────────────────

async def compact_summarize(
    groups: list[MessageGroup],
    client,
    model: str,
    keep_last_groups: int = 6,
    token_budget: int = 24_000,
) -> tuple[list[MessageGroup], bool]:
    """
    When token count exceeds token_budget, summarize the oldest user/assistant
    text groups (not tool groups — those are already collapsed) into one
    synthetic [SUMMARY] user Content.

    Equivalent to Microsoft's SummarizationStrategy(target_count=keep_last_groups).
    """
    current_tokens = sum(g.token_estimate for g in groups)
    if current_tokens <= token_budget:
        return groups, False

    # Find non-tool groups
    text_group_indices = [
        i for i, g in enumerate(groups)
        if g.kind in (GroupKind.USER_TEXT, GroupKind.ASSISTANT_TEXT)
    ]

    if len(text_group_indices) <= keep_last_groups:
        return groups, False

    # Summarize everything except the most recent keep_last_groups text groups
    to_summarize_indices = set(text_group_indices[:-keep_last_groups])

    lines: list[str] = []
    for i in sorted(to_summarize_indices):
        g = groups[i]
        role = "USER" if g.kind == GroupKind.USER_TEXT else "AGENT"
        for c in g.contents:
            for p in (c.parts or []):
                if hasattr(p, "text") and p.text:
                    # Skip internal context injections
                    txt = p.text
                    if not txt.startswith("[CONTEXT:") and not txt.startswith("[Tool:"):
                        lines.append(f"{role}: {txt[:600]}")

    if not lines:
        return groups, False

    transcript = "\n\n".join(lines)
    prompt = (
        "Summarise this conversation for a financial modelling assistant. "
        "Be factual and concise (max 250 words). Include: what model was built, "
        "key assumptions used (with numbers), key outputs, and any decisions made. "
        "Skip greetings and small talk.\n\n"
        f"TRANSCRIPT:\n{transcript}"
    )

    try:
        from google.genai import types as _types
        resp = await client.aio.models.generate_content(
            model=model,
            contents=[_types.Content(role="user", parts=[_types.Part.from_text(text=prompt)])],
            config=_types.GenerateContentConfig(temperature=0.1, max_output_tokens=400),
        )
        cands = resp.candidates or []
        cparts = (cands[0].content.parts if cands and cands[0].content else None) or []
        summary_text = (cparts[0].text or "").strip() if cparts else ""
        if not summary_text:
            raise ValueError("Empty summary response")
    except Exception as e:
        logger.warning(f"Summarization LLM call failed: {e}")
        summary_text = "(Earlier conversation summarized — context preserved.)"

    summary_content = types.Content(
        role="user",
        parts=[types.Part.from_text(
            text=f"[CONVERSATION SUMMARY]\n{summary_text}"
        )],
    )
    summary_group = MessageGroup(kind=GroupKind.USER_TEXT, contents=[summary_content])

    # Replace summarized groups with the single summary group
    new_groups: list[MessageGroup] = [summary_group]
    for i, g in enumerate(groups):
        if i not in to_summarize_indices:
            new_groups.append(g)

    logger.info(
        f"Summarization: collapsed {len(to_summarize_indices)} text groups → 1 summary. "
        f"Tokens before: {current_tokens}, after: ~{sum(g.token_estimate for g in new_groups)}"
    )
    return new_groups, True


# ── Strategy 3: SlidingWindowStrategy ────────────────────────────────────────

def compact_sliding_window(
    groups: list[MessageGroup],
    keep_last: int = 12,
    token_budget: int = 16_000,
) -> tuple[list[MessageGroup], bool]:
    """
    Hard backstop: if still over token_budget after earlier strategies,
    keep only the most recent `keep_last` non-tool groups.

    Equivalent to Microsoft's SlidingWindowStrategy(keep_last_groups=keep_last).
    """
    current_tokens = sum(g.token_estimate for g in groups)
    if current_tokens <= token_budget:
        return groups, False

    if len(groups) <= keep_last:
        return groups, False

    new_groups = groups[-keep_last:]
    logger.warning(
        f"SlidingWindow backstop: dropped {len(groups) - keep_last} groups. "
        f"Tokens before: {current_tokens}, after: ~{sum(g.token_estimate for g in new_groups)}"
    )
    return new_groups, True


# ── Pipeline: TokenBudgetComposedStrategy ────────────────────────────────────

async def run_compaction_pipeline(
    history: list[types.Content],
    client,
    model: str,
    token_budget: int = 32_000,
    tool_keep_last: int = 2,
    summary_keep_last: int = 6,
    window_keep_last: int = 12,
) -> tuple[bool, str | None]:
    """
    Run the full compaction pipeline on history (mutates in place).
    Mirrors TokenBudgetComposedStrategy from the Microsoft spec.

    Order (gentlest → most aggressive):
      1. ToolResultCompactionStrategy  — always runs, collapses old tool JSON
      2. SummarizationStrategy         — runs if still over token_budget
      3. SlidingWindowStrategy         — emergency backstop

    Returns (any_change: bool, strategy_name: str | None).
    """
    current_tokens = estimate_tokens(history)

    # Always classify into groups
    groups = classify_history(history)
    any_change = False
    strategy_used: str | None = None

    # ── Step 1: Collapse old tool results (free, always run) ──────────────────
    groups, changed = compact_tool_results(groups, keep_last=tool_keep_last)
    if changed:
        any_change = True
        strategy_used = "tool_result"
        logger.info("Compaction: ToolResultStrategy applied")

    # ── Step 2: Summarize if still over budget ────────────────────────────────
    new_tokens = sum(g.token_estimate for g in groups)
    if new_tokens > token_budget:
        groups, changed = await compact_summarize(
            groups, client, model,
            keep_last_groups=summary_keep_last,
            token_budget=token_budget,
        )
        if changed:
            any_change = True
            strategy_used = "summarization"

    # ── Step 3: Sliding window backstop ──────────────────────────────────────
    new_tokens = sum(g.token_estimate for g in groups)
    if new_tokens > token_budget:
        groups, changed = compact_sliding_window(
            groups,
            keep_last=window_keep_last,
            token_budget=token_budget,
        )
        if changed:
            any_change = True
            strategy_used = "sliding_window"

    if any_change:
        flat = flatten_groups(groups)
        history.clear()
        history.extend(flat)
        final_tokens = estimate_tokens(history)
        logger.info(
            f"Compaction complete via '{strategy_used}': "
            f"{current_tokens} → {final_tokens} estimated tokens "
            f"({round((1 - final_tokens/current_tokens)*100)}% reduction)"
        )

    return any_change, strategy_used
