"""
RightCut — Configuration
Loads environment variables and exposes typed constants.
"""

import os
from dotenv import load_dotenv

load_dotenv()


def _require(key: str) -> str:
    val = os.getenv(key)
    if not val:
        raise RuntimeError(
            f"Required environment variable '{key}' is not set. "
            "Copy .env.example to .env and fill in the values."
        )
    return val


# ── Gemini ────────────────────────────────────────────────────────────────────
GEMINI_API_KEY: str = _require("GEMINI_API_KEY")
GEMINI_MODEL: str = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

# ── Agent behaviour ───────────────────────────────────────────────────────────
MAX_TOOL_ITERATIONS: int = int(os.getenv("MAX_TOOL_ITERATIONS", "35"))
RATE_LIMIT_DELAY_BASE: float = float(os.getenv("RATE_LIMIT_DELAY_BASE", "2.0"))
MAX_BACKOFF_SECONDS: float = float(os.getenv("MAX_BACKOFF_SECONDS", "60.0"))
AGENT_TEMPERATURE: float = float(os.getenv("AGENT_TEMPERATURE", "0.1"))

# ── Uploads ───────────────────────────────────────────────────────────────────
UPLOAD_MAX_SIZE_MB: int = int(os.getenv("UPLOAD_MAX_SIZE_MB", "20"))

# ── CORS ──────────────────────────────────────────────────────────────────────
_extra_origins = [o.strip() for o in os.getenv("CORS_ORIGINS", "").split(",") if o.strip()]
CORS_ORIGINS: list[str] = [
    "http://localhost:5173",
    "http://localhost:5174",
    "http://localhost:3000",
    "http://127.0.0.1:5173",
    *_extra_origins,
]

# ── Session ───────────────────────────────────────────────────────────────────
SESSION_TTL_HOURS: int = int(os.getenv("SESSION_TTL_HOURS", "4"))

# ── History compaction pipeline ───────────────────────────────────────────────
# Estimated token budget before compaction fires.
# Gemini 2.5 Flash supports 1M tokens but we compact early to save cost.
# Tool result collapse (Strategy 1) always runs regardless of this budget.
COMPACT_TOKEN_BUDGET: int = int(os.getenv("COMPACT_TOKEN_BUDGET", "60000"))
# Keep this many recent tool-call groups verbatim when collapsing old ones (Strategy 1)
COMPACT_TOOL_KEEP_LAST: int = int(os.getenv("COMPACT_TOOL_KEEP_LAST", "2"))
# Keep this many recent text groups verbatim during LLM summarization (Strategy 2)
COMPACT_SUMMARY_KEEP_LAST: int = int(os.getenv("COMPACT_SUMMARY_KEEP_LAST", "8"))
# Sliding window backstop — set very high to avoid dropping critical context (Strategy 3)
# Effectively disabled for normal sessions; only fires on extremely long conversations
COMPACT_WINDOW_KEEP_LAST: int = int(os.getenv("COMPACT_WINDOW_KEEP_LAST", "40"))
