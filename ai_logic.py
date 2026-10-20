"""Turn plain-English logic from a mapping file into a formula, using Claude.

    "add igst amt and cgst amt then divide by 2"  ->  ([Igst Tax] + [Cgst Tax]) / 2

The generated formula is validated against the file's real columns before it is
used, and shown in the report so a person can check it.
"""
from __future__ import annotations

import os

import formula

MODEL = "claude-opus-5-5"

SYSTEM = f"""You convert a plain-English description of a spreadsheet calculation into
one formula in this exact syntax:

{formula.HELP}

Rules:
- Use only column names from the list you are given, written exactly as listed, inside [brackets].
- Match loosely-worded names to the closest listed column (e.g. "igst amt" -> [Igst Tax]).
- Reply with the formula only: one line, no explanation, no code fences, no leading "=".
- If the description cannot be expressed with the listed columns, reply exactly: CANNOT: <short reason>"""


class AIError(RuntimeError):
    pass


class Translator:
    """Caches results so each sentence is sent to the API once per run."""

    def __init__(self, api_key: str | None = None):
        self.api_key = (api_key or os.environ.get("ANTHROPIC_API_KEY") or "").strip()
        self._cache: dict[tuple, str] = {}
        self._client = None

    @property
    def available(self) -> bool:
        return bool(self.api_key)

    def to_formula(self, text: str, columns: list[str]) -> str:
        key = (text.strip(), tuple(columns))
        if key not in self._cache:
            out = self._ask(text, columns)
            if out.upper().startswith("CANNOT"):
                raise AIError(f"AI could not convert '{text}': {out.partition(':')[2].strip()}")
            out = out.strip().strip("`").strip()
            try:
                formula.validate(out, columns)
            except formula.FormulaError as e:
                raise AIError(f"AI made an invalid formula for '{text}': {out}  ({e})") from None
            self._cache[key] = out
        return self._cache[key]

    def _ask(self, text: str, columns: list[str]) -> str:
        if not self.available:
            raise AIError(
                f"'{text}' is not a column or formula. To use plain-English logic, enter a "
                "Claude API key (or set ANTHROPIC_API_KEY); otherwise write it as a formula, "
                "e.g. ([Igst Tax] + [Cgst Tax]) / 2")
        import anthropic   # imported lazily: only needed for plain-English logic

        if self._client is None:
            self._client = anthropic.Anthropic(api_key=self.api_key)
        user = "Columns:\n" + "\n".join(f"- {c}" for c in columns) + f"\n\nDescription: {text}"
        try:
            resp = self._client.beta.messages.create(
                model=MODEL,
                max_tokens=2000,
                output_config={"effort": "low"},
                betas=["server-side-fallback-2026-07-01"],
                fallbacks="default",
                system=SYSTEM,
                messages=[{"role": "user", "content": user}],
            )
        except anthropic.AuthenticationError:
            raise AIError("Claude API key is invalid") from None
        except anthropic.RateLimitError:
            raise AIError("Claude API rate limit reached; try again in a minute") from None
        except anthropic.APIConnectionError:
            raise AIError("Cannot reach the Claude API; check your internet connection") from None
        except anthropic.APIStatusError as e:
            raise AIError(f"Claude API error ({e.status_code}): {e.message}") from None
        if resp.stop_reason == "refusal":
            raise AIError(f"AI declined to convert '{text}'; write it as a formula instead")
        return next((b.text for b in resp.content if b.type == "text"), "").strip()
