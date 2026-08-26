"""Shared helpers for the MCP servers in this repo.

Extracted so p6_mcp_core.py does not repeat the copy-paste layer that already
exists between asta_mcp_core.py and asta_mcp_file.py. Pure functions only --
no COM, no JVM, no I/O -- so the whole module is unit-testable.
"""
from __future__ import annotations

import json
import re
from typing import Any, Callable, Mapping

# ---------------------------------------------------------------------------
# Response size guard (Claude Desktop tool-result limit ~8K tokens)
# ---------------------------------------------------------------------------
MAX_RESPONSE_CHARS = 25000

_TRUNCATE_HINT = (
    "\n\n... **[TRUNCATED]** Response exceeded {n} chars. "
    "Use a smaller `limit`, narrow the `source`, or query a single object."
)


def truncate_response(text: str, max_chars: int = MAX_RESPONSE_CHARS) -> str:
    """Cut `text` at a line boundary once it exceeds `max_chars`."""
    if len(text) <= max_chars:
        return text
    cut = text[:max_chars]
    nl = cut.rfind("\n")
    if nl > max_chars * 0.8:
        cut = cut[:nl]
    return cut + _TRUNCATE_HINT.format(n=max_chars)


# ---------------------------------------------------------------------------
# Secret redaction -- single exit point for every tool response and log line
# ---------------------------------------------------------------------------
_REDACT_PATTERNS = [
    re.compile(r"(/password=)(\S+)", re.I),
    re.compile(r"(/sapwd[=:])(\S+)", re.I),
    re.compile(r"(\bpassword\s*[=:]\s*)([^\s,;&\"']+)", re.I),
    re.compile(r"(\bpwd\s*[=:]\s*)([^\s,;&\"']+)", re.I),
    re.compile(r"(Password\s*=\s*)([^;\s\"']+)", re.I),  # ADO/ODBC connect strings
    re.compile(r"(\bsa/)([^@\s]+)(@)", re.I),          # dbsetup -connection user/pass@...
]

_MASK = "***"

# Literal values registered at runtime (e.g. a password read from the
# credential store) so they can never appear in output even without a pattern.
_LITERAL_SECRETS: set[str] = set()


def register_secret(value: str | None) -> None:
    """Mark a literal string as a secret; it is masked in every redact() call."""
    if value and len(value) >= 3:
        _LITERAL_SECRETS.add(value)


def redact(text: str) -> str:
    """Mask credentials in arbitrary text (responses, logs, command echoes).

    Pattern matching stops at whitespace, so a secret containing a space would
    leak its tail. Callers that hold the literal value (e.g. after reading the
    DPAPI credential store) must call register_secret() -- that path is exact
    and space-safe, and is the one relied on in production.
    """
    if not text:
        return text
    out = text
    for pat in _REDACT_PATTERNS:
        if pat.groups >= 3:
            out = pat.sub(lambda m: m.group(1) + _MASK + m.group(3), out)
        else:
            out = pat.sub(lambda m: m.group(1) + _MASK, out)
    for secret in _LITERAL_SECRETS:
        out = out.replace(secret, _MASK)
    return out


# Parameter names a tool must never accept -- forces credentials through the
# OS credential store instead of the conversation transcript.
FORBIDDEN_PARAM_KEYS = frozenset({
    "password", "passwd", "pwd", "parola", "sapwd", "secret", "token", "api_key",
})


def reject_credential_params(params: Mapping[str, Any]) -> str | None:
    """Return an error message if `params` carries a credential, else None."""
    bad = sorted(k for k in params if k.lower() in FORBIDDEN_PARAM_KEYS)
    if not bad:
        return None
    return (
        "Kimlik bilgisi tool parametresi olarak kabul edilmiyor: "
        + ", ".join(bad)
        + ". Windows Credential Manager / DPAPI kayd\u0131n\u0131 kullan\u0131n."
    )


# ---------------------------------------------------------------------------
# Uniform JSON envelope
# ---------------------------------------------------------------------------
def _dumps(payload: Any) -> str:
    return json.dumps(payload, ensure_ascii=False, indent=2, default=str)


_PREFERRED_LISTS = ("items", "rows", "jobs", "projects", "tasks", "links")


def _find_longest_list(node: Any, path: tuple = (), depth: int = 0
                       ) -> tuple[tuple, int]:
    """Longest list anywhere in the payload, as (path, length).

    Nested lists count too -- read_progress returns {"data": {"tasks": [...]}},
    and only looking at top-level keys left that payload unshrinkable.
    """
    best: tuple[tuple, int] = ((), 0)
    if depth > 4 or not isinstance(node, Mapping):
        return best
    for key, value in node.items():
        if isinstance(value, list):
            weight = len(value) * (10 if key in _PREFERRED_LISTS else 1)
            if weight > best[1]:
                best = (path + (key,), weight)
        elif isinstance(value, Mapping):
            sub = _find_longest_list(value, path + (key,), depth + 1)
            if sub[1] > best[1]:
                best = sub
    return best


def _get_path(payload: Any, path: tuple) -> Any:
    node = payload
    for key in path:
        node = node[key]
    return node


def _set_path(payload: Any, path: tuple, value: Any) -> Any:
    """Copy-on-write down `path` so the caller's payload is untouched."""
    out = dict(payload)
    node = out
    for key in path[:-1]:
        node[key] = dict(node[key])
        node = node[key]
    node[path[-1]] = value
    return out


def json_response(payload: Any, max_chars: int = MAX_RESPONSE_CHARS) -> str:
    """Serialise, redact, then shrink to fit -- keeping the JSON valid.

    Cutting the serialised string mid-object produces JSON the caller cannot
    parse, so oversized payloads lose *rows* instead of characters: the biggest
    list is halved until it fits and the response says how many were returned.
    Only if there is no list to shrink do we fall back to a text cut.
    """
    text = redact(_dumps(payload))
    if len(text) <= max_chars:
        return text

    if isinstance(payload, dict):
        path, _weight = _find_longest_list(payload)
        if path:
            items = list(_get_path(payload, path))
            total = len(items)
            label = ".".join(path)
            n = total
            while n > 0:
                n = n // 2
                shrunk = _set_path(payload, path, items[:n])
                # NOT: 'total' yazilmaz -- tool'un kendi 'count' alaniyla
                # (sorgunun toplam sonucu) karisirdi. Buradaki sayilar yalniz
                # KESILEN LISTE hakkindadir.
                shrunk["truncated"] = True
                shrunk["returned"] = n
                shrunk["list_length_before_truncate"] = total
                shrunk["truncate_note"] = (
                    "Yanit %d karakteri astigi icin '%s' listesi %d satirdan %d "
                    "satira indirildi (sorgunun toplam sonucu icin 'count' "
                    "alanina bakin). Daha kucuk 'limit' verin." 
                    % (max_chars, label, total, n))
                text = redact(_dumps(shrunk))
                if len(text) <= max_chars:
                    return text
    # No list to shrink -- last resort, still flagged.
    return truncate_response(text, max_chars)


def shrink_json_text(text: str, max_chars: int = MAX_RESPONSE_CHARS) -> str:
    """json_response for tools that already hold a serialised JSON string.

    A raw character cut on a JSON string both breaks parseability and can
    silently swallow data (the DCMA assess_all 14-rules-to-1 bug). Oversized
    text that parses as a JSON object is therefore shrunk row-wise through
    json_response; only non-JSON text falls back to the character cut.
    """
    if text is None:
        return text
    if len(text) <= max_chars:
        return redact(text)
    try:
        payload = json.loads(text)
    except (ValueError, TypeError):
        return truncate_response(redact(text), max_chars)
    if isinstance(payload, Mapping):
        return json_response(payload, max_chars)
    return truncate_response(redact(text), max_chars)


def err(message: str, **extra: Any) -> str:
    """Error envelope. Tools return errors as data, never as exceptions."""
    payload: dict[str, Any] = {"status": "error", "error": message}
    payload.update(extra)
    return json_response(payload)


def ok(payload: Mapping[str, Any] | None = None, **extra: Any) -> str:
    payload = dict(payload or {})
    payload.setdefault("status", "ok")
    payload.update(extra)
    return json_response(payload)


# ---------------------------------------------------------------------------
# action -> handler dispatch
# ---------------------------------------------------------------------------
def dispatch(
    tool: str,
    params: Mapping[str, Any],
    table: Mapping[str, Callable[[Mapping[str, Any]], Any]],
) -> str:
    """Route `params['action']` through `table`, wrapping every failure."""
    cred_error = reject_credential_params(params)
    if cred_error:
        return err(cred_error, tool=tool)

    action = params.get("action")
    if not action:
        return err(f"{tool}: 'action' zorunlu.", available_actions=sorted(table))
    handler = table.get(action)
    if handler is None:
        return err(
            f"{tool}: bilinmeyen action '{action}'.",
            available_actions=sorted(table),
        )
    try:
        result = handler(params)
    except Exception as exc:  # noqa: BLE001 - tools must not raise
        return err(f"{tool}({action}) basarisiz: {exc}", exception_type=type(exc).__name__)
    if isinstance(result, str):
        return shrink_json_text(result)
    return json_response(result)
