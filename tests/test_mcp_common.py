"""Unit tests for mcp_common -- the shared response layer of all 4 MCP servers.

The critical behaviour under test is the size guard: an oversized JSON payload
must lose *rows* (biggest list halved until it fits), never characters. A raw
character cut produced the DCMA assess_all bug where callers received 1 rule
instead of 14 with no error. shrink_json_text extends the same guarantee to
tools that hold an already-serialised JSON string (asta_mcp_core,
asta_mcp_file, and dispatch()'s str branch).
"""
import json

import pytest

import mcp_common
from mcp_common import (
    MAX_RESPONSE_CHARS,
    dispatch,
    json_response,
    redact,
    register_secret,
    reject_credential_params,
    shrink_json_text,
    truncate_response,
)


# ---------------------------------------------------------------------------
# truncate_response
# ---------------------------------------------------------------------------
class TestTruncateResponse:
    def test_short_text_untouched(self):
        assert truncate_response("hello") == "hello"

    def test_long_text_cut_with_hint(self):
        text = "x" * 30000
        out = truncate_response(text)
        assert len(out) < 30000
        assert "[TRUNCATED]" in out

    def test_cut_prefers_line_boundary(self):
        text = ("line\n" * 6000)  # 30000 chars of 5-char lines
        out = truncate_response(text)
        body = out.split("\n\n... **[TRUNCATED]**")[0]
        assert body.endswith("line")  # ends at a whole line, not mid-word


# ---------------------------------------------------------------------------
# redact / register_secret / reject_credential_params
# ---------------------------------------------------------------------------
class TestRedaction:
    def test_password_flag_masked(self):
        assert "s3cret" not in redact("/password=s3cret rest")

    def test_connect_string_masked(self):
        assert "hunter2" not in redact("Server=x;Password=hunter2;Db=y")

    def test_registered_literal_masked_even_with_space(self):
        register_secret("pa ss word")
        try:
            assert "pa ss word" not in redact("prefix pa ss word suffix")
        finally:
            mcp_common._LITERAL_SECRETS.discard("pa ss word")

    def test_reject_credential_params(self):
        msg = reject_credential_params({"action": "x", "Password": "y"})
        assert msg is not None and "password" in msg.lower()
        assert reject_credential_params({"action": "x", "limit": 5}) is None


# ---------------------------------------------------------------------------
# json_response -- row-wise shrink
# ---------------------------------------------------------------------------
def _big_rows(n, pad=200):
    return [{"id": i, "pad": "x" * pad} for i in range(n)]


class TestJsonResponse:
    def test_small_payload_verbatim(self):
        payload = {"status": "ok", "rows": [1, 2, 3]}
        assert json.loads(json_response(payload)) == payload

    def test_oversized_payload_stays_valid_json(self):
        out = json_response({"status": "ok", "rows": _big_rows(500)})
        parsed = json.loads(out)  # must not raise
        assert parsed["truncated"] is True
        assert parsed["list_length_before_truncate"] == 500
        assert parsed["returned"] == len(parsed["rows"]) < 500
        assert len(out) <= MAX_RESPONSE_CHARS

    def test_nested_list_is_found_and_shrunk(self):
        out = json_response({"status": "ok", "data": {"tasks": _big_rows(500)}})
        parsed = json.loads(out)
        assert parsed["truncated"] is True
        assert len(parsed["data"]["tasks"]) < 500
        assert len(out) <= MAX_RESPONSE_CHARS

    def test_sibling_keys_survive_shrink(self):
        """The assess_all bug: shrinking must never drop the *other* keys."""
        payload = {"status": "ok", "summary": {"score": 14},
                   "rules": _big_rows(500)}
        parsed = json.loads(json_response(payload))
        assert parsed["summary"] == {"score": 14}
        assert parsed["status"] == "ok"

    def test_no_list_falls_back_to_char_cut(self):
        out = json_response({"status": "ok", "blob": "x" * 40000})
        assert "[TRUNCATED]" in out
        assert len(out) <= MAX_RESPONSE_CHARS + 200  # hint tail allowed

    def test_caller_payload_not_mutated(self):
        payload = {"status": "ok", "rows": _big_rows(500)}
        json_response(payload)
        assert len(payload["rows"]) == 500


# ---------------------------------------------------------------------------
# shrink_json_text -- the string-side guard for asta servers
# ---------------------------------------------------------------------------
class TestShrinkJsonText:
    def test_small_text_passthrough(self):
        text = json.dumps({"status": "ok", "rows": [1, 2]})
        assert json.loads(shrink_json_text(text)) == {"status": "ok", "rows": [1, 2]}

    def test_small_text_still_redacted(self):
        out = shrink_json_text('{"cmd": "/password=abc123"}')
        assert "abc123" not in out

    def test_oversized_json_dict_shrinks_rows_not_chars(self):
        text = json.dumps({"status": "ok", "count": 500,
                           "tasks": _big_rows(500)}, default=str)
        assert len(text) > MAX_RESPONSE_CHARS
        out = shrink_json_text(text)
        parsed = json.loads(out)  # the old char-cut made this raise
        assert parsed["truncated"] is True
        assert parsed["count"] == 500          # sibling data intact
        assert len(parsed["tasks"]) < 500
        assert len(out) <= MAX_RESPONSE_CHARS

    def test_oversized_non_json_falls_back_to_char_cut(self):
        out = shrink_json_text("plain text " * 4000)
        assert "[TRUNCATED]" in out

    def test_oversized_json_array_falls_back_to_char_cut(self):
        # Top-level arrays have no key to hang truncate metadata on.
        out = shrink_json_text(json.dumps(_big_rows(500)))
        assert "[TRUNCATED]" in out

    def test_none_passthrough(self):
        assert shrink_json_text(None) is None


# ---------------------------------------------------------------------------
# dispatch
# ---------------------------------------------------------------------------
class TestDispatch:
    TABLE = {
        "ok_dict": lambda p: {"status": "ok", "echo": p.get("x")},
        "ok_str": lambda p: json.dumps(
            {"status": "ok", "rows": _big_rows(500)}, default=str),
        "boom": lambda p: 1 / 0,
    }

    def test_dict_result_wrapped(self):
        out = dispatch("t", {"action": "ok_dict", "x": 7}, self.TABLE)
        assert json.loads(out) == {"status": "ok", "echo": 7}

    def test_str_result_shrinks_as_json_not_chars(self):
        out = dispatch("t", {"action": "ok_str"}, self.TABLE)
        parsed = json.loads(out)  # old truncate_response path broke this
        assert parsed["truncated"] is True
        assert len(out) <= MAX_RESPONSE_CHARS

    def test_missing_action(self):
        parsed = json.loads(dispatch("t", {}, self.TABLE))
        assert parsed["status"] == "error"
        assert "available_actions" in parsed

    def test_unknown_action(self):
        parsed = json.loads(dispatch("t", {"action": "nope"}, self.TABLE))
        assert parsed["status"] == "error"

    def test_exception_becomes_error_envelope(self):
        parsed = json.loads(dispatch("t", {"action": "boom"}, self.TABLE))
        assert parsed["status"] == "error"
        assert parsed["exception_type"] == "ZeroDivisionError"

    def test_credential_param_rejected_before_handler(self):
        parsed = json.loads(
            dispatch("t", {"action": "ok_dict", "password": "x"}, self.TABLE))
        assert parsed["status"] == "error"


# ---------------------------------------------------------------------------
# Server wiring -- the three migrated servers must use the shared guard
# ---------------------------------------------------------------------------
REPO = __import__("pathlib").Path(__file__).resolve().parent.parent


class TestServerWiring:
    def test_asta_core_uses_shared_guard(self):
        src = (REPO / "asta_mcp_core.py").read_text(encoding="utf-8")
        assert "from mcp_common import MAX_RESPONSE_CHARS, shrink_json_text" in src
        assert "def _truncate_response" not in src

    def test_asta_file_uses_shared_guard(self):
        src = (REPO / "asta_mcp_file.py").read_text(encoding="utf-8")
        assert "from mcp_common import MAX_RESPONSE_CHARS, shrink_json_text" in src
        assert "def _truncate_response" not in src

    def test_msproject_uses_shared_envelope(self):
        src = (REPO / "msproject_mcp_core.py").read_text(encoding="utf-8")
        assert "from mcp_common import json_response" in src
        assert "return json.dumps(r, default=str, ensure_ascii=False)" not in src
