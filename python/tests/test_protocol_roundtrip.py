"""Round-trip encode/decode and envelope validation for IPC protocol modules."""

from __future__ import annotations

from al_protocol import (
    REQ_ID,
    decode_payload,
    encode_payload,
    make_request,
    validate_request,
    OP_PING,
)
from protocol import (
    REQ_ID as G_REQ_ID,
    decode_message as g_decode_message,
    encode_message as g_encode_message,
    make_request as g_make_request,
    validate_request as g_validate_request,
    OP_PING as G_OP_PING,
)
from shiftkeys_protocol import (
    decode_message as sk_decode_message,
    encode_message as sk_encode_message,
    make_request as sk_make_request,
    validate_request as sk_validate_request,
    OP_HEALTH_CHECK,
)
from wm_protocol import (
    decode_message as wm_decode_message,
    encode_message as wm_encode_message,
    make_request as wm_make_request,
    validate_request as wm_validate_request,
    OP_PING as WM_OP_PING,
)


def test_wm_protocol_roundtrip_ping():
    req = wm_make_request("rid-wm", WM_OP_PING, "", {"a": 1})
    raw = wm_encode_message(req)
    out = wm_decode_message(raw)
    assert out is not None
    assert wm_validate_request(out)
    assert out["id"] == "rid-wm"
    assert out["op"] == WM_OP_PING
    assert out["payload"] == {"a": 1}


def test_shiftkeys_protocol_roundtrip_health():
    req = sk_make_request("rid-sk", OP_HEALTH_CHECK, "", {})
    raw = sk_encode_message(req)
    out = sk_decode_message(raw)
    assert out is not None
    assert sk_validate_request(out)
    assert out["id"] == "rid-sk"
    assert out["op"] == OP_HEALTH_CHECK


def test_al_protocol_payload_roundtrip():
    req = make_request("rid-al", OP_PING, "", {"x": "y"}, 0)
    assert validate_request(req)
    blob = encode_payload(req)
    out = decode_payload(blob)
    assert out is not None
    assert validate_request(out)
    assert out[REQ_ID] == "rid-al"


def test_gemini_protocol_roundtrip_ping():
    req = g_make_request("rid-g", G_OP_PING, "", {"k": 2})
    raw = g_encode_message(req)
    out = g_decode_message(raw)
    assert out is not None
    assert g_validate_request(out)
    assert out[G_REQ_ID] == "rid-g"
    assert out["op"] == G_OP_PING


def test_invalid_envelope_rejected_by_pydantic():
    from wm_protocol import validate_request as wm_v

    assert not wm_v({"op": "Ping"})
    assert not wm_v({"id": "x"})
    assert not wm_v({})
