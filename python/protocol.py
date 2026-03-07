# Gemini IPC protocol: message framing and validation.
# Used by the persistent Python daemon and the AHK client.
# Frame format: length (4 bytes, big-endian) + UTF-8 JSON payload.

import json
import struct
import sys


def encode_message(obj: dict) -> bytes:
    payload = json.dumps(obj, ensure_ascii=False).encode("utf-8")
    return struct.pack(">I", len(payload)) + payload


def decode_message(data: bytes) -> dict | None:
    if len(data) < 4:
        return None
    (length,) = struct.unpack(">I", data[:4])
    if len(data) < 4 + length:
        return None
    payload = data[4 : 4 + length].decode("utf-8")
    return json.loads(payload)


def read_frame(stream):
    """Read one length-prefixed frame from stream. Returns decoded dict or None on EOF/error."""
    header = stream.buffer.read(4)
    if len(header) < 4:
        return None
    (length,) = struct.unpack(">I", header)
    if length > 1024 * 1024:
        return None
    payload = stream.buffer.read(length)
    if len(payload) < length:
        return None
    try:
        return json.loads(payload.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError):
        return None


def write_frame(stream, obj: dict) -> None:
    stream.buffer.write(encode_message(obj))
    stream.buffer.flush()


# Command IDs (AHK sends these; Python responds with same id in response)
CMD_PING = "ping"
CMD_PROCESS_TEXT = "process_text"
CMD_VALIDATE = "validate"

# Required request keys
REQUIRED_KEYS = ("id",)


def validate_request(obj: dict) -> bool:
    if not isinstance(obj, dict):
        return False
    return all(k in obj for k in REQUIRED_KEYS)
